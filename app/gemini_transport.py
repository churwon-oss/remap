from __future__ import annotations

"""HTTPS transport helpers for Gemini API calls.

ReMap is mainly used on Windows school networks. Many schools inspect HTTPS
traffic and deploy their inspection CA through the Windows certificate store.
Python's bundled OpenSSL trust configuration may not see that CA, and Python
3.13 also enables stricter RFC 5280 checks which can reject older inspection
certificates with messages such as "Missing Authority Key Identifier".

This module keeps certificate and hostname verification enabled at all times.
It first uses the native Windows trust store through ``truststore``. Only when
that fails with a certificate-verification error does it retry with a
Windows-root-backed OpenSSL context that relaxes *only* VERIFY_X509_STRICT.
"""

import os
import ssl
import threading
import urllib.error
import urllib.request
from dataclasses import dataclass
from typing import Any, Iterable

try:  # Optional for source users, bundled in the Windows release.
    import truststore  # type: ignore
except ImportError:  # pragma: no cover - exercised by fallback tests
    truststore = None  # type: ignore[assignment]


@dataclass(frozen=True)
class TLSAttempt:
    mode: str
    label: str
    context: ssl.SSLContext | Any | None


class GeminiCertificateError(RuntimeError):
    """Raised after all verified school-network TLS modes have failed."""

    def __init__(self, original: BaseException, attempted_modes: Iterable[str]) -> None:
        self.original = original
        self.attempted_modes = tuple(attempted_modes)
        super().__init__(friendly_certificate_message(original, self.attempted_modes))


_CONTEXT_LOCK = threading.RLock()
_CACHED_CONTEXTS: dict[str, ssl.SSLContext | Any] = {}


def _exception_chain(exc: BaseException) -> list[BaseException]:
    out: list[BaseException] = []
    seen: set[int] = set()
    current: BaseException | None = exc
    while current is not None and id(current) not in seen:
        seen.add(id(current))
        out.append(current)
        reason = getattr(current, "reason", None)
        if isinstance(reason, BaseException) and id(reason) not in seen:
            current = reason
        elif isinstance(current.__cause__, BaseException) and id(current.__cause__) not in seen:
            current = current.__cause__
        elif isinstance(current.__context__, BaseException) and id(current.__context__) not in seen:
            current = current.__context__
        else:
            current = None
    return out


def is_certificate_verification_error(exc: BaseException) -> bool:
    for item in _exception_chain(exc):
        if isinstance(item, ssl.SSLCertVerificationError):
            return True
        text = str(item).lower()
        if any(
            marker in text
            for marker in (
                "certificate_verify_failed",
                "certificate verify failed",
                "missing authority key identifier",
                "basic constraints of ca cert not marked critical",
                "unable to get local issuer certificate",
                "self-signed certificate in certificate chain",
            )
        ):
            return True
    return False


def _native_windows_context() -> ssl.SSLContext | Any:
    """Return a TLS client context backed by the Windows certificate store."""
    with _CONTEXT_LOCK:
        cached = _CACHED_CONTEXTS.get("windows_native")
        if cached is not None:
            return cached
        if truststore is None:
            raise RuntimeError("truststore 모듈이 설치되어 있지 않습니다.")
        context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        context.check_hostname = True
        context.verify_mode = ssl.CERT_REQUIRED
        _CACHED_CONTEXTS["windows_native"] = context
        return context


def _windows_root_pem_bundle() -> str:
    if os.name != "nt" or not hasattr(ssl, "enum_certificates"):
        return ""
    pem_rows: list[str] = []
    seen: set[bytes] = set()
    # Only ROOT is loaded. Loading the Windows CA/intermediate store as trust
    # anchors would grant more trust than Windows itself does.
    for cert_bytes, encoding_type, trust in ssl.enum_certificates("ROOT"):  # type: ignore[attr-defined]
        if encoding_type != "x509_asn" or not isinstance(cert_bytes, bytes):
            continue
        if cert_bytes in seen:
            continue
        # ``trust`` is True for all purposes or a set of OIDs. Server-auth roots
        # are safe to include; unrelated-purpose roots are skipped.
        if trust is not True and isinstance(trust, set):
            server_auth_oid = "1.3.6.1.5.5.7.3.1"
            if server_auth_oid not in trust:
                continue
        seen.add(cert_bytes)
        pem_rows.append(ssl.DER_cert_to_PEM_cert(cert_bytes))
    return "".join(pem_rows)


def _windows_compat_context() -> ssl.SSLContext:
    """Verified OpenSSL context for older school inspection certificates.

    Hostname verification and CERT_REQUIRED remain enabled. The only relaxed
    check is VERIFY_X509_STRICT, added to create_default_context() in Python
    3.13 and known to reject some older enterprise inspection certificates.
    """
    with _CONTEXT_LOCK:
        cached = _CACHED_CONTEXTS.get("windows_compat")
        if cached is not None:
            return cached
        context = ssl.create_default_context(purpose=ssl.Purpose.SERVER_AUTH)
        context.check_hostname = True
        context.verify_mode = ssl.CERT_REQUIRED
        strict_flag = getattr(ssl, "VERIFY_X509_STRICT", 0)
        if strict_flag:
            context.verify_flags &= ~strict_flag
        bundle = _windows_root_pem_bundle()
        if bundle:
            context.load_verify_locations(cadata=bundle)
        _CACHED_CONTEXTS["windows_compat"] = context
        return context


def _default_compat_context() -> ssl.SSLContext:
    """Strict-only compatibility context used by non-frozen source builds.

    This is intentionally not used on Render unless the default connection
    fails specifically with a certificate validation error.
    """
    with _CONTEXT_LOCK:
        cached = _CACHED_CONTEXTS.get("default_compat")
        if cached is not None:
            return cached
        context = ssl.create_default_context(purpose=ssl.Purpose.SERVER_AUTH)
        context.check_hostname = True
        context.verify_mode = ssl.CERT_REQUIRED
        strict_flag = getattr(ssl, "VERIFY_X509_STRICT", 0)
        if strict_flag:
            context.verify_flags &= ~strict_flag
        _CACHED_CONTEXTS["default_compat"] = context
        return context


def tls_attempts() -> list[TLSAttempt]:
    """Return verified TLS attempts in safest-to-most-compatible order."""
    attempts: list[TLSAttempt] = []
    if os.name == "nt":
        if truststore is not None:
            try:
                attempts.append(
                    TLSAttempt(
                        mode="windows_system",
                        label="Windows 인증서 저장소",
                        context=_native_windows_context(),
                    )
                )
            except Exception:
                # The OpenSSL Windows-root fallback remains available.
                pass
        try:
            attempts.append(
                TLSAttempt(
                    mode="school_compat",
                    label="학교망 인증서 호환 모드",
                    context=_windows_compat_context(),
                )
            )
        except Exception:
            pass
        if not attempts:
            attempts.append(TLSAttempt(mode="python_default", label="Python 기본 인증서", context=None))
        return attempts

    # Render and ordinary source environments use their normal OS/OpenSSL trust
    # configuration. A strict-only fallback is retained for Python 3.13 source
    # users, but is reached only after a certificate-specific failure.
    attempts.append(TLSAttempt(mode="python_default", label="시스템 기본 인증서", context=None))
    try:
        attempts.append(
            TLSAttempt(
                mode="strict_compat",
                label="인증서 형식 호환 모드",
                context=_default_compat_context(),
            )
        )
    except Exception:
        pass
    return attempts


def open_verified_url(
    request: urllib.request.Request | str,
    *,
    timeout: float,
) -> tuple[Any, TLSAttempt]:
    """Open an HTTPS URL using verified school-network-compatible contexts.

    Retries occur only for certificate-verification failures. HTTP errors,
    DNS errors, timeouts, and other network failures are returned immediately,
    preventing duplicate successful POST requests.
    """
    attempts = tls_attempts()
    attempted_labels: list[str] = []
    last_cert_error: BaseException | None = None

    for attempt in attempts:
        attempted_labels.append(attempt.label)
        try:
            kwargs: dict[str, Any] = {"timeout": timeout}
            if attempt.context is not None:
                kwargs["context"] = attempt.context
            response = urllib.request.urlopen(request, **kwargs)
            return response, attempt
        except urllib.error.HTTPError:
            raise
        except Exception as exc:
            if not is_certificate_verification_error(exc):
                raise
            last_cert_error = exc
            continue

    if last_cert_error is not None:
        raise GeminiCertificateError(last_cert_error, attempted_labels) from last_cert_error
    raise RuntimeError("Gemini HTTPS 연결 방식을 준비하지 못했습니다.")


def friendly_certificate_message(exc: BaseException, attempted_modes: Iterable[str] = ()) -> str:
    text = " ".join(str(item) for item in _exception_chain(exc))
    lowered = text.lower()
    modes = ", ".join(attempted_modes)
    if "missing authority key identifier" in lowered or "basic constraints" in lowered:
        reason = "학교 보안장비 인증서가 Python 3.13의 엄격한 형식 검사와 호환되지 않습니다."
    elif "unable to get local issuer certificate" in lowered or "self-signed certificate" in lowered:
        reason = "학교 보안 인증서가 이 PC의 신뢰 저장소에서 확인되지 않습니다."
    else:
        reason = "학교망의 HTTPS 인증서를 확인하지 못했습니다."
    attempted = f" 확인 방식: {modes}." if modes else ""
    return (
        f"{reason}{attempted} 인증서 검증을 끄지는 않았습니다. "
        "학교 전산 담당자에게 HTTPS 검사 루트 인증서가 Windows의 "
        "'신뢰할 수 있는 루트 인증 기관'에 설치되어 있는지 확인해 달라고 요청하세요."
    )
