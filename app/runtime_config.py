from __future__ import annotations

import base64
import ctypes
import json
import os
import shutil
import subprocess
import socket
import sys
import threading
import time
from ctypes import wintypes
from pathlib import Path
from typing import Any
from urllib.parse import urlsplit

APP_NAME = "ReMap"
APP_VERSION = os.environ.get("REMAP_VERSION", "release").strip() or "release"
DEFAULT_GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "").strip()
_ALLOWED_RUNTIME_MODES = {"developer", "windows", "render"}


def build_windows_internet_shortcut(
    target_url: str,
    icon_file: str = "%SystemRoot%\\System32\\url.dll",
    icon_index: int = 0,
) -> bytes:
    """Return a Windows .url file without a UTF-8 BOM.

    The target URL remains ASCII for maximum compatibility. ``icon_file`` may
    use Windows environment variables, which lets the local EXE launcher point
    the shortcut at ReMap's character icon without baking a user-specific path
    into the file.
    """
    value = str(target_url or "").strip()
    parsed = urlsplit(value)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("학생 접속 주소가 올바르지 않습니다.")
    if "\r" in value or "\n" in value:
        raise ValueError("학생 접속 주소에 허용되지 않는 줄바꿈이 있습니다.")

    icon_value = str(icon_file or "").strip()
    if "\r" in icon_value or "\n" in icon_value:
        raise ValueError("바로가기 아이콘 경로에 허용되지 않는 줄바꿈이 있습니다.")
    try:
        icon_index_value = int(icon_index)
    except (TypeError, ValueError):
        icon_index_value = 0

    rows = ["[InternetShortcut]", f"URL={value}"]
    if icon_value:
        rows.extend([f"IconFile={icon_value}", f"IconIndex={icon_index_value}"])
    content = "\r\n".join(rows) + "\r\n"
    try:
        return content.encode("ascii")
    except UnicodeEncodeError as exc:
        raise ValueError("바로가기 아이콘 경로는 영문 경로나 Windows 환경 변수를 사용해야 합니다.") from exc


def _powershell_literal(value: str) -> str:
    return "'" + str(value).replace("'", "''") + "'"


def build_windows_shell_link_script(link_path: Path, target_url: str, icon_path: Path) -> str:
    """Return the PowerShell used to create a character-icon Windows .lnk."""
    value = str(target_url or "").strip()
    parsed = urlsplit(value)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("학생 접속 주소가 올바르지 않습니다.")
    target_path = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32" / "rundll32.exe"
    arguments = f'url.dll,FileProtocolHandler "{value}"'
    return "\n".join(
        [
            "$ErrorActionPreference = 'Stop'",
            "$shell = New-Object -ComObject WScript.Shell",
            f"$shortcut = $shell.CreateShortcut({_powershell_literal(str(link_path.resolve()))})",
            f"$shortcut.TargetPath = {_powershell_literal(str(target_path))}",
            f"$shortcut.Arguments = {_powershell_literal(arguments)}",
            f"$shortcut.WorkingDirectory = {_powershell_literal(str(target_path.parent))}",
            f"$shortcut.IconLocation = {_powershell_literal(str(icon_path.resolve()) + ',0')}",
            "$shortcut.Description = 'ReMap 수업 시작'",
            "$shortcut.WindowStyle = 1",
            "$shortcut.Save()",
        ]
    )


def create_windows_shell_link(link_path: Path, target_url: str, icon_path: Path) -> Path:
    """Create a real Windows Shell Link using the built-in WScript COM object."""
    if os.name != "nt":
        raise OSError("Windows 바로가기는 Windows에서만 생성할 수 있습니다.")
    powershell = shutil.which("powershell.exe") or shutil.which("powershell")
    if not powershell:
        raise FileNotFoundError("Windows PowerShell을 찾지 못했습니다.")
    link_path.parent.mkdir(parents=True, exist_ok=True)
    script = build_windows_shell_link_script(link_path, target_url, icon_path)
    encoded = base64.b64encode(script.encode("utf-16le")).decode("ascii")
    completed = subprocess.run(
        [
            powershell,
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-EncodedCommand",
            encoded,
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        creationflags=0x08000000,
        timeout=20,
        check=False,
    )
    if completed.returncode != 0 or not link_path.exists() or link_path.stat().st_size < 64:
        detail = (completed.stderr or completed.stdout or "바로가기 파일이 생성되지 않았습니다.").strip()
        raise OSError(f"ReMap 학생용 바로가기를 만들지 못했습니다.\n{detail}")
    try:
        ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)
    except Exception:
        pass
    return link_path


def _runtime_mode() -> str:
    mode = os.environ.get("REMAP_RUNTIME", "developer").strip().lower()
    return mode if mode in _ALLOWED_RUNTIME_MODES else "developer"


def _default_data_dir() -> Path:
    configured = os.environ.get("REMAP_DATA_DIR", "").strip()
    if configured:
        return Path(configured).expanduser().resolve()
    if os.name == "nt":
        base = os.environ.get("APPDATA") or str(Path.home())
        return Path(base) / APP_NAME
    return Path.home() / ".remap"


def _is_private_ipv4(value: str) -> bool:
    try:
        parts = [int(p) for p in value.split(".")]
    except (TypeError, ValueError):
        return False
    if len(parts) != 4 or any(p < 0 or p > 255 for p in parts):
        return False
    return (
        parts[0] == 10
        or (parts[0] == 172 and 16 <= parts[1] <= 31)
        or (parts[0] == 192 and parts[1] == 168)
        or (parts[0] == 169 and parts[1] == 254)
    )


def get_ip_candidates() -> list[str]:
    """Return usable local IPv4 addresses, preferring the active private interface."""
    candidates: list[str] = []

    # The UDP connect does not transmit data. It asks the OS which interface would
    # be used for normal outbound traffic and is the best default on most PCs.
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.connect(("8.8.8.8", 80))
            value = sock.getsockname()[0]
            if value and not value.startswith("127."):
                candidates.append(value)
    except OSError:
        pass

    try:
        for row in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET, socket.SOCK_STREAM):
            value = row[4][0]
            if value and not value.startswith("127."):
                candidates.append(value)
    except OSError:
        pass

    # psutil is optional for developer mode but bundled in the Windows build. It
    # improves detection when Ethernet, Wi-Fi, VPN, and virtual adapters coexist.
    try:
        import psutil  # type: ignore

        for entries in psutil.net_if_addrs().values():
            for entry in entries:
                if entry.family == socket.AF_INET:
                    value = str(entry.address or "")
                    if value and not value.startswith("127."):
                        candidates.append(value)
    except Exception:
        pass

    unique: list[str] = []
    for value in candidates:
        if value not in unique:
            unique.append(value)
    if not unique:
        return ["127.0.0.1"]
    preferred = unique[0]
    remainder = sorted(
        unique[1:],
        key=lambda value: (not _is_private_ipv4(value), value.startswith("169.254."), value),
    )
    return [preferred, *[value for value in remainder if value != preferred]]


class _DataBlob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]


def _bytes_to_blob(data: bytes) -> tuple[_DataBlob, Any]:
    if not data:
        buffer = (ctypes.c_byte * 1)()
        return _DataBlob(0, buffer), buffer
    buffer = (ctypes.c_byte * len(data)).from_buffer_copy(data)
    return _DataBlob(len(data), buffer), buffer


def _dpapi_encrypt(value: str) -> str:
    if os.name != "nt":
        raise RuntimeError("Windows DPAPI is only available on Windows")
    raw = value.encode("utf-8")
    in_blob, in_buffer = _bytes_to_blob(raw)
    out_blob = _DataBlob()
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    crypt32.CryptProtectData.argtypes = [
        ctypes.POINTER(_DataBlob),
        wintypes.LPCWSTR,
        ctypes.POINTER(_DataBlob),
        ctypes.c_void_p,
        ctypes.c_void_p,
        wintypes.DWORD,
        ctypes.POINTER(_DataBlob),
    ]
    crypt32.CryptProtectData.restype = wintypes.BOOL
    kernel32.LocalFree.argtypes = [wintypes.HLOCAL]
    kernel32.LocalFree.restype = wintypes.HLOCAL
    ok = crypt32.CryptProtectData(
        ctypes.byref(in_blob),
        "ReMap Gemini API Key",
        None,
        None,
        None,
        0,
        ctypes.byref(out_blob),
    )
    _ = in_buffer
    if not ok:
        raise ctypes.WinError()
    try:
        encrypted = ctypes.string_at(out_blob.pbData, out_blob.cbData)
        return base64.b64encode(encrypted).decode("ascii")
    finally:
        kernel32.LocalFree(ctypes.cast(out_blob.pbData, wintypes.HLOCAL))


def _dpapi_decrypt(value: str) -> str:
    if os.name != "nt":
        raise RuntimeError("Windows DPAPI is only available on Windows")
    encrypted = base64.b64decode(value.encode("ascii"))
    in_blob, in_buffer = _bytes_to_blob(encrypted)
    out_blob = _DataBlob()
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    crypt32.CryptUnprotectData.argtypes = [
        ctypes.POINTER(_DataBlob),
        ctypes.POINTER(wintypes.LPWSTR),
        ctypes.POINTER(_DataBlob),
        ctypes.c_void_p,
        ctypes.c_void_p,
        wintypes.DWORD,
        ctypes.POINTER(_DataBlob),
    ]
    crypt32.CryptUnprotectData.restype = wintypes.BOOL
    kernel32.LocalFree.argtypes = [wintypes.HLOCAL]
    kernel32.LocalFree.restype = wintypes.HLOCAL
    ok = crypt32.CryptUnprotectData(
        ctypes.byref(in_blob),
        None,
        None,
        None,
        None,
        0,
        ctypes.byref(out_blob),
    )
    _ = in_buffer
    if not ok:
        raise ctypes.WinError()
    try:
        plain = ctypes.string_at(out_blob.pbData, out_blob.cbData)
        return plain.decode("utf-8")
    finally:
        kernel32.LocalFree(ctypes.cast(out_blob.pbData, wintypes.HLOCAL))


class RuntimeSettings:
    def __init__(self) -> None:
        self._lock = threading.RLock()
        self.started_at = time.time()
        self.mode = _runtime_mode()
        self.data_dir = _default_data_dir()
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.results_dir = self.data_dir / "결과"
        self.questions_dir = self.data_dir / "학생문제"
        self.shortcuts_dir = self.data_dir / "학생 접속"
        self.logs_dir = self.data_dir / "로그"
        for directory in (self.results_dir, self.questions_dir, self.shortcuts_dir, self.logs_dir):
            directory.mkdir(parents=True, exist_ok=True)
        self._settings_path = self.data_dir / "settings.json"
        self._session_api_key = ""
        self._saved_api_key = ""
        self._model = DEFAULT_GEMINI_MODEL
        self._available_models: list[dict[str, Any]] = []
        self._models_checked_at = 0.0
        self._persist_enabled = False
        self._ai_tls_mode = ""
        self._ai_tls_label = ""
        self._ai_tls_last_error = ""
        self._ai_tls_checked_at = 0.0
        self._public_host = os.environ.get("REMAP_PUBLIC_HOST", "").strip()
        self._port_explicit = "REMAP_PORT" in os.environ
        self._port = self._safe_port(os.environ.get("REMAP_PORT", "8000"))
        self._load_saved_settings()

    @staticmethod
    def _safe_port(value: Any) -> int:
        try:
            port = int(value)
        except (TypeError, ValueError):
            port = 8000
        return max(1, min(65535, port))

    def _load_saved_settings(self) -> None:
        if os.name != "nt" or not self._settings_path.exists():
            return
        try:
            data = json.loads(self._settings_path.read_text(encoding="utf-8"))
            protected = str(data.get("gemini_api_key_dpapi") or "")
            if protected:
                self._saved_api_key = _dpapi_decrypt(protected).strip()
            model = str(data.get("gemini_model") or "").strip()
            if model:
                self._model = model[:100]
            self._persist_enabled = bool(self._saved_api_key)
        except Exception:
            # Corrupted or copied DPAPI data cannot be decrypted under a different
            # Windows account. Ignore it and let the teacher enter a new key.
            self._saved_api_key = ""
            self._persist_enabled = False

    def _write_saved_settings(self, api_key: str, model: str) -> None:
        if os.name != "nt":
            raise RuntimeError("API 키 저장은 Windows 배포판에서만 지원합니다.")
        payload = {
            "gemini_api_key_dpapi": _dpapi_encrypt(api_key),
            "gemini_model": model,
            "saved_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        }
        temp_path = self._settings_path.with_suffix(".tmp")
        temp_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        temp_path.replace(self._settings_path)

    def set_network(self, public_host: str, port: int) -> None:
        with self._lock:
            host = str(public_host or "").strip()
            if host:
                self._public_host = host
            self._port = self._safe_port(port)
            self._port_explicit = True

    def public_host(self) -> str:
        with self._lock:
            if self._public_host:
                return self._public_host
        return get_ip_candidates()[0]

    def port(self) -> int:
        with self._lock:
            return self._port

    def port_is_explicit(self) -> bool:
        with self._lock:
            return self._port_explicit

    def api_key(self) -> str:
        with self._lock:
            if self._session_api_key:
                return self._session_api_key
            if self._saved_api_key:
                return self._saved_api_key
        return os.environ.get("GEMINI_API_KEY", "").strip()

    def api_key_source(self) -> str:
        with self._lock:
            if self._session_api_key:
                return "session"
            if self._saved_api_key:
                return "saved"
        return "environment" if os.environ.get("GEMINI_API_KEY", "").strip() else "none"

    def model(self) -> str:
        with self._lock:
            return self._model

    def set_model(self, model: str, save_if_persisted: bool = True) -> dict[str, Any]:
        clean_model = str(model or '').strip()[:100]
        if not clean_model:
            raise ValueError('사용할 Gemini 모델을 선택하세요.')
        with self._lock:
            self._model = clean_model
            if save_if_persisted and self._saved_api_key and os.name == 'nt':
                self._write_saved_settings(self._saved_api_key, clean_model)
            return self.ai_status()

    def update_available_models(self, models: list[dict[str, Any]]) -> None:
        clean: list[dict[str, Any]] = []
        seen: set[str] = set()
        for item in models:
            model_id = str(item.get('id') or '').strip()[:100]
            if not model_id or model_id in seen:
                continue
            seen.add(model_id)
            clean.append({
                'id': model_id,
                'display_name': str(item.get('display_name') or model_id).strip()[:128],
                'description': str(item.get('description') or '').strip()[:400],
                'version': str(item.get('version') or '').strip()[:40],
                'input_token_limit': int(item.get('input_token_limit') or 0),
                'output_token_limit': int(item.get('output_token_limit') or 0),
            })
        with self._lock:
            self._available_models = clean
            self._models_checked_at = time.time() if clean else 0.0

    def available_models(self, max_age_seconds: int = 3600) -> list[dict[str, Any]]:
        with self._lock:
            if not self._available_models:
                return []
            if max_age_seconds >= 0 and time.time() - self._models_checked_at > max_age_seconds:
                return []
            return [dict(item) for item in self._available_models]

    def configure_ai(self, api_key: str | None, model: str | None, persist: bool = False) -> dict[str, Any]:
        with self._lock:
            clean_model = str(self._model if model is None else model).strip()[:100]
            clean_key = None if api_key is None else str(api_key).strip()
            if clean_key:
                self._session_api_key = clean_key
            self._model = clean_model

            if persist:
                key_to_save = self._session_api_key or self._saved_api_key or os.environ.get("GEMINI_API_KEY", "").strip()
                if not key_to_save:
                    raise ValueError("저장할 Gemini API 키를 먼저 입력하세요.")
                self._write_saved_settings(key_to_save, clean_model)
                self._saved_api_key = key_to_save
                self._persist_enabled = True
            elif self._settings_path.exists() and os.name == "nt":
                # Leaving the box unchecked means the entered key is session-only.
                # Existing saved data is not silently deleted; use the explicit
                # clear action for that.
                self._persist_enabled = bool(self._saved_api_key)

            return self.ai_status()

    def clear_ai_key(self, clear_saved: bool = True) -> dict[str, Any]:
        with self._lock:
            self._session_api_key = ""
            if clear_saved:
                self._saved_api_key = ""
                self._persist_enabled = False
                try:
                    self._settings_path.unlink(missing_ok=True)
                except OSError:
                    pass
            return self.ai_status()


    def set_ai_tls_status(self, mode: str, label: str, error: str = "") -> None:
        with self._lock:
            self._ai_tls_mode = str(mode or "").strip()[:60]
            self._ai_tls_label = str(label or "").strip()[:100]
            self._ai_tls_last_error = str(error or "").strip()[:500]
            self._ai_tls_checked_at = time.time()

    def ai_tls_status(self) -> dict[str, Any]:
        with self._lock:
            return {
                "mode": self._ai_tls_mode,
                "label": self._ai_tls_label,
                "last_error": self._ai_tls_last_error,
                "checked_at": self._ai_tls_checked_at,
                "verification_enabled": True,
            }

    def ai_status(self) -> dict[str, Any]:
        source = self.api_key_source()
        return {
            "configured": bool(self.api_key()),
            "source": source,
            "source_label": {
                "session": "현재 실행 중 입력",
                "saved": "이 PC에 안전하게 저장",
                "environment": "서버 환경 변수",
                "none": "미설정",
            }.get(source, source),
            "model": self.model(),
            "persisted": bool(self._saved_api_key),
            "can_persist": os.name == "nt",
            "available_models_count": len(self.available_models(max_age_seconds=-1)),
            "models_checked_at": self._models_checked_at,
            "tls": self.ai_tls_status(),
        }

    def diagnostics_base(self) -> dict[str, Any]:
        host = self.public_host()
        port = self.port()
        return {
            "app": APP_NAME,
            "version": APP_VERSION,
            "runtime": self.mode,
            "runtime_label": {
                "developer": "설계자 Python",
                "windows": "Windows Local",
                "render": "Render Network",
            }.get(self.mode, self.mode),
            "started_at": self.started_at,
            "uptime_seconds": max(0, int(time.time() - self.started_at)),
            "public_host": host,
            "port": port,
            "student_base_url": f"http://{host}:{port}/" if self.mode != "render" else "",
            "data_dir": str(self.data_dir),
            "python": sys.version.split()[0],
            "ai": self.ai_status(),
            "ip_candidates": get_ip_candidates(),
        }


runtime_settings = RuntimeSettings()


def is_loopback_host(value: str | None) -> bool:
    host = str(value or "").strip().lower()
    return host in {"127.0.0.1", "::1", "localhost", "testclient"} or host.startswith("127.")


def local_settings_allowed(client_host: str | None) -> bool:
    return runtime_settings.mode in {"developer", "windows"} and is_loopback_host(client_host)
