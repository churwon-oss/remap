ReMap Render 업로드 파일
========================

이 폴더의 내용을 GitHub 저장소 또는 Render 배포 저장소의 최상위에 업로드합니다.

구성
- app: ReMap 웹 애플리케이션
- requirements.txt: Python 패키지 목록
- render.yaml: Render 배포 설정

배포 후 경로
- 학생용: https://remap-oxrv.onrender.com
- 교사용: https://remap-oxrv.onrender.com/teacher
- 방 생성 후 학생 주소: https://remap-oxrv.onrender.com/?code=방코드

온라인 주소 생성 기준
Render 버전은 로컬 IP가 아니라 REMAP_PUBLIC_BASE_URL과 현재 공개 도메인을 기준으로 학생 주소와 QR 코드를 생성합니다. 10.x.x.x 형태의 서버 내부 IP는 외부 접속 주소로 표시하지 않습니다.
이번 배포본은 기존 Render 서비스에서 환경변수 설정이 누락된 경우에도 .onrender.com 요청 주소를 자동 감지하며, 교사용 화면에는 브라우저에서 실제 접속한 공개 도메인을 우선 표시합니다.

test용 web버전의 목적
이 온라인판은 ReMap의 향후 온라인 확장 가능성을 확인하기 위해 무료 서버에서 시험한 참고용 버전입니다. 무료 서버가 대기 상태이면 최초 접속 시 약 1분이 걸릴 수 있으며, 참여 인원이 늘면 캐릭터 이동과 실시간 위치 반영이 느려질 수 있습니다. 채점 및 전체 기능 확인은 USB의 program/index.exe로 진행해 주시기 바랍니다.
