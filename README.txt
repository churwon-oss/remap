ReMap Render 전용 업로드 파일
============================

이 폴더의 내용을 GitHub 저장소 또는 Render 배포 저장소의 최상위에 업로드합니다.
로컬 Windows 실행형과 주소 생성 방식이 다른 Render 전용 구성입니다.

구성
- app: ReMap 웹 애플리케이션
- requirements.txt: Python 패키지 목록
- render.yaml: Render 배포 설정

배포 후 경로
- 학생용: https://remap-oxrv.onrender.com
- 교사용: https://remap-oxrv.onrender.com/teacher
- 방 생성 후 학생 주소: https://remap-oxrv.onrender.com/?code=방코드

온라인 주소 생성 기준
- 학생 주소, QR 코드와 주소 복사는 고정 공개 도메인 https://remap-oxrv.onrender.com을 사용합니다.
- Render 내부의 10.x.x.x 또는 172.x.x.x 주소는 외부 접속 주소로 표시하지 않습니다.
- REMAP_PUBLIC_BASE_URL 환경변수가 올바른 onrender.com HTTPS 주소일 때만 해당 값을 사용하고, 누락되거나 잘못된 경우 위 공개 도메인으로 돌아갑니다.

재배포 확인
1. 이 업로드본으로 저장소 파일을 교체합니다.
2. Render에서 새 배포가 완료될 때까지 기다립니다.
3. 교사용 페이지에서 방을 생성합니다.
4. 학생 주소와 QR 코드가 https://remap-oxrv.onrender.com/?code=방코드 형식인지 확인합니다.
5. 이전 IP 주소가 계속 보이면 브라우저 새로고침과 Render 배포 커밋을 확인합니다.

test용 web버전의 목적
이 온라인판은 ReMap의 향후 온라인 확장 가능성을 확인하기 위해 무료 서버에서 시험한 참고용 버전입니다. 무료 서버가 대기 상태이면 최초 접속 시 약 1분이 걸릴 수 있으며, 참여 인원이 늘면 캐릭터 이동과 실시간 위치 반영이 느려질 수 있습니다. 채점 및 전체 기능 확인은 USB의 program/index.exe로 진행해 주시기 바랍니다.
