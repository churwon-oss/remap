ReMap v4.2.1 - Render 업로드 파일
=====================================

이 폴더의 내용을 GitHub 저장소 루트 또는 Render 배포용 저장소 루트에 올립니다.

필수 구성
- app/                 FastAPI 애플리케이션과 정적 자료
- requirements.txt     Python 의존성
- render.yaml          Render Blueprint 설정
- VERSION              버전 정보

Render 배포 방법
1. 이 폴더의 파일을 저장소 루트에 업로드합니다.
2. Render에서 Blueprint 또는 Web Service를 생성하고 해당 저장소를 연결합니다.
3. render.yaml 사용 시 buildCommand와 startCommand가 자동 적용됩니다.
4. 배포 완료 후 교사용 주소 /teacher에서 방을 만들고 학생용 기본 주소로 접속합니다.

현재 온라인 시험판 주소
- 학생용: https://remap-oxrv.onrender.com
- 교사용: https://remap-oxrv.onrender.com/teacher

무료 서버 주의사항
- 일정 시간 사용하지 않으면 절전 상태가 될 수 있습니다.
- 최초 접속 시 서버를 깨우는 데 약 1분이 걸릴 수 있습니다.
- 자체 시험에서는 약 5명 내외가 비교적 원활했으나 무료 서버 자원과 네트워크 상황에 따라 지연될 수 있습니다.
- 정식 심사와 실제 학급 운영은 USB program 폴더의 로컬 index.exe 사용을 권장합니다.
