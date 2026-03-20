비영리기관 전자결재 시스템

`https://gw.ktbizoffice.com/sub/service.html`의 서비스 소개 기준으로 핵심 기능을 추출해, 실제 동작 가능한 웹 전자결재 시스템을 구현했습니다.

## 1) 서비스 기능 분석 요약

 전자결재 관련 기능군은 아래와 같습니다.

- 업무포탈: 대시보드 기반 통합 업무 화면
- 전자결재: 기안, 결재선 지정, 승인/반려, 결재 진행 추적
- 공지사항/게시: 공지 공유 및 조직 커뮤니케이션
- 일정/자원관리: 개인/부서 일정과 자원(회의실 등) 운영
- API/연동: 외부 시스템(ERP/인사/근태 등) 연동 확장성
- 관리자 설정: 사용자/권한/운영 정책 관리

본 프로젝트는 위 기능군 중, 전자결재를 중심으로 **포털 + 결재 + 일정/공지 + 알림**을 통합 구현했습니다.

## 2) 구현 기능

### 전자결재 핵심
- 사용자 로그인/세션 인증
- 문서 기안(양식, 우선순위, 기한, 내용)
- 작성 도구 선택(`내장 편집기` / `Google Docs`)
- Google Docs 링크 또는 문서 ID 기반 기안
- 결재선(순차) 지정
- 임시저장/상신
- 결재자 승인/반려/코멘트
- 결재선 단계별 상태 추적
- 문서함 조회(필터/검색)
- 반려 사유 기록

### 포털/부가기능
- 업무포털 KPI(결재대기, 임시저장, 진행중, 완료, 미확인 알림, 7일 일정)
- 공지사항 조회/등록(관리자)
- 일정/자원 등록 및 조회
- 알림함 조회/읽음 처리
- 휴가 양식 승인 시 일정 자동 생성(결재-일정 연계)

## 3) 기술 스택

- Backend: Python 표준 라이브러리 (`http.server`, `sqlite3`)
- Frontend: Vanilla JS + HTML + CSS
- DB: SQLite (`data/approval.db` 자동 생성)
- 외부 패키지: 없음

## 4) 실행 방법

```bash
python server.py --host 127.0.0.1 --port 8080
```

브라우저에서 `http://127.0.0.1:8080` 접속

초기 계정:
- `admin / admin123!`
- `ceo / ceo123!`
- `kim / kim123!`
- `lee / lee123!`
- `park / park123!`

## 5) 파일 구조

- `server.py`: API + 정적 파일 서버 + DB 스키마/시드
- `static/index.html`: UI 레이아웃
- `static/styles.css`: 스타일
- `static/app.js`: 화면 로직 및 API 연동
- `data/approval.db`: 런타임 생성 DB

## 6) Google Docs 기안 방식

- 기안작성 화면에서 `작성 도구`를 `Google Docs`로 선택
- `Google Docs 링크 또는 문서 ID` 입력 후 저장/상신
- 입력 즉시 작성 화면 하단에 `Google Docs 미리보기 패널`이 표시됨
- `Google Docs에서 편집` 버튼으로 원문을 새 탭에서 바로 열 수 있음
- `Google 연결` 후 `Drive Picker 문서 선택`, `새 Google 문서 생성` 지원
- 시스템은 문서 ID를 추출해 표준 URL로 저장:
  - 편집: `https://docs.google.com/document/d/{doc_id}/edit`
  - 미리보기: `https://docs.google.com/document/d/{doc_id}/preview`
- 상세 패널에서 원문 열기/미리보기를 제공

주의:
- 결재자 계정이 문서에 접근 가능하도록 Google 공유 권한을 미리 설정해야 합니다.

### OAuth + Drive Picker 설정

환경변수 설정 후 서버를 재시작하세요.

- `GOOGLE_OAUTH_CLIENT_ID`: Google Cloud 웹 OAuth Client ID
- `GOOGLE_API_KEY`: Picker용 API Key
- `GOOGLE_CLOUD_APP_ID`: (선택) Google Cloud Project Number
- `GOOGLE_DRIVE_SCOPE`: (선택) OAuth Scope
  - 기본값: `https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/documents.readonly`

OAuth 클라이언트의 Authorized JavaScript origins에 현재 접속 주소를 등록해야 합니다.

- 예시: `http://127.0.0.1:8080`, `http://localhost:8080`

## 7) API 개요

- 인증: `POST /api/auth/login`, `POST /api/auth/logout`, `GET /api/auth/me`
- 연동설정: `GET /api/integrations/google/config`
- 사용자: `GET /api/users`
- 대시보드: `GET /api/dashboard`
- 문서: `GET/POST /api/documents`, `GET /api/documents/{id}`
  - `POST /api/documents` 확장 필드: `editor_provider`, `external_doc_url`, `external_doc_id`
- 결재: `POST /api/documents/{id}/submit`, `POST /api/documents/{id}/actions`
- 결재대기: `GET /api/approvals/pending`
- 공지: `GET/POST /api/notices`
- 일정: `GET/POST /api/schedules`
- 알림: `GET /api/notifications`, `POST /api/notifications/read`

## 8) 확장 포인트 (KT BizOffice 연동 방향)

- ERP/인사/근태 API 연결 어댑터 계층 추가
- 결재 양식 디자이너(동적 폼 스키마) 도입
- 조직도/부서 권한(열람 범위) 상세화
- 전자계약/메신저/메일 모듈 연계
- 감사 로그/감사 리포트 및 보존 정책 강화


## 9) 외부 배포 (Oracle Always Free 권장)

외부 사용자(50~100명) 접속용 배포는 아래 문서를 그대로 따라 진행하세요.

- `deploy/oracle/README_ORACLE_FREE.md`
- `deploy/oracle/eapproval.service`
- `deploy/oracle/Caddyfile`
- `deploy/oracle/.env.example`
- `deploy/oracle/backup_approval_db.sh`
