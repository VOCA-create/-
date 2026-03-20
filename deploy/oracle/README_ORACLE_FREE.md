# Oracle Cloud Always Free 배포 가이드 (50~100명 조직 기준)

## 0. 개요
- 앱: `server.py` (Python stdlib) + SQLite
- 공개 방식: Oracle Free VM + Caddy(HTTPS) + systemd
- 권장 도메인: `approval.yourdomain.com`

## 1. VM 준비
1) Oracle Cloud에서 Always Free Ubuntu VM 생성
2) Security List(Ingress)
- TCP 22 (SSH)
- TCP 80 (HTTP)
- TCP 443 (HTTPS)

## 2. 서버 패키지 설치
```bash
sudo apt update
sudo apt install -y python3 python3-venv git caddy
```

## 3. 코드 배치
```bash
sudo mkdir -p /opt/eapproval
sudo chown -R ubuntu:ubuntu /opt/eapproval
cd /opt/eapproval
# 방법 A: git clone
# 방법 B: 로컬 프로젝트 업로드(scp/zip)
```

## 4. 환경변수 파일 생성
```bash
cp deploy/oracle/.env.example /opt/eapproval/.env
nano /opt/eapproval/.env
```
아래 값 입력:
- `GOOGLE_OAUTH_CLIENT_ID`
- `GOOGLE_API_KEY`
- `GOOGLE_CLOUD_APP_ID`

## 5. systemd 서비스 등록
```bash
sudo cp deploy/oracle/eapproval.service /etc/systemd/system/eapproval.service
sudo systemctl daemon-reload
sudo systemctl enable eapproval
sudo systemctl start eapproval
sudo systemctl status eapproval --no-pager
```

## 6. Caddy(HTTPS) 설정
```bash
sudo cp deploy/oracle/Caddyfile /etc/caddy/Caddyfile
# Caddyfile 안의 도메인(example.yourdomain.com)을 실도메인으로 변경
sudo systemctl reload caddy
sudo systemctl status caddy --no-pager
```

## 7. Google OAuth 콘솔 설정
Authorized JavaScript origins에 아래 등록:
- `https://approval.yourdomain.com`

## 8. 백업(권장)
```bash
chmod +x /opt/eapproval/deploy/oracle/backup_approval_db.sh
(crontab -l 2>/dev/null; echo "0 3 * * * /opt/eapproval/deploy/oracle/backup_approval_db.sh >> /opt/eapproval/backups/backup.log 2>&1") | crontab -
```

## 9. 점검
- 헬스체크: `https://approval.yourdomain.com/api/health`
- 로그인 화면에서 Google 버튼 노출 확인
- 관리자/일반 계정 로그인, 문서 생성/결재 플로우 확인

## 10. 운영 팁
- DB 파일: `/opt/eapproval/data/approval.db`
- 로그 확인:
```bash
journalctl -u eapproval -f
journalctl -u caddy -f
```
