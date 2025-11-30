# Docker 배포 가이드

SheetLink를 Docker로 배포하는 방법입니다.

## 🐳 Docker Hub에서 실행

### 빠른 시작

```bash
docker run -d \
  -p 5050:5050 \
  --name sheetlink \
  hyunjojung/sheetlink:latest
```

브라우저에서 `http://localhost:5050` 접속

### docker-compose 사용

```bash
# docker-compose.yml 다운로드
curl -O https://raw.githubusercontent.com/your-repo/ExcelLinkExtractor/master/docker-compose.yml

# 실행
docker-compose up -d

# 로그 확인
docker-compose logs -f

# 중지
docker-compose down
```

## ⚙️ 환경 변수 설정

```yaml
environment:
  ASPNETCORE_ENVIRONMENT: Production
  ExcelProcessing__MaxFileSizeMB: 10
  ExcelProcessing__MaxHeaderSearchRows: 10
  ExcelProcessing__MaxUrlLength: 2000
  ExcelProcessing__RateLimitPerMinute: 500
```

## 🔄 업데이트

```bash
# 최신 이미지 가져오기
docker pull hyunjojung/sheetlink:latest

# 재시작
docker-compose down
docker-compose up -d
```

## 🏗️ 로컬에서 빌드

```bash
# 이미지 빌드
docker build -t sheetlink:local .

# 실행
docker run -d -p 5050:5050 sheetlink:local
```

## 📊 Health Check

```bash
curl http://localhost:5050/health
```

## 🔍 문제 해결

### 컨테이너 로그 확인
```bash
docker logs sheetlink
```

### 컨테이너 상태 확인
```bash
docker ps -a
```

### 컨테이너 내부 접속
```bash
docker exec -it sheetlink /bin/bash
```

## 🚀 서버 배포 (Production)

### 1. 서버에 Docker 설치

```bash
# Ubuntu/Debian
curl -fsSL https://get.docker.com -o get-docker.sh
sudo sh get-docker.sh

# Docker Compose 설치
sudo apt-get update
sudo apt-get install docker-compose-plugin
```

### 2. docker-compose.yml 생성

```yaml
version: "3.9"
services:
  sheetlink:
    image: hyunjojung/sheetlink:latest
    ports:
      - "5050:5050"
    environment:
      ASPNETCORE_ENVIRONMENT: Production
      ExcelProcessing__RateLimitPerMinute: 500
    restart: unless-stopped
    logging:
      driver: "json-file"
      options:
        max-size: "10m"
        max-file: "3"
```

### 3. 실행

```bash
docker-compose up -d
```

### 4. Nginx 리버스 프록시 (선택사항)

```nginx
server {
    listen 80;
    server_name sheetlink.hyunjo.uk;

    location / {
        proxy_pass http://localhost:5050;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

## 📦 이미지 정보

- **이미지**: `hyunjojung/sheetlink`
- **태그**:
  - `latest` - 최신 master 브랜치
  - `v1.0.0` - 특정 버전
  - `master-sha` - 특정 커밋
- **Base 이미지**: `mcr.microsoft.com/dotnet/aspnet:10.0`
- **포트**: 5050
- **크기**: ~200MB (예상)

## 🔐 보안

- 컨테이너는 non-root 사용자로 실행
- 파일은 메모리에서만 처리 (디스크 저장 없음)
- Rate limiting 기본 적용 (500 req/min)
- HTTPS는 리버스 프록시에서 처리 권장

## 📝 참고 사항

- 파일 처리는 메모리에서만 수행되므로 볼륨 마운트 불필요
- 컨테이너 재시작 시 메트릭 데이터는 초기화됨
- 로그는 stdout/stderr로 출력 (docker logs로 확인)

---

**마지막 업데이트**: 2025-11-30
