# SheetLink - Excel Hyperlink Tool

Free, open-source web tool for extracting and merging hyperlinks in Excel files.

**Live Demo:** [sheetlink.hyunjo.uk](https://sheetlink.hyunjo.uk)

---

## Quick Start

```bash
docker run -d -p 5050:5050 hyunjojung/sheetlink:latest
```

Then visit `http://localhost:5050`

---

## Features

- **Extract Links**: Extract hyperlinks from Excel cells and export URLs
- **Merge Links**: Combine Title + URL columns into clickable hyperlinks
- **Privacy-First**: All processing in-memory, no file storage
- **Excel Support**: .xlsx and .xls formats
- **Free & Open Source**: Apache 2.0 license

---

## Using docker-compose

Create `docker-compose.yml`:

```yaml
version: "3.9"
services:
  sheetlink:
    image: hyunjojung/sheetlink:latest
    ports:
      - "5050:5050"
    environment:
      ASPNETCORE_ENVIRONMENT: Production
      ASPNETCORE_URLS: "http://+:5050"
      ExcelProcessing__MaxFileSizeMB: 10
      ExcelProcessing__RateLimitPerMinute: 100
    restart: unless-stopped
```

Then run:

```bash
docker-compose up -d
```

---

## Configuration

Environment variables:

| Variable | Default | Description |
|----------|---------|-------------|
| `ASPNETCORE_URLS` | `http://+:5050` | Bind address |
| `ExcelProcessing__MaxFileSizeMB` | `10` | Max upload size (MB) |
| `ExcelProcessing__MaxHeaderSearchRows` | `10` | Header search depth |
| `ExcelProcessing__MaxUrlLength` | `2000` | Max URL length |
| `ExcelProcessing__RateLimitPerMinute` | `100` | Rate limit per IP |

---

## Health & Metrics

- **Health Check**: `GET /health` - Returns "Healthy" when app is running
- **Prometheus Metrics**: `GET /metrics` - Prometheus exposition format

Example health check in docker-compose:

```yaml
services:
  sheetlink:
    image: hyunjojung/sheetlink:latest
    healthcheck:
      test: ["CMD", "curl", "-f", "http://localhost:5050/health"]
      interval: 30s
      timeout: 10s
      retries: 3
      start_period: 40s
```

---

## Architecture

**Tech Stack:**
- ASP.NET Core 10.0 (Blazor Server)
- DocumentFormat.OpenXml
- prometheus-net

**Image Details:**
- Base: `mcr.microsoft.com/dotnet/aspnet:10.0`
- Size: ~256 MB
- Architecture: linux/amd64
- Multi-stage build optimized

---

## Supported Tags

- `latest` - Latest stable release from master branch
- `v*.*.*` - Specific version tags
- `master-<sha>` - Specific commit from master

---

## Volumes

No volumes required. All processing is done in-memory for security.

If you need persistent data protection keys (for multi-instance deployments):

```yaml
services:
  sheetlink:
    image: hyunjojung/sheetlink:latest
    volumes:
      - dataprotection-keys:/root/.aspnet/DataProtection-Keys
volumes:
  dataprotection-keys:
```

---

## Security

- **Input Validation**: File size, format, and content validation
- **URL Sanitization**: Only http/https/mailto schemes allowed
- **Rate Limiting**: Configurable per-IP rate limits
- **No Storage**: Files processed in-memory only

---

## Examples

### Basic Usage

```bash
docker run -d \
  --name sheetlink \
  -p 5050:5050 \
  hyunjojung/sheetlink:latest
```

### Production with Custom Settings

```bash
docker run -d \
  --name sheetlink \
  -p 5050:5050 \
  -e ExcelProcessing__MaxFileSizeMB=20 \
  -e ExcelProcessing__RateLimitPerMinute=500 \
  --restart unless-stopped \
  hyunjojung/sheetlink:latest
```

### With Prometheus Monitoring

```bash
# Start SheetLink
docker run -d --name sheetlink -p 5050:5050 hyunjojung/sheetlink:latest

# Scrape metrics
curl http://localhost:5050/metrics
```

---

## Support

- **GitHub**: [github.com/HyunjoJung/sheetlink](https://github.com/HyunjoJung/sheetlink)
- **Issues**: [github.com/HyunjoJung/sheetlink/issues](https://github.com/HyunjoJung/sheetlink/issues)
- **Documentation**: See repository docs/ folder

---

## License

Apache License 2.0 - see [LICENSE](https://github.com/HyunjoJung/sheetlink/blob/master/LICENSE)

**Disclaimer**: SheetLink is an independent open-source project and is not affiliated with Microsoft Corporation. "Excel" and "Microsoft Excel" are trademarks of Microsoft Corporation.

---

**Built with ❤️ by [HyunjoJung](https://github.com/HyunjoJung)**
