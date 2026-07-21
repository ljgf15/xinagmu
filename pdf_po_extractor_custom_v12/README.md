# 数百 PDF 批量识别版本

主要改动：
- 默认单次最多 500 个 PDF，可通过环境变量 `MAX_FILE_COUNT` 调整。
- 上传采用分块流式写盘，不再把所有文件一次性读入内存。
- PDF 使用有限并发解析，默认并发数为 `min(4, CPU 核心数)`。
- 前端只展示前 20 个文件，避免选择数百文件后页面卡顿。
- 默认单文件限制 25MB，总上传限制 5120MB，可用环境变量调整。

可配置环境变量：
- `MAX_FILE_COUNT=500`
- `MAX_UPLOAD_SIZE_MB=25`
- `MAX_TOTAL_UPLOAD_SIZE_MB=5120`
- `UPLOAD_CHUNK_SIZE_MB=1`
- `PARSE_CONCURRENCY=4`
- `PREVIEW_ROW_LIMIT=500`

部署时还需确认反向代理限制，例如 Nginx：
- `client_max_body_size 5g;`
- `proxy_read_timeout 3600s;`
- `proxy_send_timeout 3600s;`

如果服务器内存较小，建议把 `PARSE_CONCURRENCY` 设为 2。
