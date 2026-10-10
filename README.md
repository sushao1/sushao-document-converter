# 苏少文档格式转换

线上网页：https://sushao-document-converter.vercel.app/

- PDF 转可编辑 Word：支持含文字层的 PDF，最多 30 页。
- PDF 转图片：支持 PNG、JPG，多页按页打包为 ZIP，最多 30 页。
- Word、Excel、PowerPoint 转 PDF：支持 DOCX、XLSX、PPTX。
- 上传限 4MB；Word/PDF 输出限 4MB，图片压缩包限 20MB；加密 PDF 会明确提示。

PDF 转 Word 使用 [pdf2docx 0.5.13](https://github.com/ArtifexSoftware/pdf2docx)（MIT）与 [PyMuPDF 1.28.2](https://github.com/pymupdf/PyMuPDF)（AGPL/commercial）。Office 转 PDF 使用 [Gotenberg 8.37.0](https://github.com/gotenberg/gotenberg)（MIT）。各依赖保留其原有许可证；pdf2docx 的原厂已停止主动维护。

## 部署与测试

Vercel 自动识别 Dockerfile.vercel。GitHub main 分支提交会触发部署。

```sh
docker build -f Dockerfile.vercel -t sushao-document-converter .
docker run --rm -p 127.0.0.1:3100:80 sushao-document-converter
docker run --rm --entrypoint /opt/pdf-converter/bin/python -w /app -v "$PWD/test_server.py:/app/test_server.py:ro" sushao-document-converter -m unittest -q test_server.py
```

PDF 转 Word 接口为 `/api/pdf-to-word`，PDF 转图片接口为 `/api/pdf-to-images`（额外接收 `format=png|jpg`），Office 转 PDF 接口为 `/api/convert`，文件字段均为 `files`。临时转换文件在请求结束后清理。PDF 转换在独立进程执行，120 秒超时；每个实例同一时间接收一个 PDF 转换任务。

已测试可编辑文字、中文表格、原有三种 Office 格式以及错误文件处理。复杂排版、字体替代可能影响还原效果。
