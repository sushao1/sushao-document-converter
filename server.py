import http.client
import io
import os
from pathlib import Path
import subprocess
import socket
import time
import sys
import tempfile
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from email.parser import BytesParser
from email.policy import default
import zipfile

LIMIT = 4 * 1024 * 1024
IMAGE_OUTPUT_LIMIT = 20 * 1024 * 1024
OFFICE = {'.docx': 'word/document.xml', '.xlsx': 'xl/workbook.xml', '.pptx': 'ppt/presentation.xml'}
DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
PDF_WORKER = threading.BoundedSemaphore(1)


def validate(body, content_type, pdf=False, image_format=False):
    if not content_type.startswith('multipart/form-data;'):
        raise ValueError('请选择 Office 文件。')
    message = BytesParser(policy=default).parsebytes(
        b'Content-Type: ' + content_type.encode('ascii') + b'\r\nMIME-Version: 1.0\r\n\r\n' + body)
    parts = list(message.iter_parts())
    files = [part for part in parts if part.get_param('name', header='content-disposition') == 'files']
    formats = [part for part in parts if part.get_param('name', header='content-disposition') == 'format']
    if len(files) != 1 or len(parts) != 1 + bool(image_format) or len(formats) != bool(image_format):
        raise ValueError('每次请选择一个文件。')
    filename = files[0].get_filename() or ''
    extension = Path(filename).suffix.lower()
    payload = files[0].get_payload(decode=True)
    if pdf:
        if extension != '.pdf' or not payload or not payload.startswith(b'%PDF-'):
            raise ValueError('请选择有效的 PDF 文件。')
        if image_format:
            value = formats[0].get_payload(decode=True).decode('ascii').lower()
            if value not in ('png', 'jpg'):
                raise ValueError('图片格式仅支持 PNG 或 JPG。')
            return payload, value
        return payload
    if extension not in OFFICE:
        raise ValueError('仅支持 DOCX、XLSX、PPTX。')
    try:
        with zipfile.ZipFile(io.BytesIO(payload)) as archive:
            entries = archive.infolist()
            if OFFICE[extension] not in archive.namelist():
                raise ValueError('文件内容与扩展名不匹配。')
            if len(entries) > 5000 or sum(item.file_size for item in entries) > 100 * 1024 * 1024:
                raise ValueError('文件解压后过大。')
    except zipfile.BadZipFile:
        raise ValueError('文件损坏或不是有效的 Office 文件。') from None


class Handler(BaseHTTPRequestHandler):
    def reply(self, status, body, content_type='text/plain; charset=utf-8'):
        self.send_response(status)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path == '/':
            self.reply(200, Path(__file__).with_name('index.html').read_bytes(), 'text/html; charset=utf-8')
        elif self.path == '/health':
            self.forward('GET', '/health')
        else:
            self.reply(404, b'Not found')

    def do_POST(self):
        if self.path not in ('/api/convert', '/api/pdf-to-word', '/api/pdf-to-images'):
            self.reply(404, b'Not found')
            return
        try:
            length = int(self.headers.get('Content-Length', '0'))
            if not 0 < length <= LIMIT:
                self.reply(413, '文件超过 4MB 或为空。'.encode())
                return
            body = self.rfile.read(length)
            if len(body) != length:
                raise ValueError('上传中断，请重新上传。')
            content_type = self.headers.get('Content-Type', '')
            images = self.path == '/api/pdf-to-images'
            payload = validate(body, content_type, pdf=self.path != '/api/convert', image_format=images)
        except (ValueError, UnicodeError) as error:
            self.reply(400, str(error).encode())
            return
        if self.path == '/api/pdf-to-word':
            self.pdf_to_word(payload)
        elif self.path == '/api/pdf-to-images':
            self.pdf_to_images(*payload)
        else:
            self.forward('POST', '/forms/libreoffice/convert', body, content_type)

    def pdf_to_word(self, payload):
        if not PDF_WORKER.acquire(blocking=False):
            self.reply(429, '已有 PDF 正在转换，请稍后重试。'.encode())
            return
        try:
            with tempfile.TemporaryDirectory(prefix='pdf-to-word-') as folder:
                source, output = Path(folder) / 'input.pdf', Path(folder) / 'output.docx'
                source.write_bytes(payload)
                process = subprocess.run([sys.executable, __file__, '--pdf-to-word', str(source), str(output)],
                                         capture_output=True, timeout=120)
                if process.returncode:
                    message = process.stdout.strip() if process.returncode == 2 else 'PDF 转换失败，请检查文件是否损坏。'.encode()
                    self.reply(400, message)
                elif output.stat().st_size > LIMIT:
                    self.reply(413, '转换结果超过 4MB，请缩小文件。'.encode())
                else:
                    self.reply(200, output.read_bytes(), DOCX_TYPE)
        except subprocess.TimeoutExpired:
            self.reply(504, '转换超时，请减少 PDF 页数后重试。'.encode())
        except OSError:
            self.reply(503, '转换服务暂不可用，请稍后重试。'.encode())
        finally:
            PDF_WORKER.release()

    def pdf_to_images(self, payload, image_format):
        if not PDF_WORKER.acquire(blocking=False):
            self.reply(429, '已有 PDF 正在转换，请稍后重试。'.encode())
            return
        try:
            with tempfile.TemporaryDirectory(prefix='pdf-to-images-') as folder:
                source, output = Path(folder) / 'input.pdf', Path(folder) / 'images.zip'
                source.write_bytes(payload)
                process = subprocess.run([sys.executable, __file__, '--pdf-to-images', str(source), str(output), image_format],
                                         capture_output=True, timeout=120)
                if process.returncode:
                    message = process.stdout.strip() if process.returncode == 2 else 'PDF 转图片失败，请检查文件是否损坏。'.encode()
                    self.reply(400, message)
                elif output.stat().st_size > IMAGE_OUTPUT_LIMIT:
                    self.reply(413, '图片压缩包超过 20MB，请减少 PDF 页数。'.encode())
                else:
                    self.reply(200, output.read_bytes(), 'application/zip')
        except subprocess.TimeoutExpired:
            self.reply(504, '转换超时，请减少 PDF 页数后重试。'.encode())
        except OSError:
            self.reply(503, '转换服务暂不可用，请稍后重试。'.encode())
        finally:
            PDF_WORKER.release()

    def forward(self, method, path, body=None, content_type=''):
        connection = http.client.HTTPConnection('127.0.0.1', 3001, timeout=120)
        try:
            connection.request(method, path, body, {'Content-Type': content_type})
            response = connection.getresponse()
            result = response.read(LIMIT + 1)
            if len(result) > LIMIT:
                self.reply(413, '转换结果超过 4MB，请缩小文件。'.encode())
            elif response.status != 200:
                self.reply(502, '转换失败，请检查文件是否损坏或受密码保护。'.encode())
            else:
                self.reply(200, result, response.getheader('Content-Type', 'application/pdf'))
        except (OSError, http.client.HTTPException):
            self.reply(503, '转换服务暂未就绪，请稍后重试。'.encode())
        finally:
            connection.close()


def convert_pdf(source, output):
    import pymupdf
    with pymupdf.open(source) as document:
        if document.needs_pass:
            raise ValueError('PDF 受密码保护，请先解除密码。')
        if not 0 < len(document) <= 30:
            raise ValueError('PDF 最多支持 30 页，请拆分后重试。')
        has_text = False
        for page in document:
            text = page.get_text().strip()
            if not text and page.get_images():
                raise ValueError('PDF 含扫描图片页，需要先进行 OCR 文字识别。')
            has_text = has_text or bool(text)
        if not has_text:
            raise ValueError('PDF 没有可转换的文字。')
    from pdf2docx import Converter
    converter = Converter(source)
    try:
        converter.convert(output, ignore_page_error=False)
    finally:
        converter.close()


def render_pdf_images(source, output, image_format):
    import pymupdf
    if image_format not in ('png', 'jpg'):
        raise ValueError('图片格式仅支持 PNG 或 JPG。')
    with pymupdf.open(source) as document, zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as archive:
        if document.needs_pass:
            raise ValueError('PDF 受密码保护，请先解除密码。')
        if not 0 < len(document) <= 30:
            raise ValueError('PDF 最多支持 30 页，请拆分后重试。')
        for number, page in enumerate(document, 1):
            if page.rect.width * page.rect.height * 4 > 40_000_000:
                raise ValueError('PDF 页面尺寸过大，请缩小后重试。')
            image = page.get_pixmap(matrix=pymupdf.Matrix(2, 2), alpha=False)
            data = image.tobytes('png' if image_format == 'png' else 'jpeg', jpg_quality=90)
            archive.writestr(f'page-{number:03d}.{image_format}', data)


if __name__ == '__main__' and sys.argv[1:2] in (['--pdf-to-word'], ['--pdf-to-images']):
    try:
        if sys.argv[1] == '--pdf-to-word':
            convert_pdf(sys.argv[2], sys.argv[3])
        else:
            render_pdf_images(sys.argv[2], sys.argv[3], sys.argv[4])
    except ValueError as error:
        print(str(error))
        sys.exit(2)
elif __name__ == '__main__':
    subprocess.Popen(['gotenberg', '--api-port=3001', '--api-bind-ip=127.0.0.1',
                      '--chromium-disable-routes', '--webhook-disable', '--api-disable-download-from'])
    for attempt in range(300):
        try:
            socket.create_connection(('127.0.0.1', 3001), timeout=1).close()
            break
        except OSError:
            time.sleep(0.1)
    else:
        raise RuntimeError('Gotenberg failed to start')
    ThreadingHTTPServer(('0.0.0.0', int(os.environ.get('PORT', '80'))), Handler).serve_forever()
