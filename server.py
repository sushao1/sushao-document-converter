import http.client
import io
import os
from pathlib import Path
import subprocess
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from email.parser import BytesParser
from email.policy import default
import zipfile

LIMIT = 4 * 1024 * 1024
OFFICE = {'.docx': 'word/document.xml', '.xlsx': 'xl/workbook.xml', '.pptx': 'ppt/presentation.xml'}


def validate(body, content_type):
    if not content_type.startswith('multipart/form-data;'):
        raise ValueError('请选择 Office 文件。')
    message = BytesParser(policy=default).parsebytes(
        b'Content-Type: ' + content_type.encode('ascii') + b'\r\nMIME-Version: 1.0\r\n\r\n' + body)
    parts = list(message.iter_parts())
    if len(parts) != 1 or parts[0].get_param('name', header='content-disposition') != 'files':
        raise ValueError('每次请选择一个文件。')
    filename = parts[0].get_filename() or ''
    extension = Path(filename).suffix.lower()
    if extension not in OFFICE:
        raise ValueError('仅支持 DOCX、XLSX、PPTX。')
    try:
        with zipfile.ZipFile(io.BytesIO(parts[0].get_payload(decode=True))) as archive:
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
        if self.path != '/api/convert':
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
            validate(body, content_type)
        except (ValueError, UnicodeError) as error:
            self.reply(400, str(error).encode())
            return
        self.forward('POST', '/forms/libreoffice/convert', body, content_type)

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


if __name__ == '__main__':
    subprocess.Popen(['gotenberg', '--api-port=3001', '--api-bind-ip=127.0.0.1',
                      '--chromium-disable-routes', '--webhook-disable', '--api-disable-download-from'])
    ThreadingHTTPServer(('0.0.0.0', int(os.environ.get('PORT', '80'))), Handler).serve_forever()
