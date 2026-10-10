import io
import unittest
import zipfile
import tempfile
from pathlib import Path
from server import OFFICE, validate, convert_pdf, render_pdf_images


class UploadValidation(unittest.TestCase):
    def test_office_formats(self):
        for extension, entry in OFFICE.items():
            with self.subTest(extension=extension):
                archive = io.BytesIO()
                with zipfile.ZipFile(archive, 'w') as output:
                    output.writestr(entry, '<document/>')
                body = (f'--test\r\nContent-Disposition: form-data; name="files"; filename="test{extension}"\r\n\r\n'.encode()
                        + archive.getvalue() + b'\r\n--test--\r\n')
                validate(body, 'multipart/form-data; boundary=test')

    def test_invalid_upload(self):
        with self.assertRaises(ValueError):
            validate(b'bad', 'text/plain')
        with self.assertRaises(ValueError):
            validate(b'--test\r\nContent-Disposition: form-data; name="downloadFrom"\r\n\r\nhttp://localhost\r\n--test--\r\n',
                     'multipart/form-data; boundary=test')

    def test_pdf_upload(self):
        body = b'--test\r\nContent-Disposition: form-data; name="files"; filename="input.pdf"\r\n\r\n%PDF-1.7\r\n--test--\r\n'
        self.assertTrue(validate(body, 'multipart/form-data; boundary=test', pdf=True).startswith(b'%PDF-'))
        with self.assertRaises(ValueError):
            validate(body.replace(b'%PDF-', b'wrong'), 'multipart/form-data; boundary=test', pdf=True)

    def test_pdf_image_upload(self):
        body = b'--test\r\nContent-Disposition: form-data; name="files"; filename="input.pdf"\r\n\r\n%PDF-1.7\r\n--test\r\nContent-Disposition: form-data; name="format"\r\n\r\njpg\r\n--test--\r\n'
        payload, image_format = validate(body, 'multipart/form-data; boundary=test', pdf=True, image_format=True)
        self.assertTrue(payload.startswith(b'%PDF-'))
        self.assertEqual(image_format, 'jpg')

    def test_pdf_text_is_editable(self):
        import pymupdf
        with tempfile.TemporaryDirectory() as folder:
            source, output = Path(folder) / 'input.pdf', Path(folder) / 'output.docx'
            with pymupdf.open() as document:
                document.new_page().insert_text((72, 72), 'Editable PDF text')
                document.save(source)
            convert_pdf(str(source), str(output))
            with zipfile.ZipFile(output) as archive:
                self.assertIn(b'Editable PDF text', archive.read('word/document.xml'))

    def test_unsupported_pdfs_are_rejected(self):
        import pymupdf
        with tempfile.TemporaryDirectory() as folder:
            source, output = Path(folder) / 'input.pdf', Path(folder) / 'output.docx'
            with pymupdf.open() as document:
                page = document.new_page()
                page.insert_text((72, 72), 'Scanned text')
                image = page.get_pixmap().tobytes('png')
            with pymupdf.open() as document:
                page = document.new_page()
                page.insert_image(page.rect, stream=image)
                document.save(source)
            with self.assertRaisesRegex(ValueError, 'OCR'):
                convert_pdf(str(source), str(output))
            source.unlink()
            with pymupdf.open() as document:
                document.new_page().insert_text((72, 72), 'Protected')
                document.save(source, encryption=pymupdf.PDF_ENCRYPT_AES_256,
                              owner_pw='test-owner', user_pw='test-user')
            with self.assertRaisesRegex(ValueError, '密码'):
                convert_pdf(str(source), str(output))
            source.unlink()
            with pymupdf.open() as document:
                for _ in range(31):
                    document.new_page()
                document.save(source)
            with self.assertRaisesRegex(ValueError, '30'):
                convert_pdf(str(source), str(output))

    def test_pdf_pages_render_as_png_and_jpg(self):
        import pymupdf
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / 'input.pdf'
            with pymupdf.open() as document:
                for text in ('First page', 'Second page'):
                    document.new_page().insert_text((72, 72), text)
                document.save(source)
            for image_format, signature in (('png', b'\x89PNG'), ('jpg', b'\xff\xd8')):
                output = Path(folder) / f'{image_format}.zip'
                render_pdf_images(str(source), str(output), image_format)
                with zipfile.ZipFile(output) as archive:
                    self.assertEqual(archive.namelist(), [f'page-001.{image_format}', f'page-002.{image_format}'])
                    self.assertTrue(archive.read(archive.namelist()[0]).startswith(signature))


if __name__ == '__main__':
    unittest.main()
