import io
import unittest
import zipfile
from server import OFFICE, validate


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


if __name__ == '__main__':
    unittest.main()
