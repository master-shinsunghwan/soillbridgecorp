import io
import sys
import tempfile
import unittest
import zipfile
import http.client
import threading
from http.server import ThreadingHTTPServer
from urllib.parse import quote
from unittest.mock import patch
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import workhub_catalog_documents as m


def office(kind):
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w') as z:
        z.writestr('[Content_Types].xml', '<Types/>')
        z.writestr('xl/workbook.xml' if kind == 'excel' else 'ppt/presentation.xml', '<document/>')
    return out.getvalue()


class DocumentTests(unittest.TestCase):
    def test_http_download_auth_and_upload(self):
        import workhub_delivery_app as app

        class Handler(app.WorkhubHandler):
            def current_user(self):
                return {'username': 'test-staff'} if self.headers.get('X-Test-User') else None

            def require_permission(self, user, permission, label):
                return bool(user)

        with tempfile.TemporaryDirectory() as d, patch.object(app, 'CONFIG_DIR', Path(d)):
            server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
            thread = threading.Thread(target=server.serve_forever, daemon=True)
            thread.start()
            conn = http.client.HTTPConnection(*server.server_address)
            try:
                conn.request('GET', '/api/catalog-document/excel')
                r = conn.getresponse()
                self.assertEqual(r.status, 302)
                self.assertTrue(r.getheader('Location').endswith('.xlsx'))
                r.read()
                conn.request('POST', '/api/catalog-document/excel', b'bad')
                r = conn.getresponse(); self.assertEqual(r.status, 401); r.read()
                conn.request('POST', '/api/catalog-document/excel', b'bad', {'X-Test-User': '1'})
                r = conn.getresponse(); self.assertEqual(r.status, 403); r.read()
                headers = {'X-Test-User': '1', 'X-Workhub-Catalog': '1', 'X-File-Name': quote('거래처.xlsx'), 'X-Document-Revision': '0'}
                conn.request('POST', '/api/catalog-document/excel', office('excel'), headers)
                r = conn.getresponse(); self.assertEqual(r.status, 200); r.read()
                conn.request('GET', '/api/catalog-document/excel')
                r = conn.getresponse()
                self.assertEqual(r.status, 200)
                self.assertIn(quote('거래처.xlsx'), r.getheader('Content-Disposition'))
                self.assertEqual(r.read(), office('excel'))
            finally:
                conn.close(); server.shutdown(); server.server_close(); thread.join()

    def test_publish_replace_conflict_and_history(self):
        with tempfile.TemporaryDirectory() as d:
            config = Path(d)
            self.assertEqual(m.documents(config), {})
            first = m.save(config, 'excel', '제안서.xlsx', office('excel'), 0, 'staff')['excel']
            self.assertEqual((config / 'catalog_documents' / first['file']).read_bytes(), office('excel'))
            with self.assertRaises(ValueError):
                m.save(config, 'excel', '경합.xlsx', office('excel'), 0, 'other')
            second = m.save(config, 'excel', '새 제안서.xlsx', office('excel'), 1, 'staff')['excel']
            self.assertEqual(second['revision'], 2)
            self.assertTrue((config / 'catalog_documents' / first['file']).is_file())
            ppt = m.save(config, 'ppt', '제안서.pptx', office('ppt'), 0, 'staff')['ppt']
            self.assertEqual(ppt['revision'], 1)
            self.assertEqual(m.documents(config)['excel'], second)
            c = m.connect(config)
            try:
                self.assertEqual(c.execute('SELECT count(*) FROM catalog_document_history').fetchone()[0], 3)
            finally:
                c.close()

    def test_invalid_uploads_do_not_publish(self):
        with tempfile.TemporaryDirectory() as d:
            config = Path(d)
            for kind, name, raw in [('excel', '../a.xlsx', office('excel')), ('excel', 'a.xlsx', office('ppt')),
                                    ('ppt', 'a.ppt', b'not office'), ('ppt', 'a.exe', b'bad'),
                                    ('excel', 'a.xlsx', b''), ('excel', 'a.xlsx', b'x' * (m.LIMIT + 1))]:
                with self.assertRaises(ValueError):
                    m.save(config, kind, name, raw, 0, 'staff')
            self.assertEqual(m.documents(config), {})


if __name__ == '__main__':
    unittest.main()
