import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import workhub_delivery_app as app


class ShipmentTimeTests(unittest.TestCase):
    def test_time_roundtrip_sort_and_optional_time(self):
        with tempfile.TemporaryDirectory() as d:
            db = Path(d) / 'test.db'
            with sqlite3.connect(db) as c:
                fields = ','.join(f'{field} TEXT' for field in app.IMPORT_SHIPMENT_FIELDS)
                c.execute(f'CREATE TABLE import_shipments (id INTEGER PRIMARY KEY, created_at TEXT, updated_at TEXT, completed_at TEXT, {fields})')
            c.close()
            with patch.object(app, 'DB_PATH', db), patch.object(app, 'init_db'):
                late = app.save_import_shipment({'item': 'late', 'warehouse_due_date': '2026-09-10', 'warehouse_due_time': '14:30'})
                early = app.save_import_shipment({'item': 'early', 'warehouse_due_date': '2026-09-10', 'warehouse_due_time': '09:00'})
                no_time = app.save_import_shipment({'item': 'unknown', 'warehouse_due_date': '2026-09-10'})
                rows = app.list_import_shipments()
                self.assertEqual([r['id'] for r in rows], [early, late, no_time])
                self.assertEqual(rows[1]['warehouse_due_time'], '14:30')
                app.save_import_shipment({'id': late, 'item': 'late', 'warehouse_due_date': '2026-09-11', 'warehouse_due_time': '16:45'})
                self.assertEqual(app.list_import_shipments()[-1]['warehouse_due_time'], '16:45')
                for payload in [{'warehouse_due_date': '2026-09-10', 'warehouse_due_time': '25:00'}, {'warehouse_due_time': '12:00'}]:
                    with self.assertRaises(ValueError):
                        app.save_import_shipment(payload)


if __name__ == '__main__':
    unittest.main()
