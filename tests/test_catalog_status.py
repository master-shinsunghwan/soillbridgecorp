import sys,tempfile,unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
import workhub_catalog as catalog
class CatalogTests(unittest.TestCase):
 def test_save_change_clear_and_validation(self):
  with tempfile.TemporaryDirectory() as d:
   root=Path(d)
   self.assertEqual(catalog.load(root),{})
   result=catalog.save(root,{'id':'SB-EX-006','type':'temporary','restock':'9월 말 예정','message':'입고 지연'})
   self.assertEqual(result['SB-EX-006']['restock'],'9월 말 예정')
   catalog.save(root,{'id':'SB-EX-007','type':'notice','message':'포장 변경'})
   catalog.save(root,{'id':'SB-EX-006','type':''})
   self.assertNotIn('SB-EX-006',catalog.load(root))
   self.assertIn('SB-EX-007',catalog.load(root))
   for payload in [{'id':'invalid','type':'temporary'},{'id':'SB-EX-006','type':'fake'},{'id':'SB-EX-006','type':'temporary','restock':'x'*101}]:
    with self.assertRaises(ValueError):catalog.save(root,payload)
if __name__=='__main__':unittest.main()
