import base64,io,json,sys,tempfile,unittest,zipfile
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from PIL import Image
import workhub_catalog_products as m
class ProductEditTests(unittest.TestCase):
 def test_new_product_publication_edit_and_stock(self):
  import workhub_catalog as stock
  with tempfile.TemporaryDirectory() as d:
   config=Path(d);before=m.products(config)
   out=io.BytesIO();Image.new('RGB',(8,8)).save(out,format='PNG')
   body=dict(name='새 상품',category='생활용품',shipping='무료배송',tax='VAT 포함',spec='구성 1개',features='',price=5000,sortPrice=5000,pack='8EA',images={'main':base64.b64encode(out.getvalue()).decode()})
   with self.assertRaises(ValueError):m.save(config,dict(body,images={}), 'tester',create=True)
   p=m.save(config,body,'tester',create=True)
   rows=m.products(config)
   self.assertEqual(len(rows),len(before)+1)
   self.assertEqual(rows[-1]['id'],p['id'])
   self.assertEqual(rows[:-1],before)
   self.assertEqual(m.asset(config,p['preview']).read_bytes(),out.getvalue())
   edited=m.save(config,dict(body,id=p['id'],revision=1,pack='12EA',images={}), 'tester')
   self.assertEqual(edited['pack'],'12EA')
   self.assertEqual(stock.save(config,dict(id=p['id'],type='temporary',restock='2026-10-01'))[p['id']]['restock'],'2026-10-01')
 def test_roundtrip_conflict_validation_images_and_history(self):
  with tempfile.TemporaryDirectory() as d:
   config=Path(d);oldbase=m.BASE;baseline=config/'base.json';p=json.loads(oldbase.read_text(encoding='utf8'))[0];p['main']=[];p['detail']=[];baseline.write_text(json.dumps([p]),encoding='utf8');m.BASE=baseline
   try:
    body={k:p[k] for k in ['id','spec','features','price','pack','sortPrice']};body.update(revision=0,price=12345,pack='12EA')
    saved=m.save(config,body,'tester');self.assertEqual(saved['sortPrice'],12345);self.assertEqual(m.products(config)[0]['pack'],'12EA')
    with self.assertRaises(ValueError):m.save(config,body,'other')
    body.update(revision=1,price=-1)
    with self.assertRaises(ValueError):m.save(config,body,'tester')
    stream=io.BytesIO();Image.new('RGB',(12,12)).save(stream,format='PNG');raw=stream.getvalue();body.update(price=12345,images={'main':base64.b64encode(raw).decode()})
    saved=m.save(config,body,'tester');self.assertEqual(saved['preview'],saved['main'][0]['url']);self.assertEqual(m.asset(config,saved['preview']).read_bytes(),raw)
    with zipfile.ZipFile(m.asset(config,saved['zips']['main'])) as z:self.assertEqual(z.read(z.namelist()[0]),raw)
    c=m.connect(config);self.assertEqual(c.execute('SELECT count(*) FROM product_history').fetchone()[0],2);c.close()
   finally:m.BASE=oldbase
if __name__=='__main__':unittest.main()
