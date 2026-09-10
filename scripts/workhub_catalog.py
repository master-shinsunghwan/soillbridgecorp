from pathlib import Path
import json,sqlite3
import workhub_catalog_products
from contextlib import contextmanager
PRODUCTS=Path(__file__).with_name('catalog_products.json')
TYPES={'temporary','sold-out','incoming','discontinued','notice'}
def products():
 return json.loads(PRODUCTS.read_text(encoding='utf8'))
@contextmanager
def connection(config):
 config.mkdir(parents=True,exist_ok=True)
 c=sqlite3.connect(config/'catalog_status.db',timeout=15)
 c.execute('CREATE TABLE IF NOT EXISTS issues (id TEXT PRIMARY KEY, type TEXT, restock TEXT, message TEXT)')
 try:
  with c:yield c
 finally:c.close()
def load(config):
 with connection(config) as c:
  return {i:{'type':t,'restock':r,'message':m} for i,t,r,m in c.execute('SELECT id,type,restock,message FROM issues')}
def save(config,payload):
 if not isinstance(payload,dict):raise ValueError('입력 형식이 올바르지 않습니다.')
 pid=payload.get('id');kind=payload.get('type','')
 if pid not in {p['id'] for p in workhub_catalog_products.products(config)}:raise ValueError('상품을 찾을 수 없습니다.')
 if kind and kind not in TYPES:raise ValueError('지원하지 않는 상태입니다.')
 restock=payload.get('restock','');message=payload.get('message','')
 if not isinstance(restock,str) or not isinstance(message,str) or len(restock)>100 or len(message)>300:raise ValueError('재입고 일정은 100자, 안내 문구는 300자 이내로 입력해 주세요.')
 with connection(config) as c:
  if not kind:c.execute('DELETE FROM issues WHERE id=?',(pid,))
  else:c.execute('INSERT OR REPLACE INTO issues VALUES (?,?,?,?)',(pid,kind,restock.strip() if kind=='temporary' else '',message.strip()))
 return load(config)
HTML=Path(__file__).with_name('catalog_editor.html').read_text(encoding='utf-8-sig')
