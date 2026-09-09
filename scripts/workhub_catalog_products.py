"""Persistent product overrides; image and ZIP publication is committed with the revision."""
import base64,copy,hashlib,io,json,math,re,sqlite3,time,uuid,zipfile
from pathlib import Path
from PIL import Image
BASE=Path(__file__).with_name('catalog_full_products.json')
def connect(config):
 c=sqlite3.connect(config/'catalog_status.db',timeout=20)
 c.execute('CREATE TABLE IF NOT EXISTS product_edits (id TEXT PRIMARY KEY, data TEXT NOT NULL, revision INTEGER NOT NULL)')
 c.execute('CREATE TABLE IF NOT EXISTS product_history (id TEXT, revision INTEGER, actor TEXT, changed_at TEXT, before_data TEXT, after_data TEXT)')
 return c
def products(config):
 config.mkdir(parents=True,exist_ok=True)
 rows=json.loads(BASE.read_text(encoding='utf8'))
 c=connect(config)
 try:edits={pid:(json.loads(data),rev) for pid,data,rev in c.execute('SELECT id,data,revision FROM product_edits')}
 finally:c.close()
 for p in rows:
  data,revision=edits.get(p['id'],({},0));p.update(data);p['revision']=revision
 return rows
def asset(config,url):
 if url.startswith('/assets-managed/'):
  root=(config/'catalog_assets').resolve();name=url.removeprefix('/assets-managed/')
 elif url.startswith('/assets/'):
  root=Path('/catalog-assets/current').resolve();name=url.lstrip('/')
 else:raise ValueError('지원하지 않는 이미지 경로입니다.')
 path=(root/name).resolve()
 if not path.is_relative_to(root):raise ValueError('잘못된 이미지 경로입니다.')
 return path
def save(config,payload,actor):
 if not isinstance(payload,dict):raise ValueError('입력 형식이 올바르지 않습니다.')
 old=next((p for p in products(config) if p['id']==payload.get('id')),None)
 if not old:raise ValueError('상품을 찾을 수 없습니다.')
 if payload.get('revision')!=old['revision']:raise ValueError('다른 직원이 수정한 상품입니다. 새로고침 후 다시 확인해 주세요.')
 p=copy.deepcopy(old)
 for key,limit in [('spec',12000),('features',12000),('pack',100)]:
  v=payload.get(key)
  if not isinstance(v,str) or len(v)>limit or (key=='pack' and not v.strip()):raise ValueError('스펙 및 입수량 입력을 확인해 주세요.')
  p[key]=v.strip()
 price=payload.get('price');sort=payload.get('sortPrice')
 if isinstance(price,bool) or not isinstance(price,(str,int,float)) or (isinstance(price,str) and (not price.strip() or len(price)>1000)):raise ValueError('단가를 확인해 주세요.')
 if isinstance(sort,bool) or not isinstance(sort,(int,float)) or not math.isfinite(sort) or sort<0:raise ValueError('목록 기준 단가는 0 이상의 숫자로 입력해 주세요.')
 if isinstance(price,(int,float)):
  if not math.isfinite(price) or price<0:raise ValueError('단가는 0 이상이어야 합니다.')
  sort=price
 p.update(price=price,sortPrice=sort)
 uploads=payload.get('images',{})
 if not isinstance(uploads,dict) or set(uploads)-{'main','detail'}:raise ValueError('이미지 구분을 확인해 주세요.')
 generated=[]
 folder=config/'catalog_assets';folder.mkdir(exist_ok=True)
 try:
  for kind,data in uploads.items():
   if not isinstance(data,str) or len(data)>28000000:raise ValueError('이미지는 한 장당 20MB 이내로 등록해 주세요.')
   try:
    raw=base64.b64decode(data,validate=True)
    if len(raw)>20*1024*1024:raise ValueError()
    with Image.open(io.BytesIO(raw)) as im:
     fmt=im.format;w,h=im.size;im.verify()
    if fmt not in {'JPEG','PNG','WEBP'}:raise ValueError()
   except Exception:raise ValueError('정상적인 JPG, PNG, WEBP 이미지 파일을 선택해 주세요.')
   ext={'JPEG':'jpg','PNG':'png','WEBP':'webp'}[fmt];name=old['id']+'-'+kind+'-'+uuid.uuid4().hex+'.'+ext
   target=folder/name;target.write_bytes(raw);generated.append(target)
   entry={'url':'/assets-managed/'+name,'name':name,'width':w,'height':h}
   p[kind]=[entry]+p[kind][1:]
   if kind=='main':p['preview']=entry['url'];p['excelImage']=False
   zipname=old['id']+'-'+kind+'-'+uuid.uuid4().hex+'.zip';zpath=folder/zipname;generated.append(zpath)
   with zipfile.ZipFile(zpath,'w',compression=zipfile.ZIP_DEFLATED) as z:
    for i,f in enumerate(p[kind]):
     source=asset(config,f['url'])
     if not source.is_file():raise ValueError('기존 이미지 파일을 확인할 수 없습니다. 담당자에게 문의해 주세요.')
     z.write(source,arcname=f'{i+1:02d}_{Path(f["name"]).name}')
   p['zips'][kind]='/assets-managed/'+zipname
  revision=old['revision']+1;p['revision']=revision
  c=connect(config)
  try:
   c.execute('BEGIN IMMEDIATE')
   row=c.execute('SELECT revision FROM product_edits WHERE id=?',(old['id'],)).fetchone()
   if (row[0] if row else 0)!=old['revision']:raise ValueError('다른 직원이 수정했습니다. 새로고침 후 다시 저장해 주세요.')
   data=json.dumps(p,ensure_ascii=False)
   c.execute('INSERT OR REPLACE INTO product_edits VALUES (?,?,?)',(old['id'],data,revision))
   c.execute('INSERT INTO product_history VALUES (?,?,?,?,?,?)',(old['id'],revision,actor,time.strftime('%Y-%m-%dT%H:%M:%SZ',time.gmtime()),json.dumps(old,ensure_ascii=False),data));c.commit()
  finally:c.close()
  return p
 except Exception:
  for path in generated:path.unlink(missing_ok=True)
  raise
