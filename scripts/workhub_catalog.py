from pathlib import Path
import json,sqlite3
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
 if pid not in {p['id'] for p in products()}:raise ValueError('상품을 찾을 수 없습니다.')
 if kind and kind not in TYPES:raise ValueError('지원하지 않는 상태입니다.')
 restock=payload.get('restock','');message=payload.get('message','')
 if not isinstance(restock,str) or not isinstance(message,str) or len(restock)>100 or len(message)>300:raise ValueError('재입고 일정은 100자, 안내 문구는 300자 이내로 입력해 주세요.')
 with connection(config) as c:
  if not kind:c.execute('DELETE FROM issues WHERE id=?',(pid,))
  else:c.execute('INSERT OR REPLACE INTO issues VALUES (?,?,?,?)',(pid,kind,restock.strip() if kind=='temporary' else '',message.strip()))
 return load(config)
HTML='''<!doctype html><html lang="ko"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>상품 제안서 상태 관리 · 소일브릿지</title><style>
body{font:15px/1.6 system-ui;margin:0;background:#f4f6f3;color:#1d362a}main{max-width:1000px;margin:auto;padding:28px 20px}a{color:#205540}h1{font-size:25px}input,select,button{font:inherit;padding:10px;border:1px solid #bac8bd;border-radius:6px;box-sizing:border-box}button{background:#1c503a;color:white;cursor:pointer}button:disabled{opacity:.5}#search{width:100%;margin:16px 0}.item{background:white;border:1px solid #dce3dc;border-radius:8px;padding:18px;margin-bottom:14px}.fields{display:grid;grid-template-columns:160px 1fr;gap:12px}label{display:flex;flex-direction:column;gap:5px}.wide{grid-column:1/-1}.item h2{font-size:16px;margin:0 0 12px}.save{margin-top:12px}.feedback{margin-left:12px}#status{font-weight:600}@media(max-width:550px){.fields{grid-template-columns:1fr}.wide{grid-column:auto}}</style><main><a href="/">← 업무 자동화로 돌아가기</a><h1>상품 제안서 상태 관리</h1><p>저장하면 공개 상품 제안서에 반영됩니다. 일시 품절은 재입고 일정을 함께 입력하세요. 일정이 없으면 ‘일정 미정’으로 표시됩니다.</p><a href="/catalog/" target="_blank" rel="noopener">상품 제안서 확인 ↗</a><input id="search" aria-label="상품 검색" placeholder="상품명 또는 번호 검색"><p id="status" role="status">불러오는 중…</p><div id="items"></div></main><script>
let products=[],issues={};const types=[['','표시 해제'],['temporary','일시 품절'],['sold-out','품절'],['incoming','입고예정'],['discontinued','단종'],['notice','상품 안내']];
const el=(tag,text)=>{const e=document.createElement(tag);if(text)e.textContent=text;return e};
function render(){const q=document.querySelector('#search').value.toLowerCase();const box=document.querySelector('#items');box.replaceChildren();for(const p of products.filter(p=>(p.id+' '+p.name).toLowerCase().includes(q))){const value=issues[p.id]||{};const row=el('section');row.className='item';row.append(el('h2',p.id+' · '+p.name));const fields=el('div');fields.className='fields';const type=el('select');type.setAttribute('aria-label',p.name+' 상태');for(const [v,t] of types){const o=el('option',t);o.value=v;type.append(o)}type.value=value.type||'';const restock=el('input');restock.value=value.restock||'';restock.maxLength=100;restock.placeholder='예: 9월 20일 예정 / 9월 말 / 미정';const msg=el('input');msg.value=value.message||'';msg.maxLength=300;for(const [title,input,wide] of [['상품 상태',type],['재입고 예정',restock],['추가 안내 (선택)',msg,true]]){const label=el('label',title);if(wide)label.className='wide';label.append(input);fields.append(label)}const toggle=()=>{restock.disabled=type.value!=='temporary'};type.addEventListener('change',toggle);toggle();row.append(fields);const btn=el('button','저장');btn.className='save';const feedback=el('span');feedback.className='feedback';feedback.setAttribute('role','status');btn.onclick=async()=>{btn.disabled=true;feedback.textContent='저장 중…';try{const response=await fetch('/api/catalog-status',{method:'POST',headers:{'Content-Type':'application/json','X-Workhub-Catalog':'1'},body:JSON.stringify({id:p.id,type:type.value,restock:restock.value,message:msg.value})});const data=await response.json();if(!response.ok)throw Error(data.error||'저장 실패');issues=data.issues;feedback.textContent='저장됨 · 제안서에 반영되었습니다.'}catch(e){feedback.textContent=e.message}finally{btn.disabled=false}};row.append(btn,feedback);box.append(row)}}
fetch('/api/catalog-status').then(async r=>{const data=await r.json();if(!r.ok)throw Error(data.error||'불러오기 실패');products=data.products;issues=data.issues;document.querySelector('#status').textContent=products.length+'개 상품';render()}).catch(e=>document.querySelector('#status').textContent=e.message);document.querySelector('#search').addEventListener('input',render);
</script></html>'''
