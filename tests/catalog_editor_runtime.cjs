const fs=require('fs'),vm=require('vm'),assert=require('assert');
class E{
 constructor(tag){this.tag=tag;this.children=[];this.dataset={};this.value='';this.listeners={};this.style={};this.files=[];}
 append(...xs){this.children.push(...xs)}
 replaceChildren(...xs){this.children=xs}
 setAttribute(){}
 addEventListener(n,f){this.listeners[n]=f}
 querySelector(){return null}
}
const nodes={};for(const id of ['documentUploads','refreshDocuments','search','filter','items','status','total','temporary','marked'])nodes['#'+id]=new E('div');nodes['#filter'].value='all';
const calls=[];const p={id:'SB-EX-009',name:'test',main:[{url:'/a.jpg'}],preview:'/a.jpg',spec:'',features:'',price:100,sortPrice:100,pack:'1EA',revision:0};
const ctx={console,URL,setTimeout,confirm:()=>true,location:{origin:'https://test'},window:{addEventListener(){}},document:{createElement:t=>new E(t),querySelector:s=>nodes[s]||null,querySelectorAll:()=>[]},fetch:async(url,options)=>{calls.push([url,options]);return {ok:true,json:async()=>url.includes('documents')?{documents:{}}:options?.method==='POST'?{issues:{[p.id]:{type:'temporary',restock:'2026-10-01'}}}:{products:[p],issues:{}}}}};
vm.createContext(ctx);const source=fs.readFileSync('scripts/catalog_editor.html','utf8').split('<script>')[1].split('</script>')[0];vm.runInContext(source,ctx);
(async()=>{
 await new Promise(r=>setTimeout(r,0));assert.equal(nodes['#documentUploads'].children.length,2);
 const editor=nodes['#items'].children[0].children[1],fields=editor.children[0];
 const type=fields.children[0].children[0],date=fields.children[1].children[0];type.value='temporary';date.value='2026-10-01';type.listeners.change();
 const actions=editor.children[1],button=actions.children[1];await button.onclick();
 assert(calls.some(([u,o])=>u==='/api/catalog-status'&&o?.method==='POST'));
 assert.equal(actions.children[0].textContent,'저장 완료');
 console.log('PASS: document initialization and stock/date save runtime');
})().catch(e=>{console.error(e);process.exitCode=1});
