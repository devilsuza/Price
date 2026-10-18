import os, re, sqlite3, uuid
from datetime import datetime
from flask import Flask, request, jsonify, render_template_string, Response
import pandas as pd

BASE_DIR=os.path.dirname(os.path.abspath(__file__))
DB_NAME=os.getenv("DB_PATH", os.path.join(BASE_DIR,"store_prices.db"))
UPLOAD_FOLDER=os.path.join(BASE_DIR,"temp_uploads")
os.makedirs(UPLOAD_FOLDER,exist_ok=True)
app=Flask(__name__)
app.config["MAX_CONTENT_LENGTH"]=20*1024*1024

def db():
    c=sqlite3.connect(DB_NAME); c.row_factory=sqlite3.Row; c.execute("PRAGMA foreign_keys=ON"); return c

def init_db():
    with db() as c:
        c.executescript('''
        CREATE TABLE IF NOT EXISTS products(
          article_code TEXT PRIMARY KEY,name TEXT NOT NULL DEFAULT '',price_cx TEXT DEFAULT '',
          price_emb TEXT DEFAULT '',price_pc_kgm TEXT DEFAULT '',price_bag TEXT DEFAULT '',updated_at TEXT);
        CREATE TABLE IF NOT EXISTS ean_aliases(
          ean_code TEXT PRIMARY KEY,article_code TEXT NOT NULL,created_at TEXT,
          FOREIGN KEY(article_code) REFERENCES products(article_code) ON DELETE CASCADE);
        CREATE INDEX IF NOT EXISTS idx_ean_article ON ean_aliases(article_code);
        CREATE TABLE IF NOT EXISTS print_jobs(
          id INTEGER PRIMARY KEY AUTOINCREMENT,article_code TEXT NOT NULL,item_name TEXT NOT NULL,
          price TEXT NOT NULL,unit TEXT NOT NULL DEFAULT 'UN',currency TEXT NOT NULL DEFAULT 'KZ',
          copies INTEGER NOT NULL DEFAULT 1,status TEXT NOT NULL DEFAULT 'pending',created_at TEXT NOT NULL,
          claimed_at TEXT,printed_at TEXT,error_message TEXT);
        ''')
        c.execute("INSERT OR IGNORE INTO ean_aliases(ean_code,article_code,created_at) SELECT article_code,article_code,COALESCE(updated_at,'') FROM products")
init_db()

def clean(v):
    if v is None or pd.isna(v): return ''
    s=str(v).strip()
    return s[:-2] if s.endswith('.0') else s

def price(v):
    s=clean(v)
    if s.lower() in ('','nan','none','-','0','0.0','0.00'): return ''
    return s

def now(): return datetime.now().strftime('%d-%b-%Y %H:%M')

def read_sheet(path,sheet):
    return pd.read_csv(path,dtype=str) if path.lower().endswith('.csv') else pd.read_excel(path,sheet_name=sheet,dtype=str)

HTML=r'''<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Fresmart Price Portal</title><script src="https://cdn.tailwindcss.com"></script><script src="https://unpkg.com/html5-qrcode"></script>
<style>body{background:#f1f5f9}.tab.active{background:#fff;color:#4338ca}.toast{position:fixed;right:18px;bottom:18px;z-index:99}.modal{position:fixed;inset:0;background:rgba(15,23,42,.88);z-index:100;display:flex;align-items:center;justify-content:center;padding:18px}.hidden{display:none!important}#reader video{border-radius:14px}</style></head>
<body class="text-slate-800"><div id="loginModal" class="modal"><div class="bg-white max-w-sm w-full rounded-3xl p-6 shadow-2xl"><h1 class="text-2xl font-black text-indigo-700 text-center">FRESMART LOGIN</h1><p class="text-sm text-slate-500 text-center">Select role and enter PIN</p><div class="grid grid-cols-2 gap-3 mt-5"><button id="roleEmployee" onclick="pickRole('employee')" class="border-2 border-indigo-600 bg-indigo-50 rounded-xl p-3 font-bold text-indigo-700">Employee</button><button id="roleManager" onclick="pickRole('manager')" class="border-2 border-slate-200 rounded-xl p-3 font-bold">Manager</button></div><input id="loginPin" type="password" inputmode="numeric" class="mt-4 w-full border-2 border-indigo-300 rounded-xl px-4 py-3 text-center text-xl" placeholder="Enter PIN"><button onclick="login()" class="mt-3 w-full bg-indigo-600 text-white py-3 rounded-xl font-black">LOGIN</button><p id="loginError" class="hidden mt-3 text-center text-red-600 font-bold">Invalid PIN</p><p class="mt-4 text-xs text-center text-slate-400">Employee: 123 | Manager: 2026</p></div></div><header class="bg-indigo-700 text-white shadow"><div class="max-w-7xl mx-auto p-4 flex flex-wrap gap-3 justify-between items-center"><div><h1 class="text-xl font-black">FRESMART PRICE PORTAL</h1><p class="text-xs text-indigo-200">EAN Search · Master Merge · Daily Price Update · Label Printing</p></div><div class="bg-indigo-600 p-1 rounded-xl"><button id="tSearch" onclick="showTab('search')" class="tab active px-4 py-2 rounded-lg font-bold">Search & Print</button><button id="tManager" onclick="showTab('manager')" class="tab hidden px-4 py-2 rounded-lg font-bold">Manager</button><button onclick="logout()" class="ml-2 bg-red-500 px-3 py-2 rounded-lg font-bold">Logout</button></div></div></header>
<main class="max-w-7xl mx-auto p-4 md:p-6">
<section id="searchTab"><div class="bg-white rounded-2xl shadow p-5"><div class="flex justify-between mb-3"><h2 class="font-black text-lg">Scan EAN / Search Article</h2><span id="scanStatus" class="text-xs font-bold text-slate-500">Scanner off</span></div><div class="flex flex-col sm:flex-row gap-2"><div class="relative flex-1"><input id="q" readonly inputmode="none" autocomplete="off" class="w-full border-2 border-indigo-300 rounded-xl px-4 pr-24 py-3 text-lg" placeholder="Scan barcode or tap Type"><button id="typeBtn" onclick="enableTyping()" class="absolute right-2 top-2 bg-slate-100 px-3 py-2 rounded-lg font-bold">Type</button></div><button onclick="search()" class="bg-indigo-600 text-white rounded-xl px-6 py-3 font-bold">SEARCH</button><button id="cameraBtn" onclick="toggleScanner()" class="bg-emerald-600 text-white rounded-xl px-6 py-3 font-bold">SCAN</button></div><div id="readerWrap" class="hidden mt-4 max-w-lg mx-auto"><div id="reader" class="overflow-hidden border-2 border-emerald-400 rounded-2xl"></div><button onclick="stopScanner()" class="mt-2 w-full bg-red-500 text-white py-2 rounded-xl font-bold">STOP CAMERA</button></div><div id="results" class="mt-5 space-y-4"></div></div></section>
<section id="managerTab" class="hidden space-y-5"><div class="grid lg:grid-cols-2 gap-5">
<div class="bg-white rounded-2xl shadow p-5 border-2 border-indigo-200"><h2 class="font-black text-lg">Master Sheet Merge</h2><p class="text-sm text-slate-500 mb-3">New articles/EANs add honge. Blank cells old data ko erase nahi karenge.</p><input id="masterFile" type="file" accept=".xlsx,.xls,.csv" class="w-full border rounded-xl p-2"><button onclick="preview('master')" class="mt-3 w-full bg-indigo-600 text-white py-3 rounded-xl font-bold">READ MASTER FILE</button></div>
<div class="bg-white rounded-2xl shadow p-5 border-2 border-amber-200"><h2 class="font-black text-lg">Daily Price Update</h2><p class="text-sm text-slate-500 mb-3">Sirf uploaded non-empty prices update honge.</p><input id="priceFile" type="file" accept=".xlsx,.xls,.csv" class="w-full border rounded-xl p-2"><button onclick="preview('price')" class="mt-3 w-full bg-amber-500 text-white py-3 rounded-xl font-bold">READ PRICE FILE</button></div></div>
<div id="mapping" class="hidden bg-white rounded-2xl shadow p-5"><h2 class="font-black text-lg mb-3">Column Mapping</h2><div id="mapBody"></div><button onclick="runImport()" class="mt-4 bg-emerald-600 text-white px-6 py-3 rounded-xl font-bold">SAVE / UPDATE</button></div>
<div class="bg-white rounded-2xl shadow p-5"><div class="flex justify-between mb-3"><h2 class="font-black text-lg">Master Articles</h2><button onclick="loadProducts()" class="text-indigo-700 font-bold">Refresh</button></div><div class="overflow-auto"><table class="w-full text-sm"><thead><tr class="bg-slate-100"><th class="p-2 text-left">Article</th><th class="p-2 text-left">Name</th><th>CX</th><th>EMB</th><th>PC/KGM</th><th>BAG</th><th>EANs</th></tr></thead><tbody id="products"></tbody></table></div></div></section>
</main><div id="toast" class="toast hidden px-5 py-3 rounded-xl shadow-xl font-bold"></div>
<script>
const PINS={employee:'123',manager:'2026'}; let selectedRole='employee',currentRole='',tempFile='',mode='',tabs=[],scanner=null,scanning=false,lastScan='',lastScanAt=0;
function toast(m,ok=true){let e=document.getElementById('toast');e.textContent=m;e.className='toast px-5 py-3 rounded-xl shadow-xl font-bold '+(ok?'bg-emerald-600':'bg-red-600')+' text-white';setTimeout(()=>e.classList.add('hidden'),3500)}
function showTab(x){if(x==='manager'&&currentRole!=='manager'){toast('Manager access required',false);return}stopScanner();document.getElementById('searchTab').classList.toggle('hidden',x!='search');document.getElementById('managerTab').classList.toggle('hidden',x!='manager');document.getElementById('tSearch').classList.toggle('active',x=='search');document.getElementById('tManager').classList.toggle('active',x=='manager');if(x==='manager')loadProducts();else lockKeyboard()}
function pickRole(r){selectedRole=r;document.getElementById('roleEmployee').className='border-2 rounded-xl p-3 font-bold '+(r==='employee'?'border-indigo-600 bg-indigo-50 text-indigo-700':'border-slate-200');document.getElementById('roleManager').className='border-2 rounded-xl p-3 font-bold '+(r==='manager'?'border-indigo-600 bg-indigo-50 text-indigo-700':'border-slate-200')}
function login(){if(document.getElementById('loginPin').value!==PINS[selectedRole]){document.getElementById('loginError').classList.remove('hidden');return}currentRole=selectedRole;sessionStorage.setItem('role',currentRole);document.getElementById('loginModal').classList.add('hidden');applyRole();lockKeyboard()}
function applyRole(){document.getElementById('tManager').classList.toggle('hidden',currentRole!=='manager');if(currentRole!=='manager')showTab('search')}
function logout(){stopScanner();sessionStorage.removeItem('role');currentRole='';document.getElementById('loginPin').value='';document.getElementById('loginModal').classList.remove('hidden')}
function lockKeyboard(){let q=document.getElementById('q');q.readOnly=true;q.setAttribute('inputmode','none');q.blur();document.getElementById('typeBtn').textContent='Type'}
function enableTyping(){let q=document.getElementById('q');q.readOnly=false;q.setAttribute('inputmode','text');q.focus();document.getElementById('typeBtn').textContent='Lock'}
function beep(){try{let a=new(window.AudioContext||window.webkitAudioContext)(),o=a.createOscillator(),g=a.createGain();o.connect(g);g.connect(a.destination);o.frequency.value=1050;g.gain.setValueAtTime(.2,a.currentTime);g.gain.exponentialRampToValueAtTime(.001,a.currentTime+.15);o.start();o.stop(a.currentTime+.15)}catch(e){}}
async function toggleScanner(){if(scanning)return stopScanner();if(typeof Html5Qrcode==='undefined'){toast('Scanner library load nahi hui',false);return}lockKeyboard();document.getElementById('readerWrap').classList.remove('hidden');scanner=new Html5Qrcode('reader');try{await scanner.start({facingMode:'environment'},{fps:12,qrbox:{width:280,height:150}},onScan,()=>{});scanning=true;document.getElementById('scanStatus').textContent='Scanner active';document.getElementById('cameraBtn').textContent='STOP'}catch(e){toast('Camera permission allow karein',false);document.getElementById('readerWrap').classList.add('hidden')}}
async function stopScanner(){if(scanner&&scanning){try{await scanner.stop();await scanner.clear()}catch(e){}}scanner=null;scanning=false;document.getElementById('readerWrap').classList.add('hidden');document.getElementById('scanStatus').textContent='Scanner off';document.getElementById('cameraBtn').textContent='SCAN';lockKeyboard()}
function onScan(text){let n=Date.now();if(text===lastScan&&n-lastScanAt<1800)return;lastScan=text;lastScanAt=n;beep();document.getElementById('q').value=text;lockKeyboard();search();stopScanner()}
document.getElementById('q').addEventListener('keydown',e=>{if(e.key==='Enter'){search();lockKeyboard()}});document.getElementById('q').addEventListener('blur',()=>setTimeout(lockKeyboard,100));
async function search(){let q=document.getElementById('q').value.trim();if(!q)return;lockKeyboard();let d=await fetch('/api/search?q='+encodeURIComponent(q)).then(r=>r.json());let box=document.getElementById('results');if(!d.items.length){box.innerHTML='<div class="p-6 bg-red-50 border border-red-300 rounded-xl font-bold text-red-700">Article not found</div>';return}box.innerHTML=d.items.map(item=>`<div class="bg-white border-2 border-indigo-300 rounded-2xl p-5"><div class="mb-4"><span class="bg-indigo-100 text-indigo-800 px-3 py-1 rounded-lg font-black">${esc(item.article_code)}</span><h3 class="text-xl font-black mt-2">${esc(item.name)}</h3></div><div class="grid sm:grid-cols-2 lg:grid-cols-4 gap-3">${card(item,'price_cx','CX','CX')}${card(item,'price_emb','EMB','EMB')}${card(item,'price_pc_kgm','PC / KGM','PC/KGM')}${card(item,'price_bag','BAG','BAG')}</div><div id="ps" class="mt-3"></div></div>`).join('')}
function esc(v){return String(v??'').replace(/[&<>"']/g,x=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[x]))}
function card(i,k,l,u){let v=i[k]||'';let t=encodeURIComponent(JSON.stringify({article_code:i.article_code,item_name:i.name,price:v,unit:u,currency:'KZ'}));return `<button ${v?`onclick="printLabel('${t}')"`:'disabled'} class="border-2 rounded-xl p-4 ${v?'border-indigo-300 hover:bg-indigo-50':'opacity-40'}"><b>${l}</b><span class="block text-2xl font-black">${v||'--'}</span>${v?'<small>PRINT LABEL</small>':''}</button>`}
async function printLabel(x){let t=JSON.parse(decodeURIComponent(x));let q=prompt('Label quantity (1-100)','1');if(q===null)return;let n=Number(q);if(!Number.isInteger(n)||n<1||n>100){toast('Quantity 1 se 100 tak whole number honi chahiye',false);return}t.copies=n;let r=await fetch('/api/print-jobs',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(t)});let d=await r.json();toast(d.success?`${n} label(s) print queue me add hue`:(d.error||'Print failed'),d.success)}
async function preview(m){mode=m;let f=document.getElementById(m==='master'?'masterFile':'priceFile').files[0];if(!f){toast('File select karein',false);return}let fd=new FormData();fd.append('file',f);let r=await fetch('/api/preview',{method:'POST',body:fd});let d=await r.json();if(!d.success){toast(d.error,false);return}tempFile=d.temp_filename;tabs=d.tabs;renderMap();}
function opts(cols,guess,optional=true){let a=optional?'<option value="">-- Not selected --</option>':'';return a+cols.map(c=>`<option ${guess.some(g=>c.toLowerCase().includes(g))?'selected':''}>${esc(c)}</option>`).join('')}
function renderMap(){let b=document.getElementById('mapBody');b.innerHTML=tabs.map((t,i)=>`<div class="map border rounded-xl p-4 mb-3" data-sheet="${esc(t.sheet_name)}"><b>${esc(t.sheet_name)} (${t.row_count} rows)</b><div class="grid md:grid-cols-3 gap-3 mt-3"><label>Article Code / EAN<select class="code w-full border p-2 rounded">${opts(t.columns,mode==='master'?['article','artigo','item code','sku']:['ean','barcode','article'],false)}</select></label><label class="${mode==='price'?'hidden':''}">EAN Column<select class="ean w-full border p-2 rounded">${opts(t.columns,['ean','barcode'])}</select></label><label>Name<select class="name w-full border p-2 rounded">${opts(t.columns,['name','description','descricao'])}</select></label><label>Price CX<select class="cx w-full border p-2 rounded">${opts(t.columns,['price cx','preco cx','cx'])}</select></label><label>Price EMB<select class="emb w-full border p-2 rounded">${opts(t.columns,['price emb','preco emb','emb'])}</select></label><label>Price PC/KGM<select class="pc w-full border p-2 rounded">${opts(t.columns,['pc/kgm','price pc','preco pc','kgm'])}</select></label><label>Price BAG<select class="bag w-full border p-2 rounded">${opts(t.columns,['price bag','preco bag','bag'])}</select></label></div></div>`).join('');document.getElementById('mapping').classList.remove('hidden')}
async function runImport(){let maps=[...document.querySelectorAll('.map')].map(x=>({sheet_name:x.dataset.sheet,code_col:x.querySelector('.code').value,ean_col:x.querySelector('.ean').value,name_col:x.querySelector('.name').value,cx_col:x.querySelector('.cx').value,emb_col:x.querySelector('.emb').value,pc_col:x.querySelector('.pc').value,bag_col:x.querySelector('.bag').value}));let r=await fetch('/api/import',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({temp_filename:tempFile,mode,mappings:maps})});let d=await r.json();toast(d.success?`Added ${d.added}, Updated ${d.updated}, EAN linked ${d.ean_linked}, Skipped ${d.skipped}`:d.error,d.success);if(d.success){loadProducts();document.getElementById('mapping').classList.add('hidden')}}
async function loadProducts(){let d=await fetch('/api/products').then(r=>r.json());document.getElementById('products').innerHTML=d.items.map(x=>`<tr class="border-b"><td class="p-2 font-bold">${esc(x.article_code)}</td><td class="p-2">${esc(x.name)}</td><td class="text-center">${esc(x.price_cx)}</td><td class="text-center">${esc(x.price_emb)}</td><td class="text-center">${esc(x.price_pc_kgm)}</td><td class="text-center">${esc(x.price_bag)}</td><td class="text-center">${x.ean_count}</td></tr>`).join('')}
window.addEventListener('load',()=>{let r=sessionStorage.getItem('role');if(r&&PINS[r]){currentRole=r;document.getElementById('loginModal').classList.add('hidden');applyRole()}lockKeyboard()});
</script></body></html>'''

@app.get('/')
def home(): return render_template_string(HTML)

@app.get('/health')
def health(): return jsonify(ok=True)

@app.get('/api/search')
def search():
    q=clean(request.args.get('q')); like=f'%{q}%'
    with db() as c:
        exact=c.execute('''SELECT DISTINCT p.* FROM products p LEFT JOIN ean_aliases e ON e.article_code=p.article_code WHERE p.article_code=? OR e.ean_code=? LIMIT 20''',(q,q)).fetchall()
        rows=exact or c.execute('''SELECT DISTINCT p.* FROM products p LEFT JOIN ean_aliases e ON e.article_code=p.article_code WHERE p.name LIKE ? OR p.article_code LIKE ? OR e.ean_code LIKE ? LIMIT 20''',(like,like,like)).fetchall()
    return jsonify(items=[dict(x) for x in rows])

@app.get('/api/products')
def products():
    with db() as c: rows=c.execute('''SELECT p.*,COUNT(e.ean_code) ean_count FROM products p LEFT JOIN ean_aliases e ON e.article_code=p.article_code GROUP BY p.article_code ORDER BY p.name LIMIT 5000''').fetchall()
    return jsonify(items=[dict(x) for x in rows])

@app.post('/api/preview')
def preview():
    f=request.files.get('file')
    if not f or not f.filename: return jsonify(success=False,error='File missing'),400
    ext=os.path.splitext(f.filename)[1].lower()
    if ext not in ('.csv','.xlsx','.xls'): return jsonify(success=False,error='Only CSV/XLSX/XLS allowed'),400
    name=uuid.uuid4().hex+ext; path=os.path.join(UPLOAD_FOLDER,name); f.save(path)
    try:
        sheets=['CSV_Data'] if ext=='.csv' else pd.ExcelFile(path).sheet_names
        tabs=[]
        for sh in sheets:
            d=read_sheet(path,sh).dropna(how='all'); tabs.append({'sheet_name':sh,'columns':[str(x) for x in d.columns],'row_count':len(d)})
        return jsonify(success=True,temp_filename=name,tabs=tabs)
    except Exception as e:
        try: os.remove(path)
        except OSError: pass
        return jsonify(success=False,error=str(e)),400

@app.post('/api/import')
def do_import():
    x=request.get_json(silent=True) or {}; mode=x.get('mode'); filename=os.path.basename(x.get('temp_filename',''))
    path=os.path.join(UPLOAD_FOLDER,filename)
    if mode not in ('master','price') or not os.path.exists(path): return jsonify(success=False,error='Invalid or expired import'),400
    added=updated=ean_linked=skipped=0
    try:
      with db() as c:
       for m in x.get('mappings',[]):
        d=read_sheet(path,m['sheet_name']).dropna(how='all')
        for _,r in d.iterrows():
         code=clean(r.get(m['code_col']))
         if not code: skipped+=1; continue
         vals={'price_cx':price(r.get(m.get('cx_col'))) if m.get('cx_col') else '', 'price_emb':price(r.get(m.get('emb_col'))) if m.get('emb_col') else '', 'price_pc_kgm':price(r.get(m.get('pc_col'))) if m.get('pc_col') else '', 'price_bag':price(r.get(m.get('bag_col'))) if m.get('bag_col') else ''}
         if mode=='price':
          a=c.execute('SELECT article_code FROM ean_aliases WHERE ean_code=?',(code,)).fetchone(); article=a[0] if a else code
          if not c.execute('SELECT 1 FROM products WHERE article_code=?',(article,)).fetchone(): skipped+=1; continue
          sets=[];v=[]
          for k,z in vals.items():
           if z: sets.append(k+'=?');v.append(z)
          if not sets: skipped+=1;continue
          sets.append('updated_at=?');v.extend([now(),article]);c.execute('UPDATE products SET '+','.join(sets)+' WHERE article_code=?',v);updated+=1;continue
         article=code; exists=c.execute('SELECT 1 FROM products WHERE article_code=?',(article,)).fetchone(); nm=clean(r.get(m.get('name_col'))) if m.get('name_col') else ''
         stored_name=nm if nm else (article if not exists else '')
         c.execute('''INSERT INTO products(article_code,name,price_cx,price_emb,price_pc_kgm,price_bag,updated_at) VALUES(?,?,?,?,?,?,?) ON CONFLICT(article_code) DO UPDATE SET name=CASE WHEN excluded.name<>'' THEN excluded.name ELSE products.name END,price_cx=CASE WHEN excluded.price_cx<>'' THEN excluded.price_cx ELSE products.price_cx END,price_emb=CASE WHEN excluded.price_emb<>'' THEN excluded.price_emb ELSE products.price_emb END,price_pc_kgm=CASE WHEN excluded.price_pc_kgm<>'' THEN excluded.price_pc_kgm ELSE products.price_pc_kgm END,price_bag=CASE WHEN excluded.price_bag<>'' THEN excluded.price_bag ELSE products.price_bag END,updated_at=excluded.updated_at''',(article,stored_name,vals['price_cx'],vals['price_emb'],vals['price_pc_kgm'],vals['price_bag'],now()))
         added+=int(not exists);updated+=int(bool(exists)); aliases=[article]
         if m.get('ean_col'):
          aliases += [clean(y) for y in re.split(r'[,;|\n]+',clean(r.get(m['ean_col']))) if clean(y)]
         for e in dict.fromkeys(aliases): c.execute('INSERT INTO ean_aliases(ean_code,article_code,created_at) VALUES(?,?,?) ON CONFLICT(ean_code) DO UPDATE SET article_code=excluded.article_code',(e,article,now()));ean_linked+=1
      try: os.remove(path)
      except OSError: pass
      return jsonify(success=True,added=added,updated=updated,ean_linked=ean_linked,skipped=skipped)
    except Exception as e: return jsonify(success=False,error=str(e)),500

@app.post('/api/print-jobs')
def add_print():
    x=request.get_json(silent=True) or {}
    for k in ('article_code','item_name','price'):
        if not clean(x.get(k)): return jsonify(success=False,error='Missing '+k),400
    try: copies=int(x.get('copies',1))
    except: return jsonify(success=False,error='Invalid quantity'),400
    if not 1<=copies<=100:return jsonify(success=False,error='Quantity must be 1-100'),400
    with db() as c:
        cur=c.execute("INSERT INTO print_jobs(article_code,item_name,price,unit,currency,copies,status,created_at) VALUES(?,?,?,?,?,?,'pending',?)",(clean(x['article_code']),clean(x['item_name']),clean(x['price']),clean(x.get('unit')) or 'UN',clean(x.get('currency')) or 'KZ',copies,datetime.now().isoformat(timespec='seconds')))
    return jsonify(success=True,job_id=cur.lastrowid,copies=copies)

def bridge_ok():
    key=os.getenv('PRINT_BRIDGE_KEY','');return bool(key) and request.headers.get('X-Print-Bridge-Key')==key
@app.get('/api/print-jobs/next')
def next_print():
    if not bridge_ok():return jsonify(error='Unauthorized'),401
    with db() as c:
        r=c.execute("SELECT * FROM print_jobs WHERE status='pending' ORDER BY id LIMIT 1").fetchone()
        if not r:return Response(status=204)
        c.execute("UPDATE print_jobs SET status='claimed',claimed_at=? WHERE id=?",(datetime.now().isoformat(timespec='seconds'),r['id']))
    return jsonify(dict(r))
@app.post('/api/print-jobs/<int:jid>/complete')
def complete(jid):
    if not bridge_ok():return jsonify(error='Unauthorized'),401
    x=request.get_json(silent=True) or {};ok=bool(x.get('success'))
    with db() as c:
        if ok:c.execute("UPDATE print_jobs SET status='printed',printed_at=?,error_message=NULL WHERE id=?",(datetime.now().isoformat(timespec='seconds'),jid))
        else:c.execute("UPDATE print_jobs SET status='pending',error_message=? WHERE id=?",(clean(x.get('error'))[:500],jid))
    return jsonify(success=True)

if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.getenv('PORT','5000')),debug=False)
