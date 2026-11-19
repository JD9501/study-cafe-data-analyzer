"""Build the browser edition of the Study Cafe analyzer from the Colab script.

The page is the same as v57. Its calls to the Flask server are answered by the same
Flask app running in a Web Worker on Pyodide (Python compiled to WebAssembly), so the
real analysis code runs in each visitor's browser and the site can be hosted as static
files (GitHub Pages).
"""
import os, re, shutil, sys, subprocess, importlib.util, glob

SRC = sys.argv[1] if len(sys.argv) > 1 else '/home/claude/work/study_cafe_v57.py'
OUT = sys.argv[2] if len(sys.argv) > 2 else '/home/claude/web/site'
PYODIDE = 'v0.27.7'
VERSION = re.search(r'v(\d+) \(', open(SRC, encoding='utf-8').read()).group(1)

code = open(SRC, encoding='utf-8').read()
def rep(s, old, new, count=1):
    n = s.count(old); assert n == count, (n, old[:90]); return s.replace(old, new)

# ── 1. the Python app, adapted for the browser ─────────────────────────────
core = code
i = core.index('# ── Silent dependency install'); j = core.index('    _pip(_p)\n') + len('    _pip(_p)\n')
core = core[:i] + 'import sys\n' + core[j:]
core = rep(core, 'from flask_cors import CORS\n', '')
core = rep(core, 'CORS(app)\n', '')
core = rep(core, '    threading.Thread(target=_lever_precompute, args=(_gen,), daemon=True).start()',
                 '    _store["lever_progress"] = {"state": "ondemand"}   # browser edition: no threads; drill-downs run on request')
core = core[:core.index('import socket, os\n\ndef _find_free_port')] + '# browser edition: no server start; the Web Worker calls the app directly\n'
core = '# Study Cafe Survey Analyzer: browser edition (runs on Pyodide), built from v%s\n' % VERSION + core

# ── 2. the page: final HTML from the script ────────────────────────────────
import subprocess as _sp
_sp.check_call = lambda *a, **k: 0
spec = importlib.util.spec_from_file_location('scsrc', SRC); m = importlib.util.module_from_spec(spec)
_argv = sys.argv; sys.argv = ['x']; spec.loader.exec_module(m); sys.argv = _argv
html = m.HTML_PAGE

html = rep(html, '''<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,600&family=Public+Sans:wght@400;600&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>''',
'''<link href="vendor/fonts.css" rel="stylesheet">
<script src="vendor/chart.umd.js"></script>
<script src="vendor/chartjs-plugin-datalabels.min.js"></script>''')
html = rep(html, "      s.src='https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';", "      s.src='vendor/jspdf.umd.min.js';")
html = rep(html, "async function _writeHtml(model){\n", "async function _writeHtml(model){\n  let _chartLib=''; try{ _chartLib=await (await fetch('vendor/chart.umd.js')).text(); }catch(e){}\n")
html = rep(html, """    +'<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"><\\/script>'""",
    """    +(_chartLib?'<script>'+_chartLib.replace(/<\\/script/gi,'<\\\\/script')+'<\\/script>':'<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"><\\/script>')""")
# drill-downs are worked out on request in the browser edition
html = rep(html, "    else if(p.state==='done'){ el.textContent='All '+(p.total||0).toLocaleString()+' drill-down combinations are ready.'; }",
    "    else if(p.state==='done'){ el.textContent='All '+(p.total||0).toLocaleString()+' drill-down combinations are ready.'; }\n"
    "    else if(p.state==='ondemand'){ el.textContent='Each drill-down is worked out when you choose it, so it can take a few seconds.'; }")

BOOT = r'''<script>
/* ── Browser edition: the analysis runs in Python inside this page ──
   The page's requests to the analysis "server" are answered by the same Python
   app running in a Web Worker on Pyodide. */
(function(){
  const API=new Set(['upload','clean','summary_data','analytics_data','download','bundled_info','bundled_dataset','reset','whatif',
                     'lever_status','predictive_data','predictor_options','predict_membership','shutdown']);
  const DATASET='data/study_cafe_survey_dataset.xlsx', DATASET_NAME='study_cafe_survey_dataset.xlsx';
  const worker=new Worker('app/worker.js?v=__VER__');
  let seq=0; const pending=new Map();
  const S=window.__engine={ready:false,pct:0,text:'Starting the analysis engine…',error:null};
  function paint(){
    let el=document.getElementById('engine-status');
    if(!el){ if(!document.body) return; el=document.createElement('div'); el.id='engine-status'; el.setAttribute('role','status'); el.setAttribute('aria-live','polite'); document.body.appendChild(el); }
    if(S.error){ el.className='es-error'; el.innerHTML='<b>The analysis engine could not start.</b> '+S.error+' Reload the page to try again.'; return; }
    if(S.ready){ el.className='es-ready'; el.innerHTML='<b>Ready.</b> The analysis runs in your browser; your data never leaves this computer.'; setTimeout(()=>el.classList.add('es-hide'),4000); return; }
    el.className=''; el.innerHTML='<b>'+S.text+'</b> <span class="es-sub">'+Math.round(S.pct)+'% · first visit takes about a minute; later visits are faster.</span><span class="es-bar"><i style="width:'+Math.round(S.pct)+'%"></i></span>';
  }
  document.addEventListener('DOMContentLoaded',paint);
  worker.onmessage=e=>{
    const m=e.data;
    if(m.type==='status'){ S.text=m.text; S.pct=m.pct; paint(); return; }
    if(m.type==='ready'){ S.ready=true; S.pct=100; paint(); return; }
    if(m.type==='fatal'){ S.error=m.error; paint(); return; }
    const p=pending.get(m.id); if(!p) return; pending.delete(m.id);
    if(m.error) p(new Response(JSON.stringify({error:m.error}),{status:500,headers:{'Content-Type':'application/json'}}));
    else p(new Response(m.body,{status:m.status,headers:m.headers}));
  };
  worker.onerror=e=>{ S.error=(e&&e.message)||'Unknown error.'; paint(); };
  const _fetch=window.fetch.bind(window);
  window.fetch=async function(input, init){
    const url=new URL(typeof input==='string'?input:input.url, location.href);
    const seg=url.pathname.replace(/\/+$/,'').split('/').pop();
    if(url.origin!==location.origin||!API.has(seg)) return _fetch(input, init);
    // the bundled survey is a plain file on this site: no need to wait for Python
    if(seg==='bundled_info') return new Response(JSON.stringify({available:true,name:DATASET_NAME}),{headers:{'Content-Type':'application/json'}});
    if(seg==='bundled_dataset') return _fetch(DATASET);
    if(seg==='shutdown') return new Response('{}',{headers:{'Content-Type':'application/json'}});
    init=init||{}; const method=(init.method||'GET').toUpperCase(), headers={}; let body=null;
    if(init.headers) new Headers(init.headers).forEach((v,k)=>{ headers[k]=v; });
    if(init.body!=null){ const r=new Response(init.body); body=await r.arrayBuffer(); const ct=r.headers.get('content-type'); if(ct&&!headers['content-type']) headers['content-type']=ct; }
    const id=++seq;
    return new Promise(res=>{ pending.set(id,res); worker.postMessage({id,method,path:'/'+seg+url.search,headers,body},body?[body]:[]); });
  };
  // The file upload uses XMLHttpRequest (for its progress bar): send those through the same bridge.
  const _XHR=window.XMLHttpRequest;
  function BridgedXHR(){
    const real=new _XHR(), self=this; let api=null;
    this.upload={onprogress:null}; this.readyState=0; this.status=0; this.responseText=''; this.response=''; this.timeout=0; this._h={};
    const sync=()=>{ ['readyState','status','responseText','response'].forEach(k=>{ try{ self[k]=real[k]; }catch(e){} }); };
    this.open=function(m,u){
      const url=new URL(u,location.href), seg=url.pathname.replace(/\/+$/,'').split('/').pop();
      if(url.origin===location.origin&&API.has(seg)){ api={m:m,u:url.href}; self.readyState=1; }
      else real.open.apply(real,arguments);
    };
    this.setRequestHeader=function(k,v){ if(api) self._h[k]=v; else real.setRequestHeader(k,v); };
    this.getResponseHeader=function(k){ return api?(self._rh?self._rh.get(k):null):real.getResponseHeader(k); };
    this.abort=function(){ if(!api) real.abort(); };
    this.send=function(body){
      if(!api){
        ['onload','onerror','ontimeout','onreadystatechange','onloadend'].forEach(k=>{ if(self[k]) real[k]=function(ev){ sync(); return self[k].call(self,ev); }; });
        if(self.upload.onprogress) real.upload.onprogress=self.upload.onprogress;
        real.timeout=self.timeout; return real.send(body);
      }
      window.fetch(api.u,{method:api.m,body:body,headers:self._h}).then(async r=>{
        self._rh=r.headers; self.status=r.status; self.responseText=await r.text(); self.response=self.responseText; self.readyState=4;
        try{ if(self.upload.onprogress) self.upload.onprogress({lengthComputable:true,loaded:1,total:1}); }catch(e){}
        if(self.onreadystatechange) self.onreadystatechange();
        if(self.onload) self.onload(); if(self.onloadend) self.onloadend();
      }).catch(e=>{ self.readyState=4; if(self.onerror) self.onerror(e); });
    };
  }
  window.XMLHttpRequest=BridgedXHR;
})();
</script>
<style>
#engine-status{position:fixed;left:16px;bottom:16px;z-index:9999;max-width:min(560px,calc(100vw - 32px));background:#17231F;color:#F5F7F4;
  border-radius:10px;padding:12px 14px;font:14px/1.45 'Public Sans',system-ui,sans-serif;box-shadow:0 8px 30px rgba(0,0,0,.25);transition:opacity .4s}
#engine-status .es-sub{color:#C5D0C8}
#engine-status .es-bar{display:block;height:4px;background:rgba(255,255,255,.18);border-radius:3px;margin-top:8px;overflow:hidden}
#engine-status .es-bar i{display:block;height:100%;background:#6FCFAE;transition:width .3s}
#engine-status.es-ready{background:#1F5C4A}
#engine-status.es-error{background:#B3261E}
#engine-status.es-hide{opacity:0;pointer-events:none}
@media (prefers-reduced-motion:reduce){#engine-status,#engine-status .es-bar i{transition:none}}
</style>
'''.replace('__VER__', VERSION)
assert html.index('<head>') < 200, html.index('<head>')
META = '<link rel="icon" href="vendor/app_icon.png">\n<meta name="description" content="Should a study café open near campus? Analyse a 6,870-student survey: demand, pricing and willingness to pay, customer groups, forecasts, a membership predictor and what-if profit scenarios. Runs in your browser.">\n<meta property="og:type" content="website">\n<meta property="og:title" content="Study Café Survey Analyzer: data-driven feasibility for a campus study café">\n<meta property="og:description" content="Survey cleaning, pricing and elasticity, customer groups, forecasts, a membership predictor and what-if profit scenarios. Python running in your browser.">\n<meta property="og:url" content="https://jd9501.github.io/study-cafe-data-analyzer/">\n<meta property="og:image" content="https://jd9501.github.io/study-cafe-data-analyzer/og-image.png">\n<meta property="og:image:width" content="1200"><meta property="og:image:height" content="627">\n<meta name="twitter:card" content="summary_large_image">\n'
html = html.replace('<head>', '<head>\n' + META + BOOT, 1)

WORKER = r'''/* Study Cafe analyzer, browser edition: runs the Python app on Pyodide. */
importScripts('https://cdn.jsdelivr.net/pyodide/__PYO__/full/pyodide.js');
const base=new URL('../', self.location.href).href;
const say=(text,pct)=>postMessage({type:'status',text,pct});
let handle=null;
const ready=(async()=>{
  try{
    say('Downloading Python…',5);
    const py=await loadPyodide();
    say('Loading pandas, scikit-learn and SciPy…',25);
    await py.loadPackage(['numpy','pandas','scipy','scikit-learn','micropip','markupsafe']);
    say('Loading the web framework…',70);
    const mp=py.pyimport('micropip');
    await mp.install(__WHEELS__.map(w=>base+'wheels/'+w), {deps:false});
    say('Loading the survey…',80);
    const ds=await (await fetch(base+'data/study_cafe_survey_dataset.xlsx')).arrayBuffer();
    py.FS.writeFile('/home/pyodide/study_cafe_survey_dataset.xlsx', new Uint8Array(ds));
    say('Starting the analysis code…',88);
    const src=await (await fetch(base+'app/study_cafe_core.py?v=__VER__')).text();
    py.FS.writeFile('/home/pyodide/study_cafe_core.py', src);
    py.runPython(`
import os, sys, json
os.environ["SC_BUNDLED_DATASET"] = "/home/pyodide/study_cafe_survey_dataset.xlsx"
sys.path.insert(0, "/home/pyodide")
import study_cafe_core as _sc
from pyodide.ffi import to_js
_client = _sc.app.test_client()
def _handle(method, path, headers_json, body):
    data = body.to_bytes() if body is not None else None
    r = _client.open(path, method=method, data=data, headers=json.loads(headers_json))
    out = r.get_data()
    h = {k: v for k, v in r.headers.items() if k.lower() not in ("content-length",)}
    return to_js([r.status_code, json.dumps(h), to_js(memoryview(out))])
`);
    handle=py.globals.get('_handle');
    postMessage({type:'ready'});
  }catch(e){ postMessage({type:'fatal',error:String(e&&e.message||e).slice(0,300)}); throw e; }
})();
onmessage=async e=>{
  const {id,method,path,headers,body}=e.data;
  try{
    await ready;
    const r=handle(method,path,JSON.stringify(headers||{}),body?new Uint8Array(body):null);
    const status=r[0], hdrs=JSON.parse(r[1]), bytes=r[2];
    const buf=bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset+bytes.byteLength);
    postMessage({id,status,headers:hdrs,body:buf},[buf]);
  }catch(err){ postMessage({id,error:String(err&&err.message||err).slice(0,2000)}); }
};
'''

# ── 3. write the site ───────────────────────────────────────────────────────
if os.path.isdir(OUT): shutil.rmtree(OUT)
for d in ['app', 'vendor', 'wheels', 'data']: os.makedirs(os.path.join(OUT, d), exist_ok=True)
open(os.path.join(OUT, 'index.html'), 'w', encoding='utf-8').write(html)
open(os.path.join(OUT, 'app', 'study_cafe_core.py'), 'w', encoding='utf-8').write(core)
wheels = sorted(os.path.basename(w) for w in glob.glob('/home/claude/web/wheels/*.whl'))
for w in wheels: shutil.copy('/home/claude/web/wheels/' + w, os.path.join(OUT, 'wheels', w))
open(os.path.join(OUT, 'app', 'worker.js'), 'w', encoding='utf-8').write(
    WORKER.replace('__PYO__', PYODIDE).replace('__WHEELS__', repr(wheels).replace("'", '"')).replace('__VER__', VERSION))
for f in os.listdir('/home/claude/desktop/vendor'):
    p = os.path.join('/home/claude/desktop/vendor', f)
    (shutil.copytree if os.path.isdir(p) else shutil.copy)(p, os.path.join(OUT, 'vendor', f))
shutil.copy('/home/claude/web/study_cafe_survey_dataset_public.xlsx', os.path.join(OUT, 'data', 'study_cafe_survey_dataset.xlsx'))
open(os.path.join(OUT, '.nojekyll'), 'w').write('')
shutil.copy('/home/claude/linkedin/og-image.png', os.path.join(OUT, 'og-image.png'))
print('site written to', OUT, '| version', VERSION, '| wheels', wheels)
