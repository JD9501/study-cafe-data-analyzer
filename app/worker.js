/* Study Cafe analyzer, browser edition: runs the Python app on Pyodide. */
importScripts('https://cdn.jsdelivr.net/pyodide/v0.27.7/full/pyodide.js');
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
    await mp.install(["blinker-1.8.2-py3-none-any.whl", "click-8.1.7-py3-none-any.whl", "et_xmlfile-2.0.0-py3-none-any.whl", "flask-3.0.3-py3-none-any.whl", "itsdangerous-2.2.0-py3-none-any.whl", "jinja2-3.1.4-py3-none-any.whl", "openpyxl-3.1.5-py2.py3-none-any.whl", "werkzeug-3.0.6-py3-none-any.whl"].map(w=>base+'wheels/'+w), {deps:false});
    say('Loading the survey…',80);
    const ds=await (await fetch(base+'data/study_cafe_survey_dataset.xlsx')).arrayBuffer();
    py.FS.writeFile('/home/pyodide/study_cafe_survey_dataset.xlsx', new Uint8Array(ds));
    say('Starting the analysis code…',88);
    const src=await (await fetch(base+'app/study_cafe_core.py?v=57')).text();
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
