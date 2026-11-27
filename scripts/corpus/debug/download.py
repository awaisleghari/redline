"""Download the contracts of a test set again (the files themselves are not committed).
Usage: python3 scripts/corpus/debug/download.py tests/corpus/sources-heldout-2.json tests/fixtures/private/h2-all
Honest client name only: sites that refuse automated downloads are skipped, never worked around."""
import json,os,re,subprocess,sys,hashlib,zipfile,urllib.parse,time
out=sys.argv[2].rstrip('/')+'/'
os.makedirs(out,exist_ok=True)
items=[{'url':i.get('source') or i['url']} for i in json.load(open(sys.argv[1])) if (i.get('source') or i.get('url','')).startswith('http')]
seen=set(); log=[]
for it in items:
    u=it['url']
    if u in seen: continue
    seen.add(u)
    host=urllib.parse.urlparse(u).netloc
    base=re.sub(r'[^A-Za-z0-9._-]+','_',urllib.parse.unquote(urllib.parse.urlparse(u).path.rstrip('/').split('/')[-1]))[:90] or 'file'
    if not base.lower().endswith('.docx'): base+='.docx'
    name=f"{host}__{base}"
    path=out+name
    k=1
    while os.path.exists(path) and any(l['file']==os.path.basename(path) and l['url']!=u for l in log):
        k+=1; path=out+name.replace('.docx',f'_{k}.docx')
    if not os.path.exists(path):
        r=subprocess.run(['curl','-sS','-L','--http1.1','--max-time','120','-A','Mozilla/5.0 (compatible; corpus-fetch)','-o',path,'-w','%{http_code} %{url_effective}',u],capture_output=True,text=True)
        status=(r.stdout or r.stderr).strip()[:200]
        time.sleep(1)
    else: status='cached'
    ok=False
    try:
        z=zipfile.ZipFile(path); ok='word/document.xml' in z.namelist()
    except Exception: pass
    if not ok and os.path.exists(path):
        head=open(path,'rb').read(60); os.remove(path)
        status+=' | not docx: '+repr(head[:40])
    log.append({'file':os.path.basename(path),'url':u,'host':host,'ok':ok,'status':status,'title':it.get('title'),'publisher':it.get('publisher')})
json.dump(log,open(out+'download-log.json','w'),indent=1)
import collections
c=collections.Counter((l['host'],l['ok']) for l in log)
for (h,ok),n in sorted(c.items()): print(h,'OK' if ok else 'FAILED',n)
for h in set(l['host'] for l in log if not l['ok']):
    print('  e.g.',h,[l['status'] for l in log if l['host']==h and not l['ok']][0][:160])
