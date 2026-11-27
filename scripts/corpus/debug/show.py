import sys,zipfile,glob
from lxml import etree
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
q=lambda n:'{%s}%s'%(W,n)
K={'ins':'+','del':'-','moveFrom':'<','moveTo':'>'}
for f in sorted(glob.glob(sys.argv[1]+'/*.docx')):
    root=etree.fromstring(zipfile.ZipFile(f).read('word/document.xml'))
    out=[]
    for p in root.iter(q('p')):
        s=''
        for r in p.iter(q('r')):
            if r.getparent() is not None and any(etree.QName(a).localname=='txbxContent' for a in r.iterancestors()): pass
            t=''.join((x.text or '') for x in r if etree.QName(x).localname in('t','delText','instrText','delInstrText'))
            if any(etree.QName(x).localname=='fldChar' for x in r): t='{'+r.find(q('fldChar')).get(q('fldCharType'))[0]+'}'
            k=next((K[etree.QName(a).localname] for a in r.iterancestors() if etree.QName(a).localname in K),'')
            s+=('[%s%s%s]'%(k,t,k)) if k else t
        rpr=p.find(q('pPr')+'/'+q('rPr'))
        mk=next((K[etree.QName(c).localname] for c in (rpr if rpr is not None else []) if etree.QName(c).localname in K),'')
        out.append(s+' ¶'+mk)
    print('==',f.split('/')[-1]); print('   '+' | '.join(out)[:520])
