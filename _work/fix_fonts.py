from pathlib import Path
import zipfile,re,hashlib,json
from lxml import etree as E
R=Path(__file__).resolve().parent
src=R.parent/'立项可行性报告和申请表1.0.docx'
out=R.parent/'立项可行性报告和申请表1.0（字体标点修正）.docx'
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
n={'w':W}; q=lambda x:'{'+W+'}'+x
z=zipfile.ZipFile(src); root=E.fromstring(z.read('word/document.xml'))
b=E.fromstring(zipfile.ZipFile(R/'baseline.docx').read('word/document.xml'))
pt=lambda p:''.join(p.xpath('.//w:t/text()',namespaces=n))
original_texts={pt(p) for p in b.findall('.//w:body//w:p',n)}
before=''.join(root.xpath('//w:t/text()',namespaces=n))
count=0; punctuation=0
for p in root.findall('.//w:body//w:p',n):
    t=pt(p)
    if not t.strip() or t in original_texts: continue
    in_table=any(a.tag==q('tc') for a in p.iterancestors())
    for r in p.findall('w:r',n):
        if not r.findall('.//w:t',n): continue
        rp=r.find(q('rPr'))
        if rp is None: rp=E.Element(q('rPr')); r.insert(0,rp)
        f=rp.find(q('rFonts'))
        if f is None: f=E.Element(q('rFonts')); rp.insert(0,f)
        for k in ['asciiTheme','hAnsiTheme','eastAsiaTheme','cstheme','csTheme']:
            f.attrib.pop(q(k),None)
        f.set(q('ascii'),'Times New Roman');f.set(q('hAnsi'),'Times New Roman')
        f.set(q('eastAsia'),'仿宋_GB2312' if in_table else '宋体')
        count+=1
    if t.startswith('理论学习：'):
        quote_open=True; prev=''
        for node in p.findall('.//w:t',n):
            chars=[]
            for c in node.text or '':
                if c=='"':
                    c='“' if quote_open else '”'; quote_open=not quote_open;punctuation+=1
                if c=='。' and prev=='。': punctuation+=1;continue
                chars.append(c);prev=c
            node.text=''.join(chars)
after=''.join(root.xpath('//w:t/text()',namespaces=n))
norm=lambda s:s.replace('"','').replace('“','').replace('”','').replace('。。','。')
assert norm(before)==norm(after)
# Verify all paragraph properties, table structures, page settings, sizes and bolds stay unchanged.
old=E.fromstring(z.read('word/document.xml'))
for xp in ['.//w:pPr','.//w:tblPr','.//w:tblGrid','.//w:trPr','.//w:tcPr','.//w:sectPr','.//w:rPr/w:sz','.//w:rPr/w:b']:
    assert [E.tostring(x) for x in old.findall(xp,n)]==[E.tostring(x) for x in root.findall(xp,n)],xp
with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as o:
    for i in z.infolist():
        data=z.read(i.filename)
        if i.filename=='word/document.xml':data=E.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)
        o.writestr(i,data)
print(json.dumps({'runs_font_normalized':count,'punctuation_changes':punctuation,'source_sha256':hashlib.sha256(src.read_bytes()).hexdigest(),'output':str(out)},ensure_ascii=False))
