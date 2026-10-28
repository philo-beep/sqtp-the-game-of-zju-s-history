from pathlib import Path
from copy import deepcopy
from lxml import etree as E
import zipfile,json,hashlib
R=Path(__file__).resolve().parent
src=R.parent/'立项可行性报告和申请表1.0（封面排版修正）.docx'
out=R.parent/'立项可行性报告和申请表1.0（正文排版修正）.docx'
z=zipfile.ZipFile(src);d=E.fromstring(z.read('word/document.xml'))
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';q=lambda s:'{'+W+'}'+s;n={'w':W}
b=E.fromstring(zipfile.ZipFile(R/'baseline.docx').read('word/document.xml'))
rows=d.findall('.//w:tr',n);br=b.findall('.//w:tr',n)
txt=lambda x:''.join(x.xpath('.//w:t/text()',namespaces=n))
before=txt(d)
def prop(p,name,attrs):
    x=p.find(q(name))
    if x is None:x=E.SubElement(p,q(name))
    for k,v in attrs.items():x.set(q(k),str(v))
    return x
def pp(p):
    x=p.find(q('pPr'))
    if x is None:x=E.Element(q('pPr'));p.insert(0,x)
    return x
for i,row in enumerate(rows):
    trp=row.find(q('trPr'))
    if trp is None:trp=E.Element(q('trPr'));row.insert(0,trp)
    # Recover original minimum heights, without fixing height or clipping content.
    for x in trp.findall(q('trHeight')):trp.remove(x)
    original=br[i].find('w:trPr/w:trHeight',n)
    if original is not None:trp.append(deepcopy(original))
    prop(trp,'cantSplit',{})
    for p in row.findall('.//w:p',n):
        pr=pp(p);prop(pr,'widowControl',{'val':'1'});prop(pr,'keepLines',{'val':'1'})
        # Clear user-added forced pagination only inside table.
        for x in pr.findall(q('pageBreakBefore')):pr.remove(x)
        if i<=14:
            # 14pt table fields otherwise snap to two 15.6pt grid lines for each line.
            prop(pr,'snapToGrid',{'val':'0'})
            prop(pr,'spacing',{'before':'0','after':'0','line':'320','lineRule':'exact'})
        if 26<=i<=32:
            prop(pr,'snapToGrid',{'val':'0'})
            prop(pr,'spacing',{'before':'0','after':'0','line':'240','lineRule':'auto'})
            prop(pr,'keepNext',{'val':'1' if i<32 else '0'})
        elif i in [16,18,20,22,24]:
            prop(pr,'keepNext',{'val':'1'})
    # Trim only spare answer-space paragraphs, leaving table grid and signed areas intact.
    if i in [15,17,19,21,23,25]:
        for cell in row.findall('w:tc',n):
            ps=cell.findall('w:p',n)
            for p in ps:
                if len(cell.findall('w:p',n))>1 and not txt(p).strip():cell.remove(p)
# Keep personnel headings with the first corresponding data row.
for i in [0,1,3,13]:
    for p in rows[i].findall('.//w:p',n):prop(pp(p),'keepNext',{'val':'1'})
# Begin narrative and approval sections on clean pages, matching template page patterns.
prop(pp(rows[33].find('.//w:p',n)),'pageBreakBefore',{'val':'1'})
for i in [33,34]:
    for p in rows[i].findall('.//w:p',n):prop(pp(p),'keepNext',{'val':'1'})
assert txt(d)==before
old=E.fromstring(z.read('word/document.xml'))
for xp in ['.//w:tblPr','.//w:tblGrid','.//w:tcPr','.//w:sectPr','.//w:rPr']:
    # Empty answer paragraphs may contain paragraph-mark rPr; all actual runs stay untouched.
    if xp=='.//w:rPr':xp='.//w:r/w:rPr'
    assert [E.tostring(x) for x in old.findall(xp,n)]==[E.tostring(x) for x in d.findall(xp,n)],xp
assert [E.tostring(x) for x in old.find('w:body',n) if x.tag!=q('tbl')]==[E.tostring(x) for x in d.find('w:body',n) if x.tag!=q('tbl')]
with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as o:
    for item in z.infolist():
        data=z.read(item.filename)
        if item.filename=='word/document.xml':data=E.tostring(d,encoding='UTF-8',xml_declaration=True,standalone=True)
        o.writestr(item,data)
print(out)
