from pathlib import Path
from copy import deepcopy
from lxml import etree as E
import zipfile,hashlib
R=Path(__file__).resolve().parent
src=R.parent/'立项可行性报告和申请表1.0（表达修正）.docx'
out=R.parent/'立项可行性报告和申请表1.0（封面排版修正）.docx'
z=zipfile.ZipFile(src);root=E.fromstring(z.read('word/document.xml'))
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';q=lambda x:'{'+W+'}'+x;n={'w':W}
ps=root.find('w:body',n).findall('w:p',n)
def prop(parent,name,attrs):
    x=parent.find(q(name))
    if x is None:x=E.SubElement(parent,q(name))
    for k,v in attrs.items():x.set(q(k),str(v))
    return x
def run(p,t='',under=False,br=False):
    r=E.SubElement(p,q('r'));rp=E.SubElement(r,q('rPr'))
    prop(rp,'rFonts',{'eastAsia':'宋体','ascii':'Times New Roman','hAnsi':'Times New Roman'})
    prop(rp,'b',{'val':'1'});prop(rp,'sz',{'val':'28'});prop(rp,'color',{'val':'000000'})
    if under:prop(rp,'u',{'val':'single'})
    if br:E.SubElement(r,q('br'))
    elif t=='\t':E.SubElement(r,q('tab'))
    else:E.SubElement(r,q('t'),{'{http://www.w3.org/XML/1998/namespace}space':'preserve'}).text=t
for idx in range(7,14):
    p=ps[idx];old=''.join(p.xpath('.//w:t/text()',namespaces=n)).strip()
    label,value=old.split('：',1);value=value.rstrip('_').strip()
    # Keep existing wording and personal fields; only reflow cover entries.
    if idx==7:label='项目名称'
    elif idx==9:label='立 项 人'
    elif idx==10:label='专    业'
    pp=p.find(q('pPr'))
    if pp is None:pp=E.Element(q('pPr'));p.insert(0,pp)
    for name in ['tabs','ind','spacing','jc']:
        for x in pp.findall(q(name)):pp.remove(x)
    prop(pp,'ind',{'left':'300','right':'300','firstLine':'0'})
    prop(pp,'jc',{'val':'left'})
    prop(pp,'spacing',{'before':'0','after':'0','line':'640','lineRule':'exact'})
    tabs=E.SubElement(pp,q('tabs'))
    prop(tabs,'tab',{'val':'center','pos':'4940'})
    E.SubElement(tabs,q('tab'),{q('val'):'right',q('pos'):'8000'})
    for x in list(p):
        if x.tag!=q('pPr'):p.remove(x)
    run(p,label+'：')
    lines=[value]
    for i,line in enumerate(lines):
        if i:run(p,br=True)
        run(p,'\t',True);run(p,line,True);run(p,'\t',True)
# Equalize the blank gap before and after the info block to keep cover on one page.
for i in range(14,18):
    pp=ps[i].find(q('pPr'))
    prop(pp,'spacing',{'line':'260','lineRule':'exact','before':'0','after':'0'})
# Only cover paragraphs may change.
old=E.fromstring(z.read('word/document.xml'))
oldbody=old.find('w:body',n);body=root.find('w:body',n)
for i,(a,b) in enumerate(zip(oldbody,body)):
    if E.tostring(a)!=E.tostring(b):assert 7<=i<=17,i
with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as o:
    for item in z.infolist():
        data=z.read(item.filename)
        if item.filename=='word/document.xml':data=E.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)
        o.writestr(item,data)
print(out)
