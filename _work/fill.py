from pathlib import Path
from lxml import etree as E
from copy import deepcopy
import zipfile, json, hashlib
ROOT=Path(__file__).resolve().parent
base=ROOT/'baseline.docx'
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ns={'w':W}
def q(t): return '{'+W+'}'+t
z=zipfile.ZipFile(base); doc=E.fromstring(z.read('word/document.xml'))
paras=doc.findall('.//w:body//w:p',ns)
rows=doc.findall('.//w:tr',ns)
log=[]
def put(p,text,font='仿宋_GB2312',size=12,bold=False):
    for x in list(p):
        if x.tag!=q('pPr'): p.remove(x)
    r=E.SubElement(p,q('r')); rp=E.SubElement(r,q('rPr'))
    E.SubElement(rp,q('rFonts'),{q('ascii'):'Times New Roman',q('hAnsi'):'Times New Roman',q('eastAsia'):font})
    E.SubElement(rp,q('b'),{q('val'):'1' if bold else '0'})
    E.SubElement(rp,q('sz'),{q('val'):str(int(size*2))})
    E.SubElement(rp,q('color'),{q('val'):'000000'})
    E.SubElement(r,q('t')).text=text
    log.append({'text':text,'font':font,'size':size,'bold':bold})
def slot(prefix, texts):
    p=next(p for p in paras if ''.join(p.itertext()).startswith(prefix))
    cell=p.getparent(); ps=cell.findall('w:p',ns)
    for i,t in enumerate(texts): put(ps[i],t)
    # keep all remaining empty paragraphs and their formatting, clearing only guidance.
    keep={'分析项目实践过程中':3,'分析项目的重难点和创新点':7,'拟定项目实施的预期成果':5}.get(prefix,0)
    for p in ps[len(texts):len(texts)+keep]:
        for x in list(p):
            if x.tag!=q('pPr'): p.remove(x)
    for p in ps[len(texts)+keep:]: cell.remove(p)

title='校史剧本杀微信小程序'
p=next(p for p in paras if ''.join(p.itertext()).startswith('项目名称：'))
for t in p.findall('.//w:t',ns):
    if '_' in (t.text or ''): t.text=(t.text or '').replace('_'*33,title+'_'*13)
# table name slot: first row second actual cell
cell=rows[0].findall('w:tc',ns)[1]
put(cell.find('w:p',ns),title,size=14,bold=True)
intro='本项目拟将有故事性的浙江大学校史素材转化为剧本杀式互动内容，通过微信小程序呈现，让同学在角色代入、阅读线索和推理讨论中了解校史。前期初步观察到部分同学校史了解不多、主动了解动力不足，拟通过调研进一步验证，在学校已有资料和学习活动基础上探索贴近同学兴趣的参与形式。项目从官方校史资料、正式出版物等可核验来源筛选故事，结合问卷与纸面试玩确定内容。首版实现文字图片展示、角色与章节切换、答题反馈和进度保存，形成一套可从头体验到结尾的内容。团队按资料与内容、开发、调研测试和统筹开展协作，优先使用现有设备、免费工具和学校问卷资源。项目按调研、开发、试用、评估推进，通过完成情况、卡关位置、反馈和前后知识测验评价体验及学习效果，如实说明样本与方法局限。预期形成可完整体验的小程序、测试记录、前后测数据、小程序演示与结题报告。'
intro=intro.replace('有故事性的','').replace('，形成一套可从头体验到结尾的内容','').replace('，在学校已有资料和学习活动基础上探索贴近同学兴趣的参与形式','')
assert 300<=len(intro)<=500
slot('（要求控制在', ['（要求控制在300--500字）',intro])
slot('项目的研究或实践目的',[
'实践目的：以校史故事为基础开发剧本杀式互动微信小程序，通过角色代入、线索阅读和推理讨论激发主动了解校史的兴趣。',
'',
'理论与实践意义：围绕调研、开发、试用和评估，学习提出可验证的问题、选择指标并用数据支持结论，训练工程实践、数据分析和学习效果评价能力。',
'',
'背景与发展：学校已有丰富的校史资料和学习活动。部分同学了解不多、动力不足是团队初步观察，尚待调研验证。本期不强行加入人工智能功能；后续数据和内容积累充分时，再探索学习行为分析与个性化推荐。'])
slot('项目理论分析',[
'理论学习：围绕数据分析、用户行为理解和学习效果评价，学习提出可验证问题、选择指标；相关能力与素质提升的理论综述【待补充】。',
'',
'内容与技术：从官方资料、正式出版物筛选故事，以问卷和纸面试玩确定范围；首版实现图文展示、角色与章节切换、答题反馈和进度保存。',
'组织与成本：10人团队分担内容、开发、调研测试和统筹；专业方向有电子信息、智能制造与机器人、光电及自动化、信息工程、人工智能。利用现有设备、免费工具和学校问卷资源。',
'理论实践结合：以试玩和原型验证方向，试用观察完成情况、卡关位置与反馈；以前后知识测验评价校史了解度变化。',
'',
'时间轴：2026年11—12月研读校史、整理出处；2027年1月调研与初步试玩；2—3月开发并接入内容；4月上半旬调试修复，下半旬内部试行、小范围测试并优化；5月分析效果，完成演示与结题报告。'])
slot('分析项目实践过程中',[
'主要风险：内容吸引力和实际学习效果尚不确定；开发经验、账号及发布条件需要确认，史实与艺术加工需要区分。',
'',
'应对方案：先调研、纸面试玩和简单原型，再逐步开发。优先完成完整体验流程，不开发实时多人联机、复杂后台；经验不足时收缩功能。采用可核验资料，必要时请相关老师协助核实。提前确认账号与发布条件，确有必要的费用再核定。以小范围测试和前后测验证效果，如实说明样本与方法局限。'])
slot('分析项目的重难点和创新点',[
'创新点：将可核验的校史故事转化为剧本杀式互动内容，用角色代入、线索阅读和推理讨论结合轻量微信小程序，探索贴近同学兴趣的校史学习形式。',
'重点：筛选同学感兴趣的故事，完成一套从头到尾可体验的内容，并按调研、开发、试用、评估形成验证路径。',
'难点：兼顾史实准确与艺术加工、内容吸引力和实际学习效果；在开发能力与成本约束下保障流程完整。'])
slot('拟定项目实施的预期成果',[
'预期成果：素材清单与内容构想、调研结果与首版范围、一套可完整体验的小程序原型、稳定的试用版本、测试记录及前后测数据、小程序演示与结题报告。',
'成果评估：前期调查校史了解程度和故事、体验形式偏好；试用时观察完成情况、卡关位置和反馈；结束后用简短的前后知识测验比较校史了解度变化，并说明样本与方法局限。'])
# preserve package parts; remove personal author metadata without carrying source names.
core=E.fromstring(z.read('docProps/core.xml'))
for x in core:
    if E.QName(x).localname in ('creator','lastModifiedBy'): x.text=''
out=ROOT.parent/'立项可行性报告和申请表（已填写）.docx'
with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as o:
    for item in z.infolist():
        data=z.read(item.filename)
        if item.filename=='word/document.xml': data=E.tostring(doc,xml_declaration=True,encoding='UTF-8',standalone=True)
        if item.filename=='docProps/core.xml': data=E.tostring(core,xml_declaration=True,encoding='UTF-8',standalone=True)
        o.writestr(item,data)
(ROOT/'fill-log.json').write_text(json.dumps(log,ensure_ascii=False,indent=2),encoding='utf-8')
print('intro chars',len(intro),'output',out)




