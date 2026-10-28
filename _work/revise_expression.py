from pathlib import Path
from lxml import etree as E
import zipfile,hashlib,json
R=Path(__file__).resolve().parent
src=R.parent/'立项可行性报告和申请表1.0（字体标点修正）.docx'
out=R.parent/'立项可行性报告和申请表1.0（表达修正）.docx'
z=zipfile.ZipFile(src); d=E.fromstring(z.read('word/document.xml'))
ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
paras=d.findall('.//w:body//w:p',ns)
get=lambda p:''.join(p.xpath('.//w:t/text()',namespaces=ns))
theory=next(get(p) for p in paras if get(p).startswith('理论学习：'))
risk_start='应对方案：先调研了解同学们感兴趣的校史方向，以初步保证内容吸引性，后续以纸面试玩和简单原型进行学习效果的体验，再逐步开发。'
changes={
'实践目的：':'实践目的：以校史故事为基础开发剧本杀式互动微信小程序，通过角色代入、线索阅读和推理讨论，激发同学主动了解校史的兴趣，引导同学理解浙江大学的发展历程及求是创新精神。',
'理论与实践意义：':'理论与实践意义：围绕调研、开发、试用和评估，学习提出可验证的问题、选择指标并用数据支持结论，拟提升工程实践、数据分析和学习效果评价能力。通过校史研读与剧本创作，深化对校史的认识，训练史料研读与考据、叙事与创意写作、用户洞察与同理心等能力，并在项目实施中培养闭环验证的工程思维。',
'背景与发展：':'背景与发展：学校已有丰富的校史资料和学习活动，但团队初步观察发现部分同学了解不多、动力不足，后续将通过调研进一步验证。项目拟根据同学感兴趣的剧本主题，借助AI辅助编程完成小程序开发，设计沉浸式校史互动体验，并安排专人开展相关内容宣传。现阶段不强行加入人工智能功能，后续数据和内容积累充分时，再探索基于兴趣分析的个性化剧本推荐。',
'时间轴：':'时间轴：2026年11月至12月中旬研读校史、整理出处；2026年12月下旬至2027年1月开展调研与初步试玩；2027年2—3月开发并接入内容；2027年4月上半旬调试修复，下半旬内部试行、小范围测试并优化；2027年5月分析效果，完成演示与结题报告。',
'应对方案：':risk_start+'遇到技术问题时，先借助开发文档及AI辅助工具排查问题，必要时请教指导教师；对暂时无法解决的问题，适当收缩功能，优先完成完整体验流程。采用可核验资料，必要时请相关教师协助核实，并在发布艺术加工内容前征求其意见。提前确认账号与发布条件，确有必要的费用再核定。以小范围测试和前后测验证效果，如实说明样本与方法局限。',
'创新点：':'创新点：将可核验的校史故事转化为剧本杀式互动内容，以微信小程序承载角色代入、线索阅读与推理讨论，探索贴近同学兴趣的校史学习形式。',
'重点：':'重点：筛选同学感兴趣的校史故事，完成一套可从头到尾体验的互动内容，并通过调研、开发、试用和评估形成验证路径。',
'难点：':'难点：在区分史实与艺术加工的基础上，兼顾内容吸引力与实际学习效果，并在开发能力与成本约束下保障体验流程完整。',
'成果评估：':'成果评估：前期调查同学的校史了解程度，以及对校史主题与体验形式的偏好；试用时观察完成情况、卡关位置并收集反馈；结束后通过简短的前后知识测验比较校史了解度变化，并说明样本与方法局限。',
'打印纸面试玩':'用于必要的试玩材料打印，采用普通A4黑白双面，其余材料优先电子化',
'用于结题时打印':'用于结题阶段宣传单或三折页的印制，开展项目宣传与成果展示',
'用于小程序代码辅助编写':'用于小程序开发、调试与测试所需的模型API调用或开发工具订阅；按实际购买方式择一核算，仅用于本项目',
}
log=[]; hits=set()
for p in paras:
    old=get(p); new=old
    for prefix, replacement in changes.items():
        if old.startswith(prefix):new=replacement;hits.add(prefix);break
    if old.startswith(('本项目拟将','内容与技术：')):
        new=new.replace('做到初步的玩家交互和剧本杀游戏体验','支持角色代入、线索阅读与推理讨论，形成可完整体验的剧本式互动流程')
    if old==new:continue
    nodes=p.findall('.//w:t',ns);assert nodes
    nodes[0].text=new
    for node in nodes[1:]:node.text=''
    log.append({'before':old,'after':new})
assert hits==set(changes)
assert next(get(p) for p in paras if get(p).startswith('理论学习：'))==theory
assert next(get(p) for p in paras if get(p).startswith('应对方案：')).startswith(risk_start)
assert 300<=len(next(get(p) for p in paras if get(p).startswith('本项目拟将')))<=500
original=E.fromstring(z.read('word/document.xml'))
# Removing only text payloads leaves identical XML: typography, tables and paragraph settings untouched.
for root in (original,d):
    for t in root.findall('.//w:t',ns):t.text=''
assert E.tostring(original)==E.tostring(d)
# Restore the selected text changes to the pristine source tree.
result=E.fromstring(z.read('word/document.xml'))
mapping={x['before']:x['after'] for x in log}
for p in result.findall('.//w:body//w:p',ns):
    old=get(p)
    if old in mapping:
        ts=p.findall('.//w:t',ns);ts[0].text=mapping[old]
        for t in ts[1:]:t.text=''
with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as o:
    for i in z.infolist():
        data=z.read(i.filename)
        if i.filename=='word/document.xml':data=E.tostring(result,encoding='UTF-8',xml_declaration=True,standalone=True)
        o.writestr(i,data)
(R/'expression-changes.json').write_text(json.dumps(log,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'paragraphs_changed':len(log),'source_sha256':hashlib.sha256(src.read_bytes()).hexdigest(),'output':str(out)},ensure_ascii=False))

