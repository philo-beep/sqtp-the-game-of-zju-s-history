from pathlib import Path
import zipfile,json,hashlib,re
from lxml import etree as E
R=Path(__file__).resolve().parent; root=R.parent
f=root/'立项可行性报告和申请表（已填写）.docx'
with zipfile.ZipFile(f) as z:
    data={i.filename:(i,z.read(i.filename)) for i in z.infolist()}
key='docProps/custom.xml'; c=E.fromstring(data[key][1])
for p in list(c):
    if p.get('name') in ['ICV','KSOTemplateDocerSaveRecord']: c.remove(p)
data[key]=(data[key][0],E.tostring(c,encoding='UTF-8',xml_declaration=True,standalone=True))
with zipfile.ZipFile(f,'w',zipfile.ZIP_DEFLATED) as z:
    for i,b in data.values(): z.writestr(i,b)
z=zipfile.ZipFile(f); bz=zipfile.ZipFile(R/'baseline.docx')
n={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
d=E.fromstring(z.read('word/document.xml')); b=E.fromstring(bz.read('word/document.xml'))
for xp in ['.//w:sectPr','.//w:tblPr','.//w:tblGrid','.//w:trPr','.//w:tcPr']:
    assert [E.tostring(x) for x in d.findall(xp,n)]==[E.tostring(x) for x in b.findall(xp,n)],xp
changed=[k for k in bz.namelist() if z.read(k)!=bz.read(k)]
assert set(changed)=={'word/document.xml','docProps/core.xml','docProps/custom.xml'}
text=''.join(d.xpath('//w:t/text()',namespaces=n))
src=(R/'source.txt').read_text(encoding='utf-8')
names=re.findall(r'(?:立项人：|成员：\s*1\s*、|\d\s*、)([\u4e00-\u9fff]{2,4})\s*专业',__import__('unicodedata').normalize('NFKC',src))
assert not any(x in text for x in names)
assert '8000' not in text and '公开论文' not in text
assert not d.findall('.//w:del',n) and not d.findall('.//w:commentRangeStart',n)
j=json.loads((R/'format.json').read_text(encoding='utf-8-sig'))
align={0:'左对齐',1:'居中',2:'右对齐',3:'两端对齐'}
rules={0:'单倍行距',1:'1.5倍行距',2:'双倍行距',3:'最小值',4:'固定值',5:'多倍行距'}
out=['# 字体规范清单与填写核对记录','',
'依据：完整读取原 DOC 正文、表格、页眉页脚，并用 Word 原生转换保留可编辑模板；原 DOC 未覆盖。',
'', '## 硬性填写规范','',
'- 项目名称不超过30个汉字；项目简介300—500字，实填307字。',
'- A4纸正反页打印。',
'- 汉字限仿宋_GB2312、仿宋或宋体；英文、数字使用 Times New Roman。',
'- 不改栏目、标题、表格结构与版式；保留原段落对齐、缩进、行距。',
'- 内容设计要求包含理论知识、模拟训练、实践锻炼三模块。资料筛选与理论学习、纸面试玩与原型、开发测试及效果评价分别对应相关内容；材料未提供正式理论综述，标为【待补充】。',
'', '## 栏目位置及处理','',
'|位置|栏目|处理|', '|---|---|---|',
'|第1页封面正文|项目名称|填写|',
'|第1页封面正文|所在院系、立项人、专业、联系电话、电子邮箱、填表日期|留空；个人信息或签署日期。封面专业是个人字段，团队专业方向改放可行性分析正文|',
'|第3页表格首行|项目名称|填写|',
'|第3页表格|项目负责人：姓名、所在院系、学号、联系方式|全部留空|',
'|第3页表格|项目主要参与人员：9行，每行姓名、学号、所在院系、联系方式|全部留空|',
'|第3页表格|指导教师：姓名、职称职务、所在单位、联系方式|全部留空，材料未提供非个人字段依据|',
'|第3页表格|项目简介|填写，保留300—500字要求提示|',
'|第4页表格第一部分|实践目的、理论与实践意义、背景现状与发展趋势|填写|',
'|第4页表格第二部分|理论分析、实践分析、理论实践结合、整体时间轴|填写；正式理论综述待补充|',
'|第4页表格第三部分|风险与不确定因素、应对方案|填写|',
'|第5页表格第四部分|主要创新点、重点、难点|填写|',
'|第5页表格第五部分|预期成果、成果评估|填写|',
'|第5页表格|预算5行：类别、金额、备注；合计金额|全部留空；来源未提供可核算单价、数量或金额|',
'|第6页表格|指导教师意见、签名、单位、日期|全部留空|',
'|第6页表格|评审小组意见、负责人、日期|全部留空|',
'|第6页表格|工作小组意见、备案、日期|全部留空|',
'|页眉页脚|无待填写栏目；页眉为空段落，未发现页脚填写项|保留|',
'', '## 字体执行清单','',
'|位置|原格式或要求|执行|','|---|---|---|',
'|封面项目名称|宋体14磅，加粗，居中，1.5倍行距，段前后0|继承原格式|',
'|表格项目名称|仿宋_GB2312 14磅，加粗，居中，单倍，段前后0|同原格式|',
'|简介填写内容|引导文字仿宋_GB2312 12磅、不加粗、两端对齐、单倍|同栏目12磅格式，英文数字Times New Roman|',
'|第一至三部分填写内容|引导文字12磅，主要为楷体_GB2312，与填写说明规定字库不一致|以明确规定为准，用仿宋_GB2312 12磅、不加粗；继承对齐、单倍行距和原有编号|',
'|第四、五部分填写内容|引导文字仿宋_GB2312 12磅、不加粗、两端对齐、单倍|按该格式统一填写|',
'|空白及签署字段|不同位置存在14/12/15磅、不同加粗与对齐|保持空白及原格式，逐段记录见后表|',
'', '## 示例和引导文字','',
'“项目实践报告一篇（8000字以上）、公开论文发表一篇等”为示例，已用材料中的成果替换，未作为承诺。',
'其余“简述项目目的……”及第一至五部分填写引导已替换；所有正式栏目标题保留。',
'', '## 自检结果','',
'- Word原生导出并逐页查看6页图像，无文字截断、重叠或缺字；主要栏目与模板保持相同分页。',
'- 表格属性、网格、所有行属性和单元格属性、节与页面设置均通过XML逐项相等检查。',
'- 保留全部标题；只在填写单元格内用正文占用原空白行，调整剩余空白段落以容纳内容，未改变表格行列或合并结构。',
'- 除正文及隐私元数据外，转换基线包中的所有部件逐字节不变。',
'- 原作者、最后修改者及模板软件的用户追踪元数据已清除。',
'- 材料姓名未写入，未新增学号、联系方式、签署日期、预算或论文承诺。',
'- 渲染辅助脚本因本机无LibreOffice未能执行；改用已安装Word原生PDF导出并转逐页PNG完成视觉检查。',
'', '## 原模板逐段字体证据','',
'以下段落编号按Word完整正文及表格顺序（含空段），用于定位。字号“混合”表示原段落含多种字号，不据此臆造统一值。',
'', '|段落|内容或位置标识|字体|磅|加粗|对齐|行距|段前/后（磅）|', '|---|---|---|---|---|---|---|---|']
for a in j:
    t=a['text'].replace('\r','').replace('\x07','').replace('\x0c','').strip() or '（空白填写/间隔段）'
    t=t.replace('|','／')
    out.append(f"|{a['i']}|{t}|{a['font'] or '混合'}|{a['size'] if a['size']<999999 else '混合'}|{ {-1:'是',0:'否'}.get(a['bold'],'混合')}|{align.get(a['align'],a['align'])}|{rules.get(a['lineRule'],a['lineRule'])}（记录值{a['line']}）|{a['before']}/{a['after']}|")
(root/'字体规范清单与填写核对.md').write_text('\n'.join(out),encoding='utf-8')
(R/'audit-result.json').write_text(json.dumps({'changed_parts':changed,'privacy_name_count_checked':len(names),'table_and_page_structure':'unchanged','source_sha256':hashlib.sha256((root/'立项可行性报告和申请表.doc').read_bytes()).hexdigest()},ensure_ascii=False,indent=2),encoding='utf-8')
print('Audit passed; checklist saved.')
