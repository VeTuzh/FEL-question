"""Condense the retained report into three pages using its native OOXML styles."""
from pathlib import Path
from copy import deepcopy
from hashlib import sha256
from zipfile import ZipFile, ZIP_DEFLATED
import json
from lxml import etree
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / 'outputs/2026-11-06_FEL光源与0.33NA_EUV光刻机照明匹配及去相干调研报告.docx'
OUTPUT = ROOT / 'outputs/2026-11-09_FEL相干性影响_去相干与EUV扫描机接口方案_三页摘要.docx'
QA = ROOT / 'tmp/fel_focus_summary_20261109'
QA.mkdir(parents=True, exist_ok=True)
SOURCE_HASH = '5b304da9e349545cad1d4e7e98e7cbbb8df42dc68653d20df6ddc52e7d4b50f9'
assert sha256(SOURCE.read_bytes()).hexdigest() == SOURCE_HASH
doc = Document(SOURCE)
ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
parts = {}
with ZipFile(SOURCE) as z:
    parts = {n: z.read(n) for n in z.namelist()}
inventory = {n: {'size': len(v), 'sha256': sha256(v).hexdigest()} for n, v in parts.items()}
(QA / 'source_parts.json').write_text(json.dumps(inventory, ensure_ascii=False, indent=2), encoding='utf-8')
(QA / 'artifact.md').write_text(f'''# Retained report formatting contract
Reference: {SOURCE.name}
SHA256: {SOURCE_HASH}
Reference: 19 pages, one section; source render reviewed in prior delivery and hash unchanged.
Reference render evidence: ../fel_sixpart_report/page-01.png through page-19.png and corrected-10.png.
Page system: A4 11906 x 16838 twips, top/bottom 1440, left/right 1800;
header 851, footer 992, grid 312. Preserve section XML and all header/footer parts.
Typography: retain all styles/theme/numbering bytes. Title SongTi 16 pt bold centered;
Heading 2 SongTi 15 pt bold; Normal FangSong 14 pt, Times New Roman Latin,
1.5 spacing, two-character indent. Reuse 10 pt caption and 10.5 pt reference roles.
Tables: source table 1, width 8300 twips, fixed layout, black borders via retained style;
light-blue D9E2F3 header, black HeiTi 9 pt bold, SongTi 9 pt body;
source padding top/bottom 80, left/right 90 twips, centered vertical alignment;
line=259/240, repeated headers, rows cannot split. No fixed row heights.
Compact summary tables disable paragraph snapToGrid so their existing 1.08 line
spacing is honored without the source's body-grid rounding; fonts, sizes, cell padding,
page geometry, colors and styles remain unchanged. This is the local pagination
adjustment needed for the user's three-page summary requirement. Set table
paragraph marks to the same 9 pt as table text, avoiding inherited 14 pt body
marks inflating row height.
Slots: replace all body paragraphs/tables/images with a concise title, three sections,
six tables, compact references and necessary evidence qualifiers. Remove the three
body-image relationships and their media parts; retain original external relationships
and display the 16 relevant original links, omitting the unrelated KEK source item.
Allow manual page breaks for three-page packaging. Preserve everything outside body,
the three image relationships/media, content types for removed images, and core/app metadata.
Package inventory: source_parts.json; all other parts must be byte-for-byte identical.
Page 1: coherence effects on EUV exposure, LER/LWR, CD/EPE and process window;
spatial/temporal coherence, pupil sigma, speckle and dose-integration metrics.
Page 2: four decoherence families, quantitative design conditions and combination choices.
Page 3: two scanner interface routes, optical/power/pupil/time/thermal matching fields;
seven patents are distributed according to function, followed by relevant source links.
Fidelity gates: three clean native Word pages, matching fonts and table components,
no missing quantitative source values, unchanged reference hash and preserve-only parts.
''', encoding='utf-8')

def el(name, **attrs):
    e = OxmlElement('w:' + name)
    for k, v in attrs.items(): e.set(qn('w:' + k), str(v))
    return e

tree = etree.fromstring(parts['word/document.xml'])
body = tree.find('w:body', ns)
section = deepcopy(body.find('w:sectPr', ns))
for x in list(body): body.remove(x)
body.append(section)
prototypes = {s: deepcopy(next(p._p for p in doc.paragraphs if p.style.name == s))
              for s in ['Title', 'Heading 2', 'Normal']}
caption_proto = deepcopy(next(p._p for p in doc.paragraphs if p.text.startswith('表1  ')))
reference_proto = deepcopy(next(p._p for p in doc.paragraphs if p.text.startswith('[1] ASML.')))

def put(e): body.insert(len(body)-1, e); return e

def fresh(proto):
    p = deepcopy(proto)
    for x in list(p):
        if x.tag != qn('w:pPr'): p.remove(x)
    return p

def text_run(text, size=None):
    r = el('r')
    if size:
        rp=el('rPr'); rp.append(el('sz', val=int(size*2))); rp.append(el('color',val='000000')); r.append(rp)
    for i, line in enumerate(text.split('\n')):
        if i: r.append(el('br'))
        t=el('t'); t.set('{http://www.w3.org/XML/1998/namespace}space','preserve'); t.text=line; r.append(t)
    return r

def para(text, role='Normal', break_before=False):
    p=fresh(prototypes[role]); p.append(text_run(text))
    pp=p.find('w:pPr',ns)
    if pp is None: pp=el('pPr'); p.insert(0,pp)
    if break_before: pp.append(el('pageBreakBefore'))
    put(p); return p

def note(text):
    p=fresh(reference_proto); p.append(text_run(text,10.5)); put(p); return p

def caption(text):
    p=fresh(caption_proto); p.append(text_run(text,10)); put(p)

tbl_proto=deepcopy(doc.tables[0]._tbl)
head_proto=deepcopy(doc.tables[0].cell(0,0)._tc)
cell_proto=deepcopy(doc.tables[0].cell(1,1)._tc)

def table(headers, rows, widths):
    assert sum(widths)==8300
    t=deepcopy(tbl_proto)
    for tr in list(t.findall(qn('w:tr'))): t.remove(tr)
    grid=t.find(qn('w:tblGrid'))
    for g in list(grid): grid.remove(g)
    for w in widths: grid.append(el('gridCol',w=w))
    for i, values in enumerate([headers,*rows]):
        tr=el('tr'); trp=el('trPr'); trp.append(el('cantSplit'))
        if i==0: trp.append(el('tblHeader'))
        tr.append(trp)
        for j, value in enumerate(values):
            c=deepcopy(head_proto if i==0 else cell_proto)
            c.find('w:tcPr/w:tcW',ns).set(qn('w:w'),str(widths[j]))
            for old in list(c.findall('w:p',ns)): c.remove(old)
            p=el('p'); pp=el('pPr')
            pp.append(el('spacing',before=0,after=0,line=259,lineRule='auto'))
            pp.append(el('snapToGrid',val=0))
            pp.append(el('ind',firstLine=0,firstLineChars=0))
            pp.append(el('jc',val='center' if i==0 or j==0 else 'left'))
            pp.append(el('keepNext',val=1 if i==0 else 0)); p.append(pp)
            mark=el('rPr'); mark.append(el('sz',val=18)); mark.append(el('szCs',val=18)); pp.append(mark)
            r=text_run(value); rp=el('rPr')
            rp.append(el('rFonts',ascii='Times New Roman',hAnsi='Times New Roman',eastAsia='黑体' if i==0 else '宋体'))
            rp.append(el('sz',val=18)); rp.append(el('color',val='000000'))
            if i==0: rp.append(el('b'))
            r.insert(0,rp); p.append(r); c.append(p); tr.append(c)
        t.append(tr)
    put(t)

para('FEL相干性影响 去相干与扫描机接口方案','Title')
para('一 FEL光源相干性对EUV光刻的影响','Heading 2')
caption('表1  相干性影响与光刻评价指标')
table(['影响','主要表现与作用环节','关键指标与证据'],[
 ('照明与局部剂量','互相干通道的光程/相位差形成交叉干涉；掩模粗糙度散射与成像光叠加产生散斑，造成局部明暗起伏。','狭缝均匀性、C、曝光内局部剂量；平均功率与局部均匀性分别验收。[7,8,12]'),
 ('线边与线宽粗糙度','散斑改变边缘附近的曝光强度，使边界沿线条方向摆动，产生边缘不平直及局部线宽变化。','LER/LWR及频谱；FEL仿真中，相干增强伴随散斑/粗糙度上升，窄图形更敏感。[8]'),
 ('尺寸与位置精度','局部强度起伏可改变CD及其均匀性；边缘两侧扰动不对称时，可引起位置偏移。','CD/CDU、EPE；阈值模型δx≈−δI/(dI/dx)，用于小扰动估算；斜率越小，边缘位移越大（工程推论）。'),
 ('对比度与工艺窗口','成像由掩模衍射光干涉形成；效果取决于照明角度、瞳形、图形周期/方向及离焦。','空中像对比度、NILS、焦深和曝光宽容度；按同图形/瞳形/剂量比较，保留目标配方成像能力。[2,7,8]'),
 ('重复性与相位敏感','热漂移、位移或指向变化可改变相干干涉图样；高度相关的连续图样会保留空间起伏。','通道可见度、图样相关ρij、相关时间、积分剂量；能量样本数与独立图样数分别计量。[6,12]'),
], [1400,3750,3150])
caption('表2  相干性与曝光平均的关键参数')
table(['参数','定义与定量关系','验收口径'],[
 ('空间与时间相干','空间互相干、时间相干函数γij(Δt)、相干时长τc、通道时差Δt及光程差ΔL。','在掩模/对应共轭面测干涉或相关性，注明工作谱、采样位置和曝光时窗。'),
 ('照明部分相干σ','圆形σ=照明NA/投影NA（同一共轭侧）；环形记录σ内/σ外，自由瞳形保留二维角度与权重。','σ描述照明角域；光场互相干度单独测量。目标瞳面及偏振按图形配方确定。[2]'),
 ('散斑与独立图样','C=std(I)/mean(I)；等权、互不相关且单幅C相同，C≈C1/√N；5%→1%需约25幅独立图样。','测曝光积分后的C、权重wi与相关ρij；有相关或权重失衡时用有效独立样本数评估。'),
 ('脉冲与剂量积分','N脉冲≈f重频×T驻留。演示：10 MHz×1 ms=10000脉冲；单脉冲能量相对标准差20%→累计约0.2%。','独立等权能量样本的假设算例；实际采用时间相关、扫描权重及漏脉冲数据。'),
], [1450,3800,3050])
note('CD/CDU：关键尺寸/均匀性；EPE：边缘位置误差；LER/LWR：线边/线宽粗糙度；NILS：归一化像对数斜率。FEL研究[8]为仿真；DPP实验[7]的影响依赖粗糙度及照明。')

para('二 FEL光源去相干方案','Heading 2',break_before=True)
caption('表3  四类去相干方式及关键设计条件')
table(['方式','方案与关键参数','适用作用与限制'],[
 ('分束差分时延','将光分入不同照明通道，利用分面几何或延时元件设置Δt=|L2−L1|/c；EP2823360B1权2–4要求照明场各点任意通道对Δt>τc。[12]\nxLight算例τc≈3 fs、相干长度≈0.90 μm（约1 μm）；通道光程差>10 μm对应时差>33.4 fs。[5]','主要降低通道间交叉干涉。验证全场最小时差、互相干及权重；采用子脉冲时间分离时另核对脉宽和尾部。像面散斑仍需实测。[6]'),
 ('动态相位调制','运动反射面让光束随时间采样不同微结构。ASML US11984236B2权1–3：双粗糙反射面，第二面沿传播方向布置凸起/凹陷，可旋转，表面速度约≥1 m/s。[13]\n华为CN115343915A权1、11–13：相邻微面高度差0<|Δh|≤kλ，k为正整数，配合旋转；为申请公开文本。[14]','曝光内更新相位与散斑。匹配表面速度、微结构尺度、光斑尺寸/转轴位置和驻留时窗；测图样相关时间、抖动、反射效率与热负荷。'),
 ('角度多样化与扫描','光束依次或并行照向不同场分面—瞳分面通道，按配方设置角度、停留时间、能量及偏振。EP2883112B1权1、4、5给出FEL偏振设定与扫描/双分面同步结构。[15]','将目标瞳面与时间平均结合。不同角度产生的图样需足够去相关；核对角域权重、偏振误差、控制带宽及实际散斑收益。'),
 ('反射式匀光与多路径','掠反射管通过多路径光程差和角域混合调理光场。仿真输入13.5 nm、直径200 μm、发散度20 mrad；选定结构内径1 mm、总长≥0.6 m、倾斜角10 mrad，能量透过率37.6%。[9]','兼顾强度/角域均匀化和路径混合。按实际光谱与相干模型测互相干及像面散斑；完整相位传播、检测平均、反射损耗与热负荷共同评价。'),
], [1400,4300,2600])
caption('表4  去相干技术组合与选用条件')
table(['组合','选用条件','联合评价'],[
 ('时延＋目标角度分配','作为首轮比较基线，先测既有照明通道的光程、互相干及瞳面覆盖，确定所需延时和角度采样。','通道可见度、曝光内C、图样相关ρij、瞳面误差及有效带内功率。'),
 ('时延＋动态相位/扫描','通道干涉已降低而像面散斑仍保持较高相关时，增加曝光期间的图样更新。','独立图样数、LER/LWR、CD/EPE、剂量稳定性，以及运动引入的抖动和损耗。'),
 ('匀光管＋双分面照明','需要同时调节入口强度和角域时，将匀光模块与后续瞳面/狭缝分配联合比较。','出口均匀性、接收角域、互相干、透过率、污染及长期热稳定性。'),
], [1650,3850,2800])
note('配置比较统一源谱、掩模、瞳面、剂量和曝光时窗；目标是满足图形指标时获得足够有效功率。剂量反馈作为配套控制，在第三页接口方案中统一配置。')

para('三 FEL光源与EUV扫描机接口方案','Heading 2',break_before=True)
caption('表5  两类0.33 NA扫描机接入方案')
table(['方案','接口光路与系统分工','实施条件与评价'],[
 ('A 经IF适配既有照明器','FEL→诊断/真空输运→扩束扩角/聚焦→IF→双分面照明→掩模狭缝→投影→晶圆。先测通道光程/瞳面覆盖，再选去相干位置。','以机型IF接收图为边界；IF前损耗计入交接功率，IF后计入照明器。取得接收、脉冲及通道数据。'),
 ('B 与照明器协同设计','FEL→输运/分支→分束或扫描＋时延/偏振配方→双分面/中继→掩模→晶圆。由目标瞳面/狭缝反推通道与输入。','取得内部光学/控制协同条件；比较有效曝光收益、配方、控制复杂度及维护。固定瞳形验证后扩展扫描/多机。'),
], [1550,3950,2800])
caption('表6  接口模块 关键参数与典型依据')
table(['接口环节','关键参数与公开方案','必须匹配的指标'],[
 ('光束转换与IF聚焦','US10928734B2权1–3：结构化双掠入射镜＋空心波导；首镜角1–10 mrad（相对表面），入射面夹角>45°，次镜角≥首镜2倍。[10]\nUS9958783B2权1、13、15：椭球长半轴a/抛物面焦距f>50，公共束分支及下游双分面。[11]','IF焦点/光斑、主光线、位置—角度分布、波前、首镜覆盖、带内接收率；核对专利散射角约束。'),
 ('光谱与带内功率','平台13.5 nm、投影NA 0.33，最大扫描场26 mm × 33 mm。[1]\n历史2%谱窗口径≈0.27 nm全宽。[4] xLight公开设计6 kW源端→1 kW IF，总比值约16.7%。[17]','按机型谱窗和接收图积分；预算演示6×0.80×1.00×0.60×0.35=1.008 kW（16.8%），各段系数为假设。'),
 ('瞳面与掩模狭缝','双分面通道共同确定二维瞳形、角域权重、偏振及狭缝叠加；去相干与照明配方联合设计。[2,12,15]','目标σ/瞳形、填充率、视场依赖、狭缝照度和通道互相干；具体容差由设备及工艺方给出。'),
 ('脉冲与剂量反馈','US9823572B2权1：FEL反馈＋可变光衰减；权3：各机独立衰减；权4–7：第二反馈回路与机内传感。[16]','脉宽、重频、单脉冲能量/面能量密度、宏脉冲占空比、漏脉冲、扫描驻留、回路带宽与多机串扰。'),
 ('热 真空与验收','吸收功率、照射足迹、冷却、热漂移、反射率衰减和污染；IF后计入照明、掩模与投影损耗。','IF接收→瞳面/狭缝→同剂量像面/胶试片；验收C、NILS、CD均匀性、LER/LWR及长时稳定性。'),
], [1450,4200,2650])
note('预算系数依次为公共输运、支路份额、支路镜组/整形、IF前转换/接收；分配份额只计一次。机型IF公差、脉冲限制、目标瞳面、剂量窗口及图形阈值待确认。')

# Compact traceability index. Reuse the exact 17 original hyperlink relationships.
rels=etree.fromstring(parts['word/_rels/document.xml.rels'])
relmap={r.get('Id'):r.get('Target') for r in rels}
source_links=[]
for p in doc.paragraphs:
    if p.text.startswith('['):
        links=p._p.findall('w:hyperlink',ns)
        if links: source_links.append(links[0].get(qn('r:id')))
assert len(source_links)==17
labels=['ASML规格','ZEISS/ASML照明','KEK 2023','Cymer谱窗','xLight光源',
        'FEL 2024时延','Pret 2012实验','Wei 2025仿真','匀光管2024',
        'US10928734B2','US9958783B2','EP2823360B1','US11984236B2',
        'CN115343915A','EP2883112B1','US9823572B2','xLight集成预算']
for indices in ([0,1,3,4],[5,6,7,8],[9,10,11,12],[13,14,15,16]):
    p=fresh(reference_proto)
    if indices[0]==0:p.append(text_run('原文索引  ',10.5))
    for i in indices:
        h=el('hyperlink'); h.set(qn('r:id'),source_links[i])
        r=text_run(f'[{i+1}] {labels[i]}',10.5)
        rp=r.find('w:rPr',ns);rp.append(el('u',val='single'));h.append(r);p.append(h)
        p.append(text_run('  ',10.5))
    put(p)

# Keep the native page-number fields; Word updates PAGE on pagination.
image_type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
removed=[]
for r in list(rels):
    if r.get('Type')==image_type:
        removed.append('word/'+r.get('Target')); rels.remove(r)
for name in removed: parts.pop(name,None)
ct=etree.fromstring(parts['[Content_Types].xml'])
for child in list(ct):
    if child.get('PartName','').lstrip('/') in removed: ct.remove(child)
parts['[Content_Types].xml']=etree.tostring(ct,xml_declaration=True,encoding='UTF-8',standalone=True)
parts['word/document.xml']=etree.tostring(tree,xml_declaration=True,encoding='UTF-8',standalone=True)
parts['word/_rels/document.xml.rels']=etree.tostring(rels,xml_declaration=True,encoding='UTF-8',standalone=True)
core=etree.fromstring(parts['docProps/core.xml'])
dc={'dc':'http://purl.org/dc/elements/1.1/'}
core.find('dc:title',dc).text='FEL相干性对EUV光刻的影响 去相干与扫描机接口方案'
parts['docProps/core.xml']=etree.tostring(core,xml_declaration=True,encoding='UTF-8',standalone=True)
app=etree.fromstring(parts['docProps/app.xml'])
appns={'a':'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'}
app.find('a:Pages',appns).text='3'
parts['docProps/app.xml']=etree.tostring(app,xml_declaration=True,encoding='UTF-8',standalone=True)
with ZipFile(OUTPUT,'w',compression=ZIP_DEFLATED) as z:
    for name,data in parts.items(): z.writestr(name,data)
assert sha256(SOURCE.read_bytes()).hexdigest()==SOURCE_HASH
print(OUTPUT)
print('tables',len(Document(OUTPUT).tables),'hyperlinks',len(Document(OUTPUT).element.xpath('//w:hyperlink')))
