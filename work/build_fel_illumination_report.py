from pathlib import Path
from copy import deepcopy
from math import atan2, cos, sin
from hashlib import sha256
from zipfile import ZipFile
import os

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Pt, Cm, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE
from PIL import Image, ImageDraw, ImageFont

root=Path(__file__).resolve().parents[1]
base=root/'tmp/fel_sixpart_report'
base.mkdir(parents=True,exist_ok=True)
reference=root/'references/format/先进光刻_光源与照明系统_格式参考.docx'
doc=Document(reference)
ref_table=deepcopy(doc.tables[3]._tbl)
sample_header=deepcopy(doc.tables[3].cell(0,0)._tc)
sample_body=deepcopy(doc.tables[3].cell(1,0)._tc)
for rel in list(doc.part.rels.values()):
    if rel.reltype in (RELATIONSHIP_TYPE.IMAGE,RELATIONSHIP_TYPE.HYPERLINK):
        doc.part.drop_rel(rel.rId)
for child in list(doc._element.body):
    if child.tag!=qn('w:sectPr'):doc._element.body.remove(child)
doc.core_properties.title='FEL光源与0.33 NA EUV光刻机照明匹配及去相干调研报告'
doc.core_properties.author='FEL 光刻光源研究'
doc.core_properties.subject='源端特性 照明需求 相干性影响 公开技术与专利 系统匹配'
doc.core_properties.comments=''
doc.core_properties.keywords='FEL, EUV, 0.33 NA, illumination, coherence'

# Native source styles, page geometry, header rule and page field are retained.
title_style=doc.styles['Title']
for tag in ('pPr','rPr'):
    old=title_style.element.find(qn('w:'+tag))
    if old is not None:title_style.element.remove(old)
    original=doc.styles['Heading 1'].element.find(qn('w:'+tag))
    if original is not None:title_style.element.append(deepcopy(original))
for run in doc.sections[0].footer.paragraphs[0].runs:
    if run._r.find(qn('w:t')) is not None:pass

def el(name,**attrs):
    x=OxmlElement('w:'+name)
    for k,v in attrs.items():x.set(qn('w:'+k),str(v))
    return x

def paragraph(text,style=None,indent=True):
    p=doc.add_paragraph(text,style=style)
    p.paragraph_format.widow_control=True
    if style is None and indent:p.paragraph_format.first_line_indent=Pt(28)
    if style is None and not indent:
        p.paragraph_format.first_line_indent=Pt(0)
        p._p.get_or_add_pPr().append(el('ind',firstLine=0,firstLineChars=0))
    return p

def section(n,text):
    p=doc.add_paragraph(f'{n}. {text}',style='Heading 2')
    p.paragraph_format.keep_with_next=True

def sub(n,text):
    p=doc.add_paragraph(f'{n} {text}',style='Heading 3')
    p.paragraph_format.keep_with_next=True

def detail(text):
    p=doc.add_paragraph(text,style='Heading 4')
    p.paragraph_format.keep_with_next=True

def equation(text):
    p=paragraph(text,indent=False)
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before=Pt(3)
    p.paragraph_format.space_after=Pt(3)
    for r in p.runs:r.font.name='Times New Roman';r.font.size=Pt(12)
    return p

def caption(text,keep=True):
    p=paragraph(text,indent=False);p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next=keep
    p.paragraph_format.line_spacing=1.2
    p.paragraph_format.space_before=Pt(3);p.paragraph_format.space_after=Pt(3)
    for r in p.runs:r.font.size=Pt(10)
    return p

def table(headers,rows,widths):
    template=deepcopy(ref_table)
    for tr in list(template.findall(qn('w:tr'))):template.remove(tr)
    grid=template.find(qn('w:tblGrid'))
    for e in list(grid):grid.remove(e)
    for width in widths:grid.append(el('gridCol',w=width))
    pr=template.find(qn('w:tblPr'))
    w=pr.find(qn('w:tblW'));w.set(qn('w:w'),'8300');w.set(qn('w:type'),'dxa')
    for ri,values in enumerate([headers,*rows]):
        row=el('tr');rp=el('trPr');rp.append(el('cantSplit'))
        if ri==0:rp.append(el('tblHeader'))
        row.append(rp)
        for j,value in enumerate(values):
            c=deepcopy(sample_header if ri==0 else sample_body)
            cp=c.find(qn('w:tcPr'))
            cp.find(qn('w:tcW')).set(qn('w:w'),str(widths[j]))
            for p in list(c.findall(qn('w:p'))):c.remove(p)
            p=el('p');pp=el('pPr')
            pp.append(el('spacing',before=0,after=0,line=259,lineRule='auto'))
            pp.append(el('ind',firstLine=0,firstLineChars=0))
            pp.append(el('jc',val='center' if ri==0 or j==0 else 'left'))
            pp.append(el('keepNext',val='1' if ri==0 else '0'))
            p.append(pp)
            r=el('r');rpr=el('rPr')
            rpr.append(el('rFonts',ascii='Times New Roman',hAnsi='Times New Roman',eastAsia='黑体' if ri==0 else '宋体'))
            rpr.append(el('sz',val=18))
            rpr.append(el('color',val='000000'))
            if ri==0:rpr.append(el('b'))
            r.append(rpr);tx=el('t');tx.text=value;r.append(tx);p.append(r);c.append(p);row.append(c)
        template.append(row)
    doc._element.body.insert(len(doc._element.body)-1,template)

def picture(path,text):
    p=paragraph('',indent=False);p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next=True
    p.paragraph_format.line_spacing=1
    p.add_run().add_picture(str(path),width=Cm(14.5))
    caption(text,keep=False)

def link(p,url):
    rid=doc.part.relate_to(url,RELATIONSHIP_TYPE.HYPERLINK,is_external=True)
    h=el('hyperlink');h.set(qn('r:id'),rid)
    r=el('r');rp=el('rPr')
    rp.append(el('rFonts',ascii='Times New Roman',hAnsi='Times New Roman',eastAsia='仿宋'))
    rp.append(el('sz',val=21));rp.append(el('color',val='000000'))
    rp.append(el('u',val='single'));r.append(rp);t=el('t');t.text='原文';r.append(t);h.append(r);p._p.append(h)

def ref(n,name,url):
    p=paragraph(f'[{n}] {name}  ',indent=False)
    p.paragraph_format.line_spacing=1.2;p.paragraph_format.space_after=Pt(4)
    for r in p.runs:r.font.size=Pt(10.5)
    link(p,url)

fontfile=Path(os.environ['WINDIR'])/'Fonts'/'simfang.ttf'
def text(draw,xy,string,size=34,anchor='mm'):
    draw.text(xy,string,font=ImageFont.truetype(str(fontfile),size),fill='#111111',anchor=anchor)
def node(draw,box,title,desc):
    x,y,z,w=box;draw.rounded_rectangle(box,radius=9,outline='#5B6D7F',width=3,fill='white')
    text(draw,((x+z)/2,y+43),title,35);text(draw,((x+z)/2,y+91),desc,29)
def arrow(draw,start,end):
    draw.line((start,end),fill='#263E54',width=5)
    ang=atan2(end[1]-start[1],end[0]-start[0])
    draw.polygon([end,*[(end[0]-18*cos(ang+a),end[1]-18*sin(ang+a)) for a in (-.5,.5)]],fill='#263E54')
def chain():
    im=Image.new('RGB',(1510,490),'white');d=ImageDraw.Draw(im)
    upper=[('FEL 输出 R0','谱  脉冲  相干'),('输运与整形','转向  调理  监测'),('中间焦点 R1','接收位置与角域'),('照明器','场分面与瞳分面')]
    lower=[('晶圆 R4','剂量与图形'),('投影物镜','晶圆侧 NA 0.33'),('掩模狭缝 R3','照度与残余干涉'),('照明瞳面 R2','角度  权重  偏振')]
    for i,labels in enumerate(upper):node(d,(10+i*380,15,350+i*380,150),*labels)
    for i,labels in enumerate(lower):node(d,(10+i*380,330,350+i*380,465),*labels)
    for i in range(3):
        arrow(d,(356+i*380,82),(382+i*380,82))
        arrow(d,(382+i*380,396),(356+i*380,396))
    arrow(d,(1320,158),(1320,322))
    text(d,(655,233),'IF 前完成光束接收匹配  通道内完成角度与时延配置',31)
    path=base/'chain.png';im.save(path);return path
def delay():
    im=Image.new('RGB',(1510,415),'white');d=ImageDraw.Draw(im)
    node(d,(10,137,310,277),'FEL 短脉冲','同一输入光束')
    node(d,(420,137,690,277),'分配与时延','多条照明通道')
    node(d,(825,15,1120,150),'通道 A','L1  角度 A  权重 w1')
    node(d,(825,260,1120,395),'通道 B','L2  角度 B  权重 w2')
    node(d,(1230,137,1505,277),'掩模狭缝','同一位置叠加')
    arrow(d,(318,205),(411,205));arrow(d,(698,177),(817,87));arrow(d,(698,240),(817,327))
    arrow(d,(1128,87),(1222,177));arrow(d,(1128,327),(1222,240))
    text(d,(970,205),'Δt = |L2 - L1| / c',30)
    path=base/'delay.png';im.save(path);return path
def feedback():
    im=Image.new('RGB',(1510,410),'white');d=ImageDraw.Draw(im)
    for i,(a,b) in enumerate([('FEL 与源端控制','功率与谱形'),('输运与支路调节','透过率与配准'),('扫描曝光','照明配方与剂量')]):
        node(d,(40+i*510,20,460+i*510,155),a,b)
    arrow(d,(468,85),(542,85));arrow(d,(978,85),(1052,85))
    node(d,(550,265,970,400),'传感与误差计算','源端及机内监测')
    arrow(d,(1270,163),(1270,330));arrow(d,(1270,330),(978,330))
    arrow(d,(542,330),(240,330));arrow(d,(240,330),(240,163))
    arrow(d,(760,258),(760,163))
    path=base/'feedback.png';im.save(path);return path
p=doc.add_paragraph('FEL光源与0.33 NA EUV光刻机\n照明匹配及去相干调研报告',style='Title')
p.paragraph_format.first_line_indent=Pt(0)
p=paragraph('2026年11月6日',indent=False);p.alignment=WD_ALIGN_PARAGRAPH.CENTER
paragraph('本报告以13.5 nm自由电子激光（FEL）与0.33 NA极紫外（EUV）扫描光刻系统为对象，依次讨论输出光束特性、照明需求、匹配问题、公开技术与专利、典型接入光路和系统参数。研究采用公开论文、设备厂商资料和专利原文，面向技术路线学习、管理层汇报及研发评审。核心问题是把源端辐射转换为满足目标瞳面、掩模狭缝和晶圆曝光窗口的可用光。')

section(1,'FEL输出光束的特性')
sub('1.1','光产生机制与源端参数')
paragraph('电子束通过波荡器时产生辐射，辐射场与电子束相互作用形成微聚束并放大光场。电子能量、波荡器周期和磁场参数共同决定共振波长，束流品质与运行方式进一步决定增益、光谱及脉冲输出。光学接口研究以波荡器出口为源端参考面，保留波长、能量、时间结构和空间分布的联合描述。[3]')
paragraph('工程上需要同步记录平均功率、单脉冲能量、重频、脉宽、谱形、光斑、发散角及其波动。平均功率由脉冲能量和有效重频共同决定，宏脉冲运行还需要加入占空比。透过率与光刻剂量计算使用实际工作模式下的长期平均值，镜面损伤评估同时使用单脉冲面能量密度。')
caption('表1  KEK公开设计中与光学接口有关的源端参数')
table(('参数','公开数值','位置与证据属性'),[
('波长','13.5 nm','FEL出口设计与起止端仿真。'),
('电子能量及平均流强','800 MeV，9.75 mA','该模拟工况的束流参数。'),
('束团电荷及重复频率','60 pC，162.5 MHz','电荷乘重频约为9.75 mA。'),
('光脉冲能量及平均功率','109.4 μJ，17.8 kW','波荡器出口模拟结果，采用4%线性渐变。'),
],(2150,2700,3450))
paragraph('表1给出一组有明确参考面的公开算例，出处为Nakamura等IPAC 2023论文。后续光学设计须用同一工况的光谱、脉冲统计、近远场和偏振数据扩展该表。[3]')

sub('1.2','空间分布与相干性')
paragraph('高增益FEL可以形成高亮度、较小发散角的定向光束。对接口而言，实际光斑、远场分布、波前、指向漂移和横向相干程度共同决定可收集功率。不同FEL架构和束流条件下的横模与角分布具有差异，光学系统用实测近场与远场建立输入模型。[3,5]')
paragraph('空间相干性描述不同横向位置的电场相关性，时间相干性描述同一光场在不同延迟下的相关性。自放大自发辐射（SASE）通过噪声起振，其单脉冲可包含多个纵向相干单元并产生能量和谱形涨落。窄谱种子或再生放大方案的时间相干特性又有不同。接口文件应分别给出空间互相干与时间相干函数，并注明采样时窗和光谱条件。[5,6,8]')
sub('1.3','光谱与偏振的可调性')
paragraph('FEL中心波长可围绕EUV多层膜反射带调整。谱宽和波长漂移会改变各镜面的反射率，并通过入射角与偏振依赖改变瞳面权重。系统可用功率宜按源谱、各级反射率、孔径和角分布联合积分获得。更窄的谱宽有利于集中带内能量，同时会改变相干时间和时延模块的设计范围。[3–5]')
paragraph('波荡器配置可提供不同偏振态。0.33 NA方案可以把偏振作为源端到掩模的传输参数，结合目标图形和多层膜反射特性评估其影响。对需要分区域偏振的照明方式，可同步控制扫描通道与源端偏振，相关公开结构见第4部分。[3,15]')
sub('1.4','参数之间的联动')
paragraph('增加源端平均功率会提高输运镜和整形器件的热负荷。扩展角域会改变接收比例，分束和去相干也会消耗带内功率。改变电子束条件来稳定光功率时，还需观察中心波长、发散角与偏振是否一起变化。因此，源端规范采用成套运行工况，联调期间记录光谱、功率、空间和时间参数之间的相关性。[3,16]')

section(2,'0.33 NA EUV光刻机的照明需求')
sub('2.1','从中间焦点到晶圆的功能分工')
paragraph('ASML NXE:3600D公开资料给出13.5 nm工作波长、晶圆侧0.33投影数值孔径和26 mm × 33 mm最大曝光场。0.33 NA决定投影物镜接收掩模衍射光的范围，中间焦点（IF）的光斑和接收角另由扫描机接口规定。最大曝光场表示扫描完成后的场尺寸，掩模与晶圆在瞬时曝光中经过照明狭缝。[1]')
picture(chain(),'图1  从FEL出口到晶圆的功能参考面，根据公开照明结构绘制。[1,2,11,12]')
paragraph('以波荡器出口R0、中间焦点R1、照明瞳面R2、掩模狭缝R3和晶圆像面R4建立参数传递。R1用于约定源与扫描机之间的接收边界，R2用于表示照明角度与权重，R3用于表征视场照度与通道互相干，R4用于评价扫描剂量和成像结果。实际计量可在对应共轭面完成。[2,11,12]')

sub('2.2','瞳面与狭缝的双重要求')
paragraph('双分面反射照明器由场分面镜、瞳分面镜及必要的中继光学组成。各场分面经通道叠加成像到掩模照明场，各瞳分面指定其照向掩模的方向。配对和倾转控制决定哪些角度参与曝光、各通道分得多少光以及狭缝内的叠加效果。ZEISS和ASML公开的0.33 NA资料展示了常规、离轴及自由形状瞳面设置。[2]')
paragraph('圆形照明的部分相干因子σ表示同一共轭侧照明NA与投影NA的比值，环形照明还记录内外边界，自由形状瞳面使用二维强度与偏振分布。设计需要保持狭缝不同位置看到的角分布符合配方，并通过扫描积分得到稳定剂量。图形和掩模不同，对应的照明配方与容差也会变化。[2]')

sub('2.3','功率光谱与脉冲的接收口径')
paragraph('IF指标同时约定带内平均功率、光谱窗口、有效孔径、角分布和运行时窗。Cymer的公开光源论文使用过13.5 nm附近2%带宽的功率口径，对应约0.27 nm全宽。该数值用于说明带内计量方法，特定扫描机的接收窗口以实际接口为准。[4]')
paragraph('短脉冲FEL与扫描曝光之间通过积分时窗衔接。每个晶圆位置的累计能量取决于该位置驻留狭缝期间收到的脉冲数、单脉冲能量、通道权重和扫描速度。稳定性记录覆盖脉冲间随机涨落、串间变化、慢漂移及漏脉冲。单脉冲强度则与镜面和掩模等位置的面能量密度共同核查。[5,16]')

sub('2.4','照明需求的参数化表示')
caption('表2  光刻照明需求与对应计量')
table(('位置','需求参数','计量或验证方法'),[
('IF入口','带内功率，谱形，光斑位置与尺寸，主光线及角分布，偏振。','校准功率与光谱测量，近远场采样，按机型接收边界积分。'),
('照明瞳面','二维瞳形，各角域权重，瞳面填充及其视场依赖。','测量瞳面或其共轭面，与目标配方逐区对照。'),
('掩模狭缝','视场均匀性，通道互相干，照明方向及偏振。','狭缝照度图、通道干涉或相关测量，结合扫描轨迹。'),
('晶圆像面与曝光后','积分剂量、散斑对比度、空中像对比度、CD均匀性、LER/LWR。','同掩模同剂量空中像对照，再做代表性光刻胶试片。'),
],(1700,3300,3300))
paragraph('表2中，CD表示临界尺寸，LER和LWR分别表示线边及线宽粗糙度，NILS表示归一化像对数斜率，用于评价图形边缘的空中像变化。后续由设备方提供机型接收公差，工艺方提供目标图形、瞳面配方和成像阈值，源端团队据此反推工况及光学转换方案。')
section(3,'FEL与光刻照明之间的匹配问题')
sub('3.1','窄束与照明角域的转换')
paragraph('FEL小发散光束便于定向输运，而扫描机需要在狭缝内叠加规定角度和权重的照明。若窄束只覆盖少量场分面，瞳面可能出现缺光或通道权重偏差。光学匹配因此包括焦点位置、近场、远场和首镜照射足迹的联合转换，目标是把足够的带内功率送入所选照明通道。[2,10,11]')
paragraph('扩束、扩角和光束分配受光学接收范围、辐射亮度及损耗约束。确定的相干光场经过静态光学传播，仍需用完整波动模型追踪相位。工程上增加输出角域以后，还要以实际带宽、通道时间差和曝光积分评价有效相干性及散斑平均。光强平坦度、角域覆盖和相干性使用独立指标记录。[9,10,12]')

sub('3.2','相干光经掩模形成散斑的机制')
paragraph('两条同偏振标量光场照到同一位置时，时间平均强度可写成各自强度之和与一个交叉项。交叉项幅度由归一化互相干函数决定，说明了通道延迟和相位相关性为什么会影响照度。这里的平均时窗及光谱条件必须与检测或曝光条件一致。')
equation('I = I1 + I2 + 2√(I1I2) Re[γ12(Δt)]')
paragraph('反射掩模的表面及多层膜粗糙度会引入局部相位变化，散射光与成像光共同产生像面强度纹理。若该纹理在曝光期间保持相关，其空间分布可影响图形阈值位置和边缘。Pret等采用EUV放电等离子体光源研究了掩模散斑，提供粗糙度到像面强度起伏的实验依据。Wei等进一步用FEL相干模型计算了散斑、照度及LER/LWR的变化，证据属性为成像仿真。[7,8]')
paragraph('相干性评价还需保留图形、焦点、瞳形和掩模状态。同一光源在不同照明配方下产生的图样相关性与成像对比度会变化。报告采用“源端相干描述—通道相干—像面散斑—图形指标”的验证顺序，把光源指标与工艺结果连接起来。[2,7,8]')

sub('3.3','独立图样与曝光平均')
paragraph('散斑对比度C定义为所选图像区域强度标准差与均值之比。对于等强度、互不相关且单幅对比度相同的N幅图样，积分后的对比度近似按C1/√N下降。该理想模型可用于从目标散斑水平估计需要的有效图样数，例如从5%降至1%对应25幅独立图样。[5,6]')
equation('C = std(I) / mean(I)       C ≈ C1 / √N')
paragraph('当通道权重不同或图样相关时，可以用归一化能量权重wi和图样相关系数ρij修正。对单幅对比度相同的情形，下面的关系由归一化强度起伏的方差展开得到。它表明，通道数增加的收益取决于权重均衡程度及实测相关性。')
equation('C²/C1² = Σwi² + 2Σ(i<j) ρij wi wj       Σwi = 1')
paragraph('光脉冲能量平均与散斑平均各有自己的独立样本数。高重频可增加曝光窗口内的能量采样，而角度、光程或动态相位改变决定散斑图样的独立程度。测试需同时保存能量时间序列和对应图样，分别估计积分剂量方差及有效散斑样本数。[5,6]')

sub('3.4','功率热负荷与稳定性的联动')
paragraph('EUV反射光路的总效率由各级反射、孔径和整形效率共同决定。去相干模块增加反射次数或散射角域后，接收功率和器件吸收都会变化。每面镜的热设计使用实际吸收谱、照射足迹、单脉冲面能量密度及冷却条件，并追踪热形变造成的IF漂移、瞳面偏移和波前误差。[4,9,10]')
paragraph('源端功率控制、支路透过率和镜面热稳态也会彼此影响。剂量下降可能来自源功率、谱移、角域偏移或反射率变化。多位置诊断可把这些变化分开定位，控制回路则按扰动时标和执行器带宽分配任务。[16]')

section(4,'公开的匹配与去相干技术')
paragraph('本轮以公开可访问的中、美、欧专利文本、论文和技术会议材料为检索范围。组合词包括“free electron laser”与“EUV lithography”“illumination”“intermediate focus”，以及“speckle”“optical delay”“dose control”和“极紫外 自由电子激光 去相干 匀光”等。以下按光学功能选取七件代表性专利，归纳其结构和工作过程。专利文本支持技术披露与权利要求分析，装置性能另按实验、仿真或公开设计算例标注。')

sub('4.1','光束输运与角域整形')
detail('结构化掠入射镜和空心波导')
paragraph('ZEISS的US10928734B2针对FEL输出束的引导与调理。其权利要求1采用两面结构化掠入射镜及下游空心波导。第一面相对反射表面的掠入射角为1–10 mrad，两镜入射面夹角大于45°，第二面的掠入射角至少为第一面的两倍，并分别约束最大散射角。权利要求2、3涉及EUV与FEL输出。[10]')
paragraph('工作时，第一面镜先沿一个入射平面改变角分布，第二面镜再沿不同平面扩展角域，空心波导接续约束和调理光束。通过分级展开光束，可为下游器件分配合适的光束足迹。评估该构型时，以带内接收功率、出射相空间、局部热负荷和污染敏感性作为设计比较量。[10]')
detail('支路聚焦与中间焦点重建')
paragraph('ZEISS的US9958783B2将公共束整形、支路耦出和IF重建组织成分层光路。权利要求1给出椭球面镜与抛物面镜聚焦组件，并约束椭球长半轴a与抛物面焦距f之比大于50。权利要求13涉及公共束和多个支路，权利要求15涉及聚焦组件下游的场分面镜和瞳分面镜。[11]')
paragraph('这种结构可用于比较单机输运与多机分配。公共束线先完成稳定传输，各支路再独立调节光斑、角域和主光线以匹配IF。具体工程输入来自接收图和安装空间，光学处方与各段效率据此设计。[11]')

sub('4.2','分束与差分时延')
paragraph('同一FEL脉冲分入多个照明通道以后，相对光程差产生到达时差。时差超过相关函数的有效宽度时，交叉干涉减弱。若同时选择不同照明角度，各通道的散斑图样还可能发生空间去相关。设计分别测量通道互相干及图样相关性，确定这两项收益。[5,6,12]')
picture(delay(),'图2  分束时延与角度多样性的作用位置，根据公开原理绘制。[6,12]')
paragraph('ZEISS的EP2823360B1以双分面照明通道为载体。权利要求2要求在照明场各点，任意两条子束的传播时差大于相干时长。权利要求3增加子束延时分量，权利要求4利用瞳分面镜布置形成不同通道长度。工程实现可通过分面几何取得光程差，也可加入延迟元件。[12]')
equation('Δt = |L2 − L1| / c')
paragraph('xLight的2024年公开材料给出约3 fs相干时间、约1 μm相干长度，以及超过10 μm的分面通道光程差。按真空光程换算，3 fs对应0.90 μm，10 μm对应33.4 fs。这些数值是其算例量级。实际时延根据工作谱和测得的相关函数选择，要求子脉冲包络分开的设计还需检查脉宽与尾部。[5]')
paragraph('Stupakov在FEL 2024报告中用简化模型研究了短脉冲分束、时延和空间叠加对散斑抑制的影响，为该机制提供模型分析。全系统设计进一步检查视场各点的最小光程差、通道强度均衡、热漂移和反射损耗。[6]')

sub('4.3','动态角度和相位调制')
detail('两级运动粗糙反射面')
paragraph('ASML的US11984236B2在权利要求1中提出两级粗糙反射面和驱动机构，并规定第二级表面以不同传播位置的凸起或凹陷实现扩散。权利要求2、3涉及旋转及约1 m/s或更高的表面运动速度。光束依次采样变化的微结构，出射角分布和相位图样随时间更新，说明书还讨论了分光前后布置。[13]')
paragraph('该模块可同时承担角域调理和曝光平均。设计变量包括粗糙面统计、入射角、表面速度、采样光斑尺寸和两级相对运动。验证使用实际曝光时窗内的图样相关衰减、可收集功率、指向抖动和镜面热负荷。专利中速度为技术披露值，目标装置速度由这些量共同确定。[13]')
detail('旋转微结构反射镜')
paragraph('华为CN115343915A的权利要求1涉及相邻微反射面的高度差，绝对值位于(0，kλ]，k为正整数。权利要求11–13涉及可旋转反射镜、照明系统分束至掩模、FEL应用和控制方法。高度差在反射后形成相对相位变化，转镜让光斑在曝光期间采样不同微面，从而更新照明干涉图样。[14]')
paragraph('这一路线把静态微结构与时间调制结合。参数选择需要配合单个晶圆位置的曝光时窗，测量转速、光斑相对转轴的位置、相位分布、振动及光收集比例。该文件为申请公开文本，所示机制可用于研究比较。实际散斑和图形收益通过同剂量试验确定。[14]')
detail('扫描照明与偏振同步')
paragraph('ZEISS的EP2883112B1将FEL偏振设定与扫描光路相连。权利要求1涉及影响电子束偏转的偏振设定装置，权利要求4、5涉及扫描与偏振同步以及双分面通道。工作时，窄束依次投向不同分面，每个被选角域获得预定偏振和能量。[15]')
paragraph('该专利提供照明角度和偏振分配的结构例子。与时延或动态相位结合后，可同时比较瞳面配方和时间平均。评价量包括各角域的停留时间、能量权重、偏振误差与图样相关性，并核对扫描装置和波荡器控制的可实现带宽。[15]')
sub('4.4','反射式匀光与多通道叠加')
paragraph('反射式匀光管通过多次掠反射改变传播方向和光程，使不同路径在出口叠加。倾斜管段可加强两个横向方向的混合。李慧等对13.5 nm、直径200 μm、发散度20 mrad的高斯输入开展仿真，在内径1 mm、总长至少0.6 m、倾斜角10 mrad的一组选定结构下给出37.6%能量透过率。[9]')
paragraph('该结果支持把匀光管作为强度与角域调理的候选。工程比较使用实际输入光谱、脉冲与相干模型，分别记录出口均匀性、通道互相干和晶圆像面散斑。对静态器件，完整相位传播与检测时空平均决定所观察到的相干指标，验证中同时保留相位或互相干测量。[9]')
paragraph('匀光管与双分面照明的功能可以串联。前者提供预定出射角域，后者按瞳面和狭缝要求分配光束。匹配时检查多次反射后的光程分布、角域接收比例、真空与污染条件，并把各段损耗分开记入预算。')

sub('4.5','扫描积分与剂量反馈')
paragraph('在10 MHz重复频率和1 ms有效驻留时间的演示条件下，一个晶圆位置接收约10000个脉冲。若单脉冲能量的相对标准差为20%，且各脉冲独立、等权，则累计能量相对标准差约为0.2%。这些输入是统计演示假设，用于说明重频与积分的关系；实际计算采用测得的时间相关性、扫描权重和控制状态。')
equation('Npulse ≈ frep × Texp       s(sum)/mean(sum) ≈ spulse/√Npulse')
paragraph('ASML的US9823572B2在权利要求1中提出FEL反馈控制与输出光可变衰减共同调节送入光刻机的功率，权利要求3涉及各光刻机独立衰减，权利要求4–7涉及第二反馈回路及机内传感位置。源端回路稳定公共输入，支路衰减与机内测量补偿各机实际收到的能量，形成两层控制结构。[16]')
picture(feedback(),'图3  源端与支路反馈的功能示意，根据公开剂量控制思路绘制。[16]')
paragraph('多机运行时，分配比例、各支路透过率和源端工况一起决定剂量。独立支路控制需要保留源端共享扰动的诊断量，同时检查某一支路关断或调节时其他支路的响应。源端调节若改变光谱，还需把波长依赖的反射率加入反馈模型。[16]')

sub('4.6','技术作用与证据的比较')
caption('表3  公开技术的作用与验证重点')
table(('技术','公开依据及作用','设计比较量'),[
('结构化双镜与波导','US10928734B2，改变输入束的角域并引导至照明系统。','相空间、带内接收率、峰值负载。'),
('支路IF重建','US9958783B2，公共束分配与单机聚焦。','焦点、角锥、支路透过率及主光线。'),
('分束时延','EP2823360B1及FEL 2024模型，降低通道交叉干涉。','最小通道时差、互相干、独立图样数。'),
('动态反射与扫描','US11984236B2、CN115343915A、EP2883112B1，更新相位或角度并按配方分配。','时间相关、平均瞳面、抖动与效率。'),
('反射式匀光管','物理学报2024仿真，出口强度与角域调理。','透过率、光程分布、像面散斑和热负荷。'),
('双层剂量反馈','US9823572B2，源端反馈加支路衰减。','积分剂量误差、带宽、串扰和谱形联动。'),
],(1900,3500,2900))
paragraph('表3中的专利披露提供结构与控制方法，匀光管和短脉冲时延资料提供模型或仿真结果。具体接入方案用同一源谱、瞳面、狭缝、剂量和图形进行对照，才能量化各技术带来的有效功率与成像收益。')

section(5,'典型FEL接入方案及技术组合')
sub('5.1','经IF适配既有照明器')
paragraph('该路线以扫描机IF为设备交接面。FEL出口先经过诊断采样和真空输运，再由反射镜组完成扩束、角域调理及末级聚焦，送入既有双分面照明器。光束转换可借鉴US10928734B2的分级调理和US9958783B2的支路聚焦。对既有照明器，先测各通道覆盖、光程及瞳面权重，再确定外置时延或动态模块的位置。[10–12]')
paragraph('IF前模块的损耗体现在交接功率，IF后模块的损耗计入照明器。若内部通道本身提供足够时差，联调重点转向角度分布和图样相关性。若需要调整内部通道，设备方需提供相关光学和控制接口。验证按“IF接收—瞳面与狭缝—像面与试片”逐级进行。')
sub('5.2','与照明器协同设计')
paragraph('该路线从目标瞳面和掩模狭缝反推通道。窄束通过分束或扫描进入选定场分面，每条通道设置能量权重、到达时刻和偏振，再由瞳分面和中继镜形成照明。时延设计参照EP2823360B1，扫描与偏振配方参照EP2883112B1，支路剂量回路参照US9823572B2。[12,15,16]')
paragraph('协同设计可直接利用FEL的方向性，比较外置整形与通道内分配的功率代价。其实施涉及照明器处方、分面执行器、13.5 nm计量和曝光控制。台架先用固定瞳形测量通道模型，再扩展到自由形状配方和扫描同步，最终以晶圆图形确定参数范围。[2]')
caption('表4  两类接入方案的比较')
table(('比较项','IF适配路线','协同设计路线'),[
('光学边界','以机型IF接收图为主要约束。','从瞳面、狭缝和通道功能反推源端输入。'),
('关键工作','转换光斑、角域和主光线，核对既有通道覆盖。','设计分束或扫描、时延、偏振和通道配方。'),
('主要评价','IF有效功率与瞳面适配程度。','每瓦源端功率的有效曝光收益及控制复杂度。'),
('实施条件','取得IF、脉冲与照明器接收数据。','取得内部光学和控制的协同设计条件。'),
],(1600,3350,3350))

sub('5.3','带内功率预算示例')
paragraph('单机IF功率等于源端带内功率乘公共输运效率、支路分配份额和该支路的转换效率。各效率因子按器件参考面定义。多机分配份额只计一次，IF后照明与投影损耗继续在下游预算中追踪。')
equation('PIF,i = P0,band × Tcommon × fi × Tbranch,i × Tmatch,i')
caption('表5  单机IF功率预算演示')
table(('参考段','假设系数','段后带内功率','口径'),[
('源端R0','—','6.00 kW','演示输入，统一工作谱窗。'),
('公共输运','0.80','4.80 kW','反射和孔径损失。'),
('所选支路分配','1.00','4.80 kW','单机演示，分配份额为1。'),
('支路镜组与整形','0.60','2.88 kW','该段包含的器件损耗。'),
('IF前转换及接收','0.35','1.008 kW','含所选去相干、孔径和对准代价。'),
],(2300,1300,1800,2900))
paragraph('表5各级系数是预算演示假设，总效率为16.8%。xLight在2024年光源研讨会摘要中给出6 kW源端至1 kW IF的设计量级，对应约16.7%，可用于对照总量级。该公开材料的逐级器件效率还需另取输入。晶圆剂量预算继续加入实际照明器、掩模和投影链路的波长与角度依赖透过率。[17]')

sub('5.4','去相干配置的联合对照')
paragraph('可以用相同的源端模式、掩模和照明瞳面，比较“固定光路”“增加通道时延”“时延加动态平均”三种台架状态。每种状态记录IF功率、瞳面误差、狭缝照度、散斑对比度、积分剂量和图形成像。配置选择以满足目标工艺指标时的有效功率、稳定性和维护需求为依据。')
paragraph('若测试显示时延降低了通道干涉，而像面散斑仍具有较高相关性，下一步调整角度采样或动态相位参数。若散斑指标满足要求而有效功率偏低，则优化反射次数、收集角域和通道分配。这个对照把机制验证与功率预算放在同一组工作条件中。')

section(6,'重要参数与系统匹配总结')
sub('6.1','各参考面的参数关系')
paragraph('FEL用于0.33 NA EUV光刻时，光源、束线、IF接口、照明器、投影系统及曝光控制共同决定最终结果。参数定义需写明波长、位置、统计时窗、运行模式和测量方法。表6将这些匹配关系汇总成后续接口文件的字段。')
caption('表6  重要参数与系统匹配总表')
table(('参数组','参考面及匹配对象','关键字段','验证方法'),[
('光谱与带内功率','FEL、各级镀层及IF接收','中心波长、谱宽、带内功率、带外份额与漂移。','校准谱仪及功率计，按工作谱窗积分。'),
('光斑与角域','R0输运至R1，照明器首镜','焦点、光斑、主光线、角分布、波前与指向。','近远场及共轭面测量，接收图积分。'),
('脉冲结构','源端、镜面与曝光积分','脉宽、重频、单脉冲能量、面能量密度、丢脉冲。','时间诊断与能量序列，核对实际驻留时窗。'),
('空间和时间相干','通道与掩模R3','互相干函数、相关宽度、两两时差及漂移。','通道干涉可见度、延迟扫描和相干模型。'),
('瞳面与偏振','R2和掩模照明角度','二维瞳形、各角域权重、σ与偏振分布。','瞳面或共轭面成像，按配方比对。'),
('狭缝与扫描剂量','R3至R4曝光控制','照度均匀性、积分剂量、扫描速度和低频漂移。','带时戳的照度与能量测量，按轨迹积分。'),
('散斑与图形','晶圆像面和曝光后','C、图样相关、NILS、CD均匀性、LER/LWR。','同瞳形同剂量空中像与光刻胶试片。'),
('热污染与运行','镜组、真空与维护','吸收功率、温度、热漂移、反射率衰减和故障恢复。','热态长时运行、功率闭合及维护前后复测。'),
],(1450,2100,2450,2300))
paragraph('具体阈值由目标扫描机和工艺输入确定。公开的0.33 NA平台规格支持波长、投影孔径和总体照明架构的选择，IF位置与角域公差、脉冲限制、目标瞳面、剂量窗口和图形验收值仍需要设备方与工艺方确认。[1,2]')

sub('6.2','验证顺序与结论')
paragraph('第一阶段建立源端统计和器件热态效率，形成可重复的工作工况。第二阶段在IF接收边界内测功率、光斑和角域，并核对照明器首镜覆盖。第三阶段按目标瞳面测通道时差、互相干、狭缝照度与曝光窗口内的图样相关性。第四阶段以同剂量空中像和代表性试片验证CD与边缘指标，再扩展运行时间、配方及多机联调。')
paragraph('现有公开资料给出了从FEL窄束输运到多通道照明的完整功能链，以及分束时延、动态相位、角度扫描、匀光和剂量反馈的具体技术结构。对本项目，单机IF适配加照明通道差分时延可作为首轮比较基线，动态平均和照明器协同设计作为条件化扩展。最终配置由实际接收参数、散斑与图形结果及有效功率共同确定。')

p=doc.add_paragraph('参考资料与专利原文',style='Heading 2')
p.paragraph_format.page_break_before=True
ref(1,'ASML. TWINSCAN NXE:3600D产品资料，13.5 nm、0.33 NA与曝光场。','https://www.asml.com/en/products/euv-lithography-systems/twinscan-nxe-3600d')
ref(2,'Zimmermann等. Flexible illumination for ultra-fine resolution with 0.33 NA EUV lithography，EUVL 2016。','https://euvlsymposium.lbl.gov/pdf/2016/Oral/Mon_S3-3.pdf')
ref(3,'Nakamura等. EUV-FEL light source for future lithography，IPAC 2023，THPM123。','https://doi.org/10.18429/JACoW-IPAC2023-THPM123')
ref(4,'Cymer. SPIE Advanced Lithography 2012光源论文，IF带内功率口径。','https://www.cymer.com/wp-content/uploads/2018/12/Cymer_SPIE_AdvancedLithography_2012.pdf')
ref(5,'Nguyen代表xLight团队. High-average-power ERL FEL for EUV Lithography，2024 EUVL Workshop。','https://euvlitho.com/2024/P43.pdf')
ref(6,'Stupakov. Speckle suppression by splitting and delaying short FEL pulses，FEL 2024。','https://indico.jacow.org/event/72/contributions/5820/')
ref(7,'Pret等. Evidence of speckle in extreme-UV lithography，Optics Express 20，25970，2012。','https://doi.org/10.1364/OE.20.025970')
ref(8,'Wei等. Impact of free electron laser coherence on imaging quality，Chinese Physics B 34，054201，2025。','https://doi.org/10.1088/1674-1056/adb734')
ref(9,'李慧等. 基于匀光管的极紫外消相干和光强均匀化仿真研究，物理学报73，114201，2024。','https://doi.org/10.7498/aps.73.20240335')
ref(10,'ZEISS. US10928734B2，Optical assembly for guiding an output beam of a free electron laser。','https://patents.google.com/patent/US10928734B2/en')
ref(11,'ZEISS. US9958783B2，Illumination system for EUV projection lithography。','https://patents.google.com/patent/US9958783B2/en')
ref(12,'ZEISS. EP2823360B1，Illumination optics for EUV projection lithography and optical system having such an illumination optics。','https://patents.google.com/patent/EP2823360B1/en')
ref(13,'ASML. US11984236B2，Radiation system。','https://patents.google.com/patent/US11984236B2/en')
ref(14,'华为. CN115343915A，反射镜、光刻装置及其控制方法，申请公开文本。','https://patents.google.com/patent/CN115343915A/zh')
ref(15,'ZEISS. EP2883112B1，Illumination system for an EUV projection lithographic projection exposure apparatus and method for operating the same。','https://patents.google.com/patent/EP2883112B1/en')
ref(16,'ASML. US9823572B2，Lithographic method。','https://patents.google.com/patent/US9823572B2/en')
ref(17,'2024 Source Workshop摘要集，xLight光束输运与扫描机集成设计算例，PDF第21页。','https://www.euvlitho.com/2024/2024%20Source%20Workshop%20Abstract%20Book.pdf')
p=paragraph('公开网页与专利文本访问日期为2026年11月6日。原文链接用于技术内容追溯，地域法律状态与自由实施需另行核查。',indent=False)
for r in p.runs:r.font.size=Pt(10.5)
p.paragraph_format.line_spacing=1.2
# Source custom XML and notes are kept as package structures, while user-facing
# body content and figures are wholly authored for the approved six-part scope.
out=root/'outputs/2026-11-06_FEL光源与0.33NA_EUV光刻机照明匹配及去相干调研报告.docx'
doc.save(out)
print(out)
print('paragraphs',len(doc.paragraphs),'tables',len(doc.tables),'figures',len(doc.inline_shapes))
print('body characters',sum(len(p.text) for p in doc.paragraphs))
