"""Build a feasibility-style 13.5 nm FEL report with reliably visible equations.

Equation PNGs and diagrams are reproducible temporary assets. The formula text
is also recorded in captions/descriptions and the source for review.
"""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'outputs' / '2026-11-04_13.5nm_EUV_FEL原理技术与系统参数研究报告.docx'
TMP = ROOT / 'tmp' / 'euv_fel_report_assets'
TMP.mkdir(parents=True, exist_ok=True)
FONT_MATH = r'C:\Windows\Fonts\cambria.ttc'
FONT_CN = r'C:\Windows\Fonts\msyh.ttc'
EQ_COUNT = 0

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Inches(8.5), Inches(11)
sec.top_margin = Cm(2.0)
sec.bottom_margin = Cm(1.8)
sec.left_margin = Cm(2.1)
sec.right_margin = Cm(2.1)

styles = doc.styles
normal = styles['Normal']
normal.font.name = 'Microsoft YaHei'
normal.font.size = Pt(10.5)
normal.font.color.rgb = RGBColor(35, 42, 48)
normal._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
normal.paragraph_format.space_after = Pt(6)
normal.paragraph_format.line_spacing = 1.22
for name, size, before, after in [('Title', 20, 0, 13), ('Heading 1', 14, 17, 9), ('Heading 2', 11.5, 12, 5)]:
    s = styles[name]
    s.font.name = 'Microsoft YaHei'
    s.font.size = Pt(size)
    s.font.bold = name != 'Title'
    s.font.color.rgb = RGBColor(20, 40, 62) if name != 'Title' else RGBColor(0, 0, 0)
    s._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
    s.paragraph_format.space_before = Pt(before)
    s.paragraph_format.space_after = Pt(after)
    s.paragraph_format.keep_with_next = True

if 'Equation' not in styles:
    st = styles.add_style('Equation', WD_STYLE_TYPE.PARAGRAPH)
else:
    st = styles['Equation']
st.font.name = 'Cambria Math'
st.font.size = Pt(11)
st.paragraph_format.space_before = Pt(3)
st.paragraph_format.space_after = Pt(7)

def p(text='', boldlead=None, style=None):
    x = doc.add_paragraph(style=style)
    if boldlead:
        x.add_run(boldlead).bold = True
    x.add_run(text)
    return x

def h(text, level=1):
    doc.add_heading(text, level)

def math(line):
    """Render one or two compact formula lines at 300 dpi, never crop glyphs."""
    global EQ_COUNT
    EQ_COUNT += 1
    lines = [part.strip() for part in line.split('；')]
    font = ImageFont.truetype(FONT_MATH, 59)
    scratch = Image.new('RGB', (10, 10), 'white')
    draw = ImageDraw.Draw(scratch)
    widths = [draw.textbbox((0, 0), x, font=font)[2] for x in lines]
    max_w = max(widths)
    # 6.05 in at 300 dpi. The physical font scales if an expression is longer.
    scale = min(1, 1770 / max_w)
    width = int(max_w * scale + 55)
    height = int((len(lines) * 91 + 12) * scale + 18)
    big = Image.new('RGB', (max_w + 70, len(lines) * 91 + 28), 'white')
    bd = ImageDraw.Draw(big)
    for i, part in enumerate(lines):
        bd.text(((big.width - widths[i]) / 2, i * 91 + 8), part, fill=(24, 38, 54), font=font)
    if scale < 1:
        big = big.resize((width, height), Image.Resampling.LANCZOS)
    path = TMP / f'equation_{EQ_COUNT:02}.png'
    big.save(path, dpi=(300, 300))
    x = doc.add_paragraph(style='Equation')
    x.paragraph_format.keep_with_next = True
    x.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = x.add_run()
    shape = run.add_picture(str(path), width=Inches(min(big.width / 300, 6.05)))
    shape._inline.docPr.set('descr', f'公式 {EQ_COUNT}：{line}')
    caption = doc.add_paragraph(f'式（{EQ_COUNT}）', style='Caption')
    caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption.paragraph_format.space_after = Pt(7)
    return x

def diagram(which):
    W, H = (1800, 580) if which == 'architecture' else (1800, 500)
    im = Image.new('RGB', (W, H), 'white')
    d = ImageDraw.Draw(im)
    cn = ImageFont.truetype(FONT_CN, 31)
    small = ImageFont.truetype(FONT_CN, 26)
    dark, blue, gray = (24, 52, 75), (35, 103, 143), (88, 104, 115)
    def box(x, y, w, h, title, sub, fill=(235, 244, 248)):
        d.rounded_rectangle((x,y,x+w,y+h), radius=16, outline=blue, width=3, fill=fill)
        d.text((x+18,y+19), title, font=cn, fill=dark)
        d.text((x+18,y+68), sub, font=small, fill=gray)
    def arrow(x1,y1,x2,y2):
        d.line((x1,y1,x2,y2), fill=blue, width=5)
        d.polygon([(x2,y2),(x2-18,y2-10),(x2-18,y2+10)], fill=blue)
    if which == 'architecture':
        d.text((70,25),'中红外验证基线',font=cn,fill=dark)
        box(70,83,490,145,'几十 MeV 电子束','35 MeV 实测例 [1]')
        box(660,83,490,145,'中红外波荡器与光腔','15 μm 出光')
        box(1250,83,480,145,'诊断与应用','验证波长和脉冲')
        arrow(560,155,650,155); arrow(1150,155,1240,155)
        d.text((70,287),'13.5 nm 高平均功率候选',font=cn,fill=dark)
        box(70,342,370,155,'高亮度电子枪','60 pC / 162.5 MHz')
        box(505,342,370,155,'SRF 加速与压缩','800 MeV / 9.75 mA')
        box(940,342,370,155,'波荡器与渐变','13.5 nm / 17.8 kW*')
        box(1375,342,355,155,'光束整形到 IF','功率与照明验收')
        for x in [440,875,1310]: arrow(x,420,x+60,420)
        d.text((620,520),'电子束经减速段回收。* 为公开仿真结果。',font=small,fill=gray)
    else:
        d.text((70,25),'同一功率口径下的损失账本',font=cn,fill=dark)
        names=[('波荡器出口','源端功率'),('分光与带通','分光 × 带通'),('传输镜组','镜组透过率'),('整形去相干','有效透过率'),('中间焦点','IF 功率')]
        xs=[70,425,780,1135,1490]
        for i,(a,b) in enumerate(names):
            box(xs[i],125,285,150,a,b)
            if i<4: arrow(xs[i]+285,200,xs[i+1]-12,200)
        d.text((70,335),'各段同时记录谱宽、指向、光瞳、相干、热负荷和时间稳定性。',font=small,fill=gray)
        d.text((70,390),'示例：17.8 kW × 假设链路效率 20% = 3.56 kW IF。20% 为方案比较用的假设值。',font=small,fill=gray)
    path = TMP / f'figure_{which}.png'
    im.save(path, dpi=(300,300))
    para=doc.add_paragraph()
    para.paragraph_format.keep_with_next = True
    para.alignment=WD_ALIGN_PARAGRAPH.CENTER
    shape=para.add_run().add_picture(str(path),width=Inches(6.20))
    shape._inline.docPr.set('descr', '系统架构示意图' if which=='architecture' else 'EUV 功率损失账本示意图')
    cap=doc.add_paragraph('图 1  中红外验证与 EUV ERL SASE 方案的系统差异' if which=='architecture' else '图 2  源端至中间焦点的功率与照明指标链',style='Caption')
    cap.alignment=WD_ALIGN_PARAGRAPH.CENTER
    return path

def table(headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.style = 'Light Shading Accent 1'
    t.autofit = False
    if widths:
        for i,w in enumerate(widths):
            t.columns[i].width = Cm(w)
    for i,z in enumerate(headers):
        t.rows[0].cells[i].text = z
    t.rows[0]._tr.get_or_add_trPr().append(OxmlElement('w:tblHeader'))
    for row in rows:
        cells = t.add_row().cells
        for i,z in enumerate(row):
            cells[i].text = str(z)
    for row in t.rows:
        row._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))
        for col_idx, cell in enumerate(row.cells):
            if widths:
                cell.width = Cm(widths[col_idx])
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            tcMar = cell._tc.get_or_add_tcPr().first_child_found_in('w:tcMar')
            if tcMar is None:
                tcMar = OxmlElement('w:tcMar')
                cell._tc.get_or_add_tcPr().append(tcMar)
            for side in ('top','bottom','left','right'):
                m = OxmlElement('w:'+side)
                m.set(qn('w:w'), '85' if side in ('top','bottom') else '95')
                m.set(qn('w:type'), 'dxa')
                tcMar.append(m)
            for para in cell.paragraphs:
                para.paragraph_format.space_after = Pt(1)
                para.paragraph_format.line_spacing = 1.12
                for run in para.runs:
                    run.font.size = Pt(8.5)
                    run.font.name = 'Microsoft YaHei'
                    run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return t

def bullet(text):
    return p(text, style='List Bullet')

doc.add_paragraph('13.5 nm 极紫外 FEL 光源的产生', style='Title')
p('中红外至极紫外的参数调整与实现方案  |  2026 年 11 月 5 日')
p('13.5 nm 自由电子激光由高能电子束在波荡器中直接产生。基频方案通过提高电子能量获得短波长，三次谐波方案通过选择更高的辐射阶次降低所需能量。两条路线分别以高平均功率和降低加速器能量为设计重点。')

h('1 FEL 光源的产生机理')
p('电子枪产生电子束团，加速器提高电子能量，波荡器的周期磁场使电子横向振荡并辐射。辐射场对电子施加能量调制，电子逐渐在辐射波长尺度形成微聚束。微聚束内的电子同相辐射，使光功率沿波荡器增长。中红外装置常用光腔让辐射多次通过电子束，短波高增益装置可在一次通过中达到饱和。[1,12]')
math('λₙ = λᵤ(1 + K²/2)/(2nγ²)；γ = 1 + Eₖ/(0.511 MeV)')
p('上式给出平面波荡器的轴上共振波长。λᵤ 为波荡器周期，K 为按峰值磁场定义的无量纲参数，Eₖ 为电子动能，n 为谐波阶次。采用均方根参数 aᵥ 时，波长公式的磁场项写为 1+aᵥ²。13.5 nm 对应光子能量约 91.84 eV。[2,6,12]')

h('2 从中红外到极紫外需要调整的参数')
p('15 μm 到 13.5 nm 的波长缩短约 1111 倍。在固定周期、磁场和基频的条件下，电子总能量随波长的平方根倒数增加，所需洛伦兹因子约增至 33.3 倍。以 28 mm 周期和 aᵥ＝1.168 计算，15 μm 对应约 23.5 MeV，13.5 nm 对应约 800 MeV。这组数值表示同一波荡器条件下的计算结果。[2,3]')
math('Eₖ(λ) = 0.511 MeV × {√[λᵤ(1+aᵥ²)/(2λ)] − 1}')
table(['参数','中红外参照','13.5 nm 调整方向'],[
    ('电子动能','FELiChEM 在 35 MeV 首次获得 15 μm','基频公开参考点为 800 MeV，三次谐波参考点为 331 MeV [1–3,6]'),
    ('波荡器周期与磁场','按中红外波长配置周期和磁隙','联合选定周期、K 和谐波阶次。公开例为基频 28 mm、三次谐波 16 mm [2,6]'),
    ('峰值电流与束质','光腔多次通过可累积增益','提高峰值电流，同时控制切片能散、归一化发射度和轨道误差 [3,12]'),
    ('平均电流','可采用宏脉冲运行，也有高平均功率红外实例','千瓦级参考设计采用百 MHz 级束团频率、约 10 mA 平均电流 [3,4,6–8]'),
    ('反馈与输出光学','使用红外光腔或红外光束线','基频 SASE 采用单程增益。RAFEL 使用 EUV 腔镜，出光后接真空反射光束线 [2,6]')],[3.1,5.0,8.2])
p('缩短周期和选用高阶谐波均可降低共振所需能量。短周期方案重点校核磁隙、束流孔径与磁场质量。谐波方案重点校核有效增益和基频抑制。提高峰值电流可缩短增益长度，束团压缩则会提高相干同步辐射和能散控制的难度。[6,12]')

h('3 参数调整所需系统及关键问题')
p('电子枪与注入器提供高品质电子束，压缩段提高峰值电流，波荡器将部分电子能量转换为 EUV。出光后，电子束进入能量回收段，EUV 经传输及照明系统到达光刻机中间焦点。')
table(['系统','承担的任务','关键问题及处理方式'],[
    ('电子枪与注入器','提供高频率、低发射度电子束团','同步优化阴极激光、束团电荷和注入光学，测量阴极寿命、发射度和暗电流。'),
    ('超导直线加速器与压缩段','提升能量并压缩束团，获得足够峰值电流','优化磁压缩与色散，测量切片能散及到达时间抖动。超导腔配置高阶模阻尼并测束流破裂阈值。'),
    ('分段波荡器与诊断','控制共振、增益和饱和后的能量提取','校准磁场、段间相位和电子轨道。沿程渐变 K，使失去能量的电子继续与辐射场保持共振。'),
    ('能量回收与束流收集','让出光后电子束返回主加速器减速，将能量送回射频场','按出光后能散和能量尾设计回收线接受度，测量回收效率、沿程束损和收集端能谱。'),
    ('EUV 传输及照明','将带内功率送至中间焦点，形成目标光瞳','控制反射镜热漂移、污染及传输损失。联合测量 IF 带内功率、光瞳、散斑和剂量稳定性。')],[3.9,4.4,8.0])

h('3.1 电子枪与注入器',2)
p('提高平均电流有两个途径，即提高束团重复频率 f 或提高单束团电荷 q。保持单束团光脉冲能量时，提高重复频率可使平均 EUV 光功率近似同比增加。提高单束团电荷会增强低能段的空间电荷排斥，使束斑膨胀，并因横向受力随位置变化而增加发射度，影响后续聚焦和 FEL 增益。长时间出束还会消耗阴极工作寿命。残余气体被束流电离后，部分正离子可回轰阴极，使量子效率下降并改变发射分布。[2,3]')
math('I(avg) = qf；P(beam)[kW] = Eₖ[MeV] × I(avg)[mA]')
p('先在目标单束团电荷下优化阴极驱动激光的横向光斑和时间分布、枪场与注入聚焦，再逐级提高重复频率。增加光斑尺寸可降低电荷密度，同时会增加由阴极初始横向动量引起的发射度，需结合束流仿真确定工作点。采用高真空、阴极量子效率监测和驱动激光功率反馈保持稳定出束。每个电流台阶记录阴极量子效率随累计抽取电荷的变化、束团电荷波动、归一化发射度及暗电流，并检查送到波荡器的单束团品质。[2,3]')

h('3.2 超导直线加速器与束团压缩',2)
p('压缩前，射频相位在束团前后形成能量随纵向位置变化的相关项，称为能量啁啾。弯转磁铁使不同能量的电子经历不同路程。以 z 为束团内纵向坐标、δ 为相对动量偏差，一阶传输可写为 z(out) ≈ z(in) + R₅₆δ。选择与啁啾相匹配的纵向色散 R₅₆，使束团头尾靠近并提高峰值电流。强压缩时，短束团在弯道中产生的相干同步辐射（CSR）会改变束团各处的能量，弧段色散再将部分能量扰动转化为横向位置和角度偏差。射频相位和幅度波动也会通过压缩段变成峰值电流与到达时间抖动。[2,3,12]')
p('联合扫描射频相位、压缩段纵向色散和压缩分配，用含空间电荷、CSR 与尾场的粒子跟踪比较峰值电流、切片能散和发射度。具备条件时，用谐波腔或高阶束流光学修正纵向非线性。调试时测量束团长度、纵向相空间和到达时间，结合横向发射度和光谱，选择 FEL 增益与束质均稳定的压缩区间。')
p('高平均电流还会增强超导腔高阶模（HOM）的激励。偏轴束团激发的横向偶极模会踢偏后续通过腔体的束团。回流光学将角度扰动转化为再次入腔的位置偏差，在特定相位下形成正反馈，触发多通束流破裂（BBU）。采用 HOM 耦合器或吸收器降低危险模式的加载品质因数，联合优化回流横向光学和束流轨道。根据实测腔模频率和阻尼复算 BBU 阈值，并在逐级提流时监测轨道振荡、HOM 功率及腔体热负荷。[3,9]')

h('3.3 分段波荡器与束光诊断',2)
p('FEL 增益依赖电子束和辐射场在波荡器中持续重叠并保持合适相位。切片内能散过大时，部分电子偏离共振。发射度和轨道偏差增大时，电子与光场的重叠变差。段间磁场或相位误差也会削弱后续放大。在一维高增益近似下，可用 Pierce 参数 ρ 判断增益长度和切片能散的量级，再通过三维时变计算确定实际波荡器长度。[12]')
math('L(gain) ≈ λᵤ/(4π√3 ρ)；σδ ≲ ρ')
p('先校准各段磁场和间隙，利用束流位置监测、校正磁铁及段间相移器调整轨道和相位。达到饱和附近后，电子向光场交出能量，洛伦兹因子 γ 下降。在固定波荡器周期和目标波长时，沿程降低 K 可让仍被俘获的电子继续满足共振条件，提高提取效率。用单脉冲能量、中心波长、谱宽及出光后能量分布联合选择渐变起点和梯度，同时把回收线允许的能量尾部纳入优化。[2,3,12]')

h('3.4 能量回收与束流收集',2)
p('出光后的电子束经回流弧以减速相位进入超导腔，将动能交回射频场，再以较低能量进入收集端。减速时，给定绝对能量偏差所占的相对比例会增大，电子束更容易超出回收线的能量接受范围或在有色散的位置碰到孔径。KEK 仿真给出的能散由出光前约 0.10% 增至出光后约 0.34%。在 800 MeV、9.75 mA 工况下，束流功率约 7.8 MW。如果全能量处损失百万分之一，局部损失功率约为 7.8 W。[3]')
p('把 FEL 输出的完整纵向相空间和能量尾部送入回收段跟踪，沿回流、减速和收集路径检查接受度与束流包络。通过调整回流路长、减速相位和纵向色散，控制收集端能量分布，并为偏离参考轨道的电子保留孔径裕量。先测无出光时的回收效率和束损基线，再逐级增加 FEL 提取量。联测分段束损、回流电流、射频功率和收集端能谱，验证持续运行和异常切断能力。[2,3,8]')

h('3.5 EUV 传输及照明',2)
p('波荡器输出经过多次反射、带通筛选和照明整形后到达中间焦点（IF），各环节的有效效率相乘。反射率随波长、入射角和偏振变化，镜面吸收引起的温升会改变面形和指向，污染也会降低反射率。高空间相干的 FEL 光经照明光学和掩模传播时，多路径的相干叠加可能形成散斑，影响曝光强度均匀性和图形质量。[2,10,13–15]')
p('按目标 IF 的带内功率、光束尺寸、发散角和光瞳反推源端需求。传输段联合优化镜面反射率、入射角和镜上光斑尺寸，采用冷却和低污染真空设计控制热漂移与反射率衰减。照明段可采用分面整形与分束延迟。分面整形配置目标照明光瞳，路径延迟可降低支路间的相干叠加。散斑抑制还依赖支路形成的图样差异及曝光时间内的积分效果。用同一组工况联测 IF 带内功率、瞳面分布、场均匀性、散斑对比度和剂量稳定性，比较各方案的抑斑效果和通光损失。具体限值按目标扫描机接口确定。[10,13–15]')

h('4 典型极紫外 FEL 方案')
h('4.1 基频 ERL SASE',2)
p('KEK 方案用光阴极枪以 162.5 MHz 产生 60 pC 束团。注入器将电子加速到约 11 MeV，超导主加速器升至 800 MeV，弧段压缩束团。电子束进入 28 mm 周期波荡器，由散粒噪声启动自放大自发辐射（SASE），再通过渐变段提高脉冲能量。出光后电子束返回主加速器减速，向射频场回馈能量。[2,3]')
diagram('architecture')
h('4.2 三次谐波 ERL RAFEL',2)
p('He 等提出用 331 MeV 电子束和 16 mm 周期波荡器，将基频共振设在约 40.5 nm，选择三次谐波产生 13.5 nm。分段相移器抑制基频增长，EUV 腔镜把部分 13.5 nm 辐射反馈给后续电子束团，形成再生放大。公开仿真设定 77 pC、130 MHz、约 370 A 峰值电流和 0.01% 相对能散。腔镜热漂移、反射率和电子束到达时间是工程验证重点。[6]')
math('λ₃ = λ₁/3')
p('在相同周期、磁场和目标输出波长下，三次谐波所需的洛伦兹因子约为基频的 1/√3。两项公开方案采用各自的波荡器周期和磁场，设计时分别代入共振式。[2,6]')

h('5 典型结果')
table(['装置或方案','代表结果','证据及功率位置'],[
    ('FELiChEM [1]','35 MeV 电子束产生 15 μm','中红外首次出光实测'),
    ('Jefferson Lab [4]','1.6 μm，14.2 kW','近红外平均光功率实测'),
    ('SXFEL [5]','约 92 eV，对应约 13.5 nm','低重复频率 SASE 光谱实测'),
    ('KEK ERL SASE [3]','109.4 μJ × 162.5 MHz ≈ 17.8 kW','2023 年起点至终点仿真，波荡器出口'),
    ('三次谐波 RAFEL [6]','15.6 μJ × 130 MHz ≈ 2.03 kW','2025 年仿真，腔输出端')],[4.4,6.6,5.3])
math('P(EUV) = U(pulse) × f')
p('KEK 工况的 60 pC × 162.5 MHz 对应 9.75 mA，800 MeV × 9.75 mA 对应 7.8 MW 束流功率。17.8 kW 源端光功率对应约 0.23% 的光能提取比例。仿真中，电子束能散由 FEL 入口的约 0.10% 增至出光后的约 0.34%，可据此设置回收线设计工况。[3]')

h('6 建议实现方案及典型系统参数')
p('采用基频 ERL SASE 作为高平均功率基准，按下表建立第一轮模拟和试验条件。三次谐波 RAFEL 作为降低加速能量的备选，先验证 EUV 腔镜热稳定性、腔反馈及束光同步。[2,3,6]')
table(['设计环节','基频参考参数','验证量'],[
    ('注入与主加速','约 11 MeV 注入，800 MeV 主束流，1.3 GHz 超导射频','工作电流下的束质、射频稳定性和束损'),
    ('电荷与频率','60 pC，162.5 MHz，9.75 mA','连续运行电流、阴极寿命和平均束流稳定性'),
    ('FEL 入口','束团长度约 39 fs，能散约 0.10%，归一化发射度约 2.0/0.9 mm·mrad（水平/垂直）','到达时间、发射度、切片能散和峰值电流'),
    ('波荡器与光输出','28 mm 周期。2023 年仿真采用 4% 渐变，输出 109.4 μJ、17.8 kW','13.5 nm 光谱、平均功率和运行稳定性'),
    ('回收及光刻接口','出光后约 0.34% 能散作为仿真工况。IF 指标按机型确定','回收效率和束损。IF 带内功率、光瞳、散斑、偏振和剂量稳定性')],[3.5,7.6,5.2])
p('研发按三步实施。第一步验证连续束流品质和减速回收。第二步逐级提高 EUV 平均功率，测量波长、光谱和运行稳定性。第三步接入真空反射与整形光路，同步测量中间焦点（IF）带内功率和照明质量。分束延迟、角度扫描及瞳面整形可结合实测散斑和通光效率选型。[10,13–15]')
math('P(IF) = P(source) × f(split) × T(route)')
p('f(split) 为该支路分到的源端功率比例，T(route) 为分光后带通、传输、整形和接口效率的乘积。以 17.8 kW 源端仿真值、单路取光和 20% 假设链路效率计算，IF 约为 3.56 kW。20% 是方案比较用的假设值，目标机型的带宽、光瞳和 IF 功率要求待确认。')
diagram('power')

h('7 极紫外 FEL 相对中红外 FEL 的主要改进')
bullet('将几十 MeV 的中红外工作点调整到与 13.5 nm 匹配的电子能量、波荡器周期和谐波阶次。')
bullet('在波荡器入口同时维持足够的峰值电流、低能散和低发射度，通过分段调节磁场提高光能提取。')
bullet('采用高重频注入、超导加速和能量回收，建立约 10 mA 平均电流下的连续运行能力。')
bullet('采用真空 EUV 反射、整形和去相干光学，将源端输出转化为满足光刻照明要求的 IF 带内功率。')
p('基频 ERL SASE 优先验证连续束流、能量回收和 EUV 功率扩展。三次谐波 RAFEL 优先验证 EUV 腔反馈与热稳定性。两条路线均以源端输出和光刻接口的联测结果评价工程可行性。')


h('参考资料')
refs = [
 '[1] Li 与 Jia，FELiChEM 首次出光，FEL 2019，doi:10.18429/JACoW-FEL2019-MOA07。',
 '[2] Kawata 等，ERL EUV FEL 设计，JMM 21，021210 (2022)，doi:10.1117/1.JMM.21.2.021210。',
 '[3] Nakamura 等，KEK EUV FEL 起点到终点模拟，IPAC 2023 THPM123，doi:10.18429/JACoW-IPAC2023-THPM123。',
 '[4] Jefferson Lab FEL 历史运行记录，https://www.jlab.org/FEL/。',
 '[5] Yang 等，SXFEL 约 92 eV 实验，JSR 31，1373 (2024)，doi:10.1107/S1600577524005812。',
 '[6] He 等，三次谐波 RAFEL，PRAB 28，030702 (2025)，doi:10.1103/PhysRevAccelBeams.28.030702。',
 '[7] 周可等，中红外太赫兹 FEL 设计，中国激光 50，1718001 (2023)，doi:10.3788/CJL230786。',
 '[8] 冯汉勋等，能量回收实验，强激光与粒子束 38，064001 (2026)，doi:10.11884/HPLPB202638.260170。',
 '[9] Konomi 等，EUV ERL 超导腔 HOM 研究，PRAB 26，121601 (2023)，doi:10.1103/PhysRevAccelBeams.26.121601。',
 '[10] 邱灏等，EUV 去相干光学模拟，物理学报 73，114201 (2024)，doi:10.7498/aps.73.20240335。',
 '[12] Huang 与 Kim，X 射线 FEL 理论综述，PRSTAB 10，034801 (2007)，doi:10.1103/PhysRevSTAB.10.034801。用于高增益量级判据。',
 '[13] Lüttgenau 等，部分相干 EUV 照明器设计，Applied Optics 61，3026 (2022)，doi:10.1364/AO.452204。用于瞳面与空间相干联调。',
 '[14] Stupakov 等，短 FEL 脉冲分束延迟抑散斑，FEL 2024 TUP142-WED，https://indico.jacow.org/event/72/contributions/5820/。模型研究。',
 '[15] Wei 等，FEL 相干对成像的影响，Chinese Physics B 34，054201 (2025)，doi:10.1088/1674-1056/adb734。成像模拟。',
]
for ref in refs:
    q=p(ref)
    q.paragraph_format.space_after=Pt(3)
    for run in q.runs: run.font.size=Pt(8.5)

footer = sec.footer.paragraphs[0]
footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
footer.add_run('FEL 光源专题研究  •  2026-11-05')
for run in footer.runs:
    run.font.size=Pt(8)
    run.font.color.rgb=RGBColor(102, 112, 120)

doc.core_properties.title = '13.5 nm 极紫外 FEL 光源的产生'
doc.core_properties.subject = 'FEL 中红外至极紫外路线、关键参数与光刻接口'
OUT.parent.mkdir(exist_ok=True)
doc.save(OUT)
print(OUT)
