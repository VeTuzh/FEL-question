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
        d.text((620,520),'电子束经减速段回收；* 为公开模拟值，不是实测',font=small,fill=gray)
    else:
        d.text((70,25),'同一功率口径下的损失账本',font=cn,fill=dark)
        names=[('波荡器出口','源端功率'),('分光与带通','分光 × 带通'),('传输镜组','镜组透过率'),('整形去相干','有效透过率'),('中间焦点','IF 功率')]
        xs=[70,425,780,1135,1490]
        for i,(a,b) in enumerate(names):
            box(xs[i],125,285,150,a,b)
            if i<4: arrow(xs[i]+285,200,xs[i+1]-12,200)
        d.text((70,335),'各段同时记录谱宽、指向、光瞳、相干、热负荷和时间稳定性。',font=small,fill=gray)
        d.text((70,390),'示例：17.8 kW × 假设链路效率 20% = 3.56 kW IF；20% 仅用于灵敏度分析。',font=small,fill=gray)
    path = TMP / f'figure_{which}.png'
    im.save(path, dpi=(300,300))
    para=doc.add_paragraph()
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

doc.add_paragraph('13.5 nm EUV FEL 光源技术可行性初步研究报告', style='Title')
p('中红外技术基线至极紫外光刻应用  |  2026 年 11 月 4 日  |  可研初判稿')
p('研究目标：以中红外 FEL 装置为认知基线，回答怎样将共振波长推进到 13.5 nm，同时维持足够的平均功率、可靠的能量回收及可用于光刻机的照明。报告给出设计方程、公开方案参数、风险分解和分阶段试验判据。')
h('执行摘要')
p('路线结论：13.5 nm 不是对已输出的中红外激光作外置倍频，而是改变电子束能量、波荡器和放大机制，在电子束与辐射场相互作用过程中直接生成 EUV。中红外出光验证只能迁移电子枪、加速器、束流诊断与能量回收的部分经验，不能证明高功率 EUV 光学与光刻接口。',boldlead='原理判断。')
p('优先路线：以连续波超导 ERL 加基频 SASE、分段渐变波荡器作为高平均功率基准；公开起点到终点模拟为 800 MeV、9.75 mA、17.8 kW 源端。以 331 MeV 三次谐波 ERL RAFEL 作为紧凑路线，其约 2.03 kW 为模型结果，EUV 腔镜和同步是新的主风险。[2,3,6]')
p('决策条件：先取得目标扫描机 IF 功率、带宽、光瞳、偏振、相干、剂量稳定性、可用率和空间接口，再倒推源端功率及装置占地。尚无公开证据证明千瓦级 13.5 nm FEL 与工业曝光机实现一体化长时运行。')
table(['结论等级','本报告采用的代表事实','对研发决策的含义'],[
    ('已实测','15 μm FEL 出光；近红外 14.2 kW；约 13.5 nm 低频 SASE 出光','能力分散在不同装置，不能叠加视为整机验证'),
    ('设计与模拟','KEK 17.8 kW；He 等 2.03 kW','可供参数闭合与试验方案设计，不能作为采购保证'),
    ('待确认','目标扫描机 IF、光瞳、带宽、运行可用率及光学效率','决定源端余量、去相干方案和整机经济性')],[2.5,7.1,6.7])

h('研究边界与可研判据')
p('本报告仅对 13.5 nm 的科学与工程可行性进行初判，不承诺整机规格、建设工期或成本。可研验收应分为六层：共振波长、脉冲能量及平均功率、束流能量回收、EUV 传输和 IF 功率、成像照明质量、长时间可用率。每层记录测量面、占空比与不确定度，不能跨层替代。')
table(['评价量','定义或测量面','本轮状态'],[
    ('Pᵤ','波荡器出口、全部脉冲的时间平均 EUV 功率','有公开模拟值'),
    ('PIF','光刻机中间焦点、指定带宽内的时间平均功率','目标与实测待确认'),
    ('照明质量','光瞳角分布、场均匀性、偏振、相干、谱纯度、剂量噪声','须同 IF 功率联合验收'),
    ('工程性','墙插电耗、束损、光学热负荷、可用率及维护间隔','无集成系统实测')],[2.4,9.0,4.9])

h('1 FEL 光源的产生机理')
p('自由电子激光由相对论电子束穿过周期磁场产生。电子在波荡器内横向振荡，与共传播辐射场交换能量；微聚束使相位相近的电子相干辐射并放大光场。电子束能量、波荡器周期与磁场强度决定共振波长，束流亮度与波荡器长度决定增益、饱和和可提取功率。[1–3]')
math('λₙ = λᵤ(1 + K²/2) / (2nγ²)；γ = 1 + Eₖ/(0.511 MeV)')
p('式中 λᵤ 为波荡器周期；K 为按峰值磁场定义的无量纲强度；n 为谐波阶次，基频 n=1；Eₖ 为电子动能。若论文采用 RMS 定义 aᵥ，则平面波荡器的分子写为 1+aᵥ²，不能将两种 K 直接混用。近似有 λ∝λᵤ(1+K²/2)/(nEₖ²)。')
p('中红外 FELiChEM 在 35 MeV 得到 15 μm 首次出光；JLab 在 1.6 μm 实测 14.2 kW 平均功率；SXFEL 在约 92 eV 测到约 13.5 nm 的 SASE 谱，重复频率低于 10 Hz。这分别说明波长可调、中近红外高平均功率和 EUV 可出光，尚未把三种能力集于一个工业装置。[1,4,5]')
math('K = 0.934 × B₀[T] × λᵤ[cm]；Eγ[eV] = 1239.84/λ[nm]')
p('13.5 nm 对应光子能量约 91.84 eV。K 由波荡器峰值磁场与周期决定；选更短周期或更弱 K 可降所需电子能量，但会降低耦合、提高磁隙和磁场误差难度。谐波 n>1 也降低电子能量，却需抑制基频并保持足够谐波增益。[2,6]')
diagram('architecture')
table(['实例','波长和工作点','证据等级与边界'],[
    ('FELiChEM','15 μm；35 MeV 首次出光','实测中红外；不能外推 EUV 平均功率'),
    ('JLab ERL FEL','1.6 μm；14.2 kW','历史实测近红外；不是 EUV'),
    ('SXFEL','约 92 eV，即约 13.5 nm；<10 Hz','EUV SASE 实测；不是千瓦级'),
    ('KEK ERL SASE','13.5 nm；17.8 kW 源端','设计与起点到终点模拟'),
    ('三次谐波 RAFEL','13.5 nm；约 2.03 kW 源端','2025 年模型模拟')], [3.2,6.0,7.1])

h('2 从中红外走向 13.5 nm 需要优化的参数')
p('将 15 μm 缩至 13.5 nm，波长比约 1111。只作同一波荡器周期、同一磁场和基频的尺度估算，电子总能量需约增加 √1111≈33.3 倍：采用 KEK 的 28 mm 周期和 aᵥ=1.168，约 23.5 MeV 对应 15 μm，约 800 MeV 对应 13.49 nm。这不是将既有中红外装置直接升级的工程方案，因为腔、束线、枪与加速器架构均不同。[2,3]')
math('λ₁ = 28 mm × (1+1.168²)/(2γ²) ≈ 13.49 nm；Eₖ = 800 MeV')
p('这里使用的是同一 28 mm 波荡器的思想实验，且在约 23.5 MeV 时电子已经相对论化但 γ=1+Eₖ/0.511 不能直接用 Eₖ/0.511 代替。实际中红外装置可能采用不同波荡器、光腔和束流能量，不能用 33.3 倍作为采购或改造系数。')
math('Eₖ(λ) = 0.511 MeV × {√[λᵤ(1+aᵥ²)/(2λ)] − 1}')
p('用上式反算给出 13.5 nm 所需电子能量；改变 n、λᵤ、K 还有多组可行组合。共振匹配只保证“波长可达”，高增益和长时平均功率还取决于电子束六维相空间、束流平均电流和波荡器误差。')
table(['参数','中红外常见做法','13.5 nm 设计要求及影响'],[
    ('电子能量','几十 MeV 可覆盖部分中红外','基频多进入数百 MeV 至约 0.8 GeV；三次谐波例约 331 MeV'),
    ('峰值电流与束质','光学腔容许多次通过','单程高增益需高峰值电流、低归一化发射度和低能散'),
    ('平均电流及频率','可采用宏脉冲和有限占空比','千瓦级需 MHz 级至百 MHz 级持续微脉冲'),
    ('波荡器','波长较长，磁铁容差相对宽','短周期、磁场误差和轨道容差收紧，常需渐变段'),
    ('光学反馈','中红外高反射腔成熟','EUV 镜面反射、热漂移、污染及腔同步成为瓶颈'),
    ('接口','实验线出光即可','需满足中间焦点、带宽、远心性、瞳面填充及稳定性')],[3.2,5.0,8.1])
p('关键参数联动：提高 Eₖ 可缩短波长，却使设施规模与束流功率增长；缩小 λᵤ 可降低 Eₖ，却增加磁场与束流孔径约束；提高 K 增强耦合，但在固定 Eₖ 时共振波长变长；提高峰值电流利于增益，却加重压缩和 CSR；提高平均电流提升功率，却推高 SRF、HOM、束损和回收风险。路线选择因此必须联合优化，而非单参数取极值。')
h('2.1 高增益条件与缩短波长的束质代价',2)
p('以一维高增益模型给出量级判据。Pierce 参数 ρ 衡量电子束与光场耦合强度，具体数值依赖峰值电流、束斑、波荡器耦合和三维修正，不能只由共振式求得；实际设计应以三维时变 FEL 程序和真实束流切片分布求解。[12]')
math('L(gain) ≈ λᵤ/(4π√3 ρ)；L(sat) ∼ (15–25) L(gain)')
p('公式中的饱和长度系数量级适用于理想高增益 SASE，工程结果受衍射、发射度、能散、轨道误差、尾场和渐变段影响。通常要求切片相对能散 σδ 与 ρ 同一量级或更小；发射度越大，短波长处的模式耦合越差。因此，极紫外路线需要控制的是“波荡器入口实际切片束质”，并非仅电子枪铭牌发射度。')
math('σδ ≲ ρ；ε(geom) = εₙ/γ；σₓ ≈ √[βₓ ε(geom)]')
p('后一式把归一化发射度、光学 β 函数与束斑尺寸相连，可用于约束聚焦晶格和波荡器孔径。该量级判据并非对所有三维情形严格充分。')

h('3 参数优化所需系统和主要困难')
h('3.1 电子束与加速器',2)
p('高电荷、高频率的光阴极电子枪必须同时满足平均电流、阴极寿命及注入发射度；压缩器提高峰值电流，却可因相干同步辐射 CSR、微束团化与色散而增加能散及发射度。应分阶段测量枪出口、压缩后与 FEL 入口的切片发射度、切片能散、峰值电流和到达时间抖动。[2,3]')
math('Iavg = q f；Pbeam [kW] = Eₖ [MeV] × Iavg [mA]')
p('KEK 参考点 q=60 pC、f=162.5 MHz，I_avg=9.75 mA；800 MeV 束流功率约 7.8 MW。超导 RF 加速及 ERL 减小持续 RF 供电需求，却必须处理高阶模 HOM、束流破裂 BBU、束损、能量回收纵向接受度及失超保护。FEL 提取后的能散上升会压缩可回收工作窗。[2,3,9]')
table(['关键调整','风险机制','可行解决路径与验证量'],[
    ('高亮度 CW 电子枪','高电荷与高频率下阴极寿命、发射度、暗电流及发射热负荷互相制约','以真实 60 pC × 162.5 MHz 的持续工作时间、切片发射度、暗电流和阴极更换周期验收'),
    ('压缩与到达时间','CSR、纵向空间电荷及微束团化放大能散和抖动','优化压缩因子、晶格色散与激光/RF 同步，测 FEL 入口切片参数'),
    ('CW 超导 RF','高阶模耦合、BBU、失超与冷却功耗','HOM 阻尼、BBU 阈值测量、独立束损联锁和低温热负荷预算'),
    ('ERL 回收','FEL 后能散、能量尾与轨道误差使减速束流易丢失','设计纵向接受度与大孔径回收线，记录每段束损、回收效率和 dump 能谱')],[3.0,5.6,7.7])
h('3.2 波荡器与光学接口',2)
p('单程 SASE 对 EUV 腔镜依赖较少，但输出时域和横向相干性强、脉冲和指向波动需系统测量。渐变 K 可在饱和后继续提取电子能量，但束流能散、轨道误差和相位误差将削弱增益。EUV 传输需真空、多层反射镜的热形变与污染控制，以及受控的谱过滤和瞳面整形。[2,3,10]')
math('P(abs,mirror) ≈ P(incident) × [1 − R − T]')
p('R 与 T 必须按真实光谱、入射角和偏振积分，不能直接用峰值反射率估算镜热。即使单面损耗比例较小，千瓦级束光的局部吸收、面形漂移与污染膜演化也会改变下一镜面的谱和瞳面，因此应以热态迭代光线追迹和实测热负荷闭环设计。RAFEL 的腔镜热约 350 W 是其模型设定下的估计。[6]')
math('P(IF) = P(u) × f(split) × T(band)；× T(transport) × T(decoh) × T(interface)')
p('P_u 为波荡器出口平均功率，f_split 为分光份额，T 为各段实测透过率。必须定义在中间焦点 IF 的测量截面，逐项记录功率损耗、光谱、时间波动与偏振；缺乏具体扫描机接口数据时，不得由源端 17.8 kW 直接声称可供某台光刻机使用。')
p('高空间相干性可使掩模和光学面粗糙度产生散斑或非均匀照明；低相干并非越低越好，瞳面形状还要支持成像。可研究延迟分束、扫描相位或角度、多模混合、扩散元件及微粗糙表面散射等方式，并联测干涉对比度、瞳面角分布、视场均匀性及吞吐损失。现有去相干模拟给出的约 37.6% 透过率只属于特定模型，不能代入本项目功率预算。[10]')
diagram('power')
h('3.3 照明相干性与扫描机匹配',2)
p('FEL 的高空间相干性会使掩模多层膜相位粗糙度和照明结构产生稳定干涉，表现为散斑、图形边缘及线宽噪声；其强度取决于掩模、瞳面和时间积分条件，不能用单一“全局相干度”代替。扫描机需要的是设计指定的部分相干照明与可编程瞳面，而不是任意漫射。[10,13–15]')
math('C = σ(I)/⟨I⟩ ≈ 1/√M(eff)；M(eff) ≈ 1/C²')
p('该散斑对比度关系仅适用于等权、统计独立且被曝光时间有效平均的 M_eff 个模式。若以 C≤0.10 作分析目标，理想模型需约 100 个独立模式；这不是扫描机指标，也不能由分束数量直接等同于独立模式数。必须在曝光积分窗口内测量 C、光瞳填充及成像 LER/LWR。')
math('τc ∼ λ²/(c Δλ)；Δt ≳ max(τc, τp)')
p('τ_c 为谱宽对应的相干时间，τ_p 为相关脉冲持续时间。时间延迟分束使副束的干涉项在曝光积分中消失；实际所需延迟要由互相关和光路色散测试决定。角度/相位扫描可在多次曝光样本中平均，但还要保证目标瞳面和剂量控制。被动混合会牺牲透过率，增加热负荷与像差。')
table(['可研接口','建议测量与设计输入','待确认项'],[
    ('功率和光谱','IF 指定带宽平均功率、带外谱、波长漂移、单脉冲能量分布','扫描机 IF 目标与带宽'),
    ('瞳面及偏振','入瞳角分布、环形/多极模式、均匀性、偏振态','NA、照明模式和偏振策略'),
    ('时间与空间相干','互相干函数、散斑对比度、曝光积分后的 LER/LWR','掩模和工艺可接受阈值'),
    ('机光接口','源尺寸/发散角、IF 位置/数值孔径、真空、振动、热漂移','具体设备机械与光学接口'),
    ('运行','剂量稳定性、源和传输可用率、污染维护周期','生产节拍及可靠性目标')],[2.9,7.0,6.4])

h('4 目前典型技术方案')
h('4.1 基频 ERL SASE 加渐变波荡器',2)
p('架构为高重复频率电子枪、注入器、1.3 GHz 超导主直线加速器、压缩与束流诊断、约 28 mm 周期波荡器、EUV 输出光学、减速能量回收及束流收集器。KEK 设计提供可量化的高功率基准，但仍属于方案与模拟。[2,3]')
table(['节点','公开工作点','验证性质'],[
    ('电子枪与注入','60 pC × 162.5 MHz＝9.75 mA；注入约 10.5–11 MeV','设计值；连续运行和寿命待证'),
    ('主加速与 FEL 入口','800 MeV；1.3 GHz SRF；约 39 fs、0.1% 能散、εₙx/εₙy＝2.0/0.9 mm mrad','2023 模拟入口分布'),
    ('波荡器','28 mm 周期；2022 版本约 18 节、每节约 5 m','不同版本应分别计算'),
    ('输出','2022：88.5 μJ、14.4 kW，2% 渐变；2023：109.4 μJ、17.8 kW，4% 渐变','均为模型模拟，非实测 EUV 千瓦级')],[3.1,8.4,4.8])
h('4.2 三次谐波 ERL RAFEL',2)
p('He 等提出基频约 40.5 nm、三次谐波 13.5 nm 的再生放大。高选择性 Mo/Si 腔镜将 13.5 nm 反馈到下一电子束；这不是把已产生的中红外光经外置倍频器直接转换。更低能量换来 EUV 腔镜、输出耦合孔和 130 MHz 同步的工程风险。[6]')
table(['节点','模型输入或结果','设计风险'],[
    ('电子束','331 MeV；77 pC × 130 MHz≈10.0 mA；370 A 峰值','约 3.31 MW 束流功率；需回收'),
    ('束质','25 μm rms 束长；归一化发射度 1 mm mrad；能散 0.01%','入口指标严苛'),
    ('波荡器','16 mm 周期；K_peak＝1.5；n＝3；五节各 2 m','短周期制造及谐波耦合'),
    ('光腔','23.077 m；假设镜面反射 70%；约 72% 输出耦合','腔长、热漂移与污染'),
    ('输出','15.6 μJ × 130 MHz≈2.03 kW；镜面热负荷约 350 W','模拟而非运行记录')],[3.1,8.2,5.0])
p('在此方案中，n=3 将基频共振约 40.5 nm 对应到 13.5 nm。分段相移器抑制基频而保持三次谐波耦合，EUV 腔的反馈增益需要超过每回合损耗；70% 反射和 72% 输出耦合是论文计算设定，须经原型镜片、热态光学和同步误差试验证实。[6]')
math('λ₃ = λ₁/3；Eₖ,3 ≈ Eₖ,1/√3')
p('能量比例仅在相同波荡器周期与 K、且 γ 足够大时成立。He 与 KEK 方案分别采用 16 mm 与 28 mm 周期，故不能用 800/√3 MeV 直接复算 331 MeV；须代入各自的磁场和谐波阶次。')
h('4.3 备选路线及边界',2)
p('常规直线加速器的低重频 SASE 可先完成波长、谱和成像实验，但平均功率不足以证明量产接口；外部种子、高次谐波和储存环方案可提供特定光谱或占地优势，却分别面临种子系统、谐波效率及储存环束流扰动平衡。短周期波荡器还面临磁隙、束流孔径、相位误差与制造一致性。此阶段不宜给这些方案赋予未经证明的源端千瓦级成熟度。[5,11,12]')

h('5 典型结果及其可比性')
math('P(u) = U(pulse) f；η(extract) ≈ P(u)/Pbeam')
p('KEK 2023 版本的 109.4 μJ×162.5 MHz＝17.78 kW，除以 7.8 MW 约得 0.23% 提取率。RAFEL 15.6 μJ×130 MHz＝2.03 kW。两项均为波荡器出口或模型定义的源端结果，不是 IF 功率。中红外装置的 100 W 宏脉冲功率若占空比不超过 10%，时间平均功率上限仅约 10 W；1 mA、500 μs、1 Hz 的能量回收实验时间平均电流约 0.5 μA，不能表述为 1 mA 连续运行。[3,6–8]')
math('I(time) = I(macro) × τ(train) × f(train)；P(time) = P(macro) × τ(train) × f(train)')
p('KEK 2023 模拟还报告 FEL 后能散由入口约 0.1% 升至约 0.34%，在其模型孔径内无束损；它同时指出未计入的误差和束流动力学仍需研究。这不是“10 mA CW 回收并具工业可用率”的实验证据。[3]')
table(['路线','13.5 nm 证据','功率与关键缺口'],[
    ('SXFEL 短脉冲 SASE','约 92 eV 谱实测','出光已证；<10 Hz，平均功率不符合目标'),
    ('KEK ERL SASE','17.8 kW 源端模拟','最高功率基准；10 mA 级持续高亮度与 ERL 待集成验证'),
    ('三次谐波 ERL RAFEL','约 2.03 kW 源端模拟','电子能量较低；EUV 镜热、反馈和同步待验证'),
    ('储存环 FEL','公开理论与模拟','可借基础设施；辐射阻尼与束流扰动平衡需验证')],[4.1,4.2,8.0])

h('6 未来实现路径与初步系统参数')
p('建议采用“两级判据”：先以 KEK 量级 ERL SASE 作为基频高平均功率验证平台，建立同一测量口径的 ERL、FEL、EUV 传输与 IF 功率账本；并行在小功率平台验证三次谐波 RAFEL 的 EUV 腔镜热稳定性。以下参数为公开方案参考点和本报告提出的验证指标，不构成经工程确认的整机规格。')
table(['设计层级','初步输入或目标','验收问题'],[
    ('电子束','约 800 MeV、60 pC、162.5 MHz、平均 9.75 mA','持续亮度、束损、阴极与 SRF 寿命'),
    ('FEL 入口','约 39 fs、0.1% 相对能散、2.0/0.9 mm mrad 归一化发射度','重复频率全范围下的切片束质'),
    ('波荡器','28 mm；aᵥ=1.168；分段调 K，目标 13.5 nm','三维增益、误差预算、残余能散'),
    ('源端','参考模拟 109.4 μJ、17.8 kW','长时功率、谱宽、指向、稳定性'),
    ('ERL','约 7.8 MW 束功率，减速回收','回收效率、能量接受度、HOM/BBU 阈值、失束'),
    ('光刻接口','IF 功率、带宽、光瞳、偏振、相干与稳定性待确认','样机光学传输及掩模/晶圆试验')],[3.0,6.3,7.0])
p('设计迭代按共振计算→3D FEL 增益与渐变优化→束流和 ERL 失稳/束损→EUV 镜热与污染→去相干和照明瞳面→IF 实测闭环进行。阶段门应将 13.5 nm 谱、源端平均功率、整机电耗、IF 功率、瞳面均匀性、散斑和运行可用率放在同一数据记录中。')
math('P(grid) = P(RF,net) + P(cryo) + P(aux)；η(wall) = P(IF)/P(grid)')
p('P_RF,net 应在能量回收后的 RF 电源边界测量，避免把 P_recovered 在同一账本里再次扣除；低温、辅助设备和基础设施均计入电耗。KEK 2023 论文约 7 MW 电耗是设计估算，不能作为已测系统墙插效率。[3]')
h('6.1 源端至 IF 的灵敏度分析',2)
math('P(u,required) = N × P(IF,target) / T(route)')
p('上式仅适用于 N 条等功率、等透过率支路：每路分光为 1/N，T(route) 是分光之后至 IF 的带通、传输、去相干和接口透过率乘积，因而不能再把 1/N 计入 T(route)。下表的“总 IF 效率”则定义为所有支路 IF 功率之和除以源端功率。下表以 17.8 kW 模拟源端作情景计算，假定总 IF 效率为 5%、10%、20%、40%；这些值均非实测，不包含长期停机率。')
table(['假设链路效率','17.8 kW 对应总 IF','若单台需要 1 kW 可分配台数上限*'],[
    ('5%','0.89 kW','0 台'),('10%','1.78 kW','1 台'),('20%','3.56 kW','3 台'),('40%','7.12 kW','7 台')],[4.3,5.3,6.7])
p('* 向下取整，仅为功率算术上限，未计备用余量、光瞳整形、机台差异与可用率。若要以 17.8 kW 同时提供 10 台各 1 kW IF，所需总链路效率至少约 56.2%；具体目标应以用户扫描机指标代入而非沿用本例。')
h('6.2 分阶段验证与退出准则',2)
table(['阶段门','试验交付证据','未达标时的处理'],[
    ('G0 需求与模型','扫描机接口条件；共振、3D FEL、ERL 与光学统一误差预算','缺接口不冻结功率、占地和成本'),
    ('G1 束流与回收','规定电荷、频率和持续时间下的切片束质、束损与回收热平衡','优先回退电荷/压缩，不把脉冲模式当 CW'),
    ('G2 EUV 产生','13.5 nm 谱、脉冲能量、平均功率、渐变增益与多小时稳定性','区分波长达到与功率达到'),
    ('G3 光学接口','IF 功率、瞳面、散斑、热漂移、污染和剂量控制同时达标','优化链路损耗或减少分光数'),
    ('G4 系统工程','可靠性、维护、能耗、安全联锁和成本实测','重新评估商业化边界')],[3.0,8.1,5.2])

h('7 与中红外 FEL 相比的改进及决策结论')
bullet('能量与束质：从几十 MeV 量级转到数百 MeV 级，并将高峰值、低发射度和低能散保持到波荡器；三次谐波可降低能量但提高反馈复杂度。')
bullet('功率架构：从宏脉冲或低频验证转到百 MHz 量级持续微脉冲和约 10 mA 平均电流；使用 ERL 控制兆瓦级束流功率的损耗与 RF 负担。')
bullet('光学架构：从成熟红外腔与输出光学转到真空 EUV 多层膜、热/污染管理以及照明瞳面和相干性控制。')
bullet('验收口径：用源端→传输→去相干→中间焦点→晶圆的实测账本判断可用性。缺少具体扫描机 IF、光谱、瞳面、偏振和稳定性接口时，工程选型保持待确认。')
p('研发建议：优先验证 ERL SASE 的持续束流、能量回收和源端功率；同步建立 EUV 光学及去相干样机。RAFEL 以腔镜热漂移、同步容差和污染寿命为独立阶段门。公开资料尚不支持宣称任何路线已经满足工业 13.5 nm 光刻机全套照明需求。')
p('组织实施上建议由加速器与束流团队负责枪、SRF、压缩及 ERL；FEL 团队负责波荡器、渐变和波长稳定；EUV 光学团队负责镜热、污染、分光、瞳面和去相干；整机团队明确 IF、剂量与系统可靠性接口。跨团队统一功率测量面和占空比，才能避免源端优化与光刻应用脱节。')
table(['技术问题','优先级依据','拟解决动作'],[
    ('10 mA 级 CW 高亮度束流和低损失 ERL','电子束约 7.8 MW，束损与 HOM/BBU 直接决定安全和能耗','先做 G1，不达到持续束质不承诺千瓦 EUV'),
    ('EUV 镜热、污染和功率链','每次反射与去相干损耗决定 PIF 和镜寿命','早建功率账本和热态光学原型'),
    ('相干性与照明瞳面','过度相干会影响散斑和成像，过度去相干损失功率','以 G3 的光瞳、散斑、成像和透过率联合优化'),
    ('方案取舍','SASE 尺度大，RAFEL EUV 反馈有新风险','SASE 为主线，RAFEL 以镜和同步原型设置独立门槛')],[4.0,6.0,6.3])

h('资料与证据边界')
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
 '[11] Lee 等，储存环 FEL 模拟，JSR 27，864 (2020)，doi:10.1107/S1600577520005676。',
 '[12] Huang 与 Kim，X 射线 FEL 理论综述，PRSTAB 10，034801 (2007)，doi:10.1103/PhysRevSTAB.10.034801。用于高增益量级判据。',
 '[13] Lüttgenau 等，部分相干 EUV 照明器设计，Applied Optics 61，3026 (2022)，doi:10.1364/AO.452204。用于瞳面与空间相干联调。',
 '[14] Stupakov 等，短 FEL 脉冲分束延迟抑散斑，FEL 2024 TUP142-WED，https://indico.jacow.org/event/72/contributions/5820/。模型研究。',
 '[15] Wei 等，FEL 相干对成像的影响，Chinese Physics B 34，054201 (2025)，doi:10.1088/1674-1056/adb734。成像模拟。',
]
for ref in refs:
    q=p(ref)
    q.paragraph_format.space_after=Pt(3)
    for run in q.runs: run.font.size=Pt(8.5)
p('资料核对范围：用户提供的部分论文只核对至指定页，其他结论以公开论文摘要、正文和原始装置资料为依据；专利检索、具体扫描机接口与工业可用率不属于本报告已验证成果。')

footer = sec.footer.paragraphs[0]
footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
footer.add_run('FEL 光源专题研究  •  2026-11-04  •  内部研发初判')
for run in footer.runs:
    run.font.size=Pt(8)
    run.font.color.rgb=RGBColor(102, 112, 120)

doc.core_properties.title = '13.5 nm EUV FEL 光源技术可行性初步研究报告'
doc.core_properties.subject = 'FEL 中红外至极紫外路线、关键参数与光刻接口'
OUT.parent.mkdir(exist_ok=True)
doc.save(OUT)
print(OUT)
