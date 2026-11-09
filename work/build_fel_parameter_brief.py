"""Condense the FEL production report into three pages using its actual styles."""
from copy import deepcopy
from pathlib import Path
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / 'outputs/2026-11-04_13.5nm_EUV_FEL原理技术与系统参数研究报告.docx'
OUTPUT = ROOT / 'outputs/2026-11-09_13.5nm_FEL光源关键参数与方案_三页摘要.docx'
doc = Document(SOURCE)
title_sample = next(p for p in doc.paragraphs if p.style.name == 'Title')
heading_sample = next(p for p in doc.paragraphs if p.style.name == 'Heading 1')
body_sample = next(p for p in doc.paragraphs if p.style.name == 'Normal' and len(p.text)>60)
header_sample = deepcopy(doc.tables[0].cell(0,0))
cell_sample = deepcopy(doc.tables[0].cell(1,0))
table_properties = deepcopy(doc.tables[0]._tbl.tblPr)
samples = {'title':deepcopy(title_sample._p), 'heading':deepcopy(heading_sample._p),
           'body':deepcopy(body_sample._p)}
for element in list(doc._element.body):
    if element.tag != qn('w:sectPr'):
        doc._element.body.remove(element)

def paragraph(text, role='body', new_page=False):
    sample = samples[role]
    p = doc.add_paragraph()
    p._p.append(deepcopy(sample.find(qn('w:pPr'))))
    r = p.add_run(text)
    rpr = sample.find('.//'+qn('w:rPr'))
    if rpr is not None:
        r._r.insert(0,deepcopy(rpr))
    if new_page:
        p.paragraph_format.page_break_before=True
    return p

def table(rows, widths_cm=None):
    t=doc.add_table(rows=0, cols=len(rows[0]))
    t.autofit=False
    t._tbl.remove(t._tbl.tblPr)
    t._tbl.insert(0,deepcopy(table_properties))
    body_width=doc.sections[0].page_width-doc.sections[0].left_margin-doc.sections[0].right_margin
    widths=[Cm(w) for w in (widths_cm or [3.9])]
    widths.append(body_width-sum(widths))
    for column,width in zip(t.columns,widths):
        column.width=width
    for row_index,row in enumerate(rows):
        cells=t.add_row().cells
        for cell_index,text in enumerate(row):
            cell=cells[cell_index]
            sample=header_sample if row_index==0 else cell_sample
            cell._tc.remove(cell._tc.tcPr)
            cell._tc.insert(0,deepcopy(sample._tc.tcPr))
            cell.width=widths[cell_index]
            p=cell.paragraphs[0]
            ppr=sample.paragraphs[0]._p.pPr
            p._p.append(deepcopy(ppr))
            # Each row is independent and the header remains attached to its first row.
            p.paragraph_format.keep_with_next=(row_index==0)
            run=p.add_run(text)
            run._r.insert(0,deepcopy(sample.paragraphs[0].runs[0]._r.rPr))
        trpr=t.rows[-1]._tr.get_or_add_trPr()
        trpr.append(OxmlElement('w:cantSplit'))
        if row_index==0:
            trpr.append(OxmlElement('w:tblHeader'))
    return t

paragraph('EUV FEL 的难点与功率扩展','title')
paragraph('一、相较中红外的主要难点','heading')
paragraph('13.5 nm FEL 需同时实现短波共振、高增益、持续出束与 EUV 光学匹配。中红外至 EUV 的调整如下。')
table([
    ('难点','变化及技术方向'),
    ('共振波长','典型电子能量从几十增至数百 MeV。\n联合选取能量、周期、磁场与谐波阶次。'),
    ('短波增益','短波增益对能散、发射度和轨道误差更敏感。\n控制 FEL 入口切片束质，校准磁场与相位。'),
    ('束团压缩','短波高增益需高峰值电流，强压缩加重 CSR。\n联合优化射频啁啾、压缩色散及到达时间。'),
    ('放大架构','中红外高反腔到 EUV 面临反射、镜热与同步约束。\n基频采用单程 SASE，谐波采用再生放大。'),
    ('EUV 光学','真空反射引入镜热、污染与通光损失。\n配置冷却、分面整形及去相干，满足 IF 照明。'),
])
paragraph('固定周期与磁场时，λ∝1/γ²。15 μm→13.5 nm 缩短约 1111 倍，γ 增加约 33.3 倍。28 mm、aᵥ＝1.168 时，动能约从 23.5 MeV 提至 800 MeV。')

paragraph('二、平均功率提升的困难','heading',new_page=True)
paragraph('平均电流 I＝qf，平均光功率 P＝U·f。保持单束团出光时，提高重频可同比提高平均功率。增加电荷、提高提取率还会改变束质和回收工况。')
table([
    ('功率瓶颈','原因及解决方向'),
    ('高平均电流','高重频要求阴极寿命与持续稳定出束。\n优化激光分布和真空，逐级提重频并监测束质。'),
    ('单束团提能','增加电荷加重空间电荷与压缩段 CSR。\n渐变 K 提高提取率，并校核能散及能量尾。'),
    ('高流强加速','高阶模 HOM 与回流束耦合可触发 BBU。\n采用 HOM 阻尼、轨道控制和回流光学优化。'),
    ('兆瓦级回收','减速放大相对能散，增加束损和收集负担。\n匹配纵向接受度、回流路长与减速相位。'),
    ('可用光功率','镜热、传输和去相干损耗压低 IF 功率。\n联合优化冷却、反射镜数量与照明通光效率。'),
])
paragraph('800 MeV×9.75 mA 对应 7.8 MW 束功率。17.8 kW 光输出的提取比例约 0.23%。源端与 IF 分别验收，IF 功率＝源端功率×分光份额×支路效率。')

paragraph('三、典型方案与关键参数','heading',new_page=True)
table([
    ('参数/技术','基频 ERL SASE','三次谐波 ERL RAFEL'),
    ('代表研究','KEK，2023 仿真','He 等，2025 仿真'),
    ('能量/射频','800 MeV，注入约\n11 MeV，1.3 GHz','331 MeV'),
    ('电荷/重频','60 pC / 162.5 MHz','77 pC / 130 MHz'),
    ('平均电流','9.75 mA','约 10 mA'),
    ('入口束质','39 fs，能散 0.10%\nεₙ水平/垂直\n2.0/0.9 mm·mrad','峰值 370 A\n能散 0.01%'),
    ('周期/共振','28 mm，aᵥ＝1.168\n基频 13.5 nm','16 mm，40.5 nm\n基频的三次谐波'),
    ('提取/反馈','4% 渐变 K\n单程高增益','相移抑制基频\nEUV 腔反馈'),
    ('脉冲/功率','109.4 μJ / 17.8 kW\n波荡器出口','15.6 μJ / 2.03 kW\n腔输出端'),
    ('验证重点','出光后能散 0.34%\n回收、束损及稳定性','腔镜热稳定性\n同步及反馈增益'),
], widths_cm=[3.0,5.8])
paragraph('两方案均输出 13.5 nm。参数和功率均为仿真工况。IF 带宽、光瞳、偏振、散斑与剂量限值按机型确认。来源沿用原报告 [2、3、6、9、12–15]。')
doc.core_properties.title='EUV FEL 的难点与功率扩展'
doc.core_properties.subject='中红外至EUV难点、功率扩展困难与典型方案'
doc.save(OUTPUT)
print(OUTPUT)
