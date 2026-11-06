"""Apply the supplied Cymer reference layout to the FEL production report.

Run after build_euv_fel_word.py. Only paragraph, run, table, page and
header/footer formatting is changed. Text and embedded images are verified.
"""
from pathlib import Path
from copy import deepcopy
from zipfile import ZipFile
from lxml import etree
from docx import Document
from docx.oxml.ns import qn

root = Path(__file__).resolve().parents[1]
base = root / 'tmp/fel_production_style'
source = root / 'outputs/2026-11-04_13.5nm_EUV_FEL原理技术与系统参数研究报告.docx'
reference = root / 'references/format/Cymer供应商清单_格式参考.docx'
out = base / 'styled.docx'
base.mkdir(parents=True,exist_ok=True)
NS = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
      'wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'}
ref = Document(reference)
para_samples = {'Title':ref.paragraphs[0], 'Heading1':ref.paragraphs[2],
                'Heading2':ref.paragraphs[16], 'Normal':ref.paragraphs[3],
                'Caption':ref.paragraphs[3], 'Equation':ref.paragraphs[3],
                'Figure':ref.paragraphs[3], 'Reference':ref.paragraphs[3]}
tab = ref.tables[1]
table_samples = [tab.cell(0,0).paragraphs[0],tab.cell(1,0).paragraphs[0]]

def replace_child(parent, tag, value):
    old = parent.find(qn(tag))
    if old is not None:
        parent.remove(old)
    if value is not None:
        parent.insert(0,deepcopy(value))

def setval(parent, tag, **values):
    item=parent.find(qn('w:'+tag))
    if item is None:
        item=etree.SubElement(parent,qn('w:'+tag))
    for key,val in values.items():
        item.set(qn('w:'+key),str(val))
    return item

def get_format(sample, keep_style=None, role='Normal'):
    ppr=deepcopy(sample._p.pPr)
    for t in ('pStyle','outlineLvl'):
        old=ppr.find(qn('w:'+t))
        if old is not None:ppr.remove(old)
    if keep_style:setval(ppr,'pStyle',val=keep_style)
    # Explicit values stop inheritance from the report's old styles.
    setval(ppr,'spacing',before=0,after=0,line=570,lineRule='exact')
    setval(ppr,'keepNext',val='1' if role in ('Title','Heading1','Heading2','Figure') else '0')
    setval(ppr,'keepLines',val='0')
    setval(ppr,'pageBreakBefore',val='0')
    setval(ppr,'widowControl',val='1')
    setval(ppr,'snapToGrid',val='0')
    if role in ('Title','Heading1','Heading2','Caption','Equation','Figure','Reference','Table'):
        for t in ('ind',):
            old=ppr.find(qn('w:'+t))
            if old is not None:ppr.remove(old)
        setval(ppr,'ind',firstLine=0,firstLineChars=0)
    if role in ('Equation','Figure'):
        setval(ppr,'jc',val='center')
    if role in ('Figure','Equation'):
        setval(ppr,'spacing',before=0,after=0,line=240,lineRule='auto')
    return ppr

def format_runs(p, sample, role='Normal'):
    reference_rpr=sample.runs[0]._r.rPr
    for run in p.iter(qn('w:r')):
        old=run.find(qn('w:rPr'))
        bold = old is not None and old.find(qn('w:b')) is not None
        new=deepcopy(reference_rpr)
        if role in ('Normal','Reference') and bold:
            setval(new,'b',val='1')
        setval(new,'color',val='000000')
        setval(new,'i',val='0')
        replace_child(run,'w:rPr',new)

with ZipFile(source) as archive:
    parts={i.filename:archive.read(i.filename) for i in archive.infolist()}
    infos=archive.infolist()
doc=etree.fromstring(parts['word/document.xml'])
body=doc.find(qn('w:body'))
roles=[]
for p in body.findall(qn('w:p')):
    pstyle=p.find('w:pPr/w:pStyle',NS)
    style=pstyle.get(qn('w:val')) if pstyle is not None else 'Normal'
    text=''.join(p.itertext()) if False else ''.join(p.xpath('.//w:t/text()',namespaces=NS))
    if p.find('.//w:drawing',NS) is not None:role='Figure'
    elif style=='Title':role='Title'
    elif style in ('Heading1','Heading2'):role=style
    elif style=='Caption':role='Caption'
    elif style=='Equation':role='Equation'
    elif p.find('w:pPr/w:jc',NS) is not None and p.find('w:pPr/w:jc',NS).get(qn('w:val'))=='center':role='Equation'
    elif text.startswith('[') and ']' in text[:5]:role='Reference'
    else:role='Normal'
    sample=para_samples[role]
    replace_child(p,'w:pPr',get_format(sample,style,role))
    if role=='Caption' and text.startswith('表 '):
        setval(p.find(qn('w:pPr')),'keepNext',val='1')
    format_runs(p,sample,role)
    roles.append((role,text[:35]))
    # A4 reference has a narrower text area. Preserve image aspect ratios.
    if role=='Figure':
        max_cx=5273675 # 8305 twips
        for drawing in p.findall('.//wp:inline',NS):
            ext=drawing.find('wp:extent',NS)
            oldx,oldy=int(ext.get('cx')),int(ext.get('cy'))
            if oldx>max_cx:
                newy=round(oldy*max_cx/oldx)
                ext.set('cx',str(max_cx));ext.set('cy',str(newy))
                for pic_ext in drawing.findall('.//a:xfrm/a:ext',NS):
                    pic_ext.set('cx',str(max_cx));pic_ext.set('cy',str(newy))

for table in body.findall(qn('w:tbl')):
    replace_child(table,'w:tblPr',tab._tbl.tblPr)
    grid=table.find(qn('w:tblGrid'))
    old=[int(e.get(qn('w:w'))) for e in grid]
    widths=[round(w*8305/sum(old)) for w in old]
    widths[-1]+=8305-sum(widths)
    for e,w in zip(grid,widths):e.set(qn('w:w'),str(w))
    for ri,row in enumerate(table.findall(qn('w:tr'))):
        trpr=row.find(qn('w:trPr'))
        if trpr is not None:
            for e in list(trpr):
                if e.tag in (qn('w:cantSplit'),qn('w:trHeight')):trpr.remove(e)
        if ri==0:
            if trpr is None:trpr=etree.SubElement(row,qn('w:trPr'))
            setval(trpr,'tblHeader',val='1')
        if trpr is None:trpr=etree.SubElement(row,qn('w:trPr'))

        for ci,cell in enumerate(row.findall(qn('w:tc'))):
            cell_sample=tab.cell(0 if ri==0 else 1,0)
            replace_child(cell,'w:tcPr',cell_sample._tc.tcPr)
            cp=cell.find(qn('w:tcPr'))
            setval(cp,'tcW',w=widths[ci],type='dxa')
            setval(cp,'vAlign',val='center')
            sample=table_samples[0 if ri==0 else 1]
            for p in cell.findall(qn('w:p')):
                replace_child(p,'w:pPr',get_format(sample,role='Table'))
                if ri==0:
                    setval(p.find(qn('w:pPr')),'keepNext',val='1')
                format_runs(p,sample,'Table')

sect=body.find(qn('w:sectPr'))
for tag in ('pgSz','pgMar','cols'):
    replace_child(sect,'w:'+tag,ref.sections[0]._sectPr.find(qn('w:'+tag)))
parts['word/document.xml']=etree.tostring(doc,xml_declaration=True,encoding='UTF-8',standalone=True)

styles=etree.fromstring(parts['word/styles.xml'])
for s in styles.findall(qn('w:style')):
    sid=s.get(qn('w:styleId'))
    if sid in ('Normal','Title','Heading1','Heading2','Caption','Header','Footer'):
        role=sid if sid in para_samples else 'Normal'
        replace_child(s,'w:pPr',get_format(para_samples[role],role=role))
        replace_child(s,'w:rPr',para_samples[role].runs[0]._r.rPr)
parts['word/styles.xml']=etree.tostring(styles,xml_declaration=True,encoding='UTF-8',standalone=True)
for path in list(parts):
    if path.startswith('word/header') and path.endswith('.xml'):
        parts[path]=etree.tostring(ref.sections[0].header._element,xml_declaration=True,encoding='UTF-8',standalone=True)
    if path.startswith('word/footer') and path.endswith('.xml'):
        parts[path]=etree.tostring(ref.sections[0].footer._element,xml_declaration=True,encoding='UTF-8',standalone=True)


with ZipFile(out,'w') as result:
    for info in infos:
        result.writestr(info,parts[info.filename])
before, after = Document(source), Document(out)
assert [p.text for p in before.paragraphs] == [p.text for p in after.paragraphs]
assert [[[c.text for c in r.cells] for r in t.rows] for t in before.tables] == [[[c.text for c in r.cells] for r in t.rows] for t in after.tables]
with ZipFile(source) as original, ZipFile(out) as formatted:
    changed={n for n in original.namelist() if original.read(n)!=formatted.read(n)}
    allowed={'word/document.xml','word/styles.xml','word/header1.xml','word/footer1.xml'}
    assert changed <= allowed, changed
assert len(after.inline_shapes)==9
source.write_bytes(out.read_bytes())
print('Applied reference formatting. Text and all embedded image bytes preserved.')
