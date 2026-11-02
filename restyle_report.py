from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from copy import deepcopy
from pathlib import Path
import hashlib

ref=Path('C:/Users/LISIMING/Desktop/计算机组成原理—物联网三班李思民/2405040303-李思民-计组课设报告.docx')
src=Path('output/智能终端系统及应用开发课程设计报告_2405040303_李思民.docx')
d=Document(src); r=Document(ref)
def props(p, sample):
 old=p._p.pPr
 if old is not None: p._p.remove(old)
 pr=deepcopy(sample._p.pPr)
 for n in ['pStyle','numPr','sectPr']:
  for e in list(pr.findall(qn('w:'+n))): pr.remove(e)
 p._p.insert(0,pr)
def font(p,size=12,bold=False,name='宋体'):
 for run in p.runs:
  if run._r.rPr is not None: run._r.remove(run._r.rPr)
  run.font.name=name; run.font.size=Pt(size); run.font.bold=bold; run.font.color.rgb=RGBColor(0,0,0)
  f=run._r.get_or_add_rPr().rFonts
  for key in ['ascii','hAnsi','eastAsia','cs']: f.set(qn('w:'+key),name)
def clean_style(s):
 for n in ['pPr','rPr']:
  e=s.element.find(qn('w:'+n))
  if e is not None: s.element.remove(e)
 s.font.name='宋体'; s.font.size=Pt(12); s.font.color.rgb=RGBColor(0,0,0)
 s.element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'宋体')
for name in ['Normal','Title','Heading 1','Heading 2','Heading 3','Caption']:
 clean_style(d.styles[name])
main_i=next(i for i,p in enumerate(d.paragraphs) if p.text=='一　实验题目')
for i,p in enumerate(d.paragraphs):
 role=p.style.name; text=p.text.strip()
 if i==0:
  props(p,r.paragraphs[0]); p.paragraph_format.right_indent=Pt(0); p.paragraph_format.space_before=Pt(36); font(p,18)
 elif role=='Title':
  props(p,r.paragraphs[3]); p.paragraph_format.right_indent=Pt(0); p.alignment=1; p.paragraph_format.first_line_indent=Pt(0); p.paragraph_format.space_before=Pt(28); p.paragraph_format.space_after=Pt(24); font(p,26)
 elif i==2:
  props(p,r.paragraphs[18]); font(p,14,True); p.paragraph_format.space_after=Pt(28)
 elif i in range(3,9):
  props(p,r.paragraphs[10]); p.paragraph_format.right_indent=Pt(0); p.paragraph_format.first_line_indent=Cm(.4); p.paragraph_format.left_indent=Pt(0)
  ind=p._p.pPr.find(qn('w:ind'))
  for a in ['rightChars','firstLineChars']: ind.attrib.pop(qn('w:'+a),None)
  font(p,16,False,'微软雅黑')
 elif role.startswith('Heading'):
  props(p,r.paragraphs[19 if role=='Heading 1' else 27]); font(p,14 if role=='Heading 1' else 12,True)
  p.paragraph_format.keep_with_next=True
  if text in ['目录','一　实验题目']: p.paragraph_format.page_break_before=True
 elif role=='Caption':
  props(p,r.paragraphs[329]); font(p,12)
 elif 9<i<main_i:
  props(p,r.paragraphs[20]); p.paragraph_format.first_line_indent=Pt(0); font(p,12)
 else:
  iscode=any(run.font.name=='Consolas' for run in p.runs)
  props(p,r.paragraphs[22]); font(p,12)
  if iscode:
   p.paragraph_format.first_line_indent=Pt(0); p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0); font(p,10,False,'Consolas')
  if text.startswith('【插图位置】'): p.paragraph_format.keep_with_next=True
for t in d.tables:
 total=sum(c.width for c in t.columns)
 ratio=(r.sections[-1].page_width-r.sections[-1].left_margin-r.sections[-1].right_margin)/total
 for c in t.columns: c.width=int(c.width*ratio)
 for row in t.rows:
  for cell in row.cells: cell.width=int(cell.width*ratio)
 for ri,row in enumerate(t.rows):
  for cell in row.cells:
   for p in cell.paragraphs:
    props(p,r.paragraphs[22]); font(p,12,ri==0)
    p.paragraph_format.first_line_indent=Pt(0)
    ind=p._p.pPr.find(qn('w:ind'))
    if ind is not None: ind.attrib.pop(qn('w:firstLineChars'),None)
    p.paragraph_format.space_before=Pt(3); p.paragraph_format.space_after=Pt(3)
    spacing=p._p.pPr.find(qn('w:spacing'))
    for a in ['beforeLines','afterLines']: spacing.attrib.pop(qn('w:'+a),None)
for sec in d.sections:
 rs=r.sections[-1]
 for attr in ['page_width','page_height','top_margin','bottom_margin','left_margin','right_margin','header_distance','footer_distance']: setattr(sec,attr,getattr(rs,attr))
 for e in list(sec._sectPr.findall(qn('w:docGrid'))): sec._sectPr.remove(e)
 sec._sectPr.append(deepcopy(rs._sectPr.find(qn('w:docGrid'))))
out=Path('output/智能终端系统及应用开发课程设计报告_李思民_参考格式版.docx'); d.save(out)
qa=Path('output/style_qa'); qa.mkdir(exist_ok=True)
(qa/'artifact.md').write_text('Reference: '+str(ref)+'\nSHA256: '+hashlib.sha256(ref.read_bytes()).hexdigest()+'\nBody: Songti 12pt; firstLine 480 twips, firstLineChars 200; line 300 auto; before/after 157 twips and 50 hundredths of a line.\nMain headings: Songti 14pt bold, source paragraph 19. Subheadings: Songti 12pt bold, paragraph 27. Captions: source paragraph 329, Songti 12pt centered.\nCover: Songti 18pt institution; Songti 26pt title; Microsoft YaHei 16pt metadata.\nBody page system: copied final reference section geometry and grid. Existing content, tables and image instructions retained. Cover indents adapted to body margins for fit.\n',encoding='utf-8')
assert [p.text for p in d.paragraphs]==[p.text for p in Document(src).paragraphs]
print(out.resolve())

