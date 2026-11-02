from pathlib import Path
from docx import Document
from docx.shared import Pt,Cm
from docx.oxml.ns import qn
import re
src=Path('C:/Users/LISIMING/Pictures/Screenshots/智能终端系统及应用开发课程设计报告_李思民_图表引用规范版.docx')
d=Document(src)
for i,p in enumerate(d.paragraphs):
 if i<37:continue
 f=p.paragraph_format;t=p.text.strip();img=bool(p._p.xpath('.//w:drawing'))
 if not t and not img and not p._p.xpath('.//w:sectPr'):
  p._p.getparent().remove(p._p);continue
 f.page_break_before=(i==37)
 f.widow_control=True
 heading=bool(re.match(r'^(?:[一二三四五六]\s|\d+\.\d+(?:\.\d+)?\s)',t))
 if heading:
  f.keep_with_next=True;f.keep_together=True
  f.space_before=Pt(10);f.space_after=Pt(6)
 elif img:
  f.keep_with_next=True;f.space_before=Pt(4);f.space_after=Pt(2)
 elif re.match(r'^图\s*\d',t):
  f.keep_with_next=False;f.keep_together=True;f.space_before=Pt(3);f.space_after=Pt(6)
 elif re.match(r'^表\s*\d',t):
  f.keep_with_next=True;f.keep_together=True;f.space_before=Pt(4);f.space_after=Pt(4)
 elif f.space_before!=0 or f.space_after!=0:
  f.space_before=Pt(3);f.space_after=Pt(3)
 # Remove line-based spacing overrides so point spacing takes effect.
 sp=p._p.pPr.find(qn('w:spacing')) if p._p.pPr is not None else None
 if sp is not None:
  for a in ['beforeLines','afterLines','beforeAutospacing','afterAutospacing']:
   sp.attrib.pop(qn('w:'+a),None)
for t in d.tables:
 for ri,row in enumerate(t.rows):
  for cell in row.cells:
   for p in cell.paragraphs:
    p.paragraph_format.keep_with_next=(ri==0)
    p.paragraph_format.keep_together=True
# Place the first explanation next to the lead-in before each figure. This
# lets prose fill the preceding page while keeping the figure and caption paired.
for cap in list(d.paragraphs):
 if not re.match(r'^图\s*\d',cap.text.strip()):continue
 im=cap._p.getprevious()
 if im is None or not im.xpath('.//w:drawing'):continue
 nxt=cap._p.getnext()
 if nxt is None or nxt.tag!=qn('w:p'):continue
 from docx.text.paragraph import Paragraph
 para=Paragraph(nxt,d)
 if not para.text.strip() or re.match(r'^(?:[一二三四五六]\s|\d+\.\d+|表\s*\d|代码)',para.text.strip()):continue
 im.addprevious(nxt)
 para.paragraph_format.keep_with_next=False
 lead=nxt.getprevious()
 if lead is not None and lead.tag==qn('w:p'):
  Paragraph(lead,d).paragraph_format.keep_with_next=False
for s in d.inline_shapes:
 # Phone screenshots remain legible, at a consistent height.
 if s.width/s.height<0.6:
  ratio=s.width/s.height;s.height=Cm(10.5);s.width=int(s.height*ratio)
compact=False
for p in d.paragraphs:
 if p.text.startswith('5.2 运行现象'):compact=True
 if compact and not re.match(r'^(?:[一二三四五六]\s|\d+\.\d+|代码\s|表\s)',p.text):
  p.paragraph_format.space_before=Pt(0);p.paragraph_format.space_after=Pt(0)
out=Path('output/layout');out.mkdir(exist_ok=True)
d.save(out/'智能终端系统及应用开发课程设计报告_李思民_排版均衡版.docx')
print('saved')
