from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.text.paragraph import Paragraph
src=Path('output/submission/智能终端系统及应用开发课程设计报告_李思民_补充截图提交版.docx')
d=Document(src)
a=next(p for p in d.paragraphs if p.text.startswith('客户端采用 Java'))
e=OxmlElement('w:p');a._p.addnext(e);p=Paragraph(e,d)
if a._p.pPr is not None:e.append(deepcopy(a._p.pPr))
p.paragraph_format.keep_with_next=False
r=p.add_run('本项目源代码已上传至 GitHub，项目仓库地址：')
if a.runs[0]._r.rPr is not None:r._r.insert(0,deepcopy(a.runs[0]._r.rPr))
url='https://github.com/litianyan52/shortVideoRepository'
h=OxmlElement('w:hyperlink');h.set(qn('r:id'),d.part.relate_to(url,RT.HYPERLINK,is_external=True))
r=OxmlElement('w:r');pr=OxmlElement('w:rPr');ft=OxmlElement('w:rFonts');ft.set(qn('w:ascii'),'Times New Roman');ft.set(qn('w:hAnsi'),'Times New Roman');pr.append(ft)
sz=OxmlElement('w:sz');sz.set(qn('w:val'),'24');pr.append(sz)
c=OxmlElement('w:color');c.set(qn('w:val'),'0563C1');pr.append(c)
u=OxmlElement('w:u');u.set(qn('w:val'),'single');pr.append(u);r.append(pr);t=OxmlElement('w:t');t.text=url;r.append(t);h.append(r);e.append(h)
out=src.with_name('智能终端系统及应用开发课程设计报告_李思民_含GitHub地址版.docx');d.save(out)
print('Saved with external hyperlink')
