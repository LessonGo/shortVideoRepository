from pathlib import Path
from copy import deepcopy
import re
from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement
from docx.shared import Pt,RGBColor

src=Path('D:/qq.text/智能终端系统及应用开发课程设计报告_李思民_图文优化版.docx')
d=Document(src)
ps=d.paragraphs
body_sample=ps[61]
caption_sample=ps[60]
def before(element,text,caption=False):
 e=OxmlElement('w:p');element.addprevious(e);p=Paragraph(e,d)
 sample=caption_sample if caption else body_sample
 if sample._p.pPr is not None:p._p.insert(0,deepcopy(sample._p.pPr))
 r=p.add_run(text)
 if sample.runs and sample.runs[0]._r.rPr is not None:r._r.insert(0,deepcopy(sample.runs[0]._r.rPr))
 if caption:
  p.alignment=1;p.paragraph_format.first_line_indent=Pt(0);p.paragraph_format.keep_with_next=True
  p.paragraph_format.space_before=Pt(6);p.paragraph_format.space_after=Pt(6)
 else:p.paragraph_format.keep_with_next=True
 return p
fig_references={
 '3-1':'系统总体架构如图 3-1 所示。',
 '4-1':'首页与发现界面的运行效果如图 4-1 所示。',
 '4-2':'分页请求与列表更新流程如图 4-2 所示。',
 '4-3':'视频播放与浏览记录的数据流程如图 4-3 所示。',
 '4-4':'视频播放页的简介与评论界面如图 4-4 所示。',
 '4-5':'个人中心与系统相册选择头像的界面如图 4-5 所示。',
 '4-6':'广场的轮播与双列图片布局如图 4-6 所示。',
 '4-7':'收藏列表与浏览记录界面如图 4-7 所示。'}
figs=[]
for i,p in enumerate(ps):
 m=re.match(r'^图\s*(\d+-\d+)',p.text.strip())
 if not m or len(p.text)>100:continue
 num=m.group(1)
 image=p._p.getprevious()
 while image is not None and not image.xpath('.//w:drawing'):
  if ''.join(image.itertext()).strip():break
  image=image.getprevious()
 assert image is not None and image.xpath('.//w:drawing'),num
 prior=image.getprevious();context=''
 # Only accept a citation in the immediately preceding prose paragraph.
 if prior is not None:
  context=''.join(prior.xpath('.//w:t/text()'))
 if not re.search(r'图\s*'+re.escape(num)+r'(?!\d)',context):before(image,fig_references[num])
 Paragraph(image,d).paragraph_format.keep_with_next=True
 p.alignment=1;p.paragraph_format.first_line_indent=Pt(0);p.paragraph_format.keep_together=True
 figs.append(num)
tables=[
 ('3-1','系统模块划分与主要职责','系统各模块的划分及主要职责见表 3-1。'),
 ('4-1','网络请求状态与界面处理','不同请求状态下的界面处理与核对要点见表 4-1。'),
 ('4-2','播放器生命周期处理与验证重点','播放器在各生命周期入口中的处理与验证重点见表 4-2。'),
 ('4-3','浏览记录数据字段','浏览记录的主要字段、类型及作用见表 4-3。'),
 ('5-1','系统功能测试用例与记录状态','系统功能测试步骤、预期结果及当前记录状态见表 5-1。'),
 ('5-2','现有问题与改进措施','当前实现中需要关注的问题及相应改进措施见表 5-2。')]
assert len(d.tables)==len(tables)
for t,(num,title,cite) in zip(d.tables,tables):
 before(t._tbl,cite)
 before(t._tbl,'表 '+num+'　'+title,True)
out=Path('output/revision/智能终端系统及应用开发课程设计报告_李思民_图表引用规范版.docx');d.save(out)
# Structural checks ensure every caption is adjacent to its figure/table.
v=Document(out)
for t,(num,title,cite) in zip(v.tables,tables):
 e=t._tbl.getprevious();assert ''.join(e.xpath('.//w:t/text()')).startswith('表 '+num)
 assert num in ''.join(e.getprevious().xpath('.//w:t/text()'))
assert len(figs)==8
assert len(v.inline_shapes)==len(d.inline_shapes)
print(out.resolve());print('Figure groups:',len(figs),'Tables:',len(tables))

