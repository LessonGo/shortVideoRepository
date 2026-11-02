from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement
from docx.shared import Pt,Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
src=Path('output/submission/智能终端系统及应用开发课程设计报告_李思民_提交修订版.docx');d=Document(src)
anchor=next(p for p in d.paragraphs if p.text=='5.3 问题分析与改进方向')._p
body=next(p for p in d.paragraphs if p.text.startswith('图 4-1 展示首页卡片'))
cap=next(p for p in d.paragraphs if p.text.startswith('图 4-7'))
heading=next(p for p in d.paragraphs if p.text=='5.3.1 收藏数量更新的代码分析')
def add(text='',sample=body):
 e=OxmlElement('w:p');anchor.addprevious(e);p=Paragraph(e,d)
 if sample._p.pPr is not None:e.append(deepcopy(sample._p.pPr))
 run=p.add_run(text)
 if sample.runs and sample.runs[0]._r.rPr is not None:run._r.insert(0,deepcopy(sample.runs[0]._r.rPr))
 p.paragraph_format.page_break_before=False
 return p
root=Path('C:/Users/LISIMING/Pictures/Screenshots')
groups=[
('5.2.1 登录前后界面对比','登录前后的界面如图 5-1 所示。登录页提供账号与验证码输入入口，登录后的设置页展示账号绑定信息和退出登录入口。','105953','105914','图 5-1　登录前后界面对比\n（a）登录前的验证码登录页　　（b）登录后的账号设置页','登录前，用户在登录页输入账号并获取验证码；登录后，设置页显示脱敏账号及“退出登录”按钮，账号管理入口随登录状态呈现相应内容。'),
('5.2.2 点赞与收藏操作对比','同一视频在点赞与收藏操作前后的界面如图 5-2 所示，视频标题和播放位置保持一致。','105641','105702','图 5-2　点赞与收藏操作前后对比\n（a）操作前　　（b）操作后','操作前，点赞和收藏图标为空心；操作后，两者变为实心选中状态，点赞数量由 1 增加到 2。界面通过图标状态与点赞数值反馈用户的互动操作，视频内容和播放位置保持不变。'),
('5.2.3 浏览记录删除前后对比','浏览记录的选择删除过程如图 5-3 所示。用户进入编辑状态，选中目标记录后完成删除。','105819','105829','图 5-3　浏览记录删除前后对比\n（a）删除前选中目标记录　　（b）删除后的记录列表','删除前，“刺客信条：信仰之跃实验”位于列表首项并处于选中状态；删除后，该记录从列表移除，“我为什么要健身？”成为首项，其余记录依次前移，页面返回普通浏览状态。')]
for title,intro,a,b,label,desc in groups:
 p=add(title,heading);p.paragraph_format.keep_with_next=True
 p=add(intro);p.paragraph_format.keep_with_next=True;p.paragraph_format.space_after=Pt(4)
 p=add();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.first_line_indent=Pt(0)
 ind=p._p.pPr.find(qn('w:ind'))
 if ind is not None:ind.attrib.pop(qn('w:firstLineChars'),None)
 p.paragraph_format.line_spacing=1;p.paragraph_format.keep_with_next=True
 for j,s in enumerate([a,b]):
  if j:p.add_run('    ')
  p.add_run().add_picture(str(root/f'屏幕截图 2026-11-02 {s}.png'),height=Cm(9))
 p=add(label,cap);p.paragraph_format.keep_with_next=False;p.paragraph_format.keep_together=True
 p=add(desc);p.paragraph_format.space_after=Pt(6)
out=Path('output/submission/智能终端系统及应用开发课程设计报告_李思民_补充截图提交版.docx');d.save(out)
print('Inserted 6 images, 3 captions and 3 explanations')
