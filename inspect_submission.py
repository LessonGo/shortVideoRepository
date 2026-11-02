from pathlib import Path
import pypdfium2 as pdf
from PIL import Image,ImageDraw
from docx import Document
p=Path('output/submission'); d=pdf.PdfDocument(p/'check.pdf');print('pages',len(d))
for k in range(0,len(d),6):
 sheet=Image.new('RGB',(1500,1500),'#bbb');dr=ImageDraw.Draw(sheet)
 for j in range(k,min(k+6,len(d))):
  im=d[j].render(scale=1.5).to_pil();im.save(p/f'page-{j+1}.png');im.thumbnail((490,710));x=(j-k)%3*500;y=(j-k)//3*750;sheet.paste(im,(x,y+25));dr.text((x+10,y+4),str(j+1),fill='black')
 sheet.save(p/f'contact-{k//6}.png')
text='\n'.join(x.text for x in Document(p/'智能终端系统及应用开发课程设计报告_李思民_提交修订版.docx').paragraphs)
for term in ['不能证明','不能替代','仍需验证','尚不能','待验证','待实测','不等于','尚未替换','需核对']:
 print(term,text.count(term))
