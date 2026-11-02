import pypdfium2 as pdf
from PIL import Image,ImageDraw
from pathlib import Path
p=Path('output/layout');d=pdf.PdfDocument(p/'layout_check.pdf'); print('pages',len(d))
for k in range(0,len(d),6):
 sheet=Image.new('RGB',(1200,1700),'#bbb');dr=ImageDraw.Draw(sheet)
 for j in range(k,min(k+6,len(d))):
  im=d[j].render(scale=1.5).to_pil();im.save(p/f'page-{j+1}.png');im.thumbnail((390,800));x=((j-k)%3)*400;y=((j-k)//3)*850;sheet.paste(im,(x,y+25));dr.text((x+10,y+4),str(j+1),fill='black')
 sheet.save(p/f'contact-{k//6}.png')
for i,page in enumerate(d):
 t=page.get_textpage().get_text_range();print(i+1,t[:60].replace('\r\n',' '),t[-65:].replace('\r\n',' '))
