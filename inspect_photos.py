from pathlib import Path
import pypdfium2 as pdf
from PIL import Image,ImageDraw
p=Path('output/submission/photo_qa');p.mkdir(exist_ok=True);d=pdf.PdfDocument('output/submission/photos_check.pdf');print('pages',len(d))
for i,pg in enumerate(d):
 im=pg.render(scale=1.5).to_pil();im.save(p/f'page-{i+1}.png')
 if i>=15:print(i+1,pg.get_textpage().get_text_range()[:65])
for k in range(0,len(d),6):
 out=Image.new('RGB',(1500,1460),'#bbb');dr=ImageDraw.Draw(out)
 for j in range(k,min(k+6,len(d))):
  im=Image.open(p/f'page-{j+1}.png');im.thumbnail((490,690));x=(j-k)%3*500;y=(j-k)//3*730;out.paste(im,(x,y+25));dr.text((x+10,y+3),str(j+1),fill='black')
 out.save(p/f'contact-{k//6}.png')
