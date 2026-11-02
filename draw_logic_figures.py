from pathlib import Path
from PIL import Image,ImageDraw,ImageFont
from html import escape
import math,zipfile
OUT=Path('output/系统逻辑架构优化图');OUT.mkdir(parents=True,exist_ok=True)
class Canvas:
 def __init__(self,w,h):
  self.w=w;self.h=h;self.s=2;self.im=Image.new('RGB',(w*2,h*2),'white');self.d=ImageDraw.Draw(self.im);self.svg=[f'<svg xmlns="http://www.w3.org/2000/svg" width="{w}" height="{h}" viewBox="0 0 {w} {h}"><rect width="100%" height="100%" fill="white"/>']
 def text(self,x,y,t,size=30,bold=False,bg=False):
  lines=t.split('\n');font=ImageFont.truetype('C:/Windows/Fonts/'+('simhei.ttf' if bold else 'simsun.ttc'),size*2)
  for i,line in enumerate(lines):
   yy=y+(i-(len(lines)-1)/2)*size*1.3
   bb=self.d.textbbox((0,0),line,font=font);ww=bb[2]-bb[0];hh=bb[3]-bb[1]
   if bg:
    self.d.rectangle((x*2-ww/2-8,yy*2-hh/2-5,x*2+ww/2+8,yy*2+hh/2+5),fill='white')
    self.svg.append(f'<rect x="{x-ww/4-4}" y="{yy-hh/4-3}" width="{ww/2+8}" height="{hh/2+6}" fill="white"/>')
   self.d.text((x*2-ww/2,yy*2-hh/2-bb[1]),line,font=font,fill='black')
   self.svg.append(f'<text x="{x}" y="{yy}" font-family="{ "SimHei" if bold else "SimSun"},serif" font-size="{size}" text-anchor="middle" dominant-baseline="central">{escape(line)}</text>')
 def rect(self,x,y,w,h,t='',size=30,fill='white',dash=False):
  self.d.rectangle((x*2,y*2,(x+w)*2,(y+h)*2),fill=fill,outline='#222',width=3)
  self.svg.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" fill="{fill}" stroke="#222" stroke-width="1.5"'+(' stroke-dasharray="10 7"' if dash else '')+'/>')
  if t:self.text(x+w/2,y+h/2,t,size)
 def diamond(self,cx,cy,w,h,t,size=28):
  pts=[(cx,cy-h/2),(cx+w/2,cy),(cx,cy+h/2),(cx-w/2,cy)]
  self.d.polygon([(x*2,y*2) for x,y in pts],fill='white');self.line(pts+[pts[0]],arrow=False);self.svg.append(f'<polygon points="'+ ' '.join(f'{x},{y}' for x,y in pts)+'" fill="white" stroke="#222" stroke-width="1.5"/>');self.text(cx,cy,t,size)
 def line(self,pts,arrow=True,dash=False):
  for a,b in zip(pts,pts[1:]):
   if dash:
    length=math.dist(a,b);n=max(1,int(length/10))
    for i in range(n):
     if i%2==0:self.d.line([(int((a[0]+(b[0]-a[0])*i/n)*2),int((a[1]+(b[1]-a[1])*i/n)*2)),(int((a[0]+(b[0]-a[0])*(i+1)/n)*2),int((a[1]+(b[1]-a[1])*(i+1)/n)*2))],fill='#222',width=3)
   else:self.d.line([(int(a[0]*2),int(a[1]*2)),(int(b[0]*2),int(b[1]*2))],fill='#222',width=3)
  self.svg.append('<polyline points="'+' '.join(f'{x},{y}' for x,y in pts)+'" fill="none" stroke="#222" stroke-width="1.5"'+(' stroke-dasharray="10 7"' if dash else '')+'/>')
  if arrow:
   a,b=pts[-2:];ang=math.atan2(b[1]-a[1],b[0]-a[0]);p=[b,(b[0]-15*math.cos(ang-.42),b[1]-15*math.sin(ang-.42)),(b[0]-15*math.cos(ang+.42),b[1]-15*math.sin(ang+.42))]
   self.d.polygon([(int(x*2),int(y*2)) for x,y in p],fill='#222');self.svg.append('<polygon points="'+' '.join(f'{x},{y}' for x,y in p)+'" fill="#222"/>')
 def save(self,name):
  self.im.save(OUT/(name+'.png'),dpi=(300,300));self.svg.append('</svg>');(OUT/(name+'.svg')).write_text('\n'.join(self.svg),encoding='utf-8')

# Figure 1: logical layers; separate downward operations and upward data.
c=Canvas(1800,1500)
c.text(900,42,'系统逻辑架构与数据交互',40,True)
for x,y,w,h,t in [
 (480,95,840,100,'app　MainActivity\n底部导航与页面容器'),
 (480,285,840,135,'界面层　Activity / Fragment\n首页　广场　发现　个人中心　播放页'),
 (480,525,840,115,'状态层　ViewModel\n业务状态、LiveData 与界面通知'),
 (480,745,840,115,'数据访问层　各业务 Model\n组装参数、调用接口与本地记录操作'),
 (480,985,840,115,'网络访问层　ApiService / ApiCall\nRetrofit + OkHttp + Gson'),
 (480,1250,840,100,'远程接口服务\n视频列表与详情、用户资料、评论、点赞和收藏')]:c.rect(x,y,w,h,t,31)
c.line([(900,195),(900,285)]);c.text(900,241,'ARouter 路由与参数',26,bg=True)
for ya,yb,down,up in [(420,525,'用户操作','LiveData / DataBinding'),(640,745,'请求业务数据','回调 / 状态结果'),(860,985,'参数与令牌','业务数据对象'),(1100,1250,'HTTP 请求','JSON 响应')]:
 c.line([(660,ya),(660,yb)]);c.text(660,(ya+yb)/2,down,26,bg=True)
 c.line([(1140,yb),(1140,ya)]);c.text(1140,(ya+yb)/2,up,26,bg=True)
c.rect(40,285,340,355,'公共支撑\nlibrary_base\n\nBaseActivity / Fragment\nARouter 路由\nEventBus 登录事件',27,fill='#f7f7f7')
c.line([(380,350),(480,350)],dash=True);c.line([(380,580),(480,580)],dash=True)
c.rect(40,985,340,115,'数据模型　video_data\n响应对象与业务实体',27,fill='#f7f7f7');c.line([(380,1040),(480,1040)],dash=True)
c.rect(1435,670,325,155,'UserManager\n加密偏好存储\ntoken 与用户资料',27)
c.line([(1320,790),(1435,790)]);c.text(1380,811,'读写',24,bg=True)
c.line([(1435,700),(1370,700),(1370,755),(1320,755)],dash=True)
c.rect(1435,945,325,155,'Room 本地数据库\n浏览记录\n查询、插入、更新、删除',27)
c.line([(1320,820),(1395,820),(1395,975),(1435,975)]);c.text(1460,872,'记录操作',25,bg=True)
c.line([(1435,1060),(1355,1060),(1355,835),(1320,835)],dash=True);c.text(1600,1145,'查询结果由回调返回界面',25)
c.text(900,1430,'实线：操作或请求及网络响应　　虚线：公共支撑、模型映射或本地结果返回',27)
c.save('图1_系统逻辑架构与数据流')

# Figure 2: actual Java home-list implementation, not comment paging.
c=Canvas(2200,1900);c.text(1100,40,'首页推荐与日报的分页处理流程',40,True)
c.rect(130,100,550,110,'首次进入 / 下拉刷新\nisFirst = true，mPage = 1',29)
c.rect(950,100,550,110,'上拉加载更多\nisFirst = false，mPage++',29)
c.line([(405,210),(405,245),(800,245),(800,285)])
c.line([(1225,210),(1225,245),(800,245)],arrow=False)
c.rect(470,285,660,130,'VideoListModel.requestDatas\n按列表类型调用 getRecommend / getDaily\n参数：page，limit = 10',29)
c.line([(800,415),(800,475)])
c.diamond(800,555,520,160,'HTTP 与业务响应正常？\ncode = 1',29)
c.line([(800,635),(800,700)]);c.text(830,670,'是',26,bg=True)
c.line([(1060,555),(1925,555),(1925,700)]);c.text(1195,530,'否 / 网络失败',26,bg=True)
c.diamond(800,785,520,170,'返回列表是否非空？',29)
c.rect(1675,700,500,160,'错误或空数据回调\nGetFailed(errorCode)\n更新 mErrorCode，显示状态页',28)
c.line([(1060,785),(1675,785)]);c.text(1130,760,'否',26,bg=True)
c.line([(1925,860),(1925,950)])
c.rect(1705,950,440,90,'本次请求失败 / 无内容',29)
c.line([(800,870),(800,930)]);c.text(830,900,'是',26,bg=True)
c.diamond(800,1010,500,160,'isFirst 是否为 true？',29)
c.line([(550,1010),(350,1010),(350,1150)]);c.text(445,983,'是',26,bg=True)
c.line([(1050,1010),(1250,1010),(1250,1150)]);c.text(1150,983,'否',26,bg=True)
c.rect(100,1150,500,110,'刷新 / 首次加载\nmData = 本页结果',29)
c.rect(1000,1150,500,110,'加载更多\n原列表 addAll(本页列表)',29)
c.line([(350,1260),(350,1310),(800,1310),(800,1355)])
c.line([(1250,1260),(1250,1310),(800,1310)],arrow=False)
c.diamond(800,1445,650,180,'已加载条数 ≥ 服务端 count？',29)
c.line([(475,1445),(320,1445),(320,1620)]);c.text(395,1417,'是',26,bg=True)
c.line([(1125,1445),(1280,1445),(1280,1620)]);c.text(1205,1417,'否',26,bg=True)
c.rect(90,1620,460,100,'enableLoadMore = false\n禁止继续加载',28)
c.rect(1050,1620,460,100,'保留加载更多能力\n等待下一次上拉',28)
c.line([(320,1720),(320,1760),(800,1760),(800,1790)])
c.line([(1280,1720),(1280,1760),(800,1760)],arrow=False)
c.rect(390,1790,820,70,'LiveData 通知列表刷新，并结束成功请求的加载动画',27)
c.save('图2_首页分页请求与状态分支')

# Figure 3: detail success forks to playback and local history; controls independent.
c=Canvas(2000,1890);c.text(950,40,'视频详情、播放与浏览记录处理流程',40,True)
c.rect(600,95,700,100,'点击视频卡片 → ARouter 跳转\n播放页接收 video_id',30)
c.line([(950,195),(950,255)])
c.rect(600,255,700,110,'播放页调用 ViewModel.requestVideoInfo\nviewModelScope 发起详情请求',29)
c.line([(950,365),(950,425)])
c.rect(600,425,700,110,'Model → ApiService → 远程详情接口\n携带 id 与可空 token',29)
c.line([(950,535),(950,590)])
c.diamond(950,675,550,170,'返回 Success 且有详情？',29)
c.line([(1225,675),(1540,675)]);c.text(1380,646,'Empty / Error',25,bg=True)
c.rect(1540,610,310,130,'提示无数据或错误\n结束本次加载',27)
c.line([(950,760),(950,815)]);c.text(980,790,'是',26,bg=True)
c.rect(600,815,700,115,'ViewModel.loadData(ArchivesInfo)\n更新视频详情及互动状态的 LiveData',29)
c.line([(950,930),(950,970),(430,970),(430,1020)])
c.line([(950,970),(1430,970),(1430,1020)])
c.rect(100,1020,660,115,'Activity 的详情观察者收到更新\n读取 video_file，调用播放器 play(url)',29)
c.rect(1110,1020,640,115,'ViewModel 调用 Model.InsertData\n按当前用户与视频构造浏览记录',29)
c.line([(430,1135),(430,1190)])
c.rect(100,1190,660,125,'ExoPlayer 设置 MediaItem\nprepare() → play()\n同时触发评论首屏请求',29)
c.line([(1430,1135),(1430,1190)])
c.rect(1110,1190,640,95,'工作线程查询 Room\n条件：user_id 与 video_id',28)
c.line([(1430,1285),(1430,1320)])
c.diamond(1430,1385,460,130,'记录是否已存在？',28)
c.line([(1200,1385),(1080,1385),(1080,1490)]);c.text(1140,1358,'是',25,bg=True)
c.line([(1660,1385),(1780,1385),(1780,1490)]);c.text(1720,1358,'否',25,bg=True)
c.rect(900,1490,360,85,'仅更新 view_time',28)
c.rect(1600,1490,360,85,'插入新记录',28)
c.text(430,1390,'播放与历史写入为两条分支\n记录详情访问，不代表完整观看',27)
c.rect(70,1640,1780,195,fill='#fafafa')
c.text(960,1667,'生命周期控制　各事件独立触发，不表示固定先后顺序',28,True)
for x,w,t in [(95,360,'onStart\n允许播放'),(500,395,'onPause / onStop\n暂停播放器'),(940,385,'onDestroy\n停止并释放播放器'),(1370,450,'onNewIntent\n更新 id，重新请求详情')]:c.rect(x,1705,w,100,t,28)
c.save('图3_视频播放与本地记录流程')

(OUT/'图示说明.txt').write_text('三图根据当前项目代码重绘，区分请求、结果、判断分支及独立生命周期事件。\n\n图1：逻辑分层及主要数据关系，不等同于 Gradle 模块依赖图。公共基础与数据模型是支撑关系，不是独立请求处理阶段。\n图2：对应 VideoListModel、BaseListViewmodel、BaseListFragment、ApiCall。首页按已加载数量与 count 比较，不是按本页不足十条判断。错误分支目前仅更新状态页；失败页码回退、空响应判空及异常时结束加载动画仍需完善，图中未将其标为已实现。\n图3：对应 Kotlin 播放链路与 UserLookRecordOperation。游客浏览记录 user_id 为 0；登录后使用实际用户编号。查询后更新或插入在工作线程执行。onPause 后只回到 onResume 的恢复行为仍需验证。\n\nPNG 为高清插图，SVG 为可编辑矢量图。建议在报告中分别替换图3-1、图4-2、图4-3；保留图前引用及图下题注。\n',encoding='utf-8')
with zipfile.ZipFile(OUT/'系统逻辑架构优化图.zip','w',zipfile.ZIP_DEFLATED) as z:
 for f in OUT.iterdir():
  if f.suffix in ['.svg','.png','.txt']:z.write(f,f.name)
print(OUT.resolve())
