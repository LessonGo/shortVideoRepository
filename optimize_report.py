from pathlib import Path
from copy import deepcopy
import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image, ImageDraw, ImageFont

root=Path('output/revision'); assets=root/'assets'; root.mkdir(exist_ok=True)
src=Document('D:/qq.text/智能终端系统及应用开发课程设计报告_李思民_参考格式版.docx')
d=Document('D:/qq.text/智能终端系统及应用开发课程设计报告_李思民_参考格式版.docx')
body=d._element.body
for e in list(body):
 if e.tag!=qn('w:sectPr'): body.remove(e)
for pp in src.paragraphs[:9]: body.insert(len(body)-1,deepcopy(pp._p))
for name in ['Normal','Heading 1','Heading 2','Heading 3','Caption','Title']:
 s=d.styles[name]
 for tag in ['pPr','rPr']:
  x=s.element.find(qn('w:'+tag))
  if x is not None:s.element.remove(x)
 s.font.name='宋体'; s.font.size=Pt(12); s.font.color.rgb=RGBColor(0,0,0)
 s.element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'宋体')
 s.paragraph_format.line_spacing=1.25; s.paragraph_format.space_before=Pt(6); s.paragraph_format.space_after=Pt(6)
 s.paragraph_format.widow_control=True
for name,size in [('Heading 1',14),('Heading 2',12),('Heading 3',12)]:
 s=d.styles[name];s.font.size=Pt(size);s.font.bold=True;s.paragraph_format.keep_with_next=True;s.paragraph_format.first_line_indent=Pt(0)
d.styles['Normal'].paragraph_format.first_line_indent=Pt(24)
d.styles['Caption'].paragraph_format.first_line_indent=Pt(0)
for sec in d.sections:
 sec.top_margin=Cm(2.54);sec.bottom_margin=Cm(2.54);sec.left_margin=Cm(3.175);sec.right_margin=Cm(3.175)
 for x in list(sec._sectPr.findall(qn('w:docGrid'))):sec._sectPr.remove(x)
grid=None
def p(t):return d.add_paragraph(t)
def old(i):return p(src.paragraphs[i].text)
counter=0; toc=[]
def h(t,l=1,new=False):
 global counter
 pp=d.add_paragraph(t,'Heading '+str(l));pp.paragraph_format.page_break_before=new
 counter+=1; key='sec'+str(counter)
 a=OxmlElement('w:bookmarkStart');a.set(qn('w:id'),str(counter));a.set(qn('w:name'),key)
 b=OxmlElement('w:bookmarkEnd');b.set(qn('w:id'),str(counter));pp._p.insert(0,a);pp._p.append(b)
 if l<=2:toc.append((t,key,l))
 return pp
def cap(t):
 pp=d.add_paragraph(t,'Caption');pp.alignment=1;pp.paragraph_format.space_before=Pt(3);pp.paragraph_format.space_after=Pt(8)
def code(title,text,source):
 pp=p(title);pp.paragraph_format.first_line_indent=Pt(0);pp.paragraph_format.keep_with_next=True
 for line in text.strip().splitlines():
  pp=d.add_paragraph();f=pp.paragraph_format;f.first_line_indent=Pt(0);f.space_before=Pt(0);f.space_after=Pt(0);f.line_spacing=Pt(12);f.keep_with_next=True
  r=pp.add_run(line);r.font.name='Consolas';r.font.size=Pt(9)
 pp=p('代码来源：'+source+'（节选）。');pp.paragraph_format.first_line_indent=Pt(0)
 for r in pp.runs:r.font.size=Pt(9)
def table(head,rows,widths):
 t=d.add_table(rows=1,cols=len(head));t.autofit=False
 for c,w in zip(t.columns,widths):c.width=Cm(w)
 for c,v in zip(t.rows[0].cells,head):c.text=v
 for row in rows:
  for c,v in zip(t.add_row().cells,row):c.text=v
 for ri,row in enumerate(t.rows):
  row._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))
  for j,c in enumerate(row.cells):
   c.width=Cm(widths[j]); pr=c._tc.get_or_add_tcPr()
   borders=OxmlElement('w:tcBorders')
   for side in ['top','bottom','left','right']:
    e=OxmlElement('w:'+side);e.set(qn('w:val'),'single');e.set(qn('w:sz'),'4');e.set(qn('w:color'),'B8B8B8');borders.append(e)
   pr.append(borders)
   if ri==0:
    e=OxmlElement('w:shd');e.set(qn('w:fill'),'EFEFEF');pr.append(e)
   for pp in c.paragraphs:
    f=pp.paragraph_format;f.first_line_indent=Pt(0);f.space_before=Pt(3);f.space_after=Pt(3);f.line_spacing=1.15
    for r in pp.runs:r.font.size=Pt(10.5);r.bold=ri==0
 t.rows[0]._tr.get_or_add_trPr().append(OxmlElement('w:tblHeader'))
 p('')
 return t
def photos(files,labels,title):
 pp=d.add_paragraph();pp.alignment=1;pp.paragraph_format.first_line_indent=Pt(0);pp.paragraph_format.keep_with_next=True
 for i,file in enumerate(files):
  if i:pp.add_run('    ')
  pp.add_run().add_picture(str(assets/file),height=Cm(11.8))
 cap(title+'\n'+'　　'.join(labels))
def fig(file,title,width=14.5):
 pp=d.add_paragraph();pp.alignment=1;pp.paragraph_format.first_line_indent=Pt(0);pp.paragraph_format.keep_with_next=True
 pp.add_run().add_picture(str(assets/file),width=Cm(width));cap(title)
def diagram(name,rows):
 w=1600; rowh=170; im=Image.new('RGB',(w,len(rows)*rowh+40),'white');dr=ImageDraw.Draw(im)
 ft=ImageFont.truetype('C:/Windows/Fonts/simsun.ttc',34)
 for i,labels in enumerate(rows):
  bw=(w-100-(len(labels)-1)*30)//len(labels); yy=20+i*rowh
  for j,label in enumerate(labels):
   x=50+j*(bw+30);dr.rounded_rectangle((x,yy,x+bw,yy+115),radius=8,outline='#333333',width=3,fill='#F3F5F6')
   for k,line in enumerate(label.split('\n')):
    bb=dr.textbbox((0,0),line,font=ft);dr.text((x+(bw-(bb[2]-bb[0]))/2,yy+22+k*42),line,font=ft,fill='black')
  if i<len(rows)-1:
   x=w//2;dr.line((x,yy+120,x,yy+155),fill='black',width=3);dr.polygon([(x,yy+163),(x-9,yy+148),(x+9,yy+148)],fill='black')
 im.save(assets/name)
diagram('architecture.png', [['app 应用入口与底部导航'],['首页与媒体模块','广场与发现模块','用户与头像模块'],['library_base\n基类 路由 用户管理','network 与 video_data\n接口封装 数据模型'],['远程服务\n视频 用户 互动数据','本地存储\nRoom 与加密偏好设置']])
diagram('paging.png',[['首次进入或下拉刷新\n页码设为 1','上拉加载更多\n请求下一页'],['Model 调用推荐或日报接口'],['成功且有数据\n刷新替换 加载追加','空数据或请求失败\n提示并结束加载'],['返回不足十条时停止加载更多\n失败页码回退属于后续改进']])
diagram('playflow.png',[['列表点击视频卡片\n通过路由传入 video_id'],['播放页 ViewModel 请求详情\n返回 ArchivesInfo'],['Activity 观察视频数据\nExoPlayer 准备并播放','Model 写入 Room\n保存或更新浏览时间'],['切换视频走 onNewIntent\n暂停页面走 onPause','销毁页面走 onDestroy\n停止并释放播放器']])

# Contents placeholder populated after headings exist.
toc_head=d.add_paragraph('目录','Heading 1');toc_head.paragraph_format.page_break_before=True
toc_anchor=d.add_paragraph()

h('一　实验题目',new=True);old(29);old(30)
h('二　实验目的')
for i in range(32,38):old(i)
p('本设计以“内容获取、界面展示、媒体播放、用户互动和数据保存”为学习主线，在实现功能的同时关注模块边界、异步操作与状态一致性。')

h('三　总体设计',new=True);h('3.1 需求分析与功能边界',2);old(40)
p('核心使用场景分为三类：游客浏览内容并播放视频；登录用户进行收藏、点赞和评论；用户通过收藏与本地记录回到已关注的内容。用户操作应能够获得明确反馈，网络不可用时保留必要的返回与重试入口。')
h('3.2 技术方案与模块划分',2);old(42)
table(['模块','主要职责'],[[c.text for c in row.cells] for row in src.tables[0].rows][1:],[4.6,10.05])
old(44)

h('3.3 系统架构与数据流',2,new=True)
p('系统总体架构如图 3-1 所示。业务模块通过路由协作，公共层集中提供基础能力；远程接口与本地数据库承担不同的数据职责。图中表示职责关系，不表示所有业务模块之间存在直接编译依赖。')
fig('architecture.png','图 3-1　系统总体架构')
old(48);old(49)
p('视频列表对象 ResVideo 提供卡片展示字段，ArchivesInfo 提供播放地址、正文及互动状态，ResVideoAllInfo 聚合详情相关数据。列表项只传递编号，播放页再读取完整详情，可避免把大对象随页面跳转反复传递。')

h('四　详细设计',new=True);h('4.1 主界面导航与组件路由',2)
old(52);old(53)
code('代码 4-1　视频列表跳转播放页的调用方式', '''ARouter.getInstance()
    .build(ArouterPath.Video.ACTIVITY_MEDIA_PLAY)
    .withInt(ArouterPath.Video.KEY_VIDEO_ID, videoId)
    .navigation();''','VideoListFragment.java 的路由调用方式，参数变量作简化')
p('路径常量确定目标页面，KEY_VIDEO_ID 约定双方使用同一参数名称。页面收到编号后请求详情，路由因此连接的是页面入口，而不是代替业务接口完成数据查询。编号无效或详情为空时，还需要界面提示并停止播放初始化。')
old(58)
p('底部导航使用 add、show 和 hide 保留 Fragment。保留页面可以减少重复请求，但隐藏页面仍处于工程管理之中，不能把“不可见”等同于“已销毁”。媒体等持续占用资源的行为应由播放页生命周期单独管理。')

h('4.2 首页浏览与内容发现',2,new=True)
p('首页通过“推荐”和“日报”切换视频列表，发现页通过分类、主题和话题组织内容。两者分别服务于连续浏览与定向查找。图 4-1 展示了现有运行截图中的界面结构。')
photos(['image2.png','image1.png'],['（a）首页推荐列表','（b）发现分类与主题'],'图 4-1　首页与发现界面')
p('首页卡片包含封面、标题和作者信息。封面使用 Glide 异步加载，列表使用 RecyclerView 回收条目视图；点击卡片进入播放页。发现页的分类与主题入口将浏览范围缩小到相应内容集合。截图能够证明页面布局，分页与各入口的运行稳定性仍需按测试用例验证。')

h('4.3 网络请求与首页分页',2,new=True)
old(60);old(61)
fig('paging.png','图 4-2　分页请求与列表更新流程')
p('当前 VideoListModel 在请求前自增页码，失败回调仅向上层传递错误。若某页失败后再次加载仍继续自增，可能跳过失败页。因此，改进时应保存本次目标页，成功后再提交页码，或者失败时回退，同时限制同一列表的并行加载。')

h('4.3.1 请求状态的统一处理',3,new=True)
code('代码 4-2　统一封装对业务结果的判断', '''val response = call()
if (response.getCode() == successCode) {
    val data = dataType(response)
    when {
        data != null -> NetworkStatusResult.Success(data)
        else -> NetworkStatusResult.Empty(response.getMsg())
    }
} else {
    NetworkStatusResult.Error(response.getMsg())
}''','network/.../ApiCallKotlinVersion.kt')
p('封装通过 successCode 判断业务成功，再利用 dataType 提取具体数据，将界面接收的结果归纳为 Success、Empty 和 Error。这样，页面不必重复解析响应外壳；但 Empty 的具体含义仍由业务接口决定，例如无返回实体的操作与空列表并不完全相同。')
p('该封装对 CancellationException 原样抛出，使协程取消继续传播；其余异常转换为网络错误结果。ViewModel 使用 viewModelScope 发起请求，生命周期结束时可取消协程。请求完成不等于媒体资源自动释放，播放器仍需要独立管理。')
table(['场景','界面处理与核对点'],[('成功返回列表','刷新替换或加载追加；检查是否重复'),('空列表','区分首次无内容和已经到达末页'),('网络或业务失败','结束加载状态，保留已有内容并提示原因'),('登录令牌失效','清理或更新会话，引导重新登录')],[4.2,10.45])

h('4.4 视频播放与生命周期管理',2,new=True)
old(65);old(66)
code('代码 4-3　媒体地址装载与播放', '''public void play(String url) {
    if (mPlayer != null) {
        MediaItem mediaItem = MediaItem.fromUri(url);
        mPlayer.setMediaItem(mediaItem);
        mPlayer.prepare();
        mPlayer.play();
    }
}''','feature_mediaplay/.../player/MediaPlayerManager.java')
p('MediaItem 描述媒体资源，setMediaItem 更换播放对象，prepare 进入准备阶段，play 表达开始播放的意图。准备与缓冲可能尚未结束，因此调用 play 不能直接作为首帧已显示或播放无卡顿的测试结论。')
old(75);old(76)

h('4.4.1 播放链路与资源释放',3,new=True)
fig('playflow.png','图 4-3　视频播放与浏览记录的数据流程')
p('收到新视频编号后，应使播放器、标题、互动状态与评论指向同一视频。当前 onNewIntent 会重新请求详情并恢复评论加载标记；关联推荐列表也需要在切换后核对。详情请求成功后就写入浏览记录，因此记录表示“访问过详情”，并不证明视频已完整观看。')
table(['生命周期入口','当前处理','验证重点'],[('onStart','允许播放，注册事件','页面首次显示与后台返回'),('onPause / onStop','暂停播放','短暂遮挡与切后台'),('onDestroy','停止、清空媒体并释放','退出后无音频与资源残留')],[3.4,5.3,5.95])

h('4.4.2 视频简介与评论界面',3,new=True)
photos(['image4.png','image3.png'],['（a）视频简介及互动入口','（b）同一视频的评论列表'],'图 4-4　视频播放页的简介与评论')
p('两张截图中的播放位置与视频内容一致，下部区域分别显示简介和评论。简介提供标题、描述、互动入口及相关内容，评论区域显示评论列表和输入入口。页面切换以同一视频详情为上下文，避免用户从简介切到评论时改变播放对象。')
p('截图可用于说明界面组织及评论展示；音视频同步、连续播放、评论提交成功和后台恢复需要实际操作记录支持。画面中的分享图标仅说明入口外观，不能据此认定分享链路已完成。')

h('4.5 用户登录与互动状态',2,new=True)
old(80)
code('代码 4-4　视频详情接口允许游客令牌为空', '''@GET(NetAddress.VIDEO_INFO_ADDRESS)
suspend fun getVideoInfo(
    @Query("id") id: String,
    @Header("token") token: String?
): ResBase<ResVideoAllInfo>''','feature_mediaplay/.../apiKotlinVersion/MediaApiService.kt')
p('可空 token 允许游客请求公开详情；需要身份的互动操作仍在 ViewModel 中先判断登录。客户端校验用于引导操作，服务端仍应校验令牌有效性及评论等资源的操作权限。')
p('当前 Kotlin 点赞与取消点赞在请求返回 Success 后才更新 mIsLike、mLikes 和详情对象，不属于请求前直接改变数字的处理方式。失败时应保留原状态；连续点击仍需通过请求进行中标记避免重复提交。')
p('评论删除则先从本地列表移除，再发起删除请求，失败时可能出现列表与服务端不同步。该操作适合在成功后移除，或保存原位置并在失败时恢复。按昵称判断可删除评论只是一层界面限制，最终权限应由服务端依据用户身份决定。')

h('4.5.1 个人中心与头像选择',3,new=True)
photos(['image6.png','image9.png'],['（a）已登录个人中心','（b）系统相册选择头像'],'图 4-5　用户资料与头像选择界面')
p('个人中心展示昵称或脱敏账号信息，并提供收藏、浏览记录及退出入口。右图为系统照片选择器，并非 CameraX 的相机预览；它说明资料编辑能够进入相册选择流程，尚不能单凭这一画面确认图片上传及头像更新已成功。')
p('头像处理需分别验证拍照、相册选择、取消选择和上传失败。若使用系统照片选择器，授权范围应围绕选中的照片；相机入口则需要单独核对相机权限拒绝后的提示与返回行为。')

h('4.6 广场与内容分类',2,new=True)
p('广场将轮播内容与双列图片卡片组合展示，轮播占据完整行，其余卡片按双列排列。该设计适合突出精选内容，并保留进入分类或大图浏览的入口。')
photos(['image10.png'],['广场轮播与图片网格'],'图 4-6　广场界面')
p('适配器使用不同条目类型处理轮播与普通图片。由于第一个位置被轮播占用，普通图片的数据索引需要减去相应偏移；点击事件同样需要使用转换后的索引。图片浏览页用横向 ViewPager2 展示组图，并通过页码提示当前查看位置。')
p('发现页的分类和主题播单与广场共同构成多入口浏览。对空分类、话题数量不足以及图片加载失败应提供容错，避免直接访问固定下标导致界面异常。')

h('4.7 本地浏览记录与数据持久化',2,new=True)
old(89)
table(['字段','类型','作用'],[[c.text for c in row.cells] for row in src.tables[1].rows][1:],[5.0,2.8,6.85])
code('代码 4-5　按用户读取并倒序展示浏览记录', '''@Query("SELECT* FROM user_look_record " +
       "WHERE user_id = :user_id " +
       "ORDER BY view_time DESC ")
List<UserLookRecord> getRecords(String user_id);''','library_base/.../db/UserLookRecordDao.java；长 SQL 按行拆分展示')
p('user_id 用于隔离账号记录，view_time 用于把最近访问的内容排在前面。插入前通过用户与视频组合查询，已有记录只更新时间。当前先查后写属于应用层控制，并发请求下仍应借助组合唯一约束与事务保证一致性。')
p('数据库操作在工作线程执行，回调切回主线程后再更新列表。浏览记录保存的是封面、标题和视频编号等元数据，重新播放仍依赖媒体地址及网络连接。')

h('4.7.1 收藏与历史回看界面',3,new=True)
photos(['image8.png','image7.png'],['（a）远程收藏列表','（b）本地浏览记录'],'图 4-7　收藏与浏览记录')
p('收藏表达用户主动保留的内容，浏览记录记录最近访问内容，二者的数据来源和业务含义不同。前者通过接口读取，后者从本机 Room 查询。相同视频可以同时出现在两个列表，不能因显示重复就将两类数据合并。')
p('浏览记录页面的编辑入口用于选择并删除记录。验证时应在删除后退出并重新进入页面，确认数据库已更新；重复查看同一视频时，应检查记录是否只更新时间，以及账号切换后是否只读取对应用户的数据。')

h('五　实验结果与分析',new=True);h('5.1 功能验证与测试设计',2)
p('现有截图展示了首页、发现、播放、评论、个人中心、广场、收藏、浏览记录和相册选择等页面。静态代码核对可以说明功能调用路径，但不能替代设备运行测试。以下记录以“界面可观察现象”和“待验证行为”分别说明，不将全部功能直接标为通过。')
table(['测试项目','测试步骤与预期结果','记录状态'],[
('四入口切换','切换后返回首页，列表与位置保持','已有界面图；待操作核对'),('推荐与日报','刷新替换、加载追加，失败不跳页','待实测'),('游客播放','未登录打开视频，正常读取详情','待实测'),('视频切换','从相关列表选新视频，媒体和评论同步','待实测'),('生命周期','切后台、短暂遮挡及返回，暂停恢复正确','待实测'),('登录与退出','登录后资料刷新；退出后互动被拦截','已有个人页；待流程核对'),('点赞与收藏','新增及取消后，重进页面状态一致','待实测'),('评论操作','提交后可见，删除失败时状态可恢复','已有列表图；待提交核对'),('历史持久化','重复访问更新时间，删除后重进不恢复','已有记录图；待实测'),('网络与权限','断网或拒绝权限，提示并可返回重试','待实测')],[2.7,8.2,3.75])
p('实测环境：设备型号________；Android 版本________；应用版本________；网络条件________；执行日期________。结果记录应注明账号状态、具体步骤、实际现象和对应截图编号。')

h('5.2 运行现象与结果分析',2,new=True)
p('图 4-1 显示首页卡片和发现分类，图 4-4 显示同一视频的简介与评论，表明内容浏览与信息展示已经有相应界面。图 4-5 显示个人中心及系统相册入口，图 4-6 显示广场的轮播和双列布局，图 4-7 显示收藏与浏览记录。以上证据可用于讨论界面结构，不能外推为全部接口、权限和异常场景均已通过。')
p('从架构上看，组件路由使首页可复用媒体模块中的列表和播放入口，公共列表基类减少刷新与加载逻辑的重复。MVVM 将界面状态与数据访问分开，Room 为历史回看提供本地数据基础。这些设计有助于组织功能，但其实际性能收益仍需测量，不能据此给出未经测试的耗时或内存指标。')
h('5.3 问题分析与改进方向',2)
table(['发现的问题','影响与改进方式'],[
('分页请求前推进页码','失败后可能跳页；成功后提交页码或失败回退'),('收藏数量更新表达式','正常数字未执行加减；明确括号并同步详情对象'),('评论先移除再请求','失败后列表失真；成功后删除或回滚'),('新旧播放状态并存','部分弹窗与 Kotlin 状态源不一致；统一页面共享状态'),('生命周期恢复入口','短暂遮挡后需核对恢复；尊重用户主动暂停状态'),('浏览记录先查后写','并发存在重复风险；组合唯一约束与事务')],[4.4,10.25])

h('5.3.1 收藏数量更新的代码分析',3,new=True)
code('代码 5-1　当前收藏数量的更新表达式', '''val newCollections = mCollections.value?.toInt()
    ?: ERRORCOLLECTIONS + 1''','MediaPlayViewModelKotlinVersion.kt；按原表达式换行')
p('该表达式中，加一作用于 Elvis 运算符右侧的后备值。若 mCollections.value 能正常转换为整数，结果直接取该整数，不会执行加一；取消收藏的相应表达式也有同类问题。这会使收藏状态已变化而显示数量未变化。')
code('代码 5-2　建议的数值更新方式', '''val current = mCollections.value?.toIntOrNull() ?: 0
val updated = current + 1
mCollections.value = updated.toString()
mArchivesInfo.value?.collection = updated''','改进示例，尚未替换工程实现')
p('建议仅在接口确认成功后执行数量更新，取消收藏则减一并限制不小于零。若服务端可直接返回最新计数，应优先采用服务端值。连续点击时还应禁用按钮或串行处理，防止重复请求导致状态与数量不一致。')
h('5.4 需要补充的运行证据',2)
p('建议补充三组截图：脱敏后的验证码登录页与登录成功状态；同一视频点赞或收藏前后及重新进入页面后的对照；浏览记录删除前后与重新进入后的对照。若展示相机能力，再补一张 CameraX 预览或拍照结果图。')
p('建议再记录一次断网后的错误提示及重试结果，并填写测试设备、Android 版本和执行日期。截图中的完整手机号、验证码和令牌应隐藏，现有相册截图不能代替相机截图。')

h('六　小结与心得体会',new=True)
for i in range(119,124):old(i)
p('后续完善将优先处理分页、收藏计数和评论删除的一致性，再开展生命周期与账号切换回归测试。相比单纯增加页面，先保证核心流程在失败、返回和重复操作后仍能正确工作，更有助于形成可靠的移动应用开发习惯。')

# Insert linked contents using heading bookmarks, refreshed through Word.
for text,key,level in toc:
 if level==2 and (text.startswith('4.4.') or text.startswith('4.5.') or text.startswith('4.7.')):continue
 pp=toc_anchor.insert_paragraph_before();pp.paragraph_format.first_line_indent=Pt(0);pp.paragraph_format.left_indent=Pt(12 if level==2 else 0);pp.paragraph_format.space_before=Pt(0);pp.paragraph_format.space_after=Pt(2);pp.paragraph_format.line_spacing=1.15
 pp.paragraph_format.tab_stops.add_tab_stop(Cm(14),2,1)
 pp.add_run(text+'\t');field=OxmlElement('w:fldSimple');field.set(qn('w:instr'),'PAGEREF '+key+' \\h');rr=OxmlElement('w:r');tt=OxmlElement('w:t');tt.text='1';rr.append(tt);field.append(rr);pp._p.append(field)
toc_anchor._p.getparent().remove(toc_anchor._p)
for i,pp in enumerate(d.paragraphs[:9]):
 if i==1:pp.style='Title'
d.core_properties.author='李思民';d.core_properties.title='智能终端系统及应用开发课程设计报告'
out=root/'智能终端系统及应用开发课程设计报告_李思民_图文优化版.docx';d.save(out)
print(out.resolve())
