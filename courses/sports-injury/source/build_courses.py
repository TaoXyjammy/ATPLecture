"""Build two independent, editable sports-injury teaching decks.
Python 3 + python-pptx; run from any directory. PDF export uses LibreOffice separately.
"""
from pathlib import Path
import json
ROOT=Path(__file__).resolve().parent
exec((ROOT/'knee_foundation.py').read_text())
OUT=ROOT.parent
KNEE=R
for s in R.slides:
 for sh in s.shapes:
  if sh.has_text_frame:
   for p in sh.text_frame.paragraphs:
    for run in p.runs:run.text=run.text.replace('基本检查、专项检查','基本与专项检查')
KCPG='https://doi.org/10.2519/jospt.2023.0301'
ACPG='https://doi.org/10.2519/jospt.2021.0302'
PAASS='https://doi.org/10.1136/bjsports-2021-104087'
NATA='https://doi.org/10.4085/1062-6050-99-16'
ATLAS='https://github.com/keanu77/professional-sports-injury-atlas'
COURSE='https://github.com/Sithranjan-Suresh/injury-risk-ml-curriculum'
CHAPTER='膝关节运动损伤与预防';SECTION='运动损伤与预防 · 体育学科课程'
OLD_SLIDE=slide

def course_slide(title,src='教材 p.136—139',note=''):
 s=R.slides.add_slide(R.slide_layouts[6]);s.background.fill.solid();s.background.fill.fore_color.rgb=RGBColor.from_string('FFFFFF')
 for j,c in enumerate(['344F91','4C73BC','126B97','2FA6CF','DA343A','E3E8F1']):box(s,.22+j*.29,.13,.18,.18,c)
 box(s,.22,.48,.19,.19,'91C64D',kind=MSO_SHAPE.OVAL);txt(s,CHAPTER,.51,.45,6,.35,14,'314E88',True)
 txt(s,SECTION,8.1,.22,4.6,.35,13,GRAY,True,PP_ALIGN.RIGHT);box(s,5,.59,7.7,.012,'DFE5EE')
 txt(s,title,.65,.94,12,.8,30,INK,True);box(s,0,7.07,13.333,.43,'F1F1F1');box(s,0,7.46,13.333,.04,'5579B2')
 txt(s,'运动损伤与预防  |  '+src,.58,7.17,10.7,.25,10,GRAY)
 s.notes_slide.notes_text_frame.text=note+'\n来源：'+src+'。图形为原创可编辑教学示意，非比例解剖图。'
 return s

def replace(s,old,new):
 for sh in s.shapes:
  if sh.has_text_frame and old in sh.text:
   for p in sh.text_frame.paragraphs:
    if old in p.text:
     # These paragraphs are single-run text in the source deck.
     for run in p.runs:
      if old in run.text:run.text=run.text.replace(old,new)

def erase_content(s,title):
 for sh in list(s.shapes):
  if Inches(.8)<=sh.top<Inches(7):sh._element.getparent().remove(sh._element)
 txt(s,title,.65,.94,12,.8,30,INK,True)

def notes(s,t):s.notes_slide.notes_text_frame.text=t

def insert(s,pos):
 ids=R.slides._sldIdLst;e=ids[-1];ids.remove(e);ids.insert(pos,e)

def remove_slide(idx):
 ids=R.slides._sldIdLst;e=ids[idx];R.part.drop_rel(e.rId);ids.remove(e)

def pose(s,cx,y,kind='neutral',scale=1,c=B):
 # Frontal stick figure: explicit joint positions support observation cues.
 k=scale
 def pt(dx,dy):return cx+dx*k,y+dy*k
 def seg(a,b,col=c,w=5):line(s,*pt(*a),*pt(*b),col,w*k)
 box(s,cx-.18*k,y,.36*k,.36*k,c,kind=MSO_SHAPE.OVAL)
 seg((0,.43),(0,1.25));seg((-.4,.65),(.4,.65));seg((-.4,.65),(-.8,.95));seg((.4,.65),(.8,.95));seg((-.4,1.25),(.4,1.25))
 knees=[(-.12,1.78),(.12,1.78)] if kind=='valgus' else [(-.48,1.78),(.48,1.78)]
 if kind=='balance':knees=[(-.4,1.9),(.65,1.58)]
 ankles=[(-.6,2.35),(.6,2.35)] if kind!='balance' else [(-.45,2.4),(.85,1.98)]
 for hip,knee_,ank in zip([(-.4,1.25),(.4,1.25)],knees,ankles):
  seg(hip,knee_);seg(knee_,ank);seg((ank[0]-.17,ank[1]),(ank[0]+.17,ank[1]))
  px,py=pt(*knee_);box(s,px-.075*k,py-.075*k,.15*k,.15*k,ORANGE,kind=MSO_SHAPE.OVAL)
 line(s,cx-1.15*k,y+2.5*k,cx+1.15*k,y+2.5*k,'A4B3C8',1.5)

def cards(s,items):
 for i,(title,body) in enumerate(items):
  x=.7+i*4.17;box(s,x,2.12,3.88,3.65,PALE);txt(s,f'0{i+1}',x+.24,2.35,3.3,.6,35,B,True);txt(s,title,x+.24,3.2,3.4,.55,25,NAVY,True);txt(s,body,x+.24,4.12,3.35,1.35,21)

def refs(s,items):
 for i,(a,b) in enumerate(items):
  y=1.95+i*1.12;txt(s,a,.85,y,11.6,.42,20,NAVY,True);txt(s,b,.85,y+.44,11.6,.70,14,GRAY)

def finish(name):
 from pptx.opc.packuri import PackURI
 for idx,s in enumerate(R.slides):
  s.part._partname=PackURI(f'/ppt/slides/slide{idx+1}.xml')
  s.notes_slide.part._partname=PackURI(f'/ppt/notesSlides/notesSlide{idx+1}.xml')
  for q in list(s.shapes):
   if q.has_text_frame and q.left>Inches(11.5) and q.top>Inches(7.08):q._element.getparent().remove(q._element)
  txt(s,f'{idx+1:02d} / {len(R.slides):02d}',11.65,7.14,1.05,.3,12,GRAY,False,PP_ALIGN.RIGHT)
  for sh in s.shapes:
   st=sh._element.find('{http://schemas.openxmlformats.org/presentationml/2006/main}style')
   if st is not None:
    ef=st.find('{http://schemas.openxmlformats.org/drawingml/2006/main}effectRef')
    if ef is not None:ef.set('idx','0')
 R.core_properties.title=CHAPTER;R.core_properties.subject=SECTION;R.core_properties.author='谢淘'
 R.save(OUT/name)
 (OUT/(name.replace('.pptx','_讲解备注.md'))).write_text('\n\n'.join(f'## {i+1}. '+next((q.text for q in s.shapes if q.has_text_frame and Inches(.8)<=q.top<Inches(1.4)),'')+'\n\n'+s.notes_slide.notes_text_frame.text for i,s in enumerate(R.slides)))
 print(name,len(R.slides),'slides')

# --- Knee: retain textbook mechanisms; shorten repeated clinical test slides. ---
for s in R.slides:
 replace(s,'第二章  膝关节功能评定',CHAPTER)
 replace(s,'第一节  膝关节损伤的功能障碍',SECTION)
 replace(s,'运动康复评定学  |','运动损伤与预防  |')
replace(R.slides[0],'膝关节损伤的功能障碍',CHAPTER)
replace(R.slides[0],'从损伤机制','理解运动损伤')
replace(R.slides[0],'到功能评定','学会科学预防')
replace(R.slides[0],'课堂主线：哪里受伤 → 如何判断 → 哪些功能受到影响','课堂主线：运动情境 → 损伤机制 → 识别线索 → 预防措施')
notes(R.slides[0],'体育学科课堂，以运动参与者、体育教师和教练员视角讨论。主体依据教材p.136—139，预防拓展依据JOSPT 2023与NATA 2018。全课可按约40—45分钟组织，具体按学生基础调整。')
s=R.slides[3];erase_content(s,'体育教师如何把“损伤知识”转化为预防？')
flow(s,['描述运动情境','理解受力机制','识别异常线索','安排预防措施'])
for x,t in [(.95,'接触还是非接触\n突发还是渐进'),(4.13,'方向、速度与负荷\n哪类结构受累'),(7.31,'疼痛、肿胀、不稳\n何时停止参与'),(10.49,'动作、体能、负荷\n环境与组织管理')]:txt(s,t,x,4,2.5,1.25,20,GRAY)
band(s,'课堂既要解释“为什么伤”，也要回答“下一次如何减少风险”。')
notes(s,'教学归纳。体育教师重点是识别危险情境、科学安排参与和及时转介，不是现场给出韧带分级。与后续预防练习建立联系。')
# Remove four repeated test-specific slides; keep the comparative table.
for idx in [11,9,7,5]:remove_slide(idx)
# Position 10 is now the case exercise; insert preventive teaching before it.
pos=10
s=course_slide('从损伤机制中寻找可改变的环节','预防拓展 · 教学归纳', '让学生把机制翻译成可操作的训练内容。既往损伤与个体结构影响风险，但一次动作筛查不能准确预测谁会受伤。预防是降低概率，不能保证不受伤。多因素观点依据NATA 2018及JOSPT 2023：'+NATA+'；'+KCPG)
cards(s,[('动作与控制','减速、落地、变向\n躯干与下肢协调'),('力量与负荷','髋膝力量、跳跃能力\n疲劳与恢复安排'),('环境与组织','场地、鞋与项目匹配\n对抗规则与人员间距')]);band(s,'风险来自多个因素；预防也应覆盖多个环节。');insert(s,pos);pos+=1
s=course_slide('落地教学：先看缓冲，再看对线与控制','预防拓展 · NATA 2018', '原创示意，仅展示可见动作。观察屈髋屈膝、落地稳定与膝内扣；这些表现不能单独预测ACL损伤。教学反馈可用“轻稳落地、髋膝共同缓冲、保持方向控制”。先低难度再增加速度、单脚与专项干扰。参考：'+NATA)
box(s,.7,2.03,5.28,3.92,'F7F9FC','DCE4F0');pose(s,2.12,2.5,'valgus',.95,RED);pose(s,4.55,2.5,'neutral',.95,B)
label(s,'观察内扣',1.17,5.24,2,RED);label(s,'练习稳定控制',3.55,5.24,2.3,B)
explain(s,[('看缓冲','避免僵硬落地，学习髋膝协同吸收负荷。'),('看控制','观察膝、髋、躯干的协调与落地稳定。'),('给反馈','降低高度与速度，动作稳定后逐步进阶。')]);band(s,'动作练习应伴随反馈；一张姿势图不能证明是否会发生损伤。');insert(s,pos);pos+=1
s=course_slide('预防训练：力量、跳跃与近端控制相结合','预防拓展 · JOSPT 2023', '示例供已能安全参加运动者课堂学习。以深蹲/弓步、桥式/髋控制、低幅跳跃落地为动作类别示例，不是适合所有人的固定处方。训练应由专业人员按个体调整。核心概念：多组分，不能只做静态拉伸或只练平衡。参考：'+KCPG)
cards(s,[('力量练习','深蹲、弓步等基础模式\n关注动作质量与控制'),('跳跃与落地','由低幅双脚开始\n逐步过渡到专项任务'),('近端控制','髋与躯干稳定\n结合减速、方向变化')]);band(s,'预防不是某一个“神奇动作”，而是持续实施的组合训练。');insert(s,pos);pos+=1
s=course_slide('把预防融入教学：练习—观察—反馈—复练','预防拓展 · 教学组织示例', '这是教师组织练习的设计示例，不是经验证的完整预防方案，也不替代指南所述训练剂量。按已选项目方案安排，持续实施并监督完成质量。证据要点：在季前开始并在赛季持续，重视依从性。参考：'+KCPG)
flow(s,['选基础动作','分组轮换练习','教师观察反馈','记录与持续实施'])
for x,t in [(.95,'先低难度、低速度\n明确本次观察点'),(4.13,'力量与落地任务\n保持足够空间'),(7.31,'一次给出一个提示\n疲劳时降低难度'),(10.49,'按既定方案持续练习\n记录完成度与不适')]:txt(s,t,x,4,2.5,1.25,20,GRAY)
band(s,'持续、规范地完成，通常比偶尔安排一次“预防课”更有价值。');insert(s,pos);pos+=1
s=course_slide('不同损伤，预防重点如何对应？','教材 p.136—139 · 教学归纳', '将教材机制转化为课堂训练与环境管理建议，不主张某个动作已被证实能单独预防某一病种。对髌股疼痛者，应个体化调整诱发负荷并转介评估；不要把髌骨软化和所有膝前痛视为同义词。')
items=[('ACL / 半月板','练习减速、变向与落地控制；\n减少突然大强度扭转任务。'),('副韧带损伤','规范对抗与保护动作，控制人员间距；\n关注接触后疼痛与侧向不稳。'),('髌股相关问题','渐进安排跑跳与屈膝负荷；\n出现持续膝前痛时及时调整与评估。')]
for i,(h,b) in enumerate(items):
 y=2.05+i*1.28;chip(s,h,.85,y,3.05);txt(s,b,4.28,y,8,1.04,23)
band(s,'有既往损伤或持续症状者，训练内容需与康复专业人员协同调整。');insert(s,pos);pos+=1
s=course_slide('运动现场：哪些表现需要停止参与？','教材 p.136—139 · 现场识别', '依据教材急性功能障碍内容的情境归纳。明显畸形、不能负重、严重痛肿、疑似脱位或交锁均需及时医疗评估。现场不强行复位，不让同学反复做诱痛检查。若有足部发冷、麻木、颜色异常或开放伤，应紧急处理。')
cards(s,[('停止参与','响声后急性肿胀\n不能负重或明显不稳'),('保护与转介','明显畸形、疑似脱位\n交锁或剧烈疼痛'),('清楚交接','受伤动作与接触方式\n症状时间与负重能力')]);band(s,'体育教师的关键任务：识别、保护、转介，并记录运动情境。');insert(s,pos)
# Update ending: case answers + take-home lesson.
s=R.slides[len(R.slides)-4]
# final case explanation remains evidence-based; mention prevention in its speaker notes.
s.notes_slide.notes_text_frame.text+='\n追问：在恢复运动参与后，可从规范变向、落地训练与持续组合预防训练减少再伤风险；回场由医疗团队综合判断。'
s=R.slides[len(R.slides)-2];erase_content(s,'课堂回扣：解释机制，识别风险，落实预防')
for i,(h,b) in enumerate([('机制','落地、变向、负重扭转，如何影响膝关节？'),('识别','痛、肿、卡锁、打软腿，哪些提示停止参与？'),('预防','多组分训练 + 规范组织 + 持续反馈与执行。')]):
 y=2.12+i*1.23;chip(s,h,.85,y,1.65);txt(s,b,2.95,y+.06,9.2,.75,26,INK,True)
band(s,'预防的目标是降低运动风险，并保持安全、可持续的运动参与。')
notes(s,'回扣体育学科的应用：不仅会背损伤名称，还要能设计组织方式、识别转介情况、落实持续预防练习。')
s=R.slides[len(R.slides)-1];erase_content(s,'教材与 GitHub 教学参考')
refs(s,[('教材：膝关节常见损伤','用户提供《运动康复评定学-膝关节损伤》，p.136—139。'),('GitHub：运动情境与课堂提问','Sithranjan-Suresh/injury-risk-ml-curriculum · Week_1 课程（25页）'),('GitHub：比赛事件与医学判断分开','keanu77/professional-sports-injury-atlas · 机制观察与证据链接'),('采用方式','借鉴讲解组织；示意图、练习与病例重新编写，不复用第三方影音。')])
notes(s,'GitHub实际查阅：'+COURSE+'/blob/main/Week_1/Mini_Missions_Week1_presentation.pdf\n'+ATLAS+'\n课程是机器学习入门示例，仅借鉴“受伤情境→问题→分步解释→提问”的教学方式，不采用其数据或预测结果作为医学证据。图鉴医学内容标示在审阅中；本课预防结论单独依据指南。')
s=course_slide('预防拓展的证据依据','预防拓展 · 参考文献', '新增预防部分不是所上传教材的原文。以下指南与立场声明提供基础原则；课堂中的练习编排与情境为教学设计。')
refs(s,[('JOSPT 2023：膝与 ACL 损伤的运动预防','Arundale et al. Exercise-Based Knee and Anterior Cruciate Ligament Injury Prevention.\n'+KCPG),('NATA 2018：ACL 损伤预防立场声明','Padua et al. National Athletic Trainers’ Association Position Statement.\n'+NATA),('本课的证据边界','组合训练与动作反馈有依据；不把单个姿势、单个动作当作损伤预测或预防保证。')]);band(s,'完整来源、课堂提示和病例说明见各页“备注”。')
finish('Knee_Sports_Injury_Prevention.pptx')

# --- Ankle: independent deck grounded in the 6 uploaded pages. ---
R=Presentation();R.slide_width=Inches(13.333);R.slide_height=Inches(7.5)
CHAPTER='踝足运动损伤与预防'

def ankle(s,mode='lig',x=.95,y=2.22):
 # Lateral relationship diagram, toes to the right. Bone geometry simplified.
 box(s,x+1.35,y,.68,1.55,'D8E2EF',B,MSO_SHAPE.ROUNDED_RECTANGLE)
 box(s,x+2.14,y+.18,.36,1.62,'BCCDE1',B,MSO_SHAPE.ROUNDED_RECTANGLE)
 box(s,x+1.5,y+1.62,1.3,.55,'D8E2EF',B,MSO_SHAPE.OVAL)
 box(s,x+.78,y+2.18,1.6,.63,'D8E2EF',B,MSO_SHAPE.OVAL)
 box(s,x+2,y+2.37,2.3,.42,'D8E2EF',B,MSO_SHAPE.ROUNDED_RECTANGLE)
 if mode=='lig':
  line(s,x+2.34,y+1.65,x+2.7,y+1.97,RED,5);line(s,x+2.28,y+1.72,x+1.9,y+2.46,ORANGE,5)
  label(s,'距腓前韧带',x+2.95,y+1.3,2,RED);line(s,x+3.1,y+1.72,x+2.6,y+1.85,RED)
  label(s,'跟腓韧带',x+.03,y+1.75,1.8,ORANGE);line(s,x+1.3,y+2.1,x+2.1,y+2.05,ORANGE)
 elif mode in ['achilles','rupture']:
  if mode=='achilles':line(s,x+.9,y+.38,x+1.15,y+2.4,RED,7)
  else:
   line(s,x+.9,y+.38,x+1.02,y+1.23,RED,7);line(s,x+1.08,y+1.56,x+1.15,y+2.4,RED,7)
  label(s,'跟腱',x+.02,y+.42,1.3,RED);label(s,'跟骨',x+.1,y+2.85,1.3,B)
 elif mode=='plantar':
  line(s,x+1.25,y+2.9,x+4.05,y+2.9,TEAL,5);box(s,x+1.16,y+2.66,.36,.36,RED,kind=MSO_SHAPE.OVAL)
  label(s,'足底筋膜',x+2.5,y+3.03,2,TEAL);label(s,'近足跟疼痛',x+.05,y+3.03,2,RED)
 label(s,'胫骨',x+.4,y+.1,1,B);label(s,'腓骨',x+2.8,y+.1,1,B)
 txt(s,'外侧结构关系示意 · 足尖朝右 · 非比例',.95,5.76,4.9,.25,11,GRAY)

def ankle_split(title,src,mode,items,tag,note):
 s=course_slide(title,src,note);panel(s);ankle(s,mode);explain(s,items);band(s,tag);return s

s=course_slide(CHAPTER,'教材 p.164—169 · 预防拓展', '体育学科课程，建议根据授课时长选取核心踝扭伤部分或完整讲授。全课涵盖踝扭伤、跟腱病与断裂、足底筋膜、胫骨内侧应力综合征及骨折。核心主线是运动动作、负荷、识别和预防。')
txt(s,'理解崴脚\n减少再次受伤',.85,2.03,7,1.85,43,NAVY,True);txt(s,'踝扭伤 · 跟腱 · 足跟痛\n负荷管理与安全运动参与',.9,4.36,6.7,1.05,25,GRAY);ankle(s,'lig',7.55,2.16);band(s,'课堂主线：运动情境 → 损伤机制 → 风险识别 → 预防与回场')
# Cover graphic extends within slide; remove fixed left caption (cover only).
for q in list(s.shapes):
 if q.has_text_frame and '外侧结构关系示意' in q.text:q._element.getparent().remove(q._element)
s=course_slide('落地踩到别人脚上，为什么会“崴脚”？','教材 p.164 · 教学情境', '虚构篮球落地情境。先问学生看见了什么，再问哪些诊断结论不能仅靠动作外观确定。借鉴GitHub运动伤害图鉴的“可见事件与医学判断分开”组织方式：'+ATLAS)
box(s,.7,2.04,5.25,3.92,PALE);txt(s,'篮球争抢篮板后落地，\n足部踩到他人脚面，\n随即翻转、疼痛，\n难以继续跑跳。',1.02,2.55,4.65,2.65,27,NAVY,True)
explain(s,[('先描述事件','落地、接触、足部翻转与疼痛位置。'),('再解释机制','外侧结构受到牵拉，可能出现韧带损伤。'),('保留未确定处','仅凭外观不能分级，也不能排除骨折。')]);band(s,'看见“崴脚动作”不等于已经确认“轻度扭伤”。')
ankle_split('认识踝外侧韧带：稳定从何而来？','教材 p.164','lig',[('骨性稳定','距骨位于踝穴内；踝位改变时稳定条件不同。'),('外侧韧带','距腓前韧带等协助限制过度内翻。'),('动态稳定','周围肌群与感觉运动控制共同参与稳定。')],'结构、肌力与控制共同决定运动中的稳定表现。','教材为骨性与韧带结构依据；动态稳定为JOSPT2021补充。'+ACPG)
#4
s=course_slide('跖屈 + 内翻：外侧扭伤的典型情境','教材 p.164', '教材描述踝跖屈位内翻是常见情境。足尖向下时踝穴适配较松，内翻牵拉外侧韧带，距腓前韧带常受累。不是所有扭伤都只有这一机制；外翻与外旋可累及其他结构。')
panel(s);chip(s,'足尖向下：跖屈',1.05,2.38,4.5);chip(s,'足底向内：内翻',1.05,3.52,4.5,RED);txt(s,'组合受力 → 外侧韧带牵拉',1.07,4.8,4.55,.85,25,NAVY,True)
explain(s,[('典型发生时刻','落地、变向、踩空或不平整地面。'),('常见受累结构','外侧韧带，尤其距腓前韧带。'),('不能一概而论','外翻或外旋受伤，要关注其他组织。')]);band(s,'解释机制要同时说明关节姿态与受力方向。')
#5
s=course_slide('扭伤严重度：重点看功能与稳定性','教材 p.164—165', '教材列I—III级，包含疼痛、肿胀、活动度和韧带松弛。本课简化为严重度理解，不以肿胀厘米数或疼痛程度现场定级。学生不互相进行应力试验，临床分级由专业评估完成。')
cards(s,[('I 级：较轻','功能受限相对较小\n通常无明显韧带松弛'),('II 级：部分损伤','疼痛、肿胀更明显\n活动与稳定性受影响'),('III 级：较重','功能明显受损\n可有显著不稳与韧带断裂')]);band(s,'肿胀大小不能单独决定分级，也不能排除骨折或其他损伤。')
#6
s=course_slide('“只是崴了一下”——先排查需要转介的线索','教材 p.164、169', '教材强调内外踝、第五跖骨基底与舟骨等骨性部位触诊，以及必要X线；骨折可有痛肿、不能走路与畸形。体育现场不反复弹击、扭转或强拉关节来排除骨折；不能负重、明显骨性压痛、畸形、麻木发冷或开放伤，应及时医疗评估。')
cards(s,[('观察与询问','痛在哪里？何时肿？\n能否负重和行走？'),('关注危险表现','骨性部位明显痛\n畸形、麻木或足部发冷'),('停止与转介','疑似骨折或严重损伤\n保护患部并尽快就医')]);band(s,'可见伤情、症状和负重表现需要一并记录。')
#7
s=course_slide('为什么“不痛了”仍可能反复崴脚？','教材 p.165 · JOSPT 2021', '教材提到本体感觉、平衡和不稳问题，CPG强调既往扭伤和感觉运动控制。以下为概念归纳，不表示每个人都会出现同一因果链。疼痛减轻不代表肌力、平衡和专项能力全部恢复。'+ACPG)
flow(s,['一次扭伤','功能尚未恢复','过早重返运动','再次失稳风险'])
for x,t in [(.95,'痛肿与活动受限\n改变负重方式'),(4.13,'肌力、平衡、信心\n可能仍有不足'),(7.31,'跑跳变向需求较高\n身体尚未适应'),(10.49,'关注既往扭伤者\n安排复评与预防')]:txt(s,t,x,4,2.5,1.25,20,GRAY)
band(s,'再伤预防要补足功能短板，而不只等疼痛消失。')
#8
s=course_slide('预防重点一：把平衡练习做成可进阶的任务','预防拓展 · JOSPT 2021', 'CPG支持本体感觉和平衡训练减少复发。图示为已可安全负重者的任务类别，并非急性伤后处方。可从稳地面单脚站立到触点和接传球；进阶由专业人员按耐受决定，不把闭眼或不稳定垫作为人人必须的首选。'+ACPG)
box(s,.7,2.03,5.28,3.92,'F7F9FC','DCE4F0');pose(s,3.27,2.43,'balance',1.07,B)
explain(s,[('基础任务','稳定地面单脚支撑，观察摇摆与代偿。'),('增加任务','触点、接传球或方向变化，逐步提高难度。'),('反馈与调整','保持安全支撑；质量下降时退回较易任务。')]);band(s,'平衡训练尤其值得用于有既往踝扭伤史的人群。')
#9
s=course_slide('预防重点二：力量与专项控制一起练','预防拓展 · 课堂任务示例', '动作是功能训练类别示例。小腿提踵、踝周肌群抗阻和专项控制可按能力逐步安排。CPG支持结构化康复与感觉运动训练，但这里的三个任务不等同于一项独立验证的预防方案，不提供统一组数与次数。'+ACPG)
cards(s,[('踝周力量','提踵与踝周抗阻\n观察力量与耐力差异'),('下肢协同','髋、膝、踝共同控制\n避免只盯住踝关节'),('专项衔接','稳定站立到落地\n再到减速、变向任务')]);band(s,'先能稳定完成，再增加速度、复杂度和运动情境。')
#10
s=course_slide('预防重点三：落地空间、护具与训练执行','预防拓展 · JOSPT 2021 / 教学组织', 'CPG推荐适当护具与平衡训练用于再伤预防；护具类型与使用需个体评估。场地与人员间距为体育教学组织建议，不声称某种鞋可普遍防止扭伤。预防训练及支具都不能保证零风险。'+ACPG)
cards(s,[('空间与规则','跑跳练习保持间距\n检查地面与障碍物'),('既往扭伤者','评估是否适合护踝\n同时坚持功能训练'),('规范执行','教会动作与反馈\n记录不适和完成情况')]);band(s,'护踝可提供辅助支持，不能代替功能恢复与训练。')
#11
s=course_slide('何时回到球场？用 PAASS 组织复评','预防拓展 · PAASS 共识 2021', 'PAASS是专家共识框架，不是已验证的统一达标分数。五领域：疼痛；踝关节功能；运动员感知；感觉运动控制；运动/功能表现。回场由医疗与康复团队结合专项需要综合判断，不能按固定休息天数。'+PAASS)
for i,(a,b) in enumerate([('P  疼痛','运动中与近24小时反应'),('A  踝功能','活动度、力量、耐力与爆发力'),('A  自我感知','稳定感、信心与心理准备'),('S  感觉运动控制','本体感觉与动态平衡'),('S  运动表现','跳跃、敏捷、专项练习与完整训练')]):
 y=1.99+i*.75;chip(s,a,.85,y,3.65);txt(s,b,4.85,y+.08,7.5,.55,23)
band(s,'“能走路”与“准备好比赛”之间，还需要一组功能证据。')
#12
ankle_split('跟腱病：反复跳跃与蹬地后的负荷问题','教材 p.165—166','achilles',[('常见背景','反复跑跳、负荷增加及相关活动度不足。'),('常见表现','晨僵、局部疼痛，提踵力量或耐力下降。'),('预防关注','训练渐进、恢复安排，持续不适及时评估。')],'跟腱病与急性断裂不同；不能把疼痛都当作普通训练酸痛。','教材的跟腱病变描述为主体，预防为机制导向教学归纳。跟腱痛不等于只有炎症，注意负荷反应。')
#13
ankle_split('跟腱断裂：突然“被踢感”与蹬地无力','教材 p.166','rupture',[('典型事件','突然加速、起跳或被动过度背伸。'),('重要线索','后踝突发痛或响声，蹬地无力，可有凹陷。'),('现场处理重点','停止运动、保护并尽快就医，避免拉伸测试。')],'即使仍能部分行走，也不能据此排除跟腱断裂。','教材列Thompson征及超声/MRI辅助。专门检查由专业人员进行。课堂不让学生互相模拟急性断裂检查；不强行提踵、拉伸或继续比赛。')
#14
ankle_split('足底筋膜相关疼痛：晨起第一步为什么痛？','教材 p.167','plantar',[('疼痛位置','足跟内侧及足底近跟骨区域。'),('典型线索','休息后第一步更痛，长时间负重可再次加重。'),('预防关注','逐步适应跑量与站立负荷，留意持续足跟痛。')],'位置与时间规律比笼统说“脚跟痛”更有信息量。','教材足底筋膜炎内容。预防为机制导向归纳，不将所有足跟痛都确诊为足底筋膜炎，需与跟腱和其他疾病鉴别。')
#15
s=course_slide('胫骨内侧应力综合征：训练加量后的“小腿痛”','教材 p.168', '教材指出胫骨内侧应力综合征与训练量增加过快、地面、下肢受力等有关。较弥散的内侧胫骨疼痛可见于本病，但局灶剧痛、休息痛或持续加重时需排查骨应力损伤；不能靠一个触点在课堂确诊。')
box(s,.7,2.03,5.25,3.92,'F7F9FC','DCE4F0');box(s,3.1,2.42,.7,2.85,'D8E2EF',B,MSO_SHAPE.ROUNDED_RECTANGLE);line(s,3.02,3.65,3.02,5.1,RED,6);label(s,'胫骨内侧疼痛区域',1.03,5.42,4.6,RED);txt(s,'小腿区域概念示意',1.02,2.18,4.5,.3,13,GRAY)
explain(s,[('常见诱因','跑量增加过快、反复跑跳、地面改变。'),('症状规律','沿小腿内侧疼痛，活动时加重。'),('及时调整','降低诱发负荷；持续或局灶疼痛及时评估。')]);band(s,'训练加量后的疼痛需要追踪，不应一律解释为“练开就好”。')
#16
s=course_slide('过度使用损伤：让负荷变化看得见','教材 p.165、167—168 · 教学归纳', '课表、训练日记和症状反馈为教学管理示例，不将通用固定百分比或某个负荷比值当作安全阈值。根据既往训练和个体反应安排，出现持续不适要评估。归纳自教材跟腱、足底筋膜、胫骨应力相关机制。')
flow(s,['记录原有负荷','辨认近期变化','观察身体反应','调整并复评'])
for x,t in [(.95,'跑量、跳跃、比赛\n休息与恢复'),(4.13,'强度、次数、地面\n是否同时改变'),(7.31,'运动中与次日疼痛\n僵硬、力量和疲劳'),(10.49,'减少诱发任务\n再渐进增加需求')]:txt(s,t,x,4,2.5,1.25,20,GRAY)
band(s,'没有适用于所有人的“安全加量公式”；负荷要与个体反应对应。')
#17
s=course_slide('踝部骨折：不能把所有肿痛都当成扭伤','教材 p.169', '教材骨折常见机制包括高处坠落、直接暴力及间接暴力，表现有痛肿、瘀斑、不能走路与畸形，X线等辅助。现场不重复检查骨擦感或做纵向叩击来试图确诊。保护患肢，明显畸形、开放伤或感觉循环异常需急救。')
cards(s,[('机制线索','较大暴力、坠落\n碰撞或明显扭转'),('表现线索','不能负重、明显骨痛\n畸形或严重肿胀'),('行动原则','停止、保护、转介\n不强行走动或复位')]);band(s,'肿胀、瘀斑可见于扭伤和骨折，现场无法仅凭外观区分。')
#18
s=course_slide('足踝疼痛：按“位置 + 过程 + 功能”归纳','教材 p.164—169', '此表为症状线索学习，不是自我诊断表。可将各列遮住让学生回答。多个问题可同时存在，持续症状需要专业评估。')
rows=[['问题','位置 / 过程','运动功能线索'],['外侧踝扭伤','外踝附近；急性翻转后','负重痛、跑跳受限或不稳'],['跟腱病','后踝跟腱；渐进负荷相关','晨僵、提踵耐力下降'],['跟腱断裂','后踝突发痛 / 被踢感','蹬地与踮脚明显无力'],['足底筋膜相关痛','足底近足跟；晨起第一步痛','负重后加重'],['胫骨内侧应力综合征','小腿内侧；跑量变化后','跑跳痛，需关注持续加重'],['踝部骨折','外伤后骨性部位痛肿','不能负重或明显畸形']]
widths=[3.15,4.55,4.33]
for i,row in enumerate(rows):
 x=.65
 for j,t in enumerate(row):
  box(s,x,1.95+i*.56,widths[j],.56,B if i==0 else ('EDF2FB' if i%2 else 'F7F9FC'),'FFFFFF');txt(s,t,x+.12,2.07+i*.56,widths[j]-.2,.4,18,'FFFFFF' if i==0 else INK,i==0 or j==0);x+=widths[j]
band(s,'位置用于缩小范围；病史、体征与必要影像共同支持诊断。')
#19
s=course_slide('体育课堂的现场流程：停—护—记—转','教材 p.164—169 · 教学组织', '流程为教材症状和功能障碍内容的课堂应用。停止危险活动，保护伤肢，记录动作、症状和负重状况，与医疗团队交接。紧急表现及时呼叫急救。对急性疼痛不强行“走走看”、拉伸或让同学复位。')
flow(s,['停止活动','保护患部','记录信息','及时转介'])
for x,t in [(.95,'离开危险运动情境\n防止再次碰撞'),(4.13,'按伤情减少负重\n避免强行扳动'),(7.31,'动作、痛点、时间\n肿胀与行走能力'),(10.49,'严重表现急诊处理\n持续症状专业评估')]:txt(s,t,x,4,2.5,1.25,20,GRAY)
band(s,'体育教师负责风险识别与安全组织；专业诊断与回场决策需协作。')
#20
s=course_slide('一堂体育课里，预防练习如何组织？','预防拓展 · 教学组织示例', '只示范组织方式，不宣称下列四环节已构成验证过的独立预防计划。可结合项目、年龄和能力选择指南支持的平衡与功能训练，监督执行。当前疼痛、未获安全参与许可者不直接进入跑跳站。'+ACPG)
flow(s,['渐进热身','平衡与力量','专项动作','反馈与记录'])
for x,t in [(.95,'低强度移动\n逐步进入运动'),(4.13,'安全单脚支撑\n踝周力量任务'),(7.31,'落地、减速、变向\n先可控再复杂'),(10.49,'观察质量与疲劳\n追踪不适与完成度')]:txt(s,t,x,4,2.5,1.25,20,GRAY)
band(s,'预防要长期融入日常教学，不能只在受伤后临时补做。')
#21
s=course_slide('课堂讨论：不痛了，能否马上参加比赛？','教学病例 · 虚构情境', '让学生先给出判断依据，再看下一页参考思路。病例不提供足够临床资料作诊断。重点是发现无痛步行不能代表专项功能恢复。')
box(s,.7,2.03,5.25,3.92,PALE);txt(s,'学生上周踝扭伤，\n走路已经基本不痛，\n单脚落地仍然晃动，\n现在想直接参加比赛。',1.02,2.55,4.65,2.65,27,NAVY,True)
explain(s,[('问题一','还缺少哪些功能与自我感知信息？'),('问题二','体育教师应如何安排后续参与？'),('问题三','哪些预防内容值得长期保留？')]);band(s,'请用“症状—功能—信心—专项要求”组织回答。')
#22
s=course_slide('病例解析：把“无痛”与“准备好”区分开','预防拓展 · PAASS / JOSPT 2021', '参考回答不是临床处方：需要评估疼痛反应、踝活动和力量、动态平衡、信心、专项任务及完整训练能力。体育教师按专业许可调整参与，回场决定由医疗/康复团队协同，持续平衡及适当外部支持可用于再伤预防。'+PAASS+'；'+ACPG)
cards(s,[('补充评估','活动度、力量、平衡\n信心与专项任务'),('调整参与','尚未稳定完成任务\n不直接升级到比赛'),('继续预防','平衡与功能训练\n必要护踝及负荷追踪')]);band(s,'回场依据功能与运动需求，不能仅按休息天数或走路是否疼痛。')
#23
s=course_slide('本课带走的三个判断','教材内容与预防拓展归纳', '请学生各用一句话解释：外侧踝扭伤的机制；反复崴脚为什么需要功能恢复；什么时候应停止参与并转介。最后把预防与保持运动参与联系起来。')
for i,(h,b) in enumerate([('机制','跖屈、内翻与跑跳负荷，分别影响哪些结构？'),('预防','平衡、力量、动作控制与合适支持共同参与。'),('参与','疼痛减轻不等于完全恢复；回场需要综合证据。')]):
 y=2.12+i*1.23;chip(s,h,.85,y,1.65);txt(s,b,2.95,y+.06,9.2,.75,26,INK,True)
band(s,'少一次受伤，更要少一次未恢复就重返运动的决定。')
#24
s=course_slide('教材与 GitHub 参考','资料说明', '教材医学主体来自所上传6页。GitHub课程实际查阅其第1周25页讲义；只参考情境、分步解释和课堂提问，不引用模型结论作医学依据。GitHub图鉴只参考观察与判断分开、联系回场的组织方式，未复制受保护的第三方视频和医学长文。')
refs(s,[('教材：踝足损伤的功能障碍与主观检查','用户提供《运动康复评定学-踝关节损伤》，p.164—169。'),('GitHub：从受伤情境提出学习问题','Sithranjan-Suresh/injury-risk-ml-curriculum · Week_1'),('GitHub：机制观察、证据与回场情境','keanu77/professional-sports-injury-atlas'),('GitHub：踝扭伤信息结构化记录','freshehrteam/Ankle-sprain · 症状与功能记录的组织思路')]);notes(s,s.notes_slide.notes_text_frame.text+'\n'+COURSE+'\n'+ATLAS+'\nhttps://github.com/freshehrteam/Ankle-sprain')
#25
s=course_slide('预防与回场：补充证据来源','预防拓展 · 参考文献', '教材未详细覆盖的预防及回场部分采用以下来源。PAASS属于共识框架，不是统一回场测试或保证。课堂任务示例为教学设计，不代替个体化康复方案。')
refs(s,[('JOSPT 2021：外侧踝扭伤与慢性踝不稳','Martin et al. Ankle Stability and Movement Coordination Impairments.\n'+ACPG),('BJSM 2021：PAASS 回场评估框架','Smith et al. Return to sport decisions after an acute lateral ankle sprain injury.\n'+PAASS),('本课使用范围','突出再伤预防、功能恢复与协作回场；练习不设置适用于所有人的固定剂量。')]);band(s,'完整来源、课堂提示和病例说明见各页“备注”。')
finish('Ankle_Sports_Injury_Prevention.pptx')
