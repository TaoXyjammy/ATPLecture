"""Blue-white football teaching decks, following the supplied textbook order.
Run: pip install python-pptx pillow; python source/build_courses.py
All embedded media are under source/assets; attribution is in MEDIA_SOURCES.json.
"""
from pathlib import Path
import json,re
from PIL import Image
from pptx import Presentation
from pptx.util import Inches,Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE,MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN
ROOT=Path(__file__).resolve().parent;OUT=ROOT.parent;ASSET=ROOT/'assets'
MEDIA=json.loads((ASSET/'MEDIA_SOURCES.json').read_text())
B='3C5FA4';NAVY='243F83';PALE='EDF2FB';INK='20242A';GRAY='68707A';RED='B13A47';TEAL='327C80';FONT='Noto Sans CJK SC'
KCPG='https://doi.org/10.2519/jospt.2023.0301';ACPG='https://doi.org/10.2519/jospt.2021.0302'
FIFA='https://www.fifatrainingcentre.com/en/practice/grassroots/grassroots-and-youth-football-essentials/grassroots-coaching-essentials/optimising-warm-ups-u12-to-u15.php'
ACLSTUDY='https://doi.org/10.1136/bjsports-2014-094573';ANKLESTUDY='https://doi.org/10.1002/ksa.70049'
GITHUB='https://github.com/convderover/broadcast-kick-biomechanics'
R=None;TOPIC='';CHAPTER='';SECTION='';PAGES='';RECORDS=[]

def box(s,x,y,w,h,c,line=None,kind=MSO_SHAPE.RECTANGLE):
 q=s.shapes.add_shape(kind,Inches(x),Inches(y),Inches(w),Inches(h));q.fill.solid();q.fill.fore_color.rgb=RGBColor.from_string(c);q.line.fill.background()
 if line:q.line.color.rgb=RGBColor.from_string(line);q.line.width=Pt(1)
 return q

def txt(s,t,x,y,w,h,size=22,c=INK,b=False,align=None):
 q=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h));tf=q.text_frame;tf.word_wrap=True;tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=0
 for i,line in enumerate(t.split('\n')):
  p=tf.paragraphs[0] if i==0 else tf.add_paragraph();p.text=line;p.font.name=FONT;p.font.size=Pt(size);p.font.bold=b;p.font.color.rgb=RGBColor.from_string(c);p.space_after=Pt(7)
  if align is not None:p.alignment=align
 return q

def line(s,x,y,x2,y2,c=B,width=2):
 q=s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Inches(x),Inches(y),Inches(x2),Inches(y2));q.line.color.rgb=RGBColor.from_string(c);q.line.width=Pt(width);return q

def band(s,t):
 box(s,.65,6.28,12.03,.54,PALE,'AABBD7');txt(s,t,.85,6.39,11.6,.35,18,NAVY,True)

def slide(title,src,note='',kind='教材内容',condition=None,stage=None):
 s=R.slides.add_slide(R.slide_layouts[6]);s.background.fill.solid();s.background.fill.fore_color.rgb=RGBColor.from_string('FFFFFF')
 for j,c in enumerate(['344F91','4C73BC','126B97','2FA6CF','DA343A','E3E8F1']):box(s,.22+j*.29,.13,.18,.18,c)
 box(s,.22,.48,.19,.19,'91C64D',kind=MSO_SHAPE.OVAL);txt(s,CHAPTER,.51,.45,6,.35,14,'314E88',True)
 txt(s,'足球专项 · '+kind,8.05,.22,4.65,.35,13,GRAY,True,PP_ALIGN.RIGHT);box(s,5,.59,7.7,.012,'DFE5EE')
 txt(s,title,.65,.94,12,.85,29,INK,True)
 box(s,0,7.07,13.333,.43,'F1F1F1');box(s,0,7.46,13.333,.04,'5579B2');txt(s,SECTION+'  |  '+src,.58,7.17,10.9,.25,9,GRAY)
 s.notes_slide.notes_text_frame.text=note+'\n内容依据：'+src+'。医学插图用于结构理解，不代表本页案例的诊断。'
 RECORDS.append({'slide':len(R.slides),'title':title,'condition':condition,'stage':stage,'source':src})
 return s

def add_note(s,t):s.notes_slide.notes_text_frame.text+='\n'+t

def explain(s,items):
 for i,(h,b) in enumerate(items):
  y=2.03+i*1.25;txt(s,f'{i+1:02d}',6.45,y,.55,.38,20,B,True);txt(s,h,7.1,y,5.5,.42,23,NAVY,True);txt(s,b,7.1,y+.47,5.48,.78,19)

def picture(s,name,caption='解剖参考',x=.82,y=1.98,w=5.12,h=3.83):
 box(s,.65,1.88,5.42,4.25,'FFFFFF','DCE4F0')
 p=ASSET/name;im=Image.open(p);iw,ih=im.size;ratio=min(w/iw,h/ih);dw,dh=iw*ratio,ih*ratio
 q=s.shapes.add_picture(str(p),Inches(x+(w-dw)/2),Inches(y+(h-dh)/2),width=Inches(dw),height=Inches(dh))
 m=MEDIA.get(name,{})
 creators={'knee_ligaments.jpg':'Scientific Animations','meniscus.png':'BruceBlaus','patella.png':'InjuryMap','knee_side.png':'Blausen Medical','ankle.png':'Jak / Wikimedia','achilles.jpg':'BruceBlaus','plantar.png':'InjuryMap','shin.png':'InjuryMap','fracture.png':'Servier Medical Art','foot_anatomy.png':'Blausen Medical','tendinopathy.png':'Medical illustration','acl_match.gif':'Waldén et al., 2015','ankle_frames.gif':'Buckthorpe et al., 2025','football_kick.gif':'R. Martín Gutiérrez, 2026'}
 credit=f'{caption}  |  {creators.get(name,"来源见备注")}  |  {m.get("license","")}'
 credit=credit.replace('（论文及补充材料；非商业教学）','').replace('（原文）','').replace('（仓库 docs/ 与 resultados/ 中的图文视频）','')
 txt(s,credit,.83,5.91,5.08,.2,8.1,GRAY)
 add_note(s,'素材信息：'+json.dumps(m,ensure_ascii=False))
 return q

def pitch(s,label,kind='cut'):
 box(s,.65,1.88,5.42,4.25,'FFFFFF','DCE4F0');txt(s,'足球训练组织示例',.9,2.05,4.9,.4,20,NAVY,True)
 box(s,.95,2.68,4.8,2.58,'F0F5FA',B);line(s,3.35,2.68,3.35,5.26,'AFBFDB',1);box(s,2.95,3.55,.8,.8,'F0F5FA','AFBFDB',MSO_SHAPE.OVAL)
 points=[(1.45,4.67),(2.48,3.2),(3.52,4.65),(4.75,3.18)]
 if kind in ['balance','plantar','tendon']:points=[(1.55,4.1),(3.05,3.3),(4.75,4.1)]
 if kind=='contact':points=[(1.6,3.5),(2.65,4.45),(4.65,3.5)]
 for i,(x,y) in enumerate(points):
  box(s,x,y,.20,.20,B,kind=MSO_SHAPE.OVAL);txt(s,chr(65+i),x-.04,y-.36,.3,.3,14,NAVY,True)
  if i:
   px,py=points[i-1];line(s,px+.1,py+.1,x+.1,y+.1,TEAL,2)
 box(s,4.95,4.7,.22,.22,'FFFFFF',INK,MSO_SHAPE.OVAL)
 txt(s,label,.95,5.48,4.8,.45,18,TEAL,True)
 txt(s,'路线示例，可按年龄、能力与场地调整；非完整训练处方。',.83,5.94,5.06,.18,8.4,GRAY)

def cover(anatomy):
 s=slide(SECTION,'教材 p.'+PAGES,'按教材损伤种类顺序讲解。每种损伤内部依次讲损伤机制、诊断标准/线索、主要功能障碍；随后补充足球专项预防，再进入下一种损伤。预防是教学拓展，不伪装成原教材内容。',kind='课程封面')
 txt(s,'足球运动中的\n损伤识别与预防',.85,2.1,6.9,1.7,40,NAVY,True)
 txt(s,'按教材顺序逐项讲解\n每项损伤后紧接对应预防',.9,4.4,6.55,1.05,23,GRAY)
 # Use a detailed anatomical image, cropped only by proportional placement.
 p=ASSET/anatomy;im=Image.open(p);iw,ih=im.size;scale=min(4.4/iw,4.1/ih);dw,dh=iw*scale,ih*scale
 s.shapes.add_picture(str(p),Inches(8+(4.4-dw)/2),Inches(1.9+(4.1-dh)/2),width=Inches(dw),height=Inches(dh));add_note(s,'封面配图：'+json.dumps(MEDIA[anatomy],ensure_ascii=False))
 band(s,'讲解路径：损伤机制 → 诊断线索 → 主要功能障碍 → 足球中的预防')
 return s

def roadmap(conditions):
 s=slide('教材讲解顺序：逐项学习，逐项配套预防','教材 p.'+PAGES,'本页只展示顺序。各损伤的预防跟随该项教材内容之后，不把所有预防集中放在课程末尾。',kind='课程内容')
 for i,d in enumerate(conditions):
  y=1.95+i*(.62 if len(conditions)==6 else .75);txt(s,d['num']+'、'+d['name'],.95,y,7.5,.43,23,NAVY,True);txt(s,'机制 → 诊断与功能 → 足球预防',8.15,y+.06,4.3,.35,16,GRAY)
 band(s,'贯穿情境：逼抢变向、支撑脚转身、争顶落地、冲刺蹬地与对抗。')

def teach(d,bonus=None):
 name=d['num']+'、'+d['name'];p=d['page'];img=d['img']
 s=slide(name+'：损伤机制','教材 p.'+p,d['note'],condition=d['name'],stage='损伤机制');picture(s,img,d.get('caption','解剖参考'));explain(s,d['mechanism']);band(s,d['mechanism_key'])
 if bonus:bonus(d)
 s=slide(name+'：诊断与功能障碍','教材 p.'+p,d.get('diagnosis_note','教材列举的专项试验作为判断线索，需要病史、体征与必要影像相互印证；由专业人员检查。')+'讲解顺序为先诊断线索，再主要功能障碍。',condition=d['name'],stage='诊断与功能障碍');picture(s,img,d.get('caption','结构参考'));explain(s,d['diagnosis']);band(s,d['diagnosis_key'])
 s=slide(name+'：足球中的预防','教材机制归纳 · 预防拓展',d['prevention_note']+'\n该页在该项教材内容后立即讲解。具体任务为教练组织示例，不能声称单一动作已被证明可独立预防此病。',kind='足球专项预防',condition=d['name'],stage='足球预防');pitch(s,d['drill'],d.get('drill_kind','cut'));explain(s,d['prevention']);band(s,d['prevention_key']);add_note(s,'拓展依据：'+d.get('evidence',KCPG)+'；FIFA基础训练组织：'+FIFA)

def acl_bonus(d):
 s=slide('一、前交叉韧带损伤：比赛录像观察','教材 p.136 · Waldén et al., 2015','嵌入论文补充视频1的GIF，真实足球比赛中的研究病例。先观察接触、足部落地与方向变化，再讨论机制。论文研究中确认的病例不等于可以凭课堂影片自行确诊。此短片约1.9秒，GIF循环播放；PDF只显示静态帧。原文及视频：'+ACLSTUDY,kind='真实比赛 GIF',condition=d['name'],stage='损伤机制·动图')
 picture(s,'acl_match.gif','真实比赛录像 · 转GIF')
 explain(s,[('看事件','观察逼抢接近、支撑脚着地和身体转向。'),('看负荷','留意重心变化与膝部受力，避免只盯一个角度。'),('回到足球训练','把减速、换向与争顶落地作为控制训练任务。')]);band(s,'提问：录像能说明什么动作？哪些诊断仍需要临床检查？')

def ankle_bonus(d):
 s=slide('一、踝关节扭伤：争顶落地的受伤图帧','教材 p.164 · Buckthorpe et al., 2025','原文图2(c1—c4)：争顶、落地接触对手脚、估计受伤帧、失衡后恢复。按原顺序制作GIF，是离散研究图帧，不是连续录像，也不是精确到毫秒的测量。论文及图2：'+ANKLESTUDY,kind='真实比赛图帧 GIF',condition=d['name'],stage='损伤机制·动图')
 picture(s,'ankle_frames.gif','研究图帧序列 · 非连续录像')
 explain(s,[('c1 → c2：落地条件','争顶后落地，受伤侧足部与对手脚发生接触。'),('c3：估计受伤帧','足踝受力条件突变；图帧不能单独确定分级。'),('c4：动作后果','失去平衡，随后用对侧恢复支撑。')]);band(s,'足球要点：争顶落地空间、接触干扰与踝部控制需要一起关注。')

def summary(conditions):
 s=slide('按教材回扣：把“损伤—足球情境—预防”连起来','教材 p.'+PAGES+' · 教学归纳','顺序与正文一致。让学生任选一项，说明受损结构、足球中的典型任务以及预防重点。不将一项动作观察当作诊断或伤病预测。',kind='课堂回顾')
 rows=[['损伤（教材顺序）','足球情境','对应预防重点']]+[[d['name'],d['summary_scene'],d['summary_prevent']] for d in conditions]
 widths=[3.3,4.25,4.48];rh=.56 if len(conditions)==6 else .64
 for i,row in enumerate(rows):
  x=.65
  for j,t in enumerate(row):
   box(s,x,1.96+i*rh,widths[j],rh,B if i==0 else ('EDF2FB' if i%2 else 'F7F9FC'),'FFFFFF');txt(s,t,x+.12,2.08+i*rh,widths[j]-.22,.42,17,'FFFFFF' if i==0 else INK,i==0 or j==0);x+=widths[j]
 band(s,'预防要落实到训练与组织；有急性痛肿、交锁、不稳或不能负重时及时评估。')

def reference_pages():
 s=slide('教材与足球专项拓展依据','教材与预防来源','教材内容与预防拓展分开标注。保持原文伤病顺序，不把课堂示例冒充完整的循证训练方案。',kind='参考资料')
 entries=[('教材内容','用户上传《运动康复评定学》'+TOPIC+'损伤相关页，p.'+PAGES+'。'),('预防指南',('JOSPT 2023：膝与 ACL 损伤的运动预防。\n'+KCPG) if TOPIC=='膝关节' else ('JOSPT 2021：外侧踝扭伤与慢性踝不稳。\n'+ACPG)),('足球训练组织','FIFA Training Centre：渐进热身、技术衔接与年龄适配。'),('应用方式','各页练习是足球教学示例；具体负荷与伤后回场需要个体化评估。')]
 for i,(h,b) in enumerate(entries):
  y=1.94+i*1.02;txt(s,h,.85,y,11.5,.4,21,NAVY,True);txt(s,b,.85,y+.42,11.5,.66,15,GRAY)
 add_note(s,'FIFA来源：'+FIFA)
 s=slide('解剖图、真实动图与 GitHub 来源','素材来源与播放说明','所有图片作者、许可、原始链接、加工方式记录于source/assets/MEDIA_SOURCES.json。CC BY-SA素材的改编图继续遵循原许可；ACL论文补充录像为CC BY-NC 4.0，供非商业教学使用，第三方原始标识保留。其他素材按各自许可使用。',kind='素材与说明')
 entries=[('真实结构图','Blausen Medical、Scientific Animations、InjuryMap、Servier / Wikimedia Commons。'),('比赛受伤素材',('Waldén et al., BJSM 2015：真实足球 ACL 病例补充录像。\n'+ACLSTUDY) if TOPIC=='膝关节' else ('Buckthorpe et al., KSSTA 2025：足球踝扭伤图2(c1—c4)。\n'+ANKLESTUDY)),('GitHub 实际检索与正常踢球动图','convderover/broadcast-kick-biomechanics · 作者本人踢球，CC BY 4.0。'),('播放说明','GIF已嵌入PPT，放映时查看；PDF为静态预览。完整素材与来源随源码包提供。')]
 for i,(h,b) in enumerate(entries):
  y=1.94+i*1.02;txt(s,h,.85,y,11.5,.4,21,NAVY,True);txt(s,b,.85,y+.42,11.5,.66,15,GRAY)
 add_note(s,'实际查阅GitHub：'+GITHUB+'；其正常踢球视频不是受伤录像，也不能用其模型结果进行个人损伤预测。另查阅MonaElkhouly/Biomechanics-Final-Project-ACL-Injuries-for-female-football-player-Analysi-，未将未明确授权的素材重新分发。')

KNEE=[
 dict(num='一',name='前交叉韧带损伤',page='136',img='knee_ligaments.jpg',caption='三维结构参考：1 ACL；2 PCL；3 MCL；4 LCL',
 mechanism=[('足球中的典型情境','逼抢急停、变向或争顶落地，常见非接触性机制。'),('教材所述受力','落地不良姿势、足部固定时扭转，或膝过伸。'),('结构与作用','ACL限制胫骨过度前移，参与膝关节旋转稳定。')],mechanism_key='足球中的“脚已着地、身体仍转向”，需要关注支撑与减速控制。',
 diagnosis=[('诊断线索：病史','响声、剧痛、急性肿胀、不能继续运动；后期可打软腿。'),('诊断线索：检查','教材列 Lachman、前抽屉试验；必要时结合影像。'),('主要功能障碍','急性痛肿与活动受限；慢性可肌力下降、步态和不稳异常。')],diagnosis_key='打软腿、急性肿胀和运动受限需综合评估，也要关注合并损伤。',
 prevention=[('练减速与变向','低速绕标起步，逐步增加速度与持球干扰，反馈支撑控制。'),('练力量与落地','将髋膝力量、近端控制、跳跃落地纳入组合训练。'),('把预防持续做下去','季前开始并在赛季保持；不要只在比赛前偶尔做一次。')],drill='绕标减速 → 方向变化 → 接球',prevention_key='ACL预防以持续的多组分训练为核心，单个姿势不能预测谁会受伤。',
 note='按教材一、前交叉韧带损伤之（一）损伤机制讲解。足球研究提示逼抢、踢球后恢复平衡、争顶落地等情境，录像用于机制讨论，不能代替诊断。'+ACLSTUDY,prevention_note='JOSPT2023支持运动预防计划；足球任务是其原理的教学转化，不是一套经本课独立验证的新方案。',summary_scene='逼抢急停、争顶落地',summary_prevent='减速与落地 + 组合训练'),
 dict(num='二',name='半月板损伤',page='137',img='meniscus.png',caption='三维半月板结构与撕裂示意',
 mechanism=[('正常作用','分散负荷、吸收冲击，参与稳定与润滑。'),('教材所述机制','负重扭转、过度屈曲；半月板受挤压与剪切。'),('联系足球支撑脚','带球转身或踢球时，支撑足固定而身体旋转会改变受力。')],mechanism_key='理解“负重 + 扭转”；撕裂类型、部位与血供影响恢复。',
 diagnosis=[('诊断线索：症状与触诊','交锁、卡顿、错动感；内外侧关节线压痛。'),('诊断线索：专项检查','McMurray、Apley、Thessaly；\n必要时结合 MRI 综合判断。'),('主要功能障碍','急性痛肿、活动受限；慢性可在急停急转或屈膝时疼痛。')],diagnosis_key='卡顿与真正交锁要区分，专项试验不能单独确认撕裂。',
 prevention=[('把转身拆成步法','从连续小步换向开始，减少足被固定时强行扭转身体。'),('逐步增加任务难度','无对抗带球 → 方向变化 → 可控对抗，关注疲劳下动作。'),('持续追踪症状','出现关节线痛、卡顿或交锁时调整训练，及时专业评估。')],drill='小步转身 → 带球换向',prevention_key='本页为依据机制的足球训练建议，不把某一练习当作半月板损伤预防保证。',
 note='按教材半月板损伤之损伤机制讲解。图示结构为参考，不表示某个真实球员的撕裂形态。',prevention_note='教材机制导向归纳，配合足球教练组织原则。既往损伤者由医疗与康复专业人员协同决定参与。',evidence=FIFA,summary_scene='支撑脚转身、负重扭转',summary_prevent='步法调整、渐进对抗、症状反馈'),
 dict(num='三',name='副韧带损伤',page='138',img='knee_ligaments.jpg',caption='三维结构参考：3 MCL；4 LCL',
 mechanism=[('内侧副韧带','外翻应力牵拉内侧支持结构，可合并其他损伤。'),('外侧副韧带','内翻应力相关；还需关注邻近外侧结构。'),('足球中的接触情境','铲抢或侧向碰撞时，膝关节姿态与外力方向共同作用。')],mechanism_key='先说明膝关节如何变形，再判断哪一侧结构被牵拉。',
 diagnosis=[('诊断线索：局部表现','内侧或外侧较表浅疼痛，急性期可肿胀与屈伸受限。'),('诊断线索：检查','教材列内翻、外翻应力试验；MRI明确相关损伤。'),('主要功能障碍','急性痛肿；慢性可肌力和活动度下降，行走运动不稳。')],diagnosis_key='疼痛侧、受力机制、侧向稳定与功能表现应互相印证。',
 prevention=[('规范对抗与抢断','练合法接近与身体控制，限制失控侧撞与危险铲抢。'),('分层增加对抗强度','无对抗技术 → 可控接触 → 比赛情境，保持足够间距。'),('接触后及时识别','侧方疼痛、肿胀或不稳时停止对抗，避免继续强行变向。')],drill='分区对抗，先技术后强度',drill_kind='contact',prevention_key='规范对抗能管理接触风险，不能把护具或力量训练当作零受伤保证。',
 note='教材副韧带损伤。足球情境用于说明接触受力，不把某一种铲抢外观直接等同于某条韧带撕裂。',prevention_note='足球教学组织与教材机制归纳。对抗升级按年龄、技术和身体条件进行。',evidence=FIFA,summary_scene='侧向碰撞、铲抢对抗',summary_prevent='规范对抗、渐进接触、及时停止'),
 dict(num='四',name='髌骨脱位（髌股关节不稳）',page='138—139',img='patella.png',caption='正常髌股关系与髌骨脱位医学插图',
 mechanism=[('静态支持','MPFL等内侧支持结构限制髌骨过度外移。'),('动态控制与个体因素','股内侧斜肌参与稳定；结构形态与下肢排列也可能相关。'),('联系足球动作','变向或落地失衡、直接撞击后，可能出现髌股不稳事件。')],mechanism_key='髌股不稳涉及结构与控制，不应只归因于某一块肌肉。',
 diagnosis=[('诊断线索：病史','不稳、发软、踏空感；询问既往脱位与反复发生情况。'),('诊断线索：专项检查','教材列髌骨外推、恐惧、倾斜试验。'),('主要功能障碍','疼痛、肿胀、卡顿和错动；上下楼或下蹲困难。')],diagnosis_key='急性疑似脱位需及时医疗评估，课堂不模拟复位或诱发实际脱位。',
 prevention=[('练髋膝协同控制','在基础力量与稳定任务中观察躯干、髋膝配合。'),('让变向与落地渐进','从低速、可控任务过渡到足球专项，避免失控叠加对抗。'),('关注既往脱位者','不以“不痛了”判断恢复；参与与必要支持需个体评估。')],drill='可控接球 → 小步换向',prevention_key='反复髌股不稳者需要专业评估，再决定足球参与与训练进阶。',
 note='按教材髌骨脱位损伤机制讲解，预防页是下一页之后单独展开。图片为一般医学插图，不是个人诊断。',prevention_note='教材机制及FIFA组织原则的教学转化，不声称孤立股内侧肌练习能预防全部髌骨脱位。',evidence=FIFA,summary_scene='变向、落地失衡或撞击',summary_prevent='髋膝控制、渐进任务、既往史管理'),
 dict(num='五',name='髌骨软化',page='139',img='knee_side.png',caption='膝关节侧面结构参考（非软化病灶图）',
 mechanism=[('教材所述结构改变','髌股关节软骨相关病理改变；可与过度使用或外伤有关。'),('足球中的负荷情境','反复跑跳、减速和屈膝任务，增加髌股关节负荷需求。'),('理解边界','膝前痛是症状，不能据此直接确认髌骨软化。')],mechanism_key='先理解髌股接触与负荷，再讨论膝前痛；不能把二者简单画等号。',
 diagnosis=[('诊断线索：教材试验','Clarke、磨髌试验是教材列举的检查线索。'),('诊断线索：综合判断','必要时结合影像；诱发疼痛的试验不能单独确诊。'),('主要功能障碍','大量屈伸后疼痛；严重时下楼、下蹲等日常活动也受限。')],diagnosis_key='本页按教材病名讲解；图为结构参考，不是病灶或诊断结论。',
 prevention=[('调整跑跳与屈膝负荷','关注近期训练密度，避免突然叠加大量急停、跳跃和深蹲。'),('保持下肢力量与耐力','按能力安排髋膝力量与动作控制，留意疲劳和代偿。'),('追踪疼痛反应','记录训练中及次日反应；持续膝前痛要调整并评估。')],drill='分站练习，交替技术与体能任务',prevention_key='这些是负荷管理与教学建议，不能用来替代髌股疼痛的个体化诊疗。',
 note='教材髌骨软化，侧面结构图用于定位髌骨与关节关系。未把正常结构图误标为病灶图。',prevention_note='基于教材过度使用机制的教学归纳；足球组织参考FIFA渐进原则。',evidence=FIFA,summary_scene='反复跑跳、急停、屈膝负荷',summary_prevent='渐进负荷、力量耐力、疼痛追踪')
]

ANKLE=[
 dict(num='一',name='踝关节扭伤',page='164—165',img='ankle.png',caption='踝外侧韧带与跟腱结构参考',
 mechanism=[('教材所述机制','跖屈位内翻是典型情境，外侧韧带较常受累。'),('足球中的典型事件','争顶落地踩到对手脚、变向踩空，或抢球接触后失衡。'),('需关注的结构','距腓前韧带常受累；外翻、外旋还可能涉及其他组织。')],mechanism_key='不是每次“崴脚”都只是轻度外侧扭伤，还要注意骨折和其他损伤。',
 diagnosis=[('诊断线索：症状与触诊','外踝前下方肿痛；关注内外踝及足部骨性压痛。'),('诊断线索：检查与分级','教材列前抽屉、距骨倾斜试验和I—III级；必要X线。'),('主要功能障碍','急性痛肿、活动和负重受限；可遗留本体感觉与平衡不足。')],diagnosis_key='不能只按肿胀大小现场分级；持续不稳与反复扭伤应复评。',
 prevention=[('平衡与接传球结合','从安全单脚支撑到触点、接传球，逐步增加任务难度。'),('力量与落地控制','关注踝周肌力、下肢协同及争顶后的落地空间。'),('管理再伤风险','既往扭伤者评估护踝与训练；回场要结合专项功能与信心。')],drill='单脚稳定 → 触点 → 接传球',drill_kind='balance',prevention_key='踝扭伤预防：功能训练 + 适当支持 + 场地与对抗组织。',
 note='按教材踝关节扭伤（一）损伤机制讲解，足球研究用于情境补充。'+ANKLESTUDY,prevention_note='JOSPT2021支持本体感觉/平衡训练和适当预防性支具降低复发风险；根据功能短板个体化安排。回场还需专业判断。',evidence=ACPG,summary_scene='争顶落地、抢球接触、踩空',summary_prevent='平衡、落地控制、再伤管理'),
 dict(num='二',name='跟腱病变',page='165—166',img='tendinopathy.png',caption='跟腱相关疼痛医学插图',
 mechanism=[('教材所述机制','反复跳跃、蹬地，小腿三头肌反复收缩牵拉跟腱。'),('相关负荷因素','踝活动度、肌力与耐力改变，以及地面和训练环境。'),('足球中的负荷情境','短冲刺、连续起动和争顶次数增加，恢复不足时需留意。')],mechanism_key='跟腱病变常是渐进的负荷相关问题，与突然断裂分开理解。',
 diagnosis=[('诊断线索：症状过程','晨僵、局部痛；教材按活动后、活动中和持续疼痛描述程度。'),('诊断线索：检查','局部压痛、相关跟腱试验；可见提踵力量与耐力下降。'),('主要功能障碍','休息后或晨起痛、局部增厚；跑跳、蹬地及穿鞋可不适。')],diagnosis_key='晨僵与提踵能力变化值得追踪，不应一律解释为普通运动酸痛。',
 prevention=[('逐步增加高需求动作','冲刺、跳跃与比赛负荷不要同时突然增加。'),('安排小腿能力训练','按能力渐进进行提踵及相关力量耐力任务，监督动作质量。'),('让症状参与训练决策','跟踪晨起和次日反应，持续不适时减少诱发任务并评估。')],drill='技术练习与蹬地负荷分层安排',drill_kind='tendon',prevention_key='跟腱疼痛后的足球训练调整，要与专业评估和负荷耐受相结合。',
 note='教材二、跟腱病变。病变不应被简单理解为只有炎症；足球负荷为教学联系。',prevention_note='根据教材损伤机制与功能障碍作负荷管理归纳，不提供适用于所有人的固定提踵处方。',evidence=FIFA,summary_scene='冲刺、连续起动、反复争顶',summary_prevent='渐进负荷、小腿能力、症状追踪'),
 dict(num='三',name='跟腱断裂',page='166',img='achilles.jpg',caption='跟腱断裂三维医学插图',
 mechanism=[('教材所述机制','直接或间接暴力；小腿肌群强烈收缩、跟腱过度牵拉。'),('足球中的典型任务','突然加速、蹬地起跳，或足踝被动过度背伸。'),('与跟腱病变区分','这是急性结构损伤；既往病变可存在，但不是每例都有。')],mechanism_key='突发“被踢感”、响声与蹬地无力，应首先停止运动并评估。',
 diagnosis=[('诊断线索：事件与局部','后踝突发痛、响声、肿胀或凹陷；踮脚和蹬地困难。'),('诊断线索：专业检查','教材列Thompson征；超声或MRI辅助检查。'),('主要功能障碍','跖屈蹬地与单脚踮脚能力明显受损，步行和跑跳受到影响。')],diagnosis_key='仍能部分行走，也不能排除跟腱断裂；现场不强行提踵或拉伸验证。',
 prevention=[('充分准备再进入高速任务','从一般活动逐步过渡到冲刺、起跳和专项对抗。'),('管理持续后踝症状','疼痛、晨僵或蹬地能力下降时，及时调整并专业评估。'),('避免二次损伤','出现急性异常立即停止；回场按恢复功能与足球需求判断。')],drill='渐进起动，避免直接进入全速对抗',drill_kind='tendon',prevention_key='不能承诺完全预防断裂；准备、症状识别和及时停止都有实际意义。',
 note='按教材三、跟腱断裂依次讲解。图为病理示意，不是球员影像。',prevention_note='本页为教材机制与足球训练组织的归纳。不得声称热身或单一力量动作可确保防止所有断裂。',evidence=FIFA,summary_scene='突然起动、蹬地起跳',summary_prevent='渐进准备、症状管理、及时停止'),
 dict(num='四',name='足底筋膜炎',page='167',img='plantar.png',caption='足底筋膜相关足跟痛医学插图',
 mechanism=[('教材所述机制','长期反复牵拉与过度使用，足底负重组织反复受应力。'),('相关条件','跑跳或长时间站立、足部力学和踝活动等因素。'),('足球中的负荷情境','训练跑量、球场适应与站立时间共同影响足底需求。')],mechanism_key='本课沿用教材病名；实际足跟痛仍需区分不同来源。',
 diagnosis=[('诊断线索：疼痛规律','足跟内侧或足底近跟骨痛；休息后第一步常明显。'),('诊断线索：检查','局部压痛与相关牵拉试验等线索，需结合鉴别。'),('主要功能障碍','负重与长时间行走痛；活动初期缓解后又可能复发。')],diagnosis_key='“第一步痛 + 负重反应 + 疼痛位置”，比只说“脚跟痛”更有信息。',
 prevention=[('逐步适应训练与地面','避免跑量、场地和训练鞋同时突然改变。'),('管理足底负重总量','关注球场外久站与行走，留出恢复并记录足跟反应。'),('出现持续痛及时调整','减少反复诱发任务；鞋与必要支持按舒适度和专业评估选择。')],drill='分站轮换，减少单一任务连续负重',drill_kind='plantar',prevention_key='不把某种鞋或鞋垫当作适用于所有球员的足底筋膜炎预防方案。',
 note='按教材四、足底筋膜炎讲解。图示区域是教学参考，不能用痛点自我确诊。',prevention_note='教材过度使用机制与足球日常训练管理的教学归纳。',evidence=FIFA,summary_scene='跑量增加、足底反复负重',summary_prevent='渐进适应、总负荷与足跟反馈'),
 dict(num='五',name='胫骨内侧应力综合征',page='168',img='shin.png',caption='胫骨内侧疼痛区域医学插图',
 mechanism=[('教材所述机制','训练量增加过快、反复跑跳及小腿相关组织过负荷。'),('可能相关条件','下肢排列、足部受力、踝柔韧性、肌群平衡及地面。'),('联系足球训练','季前加量、反复折返跑或场地转换后，留意小腿内侧痛。')],mechanism_key='小腿内侧训练痛不应一概当作“练开就好”，也需鉴别骨应力损伤。',
 diagnosis=[('诊断线索：时间与部位','训练中后沿胫骨内侧疼痛；训练增强时可加重。'),('诊断线索：检查与鉴别','局部触痛及相关检查；持续、局灶或休息痛需排查骨损伤。'),('主要功能障碍','跑跳时疼痛，休息可缓解；严重时影响训练与行走。')],diagnosis_key='和跟腱痛区分位置，也要关注症状是否持续或加重。',
 prevention=[('控制加量的叠加','跑量、折返次数、强度与比赛密度不要同步突增。'),('渐进适应场地','硬地或场地变化时调整跑跳需求，配合小腿与足踝能力训练。'),('建立报告与复评机制','记录疼痛位置与变化，持续或加重时停止硬顶训练并评估。')],drill='折返跑与技术站交替，分层加量',prevention_key='不存在适用于所有球员的固定“安全加量百分比”。',
 note='按教材五、胫骨内侧应力综合征讲解；未把教材单个足纵弓数值作为通用预测阈值。',prevention_note='基于教材训练量与地面等机制的足球训练归纳，非统一负荷公式。',evidence=FIFA,summary_scene='季前加量、折返跑、地面变化',summary_prevent='避免同步突增、适应地面、报告疼痛'),
 dict(num='六',name='踝关节骨折',page='169',img='fracture.png',caption='踝部骨折医学插图（非实际球员影像）',
 mechanism=[('教材所述机制','间接暴力或直接暴力，受力方向和大小不同可形成不同骨折。'),('足球中的风险情境','抢球碰撞、足被固定时扭转，或落地时受异常外力。'),('结构影响','骨性稳定受破坏，可能合并韧带和其他软组织损伤。')],mechanism_key='不能因为发生在踢球时，或看起来像扭伤，就排除骨折。',
 diagnosis=[('诊断线索：病史与局部','外伤、明显痛肿、骨性部位压痛；可有畸形或瘀斑。'),('诊断线索：影像','教材以X线等检查辅助明确；不现场诱发骨擦感来验证。'),('主要功能障碍','踝活动、负重与行走明显受限；可能无法站立行走。')],diagnosis_key='不能负重、畸形、开放伤或足部麻木发冷，需要及时医疗评估。',
 prevention=[('规范抢断与对抗','训练规则、接近路线和空间管理，减少危险接触。'),('检查场地与装备','留意坑洼和障碍；按场地选鞋，护腿板不能排除踝部骨折。'),('把现场保护做好','疑似骨折立即停训、保护和转介；不强行走动、扳动或复位。')],drill='划分对抗区与缓冲空间',drill_kind='contact',prevention_key='骨折预防还包括减少危险接触，以及发生后避免二次损伤。',
 note='教材六、踝关节骨折。该图是一般解剖病理插图，具体分型和治疗不在体育课堂展开。',prevention_note='教材暴力机制与FIFA安全组织的教学归纳。发生急性异常时按伤情联系急救或医疗团队。',evidence=FIFA,summary_scene='抢球碰撞、固定足扭转、异常落地',summary_prevent='规则与空间、场地检查、保护转介')
]

def build(topic,chapter,section,pages,data,filename):
 global R,TOPIC,CHAPTER,SECTION,PAGES,RECORDS
 R=Presentation();R.slide_width=Inches(13.333);R.slide_height=Inches(7.5);TOPIC=topic;CHAPTER=chapter;SECTION=section;PAGES=pages;RECORDS=[]
 cover('knee_side.png' if topic=='膝关节' else 'foot_anatomy.png');roadmap(data)
 for i,d in enumerate(data):
  bonus=(acl_bonus if topic=='膝关节' else ankle_bonus) if i==0 else None
  teach(d,bonus)
 summary(data);reference_pages()
 # One small normal-action GIF on the cover's picture, with an explicit caption; anatomy remains the main graphic.
 s=R.slides[0]
 s.shapes.add_picture(str(ASSET/'football_kick.gif'),Inches(6.65),Inches(3.93),width=Inches(1.6),height=Inches(1.78))
 txt(s,'正常踢球 · 非受伤录像',6.6,5.76,2.0,.23,8,GRAY,False,PP_ALIGN.CENTER)
 add_note(s,'封面正常踢球GIF：'+json.dumps(MEDIA['football_kick.gif'],ensure_ascii=False))
 for i,s in enumerate(R.slides):
  txt(s,f'{i+1:02d} / {len(R.slides):02d}',11.65,7.14,1.05,.3,12,GRAY,False,PP_ALIGN.RIGHT)
  for sh in s.shapes:
   st=sh._element.find('{http://schemas.openxmlformats.org/presentationml/2006/main}style')
   if st is not None:
    ef=st.find('{http://schemas.openxmlformats.org/drawingml/2006/main}effectRef')
    if ef is not None:ef.set('idx','0')
 R.core_properties.title=section+'（足球专项）';R.core_properties.author='谢淘';R.core_properties.subject='按教材顺序：每项损伤后紧接预防；真实结构图与GIF'
 R.save(OUT/filename)
 (OUT/(filename.replace('.pptx','_讲解备注.md'))).write_text('\n\n'.join(f'## {i+1}. {RECORDS[i]["title"]}\n\n{s.notes_slide.notes_text_frame.text}' for i,s in enumerate(R.slides)))
 (OUT/(filename.replace('.pptx','_顺序核对.json'))).write_text(json.dumps(RECORDS,ensure_ascii=False,indent=2))
 # Verify that each injury's three textbook/prevention stages remain consecutive.
 expected=[d['name'] for d in data];observed=[]
 for r in RECORDS:
  if r['condition'] and (not observed or r['condition']!=observed[-1]):observed.append(r['condition'])
 assert observed==expected,(observed,expected)
 for d in data:
  stages=[r['stage'] for r in RECORDS if r['condition']==d['name']]
  assert stages[0]=='损伤机制' and stages[-1]=='足球预防' and stages[-2]=='诊断与功能障碍',stages
 print(filename,len(R.slides),'slides; textbook order verified')

if __name__=='__main__':
 build('膝关节','第二章  膝关节功能评定','膝关节损伤的功能障碍','136—139',KNEE,'Knee_Sports_Injury_Prevention.pptx')
 build('踝足','第三章  踝关节功能评定','踝足损伤的功能障碍','164—169',ANKLE,'Ankle_Sports_Injury_Prevention.pptx')
