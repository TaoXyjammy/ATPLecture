"""General sports-injury teaching decks; soccer appears only as an illustrative example.
Run: pip install python-pptx pillow; python source/build_general_courses.py
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
MEDIA=json.loads((ROOT/'General_MEDIA_SOURCES.json').read_text())
LABELS=json.loads((ROOT/'Chinese_labels.json').read_text())
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
 txt(s,'运动损伤评定 · '+kind,8.05,.22,4.65,.35,13,GRAY,True,PP_ALIGN.RIGHT);box(s,5,.59,7.7,.012,'DFE5EE')
 txt(s,title,.65,.94,12,.85,29,INK,True)
 box(s,0,7.07,13.333,.43,'F1F1F1');box(s,0,7.46,13.333,.04,'5579B2');txt(s,SECTION+'  |  '+src,.58,7.17,10.9,.25,9,GRAY)
 s.notes_slide.notes_text_frame.text=note+'\n内容依据：'+src+'。医学插图用于结构理解，不代表本页案例的诊断。'
 RECORDS.append({'slide':len(R.slides),'title':title,'condition':condition,'stage':stage,'source':src})
 return s

def add_note(s,t):s.notes_slide.notes_text_frame.text+='\n'+t

def explain(s,items):
 for i,(h,b) in enumerate(items):
  y=2.03+i*1.25;txt(s,f'{i+1:02d}',6.45,y,.55,.38,20,B,True);txt(s,h,7.1,y,5.5,.42,23,NAVY,True);txt(s,b,7.1,y+.47,5.48,.78,19)

def chinese_labels(s,name,left,top,width,height):
 # Document annotation only: masks and Chinese text are native PPT objects.
 # Leaders, anatomical pixels and all animation frames are preserved.
 im=Image.open(ASSET/name);iw,ih=im.size
 for item in LABELS.get(name,[]):
  x=left+item['x']/iw*width;y=top+item['y']/ih*height
  w=item['w']/iw*width;h=item['h']/ih*height
  box(s,x,y,w,h,item['fill'])
  q=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h))
  tf=q.text_frame;tf.word_wrap=False;tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=0
  from pptx.enum.text import MSO_ANCHOR
  tf.vertical_anchor=MSO_ANCHOR.MIDDLE
  p=tf.paragraphs[0];p.text=item['chinese'];p.alignment=PP_ALIGN.CENTER
  p.font.name=FONT;p.font.color.rgb=RGBColor.from_string(item['color'])
  p.font.size=Pt(item['size_px']/iw*width*72)
  p.space_after=Pt(0)
  add_note(s,'图内中文标注：'+item['english']+' → '+item['chinese'])

def picture(s,name,caption='解剖参考',x=.82,y=1.98,w=5.12,h=3.83):
 box(s,.65,1.88,5.42,4.25,'FFFFFF','DCE4F0')
 p=ASSET/name;im=Image.open(p);iw,ih=im.size;ratio=min(w/iw,h/ih);dw,dh=iw*ratio,ih*ratio
 q=s.shapes.add_picture(str(p),Inches(x+(w-dw)/2),Inches(y+(h-dh)/2),width=Inches(dw),height=Inches(dh))
 chinese_labels(s,name,x+(w-dw)/2,y+(h-dh)/2,dw,dh)
 m=MEDIA.get(name,{})
 creators={'knee_ligaments.jpg':'科学动画','meniscus.png':'布劳森医学','patella.png':'损伤图谱','knee_side.png':'布劳森医学','ankle.png':'维基共享资源','achilles.jpg':'布劳森医学','plantar.png':'损伤图谱','shin.png':'损伤图谱','fracture.png':'施维雅医学图像','foot_anatomy.png':'布劳森医学','tendinopathy.png':'医学插图','acl_match.gif':'沃尔登等，2015','ankle_frames.gif':'巴克索普等，2025','football_kick.gif':'马丁·古铁雷斯，2026'}
 license_name=m.get('license','')
 if license_name.startswith('CC BY-SA'):license_name='署名—相同方式共享'
 elif license_name.startswith('CC BY-NC'):license_name='署名—非商业性使用'
 elif license_name.startswith('CC BY'):license_name='署名许可'
 elif license_name=='Public domain':license_name='公有领域'
 txt(s,caption,.83,5.86,5.08,.18,8.0,GRAY)
 txt(s,'图源：'+creators.get(name,'来源见备注')+'｜'+license_name+'；原始署名及许可见备注',.83,6.03,5.08,.16,7.0,GRAY)
 add_note(s,'素材信息：'+json.dumps(m,ensure_ascii=False))
 return q

def pitch(s,label,kind='cut'):
 box(s,.65,1.88,5.42,4.25,'FFFFFF','DCE4F0');txt(s,'基础练习组织示例',.9,2.05,4.9,.4,20,NAVY,True)
 box(s,.95,2.68,4.8,2.58,'F0F5FA',B)
 points=[(1.45,4.67),(2.48,3.2),(3.52,4.65),(4.75,3.18)]
 if kind in ['balance','plantar','tendon']:points=[(1.55,4.1),(3.05,3.3),(4.75,4.1)]
 if kind=='contact':points=[(1.6,3.5),(2.65,4.45),(4.65,3.5)]
 for i,(x,y) in enumerate(points):
  box(s,x,y,.20,.20,B,kind=MSO_SHAPE.OVAL);txt(s,'甲乙丙丁'[i],x-.04,y-.36,.3,.3,14,NAVY,True)
  if i:
   px,py=points[i-1];line(s,px+.1,py+.1,x+.1,y+.1,TEAL,2)
 txt(s,label,.95,5.48,4.8,.45,18,TEAL,True)
 txt(s,'路线示例，可按年龄、能力与场地调整；非完整训练处方。',.83,5.94,5.06,.18,8.4,GRAY)

def cover(anatomy):
 s=slide(SECTION,'教材 p.'+PAGES,'按教材损伤种类顺序讲解。每种损伤内部依次讲损伤机制、诊断标准/线索、主要功能障碍；随后补充一般损伤预防，再进入下一种损伤。预防是教学拓展，不伪装成原教材内容。',kind='课程封面')
 txt(s,'常见运动损伤\n识别、评定与预防',.85,2.1,6.9,1.7,40,NAVY,True)
 txt(s,'按教材顺序逐项讲解\n每项损伤后紧接对应预防',.9,4.4,6.55,1.05,23,GRAY)
 # Use a detailed anatomical image, cropped only by proportional placement.
 p=ASSET/anatomy;im=Image.open(p);iw,ih=im.size;scale=min(4.4/iw,4.1/ih);dw,dh=iw*scale,ih*scale
 s.shapes.add_picture(str(p),Inches(8+(4.4-dw)/2),Inches(1.9+(4.1-dh)/2),width=Inches(dw),height=Inches(dh));add_note(s,'封面配图：'+json.dumps(MEDIA[anatomy],ensure_ascii=False))
 chinese_labels(s,anatomy,8+(4.4-dw)/2,1.9+(4.1-dh)/2,dw,dh)
 band(s,'讲解路径：损伤机制 → 诊断线索 → 主要功能障碍 → 损伤预防')
 return s

def roadmap(conditions):
 s=slide('教材讲解顺序：逐项学习，逐项配套预防','教材 p.'+PAGES,'本页只展示顺序。各损伤的预防跟随该项教材内容之后，不把所有预防集中放在课程末尾。',kind='课程内容')
 for i,d in enumerate(conditions):
  y=1.95+i*(.62 if len(conditions)==6 else .75);txt(s,d['num']+'、'+d['name'],.95,y,7.5,.43,23,NAVY,True);txt(s,'机制 → 诊断与功能 → 损伤预防',8.15,y+.06,4.3,.35,16,GRAY)
 band(s,'常见情境：跑跳、转向、负重、重复活动，以及跌倒或外力接触。')

def teach(d,bonus=None):
 name=d['num']+'、'+d['name'];p=d['page'];img=d['img']
 s=slide(name+'：损伤机制','教材 p.'+p,d['note'],condition=d['name'],stage='损伤机制');picture(s,img,d.get('caption','解剖参考'));explain(s,d['mechanism']);band(s,d['mechanism_key'])
 if bonus:bonus(d)
 s=slide(name+'：诊断与功能障碍','教材 p.'+p,d.get('diagnosis_note','教材列举的专项试验作为判断线索，需要病史、体征与必要影像相互印证；由专业人员检查。')+'讲解顺序为先诊断线索，再主要功能障碍。',condition=d['name'],stage='诊断与功能障碍');picture(s,img,d.get('caption','结构参考'));explain(s,d['diagnosis']);band(s,d['diagnosis_key'])
 s=slide(name+'：损伤预防','教材机制归纳 · 预防拓展',d['prevention_note']+'\n该页在该项教材内容后立即讲解。具体任务为一般练习组织示例，不能声称单一动作已被证明可独立预防此病。',kind='损伤预防',condition=d['name'],stage='损伤预防');pitch(s,d['drill'],d.get('drill_kind','cut'));explain(s,d['prevention']);band(s,d['prevention_key']);add_note(s,'拓展依据：'+d.get('evidence',KCPG)+'；预防任务为依据机制的教学归纳，具体安排应结合个体能力与活动要求。')

def acl_bonus(d):
 s=slide('一、前交叉韧带损伤：足球病例示例','教材 p.136 · Waldén et al., 2015','嵌入论文补充视频1的GIF，真实足球比赛中的研究病例。先观察接触、足部落地与方向变化，再讨论机制。论文研究中确认的病例不等于可以凭课堂影片自行确诊。此短片约1.9秒，GIF循环播放；PDF只显示静态帧。原文及视频：'+ACLSTUDY,kind='足球例证 · 动图',condition=d['name'],stage='损伤机制·动图')
 picture(s,'acl_match.gif','真实比赛录像 · 转GIF')
 explain(s,[('看事件','观察逼抢接近、支撑脚着地和身体转向。'),('看负荷','留意重心变化与膝部受力，避免只盯一个角度。'),('理解一般受力机制','减速、转向与跳跃落地，在多种活动中都需要稳定控制。')]);band(s,'提问：录像能说明什么动作？哪些诊断仍需要临床检查？')

def ankle_bonus(d):
 s=slide('一、踝关节扭伤：足球落地示例','教材 p.164 · Buckthorpe et al., 2025','原文图2(c1—c4)：争顶、落地接触对手脚、估计受伤帧、失衡后恢复。按原顺序制作GIF，是离散研究图帧，不是连续录像，也不是精确到毫秒的测量。论文及图2：'+ANKLESTUDY,kind='足球例证 · 研究图帧',condition=d['name'],stage='损伤机制·动图')
 picture(s,'ankle_frames.gif','研究图帧序列 · 非连续录像')
 explain(s,[('c1 → c2：落地条件','争顶后落地，受伤侧足部与对手脚发生接触。'),('c3：估计受伤帧','足踝受力条件突变；图帧不能单独确定分级。'),('c4：动作后果','失去平衡，随后用对侧恢复支撑。')]);band(s,'由示例理解机制：不稳定的支撑条件与外力干扰会改变踝部受力。')

def summary(conditions):
 s=slide('按教材回扣：把“损伤—活动情境—预防”连起来','教材 p.'+PAGES+' · 教学归纳','顺序与正文一致。让学生任选一项，说明受损结构、常见活动情境以及预防重点。不将一项动作观察当作诊断或伤病预测。',kind='课堂回顾')
 rows=[['损伤（教材顺序）','活动情境','对应预防重点']]+[[d['name'],d['summary_scene'],d['summary_prevent']] for d in conditions]
 widths=[3.3,4.25,4.48];rh=.56 if len(conditions)==6 else .64
 for i,row in enumerate(rows):
  x=.65
  for j,t in enumerate(row):
   box(s,x,1.96+i*rh,widths[j],rh,B if i==0 else ('EDF2FB' if i%2 else 'F7F9FC'),'FFFFFF');txt(s,t,x+.12,2.08+i*rh,widths[j]-.22,.42,17,'FFFFFF' if i==0 else INK,i==0 or j==0);x+=widths[j]
 band(s,'预防要落实到训练与组织；有急性痛肿、交锁、不稳或不能负重时及时评估。')

def reference_pages():
 s=slide('教材与一般预防拓展依据','教材与预防来源','教材内容与预防拓展分开标注。保持原文伤病顺序，不把课堂示例冒充完整的循证训练方案。',kind='参考资料')
 entries=[('教材内容','用户上传《运动康复评定学》'+TOPIC+'损伤相关页，p.'+PAGES+'。'),('预防指南',('JOSPT 2023：膝与 ACL 损伤的运动预防。\n'+KCPG) if TOPIC=='膝关节' else ('JOSPT 2021：外侧踝扭伤与慢性踝不稳。\n'+ACPG)),('一般活动管理','根据年龄、活动基础和既往损伤，安排渐进负荷、动作练习与安全环境。'),('应用方式','各页练习为一般教学任务示例；具体负荷与伤后回场需要个体化评估。')]
 for i,(h,b) in enumerate(entries):
  y=1.94+i*1.02;txt(s,h,.85,y,11.5,.4,21,NAVY,True);txt(s,b,.85,y+.42,11.5,.66,15,GRAY)
 add_note(s,'一般预防部分采用已列临床指南及教材损伤机制作教学归纳；不从单一项目研究推定所有项目的风险。')
 s=slide('中文解剖图与运动损伤示例来源','素材来源与播放说明','所有图片作者、许可、原始链接、加工方式记录于source/General_MEDIA_SOURCES.json。CC BY-SA素材的改编图继续遵循原许可；ACL论文补充录像为CC BY-NC 4.0，供非商业教学使用，第三方原始标识保留。其他素材按各自许可使用。',kind='素材与说明')
 entries=[('真实结构图','Blausen Medical、Scientific Animations、InjuryMap、Servier / Wikimedia Commons。'),('比赛受伤素材',('Waldén et al., BJSM 2015：真实足球 ACL 病例补充录像。\n'+ACLSTUDY) if TOPIC=='膝关节' else ('Buckthorpe et al., KSSTA 2025：足球踝扭伤图2(c1—c4)。\n'+ANKLESTUDY)),('足球示例的使用范围','比赛动图仅用于解释损伤机制，课程重点为通用评定与预防。'),('播放说明','GIF已嵌入PPT，放映时查看；PDF为静态预览。完整素材与来源随源码包提供。')]
 for i,(h,b) in enumerate(entries):
  y=1.94+i*1.02;txt(s,h,.85,y,11.5,.4,21,NAVY,True);txt(s,b,.85,y+.42,11.5,.66,15,GRAY)
 add_note(s,'本通用版使用原有中文标注和真实足球病例动图；动图作为一种活动情境示例，不据此推定其他活动的风险或确诊个人损伤。')

KNEE=[{'num': '一',
  'name': '前交叉韧带损伤',
  'page': '136',
  'img': 'knee_ligaments.jpg',
  'caption': '三维参考：①前交叉 ②后交叉 ③内侧副韧带 ④外侧副韧带',
  'mechanism': [('结构与作用', '前交叉韧带限制胫骨过度前移，参与膝关节旋转稳定。'),
                ('教材所述受力', '足部固定时膝部扭转、落地控制不良或膝过伸。'),
                ('常见活动情境', '急停、变向、跳跃落地；足球中的逼抢转向可作一例。')],
  'mechanism_key': '从关节姿态、支撑条件与受力方向理解损伤，不能只看运动项目名称。',
  'diagnosis': [('诊断线索：病史', '响声、剧痛、急性肿胀、不能继续运动；后期可打软腿。'),
                ('诊断线索：检查', '教材列 Lachman、前抽屉试验；必要时结合影像。'),
                ('主要功能障碍', '急性痛肿与活动受限；慢性可肌力下降、步态和不稳异常。')],
  'diagnosis_key': '打软腿、急性肿胀和运动受限需综合评估，也要关注合并损伤。',
  'prevention': [('练习减速与方向控制', '从低速、简单任务开始，逐步增加速度与动作复杂度。'),
                 ('结合力量与落地训练', '将下肢力量、髋躯干控制和跳跃落地纳入组合练习。'),
                 ('重视持续实施', '根据年龄、能力和既往损伤调整，并在日常训练中持续开展。')],
  'drill': '减速 → 变向 → 稳定停住',
  'prevention_key': '预防以持续的多组分训练为核心，单个姿势不能预测谁会受伤。',
  'note': '按上传教材的病种顺序及“损伤机制—诊断线索—主要功能障碍”组织。本通用版面向一般体育活动与相关日常功能；如涉及足球，仅作损伤情境示例。图为一般结构参考。',
  'prevention_note': '此页在该项损伤内容之后立即讲解。预防建议依据教材受力与功能特点作一般教学归纳；临床指南适用的结论按来源范围使用，不将足球病例直接推广为所有活动的风险。具体任务按年龄、能力、症状和活动需求调整。',
  'summary_scene': '急停、变向、跳跃落地',
  'summary_prevent': '减速与落地控制、组合训练',
  'evidence': 'https://doi.org/10.2519/jospt.2023.0301'},
 {'num': '二',
  'name': '半月板损伤',
  'page': '137',
  'img': 'meniscus.png',
  'caption': '三维半月板结构与撕裂示意',
  'mechanism': [('正常作用', '半月板分散负荷、吸收冲击，参与关节稳定与润滑。'),
                ('教材所述机制', '负重扭转、过度屈曲，使半月板受到挤压与剪切。'),
                ('联系活动实例', '转身、蹲起等负重任务；足球支撑脚转身可说明受力变化。')],
  'mechanism_key': '理解“负重 + 扭转”；撕裂类型、部位与血供影响恢复。',
  'diagnosis': [('诊断线索：症状与触诊', '交锁、卡顿、错动感；内外侧关节线压痛。'),
                ('诊断线索：专项检查', 'McMurray、Apley、Thessaly；\n必要时结合 MRI 综合判断。'),
                ('主要功能障碍', '急性痛肿、活动受限；慢性可在急停急转或屈膝时疼痛。')],
  'diagnosis_key': '卡顿与真正交锁要区分，专项试验不能单独确认撕裂。',
  'prevention': [('学习协调转身', '结合脚步与身体方向变化，减少足部固定时的强行扭转。'),
                 ('逐步增加负重需求', '由简单动作到较复杂任务，关注深屈膝、疲劳和负荷耐受。'),
                 ('及时报告机械性症状', '出现关节线痛、卡顿或交锁时调整活动，及时专业评估。')],
  'drill': '小步转身 → 控制方向变化',
  'prevention_key': '这些是依据机制的动作与负荷管理建议，不能保证避免所有半月板损伤。',
  'note': '按上传教材的病种顺序及“损伤机制—诊断线索—主要功能障碍”组织。本通用版面向一般体育活动与相关日常功能；如涉及足球，仅作损伤情境示例。图为一般结构参考。',
  'prevention_note': '此页在该项损伤内容之后立即讲解。预防建议依据教材受力与功能特点作一般教学归纳；临床指南适用的结论按来源范围使用，不将足球病例直接推广为所有活动的风险。具体任务按年龄、能力、症状和活动需求调整。',
  'evidence': '上传教材 p.137 的损伤机制与功能障碍；一般活动管理的教学归纳',
  'summary_scene': '负重转身、扭转、深屈膝',
  'summary_prevent': '协调步法、渐进负荷、症状反馈'},
 {'num': '三',
  'name': '副韧带损伤',
  'page': '138',
  'img': 'knee_ligaments.jpg',
  'caption': '三维结构参考：③内侧副韧带 ④外侧副韧带',
  'mechanism': [('内侧副韧带', '外翻应力牵拉内侧支持结构，可合并其他损伤。'),
                ('外侧副韧带', '内翻应力相关；还需关注邻近外侧结构。'),
                ('接触与意外受力', '侧向碰撞、跌倒等可改变膝部受力；足球对抗可作一例。')],
  'mechanism_key': '先说明膝关节如何变形，再判断哪一侧结构受到牵拉。',
  'diagnosis': [('诊断线索：局部表现', '内侧或外侧较表浅疼痛，急性期可肿胀与屈伸受限。'),
                ('诊断线索：检查', '教材列内翻、外翻应力试验；MRI明确相关损伤。'),
                ('主要功能障碍', '急性痛肿；慢性可肌力和活动度下降，行走运动不稳。')],
  'diagnosis_key': '疼痛侧、受力机制、侧向稳定与功能表现应互相印证。',
  'prevention': [('规范接触与活动规则', '在有接触的活动中学习身体控制，减少失控侧撞。'),
                 ('分层安排任务难度', '从无接触、可控任务开始，逐步提高速度或对抗需求。'),
                 ('关注接触后的异常', '侧方疼痛、肿胀或不稳时停止高风险活动，及时评估。')],
  'drill': '分区练习 → 可控接触 → 稳定控制',
  'drill_kind': 'contact',
  'prevention_key': '规则、空间与动作控制共同管理风险，不能将护具视为零受伤保证。',
  'note': '按上传教材的病种顺序及“损伤机制—诊断线索—主要功能障碍”组织。本通用版面向一般体育活动与相关日常功能；如涉及足球，仅作损伤情境示例。图为一般结构参考。',
  'prevention_note': '此页在该项损伤内容之后立即讲解。预防建议依据教材受力与功能特点作一般教学归纳；临床指南适用的结论按来源范围使用，不将足球病例直接推广为所有活动的风险。具体任务按年龄、能力、症状和活动需求调整。',
  'evidence': '上传教材 p.138 的损伤机制与功能障碍；一般活动管理的教学归纳',
  'summary_scene': '侧向碰撞、跌倒、异常内外翻',
  'summary_prevent': '规则与空间、控制接触、及时停止'},
 {'num': '四',
  'name': '髌骨脱位（髌股关节不稳）',
  'page': '138—139',
  'img': 'patella.png',
  'caption': '正常髌股关系与髌骨脱位医学插图',
  'mechanism': [('静态支持', '内侧髌股韧带等内侧支持结构限制髌骨过度外移。'),
                ('动态控制与个体因素', '股四头肌参与动态稳定；结构形态、排列和关节松弛也相关。'),
                ('常见活动情境', '突然转向、落地失衡或直接撞击；足球变向只是其中一例。')],
  'mechanism_key': '髌股不稳涉及结构与控制，不应只归因于某一块肌肉。',
  'diagnosis': [('诊断线索：病史', '不稳、发软、踏空感；询问既往脱位与反复发生情况。'),
                ('诊断线索：专项检查', '教材列髌骨外推、恐惧、倾斜试验。'),
                ('主要功能障碍', '疼痛、肿胀、卡顿和错动；上下楼或下蹲困难。')],
  'diagnosis_key': '急性疑似脱位需及时医疗评估，课堂不模拟复位或诱发实际脱位。',
  'prevention': [('练习下肢协同控制', '在基础力量和平衡任务中观察躯干、髋与膝的配合。'),
                 ('让转向与落地渐进', '从低速、稳定支撑任务过渡到更复杂的活动要求。'),
                 ('管理既往脱位史', '反复不稳者需专业评估，再决定参与程度与必要支持。')],
  'drill': '稳定支撑 → 小步换向',
  'prevention_key': '既往脱位者应结合结构、功能与活动需求，个体化安排参与。',
  'note': '按上传教材的病种顺序及“损伤机制—诊断线索—主要功能障碍”组织。本通用版面向一般体育活动与相关日常功能；如涉及足球，仅作损伤情境示例。图为一般结构参考。',
  'prevention_note': '此页在该项损伤内容之后立即讲解。预防建议依据教材受力与功能特点作一般教学归纳；临床指南适用的结论按来源范围使用，不将足球病例直接推广为所有活动的风险。具体任务按年龄、能力、症状和活动需求调整。',
  'evidence': '上传教材 p.138—139 的损伤机制与功能障碍；一般活动管理的教学归纳',
  'summary_scene': '转向、落地失衡或直接撞击',
  'summary_prevent': '协同控制、渐进任务、既往史管理'},
 {'num': '五',
  'name': '髌骨软化',
  'page': '139',
  'img': 'knee_side.png',
  'caption': '膝关节侧面结构参考（非软化病灶图）',
  'mechanism': [('教材所述结构改变', '髌股关节软骨相关病理改变，可与过度使用或外伤有关。'),
                ('常见负荷情境', '反复跑跳、屈膝、蹲起和上下楼，会改变髌股关节负荷。'),
                ('理解疼痛与病变', '膝前痛是症状，不能据此直接确认髌骨软化。')],
  'mechanism_key': '先理解髌股接触与负荷，再讨论膝前痛；不能把症状与结构病变等同。',
  'diagnosis': [('诊断线索：教材试验', 'Clarke、磨髌试验是教材列举的检查线索。'),
                ('诊断线索：综合判断', '必要时结合影像；诱发疼痛的试验不能单独确诊。'),
                ('主要功能障碍', '大量屈伸后疼痛；严重时下楼、下蹲等日常活动也受限。')],
  'diagnosis_key': '本页按教材病名讲解；图为结构参考，不是病灶或诊断结论。',
  'prevention': [('循序调整屈膝负荷', '关注近期活动变化，避免突然叠加大量跑跳、蹲起或上下楼。'),
                 ('保持力量与活动能力', '按能力安排髋膝力量和动作控制，关注疲劳与代偿。'),
                 ('记录疼痛反应', '结合活动中和次日症状调整；持续膝前痛要专业评估。')],
  'drill': '活动任务交替，逐步适应负重',
  'prevention_key': '负荷管理建议需要结合个人症状与功能，不能代替个体化诊疗。',
  'note': '按上传教材的病种顺序及“损伤机制—诊断线索—主要功能障碍”组织。本通用版面向一般体育活动与相关日常功能；如涉及足球，仅作损伤情境示例。图为一般结构参考。',
  'prevention_note': '此页在该项损伤内容之后立即讲解。预防建议依据教材受力与功能特点作一般教学归纳；临床指南适用的结论按来源范围使用，不将足球病例直接推广为所有活动的风险。具体任务按年龄、能力、症状和活动需求调整。',
  'evidence': '上传教材 p.139 的损伤机制与功能障碍；一般活动管理的教学归纳',
  'summary_scene': '反复屈膝、跑跳、蹲起与上下楼',
  'summary_prevent': '渐进负荷、力量控制、疼痛追踪'}]

ANKLE=[{'num': '一',
  'name': '踝关节扭伤',
  'page': '164—165',
  'img': 'ankle.png',
  'caption': '踝外侧韧带与跟腱结构参考',
  'mechanism': [('教材所述机制', '跖屈位内翻是典型情境，外侧韧带较常受累。'),
                ('常见活动情境', '行走踩空、跑跳落地、不平整地面或接触后失衡。'),
                ('需关注的结构', '距腓前韧带常受累；外翻、外旋还可能涉及其他组织。')],
  'mechanism_key': '损伤可发生于运动或日常活动；不能把每次崴脚都当作轻度扭伤。',
  'diagnosis': [('诊断线索：症状与触诊', '外踝前下方肿痛；关注内外踝及足部骨性压痛。'),
                ('诊断线索：检查与分级', '教材列前抽屉、距骨倾斜试验和I—III级；必要X线。'),
                ('主要功能障碍', '急性痛肿、活动和负重受限；可遗留本体感觉与平衡不足。')],
  'diagnosis_key': '不能只按肿胀大小现场分级；持续不稳与反复扭伤应复评。',
  'prevention': [('练习平衡与姿势控制', '由安全单脚支撑到触点或方向变化，逐步增加任务难度。'),
                 ('结合力量与落地控制', '关注踝周肌力、下肢协同和支撑面的稳定条件。'),
                 ('减少反复扭伤风险', '既往扭伤者评估适当支持与功能训练，结合活动需求复评。')],
  'drill': '单脚稳定 → 触点 → 方向变化',
  'drill_kind': 'balance',
  'prevention_key': '踝扭伤预防：平衡与功能训练、适当支持，以及安全环境。',
  'note': '按上传教材的病种顺序及“损伤机制—诊断线索—主要功能障碍”组织。本通用版面向一般体育活动与相关日常功能；如涉及足球，仅作损伤情境示例。图为一般结构参考。',
  'prevention_note': '此页在该项损伤内容之后立即讲解。预防建议依据教材受力与功能特点作一般教学归纳；临床指南适用的结论按来源范围使用，不将足球病例直接推广为所有活动的风险。具体任务按年龄、能力、症状和活动需求调整。',
  'evidence': 'https://doi.org/10.2519/jospt.2021.0302',
  'summary_scene': '踩空、不平地面、跑跳落地',
  'summary_prevent': '平衡、支撑控制、再伤管理'},
 {'num': '二',
  'name': '跟腱病变',
  'page': '165—166',
  'img': 'tendinopathy.png',
  'caption': '跟腱相关疼痛医学插图',
  'mechanism': [('教材所述机制', '反复跳跃、蹬地，小腿三头肌反复收缩牵拉跟腱。'),
                ('相关功能与环境因素', '踝活动度、肌力耐力、地面及鞋等因素影响组织需求。'),
                ('联系常见活动', '跑步、跳跃或突然增加的练习量；足球连续起动可作一例。')],
  'mechanism_key': '跟腱病变常是渐进的负荷相关问题，应与突然断裂分开理解。',
  'diagnosis': [('诊断线索：症状过程', '晨僵、局部痛；教材按活动后、活动中和持续疼痛描述程度。'),
                ('诊断线索：检查', '局部压痛、相关跟腱试验；可见提踵力量与耐力下降。'),
                ('主要功能障碍', '休息后或晨起痛、局部增厚；跑跳、蹬地及穿鞋可不适。')],
  'diagnosis_key': '晨僵与提踵能力变化值得追踪，不应一律解释为普通运动酸痛。',
  'prevention': [('逐步增加跑跳需求', '调整活动强度、频次和持续时间，避免同时突然增加。'),
                 ('培养小腿力量与耐力', '按能力渐进安排提踵等任务，关注动作质量与耐受。'),
                 ('让症状参与活动决策', '跟踪晨起和次日反应；持续不适时减少诱发活动并评估。')],
  'drill': '一般活动与蹬地任务分层安排',
  'drill_kind': 'tendon',
  'prevention_key': '跟腱疼痛后的活动调整，应结合专业评估、症状反应与负荷耐受。',
  'note': '按上传教材的病种顺序及“损伤机制—诊断线索—主要功能障碍”组织。本通用版面向一般体育活动与相关日常功能；如涉及足球，仅作损伤情境示例。图为一般结构参考。',
  'prevention_note': '此页在该项损伤内容之后立即讲解。预防建议依据教材受力与功能特点作一般教学归纳；临床指南适用的结论按来源范围使用，不将足球病例直接推广为所有活动的风险。具体任务按年龄、能力、症状和活动需求调整。',
  'evidence': '上传教材 p.165—166 的损伤机制与功能障碍；一般活动管理的教学归纳',
  'summary_scene': '反复跑跳、蹬地与活动量突增',
  'summary_prevent': '渐进负荷、小腿能力、症状追踪'},
 {'num': '三',
  'name': '跟腱断裂',
  'page': '166',
  'img': 'achilles.jpg',
  'caption': '跟腱断裂三维医学插图',
  'mechanism': [('教材所述机制', '直接或间接暴力、小腿肌群强烈收缩及跟腱过度牵拉。'),
                ('常见活动情境', '突然加速、蹬地起跳，或足踝被动过度背伸。'),
                ('与跟腱病变区分', '这是急性结构损伤；既往病变可存在，但不是每例都有。')],
  'mechanism_key': '突发“被踢感”、响声与蹬地无力，应首先停止活动并评估。',
  'diagnosis': [('诊断线索：事件与局部', '后踝突发痛、响声、肿胀或凹陷；踮脚和蹬地困难。'),
                ('诊断线索：专业检查', '教材列Thompson征；超声或MRI辅助检查。'),
                ('主要功能障碍', '跖屈蹬地与单脚踮脚能力明显受损，步行和跑跳受到影响。')],
  'diagnosis_key': '仍能部分行走，也不能排除跟腱断裂；现场不强行提踵或拉伸验证。',
  'prevention': [('逐渐进入高需求活动', '从一般准备过渡到快速起动、跳跃或较大负荷任务。'),
                 ('管理持续后踝症状', '疼痛、晨僵或蹬地能力下降时，及时调整并专业评估。'),
                 ('避免发生后的二次损伤', '急性异常立即停止；恢复活动应结合功能与实际需求判断。')],
  'drill': '一般准备 → 渐进起动',
  'drill_kind': 'tendon',
  'prevention_key': '准备活动、症状识别和及时停止有实际意义，但不能保证避免所有断裂。',
  'note': '按上传教材的病种顺序及“损伤机制—诊断线索—主要功能障碍”组织。本通用版面向一般体育活动与相关日常功能；如涉及足球，仅作损伤情境示例。图为一般结构参考。',
  'prevention_note': '此页在该项损伤内容之后立即讲解。预防建议依据教材受力与功能特点作一般教学归纳；临床指南适用的结论按来源范围使用，不将足球病例直接推广为所有活动的风险。具体任务按年龄、能力、症状和活动需求调整。',
  'evidence': '上传教材 p.166 的损伤机制与功能障碍；一般活动管理的教学归纳',
  'summary_scene': '突然起动、起跳、异常牵拉',
  'summary_prevent': '渐进准备、症状管理、保护与转介'},
 {'num': '四',
  'name': '足底筋膜炎',
  'page': '167',
  'img': 'plantar.png',
  'caption': '足底筋膜相关足跟痛医学插图',
  'mechanism': [('教材所述机制', '长期反复牵拉与过度使用，使足底负重组织反复受应力。'),
                ('相关条件', '跑跳、长时间站立、足部力学及踝活动等因素。'),
                ('联系生活与运动', '步行或久站增加、运动跑量变化，都可能改变足底负荷。')],
  'mechanism_key': '足跟痛可影响日常与运动；本课沿用教材病名，实际仍需鉴别来源。',
  'diagnosis': [('诊断线索：疼痛规律', '足跟内侧或足底近跟骨痛；休息后第一步常明显。'),
                ('诊断线索：检查', '局部压痛与相关牵拉试验等线索，需结合鉴别。'),
                ('主要功能障碍', '负重与长时间行走痛；活动初期缓解后又可能复发。')],
  'diagnosis_key': '“第一步痛 + 负重反应 + 疼痛位置”，比只说“脚跟痛”更有信息。',
  'prevention': [('逐步适应负重变化', '避免步行、站立或跑跳总量突然增加。'),
                 ('关注全日负荷与环境', '将工作、生活和运动负重一起考虑，留出恢复并选择合适鞋具。'),
                 ('持续疼痛及时调整', '减少反复诱发任务；必要支持应结合症状、舒适度与专业评估。')],
  'drill': '任务轮换，逐步适应站立与行走',
  'drill_kind': 'plantar',
  'prevention_key': '鞋、鞋垫或某个练习都不是适用于所有人的统一预防方案。',
  'note': '按上传教材的病种顺序及“损伤机制—诊断线索—主要功能障碍”组织。本通用版面向一般体育活动与相关日常功能；如涉及足球，仅作损伤情境示例。图为一般结构参考。',
  'prevention_note': '此页在该项损伤内容之后立即讲解。预防建议依据教材受力与功能特点作一般教学归纳；临床指南适用的结论按来源范围使用，不将足球病例直接推广为所有活动的风险。具体任务按年龄、能力、症状和活动需求调整。',
  'evidence': '上传教材 p.167 的损伤机制与功能障碍；一般活动管理的教学归纳',
  'summary_scene': '久站、步行增加、足底反复负重',
  'summary_prevent': '渐进适应、全日负荷、足跟反馈'},
 {'num': '五',
  'name': '胫骨内侧应力综合征',
  'page': '168',
  'img': 'shin.png',
  'caption': '胫骨内侧疼痛区域医学插图',
  'mechanism': [('教材所述机制', '活动量增加过快、反复跑跳，使小腿相关组织过负荷。'),
                ('可能相关条件', '下肢排列、足部受力、踝柔韧性、肌群平衡与地面。'),
                ('常见活动情境', '初次跑步、恢复训练或地面变化后，留意胫骨内侧痛。')],
  'mechanism_key': '活动后的胫骨内侧痛需要追踪，也要注意与骨应力损伤鉴别。',
  'diagnosis': [('诊断线索：时间与部位', '训练中后沿胫骨内侧疼痛；训练增强时可加重。'),
                ('诊断线索：检查与鉴别', '局部触痛及相关检查；持续、局灶或休息痛需排查骨损伤。'),
                ('主要功能障碍', '跑跳时疼痛，休息可缓解；严重时影响训练与行走。')],
  'diagnosis_key': '和跟腱痛区分位置，也要关注症状是否持续或加重。',
  'prevention': [('避免多种负荷同步突增', '跑量、跳跃次数、强度与频次分别渐进调整。'),
                 ('逐步适应场地与任务', '地面或活动方式改变时，结合小腿和足踝能力安排练习。'),
                 ('记录并报告疼痛变化', '关注部位、持续时间与加重趋势，不硬顶持续或局灶疼痛。')],
  'drill': '跑走或基础任务交替，分层加量',
  'prevention_key': '不存在适用于所有人的固定“安全加量百分比”。',
  'note': '按上传教材的病种顺序及“损伤机制—诊断线索—主要功能障碍”组织。本通用版面向一般体育活动与相关日常功能；如涉及足球，仅作损伤情境示例。图为一般结构参考。',
  'prevention_note': '此页在该项损伤内容之后立即讲解。预防建议依据教材受力与功能特点作一般教学归纳；临床指南适用的结论按来源范围使用，不将足球病例直接推广为所有活动的风险。具体任务按年龄、能力、症状和活动需求调整。',
  'evidence': '上传教材 p.168 的损伤机制与功能障碍；一般活动管理的教学归纳',
  'summary_scene': '跑跳加量、恢复活动、地面变化',
  'summary_prevent': '避免同步突增、逐步适应、疼痛报告'},
 {'num': '六',
  'name': '踝关节骨折',
  'page': '169',
  'img': 'fracture.png',
  'caption': '踝部骨折医学插图（非实际病例影像）',
  'mechanism': [('教材所述机制', '直接或间接暴力，受力方向和大小不同可形成不同骨折。'),
                ('运动与日常情境', '跌倒、坠落、碰撞、足部固定后的扭转或异常落地。'),
                ('结构影响', '骨性稳定受破坏，可能合并韧带和其他软组织损伤。')],
  'mechanism_key': '外伤后的踝部痛肿不能仅按扭伤处理，需关注骨折可能。',
  'diagnosis': [('诊断线索：病史与局部', '外伤、明显痛肿、骨性部位压痛；可有畸形或瘀斑。'),
                ('诊断线索：影像', '教材以X线等检查辅助明确；不现场诱发骨擦感来验证。'),
                ('主要功能障碍', '踝活动、负重与行走明显受限；可能无法站立行走。')],
  'diagnosis_key': '不能负重、畸形、开放伤或足部麻木发冷，需要及时医疗评估。',
  'prevention': [('改善环境与活动组织', '检查地面、障碍和活动空间，减少跌倒与失控碰撞风险。'),
                 ('按能力选择任务与装备', '循序增加难度，根据地面和活动要求选择合适鞋具或保护。'),
                 ('做好疑似骨折后的保护', '立即停止活动、保护并转介；不强行走动、扳动或复位。')],
  'drill': '留出活动与缓冲空间',
  'drill_kind': 'contact',
  'prevention_key': '预防既包括减少危险暴力，也包括发生后避免二次损伤。',
  'note': '按上传教材的病种顺序及“损伤机制—诊断线索—主要功能障碍”组织。本通用版面向一般体育活动与相关日常功能；如涉及足球，仅作损伤情境示例。图为一般结构参考。',
  'prevention_note': '此页在该项损伤内容之后立即讲解。预防建议依据教材受力与功能特点作一般教学归纳；临床指南适用的结论按来源范围使用，不将足球病例直接推广为所有活动的风险。具体任务按年龄、能力、症状和活动需求调整。',
  'evidence': '上传教材 p.169 的损伤机制与功能障碍；一般活动管理的教学归纳',
  'summary_scene': '跌倒、坠落、碰撞与异常扭转',
  'summary_prevent': '环境与任务适配、保护与转介'}]

def build(topic,chapter,section,pages,data,filename):
 global R,TOPIC,CHAPTER,SECTION,PAGES,RECORDS
 R=Presentation();R.slide_width=Inches(13.333);R.slide_height=Inches(7.5);TOPIC=topic;CHAPTER=chapter;SECTION=section;PAGES=pages;RECORDS=[]
 cover('knee_side.png' if topic=='膝关节' else 'foot_anatomy.png');roadmap(data)
 for i,d in enumerate(data):
  bonus=(acl_bonus if topic=='膝关节' else ankle_bonus) if i==0 else None
  teach(d,bonus)
 summary(data);reference_pages()
 for i,s in enumerate(R.slides):
  txt(s,f'{i+1:02d} / {len(R.slides):02d}',11.65,7.14,1.05,.3,12,GRAY,False,PP_ALIGN.RIGHT)
  for sh in s.shapes:
   st=sh._element.find('{http://schemas.openxmlformats.org/presentationml/2006/main}style')
   if st is not None:
    ef=st.find('{http://schemas.openxmlformats.org/drawingml/2006/main}effectRef')
    if ef is not None:ef.set('idx','0')
 R.core_properties.title=section+'（通用版）';R.core_properties.author='谢淘';R.core_properties.subject='按教材顺序：每项损伤后紧接预防；中文结构图与示例动图'
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
  assert stages[0]=='损伤机制' and stages[-1]=='损伤预防' and stages[-2]=='诊断与功能障碍',stages
 print(filename,len(R.slides),'slides; textbook order verified')

if __name__=='__main__':
 build('膝关节','第二章  膝关节功能评定','膝关节损伤的功能障碍','136—139',KNEE,'Knee_General_Injury_Prevention.pptx')
 build('踝足','第三章  踝关节功能评定','踝足损伤的功能障碍','164—169',ANKLE,'Ankle_General_Injury_Prevention.pptx')
