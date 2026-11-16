from pptx import Presentation
from pptx.util import Inches,Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN
from pathlib import Path
R=Presentation();R.slide_width=Inches(13.333);R.slide_height=Inches(7.5)
B='3C5FA4'; NAVY='243F83'; PALE='EDF2FB'; INK='20242A'; GRAY='68707A'; RED='B13A47'; ORANGE='D78C3F'; TEAL='287F8A'; FONT='Noto Sans CJK SC'
def box(s,x,y,w,h,c,line=None,kind=MSO_SHAPE.RECTANGLE):
 sh=s.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h));sh.fill.solid();sh.fill.fore_color.rgb=RGBColor.from_string(c);sh.line.fill.background()
 if line:sh.line.color.rgb=RGBColor.from_string(line);sh.line.width=Pt(1.3)
 return sh
def txt(s,t,x,y,w,h,size=22,c=INK,b=False,align=None):
 sh=s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h));tf=sh.text_frame;tf.word_wrap=True
 tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=0
 for i,line in enumerate(t.split('\n')):
  p=tf.paragraphs[0] if i==0 else tf.add_paragraph();p.text=line;p.font.name=FONT;p.font.size=Pt(size);p.font.bold=b;p.font.color.rgb=RGBColor.from_string(c);p.space_after=Pt(9)
  if align is not None:p.alignment=align
 return sh
def line(s,x,y,x2,y2,c=B,width=2):
 q=s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y), Inches(x2), Inches(y2));q.line.color.rgb=RGBColor.from_string(c);q.line.width=Pt(width);return q
def arrow(s,x,y,w=.65,c=B,down=False):
 return box(s,x,y,w,.28,c,kind=MSO_SHAPE.DOWN_ARROW if down else MSO_SHAPE.RIGHT_ARROW)
def band(s,t):
 box(s,.65,6.23,12.03,.57,PALE,'AABBD7');txt(s,t,.85,6.35,11.6,.4,19,NAVY,True)
def slide(title,src='136—139',note=''):
 s=R.slides.add_slide(R.slide_layouts[6]);s.background.fill.solid();s.background.fill.fore_color.rgb=RGBColor.from_string('FFFFFF')
 for j,c in enumerate(['344F91','4C73BC','126B97','2FA6CF','DA343A','E3E8F1']):box(s,.22+j*.29,.13,.18,.18,c)
 box(s,.22,.48,.19,.19,'91C64D',kind=MSO_SHAPE.OVAL);txt(s,'第二章  膝关节功能评定',.51,.45,6,.35,14,'314E88',True)
 txt(s,'第一节  膝关节损伤的功能障碍',8.1,.22,4.6,.35,13,GRAY,True,PP_ALIGN.RIGHT);box(s,5,.59,7.7,.012,'DFE5EE')
 txt(s,title,.65,.94,12,.8,30,INK,True)
 box(s,0,7.07,13.333,.43,'F1F1F1');box(s,0,7.46,13.333,.04,'5579B2');txt(s,f'运动康复评定学  |  教材 p.{src}',.58,7.17,10,.25,10,GRAY)
 s.notes_slide.notes_text_frame.text=note+'\n教材依据：用户提供《运动康复评定学-膝关节损伤》，印刷页码 p.'+src+'。示意图为重绘教学简图，不表示真实比例。'
 return s
def panel(s):box(s,.65,1.95,5.35,4.0,'F7F9FC','DCE4F0')
def explain(s,items,x=6.45,y=2.03,w=6.1):
 for i,(head,body) in enumerate(items):
  yy=y+i*1.25;txt(s,f'{i+1:02d}',x,yy,.55,.4,20,B,True);txt(s,head,x+.65,yy,w-.65,.4,23,NAVY,True);txt(s,body,x+.65,yy+.48,w-.65,.75,19,INK)
def label(s,t,x,y,w=2,c=B):txt(s,t,x,y,w,.4,17,c,True)
def knee(s,x=1.5,y=2.25,mode='all'):
 # Frontal conceptual schematic, editable bones and soft-tissue markers.
 box(s,x+1.0,y,.75,1.15,'D8E2EF',B,MSO_SHAPE.ROUNDED_RECTANGLE)
 for dx in [.65,1.45]:box(s,x+dx,y+.85,.85,.65,'D8E2EF',B,MSO_SHAPE.OVAL)
 box(s,x+.72,y+1.73,1.55,.42,'D8E2EF',B,MSO_SHAPE.ROUNDED_RECTANGLE);box(s,x+1.12,y+2.08,.72,.75,'D8E2EF',B,MSO_SHAPE.ROUNDED_RECTANGLE)
 line(s,x+.8,y+1.62,x+1.35,y+1.62,TEAL,7);line(s,x+1.6,y+1.62,x+2.15,y+1.62,TEAL,7)
 if mode in ['all','acl']:
  line(s,x+1.1,y+1.05,x+1.9,y+1.94,RED,5);line(s,x+1.9,y+1.1,x+1.15,y+1.92,'A0B5D5',4)
 if mode in ['all','collateral']:
  line(s,x+.58,y+1.03,x+.58,y+2.12,ORANGE,5);line(s,x+2.4,y+1.03,x+2.4,y+2.12,ORANGE,5)
 if mode=='patella':
  box(s,x+1.12,y+.85,.75,1.0,'EFC4C8',RED,MSO_SHAPE.OVAL)
 txt(s,'股骨',x+2.5,y+.1,1,.4,17,GRAY);txt(s,'胫骨',x+2.5,y+2.6,1,.4,17,GRAY)
 txt(s,'结构关系示意 · 非比例解剖图',.95,5.57,4.8,.3,11,GRAY)
def chip(s,t,x,y,w,c=B):box(s,x,y,w,.6,c);txt(s,t,x+.1,y+.11,w-.2,.4,20,'FFFFFF',True,PP_ALIGN.CENTER)
def flow(s,labels,y=2.55):
 for i,t in enumerate(labels):
  x=.8+i*3.18;box(s,x,y,2.72,1.05,PALE,'AABBD7');txt(s,t,x+.16,y+.22,2.4,.7,23,NAVY,True,PP_ALIGN.CENTER)
  if i<len(labels)-1:arrow(s,x+2.8,y+.37,.3)
# 1
s=slide('膝关节损伤的功能障碍',note='开场直接进入课程：膝关节损伤不仅涉及某一结构，还会改变负重、活动和运动表现。本节以五类常见损伤为线索，学习从机制推测受损结构，再用检查与功能表现核对。')
txt(s,'从损伤机制\n到功能评定',.85,2.05,7,1.8,43,NAVY,True);txt(s,'前交叉韧带 · 半月板 · 副韧带\n髌股关节不稳 · 髌骨软化',.9,4.32,6.7,1.1,24,GRAY)
knee(s,8.25,2.15,'all');band(s,'课堂主线：哪里受伤 → 如何判断 → 哪些功能受到影响')
# 2
s=slide('同样是“膝痛”，评定从哪里开始？',note='提出两个情境，请学生先描述损伤动作而不是立即给诊断。A提示需要关注ACL及合并损伤，B提示髌股关节相关问题。情境均为教学编写，不能据此直接确诊。')
for x,n,t,body in [(.75,'A','急停变向后','膝内“啪”的一声，迅速肿胀，\n无法继续比赛。'),(6.85,'B','反复上下楼后','膝前方疼痛，下蹲不适，\n逐渐影响日常活动。')]:
 box(s,x,2.1,5.72,3.42,PALE);txt(s,n,x+.25,2.35,.8,.8,48,B,True);txt(s,t,x+1.2,2.5,4.1,.5,27,NAVY,True);txt(s,body,x+.3,3.6,5.15,1.25,24)
band(s,'先问：受伤动作、疼痛部位、肿胀过程、卡锁或打软腿。')
#3
s=slide('结构不同，承担的稳定任务也不同',note='左图仅用于建立结构位置感。ACL重点限制胫骨过度前移并参与旋转稳定；半月板分散负荷、吸收冲击；副韧带提供内外侧稳定。髌骨与滑车配合，帮助伸膝。ACL与PCL在图中交叉不代表真实投影。')
panel(s);knee(s,1.45,2.2);label(s,'前交叉韧带',.95,2.1,2,RED);line(s,2.05,2.55,2.9,3.7,RED);label(s,'半月板',.95,4.75,1.5,TEAL);line(s,2,4.75,2.5,3.88,TEAL)
explain(s,[('前交叉韧带','限制胫骨过度前移，参与旋转稳定。'),('半月板与副韧带','前者分散负荷；后者提供内、外侧稳定。'),('髌股关节','髌骨参与伸膝；轨迹异常可伴疼痛与不稳。')]);band(s,'评定要把结构受损与活动受限联系起来。')
#4
s=slide('把零散症状组织成一条评定链',note='此页为基于教材的教学归纳，非教材原有检查量表。先明确损伤动作与症状，再观察疼痛、肿胀、活动度，结合有针对性的检查，最终落实到步行、下蹲和运动功能。明显急性疼痛、严重肿胀或疑似骨折时不强行进行诱发试验。')
flow(s,['问病史','查基本功能','选专项检查','看活动影响'])
for x,t in [(.95,'动作与接触方式\n疼痛与肿胀时间'),(4.13,'疼痛、肿胀\n活动度、肌力、负重'),(7.31,'针对疑似结构\n结合影像综合判断'),(10.49,'步行、上下楼\n下蹲、转向、运动')]:txt(s,t,x,4,2.5,1.25,20,GRAY)
band(s,'专项试验提供线索；诊断需要病史、体征与必要影像相互印证。')
#5
s=slide('前交叉韧带：落地与变向中的稳定失守','136',note='教材指出ACL损伤多见于快速变向和起跳落地项目，常为非接触性损伤。足部固定时的扭转、外翻及过伸可能参与损伤。强调不是看到一个动作就断定ACL撕裂，也要关注半月板、骨挫伤等合并问题。')
panel(s);knee(s,1.5,2.2,'acl');label(s,'旋转 / 外翻 / 过伸',1,2.06,4,RED);arrow(s,4.3,3.55,.95,RED);label(s,'异常前移与旋转',1.25,5.13,4,RED)
explain(s,[('典型场景','急停、变向、起跳落地；常见非接触性机制。'),('急性线索','响声、剧痛、急性肿胀，不能继续运动。'),('后续表现','打软腿、错动感、积液；需关注合并损伤。')]);band(s,'机制提示 ACL 风险；急性症状与慢性不稳要分别询问。')
#6
s=slide('前交叉韧带：从前向稳定性到运动功能','136',note='教材列出Lachman与前抽屉试验。此课聚焦检查目的：评价前向稳定性，观察位移和终末感；操作细节应在后续规范实训中演示。疼痛、肌肉防御等可影响结果，多个线索结合比单次检查更稳妥。X线用于排查骨性损伤，MRI评价软组织。')
box(s,.7,2.03,5.25,3.92,PALE);chip(s,'前向稳定性',1.05,2.35,4.5);txt(s,'Lachman 试验\n前抽屉试验',1.12,3.35,4.6,1.2,28,NAVY,True);txt(s,'关注：胫骨前移与终末感',1.12,5.12,4.6,.5,20,GRAY)
explain(s,[('检查的目的','为 ACL 受损提供体征线索，综合判断。'),('急性期功能','疼痛、肿胀、活动度受限，运动加重疼痛。'),('慢性期功能','股四头肌萎缩、步态异常、下肢整体功能下降。')]);band(s,'不只记录“阳性 / 阴性”，还要记录疼痛、活动与不稳的影响。')
#7
s=slide('半月板：负重与旋转为何容易造成损伤？','137',note='图为半月板俯视概念图，内侧C形、外侧近O形。教材强调承重、旋转、过度屈曲，以及半月板未能随骨面协调移动时受挤压。周缘血供相对丰富，中央血供少；恢复还与撕裂类型和位置等有关。避免直接重复教材关于内外侧“更严重”的泛化结论。')
panel(s)
# schematic C and O by arcs represented editable rings with masking
for x,t,cut in [(1.25,'内侧：C 形',True),(3.62,'外侧：近 O 形',False)]:
 box(s,x,2.65,1.65,2.05,TEAL,kind=MSO_SHAPE.OVAL);box(s,x+.33,2.98,.99,1.39,'F7F9FC',kind=MSO_SHAPE.OVAL)
 if cut:box(s,x+1.04,3.28,.78,.8,'F7F9FC')
 txt(s,t,x-.18,5.03,2.13,.45,18,TEAL,True,PP_ALIGN.CENTER)
txt(s,'半月板俯视概念图 · 非比例',1,5.57,4.8,.3,11,GRAY)
explain(s,[('正常作用','分散负荷、吸收冲击，参与稳定与润滑。'),('损伤条件','负重时扭转、急停急转或过度屈曲。'),('恢复差异','周缘血供相对丰富；中央区血供较少。')]);band(s,'关注“承重 + 旋转”的组合，也要问卡顿、交锁和关节线疼痛。')
#8
s=slide('半月板：机械性症状与专项试验相互核对','137',note='教材列出关节线压痛、McMurray、Apley和Thessaly试验及MRI。重点解释其在疑似半月板损伤中的线索作用。卡锁不等同于普通活动疼痛；不要强行完成负重旋转试验。具体操作、禁忌与诊断效能不在这四页教材的展开范围内。')
box(s,.7,2.02,5.25,3.95,PALE);txt(s,'交锁 / 卡顿\n关节线疼痛\n转向或屈膝疼痛',1.05,2.48,4.65,2.6,30,NAVY,True)
explain(s,[('询问与触诊','询问机械性症状；检查内、外侧关节线压痛。'),('专项试验线索','McMurray、Apley、Thessaly；\n与病史、体征综合判断。'),('功能与影像','急性期可痛肿、活动受限；MRI 辅助判断。')]);band(s,'出现交锁时，先区分活动真正受阻，还是因疼痛而不敢活动。')
#9
s=slide('副韧带：外翻与内翻指向不同的受力侧','138',note='教材描述内侧副韧带多与外翻应力有关，外侧副韧带多与内翻应力有关，后者还可能累及外侧相关结构。图为概念性受力示意，文字说明外翻牵拉内侧、内翻牵拉外侧，避免将外力来源与受损侧混淆。')
panel(s);knee(s,1.45,2.2,'collateral');label(s,'内侧支持结构',.9,2.1,2.5,ORANGE);label(s,'外侧支持结构',3.45,5.15,2.5,ORANGE)
explain(s,[('内侧副韧带','外翻应力牵拉内侧结构；可合并其他损伤。'),('外侧副韧带','内翻应力相关；需关注邻近肌腱与外侧结构。'),('关键区分','外翻 / 内翻描述关节姿态，不等同于撞击侧。')]);band(s,'用“关节如何变形 → 哪一侧被牵拉”理解损伤机制。')
#10
s=slide('副韧带：局部疼痛与侧向不稳','138',note='教材诊断线索包括内外侧表浅疼痛、急性屈伸范围下降、内翻或外翻应力试验和MRI。此处强调应力试验的目的是侧向稳定性评价，记录疼痛与松弛而非只记阳性。操作由具备资质人员在适当时机实施。')
flow(s,['疼痛在哪里？','活动如何受限？','侧向稳定如何？','是否影响行走？'])
for x,t in [(.95,'内侧 / 外侧\n局部表浅疼痛'),(4.13,'急性肿胀\n屈伸范围下降'),(7.31,'内翻 / 外翻应力\n观察疼痛与松弛'),(10.49,'肌力下降\n步行与运动不稳')]:txt(s,t,x,4,2.5,1.25,21,GRAY)
band(s,'疼痛侧、受力机制、稳定性与功能表现应形成一致的解释。')
#11
s=slide('髌骨脱位：静态限制与动态控制共同失衡','138—139',note='教材介绍髌骨向外侧脱位及髌股不稳。股内侧斜肌是动态稳定结构，MPFL属于内侧静态限制结构，两者不能混为同一种组织。高位髌骨、滑车发育因素、下肢排列等可参与。图示向外侧偏移为概念示意。')
panel(s);knee(s,1.2,2.2,'patella');box(s,3.7,3.05,.75,1.0,'EFC4C8',RED,MSO_SHAPE.OVAL);arrow(s,3.25,3.4,.42,RED);label(s,'外侧偏移（示意）',2.65,5.12,3,RED)
explain(s,[('静态限制','MPFL 等内侧支持结构限制髌骨外移。'),('动态控制','股内侧斜肌参与髌骨动态稳定。'),('其他相关因素','髌骨高度、滑车形态、下肢排列及关节松弛。')]);band(s,'髌股不稳涉及结构与控制，不能简单归因于某一块肌肉。')
#12
s=slide('髌股不稳：关注“不放心用力”的膝关节','139',note='教材列出不稳感、发软或踏空感，以及髌骨外推、恐惧、倾斜试验。解释恐惧是对再次脱位的不安和防御，不是要求检查者诱发实际脱位。功能表现包括卡顿、错动、疼痛肿胀、上下楼和下蹲困难。急性疑似脱位先规范医疗评估。')
box(s,.7,2.02,5.25,3.95,PALE);txt(s,'不稳 / 发软 / 踏空\n反复脱位经历\n上下楼与下蹲困难',1.05,2.48,4.65,2.6,27,NAVY,True)
explain(s,[('病史线索','询问既往脱位、外伤及反复不稳经历。'),('专项检查','髌骨外推、恐惧、倾斜试验提供线索。'),('观察实际功能','记录疼痛、肿胀、卡顿及活动时的信心。')]);band(s,'不稳感与疼痛可同时存在；评定要关注活动中的安全感。')
#13
s=slide('髌骨软化：膝前痛需要结合负荷情境理解','139',note='沿用教材“髌骨软化”名称，但膝前痛是症状，不能仅凭疼痛就确认软骨改变。教材提及过度使用、直接外伤等因素，Clarke及磨髌试验和影像辅助诊断。此类试验可诱发疼痛且特异性有限，课堂以教材列举为主，不推荐反复诱痛。')
panel(s);knee(s,1.5,2.2,'patella');label(s,'髌股接触区域',.95,2.06,3,RED);line(s,2.2,2.5,3,3.6,RED);label(s,'反复屈伸 / 直接外伤',1.05,5.12,4,RED)
explain(s,[('教材所述机制','过度使用、直接外伤及髌股受力异常。'),('常见功能表现','反复屈伸后疼痛；较重时下楼、下蹲困难。'),('判断保持审慎','教材列 Clarke、磨髌等试验；必要时结合影像。')]);band(s,'膝前痛不等于髌骨软化，诱发疼痛的试验也不能单独确诊。')
#14
s=slide('把五类问题放进同一张对照表',note='此表是用于复习的线索对照，不是诊断决策规则。多种损伤可能并存，症状存在重叠。可请学生先遮住右列，尝试说出相应的主要功能影响。')
rows=[['问题','典型机制 / 症状线索','专项检查线索','主要功能影响'],['前交叉韧带','变向、落地；打软腿','Lachman、前抽屉','前向 / 旋转不稳'],['半月板','负重扭转；交锁卡顿','关节线压痛、旋转类试验','屈膝、转向受限'],['副韧带','外翻 / 内翻；内外侧痛','外翻 / 内翻应力试验','侧向稳定下降'],['髌股不稳','发软、踏空、脱位史','外推、恐惧、倾斜试验','上下楼、下蹲困难'],['髌骨软化','膝前痛；反复屈伸相关','教材列 Clarke、磨髌等','负荷相关活动疼痛']]
widths=[2.1,3.55,3.4,2.98];x0=.65;y=1.98
for i,row in enumerate(rows):
 x=x0
 for j,t in enumerate(row):
  box(s,x,y+i*.63,widths[j],.63,B if i==0 else ('EDF2FB' if i%2 else 'F7F9FC'),'FFFFFF');txt(s,t,x+.13,y+i*.63+.15,widths[j]-.23,.43,17,'FFFFFF' if i==0 else INK,i==0 or j==0);x+=widths[j]
band(s,'同一患者可能存在合并损伤：按证据整合，不按单一症状贴标签。')
#15
s=slide('病例练习：急停后膝关节肿胀，如何评定？',note='教学虚构病例。给学生1分钟讨论：提出至少两个待排查问题，说明需要追问和检查的内容。不要要求在信息不足时给出唯一诊断。病例仅用于临床推理训练，不是治疗处方。')
box(s,.7,2.03,5.25,3.92,PALE);txt(s,'足球运动员，急停变向后\n听到响声，膝部迅速肿胀，\n无法继续比赛；数日后\n感觉“打软腿”。',1.02,2.55,4.65,2.7,27,NAVY,True)
explain(s,[('先提假设','优先关注什么结构？还可能合并什么损伤？'),('再找证据','需要补问什么？选择什么评定线索？'),('最后看功能','除疼痛外，还应记录哪些活动影响？')]);band(s,'回答结构：我的假设是……支持证据是……尚需核实的是……')
#16
s=slide('病例解析：机制、结构与功能需要相互印证','136—138',note='参考回答：优先怀疑ACL损伤，但不能确诊；同时关注半月板、其他韧带和骨性损伤。询问非接触还是接触、肿胀时程、能否负重、交锁等。基本检查后在适当条件下做ACL针对性检查，必要影像。急性剧痛肿胀和不能负重时先医疗评估，不强行试验。')
flow(s,['提出假设','补充证据','评价功能','形成综合判断'])
for x,t in [(.95,'ACL 受损可能\n兼顾合并损伤'),(4.13,'动作、肿胀、交锁\n基本检查、专项检查'),(7.31,'活动度、肌力\n负重、步态与不稳'),(10.49,'病史 + 体征\n必要影像与转诊')]:txt(s,t,x,4,2.5,1.25,20,GRAY)
band(s,'响声、肿胀与打软腿支持进一步排查，不能替代正式诊断。')
#17
s=slide('从“结构诊断”走向“功能问题清单”',note='此页为教材功能障碍内容的教学归纳，不是新增康复治疗方案。把观察到的问题转化为可重复记录的项目：疼痛肿胀、屈伸范围、肌力和稳定、活动表现。复评尽量保持相同条件与方法，按阶段选择合适活动，不在急性期强行做高负荷测试。')
for i,(h,b) in enumerate([('症状与体征','疼痛、肿胀、积液\n诱发与缓解因素'),('身体功能','屈伸活动度、肌力\n稳定性与控制'),('活动表现','步行、上下楼、下蹲\n转向及运动参与')]):
 x=.7+i*4.17;box(s,x,2.15,3.88,3.58,PALE);txt(s,f'0{i+1}',x+.25,2.4,3.3,.6,38,B,True);txt(s,h,x+.25,3.25,3.3,.5,26,NAVY,True);txt(s,b,x+.25,4.17,3.3,1.2,22)
band(s,'复评要回答：哪些功能在改善，哪些问题仍然限制日常与运动？')
#18
s=slide('课堂回扣：用三句话讲清一类损伤',note='可请学生任选一种损伤，用机制、评定、功能三句话回答。参考：ACL常见于变向落地；Lachman与前抽屉提供前向稳定性线索；急性痛肿受限，慢性可不稳和肌力下降。再次强调避免把单个试验当作确诊依据。')
for i,(h,b) in enumerate([('机制','在什么动作、受力和负荷条件下发生？'),('评定','哪些病史、体征和检查线索支持判断？'),('功能','疼痛、活动度、稳定性和日常活动如何改变？')]):
 y=2.12+i*1.23;chip(s,h,.85,y,1.65);txt(s,b,2.95,y+.06,9.2,.75,27,INK,True)
band(s,'评定的终点：解释患者做不了什么，并为后续康复提供依据。')
#19
s=slide('教材依据与制作参考',note='本页为课后查阅信息。所有医学主体内容依据用户提供的4页教材；病例、评定链与对照表为教学组织加工。GitHub资源用于课件工程实现和版式一致性，不作为医学知识来源。')
txt(s,'教材内容',.8,1.98,11,.45,25,NAVY,True);txt(s,'用户提供《运动康复评定学-膝关节损伤》\n第二章 膝关节功能评定 · 第一节 膝关节损伤的功能障碍，p.136—139。',.85,2.63,11.7,1.05,22)
txt(s,'GitHub 制作参考',.8,4.03,11,.45,25,NAVY,True);txt(s,'gitbrent/PptxGenJS — 统一母版、文本与矢量图形的组件化组织\nhttps://github.com/gitbrent/PptxGenJS\nscanny/python-pptx — 本课件采用的可编辑 PPTX 生成工具\nhttps://github.com/scanny/python-pptx',.85,4.58,11.7,1.48,16,GRAY)
band(s,'沿用原蓝白教材模板；图形为可编辑示意图，病例与归纳内容已在备注中说明。')
