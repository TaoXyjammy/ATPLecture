# 膝关节与踝足损伤课程 · 中文标注 / 教材顺序 / 足球专项版

保留原蓝白教材模板，两份课件独立授课。严格按照用户上传的教材损伤种类顺序，每项内部依次呈现：**损伤机制 → 诊断线索 → 主要功能障碍 → 紧接足球专项预防**。诊断与功能在同页按先后顺序讲解，不跨病种集中预防。

| 课件 | 页数 | PPT | PDF |
|---|---:|---|---|
| 膝关节损伤的功能障碍（足球专项） | 21 | [下载](Knee_Sports_Injury_Prevention.pptx) | [预览](Knee_Sports_Injury_Prevention.pdf) |
| 踝足损伤的功能障碍（足球专项） | 24 | [下载](Ankle_Sports_Injury_Prevention.pptx) | [预览](Ankle_Sports_Injury_Prevention.pdf) |

## 本次中文标注修订

- 解剖图内的英文术语已转为中文，覆盖膝关节侧面、半月板、髌骨及踝韧带图，并处理封面中的重复配图。
- 踢球动图的英文动作说明、髋膝名称和负荷说明，以及踝扭伤图帧中的英文说明均已中文化；训练图编号改为甲、乙、丙、丁。
- 中文标注为可编辑PPT文字层，保留原来的解剖图、引线、GIF帧和运动顺序；不通过重画更改组织结构。
- 比赛画面中的球队、电视台、广告等原始赛事标识保留，原作者的正式署名、网址和许可记录保存在备注与素材清单中。国际单位符号、文献与试验名称按学术使用保留。
- `source/Chinese_labels.json` 保存38项英文与中文标注的对应关系与位置；源码包内的原图保留原始版本，运行构建脚本会自动叠加中文标注。

## 前次三个修改要求的落实

1. **教材顺序**：伤病顺序与上传的两份PDF一致。每项损伤讲完后紧接该项预防页，再进入下一项；保存了机器核对用的“顺序核对.json”。
2. **配图与动图**：替换原来的简笔骨骼图，采用有来源的细节解剖图和三维医学插图；加入真实足球比赛ACL病例GIF、踝扭伤研究图帧GIF。正常踢球动图来自GitHub，明确标注不是受伤录像。
3. **足球专项**：解释贯穿逼抢急停、支撑脚转身、争顶落地、冲刺蹬地、铲抢对抗；对应预防也具体落实到足球任务、训练渐进、场地和对抗组织。

## 膝关节教材顺序（上传教材p.136—139）

| 教材章节 | 课件页码 | 配套预防页 |
|---|---|---:|
| 一、前交叉韧带损伤 | 3—6（含第4页真实比赛GIF） | 6 |
| 二、半月板损伤 | 7—9 | 9 |
| 三、副韧带损伤 | 10—12 | 12 |
| 四、髌骨脱位（髌股关节不稳） | 13—15 | 15 |
| 五、髌骨软化 | 16—18 | 18 |

第1—2页为封面与教材顺序，第19页回顾，第20—21页来源。

## 踝足教材顺序（上传教材p.164—169）

| 教材章节 | 课件页码 | 配套预防页 |
|---|---|---:|
| 一、踝关节扭伤 | 3—6（含第4页真实研究图帧GIF） | 6 |
| 二、跟腱病变 | 7—9 | 9 |
| 三、跟腱断裂 | 10—12 | 12 |
| 四、足底筋膜炎 | 13—15 | 15 |
| 五、胫骨内侧应力综合征 | 16—18 | 18 |
| 六、踝关节骨折 | 19—21 | 21 |

第1—2页为封面与教材顺序，第22页回顾，第23—24页来源。未将后续“主观检查”另立为新的损伤章节。

## 真实结构与动态图

图片均为细节解剖参考或医学插图，不冒称为某位球员的病灶照片。

- **ACL真实比赛GIF**：Waldén et al., BJSM 2015，补充视频1，来自足球比赛研究病例。约1.9秒，保留原画面标识；转GIF、缩放、降帧，不改事件内容。[研究与补充材料](https://pmc.ncbi.nlm.nih.gov/articles/PMC4680158/) / [DOI](https://doi.org/10.1136/bjsports-2014-094573)。原文及补充材料为CC BY-NC 4.0，供非商业教学使用，第三方标识与既有权利保留。
- **踝扭伤图帧GIF**：Buckthorpe et al., KSSTA 2025，图2(c1—c4)，争顶落地情境。是把原文四张真实研究图帧按顺序重排，不是连续录像，不能据此测量精确损伤时刻。[研究及图2](https://pmc.ncbi.nlm.nih.gov/articles/PMC12582226/) / [DOI](https://doi.org/10.1002/ksa.70049)。原文CC BY 4.0；转换方式已注明。
- **GitHub正常踢球GIF**：Roberto Martín Gutiérrez本人正常踢球慢动作，[broadcast-kick-biomechanics](https://github.com/convderover/broadcast-kick-biomechanics)，原路径 `resultados/videos/07_golpeo_real_autor.mp4`。仓库对该视频采用CC BY 4.0。它不是受伤录像，彩色模型叠加也不是伤情标记；未把该项目结果作为医学诊断依据。
- **结构图**：Blausen Medical、Scientific Animations、InjuryMap、Servier Medical Art等；具体作者、原图链接、许可证与加工方式见 [MEDIA_LICENSES.md](MEDIA_LICENSES.md) 及 [MEDIA_SOURCES.json](MEDIA_SOURCES.json)。SVG转PNG、裁切或缩放等修改已记录，CC BY-SA改编图继续遵循相应原许可。
- 另外实际查阅了 [MonaElkhouly等的足球ACL生物力学课程项目](https://github.com/MonaElkhouly/Biomechanics-Final-Project-ACL-Injuries-for-female-football-player-Analysi-)，用于了解足球课程讲解切入点；未重发该项目未明确授权的素材，也未将课程项目结论当作临床证据。

## 预防拓展依据

教材以损伤、检查与功能障碍为主；预防页明确标注“预防拓展”，并与教材内容区分。

- [JOSPT 2023：膝与ACL损伤的运动预防](https://doi.org/10.2519/jospt.2023.0301)
- [JOSPT 2021：外侧踝扭伤与慢性踝不稳](https://doi.org/10.2519/jospt.2021.0302)
- [FIFA Training Centre：渐进热身与技术衔接](https://www.fifatrainingcentre.com/en/practice/grassroots/grassroots-and-youth-football-essentials/grassroots-coaching-essentials/optimising-warm-ups-u12-to-u15.php)

各预防页中的绕标、分站、接传球等是依据机制与训练原则编写的课堂任务示例，不声称它们是已独立验证的新预防方案。专项试验作为判断线索，不单独确诊；伤后回场仍需综合评估。

## 放映与源码

- GIF已直接嵌入PPTX，放映模式查看；编辑视图可能只显示一帧。
- PDF是静态预览，不播放GIF。
- [下载完整源码与素材包](Course_Source.zip)。解压后包含 `source/build_courses.py` 和 `source/assets/`；逐页讲解备注也单独提供Markdown文件。
- 生成代码依赖 `python-pptx`、`Pillow`，中文字体为Noto Sans CJK SC，可按本机字体调整：

```sh
pip install python-pptx pillow
python source/build_courses.py
```

源码包包括需要的媒体和中文标注表，生成时不依赖在线链接。仓库单独展示的构建脚本需配合压缩包内的assets目录使用。

## 完成检查

已核对每个病种的顺序与预防页位置；两份课件分别为21和24页。已重新打开PPTX、检查ZIP部件与对象边界、检查全页缩略图，并导出PDF。PPTX内GIF保留多帧：正常踢球31帧、ACL病例19帧、踝扭伤研究图帧4帧。未在用户本地WPS中实测放映。
