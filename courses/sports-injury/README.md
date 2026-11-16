# 膝关节与踝足运动损伤预防课程

授课方向：体育学科，重点讲解运动损伤机制、风险识别、预防训练和安全参与。沿用既有白底蓝色教材模板，16:9；文本与示意图可在 WPS / PowerPoint 中编辑。

| 独立课件 | 页数 | PPTX | PDF |
|---|---:|---|---|
| 膝关节运动损伤与预防 | 22 | [下载 PPTX](Knee_Sports_Injury_Prevention.pptx) | [预览 PDF](Knee_Sports_Injury_Prevention.pdf) |
| 踝足运动损伤与预防 | 25 | [下载 PPTX](Ankle_Sports_Injury_Prevention.pptx) | [预览 PDF](Ankle_Sports_Injury_Prevention.pdf) |

## 讲解组织

两份均按“运动情境 → 损伤机制 → 识别线索 → 预防措施 → 病例应用”组织。结构相关页面采用左图右解；预防部分采用动作示意、流程与课堂任务。病例均为教学编写，图为可编辑概念示意，不是比例解剖图。

- 膝关节：ACL、半月板、副韧带、髌股不稳与髌骨软化；重点联系变向、落地、组合预防训练与体育课堂组织。
- 踝足：重点为外侧踝扭伤、再伤预防与回场，另涵盖跟腱病/断裂、足底筋膜相关痛、胫骨内侧应力综合征及踝部骨折。
- 备注：每页附来源、讲解提示、课堂问题与必要的证据边界；另有 Markdown 版本便于阅读。
- PDF 仅用于静态预览，讲解备注见 PPTX 的备注区或对应 Markdown 文件。

## 教材依据

用户提供的《运动康复评定学-膝关节损伤》p.136—139与《运动康复评定学-踝关节损伤》p.164—169。上传内容为课件重新编写的概括，不包含教材扫描页。

## GitHub 检索及实际采用方式

1. [Sithranjan-Suresh/injury-risk-ml-curriculum](https://github.com/Sithranjan-Suresh/injury-risk-ml-curriculum)
   - 实际查阅 [Week_1 的25页课程讲义](https://github.com/Sithranjan-Suresh/injury-risk-ml-curriculum/blob/main/Week_1/Mini_Missions_Week1_presentation.pdf)。其课程以 ACL 受伤情境引入，再逐步说明概念并插入提问。
   - 仅借鉴情境引入、分步解释、课堂提问的方式；这是机器学习入门课程，未将其中的数据、相关性或预测结果作为医学结论。
2. [keanu77/professional-sports-injury-atlas](https://github.com/keanu77/professional-sports-injury-atlas)
   - 实际查阅 README、ACL 与踝扭伤条目。借鉴“可见事件 / 医学判断 / 回场脉络”分开呈现的方式。
   - 它是运动伤害教学素材索引，非现成的整套课程PPT。条目含审阅状态，本课医学预防结论另依据下列原始指南。
   - 未复制其第三方影音、图片或完整医学编辑内容。
3. [freshehrteam/Ankle-sprain](https://github.com/freshehrteam/Ankle-sprain)
   - 查阅 docs/README.md，借鉴踝扭伤症状、功能信息结构化记录的思路；该项目是物理治疗记录模型项目，不是已验证的课程或预防试验。

这些资料支持讲解组织，不意味着本课已获上述作者审核、授权背书或临床验证。

## 预防拓展证据

教材以损伤及功能障碍为主，新增预防和回场部分单独标注来源。

- Arundale et al. **Exercise-Based Knee and Anterior Cruciate Ligament Injury Prevention**. JOSPT, 2023. [DOI](https://doi.org/10.2519/jospt.2023.0301)
- Padua et al. **National Athletic Trainers’ Association Position Statement: Prevention of Anterior Cruciate Ligament Injury**. 2018. [DOI](https://doi.org/10.4085/1062-6050-99-16)
- Martin et al. **Ankle Stability and Movement Coordination Impairments: Lateral Ankle Ligament Sprains Revision 2021**. JOSPT. [DOI](https://doi.org/10.2519/jospt.2021.0302)
- Smith et al. **Return to sport decisions after an acute lateral ankle sprain injury: introducing the PAASS framework**. BJSM, 2021. [DOI](https://doi.org/10.1136/bjsports-2021-104087)

课堂练习是教学任务示例，不提供适用于所有人的固定运动剂量。PAASS是回场考虑领域的共识框架，不是统一达标分数。病例用于讨论，专项试验不能单独确诊。

## 源码与生成

`source/build_courses.py` 为入口，`source/knee_foundation.py` 为膝课件基础布局。依赖 `python-pptx`，中文字体使用 Noto Sans CJK SC；可按本机字体调整。

```sh
pip install python-pptx
python source/build_courses.py
```

运行后在源码目录的上一级生成两份 PPTX 与对应讲解备注。PDF 可用 WPS / PowerPoint / LibreOffice 导出。制作工具参考 [python-pptx](https://github.com/scanny/python-pptx)，以原生文本与形状保留可编辑性。

## 验证

已通过 PPTX 重新打开、ZIP部件唯一性、对象边界检查，并以 LibreOffice 导出PDF；两份页数分别为22、25页。已检查全页缩略图及关键页面，修正文字换行。不同版本 WPS 的字体替换可能影响个别换行。
