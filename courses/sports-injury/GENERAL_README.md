# 常见运动损伤评定与预防 · 通用版

本版本面向一般体育活动与相关日常功能，足球作为损伤发生的举例素材。原有足球专项中文版完整保留，使用不同文件名并存。

| 课件 | 页数 | 可编辑PPT | PDF |
|---|---:|---|---|
| 膝关节损伤的功能障碍（通用版） | 21 | [下载](Knee_General_Injury_Prevention.pptx) | [预览](Knee_General_Injury_Prevention.pdf) |
| 踝足损伤的功能障碍（通用版） | 24 | [下载](Ankle_General_Injury_Prevention.pptx) | [预览](Ankle_General_Injury_Prevention.pdf) |

## 课程定位

- 讲解运动及相关日常活动中的损伤机制、诊断线索、功能障碍与一般预防。
- 活动情境覆盖跑跳、急停、转向、负重、久站、行走，以及跌倒或接触外力。
- 预防围绕渐进负荷、力量与平衡、动作控制、环境安全、既往损伤和症状反馈展开。
- 足球病例及动图保留在“损伤机制”环节，明确标注“足球示例”。足球研究中的事件只用于解释受力，不推定其他运动项目具有相同风险。
- 保留蓝白模板、中文可编辑标注、教材顺序及逐项配套预防。封面与预防示意采用一般活动表达。

## 教材顺序

膝关节依据上传教材p.136—139：前交叉韧带 → 半月板 → 副韧带 → 髌骨脱位（髌股关节不稳） → 髌骨软化。

踝足依据上传教材p.164—169：踝关节扭伤 → 跟腱病变 → 跟腱断裂 → 足底筋膜炎 → 胫骨内侧应力综合征 → 踝关节骨折。

每项均为“损伤机制 → 诊断线索与主要功能障碍 → 损伤预防”，然后进入下一项。足球例证安排在对应的机制环节。顺序核对文件与逐页讲解备注单独保存。

## 原足球专项中文版

- [膝关节足球专项版](Knee_Sports_Injury_Prevention.pptx)
- [踝足足球专项版](Ankle_Sports_Injury_Prevention.pptx)
- [原版说明](README.md)

原版PPT、PDF、讲解备注、来源资料与构建脚本均保留；通用版新增文件带 `General` 标记。

## 来源与素材

教材内容与预防拓展分开标注。膝与ACL运动预防参考[JOSPT 2023](https://doi.org/10.2519/jospt.2023.0301)，踝扭伤再伤预防参考[JOSPT 2021](https://doi.org/10.2519/jospt.2021.0302)。其他预防任务依据教材机制作教学归纳，结合能力、年龄、症状和活动需求调整。

真实足球ACL病例动图来自[Waldén等，BJSM 2015](https://doi.org/10.1136/bjsports-2014-094573)；踝扭伤为[Buckthorpe等，KSSTA 2025](https://doi.org/10.1002/ksa.70049)图2(c1—c4)的原序图帧动画，后者是离散图帧。医学插图作者、许可、中文标注方式及原始来源见 [General_MEDIA_LICENSES.md](General_MEDIA_LICENSES.md) 和 [source/General_MEDIA_SOURCES.json](source/General_MEDIA_SOURCES.json)。原始赛事标识和完整署名保留。

## 放映与修改

GIF已嵌入PPT，可在放映模式查看；PDF为静态预览。中文标注是可编辑的PPT文字层，原始解剖图和GIF各帧保留。

[下载通用版完整源码与素材](General_Course_Source.zip)，解压后运行：

```sh
pip install python-pptx pillow
python source/build_general_courses.py
```

压缩包包含通用版脚本、中文标注表、通用版来源清单和所需媒体，不依赖在线下载。运行只生成通用版文件，不覆盖足球专项版。

## 检查

已检查两份通用版的教材顺序、对象边界、PPTX部件唯一性与PDF版面；分别为21、24页。动图分别保留19帧的ACL片段和4帧的踝扭伤研究图帧。原足球专项PPT、PDF及构建脚本在工作区逐字节校验一致。
