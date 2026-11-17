# 第二章 疼痛的运动康复评定｜教学参考改版 V2

29页，16:9。沿用用户提供《肌骨概述.pptx》的白底橙红模板。正文知识仍以《疼痛.pdf》章首页及第15—19页为依据。

- [可编辑PPT](Chapter2_Pain_Assessment_Reference_V2.pptx)
- [PDF预览](Chapter2_Pain_Assessment_Reference_V2.pdf)
- [逐页讲解备注](逐页讲解备注.md)
- [教材顺序核对](顺序核对.json)
- [制作源文件](Pain_Reference_V2_Source.zip)

## 本次调整

- 第3页加入两个虚构的运动参与者案例：同样报告下楼时4/10，活动、睡眠及担忧程度不同；由此引出完整评定的必要性。案例甲贯穿后续定位、量表、进程和结果记录。
- 第13页重绘正背面身体定位图，红点对应案例甲所述右膝前侧；右侧同步记录部位、范围、分布、内外侧和深浅。图形为可编辑轮廓示意，不是组织结构解剖图。正面图左对应受评者右侧。
- 第17—20页将VAS、NRS、VRS、FPS改为左右分栏，分别呈现工具或使用依据，以及操作和记录。
- 第22页与第24页同样采用问诊要点／案例信息分栏，结果仍用教材SAND汇总。

第一节顺序：评估概述→疼痛分类。第二节顺序：病史→部位→性质→程度（VAS→NRS→VRS→FPS）→发展进程→加重缓解因素→结果记录。导入、案例应用及来源页为教学补充，不改变教材知识点顺序。

## 参考来源及使用范围

1. [Sleep, Comfort & Pain Management 网页幻灯片源文件](https://github.com/sagreenxyz/sagreenxyz.github.io2/blob/main/src/pages/private/d439/sleep-comfort-pain-slides.astro)：借鉴评定内容左右分栏及工具选择与操作要点并列的组织方式。未照搬其用药方案、固定年龄界限及复测时间。
2. [PlymouthPsychology / PEARS](https://github.com/PlymouthPsychology/pears)：借鉴从日常体验进入疼痛认识，以及关注想法、感受与活动的讲解方式。本课件案例自行编写，不使用其术后建议作为普通运动者的处方。
3. [Symptom-Assesment-Interface](https://github.com/iamdigiovanni/Symptom-Assesment-Interface)：借鉴身体图定位＋文字记录的界面结构。身体图自行绘制，未复制原截图或代码。
4. [IASP 2020 疼痛定义](https://www.iasp-pain.org/publications/iasp-news/iasp-announces-revised-definition-of-pain/)
5. [NSW ACI VAS操作资料](https://aci.health.nsw.gov.au/__data/assets/pdf_file/0018/212913/ACI-Pain-assessment-tools-visual-analogue-scale.pdf)

仅借鉴教学呈现思路，没有将GitHub项目视为临床指南。正文为中文，量表缩写及SAND保留。VAS投影线段为示意；FPS页讲授选用和操作，不把自绘表情冒充标准量表。教材扫描全文未上传。

## 重建

源文件压缩包解压后，安装 `python-pptx`，执行 `python source/build_pain.py`。字体为Noto Sans CJK SC；可根据本机字体替换。逐页讲解已写入PPT演讲者备注，Markdown亦可单独阅读。

原27页课件与既有膝、踝课程保持不变，本版独立保存。
