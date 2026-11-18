# 骨骼肌的结构特点与收缩原理｜授课教案

授课教师：谢淘。14页，A4，可编辑Word与PDF。

- [下载Word教案](Skeletal_Muscle_Lesson_Plan.docx)
- [下载PDF教案](Skeletal_Muscle_Lesson_Plan.pdf)
- [源文件与PPT配图](Skeletal_Muscle_Lesson_Plan_Source.zip)
- [课时与内容核对](教案核对.json)

## 模板选择

从GitHub检索中文教案模板及lesson plan template后，参考 [registor/nwafuTeachingPlan](https://github.com/registor/nwafuTeachingPlan) 的结构。该仓库是作者个人编排的高校教案LaTeX模板，包含表格式授课计划、学情、目标、重点难点、教法、正文旁注、小结、作业与后记。

本文件采用上述结构思路，以Word独立排版：首页为授课信息和目标，正文左栏标时间、教法及PPT页码，右栏为教材内容、配图、教师活动、学生活动与检查要点。未套用该校校名、标识或作者个人信息，未将其称为学校官方统一模板。

## 内容与范围

对应最新《Skeletal_Muscle_Textbook_Text.pptx》26页教材原文配文版，覆盖《运动生理学》（第6版）第12—18页第一节第一、第二部分。知识内容保留PPT中的45段教材文字，增加教案必需的课堂组织、活动安排、检查问题、作业及反思栏。

按完整内容安排2课时、每课时45分钟，共90分钟；不是压缩成10—15分钟的单知识点试讲。授课班级、日期留空，教学后记留待实际授课后填写。

## 配图

图2-1至图2-7及肌小节局部均直接从该PPT的PDF预览中导出，保留PPT中的中文标注与彩色结构。没有再次生成或替换为新图片。源文件`figure_map.json`记录每幅图所在PPT页码与导出区域。Word中的每幅图可单独替换、缩放。

## 验证与重建

- 第一、第二课时分别45分钟，总计90分钟。
- 教材45段文字在导出PDF中完整显示。
- DOCX图片均与PPT配图导出文件的哈希对应。
- 人工检查14页PDF，正文、图像及图注没有被裁切或跨页打断。

源文件包含`build_lesson_plan.py`、PPT原文内容表、配图导出记录与PNG图像。安装`python-docx`、`Pillow`，执行`python source/build_lesson_plan.py`重建Word；可用WPS或LibreOffice导出PDF。字体使用Noto Serif CJK与Noto Sans CJK。

原PPT及此前课程均保留。此教案独立保存在`lesson-plans/skeletal-muscle/`。
