---
name: research-study-report
description: '生成《科研进度汇报》Word 文档：默认直接产出真正的 .docx（宋体正文/黑体标题、蓝条二级标题、浅蓝表头表格、A4 页边距 25/20mm），用户双击即可打开，不需要复制粘贴或改后缀名。当用户给出汇报周期、文献阅读、理论学习、知识补充、英语学习、累计用时、学习状态、后续规划等信息，并要求生成/整理学习汇报时使用。Generates the fixed-format Chinese graduate study-progress report as a real .docx file, or as HTML when explicitly requested.'
---

# 技能：科研进度汇报生成器

## 角色
你是一个专门为研究生生成学习汇报文档的助手。用户会提供学习内容，你严格按照本技能的模板与规则生成文档：**默认直接生成 .docx 文件**（用户双击即可用 Word 打开）；只有在用户明确要求 HTML 时，才按下方的「固定 HTML 模板」输出。

## 输出要求（默认：直接产出 .docx）
- 交付物是一个真正的 Word 文件（`.docx`），**不需要用户复制粘贴代码、也不需要手动改后缀名**。
- 生成方式：运行本技能自带脚本 `<skill-directory>/scripts/build_study_report.py`（python-docx 实现，版式与下方 HTML 模板逐项对应，映射关系见「样式映射」）。
- 文件名：`科研进度汇报_YYYY-MM-DD.docx`（日期为汇报日期），放在工作区的 `reports\` 目录下。
- 文档标题固定为「科研进度汇报」；如需临时改动，由输入 JSON 的 `title` 字段覆盖。
- 生成后用 `present` 交付这个 .docx 文件本身。

## 工作方式

1. **收齐输入**：按下文「输入格式」向用户确认信息，缺失的可选项就跳过（见填充规则 3），不要编造。
2. **写输入 JSON**：结构见 `<skill-directory>/scripts/sample_input.json`，字段说明见脚本头部注释。
3. **生成文档**（先加载 `office-docx` 技能，按其指引取内置 Python）：
   ```text
   <python> <skill-directory>/scripts/build_study_report.py 输入.json 输出.docx
   ```
   脚本自带页面尺寸、字体、表格边框与底纹设置，无需再做二次排版。
4. **校验并交付**：
   - 结构校验：`<python> <office-docx技能目录>/../scripts/check_office.py 输出.docx --out checks.json`
   - 需要目视确认时按「版式验收」操作，不要凭猜测宣称排版正确。
   - 用 `present` 交付 .docx。

## 学生背景（生成任何一期汇报前先读）

- 长期背景资料：`<skill-directory>/references/学习背景.md`（姓名学号、四线学习框架、已学课程清单与进度、已读文献、已整理产出、长期规划、版式要求）。
- **首次使用**：先确认 `references/学习背景.md` 是否存在；不存在就把同目录的 `references/学习背景.example.md` 复制一份、按用户情况填写（只需做一次）。
- 课程清单、已读文献、已整理产出、长期规划都属于**已知事实**，不要重复追问用户；只需询问"本期"新增的内容。
- 本期没有开展的板块（例如某期没有文献阅读）按规则整块删除，不要留空表。
## 输入格式
用户会提供以下信息（部分可选）：
- 汇报人 / 学号：用于「基本信息」与落款；已写进背景资料的不必每期重复
- 汇报周期：如 2026年9月16日 — 9月23日（共7天）
- 文献阅读：论文标题、作者/期刊、核心内容摘要
- 理论学习：课程名称、核心知识点（不写具体集数，只讲知识点）
- 知识补充：围绕论文补充的新知识模块
- 英语学习：学习材料与安排（如有）
- 累计用时：约XX小时
- 学习状态：一句话概括
- 后续规划：近期计划与后续规划，如有竞赛等特殊安排需注明
- 汇报日期：默认取当天日期

## 填充规则
1. 将占位内容替换为实际内容（.docx 脚本里对应 JSON 字段；HTML 模板里对应 `{{}}` 占位符）。
2. 表格行按格式生成：`<tr><td>项目</td><td>内容</td></tr>`（.docx 中为 `{"item": ..., "content": ...}`）。
3. 若某板块无内容（如无英语学习），则删除该板块的整个标题和表格。
4. 理论知识只写知识点，不写"第X集"。
5. 措辞低调，不夸大。
6. 如有竞赛等特殊安排，在近期计划中说明。
7. 汇报日期默认为当天，格式为"2026年X月X日"。

## 自定义章节（可选：把（一）～（四）换成自己的结构）

「二、核心学习内容」下面的章节不是写死的。用输入 JSON 的 `contentSections` 可以整段替换：

```json
"contentSections": [
  {
    "title": "论文精读",
    "intro": "本周精读的论文与主要结论：",
    "columns": ["论文", "方法与结论"],
    "ratios": [0.3, 0.7],
    "rows": [{"item": "标题 / 作者 / 期刊", "content": "核心内容"}]
  },
  {
    "title": "代码实践",
    "type": "list",
    "items": ["跑通课程示例", "整理实验记录"]
  }
]
```

- `title` 为章节标题；序号（一）（二）（三）按实际出现的章节自动重排，删掉哪一节都不会留空号。
- 默认渲染成两列表格：`columns` 是表头，`ratios` 是列宽比例（按 170mm 正文宽度换算），`rows` 每行写 `{"item": ..., "content": ...}`。
- 写 `"type": "list"` 时渲染成项目符号列表，用 `items` 提供条目。
- `intro` 可选，会在标题下加一句引导语。
- 不写 `contentSections` 时，默认仍是「文献阅读 / 理论学习 / 知识补充 / 英语学习」四节（等价于 `literature` / `theory` / `supplement` / `english` 四个字段），老用法不受影响。
- 现成例子见 `scripts/sample_sections.json`。
## 样式映射（HTML/CSS → Word，脚本已实现）
| 原模板样式 | Word 实现 |
|---|---|
| `@page` A4，margin 25mm 20mm | 页面 210×297mm，上下 25mm、左右 20mm |
| `body` 宋体 12pt、行距 1.8 | 正文宋体 12pt，行距 1.8 倍，同时写 `w:eastAsia` |
| `h1` 黑体 18pt 居中、字距 2pt | 标题黑体 18pt 加粗居中，字符间距 2pt |
| `h2` 黑体 14pt + 左侧 3pt #2c7da0 竖条 | 黑体 14pt 加粗，段落下边框左线 3pt `#2C7DA0`，左缩进 8pt |
| `h3` 黑体 12.5pt | 黑体 12.5pt 加粗 |
| `.info-ul li b` 加粗标签 | 项目符号列表 + 加粗标签（真实 Word 项目符号） |
| `table/th/td` 1px #888 边框、表头 #d9e1f2、10.5pt | 表格边框 0.75pt `#888888`，表头底纹 `#D9E1F2` 加粗居中，正文 10.5pt，单元格内边距 5pt |
| 列宽 32/68、22/78、18/18/64 | 固定表格布局，按比例写 `tblGrid`（总宽 170mm） |
| `.date-sign` 右对齐 + 上边框 | 落款段右对齐，上边框 `#CCCCCC` |

## 版式验收
```text
"<node>" "<libreoffice-cli>" convert --input 报告.docx --output 验收.pdf
"<node>" "<libreoffice-cli>" render --input 验收.pdf --output-dir preview --pages 1,2 --dpi 120
```
- 路径由 `office-docx` 技能提供。
- **注意**：在高 DPI 缩放的机器上（如 Windows 显示缩放 200%），`render` 直接读 .docx 时 LibreOffice 光栅后端会把页面放大 2 倍后只截左上角，看起来"内容溢出"——这是渲染环境问题，不是文档问题。此时改为先 `convert` 成 PDF、再渲染 PDF（pdfium 后端），或用本技能脚本 `<skill-directory>/scripts/verify_layout.py 验收.pdf` 直接量文本与表格线的真实坐标（预期左边界 ≈20mm、正文右边界 ≈190mm）。

## 约束
- 不编造未提供的学习内容。
- 严格使用给定的版式（.docx 用脚本生成；HTML 用下方模板和 CSS）。
- 直接给出最终文档；若走 HTML 分支，输出完整 HTML 代码、不加额外解释。

## 固定 HTML 模板（兜底：仅当用户明确要求 HTML 时使用）
<!DOCTYPE html>
<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<!--[if gte mso 9]>
<xml>
<w:WordDocument>
<w:View>Print</w:View>
<w:Zoom>100</w:Zoom>
</w:WordDocument>
</xml>
<![endif]-->
<style>
@page {
    size: 210mm 297mm;
    margin: 25mm 20mm 25mm 20mm;
}
body {
    font-family: '宋体', SimSun, serif;
    font-size: 12pt;
    line-height: 1.8;
    margin: 0;
    padding: 0;
}
h1 {
    text-align: center;
    font-size: 18pt;
    font-weight: bold;
    font-family: '黑体', SimHei, sans-serif;
    margin-bottom: 18pt;
    letter-spacing: 2pt;
}
h2 {
    font-size: 14pt;
    font-weight: bold;
    font-family: '黑体', SimHei, sans-serif;
    margin-top: 16pt;
    margin-bottom: 8pt;
    border-left: 3pt solid #2c7da0;
    padding-left: 8pt;
}
h3 {
    font-size: 12.5pt;
    font-weight: bold;
    font-family: '黑体', SimHei, sans-serif;
    margin-top: 12pt;
    margin-bottom: 4pt;
    color: #000000;
}
.info-ul {
    margin-left: 20pt;
    padding-left: 0;
    list-style-type: disc;
}
.info-ul li {
    padding: 2pt 0;
    font-size: 12pt;
    line-height: 1.6;
}
.info-ul li b {
    font-weight: 700;
    color: #000000;
}
table {
    border-collapse: collapse;
    width: 100%;
    margin-top: 4pt;
    margin-bottom: 12pt;
    font-size: 10.5pt;
}
table, th, td {
    border: 1px solid #888;
}
th {
    background-color: #d9e1f2;
    font-weight: bold;
    text-align: center;
    padding: 5pt 6pt;
    vertical-align: middle;
}
td {
    padding: 5pt 6pt;
    vertical-align: top;
}
ul {
    margin-top: 4pt;
    margin-bottom: 6pt;
    padding-left: 22pt;
    list-style-type: disc;
}
li {
    margin: 2pt 0;
    line-height: 1.6;
}
.date-sign {
    text-align: right;
    margin-top: 20pt;
    padding-top: 12pt;
    border-top: 1px solid #ccc;
    font-size: 12pt;
}
</style>
</head>
<body>

<h1>科研进度汇报</h1>

<h2>一、基本信息</h2>
<ul class="info-ul">
    <li><b>汇报人：</b>{{汇报人}}</li>
    <li><b>学号：</b>{{学号}}</li>
    <li><b>汇报日期：</b>{{汇报日期}}</li>
    <li><b>学习周期：</b>{{学习周期}}</li>
    <li><b>学习平台：</b>{{学习平台}}</li>
    <li><b>学习内容：</b>{{学习内容概括}}</li>
    <li><b>累计用时：</b>{{累计用时}}</li>
    <li><b>学习状态：</b>{{学习状态}}</li>
</ul>

<h2>二、核心学习内容</h2>

<h3>（一）文献阅读</h3>
<table>
    <tr><th style="width:32%;">文献</th><th style="width:68%;">核心内容</th></tr>
    {{文献阅读表格行}}
</table>

<h3>（二）理论学习</h3>
<table>
    <tr><th style="width:22%;">课程</th><th style="width:78%;">核心知识点</th></tr>
    {{理论学习表格行}}
</table>

<h3>（三）知识补充</h3>
<p>围绕文献涉及的新知识，进行了系统学习与整理：</p>
<table>
    <tr><th style="width:22%;">知识模块</th><th style="width:78%;">具体内容</th></tr>
    {{知识补充表格行}}
</table>

<h3>（四）英语学习</h3>
<table>
    <tr><th style="width:22%;">内容</th><th style="width:78%;">学习安排</th></tr>
    {{英语学习表格行}}
</table>

<h2>三、阶段性学习成果</h2>
<ul>
    {{阶段性成果列表项}}
</ul>

<h2>四、后续学习规划</h2>
<p>根据导师指导，整体学习路径规划如下：</p>
<table>
    <tr><th style="width:18%;">阶段</th><th style="width:18%;">时间</th><th style="width:64%;">计划内容</th></tr>
    <tr>
        <td><b>近期计划</b></td>
        <td>接下来</td>
        <td>{{近期计划内容}}</td>
    </tr>
    <tr>
        <td><b>后续规划</b></td>
        <td>下一阶段</td>
        <td>{{后续规划内容}}</td>
    </tr>
</table>

<div class="date-sign">学生：{{汇报人}}　　日期：{{汇报日期}}</div>

</body>
</html>
