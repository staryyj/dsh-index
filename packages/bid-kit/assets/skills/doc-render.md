---
name: doc-render
description: 文档渲染与页面校核：docx 转 PDF、统计页数、逐页截图核对、图片尺寸与页边距检查、LibreOffice Kit 用法速查
whenToUse: 需要确认"还是不是原来的页数""渲染出来看看""看看这页排版对不对"，或需要把 docx 批量转 PDF 核对时
---

# 文档渲染与页面校核

## 一、LibreOffice Kit CLI（本机可用的渲染工具）

```bash
# 环境变量（Windows Desktop 版）
NODE="$(node 路径)/node.exe"          # D:\deepseek\resources\runtime\primary-runtime\dependencies\node\bin\node.exe
CLI="D:\deepseek\resources\app.asar.unpacked\dsh\node_modules\@deepseek-ai\libreoffice-kit\lib\cli.js"

# 转换
"$NODE" "$CLI" convert --input 甲.docx --output 甲.pdf

# 渲染指定页（1-based，只接受逗号列表或 all）
"$NODE" "$CLI" render --input 甲.pdf --output-dir 输出目录 --pages "1,2,3" --dpi 60
```

**限制（踩过的坑）**：

| 限制 | 说明 |
|---|---|
| `--pages` | 只能 `all` 或逗号列表，不能写范围（`1-3` 不行） |
| 页数上限 | 一次渲染 ≤ 24 页；超过报 `output-too-large`，分批 |
| 输出目录 | **不能预先创建**（CLI 自己 mkdir，已存在会 EEXIST 失败） |
| `--dpi` | 24～600 |

## 二、页数核对标准流程

```bash
# 1) docx → pdf
# 2) 渲染第 1 页拿 pageCount
# 3) 与目标页数比对
```

目标页数约定：

| 文件 | 目标 |
|---|---|
| 技术标 A3 版 | 4 页 |
| 技术标 A4 版 | ≤10 页（招标文件上限） |
| 附表 | 5～6 页 |
| 整册 | 按模板页数，前后差 ≤2 页 |

**页数变了先查三处**：图片高度、表格行高、段前段后间距——不要靠删内容调。

## 三、逐页截图核对（视觉检查）

```bash
render --pages "1,2,3,4" --dpi 60 → 得到 page-0001.png …
```

重点看：

- 图有没有超出页面、被裁切
- 表格是否跨页错位、表头是否重复
- 页码、页眉页脚位置
- 中文是否缺字（渲染缺字说明字体缺失）

## 四、PDF 文本抽取的现实问题

平台生成的招标文件 PDF **抽出来的文字常常是乱序的**（同一行的标签和数值被拆开），所以：

- 抽文本只用于**关键词检索**（"暗标""技术标文件…页""否决"）
- **数值以页面图为准**：把公告/前附表渲染成 PNG，人眼或视觉模型读

## 五、docx 内部检查（不用渲染也能查）

```python
from docx import Document
d = Document('甲.docx')
# 图片尺寸
for ext in d.element.body.iter('{...}extent'):  # wp:extent
    print(ext.get('cx'), ext.get('cy'))        # 单位 EMU，÷360000 = cm
# 页数/边距
for s in d.sections:
    print(s.page_width, s.page_height, s.left_margin, s.right_margin)
```

**图片定位**：`w:drawing` → `wp:extent` 的 cx/cy ÷ 360000 得到厘米值；Word 里图片宽度是硬约束，改宽度会改页数。

## 六、常见问题速查

| 现象 | 原因 | 处理 |
|---|---|---|
| 转换后页数变多 | 字体缺失替换、行距不同 | 装字体或统一行距 |
| 图片模糊 | 插入分辨率低 | 生成 2600px/300dpi |
| 表格列宽错乱 | 只改了 tcW 没改 tblGrid | 同时设置 `w:tblGrid/gridCol` |
| 渲染 0 字节图 | 输出目录预建导致失败 | 删除目录重跑 |
| PDF 有字体缺失 | 中文未嵌入 | 转 PDF 时勾选嵌入或换通用字体 |
