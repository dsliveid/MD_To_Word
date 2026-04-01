# MD to Word Converter

一个将 Markdown 文件转换为 Word 文档和 PDF 文档的 Python 工具，同时也支持 Word 文档转换为 Markdown 文件。

## 功能特性

### Markdown 转 Word

- 支持将 Markdown 文件转换为 Word (.docx) 格式
- 支持标题（H1, H2, H3）转换
- 支持段落文本转换
- 支持代码块（包含语法高亮）
- 支持图片插入（自动居中）
- 支持无序列表
- 支持表格
- 自动处理相对路径图片
- 中文字体支持

### Markdown 转 PDF

- 支持将 Markdown 文件转换为 PDF 格式
- 支持标题（H1-H4）转换
- 支持段落文本、加粗、斜体、行内代码
- 支持代码块（灰色背景，保留缩进，自动分页）
- 支持图片插入（自动缩放，居中显示）
- 支持有序/无序列表
- 支持表格（自动换行）
- 支持引用块
- 中文字体支持（SimHei 黑体、SimSun 宋体）

### Word 转 Markdown

- 支持将 Word 文档转换为 Markdown 格式
- 支持标题转换（Heading 1-3）
- 支持行内格式（加粗、斜体、行内代码）
- 支持代码块检测
- 支持有序/无序列表
- 支持表格转换
- 支持图片提取并保存到 images 目录

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### Markdown 转 Word

#### 转换单个文件

```python
from md_to_word import md_to_word

md_to_word('MdToWord/input.md', 'MdToWord/output.docx')
```

#### 批量转换所有 Markdown 文件

直接运行脚本：

```bash
python md_to_word.py
```

这会自动转换 `MdToWord` 文件夹下所有的 `.md` 文件为对应的 `.docx` 文件（保存在同一文件夹中）。

### Markdown 转 PDF

#### 转换单个文件

```python
from md_to_pdf import md_to_pdf

md_to_pdf('MdToPDF/input.md', 'MdToPDF/output.pdf')
```

#### 批量转换所有 Markdown 文件

直接运行脚本：

```bash
python md_to_pdf.py
```

这会自动转换 `MdToPDF` 文件夹下所有的 `.md` 文件为对应的 `.pdf` 文件。

也可以指定其他目录：

```bash
python md_to_pdf.py path/to/markdown/files
```

### Word 转 Markdown

#### 转换单个文件

```python
from word_to_md import word_to_md

word_to_md('WordToMd/input.docx', 'WordToMd/output.md')
```

#### 批量转换所有 Word 文件

直接运行脚本：

```bash
python word_to_md.py
```

这会自动转换 `WordToMd` 文件夹下所有的 `.docx` 文件为对应的 `.md` 文件。

**注意**：转换后的图片会保存到 `WordToMd/images/` 目录中。

## 依赖库

- `markdown>=3.4.1` - Markdown 转 HTML
- `python-docx>=0.8.11` - Word 文档操作
- `reportlab>=4.0.4` - PDF 文档生成
- `Pillow>=9.0.0` - 图片处理（用于 PDF 图片缩放）

## 注意事项

### Markdown 转 Word

- 图片路径支持相对路径和绝对路径
- 如果输出文件已存在且无法删除，会自动生成带 `_new` 后缀的新文件
- 代码块使用 Courier New 字体，字号 9pt
- 图片默认宽度为 5 英寸并居中显示

### Markdown 转 PDF

- 图片路径支持相对路径和绝对路径
- 图片自动缩放以适应页面宽度（最大 15cm）
- 代码块使用浅灰色背景，保留缩进
- 表格单元格支持自动换行
- 中文字体使用黑体（标题）和宋体（正文）

### Word 转 Markdown

- 图片会提取并保存到 `WordToMd/images/` 目录
- 代码块通过检测 Courier New 字体和缩进自动识别
- 列表嵌套级别通过段落缩进判断
- 转换后的 Markdown 文件保存在同一目录

## 目录结构

```
MD_To_Word/
├── md_to_word.py      # Markdown 转 Word 转换脚本
├── md_to_pdf.py       # Markdown 转 PDF 转换脚本
├── word_to_md.py      # Word 转 Markdown 转换脚本
├── requirements.txt   # 依赖库列表
├── MdToWord/          # Markdown 转 Word 的输入/输出目录
│   ├── *.md           # 输入的 Markdown 文件
│   └── *.docx         # 输出的 Word 文件
├── MdToPDF/           # Markdown 转 PDF 的输入/输出目录
│   ├── *.md           # 输入的 Markdown 文件
│   ├── *.pdf          # 输出的 PDF 文件
│   └── images/        # 图片目录
├── WordToMd/          # Word 转 Markdown 的输入/输出目录
│   ├── *.docx         # 输入的 Word 文件
│   ├── *.md           # 输出的 Markdown 文件
│   └── images/        # 提取的图片目录
└── README.md          # 说明文档
```

## 示例

### Markdown 转 Word 示例

假设有以下 Markdown 文件 `example.md`：

````markdown
# 示例文档

这是一段普通文本。

## 代码示例

```python
def hello():
    print("Hello, World!")
```
````

## 列表示例

- 项目一
- 项目二
- 项目三

````

运行转换后，会生成 `example.docx` 文件，保留所有格式和样式。

### Word 转 Markdown 示例

将 Word 文档放入 `md` 目录：

```bash
# 运行转换
python word_to_md.py
````

转换后会生成对应的 `.md` 文件，图片会保存到 `WordToMd/images/` 目录中。
