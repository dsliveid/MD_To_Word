# MD to Word Converter

一个将 Markdown 文件转换为 Word 文档的 Python 工具。

## 功能特性

- 支持将 Markdown 文件转换为 Word (.docx) 格式
- 支持标题（H1, H2, H3）转换
- 支持段落文本转换
- 支持代码块（包含语法高亮）
- 支持图片插入（自动居中）
- 支持无序列表
- 支持表格
- 自动处理相对路径图片
- 中文字体支持

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 转换单个文件

```python
from md_to_word import md_to_word

md_to_word('doc/input.md', 'doc/output.docx')
```

### 批量转换所有 Markdown 文件

直接运行脚本：

```bash
python md_to_word.py
```

这会自动转换 `doc` 文件夹下所有的 `.md` 文件为对应的 `.docx` 文件（保存在同一文件夹中）。

## 依赖库

- `markdown>=3.4.1` - Markdown 转 HTML
- `python-docx>=0.8.11` - Word 文档操作

## 注意事项

- 图片路径支持相对路径和绝对路径
- 如果输出文件已存在且无法删除，会自动生成带 `_new` 后缀的新文件
- 代码块使用 Courier New 字体，字号 9pt
- 图片默认宽度为 5 英寸并居中显示

## 示例

假设有以下 Markdown 文件 `example.md`：

```markdown
# 示例文档

这是一段普通文本。

## 代码示例

```python
def hello():
    print("Hello, World!")
```

## 列表示例

- 项目一
- 项目二
- 项目三
```

运行转换后，会生成 `example.docx` 文件，保留所有格式和样式。
