import os
import re
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Preformatted, Image, KeepTogether, Flowable
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# 注册中文字体
pdfmetrics.registerFont(TTFont('SimSun', 'C:/Windows/Fonts/simsun.ttc'))
pdfmetrics.registerFont(TTFont('SimHei', 'C:/Windows/Fonts/simhei.ttf'))
# 注册等宽字体用于代码
pdfmetrics.registerFont(TTFont('Consolas', 'C:/Windows/Fonts/consola.ttf'))


def create_code_block(code_text, width=15*cm):
    """创建代码块，支持自动分页，使用灰色背景，保留缩进"""
    # 代码块样式
    code_style = ParagraphStyle(
        name='CodeBlockStyle',
        fontName='SimSun',  # 支持中文
        fontSize=9,
        leading=12,
        textColor=colors.black,
        backColor=colors.Color(0.92, 0.92, 0.92),  # 浅灰色背景
        leftIndent=0.3*cm,
        rightIndent=0.3*cm,
        spaceBefore=0,
        spaceAfter=0,
    )

    # 每行作为单独的 Paragraph，这样可以自动分页
    elements = []
    lines = code_text.split('\n')
    for i, line in enumerate(lines):
        # 转义特殊字符
        safe_line = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        safe_line = safe_line.replace('\t', '    ')  # 制表符转4空格
        # 将空格转换为 &nbsp; 以保留缩进
        # 先转换前导空格（缩进）
        leading_spaces = len(safe_line) - len(safe_line.lstrip(' '))
        safe_line = '&nbsp;' * leading_spaces + safe_line.lstrip(' ')
        if safe_line.strip() == '' or safe_line == '&nbsp;' * leading_spaces:
            safe_line = '&nbsp;'  # 空行用空格替代
        para = Paragraph(safe_line, code_style)
        elements.append(para)

    return elements


def get_styles():
    """获取样式"""
    styles = getSampleStyleSheet()

    # 标题样式
    styles.add(ParagraphStyle(
        name='ChineseH1',
        fontName='SimHei',
        fontSize=24,
        leading=30,
        spaceAfter=12,
        spaceBefore=20,
    ))

    styles.add(ParagraphStyle(
        name='ChineseH2',
        fontName='SimHei',
        fontSize=18,
        leading=24,
        spaceAfter=10,
        spaceBefore=16,
    ))

    styles.add(ParagraphStyle(
        name='ChineseH3',
        fontName='SimHei',
        fontSize=14,
        leading=18,
        spaceAfter=8,
        spaceBefore=12,
    ))

    styles.add(ParagraphStyle(
        name='ChineseH4',
        fontName='SimHei',
        fontSize=12,
        leading=16,
        spaceAfter=6,
        spaceBefore=10,
    ))

    # 正文样式
    styles.add(ParagraphStyle(
        name='ChineseBody',
        fontName='SimSun',
        fontSize=12,
        leading=18,
        spaceAfter=6,
    ))

    # 列表样式
    styles.add(ParagraphStyle(
        name='ChineseBullet',
        fontName='SimSun',
        fontSize=12,
        leading=18,
        leftIndent=1*cm,
        bulletIndent=0.5*cm,
    ))

    # 引用样式
    styles.add(ParagraphStyle(
        name='ChineseQuote',
        fontName='SimSun',
        fontSize=12,
        leading=18,
        leftIndent=1*cm,
        backColor=colors.Color(0.97, 0.97, 0.97),
        textColor=colors.Color(0.4, 0.4, 0.4),
    ))

    # 代码块外层容器样式（带背景和边框）
    return styles


def md_to_pdf(md_file, output_file):
    """将 Markdown 文件转换为 PDF 文档"""

    # 读取 Markdown 文件
    with open(md_file, 'r', encoding='utf-8') as f:
        md_content = f.read()

    # 获取 Markdown 文件所在目录，用于解析相对路径图片
    md_dir = os.path.dirname(os.path.abspath(md_file))

    # 创建 PDF
    doc = SimpleDocTemplate(
        output_file,
        pagesize=A4,
        leftMargin=2*cm,
        rightMargin=2*cm,
        topMargin=2*cm,
        bottomMargin=2*cm,
    )

    styles = get_styles()
    story = []

    # 解析 Markdown 内容
    lines = md_content.split('\n')
    in_code_block = False
    code_content = []
    in_table = False
    table_data = []

    for line in lines:
        # 处理代码块
        if line.startswith('```'):
            if in_code_block:
                # 结束代码块
                in_code_block = False
                if code_content:
                    code_text = '\n'.join(code_content)
                    story.extend(create_code_block(code_text))
                    story.append(Spacer(1, 0.3*cm))
                code_content = []
            else:
                # 开始代码块
                in_code_block = True
            continue

        if in_code_block:
            code_content.append(line)
            continue

        # 处理标题（从最长匹配开始）
        if line.startswith('#### '):
            text = line[5:].strip()
            story.append(Paragraph(text, styles['ChineseH4']))
            continue
        elif line.startswith('### '):
            text = line[4:].strip()
            story.append(Paragraph(text, styles['ChineseH3']))
            continue
        elif line.startswith('## '):
            text = line[3:].strip()
            story.append(Paragraph(text, styles['ChineseH2']))
            continue
        elif line.startswith('# '):
            text = line[2:].strip()
            story.append(Paragraph(text, styles['ChineseH1']))
            continue

        # 处理无序列表
        stripped = line.lstrip()
        if stripped.startswith('- ') or stripped.startswith('* '):
            text = stripped[2:].strip()
            # 处理内联格式
            text = process_inline(text)
            story.append(Paragraph('• ' + text, styles['ChineseBullet']))
            continue

        # 处理有序列表
        list_match = re.match(r'^(\d+)\.\s+(.+)$', stripped)
        if list_match:
            num = list_match.group(1)
            text = list_match.group(2).strip()
            text = process_inline(text)
            story.append(Paragraph(f'{num}. ' + text, styles['ChineseBullet']))
            continue

        # 处理引用
        if line.startswith('> '):
            text = line[2:].strip()
            text = process_inline(text)
            story.append(Paragraph(text, styles['ChineseQuote']))
            continue

        # 处理图片 ![alt](path)
        img_match = re.match(r'^!\[.*?\]\((.+?)\)', stripped)
        if img_match:
            img_path = img_match.group(1)
            # 处理相对路径
            if not os.path.isabs(img_path):
                img_path = os.path.join(md_dir, img_path)
            # 如果图片存在，添加到PDF
            if os.path.exists(img_path):
                try:
                    # 设置最大宽度，保持原始比例
                    from PIL import Image as PILImage
                    pil_img = PILImage.open(img_path)
                    img_width, img_height = pil_img.size
                    # 计算缩放比例，最大宽度15cm
                    max_width = 15 * cm
                    scale = max_width / img_width
                    final_width = max_width
                    final_height = img_height * scale
                    # 如果高度超过页面，则按高度缩放
                    max_height = 18 * cm
                    if final_height > max_height:
                        scale = max_height / img_height
                        final_width = img_width * scale
                        final_height = max_height
                    img = Image(img_path, width=final_width, height=final_height)
                    img.hAlign = 'CENTER'
                    story.append(img)
                    story.append(Spacer(1, 0.3*cm))
                except Exception as e:
                    print(f"警告: 无法添加图片 {img_path}: {e}")
            else:
                print(f"警告: 图片文件不存在 {img_path}")
            continue

        # 处理表格
        if line.startswith('|') and '|' in line:
            cells = [c.strip() for c in line.split('|')[1:-1]]
            # 检查是否是分隔行
            if all(c.replace('-', '').replace(':', '') == '' for c in cells):
                in_table = True
                continue
            if cells:
                if not in_table:
                    in_table = True
                    table_data = []
                table_data.append(cells)
            continue
        elif in_table and table_data:
            # 表格结束，创建表格
            in_table = False
            if table_data:
                # 创建表格样式（用于单元格内容）
                cell_style = ParagraphStyle(
                    name='TableCell',
                    fontName='SimSun',
                    fontSize=10,
                    leading=14,
                    wordWrap='CJK',  # 支持中文换行
                )
                header_style = ParagraphStyle(
                    name='TableHeader',
                    fontName='SimHei',
                    fontSize=10,
                    leading=14,
                    wordWrap='CJK',
                )
                # 将单元格内容转换为 Paragraph 以支持自动换行
                para_table_data = []
                for row_idx, row in enumerate(table_data):
                    para_row = []
                    for cell in row:
                        cell_text = process_inline(cell)
                        if row_idx == 0:
                            para_row.append(Paragraph(cell_text, header_style))
                        else:
                            para_row.append(Paragraph(cell_text, cell_style))
                    para_table_data.append(para_row)
                # 创建表格
                col_count = len(table_data[0])
                col_width = 15*cm / col_count
                t = Table(para_table_data, colWidths=[col_width]*col_count)
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.96, 0.96, 0.96)),
                    ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # 垂直对齐顶部
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.Color(0.8, 0.8, 0.8)),
                    ('LEFTPADDING', (0, 0), (-1, -1), 4),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                    ('TOPPADDING', (0, 0), (-1, -1), 4),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ]))
                story.append(t)
                story.append(Spacer(1, 0.3*cm))
                table_data = []

        # 处理空行
        if not line.strip():
            story.append(Spacer(1, 0.3*cm))
            continue

        # 处理普通段落
        text = line.strip()
        if text:
            text = process_inline(text)
            story.append(Paragraph(text, styles['ChineseBody']))

    # 如果文件已存在，先删除
    if os.path.exists(output_file):
        try:
            os.remove(output_file)
        except Exception as e:
            print("警告: 无法删除现有文件 " + output_file + ": " + str(e))
            output_file = output_file.replace('.pdf', '_new.pdf')

    # 构建 PDF
    doc.build(story)
    print("转换完成: " + output_file)


def process_inline(text):
    """处理内联格式，转换为 reportlab 格式"""
    # 加粗 **text** -> <b>text</b>
    text = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)
    # 斜体 *text* -> <i>text</i>
    text = re.sub(r'\*(.+?)\*', r'<i>\1</i>', text)
    text = re.sub(r'_(.+?)_', r'<i>\1</i>', text)
    # 行内代码 `text` -> <font name="SimSun" size="10">text</font>
    text = re.sub(r'`(.+?)`', r'<font name="SimSun" size="10" backColor="#f4f4f4">\1</font>', text)
    return text


def batch_convert(folder_path):
    """批量转换指定文件夹中的所有 Markdown 文件"""
    if not os.path.exists(folder_path):
        print("错误: 文件夹 '" + folder_path + "' 不存在")
        return

    md_files = [f for f in os.listdir(folder_path) if f.endswith('.md')]

    if not md_files:
        print("在 '" + folder_path + "' 文件夹中没有找到 .md 文件")
        return

    for md_file in md_files:
        md_path = os.path.join(folder_path, md_file)
        output_file = md_file.replace('.md', '.pdf')
        output_path = os.path.join(folder_path, output_file)
        md_to_pdf(md_path, output_path)


if __name__ == "__main__":
    import sys
    # 获取项目根目录
    root_dir = os.path.dirname(os.path.abspath(__file__))
    # 默认转换 MdToPDF 目录，也可通过参数指定
    if len(sys.argv) > 1:
        target_dir = sys.argv[1]
    else:
        target_dir = os.path.join(root_dir, 'MdToPDF')
    batch_convert(target_dir)