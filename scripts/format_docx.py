#!/usr/bin/env python3
"""
format_docx.py — 将 Word 文档格式化为中文学术论文或公文样式。

用法（基础）:
    python format_docx.py --input 草稿.docx --style academic --output 输出.docx

用法（AI 语义模式，推荐）:
    # 第一步：提取段落列表让 Claude 分析
    python extract_paragraphs.py --input 草稿.docx
    # Claude 读取输出，生成段落类型映射 JSON
    # 第二步：携带映射执行格式化
    python format_docx.py --input 草稿.docx --style academic \
        --heading-map '{"0":"title","4":"h1","11":"h1"}' --output 输出.docx

--heading-map 格式：JSON 字符串，key 为段落索引（字符串），value 为类型
    类型值：title / h1 / h2 / h3 / body
    未在 map 中指定的段落使用自动检测（样式名 + 内容规则）作为兜底

处理范围（方案 A）：
  - 正文段落、标题段落（H1/H2/H3）、文档标题
  - 页面边距
  - 不处理：表格、图片、页眉页脚、脚注
"""

import argparse
import json
import re
import sys
from pathlib import Path

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ─────────────────────────────────────────────────────────────────────────────
# 样式配置
# ─────────────────────────────────────────────────────────────────────────────

ACADEMIC_STYLE = {
    # 页面设置（A4，上下2.5cm，左3cm，右2.5cm）
    'page': {
        'width': Cm(21), 'height': Cm(29.7),
        'top': Cm(2.5), 'bottom': Cm(2.5),
        'left': Cm(3.0), 'right': Cm(2.5),
    },
    # 论文封面标题（Title 样式）：黑体/华文中宋 小二 18pt 居中加粗
    'title': {
        'font_en': 'Times New Roman', 'font_cn': '黑体',
        'size': Pt(18), 'bold': True,
        'align': WD_ALIGN_PARAGRAPH.CENTER,
        'space_before': Pt(0), 'space_after': Pt(12),
        'first_line': None,
    },
    # 章标题 H1：黑体 三号 16pt 居中加粗，段前12pt
    'h1': {
        'font_en': 'Times New Roman', 'font_cn': '黑体',
        'size': Pt(16), 'bold': True,
        'align': WD_ALIGN_PARAGRAPH.CENTER,
        'space_before': Pt(12), 'space_after': Pt(6),
        'first_line': None,
    },
    # 节标题 H2：黑体 四号 14pt 左对齐加粗，段前6pt
    'h2': {
        'font_en': 'Times New Roman', 'font_cn': '黑体',
        'size': Pt(14), 'bold': True,
        'align': WD_ALIGN_PARAGRAPH.LEFT,
        'space_before': Pt(6), 'space_after': Pt(6),
        'first_line': None,
    },
    # 小节标题 H3：黑体 小四 12pt 左对齐加粗，段前6pt
    'h3': {
        'font_en': 'Times New Roman', 'font_cn': '黑体',
        'size': Pt(12), 'bold': True,
        'align': WD_ALIGN_PARAGRAPH.LEFT,
        'space_before': Pt(6), 'space_after': Pt(6),
        'first_line': None,
    },
    # 正文：宋体 小四 12pt，固定行距20pt，首行缩进2字符（24pt）
    'body': {
        'font_en': 'Times New Roman', 'font_cn': '宋体',
        'size': Pt(12), 'bold': False,
        'align': WD_ALIGN_PARAGRAPH.JUSTIFY,
        'space_before': Pt(0), 'space_after': Pt(0),
        'line_spacing': Pt(20),
        'first_line': Pt(24),
    },
}

GONGWEN_STYLE = {
    # 页面设置（A4，GB/T 9704-2012 标准边距）
    'page': {
        'width': Cm(21), 'height': Cm(29.7),
        'top': Cm(3.7), 'bottom': Cm(3.5),
        'left': Cm(2.8), 'right': Cm(2.6),
    },
    # 公文主标题：小标宋 二号 22pt 居中加粗，固定行距28pt
    'title': {
        'font_en': 'FZXiaoBiaoSong-B05S', 'font_cn': 'FZXiaoBiaoSong-B05S',
        'size': Pt(22), 'bold': True,
        'align': WD_ALIGN_PARAGRAPH.CENTER,
        'space_before': Pt(0), 'space_after': Pt(0),
        'line_spacing': Pt(28),
        'first_line': None,
    },
    # 一级标题：黑体 三号 16pt 左对齐，固定行距28pt
    'h1': {
        'font_en': 'SimHei', 'font_cn': '黑体',
        'size': Pt(16), 'bold': False,
        'align': WD_ALIGN_PARAGRAPH.LEFT,
        'space_before': Pt(0), 'space_after': Pt(0),
        'line_spacing': Pt(28),
        'first_line': None,
    },
    # 二级标题：楷体 三号 16pt 左对齐缩进2字符，固定行距28pt
    'h2': {
        'font_en': 'KaiTi_GB2312', 'font_cn': '楷体_GB2312',
        'size': Pt(16), 'bold': False,
        'align': WD_ALIGN_PARAGRAPH.LEFT,
        'space_before': Pt(0), 'space_after': Pt(0),
        'line_spacing': Pt(28),
        'first_line': Pt(32),
    },
    # 三级标题：仿宋加粗 三号 16pt 左对齐缩进2字符，固定行距28pt
    'h3': {
        'font_en': 'FangSong_GB2312', 'font_cn': '仿宋_GB2312',
        'size': Pt(16), 'bold': True,
        'align': WD_ALIGN_PARAGRAPH.LEFT,
        'space_before': Pt(0), 'space_after': Pt(0),
        'line_spacing': Pt(28),
        'first_line': Pt(32),
    },
    # 正文：仿宋 三号 16pt，首行缩进2字符（32pt），固定行距28pt
    'body': {
        'font_en': 'FangSong_GB2312', 'font_cn': '仿宋_GB2312',
        'size': Pt(16), 'bold': False,
        'align': WD_ALIGN_PARAGRAPH.JUSTIFY,
        'space_before': Pt(0), 'space_after': Pt(0),
        'line_spacing': Pt(28),
        'first_line': Pt(32),
    },
}


# ─────────────────────────────────────────────────────────────────────────────
# 段落类型检测
# ─────────────────────────────────────────────────────────────────────────────

# 所有已知的标题样式名称 → 标题级别
# 覆盖 Word（中文版/英文版）、WPS、Google Docs 导出的常见变体
HEADING_LEVEL_MAP = {
    # 英文 Word 默认
    'heading 1': 1, 'heading1': 1,
    'heading 2': 2, 'heading2': 2,
    'heading 3': 3, 'heading3': 3,
    # 中文 Word 数字 ID（最常见）
    '1': 1, '2': 2, '3': 3,
    # 中文样式名
    '标题1': 1, '标题 1': 1, '标题一': 1,
    '标题2': 2, '标题 2': 2, '标题二': 2,
    '标题3': 3, '标题 3': 3, '标题三': 3,
    # WPS 变体
    'heading 1 (body)': 1, 'heading 2 (body)': 2,
}

# 文档标题样式名（用于封面标题 / 公文发文标题）
TITLE_STYLE_NAMES = {'title', '标题', 'document title', '文档标题'}


def detect_paragraph_type(para) -> str:
    """
    检测段落类型，返回 'title' / 'h1' / 'h2' / 'h3' / 'body'。

    检测顺序：
    1. 样式名精确/前缀匹配
    2. XML OutlineLevel 属性
    3. 基于内容的启发式匹配（兜底，处理全用 Normal 样式的文档）
    """
    style_name = (para.style.name or '').lower().strip()

    # 文档标题
    if style_name in TITLE_STYLE_NAMES:
        return 'title'

    # 精确匹配标题级别
    if style_name in HEADING_LEVEL_MAP:
        return f'h{HEADING_LEVEL_MAP[style_name]}'

    # 前缀匹配（例如 "heading 1 char" 也应识别为 h1）
    for pattern, level in HEADING_LEVEL_MAP.items():
        if style_name.startswith(pattern + ' ') or style_name.startswith(pattern + '_'):
            return f'h{level}'

    # 通过 XML OutlineLevel 兜底检测（Word 会在标题段落里设置此属性）
    pPr = para._p.find(qn('w:pPr'))
    if pPr is not None:
        outline = pPr.find(qn('w:outlineLvl'))
        if outline is not None:
            lvl = int(outline.get(qn('w:val'), '9'))
            if lvl == 0:
                return 'h1'
            if lvl == 1:
                return 'h2'
            if lvl == 2:
                return 'h3'

    # 基于内容的启发式检测（处理全用 Normal 样式的中文文档）
    return _detect_by_content(para.text.strip())


# 中文章节编号的正则模式

_H1_PATTERNS = re.compile(
    r'^('
    r'第[一二三四五六七八九十百]+[章节篇部分]'   # 第一章 第二节
    r'|[一二三四五六七八九十]+[、．.]'            # 一、 二、
    r'|\d+\s+[\u4e00-\u9fff]'                    # 1 绪论（数字+空格+中文）
    r'|第\s*\d+\s*[章节]'                         # 第1章 第 1 章
    r')'
)

_H2_PATTERNS = re.compile(
    r'^('
    r'[（(][一二三四五六七八九十]+[）)]'          # （一）(一)
    r'|\d+\.\d+\s'                               # 1.1 + 空格
    r'|\d+\.\d+[\u4e00-\u9fff]'                  # 1.1中文
    r')'
)

_H3_PATTERNS = re.compile(
    r'^('
    r'\d+\.\d+\.\d+'                             # 1.1.1
    r'|[（(]\d+[）)]\s'                           # (1) + 空格
    r')'
)


def _detect_by_content(text: str) -> str:
    """
    通过段落文字内容判断是否为标题。
    仅在样式无法识别时作为兜底，避免误判正文中的序号列表。
    短段落（≤40字）且匹配编号模式才认定为标题。
    """
    if not text or len(text) > 60:
        return 'body'

    if _H1_PATTERNS.match(text):
        return 'h1'
    if _H2_PATTERNS.match(text):
        return 'h2'
    if _H3_PATTERNS.match(text):
        return 'h3'

    return 'body'


# ─────────────────────────────────────────────────────────────────────────────
# 格式化工具函数
# ─────────────────────────────────────────────────────────────────────────────

def apply_run_format(run, fmt: dict):
    """
    对单个 run 应用字体和字号。

    必须通过 XML 操作设置 eastAsia 字体，因为 python-docx 的高层 API
    不直接暴露 w:rFonts 的 eastAsia 槽——而 CJK 字符只认这个槽。
    """
    rPr = run._r.get_or_add_rPr()

    # 先移除会污染结果的直接格式属性
    for tag in ('w:rFonts', 'w:sz', 'w:szCs', 'w:b', 'w:bCs', 'w:color',
                'w:highlight', 'w:shd', 'w:u', 'w:spacing'):
        for el in rPr.findall(qn(tag)):
            rPr.remove(el)

    # 写入新的字体配置（四槽）
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'),    fmt['font_en'])
    rFonts.set(qn('w:hAnsi'),    fmt['font_en'])
    rFonts.set(qn('w:eastAsia'), fmt['font_cn'])   # ← 中文字符专用槽
    rFonts.set(qn('w:cs'),       fmt['font_en'])
    rPr.insert(0, rFonts)

    # 字号（OpenXML 单位：半磅，即 points × 2）
    size_hp = str(int(fmt['size'].pt * 2))
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), size_hp)
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), size_hp)
    rPr.append(sz)
    rPr.append(szCs)

    # 加粗
    if fmt.get('bold'):
        rPr.append(OxmlElement('w:b'))


def apply_paragraph_format(para, fmt: dict):
    """对段落应用对齐、间距、缩进等属性。"""
    pf = para.paragraph_format

    pf.alignment     = fmt.get('align', WD_ALIGN_PARAGRAPH.JUSTIFY)
    pf.space_before  = fmt.get('space_before', Pt(0))
    pf.space_after   = fmt.get('space_after',  Pt(0))

    # 固定行距（如果配置了 line_spacing）
    if 'line_spacing' in fmt:
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing      = fmt['line_spacing']

    # 首行缩进
    pf.first_line_indent = fmt.get('first_line') or Pt(0)
    pf.left_indent       = Pt(0)


def set_page_margins(doc, page_cfg: dict):
    """设置文档所有分节的页面尺寸和边距。"""
    for section in doc.sections:
        section.page_width    = page_cfg['width']
        section.page_height   = page_cfg['height']
        section.top_margin    = page_cfg['top']
        section.bottom_margin = page_cfg['bottom']
        section.left_margin   = page_cfg['left']
        section.right_margin  = page_cfg['right']


# ─────────────────────────────────────────────────────────────────────────────
# 主处理函数
# ─────────────────────────────────────────────────────────────────────────────

def format_document(input_path: str, style_name: str, output_path: str,
                    heading_map: dict[str, str] | None = None) -> dict:
    """
    执行文档格式化，返回统计信息字典。

    heading_map: 由 Claude 语义分析产出的段落类型映射。
        key   = 段落索引（字符串，如 "0", "4"）
        value = 段落类型（"title" / "h1" / "h2" / "h3" / "body"）
        未指定的段落自动检测（样式名 + 内容规则兜底）。
    """
    style_map = ACADEMIC_STYLE if style_name == 'academic' else GONGWEN_STYLE
    heading_map = heading_map or {}

    doc = Document(input_path)

    # 1. 页面设置
    set_page_margins(doc, style_map['page'])

    # 2. 逐段格式化
    stats = {'title': 0, 'h1': 0, 'h2': 0, 'h3': 0, 'body': 0, 'skipped': 0}
    source = {'ai_map': 0, 'auto': 0}  # 统计两种检测路径的使用次数

    for idx, para in enumerate(doc.paragraphs):
        # 优先使用 Claude 给出的语义映射，其次回退到自动检测
        if str(idx) in heading_map:
            para_type = heading_map[str(idx)]
            source['ai_map'] += 1
        else:
            para_type = detect_paragraph_type(para)
            source['auto'] += 1

        fmt = style_map.get(para_type) or style_map['body']
        if para_type not in style_map:
            para_type = 'body'

        apply_paragraph_format(para, fmt)

        if para.runs:
            for run in para.runs:
                apply_run_format(run, fmt)
            stats[para_type] += 1
        elif para.text.strip():
            stats['skipped'] += 1

    doc.save(output_path)
    stats['_source'] = source
    return stats


# ─────────────────────────────────────────────────────────────────────────────
# CLI 入口
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description='将 Word 文档格式化为中文学术论文或公文样式'
    )
    parser.add_argument('--input',  required=True,
                        help='输入 .docx 文件路径')
    parser.add_argument('--style',  required=True,
                        choices=['academic', 'gongwen'],
                        help='目标样式：academic（学术论文）或 gongwen（公文）')
    parser.add_argument('--output', required=True,
                        help='输出 .docx 文件路径')
    parser.add_argument('--heading-map', default=None,
                        help='Claude 语义分析产出的段落类型映射，JSON 字符串。'
                             '例：\'{"0":"title","4":"h1","11":"h1","18":"h2"}\'')
    args = parser.parse_args()

    if not Path(args.input).exists():
        print(f'错误：输入文件不存在：{args.input}', file=sys.stderr)
        sys.exit(1)

    # 解析 heading_map
    heading_map = {}
    if args.heading_map:
        try:
            heading_map = json.loads(args.heading_map)
        except json.JSONDecodeError as e:
            print(f'错误：--heading-map 不是合法 JSON：{e}', file=sys.stderr)
            sys.exit(1)

    style_label = '中文学术论文' if args.style == 'academic' else '中文公文（GB/T 9704）'
    mode_label  = 'AI 语义映射' if heading_map else '自动检测（样式名 + 内容规则）'
    print(f'正在处理：{args.input}')
    print(f'目标样式：{style_label}')
    print(f'检测模式：{mode_label}')

    stats = format_document(args.input, args.style, args.output, heading_map)

    src = stats.get('_source', {})
    print(f'\n✓ 格式化完成：{args.output}')
    print(f'  文档标题段落：{stats["title"]}')
    print(f'  一级标题：{stats["h1"]}，二级标题：{stats["h2"]}，三级标题：{stats["h3"]}')
    print(f'  正文段落：{stats["body"]}')
    if src:
        print(f'  来源：AI映射 {src.get("ai_map",0)} 段，自动检测 {src.get("auto",0)} 段')
    if stats['skipped']:
        print(f'  跳过（无 run 的有文字段落）：{stats["skipped"]}')
    if not heading_map and stats['h1'] == 0 and stats['h2'] == 0 and stats['h3'] == 0:
        print('\n  提示：未检测到任何标题。建议先运行 extract_paragraphs.py，')
        print('  让 Claude 语义分析后用 --heading-map 参数重新处理。')


if __name__ == '__main__':
    main()
