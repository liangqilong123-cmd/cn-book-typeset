#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""build_book —— 中文书籍 Markdown → 印刷级 DOCX

    标准 32 开（147×210mm）· 镜像边距 · 奇偶页独立页眉 · 自动目录域 · 区级页码

核心版式为真实书籍规格，不是「Word 文档加个标题」：

    · 镜像边距（mirrorMargins）——内边距宽、外边距窄，装订侧永远在里
    · 奇偶页页眉分开 —— 奇数页印章名，偶数页印书名（出版惯例）
    · 页码贴在切口侧 —— 奇数页右下、偶数页左下
    · 每章独立小节 —— 章起新页、页眉随章变、首章页码重启
    · 目录用 Word 域 —— 打开后全选按 F9 自动生成页码

用法：

    python build_book.py --config examples/book.yaml            # 全本
    python build_book.py --config examples/book.yaml --sample   # 样张（只取第一章）

配置见 examples/book.yaml，字段说明见 docs/32开版式规格.md。
"""

import argparse
import os
import re
import sys
from pathlib import Path

import yaml
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt, RGBColor

from pii_guard import audit, mask

# ────────────────────────────── 默认版式 ──────────────────────────────
DEFAULTS = {
    'page_width_mm': 147,      # 32 开成品宽
    'page_height_mm': 210,     # 32 开成品高
    'top_margin_cm': 2.0,
    'bottom_margin_cm': 1.8,
    'inner_margin_cm': 2.0,    # 装订侧
    'outer_margin_cm': 1.6,    # 切口侧
    'body_font': '宋体',
    'latin_font': 'Times New Roman',
    'body_size_pt': 10.5,      # 五号
    'line_spacing': 1.5,
    'h1_font': '黑体',
    'h1_size_pt': 16,
    'h2_font': '黑体',
    'h2_size_pt': 13,
    'quote_font': '楷体',      # 题词/引文短行用楷体区分
    'header_font_size_pt': 9,
    'header_color': '555555',
    'footer_font_size_pt': 9,
    'image_width_mm': 96,
    'short_line_max_chars': 14,   # 不超过这么长、且无句末标点 → 判为题词
}

ENDS = ('。', '；', '，', '：', '"', '"', '）', ')', '！', '？', '…', '”', '’')


class Layout:
    """版式参数集合，缺项自动落回 DEFAULTS"""

    def __init__(self, raw=None):
        raw = raw or {}
        for k, v in DEFAULTS.items():
            setattr(self, k, raw.get(k, v))
        # 手工覆盖项（未在 DEFAULTS 里的也允许）
        for k, v in raw.items():
            if not hasattr(self, k):
                setattr(self, k, v)


# ────────────────────────────── 基础写入 ──────────────────────────────
def set_run(run, cn, en, size, bold=False, color=None):
    run.font.name = en
    run.font.size = Pt(size)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), cn)
    if color:
        run.font.color.rgb = RGBColor(*color)


def add_para(doc, text='', L=None, font=None, size=None, bold=False,
             align=None, indent=True, space_before=0, space_after=0):
    L = L or Layout()
    p = doc.add_paragraph()
    pf = p.paragraph_format
    if align is not None:
        pf.alignment = align
    size = size or L.body_size_pt
    if indent and align is None:
        pf.first_line_indent = Pt(size * 2)     # 首行缩进两个字符
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = L.line_spacing
    if text:
        set_run(p.add_run(text), font or L.body_font, L.latin_font, size, bold)
    return p


def add_page_break(doc):
    doc.add_page_break()


def add_field(paragraph, instr, placeholder='', size=10.5, L=None):
    """在段落里插入一个 Word 域（PAGE / TOC 都用它）"""
    L = L or Layout()
    run = paragraph.add_run()
    set_run(run, L.body_font, L.latin_font, size)
    begin = OxmlElement('w:fldChar')
    begin.set(qn('w:fldCharType'), 'begin')
    it = OxmlElement('w:instrText')
    it.set(qn('xml:space'), 'preserve')
    it.text = f' {instr} '
    sep = OxmlElement('w:fldChar')
    sep.set(qn('w:fldCharType'), 'separate')
    tip = OxmlElement('w:t')
    tip.text = placeholder
    end = OxmlElement('w:fldChar')
    end.set(qn('w:fldCharType'), 'end')
    for el in (begin, it, sep, tip, end):
        run._element.append(el)


def add_page_number(para, align, L):
    para.alignment = align
    add_field(para, 'PAGE', size=L.footer_font_size_pt, L=L)


def _header_rule(para):
    """页眉下的一条细灰线"""
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '4')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '999999')
    pBdr.append(bottom)
    pPr.append(pBdr)


def write_header(para, text, L):
    para.text = ''
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    gray = tuple(int(L.header_color[i:i + 2], 16) for i in (0, 2, 4))
    set_run(para.add_run(text), L.h1_font, L.latin_font,
            L.header_font_size_pt, color=gray)
    _header_rule(para)


def setup_section(doc, sec, L, odd_header, even_header, restart_number=False):
    """给一节套上书籍版式：镜像边距 + 奇偶页分别的页眉页脚

    踩坑记录：doc.add_section() 新建的节会**继承**上一节的 w:pgNumType
    （带 start=1），所以续编页码时必须显式删掉它，否则每章页码都重新从 1 开始。
    """
    sec.page_width = Mm(L.page_width_mm)
    sec.page_height = Mm(L.page_height_mm)
    sec.top_margin = Cm(L.top_margin_cm)
    sec.bottom_margin = Cm(L.bottom_margin_cm)
    sec.left_margin = Cm(L.inner_margin_cm)   # 镜像后 = 内侧
    sec.right_margin = Cm(L.outer_margin_cm)  # 镜像后 = 外侧

    settings = doc.settings.element
    if settings.find(qn('w:mirrorMargins')) is None:
        settings.append(OxmlElement('w:mirrorMargins'))
    doc.settings.odd_and_even_pages_header_footer = True

    sectPr = sec._sectPr
    pg = sectPr.find(qn('w:pgNumType'))
    if restart_number:
        if pg is None:
            pg = OxmlElement('w:pgNumType')
            sectPr.append(pg)
        pg.set(qn('w:start'), '1')
    else:
        if pg is not None:
            sectPr.remove(pg)

    sec.header.is_linked_to_previous = False
    write_header(sec.header.paragraphs[0], odd_header, L)
    sec.even_page_header.is_linked_to_previous = False
    write_header(sec.even_page_header.paragraphs[0], even_header, L)

    sec.footer.is_linked_to_previous = False
    add_page_number(sec.footer.paragraphs[0], WD_ALIGN_PARAGRAPH.RIGHT, L)
    sec.even_page_footer.is_linked_to_previous = False
    add_page_number(sec.even_page_footer.paragraphs[0], WD_ALIGN_PARAGRAPH.LEFT, L)


def setup_front_section(doc, L):
    """前置页（扉页/说明/目录）：同纸张、无页眉、无页码"""
    sec = doc.sections[0]
    sec.page_width = Mm(L.page_width_mm)
    sec.page_height = Mm(L.page_height_mm)
    sec.top_margin = Cm(L.top_margin_cm)
    sec.bottom_margin = Cm(L.bottom_margin_cm)
    sec.left_margin = Cm(L.inner_margin_cm)
    sec.right_margin = Cm(L.outer_margin_cm)
    sec.header.is_linked_to_previous = False
    sec.header.paragraphs[0].text = ''
    sec.footer.is_linked_to_previous = False
    sec.footer.paragraphs[0].text = ''


# ────────────────────────────── Markdown 解析 ──────────────────────────────
def parse_md(path, L, whitelist=()):
    """把 md 拆成 [(类型, 内容)]，类型为 h1 / h2 / center / img / para

    识别规则（刻意保守，宁可多判正文）：
        # / ## / ###  → 标题（### 并入二级，书里很少用三级）
        ![alt](src)  → 插图
        > 引用 / - 列表 → 正文段
        短行（≤ short_line_max_chars 且无句末标点）→ 题词，用楷体居中
    """
    text = Path(path).read_text(encoding='utf-8-sig')
    items = []
    for raw in text.splitlines():
        s = raw.rstrip()
        if not s.strip():
            continue
        if s.startswith('### '):
            items.append(('h2', s[4:].strip()))
        elif s.startswith('## '):
            items.append(('h2', s[3:].strip()))
        elif s.startswith('# '):
            items.append(('h1', s[2:].strip()))
        elif s.lstrip().startswith('!['):
            m = re.search(r'\(([^)]+)\)', s)
            if m:
                items.append(('img', m.group(1)))
        elif s.startswith('> '):
            items.append(('para', s[2:].strip()))
        elif s.startswith('- '):
            items.append(('para', '　' + s[2:].strip()))
        else:
            t = s.strip()
            if len(t) <= L.short_line_max_chars and not t.endswith(ENDS):
                items.append(('center', t))
            else:
                items.append(('para', t))

    # 统一脱敏（白名单除外）
    return [(typ, mask(val, whitelist)[0]) for typ, val in items]


def first_h1(path, L, whitelist=()):
    for typ, val in parse_md(path, L, whitelist):
        if typ == 'h1':
            return val
    return None


# ────────────────────────────── 渲染 ──────────────────────────────
def render_item(doc, typ, payload, L, img_dir):
    if typ == 'h1':
        p = doc.add_paragraph(style='Heading 1')
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.page_break_before = True   # 一级标题另起一页
        p.paragraph_format.space_before = Pt(24)
        p.paragraph_format.space_after = Pt(18)
        p.paragraph_format.line_spacing = L.line_spacing
        set_run(p.add_run(payload), L.h1_font, L.latin_font, L.h1_size_pt, bold=True)
    elif typ == 'h2':
        p = doc.add_paragraph(style='Heading 2')
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.line_spacing = L.line_spacing
        set_run(p.add_run(payload), L.h2_font, L.latin_font, L.h2_size_pt, bold=True)
    elif typ == 'center':
        add_para(doc, payload, L=L, font=L.quote_font,
                 align=WD_ALIGN_PARAGRAPH.CENTER, indent=False,
                 space_before=2, space_after=2)
    elif typ == 'img':
        base = os.path.splitext(os.path.basename(payload))[0]
        fp = os.path.join(img_dir, base + '.png') if img_dir else None
        if fp and os.path.exists(fp):
            pic = doc.add_paragraph()
            pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pic.paragraph_format.space_before = Pt(10)
            pic.paragraph_format.space_after = Pt(2)
            pic.add_run().add_picture(fp, width=Mm(L.image_width_mm))
            num = re.search(r'(\d+)', base)
            cap = doc.add_paragraph()
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap.paragraph_format.space_after = Pt(10)
            set_run(cap.add_run(f'图 {num.group(1) if num else base}'),
                    L.h2_font, L.latin_font, L.body_size_pt, bold=True)
        else:
            add_para(doc, f'〔缺图：{base}〕', L=L, size=10, bold=True,
                     align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
    else:
        add_para(doc, payload, L=L)


# ────────────────────────────── 主流程 ──────────────────────────────
def build(cfg_path, sample=False, out_path=None):
    cfg = yaml.safe_load(Path(cfg_path).read_text(encoding='utf-8'))
    base = Path(cfg_path).parent
    book = cfg.get('book', {})
    L = Layout(cfg.get('layout'))
    pii = cfg.get('pii', {}) or {}
    whitelist = set(pii.get('whitelist') or [])
    placeholder = pii.get('placeholder', '〔联系方式已隐去〕')
    src_dir = base / (cfg.get('source_dir') or '.')

    chapters = cfg.get('chapters') or []
    if not chapters:
        chapters = [{'file': p.name} for p in sorted(src_dir.glob('*.md'))
                    if not p.name.startswith('00_')]
    if sample:
        chapters = chapters[:1]

    doc = Document()
    setup_front_section(doc, L)

    # 前置页文字同样要过脱敏（否则只能靠最后的自检兜底，白排一遍）
    def m(s):
        return mask(s, whitelist, placeholder)[0] if isinstance(s, str) else s

    # ---- 扉页 ----
    for _ in range(5):
        add_para(doc, '', L=L, indent=False)
    add_para(doc, m(book.get('title', '书 名')), L=L, font=L.h1_font, size=36,
             bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
    if book.get('subtitle'):
        add_para(doc, '', L=L, indent=False)
        add_para(doc, m(book['subtitle']), L=L, size=14,
                 align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
    for _ in range(4):
        add_para(doc, '', L=L, indent=False)
    if book.get('edition'):
        add_para(doc, m(book['edition']), L=L,
                 align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
    add_page_break(doc)

    # ---- 出版说明 ----
    notes = book.get('publisher_notes') or []
    if notes:
        add_para(doc, m(book.get('notes_heading', '出 版 说 明')), L=L, font=L.h1_font,
                 size=15, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                 indent=False, space_after=16)
        for t in notes:
            add_para(doc, m(t), L=L, space_after=8)
        add_page_break(doc)

    # ---- 目录（Word 域，打开后 Ctrl+A → F9 生成页码）----
    add_para(doc, m(book.get('toc_heading', '目  录')), L=L, font=L.h1_font, size=15,
             bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, space_after=16)
    toc_para = doc.add_paragraph()
    add_field(toc_para, r'TOC \o "1-2" \h \z \u',
              placeholder='（目录域：全选后按 F9 更新页码）', L=L)
    add_page_break(doc)

    # ---- 正文：每章一个独立小节 ----
    even_head = m(book.get('running_head_even') or book.get('title', ''))
    for idx, ch in enumerate(chapters):
        path = src_dir / ch['file']
        if not path.exists():
            print(f'  ! 跳过（不存在）：{path}', file=sys.stderr)
            continue
        head = first_h1(path, L, whitelist) or ch.get('title') or path.stem
        odd_head = m((book.get('running_head_odd_template') or '{book} · {chapter}').format(
            book=book.get('title', ''), chapter=head))

        # 每章开新节；第一章重启页码（前置页不计入正文页码）
        sec = doc.add_section(WD_SECTION.NEW_PAGE)
        setup_section(doc, sec, L, odd_head, even_head, restart_number=(idx == 0))
        for typ, payload in parse_md(path, L, whitelist):
            render_item(doc, typ, payload, L, img_dir=ch.get('image_dir'))

    # ---- 落盘前自检：隐私信息不得漏进印张 ----
    plain = re.sub(r'<[^>]+>', '', doc.element.xml)
    leaks = audit(plain, whitelist)
    must = cfg.get('verify', {}).get('must_contain') or []
    missing = [m for m in must if m not in plain]
    if leaks:
        raise SystemExit(f'✗ 印张仍含联系方式，已中止保存：{leaks[:3]}')
    if missing:
        raise SystemExit(f'✗ 必含内容缺失（可能是免责声明被删）：{missing}')

    out = Path(out_path) if out_path else base / (
        book.get('output') or cfg.get('output') or 'book.docx')
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    print(f'✓ 已生成：{out}')
    print(f'  章节 {len(chapters)} 篇 · 隐私自检通过'
          f'{(" · 必含校验通过") if must else ""}')
    return out


def main():
    ap = argparse.ArgumentParser(description='Markdown → 32 开印刷版 DOCX')
    ap.add_argument('--config', '-c', required=True, help='配置文件（YAML）')
    ap.add_argument('--sample', action='store_true', help='只排第一章做样张')
    ap.add_argument('--out', '-o', help='输出路径，默认取配置里的 book.output')
    args = ap.parse_args()
    build(args.config, sample=args.sample, out_path=args.out)


if __name__ == '__main__':
    main()
