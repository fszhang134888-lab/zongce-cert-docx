#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""gen_cert.py — 生成"综测加分证明"docx，版式与字体按学院最新模板锁死。

模板依据（2026-11 版《格式示例.docx》，学院下发）：
  页面  A4，上下边距 2.54cm、左右 3.17cm，文档网格 15.6pt/行
  标题  宋体 22pt(二号)，居中，段前 17pt、段后 16.5pt，其下空一行
  正文  仿宋 16pt(三号)，首行缩进 2 字符，左对齐，单倍行距
  名单  无框线表格，第 1 列只放班级、其后每行 4 个姓名
  落款  日期 + 单位，各一段，右对齐

用法：
  python gen_cert.py --demo out.docx                  # 出一份样例
  python gen_cert.py --dump-config cert.json          # 导出配置模板
  python gen_cert.py --config cert.json --out x.docx  # 正式生成
  可选：--preset image 复刻早期截图的黑体外观；--force-canonical-fonts 写死规范字体名

依赖：python-docx、Pillow
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from PIL import ImageFont

# ---------------------------------------------------------------- 版式常量（照官方模板）
PAGE_W_CM, PAGE_H_CM = 21.0, 29.7
MARGIN_TOP_CM = 2.54
MARGIN_BOTTOM_CM = 2.54
MARGIN_LEFT_CM = 3.17
MARGIN_RIGHT_CM = 3.17
DOC_GRID_LINE_PITCH = 312      # twips，15.6pt/行；16pt 正文实际行高 = 2×15.6 = 31.2pt

TITLE_PT = 22.0                # 二号
BODY_PT = 16.0                 # 三号
TITLE_SPACE_BEFORE = 17.0
TITLE_SPACE_AFTER = 16.5
LINE_SPACING = 1.0             # 单倍（配合文档网格 = 31.2pt 行距）
INDENT_CHARS = 2               # 首行缩进 2 字符
CELL_PAD_PT = 5.4              # 单元格左右内边距（108 twips）
# 姓名列宽 = 文字实测宽 + CELL_PAD_PT*2 + 2.0 安全余量（见 plan_columns）
LABEL_COL_MIN_PT = 150.0       # 第 1 列不窄于 5.3cm（官方要求：放宽，只放班级、不许折行）
NAMES_PER_ROW = 4              # 一行 4 个名字

TITLE_CANDIDATES = ["方正小标宋_GBK", "方正小标宋简体", "宋体"]
BODY_CANDIDATES = ["仿宋_GB2312", "仿宋"]
PRESETS = {
    "spec": (TITLE_CANDIDATES, BODY_CANDIDATES),
    "image": (["黑体", "SimHei"], ["黑体", "SimHei"]),
}

DEFAULT_CONFIG = {
    "title": "综测加分证明",
    "body": [
        # 数字与汉字之间不要手打空格：Word 的"自动调整中西文间距"会自己加，
        # 手打会变成双倍间距（官方模板原文同样不带空格）
        # {first} = 名单第一位姓名，{n} = 名单总人数（学院成品写法：兹有陈小明等12位同学…）
        "兹有{first}等{n}位同学于2026年10月25日参加2026年校园文化艺术节开幕式"
        "（活动全称），为活动的开展贡献了积极力量。"
    ],
    "certify": "特此证明。",
    "bonus_line": "名单如下：",
    # 两字姓名中间补全角空格撑成三字宽（学院成品做法）；不想补就改成 false
    "pad_two_char_names": True,
    "roster": [
        {"label": "26电信1班", "names": ["张浩然", "李思远", "王雨桐", "赵梓涵"]},
        {"label": "26电信2班", "names": ["周子墨", "吴俊熙", "郑可欣", "孙嘉懿"]},
        {"label": "26电信3班", "names": ["陈奕辰", "林诗涵", "黄博文", "徐若曦"]},
        {"label": "26电信4班", "names": ["何雨泽", "高梦琪", "罗嘉豪", "谢欣怡"]},
        {"label": "26电信5班", "names": ["唐子涵", "曹梓萱", "邓皓宇", "冯乐瑶"]},
        {"label": "26电信6班", "names": ["蒋思彤", "韩明轩", "沈可怡", "卢俊杰"]},
    ],
    "sign": {
        "date": "2026年11月4日",
        "unit": "广东技术师范大学集成电路与信息工程学院团委",
    },
    "layout": {
        "margin_top_cm": MARGIN_TOP_CM,
        "margin_bottom_cm": MARGIN_BOTTOM_CM,
        "margin_left_cm": MARGIN_LEFT_CM,
        "margin_right_cm": MARGIN_RIGHT_CM,
        "line_spacing": LINE_SPACING,
        "first_line_indent_chars": INDENT_CHARS,
        "blank_after_title": 1,
        "blank_before_bonus": 2,
        "blank_before_sign": 2,
        "title_space_before_pt": TITLE_SPACE_BEFORE,
        "title_space_after_pt": TITLE_SPACE_AFTER,
        "names_per_row": NAMES_PER_ROW,   # 0 = 按可用宽度自动降级
    },
}


# ------------------------------------------------------------ 字体探测
def _font_registry() -> dict:
    """返回 {字体显示名: 文件绝对路径}，合并 HKLM / HKCU。

    中文 Windows 的注册表里 simsun.ttc 记作 "SimSun & NSimSun"，所以把
    "&" 分隔的英文别名也一并登记，才能用中文名查到文件。
    """
    out: dict[str, str] = {}
    try:
        import winreg
    except ImportError:
        return out
    for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        sub = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
        try:
            key = winreg.OpenKey(root, sub)
        except OSError:
            continue
        i = 0
        while True:
            try:
                name, value, _ = winreg.EnumValue(key, i)
            except OSError:
                break
            i += 1
            if not isinstance(value, str):
                continue          # 注册表里混有 DWORD/MULTI_SZ 项
            disp = re.sub(r"\s*\((TrueType|OpenType|Type 1|All res)\)\s*$", "", name).strip()
            if not os.path.isabs(value):
                value = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", value)
            if not os.path.exists(value):
                continue
            out.setdefault(disp, value)
            for alias in disp.split("&"):
                out.setdefault(alias.strip(), value)
    for cn, en in {"宋体": "SimSun", "仿宋": "FangSong", "黑体": "SimHei",
                   "楷体": "KaiTi", "微软雅黑": "Microsoft YaHei"}.items():
        if cn not in out and en in out:
            out[cn] = out[en]
    return out


def resolve_font(candidates: list[str], override: str | None = None) -> tuple[str, str]:
    """按候选顺序取第一个本机已安装的字体，返回 (字体名, 文件路径)。"""
    reg = _font_registry()
    pool = [override] if override else candidates
    for name in pool:
        if name in reg:
            return name, reg[name]
    for name in reversed(candidates):
        if name in reg:
            if override:
                print(f"[warn] 本机未安装「{override}」，实际使用「{name}」"
                      f"（要写死规范字体名加 --force-canonical-fonts）")
            return name, reg[name]
    raise SystemExit("[error] 找不到宋体/仿宋，无法生成（缺中文字体的环境不支持本 skill）")


def text_width_pt(text: str, font_path: str, size_pt: float) -> float:
    """用真实字体度量文本宽度（pt），比按字符数估算可靠。"""
    f = ImageFont.truetype(font_path, int(round(size_pt * 8)))
    return f.getlength(text) / 8.0


# ------------------------------------------------------------ docx 工具
def set_run(run, font_name: str, size_pt: float, bold: bool | None = None):
    """给 run 同时设 ascii/hAnsi/eastAsia/cs —— 只设 font.name 中文不生效。"""
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    if bold is not None:
        run.font.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), font_name)
    szcs = rpr.find(qn("w:szCs"))
    if szcs is None:
        szcs = OxmlElement("w:szCs")
        rpr.append(szcs)
    szcs.set(qn("w:val"), str(int(round(size_pt * 2))))


def set_par_mark(p, font_name: str, size_pt: float):
    """设定段落标记的字体字号——空段落的高度由它决定。"""
    ppr = p._p.get_or_add_pPr()
    rpr = ppr.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        ppr.append(rpr)
    for tag in ("w:rFonts", "w:sz", "w:szCs"):
        old = rpr.find(qn(tag))
        if old is not None:
            rpr.remove(old)
    rfonts = OxmlElement("w:rFonts")
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), font_name)
    rpr.append(rfonts)
    for tag in ("w:sz", "w:szCs"):
        el = OxmlElement(tag)
        el.set(qn("w:val"), str(int(round(size_pt * 2))))
        rpr.append(el)


def set_doc_grid(doc, line_pitch: int = DOC_GRID_LINE_PITCH):
    """写入"只指定行网格"，与官方模板一致；16pt 正文因此占 2 格 = 31.2pt 行高。"""
    for sect in doc.sections:
        sect_el = sect._sectPr
        old = sect_el.find(qn("w:docGrid"))
        if old is not None:
            sect_el.remove(old)
        grid = OxmlElement("w:docGrid")
        grid.set(qn("w:type"), "lines")
        grid.set(qn("w:linePitch"), str(line_pitch))
        grid.set(qn("w:charSpace"), "0")
        sect_el.append(grid)


def enable_table_grid(doc):
    """加 w:adjustLineHeightInTable 兼容标志：让表格内段落也套用文档网格。

    现代模式（compatibilityMode 14+）下 Word 表格不套网格，行高会掉到
    20.8pt；学院下发/产出的老 .doc 体系都带这枚 Word 2003 兼容标志，
    表格行距才是 31.2。实测在 mode 14 文档上加这枚标志同样生效。
    保留 trHeight=624 atLeast 作双保险：即使标志被剥离，行高也不掉。
    """
    settings = doc.settings.element
    compat = settings.find(qn("w:compat"))
    if compat is None:
        compat = OxmlElement("w:compat")
        settings.append(compat)
    if compat.find(qn("w:adjustLineHeightInTable")) is None:
        compat.insert(0, OxmlElement("w:adjustLineHeightInTable"))


def pad_two_char(names: list[str]) -> list[str]:
    """两字姓名中间补一个全角空格，撑成三字宽。

    学院成品的标准做法（选手/嘉宾/观众三份名单全部如此：林　霈、肖　菲、
    罗　洁、彭　冰），对应说明 PDF 的"姓名只有两个字则单独选中，调整宽度"。
    """
    return [n if len(n) != 2 else n[0] + "\u3000" + n[1] for n in names]


# pPr 里排在 autoSpaceDE/DN 之后的元素
_PPR_AFTER_AUTOSPACE = (
    "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
    "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
    "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
)


def disable_auto_space(p):
    """关掉"自动调整中文与西文/数字的间距"，让表格列宽与实测宽度一致。"""
    ppr = p._p.get_or_add_pPr()
    for tag in ("w:autoSpaceDE", "w:autoSpaceDN"):
        old = ppr.find(qn(tag))
        if old is not None:
            ppr.remove(old)
        el = OxmlElement(tag)
        el.set(qn("w:val"), "0")
        anchor = None
        for child in ppr:
            if child.tag in {qn(t) for t in _PPR_AFTER_AUTOSPACE}:
                anchor = child
                break
        if anchor is None:
            ppr.append(el)
        else:
            anchor.addprevious(el)


def add_blank(doc, font_name: str, size_pt: float, line_spacing: float):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = line_spacing
    set_par_mark(p, font_name, size_pt)
    return p


def add_text_par(doc, text, font_name, size_pt, line_spacing,
                 indent_chars=0.0, align=WD_ALIGN_PARAGRAPH.LEFT):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = line_spacing
    pf.alignment = align
    if indent_chars:
        pf.first_line_indent = Pt(size_pt * indent_chars)
    run = p.add_run(text)
    set_run(run, font_name, size_pt, bold=False)
    set_par_mark(p, font_name, size_pt)
    return p


# ------------------------------------------------------------ 表格（名单）
_TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
                "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd",
                "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
                "tblCaption", "tblDescription"]


def _tblpr_insert(tblPr, el):
    """按 OOXML schema 顺序插入 tblPr 子元素，避免 WPS 解析异常。"""
    tag = el.tag.split("}")[1]
    idx = _TBLPR_ORDER.index(tag)
    for child in list(tblPr):
        ctag = child.tag.split("}")[1]
        if ctag in _TBLPR_ORDER and _TBLPR_ORDER.index(ctag) > idx:
            child.addprevious(el)
            return
    tblPr.append(el)


def _no_borders(table):
    """整表无框线（官方模板要求）。"""
    tblPr = table._tbl.tblPr
    old = tblPr.find(qn("w:tblBorders"))
    if old is not None:
        tblPr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "auto")
        borders.append(el)
    _tblpr_insert(tblPr, borders)


def _cell_margin(table, left_twips=108, right_twips=108):
    tblPr = table._tbl.tblPr
    old = tblPr.find(qn("w:tblCellMar"))
    if old is not None:
        tblPr.remove(old)
    mar = OxmlElement("w:tblCellMar")
    for side, val in (("top", 0), ("left", left_twips), ("bottom", 0), ("right", right_twips)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    _tblpr_insert(tblPr, mar)


def plan_columns(roster, font_file, size_pt, text_w_pt, names_per_row_cfg=0):
    """算列数与各列宽度：第 1 列放班级，其后 k 列放姓名。

    返回 (k, [col_width_pt...])，宽度之和等于版心宽度。
    宽度预算 = 文字实测宽 + 单元格左右边距(2×5.4) + 安全余量——余量不能省，
    否则 Word 渲染时的取整误差会让"26汉语言文学1班"这类长班名折行。
    """
    labels = [g.get("label", "") for g in roster if g.get("label")]
    label_need = max([text_width_pt(lb, font_file, size_pt) + 32 + CELL_PAD_PT * 2 + 4.0
                      for lb in labels] + [LABEL_COL_MIN_PT])
    # 指定的 names_per_row 只是偏好（官方"一行一般放 4 个"）；放不下就降级，
    # 而不是直接塌到 1 列——学院实际成品也常用一行 2~3 个名字
    pref = [names_per_row_cfg] if names_per_row_cfg else []
    cand = pref + [k for k in (5, 4, 3, 2, 1) if k not in pref]
    for k in cand:
        if k < 1:
            continue
        need = [0.0] * k
        for g in roster:
            for i, n in enumerate(g.get("names", [])):
                need[i % k] = max(need[i % k], text_width_pt(n, font_file, size_pt) + CELL_PAD_PT * 2 + 2.0)
        need = [max(v, BODY_PT * 3 + 10.8 + 2.0) for v in need]
        total = label_need + sum(need)
        if total <= text_w_pt or k == 1:
            widths = [label_need] + need
            slack = text_w_pt - sum(widths)
            if slack > 0:                    # 富余宽度均分给姓名列
                for i in range(1, len(widths)):
                    widths[i] += slack / (len(widths) - 1)
            elif slack < -0.5:
                print(f"[warn] 名单内容超出 {abs(slack):.1f}pt，k={k} 时可能折行")
            return k, widths
    return 1, [text_w_pt]


def build_roster_table(doc, roster, font_name, font_file, size_pt, line_spacing,
                       text_w_pt, names_per_row_cfg=0):
    """生成无框线名单表格，返回 (行数, 列数, 列宽列表)。"""
    groups = [g for g in roster if g.get("names")]
    if not groups:
        return 0, 0, []
    k, widths = plan_columns(groups, font_file, size_pt, text_w_pt, names_per_row_cfg)
    n_cols = k + 1

    layout_rows = []
    for g in groups:
        names = g["names"]
        for ci in range(0, len(names), k):
            layout_rows.append((g.get("label", "") if ci == 0 else "", names[ci:ci + k]))

    table = doc.add_table(rows=len(layout_rows), cols=n_cols)
    table.autofit = False
    _no_borders(table)
    _cell_margin(table)

    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        _tblpr_insert(tblPr, tblW)
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(int(round(text_w_pt * 20))))

    for ci, w in enumerate(widths):
        table.columns[ci].width = Pt(w)
        for cell in table.columns[ci].cells:
            cell.width = Pt(w)

    for row in table.rows:
        row.height = Pt(BODY_PT * 1.95)          # 31.2pt：官方表格实测行高
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

    for ri, (label, names) in enumerate(layout_rows):
        for ci in range(n_cols):
            cell = table.cell(ri, ci)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            pf = p.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = None               # 继承网格，别写死单倍
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            if ci == 0:
                pf.left_indent = Pt(size_pt * INDENT_CHARS)   # 与"名单如下："左对齐
                text = label
            else:
                text = names[ci - 1] if ci - 1 < len(names) else ""
            disable_auto_space(p)
            run = p.add_run(text)
            set_run(run, font_name, size_pt, bold=False)
            set_par_mark(p, font_name, size_pt)
    return len(layout_rows), n_cols, widths


# ------------------------------------------------------------ 主流程
def build(cfg: dict, out_path: str, title_font: tuple, body_font: tuple,
          force_canonical: bool = False):
    lay = cfg.setdefault("layout", {})
    m_top = lay.get("margin_top_cm", MARGIN_TOP_CM)
    m_bot = lay.get("margin_bottom_cm", MARGIN_BOTTOM_CM)
    m_left = lay.get("margin_left_cm", MARGIN_LEFT_CM)
    m_right = lay.get("margin_right_cm", MARGIN_RIGHT_CM)
    ls = lay.get("line_spacing", LINE_SPACING)
    ind_chars = lay.get("first_line_indent_chars", INDENT_CHARS)
    blank_title = lay.get("blank_after_title", 1)
    blank_bonus = lay.get("blank_before_bonus", 2)
    blank_sign = lay.get("blank_before_sign", 2)
    title_before = lay.get("title_space_before_pt", TITLE_SPACE_BEFORE)
    title_after = lay.get("title_space_after_pt", TITLE_SPACE_AFTER)

    title_name, title_file = title_font
    body_name, body_file = body_font
    if force_canonical:
        title_name = TITLE_CANDIDATES[0]
        body_name = BODY_CANDIDATES[0]

    text_w_pt = (PAGE_W_CM - m_left - m_right) * 72 / 2.54

    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(PAGE_W_CM)
    sec.page_height = Cm(PAGE_H_CM)
    sec.top_margin = Cm(m_top)
    sec.bottom_margin = Cm(m_bot)
    sec.left_margin = Cm(m_left)
    sec.right_margin = Cm(m_right)
    set_doc_grid(doc)
    enable_table_grid(doc)

    normal = doc.styles["Normal"]
    normal.font.name = body_name
    normal.font.size = Pt(BODY_PT)
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), body_name)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing = ls

    # 标题：二号，居中，段前 17 / 段后 16.5
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.space_before = Pt(title_before)
    pf.space_after = Pt(title_after)
    pf.line_spacing = 1.0
    run = p.add_run(cfg.get("title", "综测加分证明"))
    set_run(run, title_name, TITLE_PT, bold=False)
    set_par_mark(p, title_name, TITLE_PT)

    for _ in range(blank_title):
        add_blank(doc, body_name, BODY_PT, ls)

    # 正文里的 {n} 自动替换为名单总人数、{first} 替换为第一位姓名，
    # 避免"兹有 N 位"与名单对不上（学院成品写法：兹有陈小明等12位同学…）
    raw_roster = [g for g in cfg.get("roster", []) if g.get("names")]
    n_names = sum(len(g["names"]) for g in raw_roster)
    first_name = raw_roster[0]["names"][0] if raw_roster else ""

    def fill(s: str) -> str:
        return (s.replace("{n}", str(n_names))
                 .replace("{N}", str(n_names))
                 .replace("{first}", first_name))

    # 两字姓名补全角空格（学院做法），放在列宽规划之前，规划才能量到真实宽度
    roster = raw_roster
    if cfg.get("pad_two_char_names", True):
        roster = [{**g, "names": pad_two_char(g["names"])} for g in raw_roster]

    for para in cfg.get("body", DEFAULT_CONFIG["body"]):
        add_text_par(doc, fill(para), body_name, BODY_PT, ls, ind_chars)

    if cfg.get("certify"):
        add_text_par(doc, cfg["certify"], body_name, BODY_PT, ls, ind_chars)

    for _ in range(blank_bonus):
        add_blank(doc, body_name, BODY_PT, ls)

    if cfg.get("bonus_line"):
        add_text_par(doc, fill(cfg["bonus_line"]), body_name, BODY_PT, ls, ind_chars)

    rows, cols, widths = build_roster_table(
        doc, roster, body_name, body_file, BODY_PT, ls, text_w_pt,
        lay.get("names_per_row", NAMES_PER_ROW),
    )

    for _ in range(blank_sign):
        add_blank(doc, body_name, BODY_PT, ls)

    sign = cfg.get("sign", {})
    for key in ("date", "unit"):
        if sign.get(key):
            add_text_par(doc, sign[key], body_name, BODY_PT, ls, 0,
                         align=WD_ALIGN_PARAGRAPH.RIGHT)

    doc.save(out_path)
    print(f"OK {out_path}")
    print(f"  标题字体: {title_name} {TITLE_PT:.0f}pt")
    print(f"  正文字体: {body_name} {BODY_PT:.0f}pt  (度量文件 {os.path.basename(body_file)})")
    print(f"  页边距: 上{m_top} 下{m_bot} 左{m_left} 右{m_right} cm  版心宽 {text_w_pt:.1f}pt")
    if rows:
        print(f"  名单表: {rows} 行 x {cols} 列，列宽 " + "/".join(f"{w:.1f}" for w in widths) + " pt")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config")
    ap.add_argument("--out")
    ap.add_argument("--demo", action="store_true", help="用内置样例数据生成")
    ap.add_argument("--dump-config", metavar="PATH")
    ap.add_argument("--title-font")
    ap.add_argument("--body-font")
    ap.add_argument("--preset", choices=["spec", "image"], default="spec",
                    help="spec=模板要求的宋体+仿宋（默认）; image=复刻早期截图的黑体外观")
    ap.add_argument("--force-canonical-fonts", action="store_true")
    args = ap.parse_args()

    if args.dump_config:
        with open(args.dump_config, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)
        print(f"已写出配置模板: {args.dump_config}")
        return

    if args.config:
        with open(args.config, encoding="utf-8") as f:
            cfg = json.load(f)
        for k, v in DEFAULT_CONFIG.items():
            cfg.setdefault(k, v)
        merged = dict(DEFAULT_CONFIG["layout"])
        merged.update(cfg.get("layout") or {})
        cfg["layout"] = merged
    elif args.demo:
        cfg = json.loads(json.dumps(DEFAULT_CONFIG))
    else:
        ap.error("需要 --config 或 --demo")

    out = args.out or (os.path.splitext(args.config)[0] + ".docx" if args.config else "综测加分证明样例.docx")
    title_pool, body_pool = PRESETS[args.preset]
    build(cfg, out, resolve_font(title_pool, args.title_font),
          resolve_font(body_pool, args.body_font), args.force_canonical_fonts)


if __name__ == "__main__":
    main()
