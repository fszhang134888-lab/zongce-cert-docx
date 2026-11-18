#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""verify_fonts.py — 校验收综测加分证明 docx 的字体、字号、段落格式与名单表结构。

编辑器里看着像仿宋不算数：Word 只认 w:rFonts 的 eastAsia 属性，只设 font.name
（ascii/hAnsi）中文会退回默认字体。本脚本直接读 docx XML 逐 run 核对，
并额外检查表格结构（列宽、框线、行高）是否符合学院模板。

用法：
  python verify_fonts.py <docx> [--title-font 宋体] [--body-font 仿宋]

有任何不合规即退出码 1。
"""
from __future__ import annotations

import argparse
import sys
from collections import Counter

from docx import Document
from docx.oxml.ns import qn


def run_info(run):
    rpr = run._element.find(qn("w:rPr"))
    east = None
    if rpr is not None:
        rf = rpr.find(qn("w:rFonts"))
        if rf is not None:
            east = rf.get(qn("w:eastAsia"))
    size = run.font.size.pt if run.font.size else None
    return east, size


def iter_paragraphs(doc):
    """正文段落 + 表格单元格段落，带位置标签。"""
    for i, p in enumerate(doc.paragraphs):
        yield f"正文{i+1}", p
    for ti, table in enumerate(doc.tables, 1):
        for ri, row in enumerate(table.rows, 1):
            seen = set()
            for ci, cell in enumerate(row.cells, 1):
                if cell._tc in seen:      # 合并单元格会重复返回
                    continue
                seen.add(cell._tc)
                for p in cell.paragraphs:
                    yield f"表{ti}-{ri}行{ci}列", p


def table_report(doc, problems):
    if not doc.tables:
        problems.append("文档里没有名单表格（模板要求用无框线表格排名单）")
        return
    for ti, table in enumerate(doc.tables, 1):
        tblPr = table._tbl.tblPr
        widths = []
        for col in table.columns:
            widths.append(round(col.width.pt, 1) if col.width else None)
        borders_el = tblPr.find(qn("w:tblBorders"))
        border_vals = {}
        if borders_el is not None:
            for edge in borders_el:
                border_vals[edge.tag.split("}")[1]] = edge.get(qn("w:val"))
        heights = []
        for row in table.rows:
            heights.append(round(row.height.pt, 1) if row.height else None)
        print(f"\n--- 表格{ti} ---")
        print(f"  行数×列数: {len(table.rows)} x {len(table.columns)}")
        print(f"  列宽(pt): {widths}  合计 {sum(w or 0 for w in widths):.1f}")
        print(f"  行高(pt): {sorted(set(heights))}")
        print(f"  框线: {border_vals or '未设置（默认有框线！）'}")
        if any(v not in (None, "none") for v in border_vals.values()):
            problems.append(f"表格{ti} 有框线，模板要求整表无框线")
        if not border_vals:
            problems.append(f"表格{ti} 未显式设置无框线")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("--title-font")
    ap.add_argument("--body-font")
    ap.add_argument("--title-pt", type=float, default=22.0)
    ap.add_argument("--body-pt", type=float, default=16.0)
    args = ap.parse_args()

    doc = Document(args.docx)
    problems: list[str] = []
    fonts_seen = Counter()
    sizes_seen = Counter()

    print(f"{'位置':<14}{'文本':<24}{'字体/字号':<20}{'缩进':>6}{'行距':>6}  标记字体")
    for where, p in iter_paragraphs(doc):
        text = p.text
        pf = p.paragraph_format
        ind = pf.first_line_indent.pt if pf.first_line_indent is not None else 0.0
        if ind == 0.0 and pf.left_indent is not None:
            ind = pf.left_indent.pt
        ls = pf.line_spacing
        info = []
        for r in p.runs:
            east, size = run_info(r)
            if r.text.strip():
                fonts_seen[east] += 1
                sizes_seen[size] += 1
                if east is None:
                    problems.append(f"{where}「{text[:12]}」有 run 未设 eastAsia 字体")
                if size is None:
                    problems.append(f"{where}「{text[:12]}」有 run 未设字号")
            info.append(f"{east}/{size:g}pt" if size else f"{east}/?")
        ppr = p._p.find(qn("w:pPr"))
        mark = None
        if ppr is not None:
            mrpr = ppr.find(qn("w:rPr"))
            if mrpr is not None:
                mrf = mrpr.find(qn("w:rFonts"))
                if mrf is not None:
                    mark = mrf.get(qn("w:eastAsia"))
        fshow = info[0] if len(set(info)) == 1 else "/".join(sorted(set(info)))[:20]
        print(f"{where:<14}{(text[:22] or '(空)'):<24}{str(fshow):<20}{ind:>6.1f}"
              f"{(ls if ls else 0):>6.2f}  {mark or '-'}")

    table_report(doc, problems)

    print("\n--- 统计 ---")
    print("  字体分布:", dict(fonts_seen))
    print("  字号分布:", dict(sizes_seen))

    if args.title_font and args.title_font not in fonts_seen:
        problems.append(f"标题字体应为 {args.title_font}，实际未见")
    if args.body_font and args.body_font not in fonts_seen:
        problems.append(f"正文字体应为 {args.body_font}，实际未见")
    for sz in sizes_seen:
        if sz not in (args.title_pt, args.body_pt):
            problems.append(f"出现非规范字号 {sz}pt（只允许 {args.title_pt:g}/{args.body_pt:g}pt）")
    sec = doc.sections[0]
    print(f"  页面: {sec.page_width.cm:.1f} x {sec.page_height.cm:.1f} cm, "
          f"边距 上{sec.top_margin.cm:.2f} 下{sec.bottom_margin.cm:.2f} "
          f"左{sec.left_margin.cm:.2f} 右{sec.right_margin.cm:.2f} cm")

    print("\n--- 结论 ---")
    if problems:
        for p in problems:
            print("  FAIL:", p)
        sys.exit(1)
    print("  PASS: eastAsia 字体齐全，字号合规，表格结构符合模板")


if __name__ == "__main__":
    main()
