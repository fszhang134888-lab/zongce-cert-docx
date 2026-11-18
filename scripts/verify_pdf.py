#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""verify_pdf.py — 交付前验收：PDF 里"到底用了什么字体"+"名单表格有没有对齐、有没有折行"。

Word 里看着对不等于交上去对。本脚本直接读 PDF 的字体资源和字符坐标：
  1. 字体清单：打印实际字体名 / 字号，证明东亚字体没被替换
  2. 名单表：把每个单元格文字与 docx 表格的列坐标比对（偏差 < 1.5pt），
     并检查行数一致——PDF 行数变多说明有姓名折行（模板明确要求一行放下）
  3. 落款两行右边界是否齐

用法：
  python verify_pdf.py <pdf> --docx <docx> [--title-font 宋体] [--body-font 仿宋]

依赖：pymupdf、python-docx
"""
from __future__ import annotations

import argparse
import sys

import pymupdf

PDF_NAME = {
    "宋体": "SimSun", "仿宋": "FangSong", "仿宋_GB2312": "FangSong",
    "黑体": "SimHei", "楷体": "KaiTi", "微软雅黑": "MicrosoftYaHei",
    "方正小标宋_GBK": "FZXiaoBiaoSong", "方正小标宋简体": "FZXiaoBiaoSong",
}

CELL_GAP_PT = 5.0     # 字符间隙 > 5pt 视为换格
COL_TOL_PT = 2.5      # 单元格起点与列位允许偏差
FIRST_COL_INDENT_PT = 32.0   # 首列文字左缩进（2 字符 @16pt）


def norm(font: str) -> str:
    return font.split("+")[-1].replace("-", "").replace(" ", "")


def split_cells(chars):
    """把一行字符切成"格"：遇到空白字符或间隙 > 5pt 就断开。

    注意一：空单元格在 PDF 里只剩一个空格字符，会被跳过而不产生"格"，
    所以末行只剩 1 个姓名时 cells 长度可能就是 1，行判定不能依赖长度。
    注意二：全角空格（U+3000）不是断格符——两字姓名按学院做法补了全角
    空格（林　霈），它属于姓名的一部分，必须留在格内。
    """
    cells, prev_end, brk = [], None, False
    for ch in chars:
        c = ch["c"]
        if (not c.strip()) and c != "\u3000":
            prev_end = ch["bbox"][2]
            brk = True
            continue
        x0, x1 = ch["bbox"][0], ch["bbox"][2]
        if cells and not brk and prev_end is not None and x0 - prev_end <= CELL_GAP_PT:
            px0, _, ptxt = cells[-1]
            cells[-1] = (px0, x1, ptxt + c)
        else:
            cells.append((x0, x1, c))
        prev_end = x1
        brk = False
    return cells


def merge_fragments(cells, expected, col_right):
    """把同一单元格被空白/间隙切碎的碎片粘回去。

    两字姓名按学院做法补了全角空格（林　霈），Word 导出 PDF 时它变成
    "8pt 空格字形 + 8pt 定位间隙"，按字符判断永远会把它切成两格。
    规则：碎片起点对不上任何列起点、但落在上一格所在列的范围之内，
    就属于上一格——真单元格的文字永远从自己的列起点开始。
    """
    out = []  # [x0, x1, text, col_idx]
    for x0, x1, txt in cells:
        k = min(range(len(expected)), key=lambda i: abs(x0 - expected[i]))
        aligned = abs(x0 - expected[k]) < COL_TOL_PT
        if aligned:
            out.append([x0, x1, txt, k])
        elif out and x0 < col_right[out[-1][3]] - 1.0:
            out[-1][1] = x1
            out[-1][2] += txt
        else:
            out.append([x0, x1, txt, k])
    return [(x0, x1, txt) for x0, x1, txt, _ in out]


def docx_table_grid(path):
    """从 docx 读表格几何：边距、各列宽、行数、每行单元格文字。"""
    from docx import Document
    d = Document(path)
    if not d.tables:
        return None
    t = d.tables[0]
    left = d.sections[0].left_margin.pt
    widths = [c.width.pt for c in t.columns]
    rows = []
    for row in t.rows:
        seen, cells = set(), []
        for c in row.cells:
            if c._tc in seen:
                continue
            seen.add(c._tc)
            cells.append(c.text.strip())
        rows.append(cells)
    return {"left": left, "widths": widths, "rows": rows}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("--docx")
    ap.add_argument("--title-font")
    ap.add_argument("--body-font")
    ap.add_argument("--title-pt", type=float, default=22.0)
    ap.add_argument("--body-pt", type=float, default=16.0)
    args = ap.parse_args()

    doc = pymupdf.open(args.pdf)
    print(f"pages: {len(doc)}")
    problems: list[str] = []

    lines_all = []       # [(page_no, y, font, size, [(x0, x1, text)])]
    for pno, page in enumerate(doc):
        print(f"page {pno+1}: {page.rect.width:.1f} x {page.rect.height:.1f} pt")
        buckets: dict[int, dict] = {}
        for blk in page.get_text("rawdict")["blocks"]:
            for ln in blk.get("lines", []):
                y = int(round(ln["bbox"][1]))
                key = min((k for k in buckets if abs(k - y) <= 2), default=y)
                b = buckets.setdefault(key, {"chars": [], "font": None, "size": None})
                for sp in ln["spans"]:
                    b["font"] = b["font"] or norm(sp["font"])
                    b["size"] = b["size"] or round(sp["size"], 1)
                    b["chars"].extend(sp["chars"])
        for y in sorted(buckets):
            cells = split_cells(buckets[y]["chars"])
            if cells:
                lines_all.append((pno + 1, y, buckets[y]["font"], buckets[y]["size"], cells))

    # ---- 字体 ----
    print("\n--- 字体清单 ---")
    seen = {}
    for _, _, f, s, cells in lines_all:
        seen[(f, s)] = seen.get((f, s), 0) + sum(len(c[2]) for c in cells)
    for (f, s), n in sorted(seen.items(), key=lambda kv: -kv[1]):
        print(f"  {f:<24} {s:5.1f}pt  {n:3d} 字")
    for s in {s for (_, s) in seen}:
        if s not in (args.title_pt, args.body_pt):
            problems.append(f"出现非规范字号 {s}pt")
    for key, want in (("title", args.title_font), ("body", args.body_font)):
        if want:
            target = PDF_NAME.get(want, want)
            if any(target.lower() in f.lower() for (f, _) in seen):
                print(f"  OK {key} 字体 = {want}({target})")
            else:
                problems.append(f"{key} 字体应为 {want}({target})，PDF 里没找到")

    # ---- 名单表 ----
    print("\n--- 名单表 ---")
    grid = docx_table_grid(args.docx) if args.docx else None
    if not grid:
        print("  未提供 --docx 或无表格，跳过")
    else:
        left = grid["left"]
        widths = grid["widths"]
        col_x, col_right = [left], []          # 列左边界（不含首列缩进）
        for w in widths:
            col_right.append(col_x[-1] + w)
            col_x.append(col_right[-1])
        expected = [x + (FIRST_COL_INDENT_PT if i == 0 else 0.0) for i, x in enumerate(col_x[:-1])]
        n_cols = len(widths)
        print("  期望列起点: " + ", ".join(f"{x:.1f}" for x in expected))

        def nearest(x):
            return min(range(n_cols), key=lambda i: abs(x - expected[i]))

        # 行判定：靠锚点定位，不靠列位。
        # 正文里数字与汉字之间会被抽出空格字符，断格后某一段起点可能碰巧落在第 2 列列位上，
        # 单看列位必然误判。表格一定在"名单如下："行之后、落款日期行之前。
        anchor_pg = anchor_y = None
        for (pg, y, _, _, cells) in lines_all:
            if "名单如下" in "".join(c[2] for c in cells):
                anchor_pg, anchor_y = pg, y
        end_pg, end_y = (lines_all[-2][0], lines_all[-2][1]) if len(lines_all) >= 2 else (None, None)

        def in_table_zone(pg, y):
            if anchor_pg is None:
                return False
            if not (anchor_pg <= pg <= end_pg):
                return False
            if pg == anchor_pg and y <= anchor_y + 4:
                return False
            if pg == end_pg and y >= end_y - 2:
                return False
            return True

        if anchor_pg is None:
            print("  未找到'名单如下：'锚点，跳过表格校验")
            table_lines = []
        else:
            table_lines = []
            for (pg, y, _, _, cells) in lines_all:
                if in_table_zone(pg, y):
                    table_lines.append((y, merge_fragments(cells, expected, col_right)))

        n_docx, n_pdf = len(grid["rows"]), len(table_lines)
        print(f"  docx 表格 {n_docx} 行 / PDF 命中 {n_pdf} 行")
        if table_lines:
            ys = [y for y, _ in table_lines]
            print("  行 y 坐标: " + ", ".join(f"{y:.1f}" for y in ys))
        if n_pdf > n_docx:
            problems.append(f"表格行数变多（docx {n_docx} 行，PDF {n_pdf} 行）"
                            f"——有姓名折行，模板要求一行放下、不许换行")
        elif n_pdf < n_docx:
            problems.append(f"表格行数变少（docx {n_docx} 行，PDF 仅 {n_pdf} 行）"
                            f"——有整行丢失，请人工核对")

        worst, worst_at = 0.0, ""
        over = []
        for y, cells in table_lines:
            for c in cells:
                k = nearest(c[0])
                dev = abs(c[0] - expected[k])
                if dev > worst:
                    worst, worst_at = dev, c[2]
                if c[1] > col_right[k] + 1.0:
                    over.append((c[2], k + 1, c[1], col_right[k]))
        print(f"  最大列偏差 {worst:.2f}pt" + (f"（'{worst_at}'）" if worst_at else "")
              + f" -> {'PASS (<1.5pt)' if worst < 1.5 else 'FAIL'}")
        if worst >= 1.5:
            problems.append(f"单元格没对齐列位置，最大偏差 {worst:.1f}pt（'{worst_at}'）")
        print(f"  越出所在列右界的单元格 {len(over)} 个")
        for txt, col, x1, lim in over:
            print(f"    '{txt}' 在第 {col} 列，文字右缘 {x1:.1f}pt > 列右界 {lim:.1f}pt")
        if over:
            problems.append("有单元格文字溢出到相邻列（该列宽度不够放这个名字）")

    # ---- 落款 ----
    if len(lines_all) >= 2:
        (_, _, _, _, c1), (_, _, _, _, c2) = lines_all[-2], lines_all[-1]
        r1 = max(c[1] for c in c1)
        r2 = max(c[1] for c in c2)
        print(f"\n--- 落款 ---\n  日期行右缘 {r1:.1f}pt / 单位行右缘 {r2:.1f}pt  差 {abs(r1-r2):.1f}pt")
        if abs(r1 - r2) > 1.5:
            problems.append(f"落款两行右边界不齐 {abs(r1-r2):.1f}pt")

    print("\n--- 结论 ---")
    if problems:
        for p in problems:
            print("  FAIL:", p)
        sys.exit(1)
    print("  PASS")


if __name__ == "__main__":
    main()
