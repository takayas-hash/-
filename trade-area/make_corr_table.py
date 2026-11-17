"""値上げ前後の分析結果（analyze_price.py / analyze_dow.py の出力）から、相関の一覧表 .xlsx を作る。

使い方: python make_corr_table.py output/相関一覧.xlsx
"""
import csv
import sys

import numpy as np
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from scipy import stats

from analyze_price import VARS

FONT = "Arial"
HEAD = PatternFill("solid", fgColor="1F3864")

SETS = [
    ("曜日差抜き前年比（4〜6月 → 8〜10月）", "output/price_analysis_dow.csv", [
        ("売上前年比の変化（後−前）", "売上前年比の変化pt"),
        ("客数前年比の変化（後−前）", "客数前年比の変化pt"),
        ("値上げ後の売上前年比", "値上げ後_売上前年比"),
        ("値上げ後の客数前年比", "値上げ後_客数前年比"),
    ]),
    ("前年比（曜日差込み、4〜6月 → 8〜10月）と売上の大きさ", "output/price_analysis.csv", [
        ("売上前年比の変化（後−前）", "売上前年比の変化pt"),
        ("客数前年比の変化（後−前）", "客数前年比の変化pt"),
        ("値上げ後の月平均売上", "値上げ後_月平均売上"),
        ("値上げ前の月平均売上", "値上げ前_月平均売上"),
    ]),
]


def load(path):
    rows = list(csv.DictReader(open(path, encoding="utf-8-sig")))
    for r in rows:
        for k in list(r):
            try:
                r[k] = float(r[k])
            except (TypeError, ValueError):
                r[k] = None if r[k] == "" else r[k]
    return rows


def judge(r, p):
    a = abs(r)
    if p >= 0.05 or a < 0.2:
        return "ほぼ無し"
    s = "強い" if a >= 0.5 else "中くらいの" if a >= 0.3 else "弱い"
    return f"{s}{'正' if r > 0 else '負'}の関係"


def main(dst):
    wb = Workbook()
    ws = wb.active
    ws.title = "相関一覧"
    ws.append(["相関係数 r：0なら無関係、±1に近いほど強い関係。目安は|r|0.3以上で関係あり、0.1前後はほぼ無し。"
               "p値が0.05未満なら偶然とは言いにくい。色：青＝正の関係、赤＝負の関係。"])
    ws.append([])
    for title, path, ys in SETS:
        rows = load(path)
        ws.append([f"{title}（対象 {len(rows)}店）"])
        ws.cell(ws.max_row, 1).font = Font(name=FONT, bold=True, size=11)
        head = ["商圏・地域の指標"]
        for yl, _ in ys:
            head += [f"{yl} r", "p値", "判定"]
        ws.append(head)
        hr = ws.max_row
        for c in ws[hr]:
            c.fill = HEAD
            c.font = Font(name=FONT, bold=True, color="FFFFFF", size=9)
            c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        start = hr + 1
        for k, label in VARS:
            line = [label]
            for _, yk in ys:
                pairs = [(r[k], r[yk]) for r in rows if isinstance(r.get(k), float) and isinstance(r.get(yk), float)]
                x, y = np.array(pairs).T
                rr, pp = stats.pearsonr(x, y)
                line += [round(float(rr), 2), round(float(pp), 3), judge(rr, pp)]
            ws.append(line)
        end = ws.max_row
        for j in range(len(ys)):
            col = get_column_letter(2 + j * 3)
            ws.conditional_formatting.add(f"{col}{start}:{col}{end}", ColorScaleRule(
                start_type="num", start_value=-0.6, start_color="F8696B",
                mid_type="num", mid_value=0, mid_color="FFFFFF",
                end_type="num", end_value=0.6, end_color="5A8AC6"))
        ws.append([])

    # 5分位
    ws2 = wb.create_sheet("グループ別の平均")
    rows = load("output/price_analysis_dow.csv")
    ws2.append(["曜日差抜き前年比：指標の小さい順に5グループに分けた平均（4〜6月 → 8〜10月）"])
    ws2.cell(1, 1).font = Font(name=FONT, bold=True, size=11)
    for k, label in VARS:
        ws2.append([])
        ws2.append([label, "下限", "上限", "店数", "値上げ前 売上前年比", "値上げ後 売上前年比", "売上の変化pt",
                    "値上げ前 客数前年比", "値上げ後 客数前年比", "客数の変化pt"])
        for c in ws2[ws2.max_row]:
            c.fill = HEAD
            c.font = Font(name=FONT, bold=True, color="FFFFFF", size=9)
            c.alignment = Alignment(wrap_text=True, horizontal="center")
        v = sorted([r for r in rows if isinstance(r.get(k), float)], key=lambda r: r[k])
        n = len(v)
        for i in range(5):
            g = v[i * n // 5:(i + 1) * n // 5]
            m = lambda c: float(np.mean([r[c] for r in g]))  # noqa: E731
            ws2.append([f"第{i + 1}グループ（{'小' if i == 0 else '大' if i == 4 else '中'}）", g[0][k], g[-1][k], len(g),
                        m("値上げ前_売上前年比"), m("値上げ後_売上前年比"), round(m("売上前年比の変化pt"), 1),
                        m("値上げ前_客数前年比"), m("値上げ後_客数前年比"), round(m("客数前年比の変化pt"), 1)])
            for c in (5, 6, 8, 9):
                ws2.cell(ws2.max_row, c).number_format = "0.0%"
            for c in (2, 3):
                ws2.cell(ws2.max_row, c).number_format = "#,##0.0" if g[-1][k] < 100 else "#,##0"

    for w in (ws, ws2):
        w.column_dimensions["A"].width = 26
        for i in range(2, 16):
            w.column_dimensions[get_column_letter(i)].width = 13
        for row in w.iter_rows():
            for c in row:
                if c.font.name != FONT:
                    c.font = Font(name=FONT, size=10, bold=c.font.bold, color=c.font.color)
    wb.save(dst)


if __name__ == "__main__":
    main(sys.argv[1])
