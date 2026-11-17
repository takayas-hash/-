"""5km商圏の人口・世帯数と月商の関係をまとめ、人口から見た「期待月商」と実績の差を店ごとに出す。

入力: data/sales/YYYYMM.csv（社内データ。コミットしない）、output/stores_trade_area.csv
使い方:
  python analyze_expected.py --months 202604,202605,202606,202608,202609,202610 --out output/期待月商.xlsx
期待月商 = a + b × ln(5km人口)。人口が増えるほど伸びが緩やかになる形（対数）で、全店の実績に当てはめる。
xlsx では a・b と期待月商・差を数式で入れるので、店を外したり月を変えたりしても再計算される。
"""
import argparse
import csv
import datetime as dt
import math
import os

import numpy as np
from openpyxl import Workbook
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

FONT = "Arial"
HEAD = PatternFill("solid", fgColor="1F3864")
INPUT = PatternFill("solid", fgColor="FFF2CC")
BANDS = [(0, 75_000), (75_000, 100_000), (100_000, 150_000), (150_000, 200_000), (200_000, 300_000),
         (300_000, 500_000), (500_000, 1_000_000), (1_000_000, 10_000_000)]


def band_label(lo, hi):
    f = lambda v: f"{v / 10000:g}万"  # noqa: E731
    return f"{f(lo)}〜{f(hi)}未満" if lo and hi < 10_000_000 else (f"{f(hi)}未満" if not lo else f"{f(lo)}以上")


def load(months, open_before):
    area = {r["店舗コード"]: r for r in csv.DictReader(open("output/stores_trade_area.csv", encoding="utf-8-sig"))}
    sales = {}
    for m in months:
        for r in csv.DictReader(open(f"data/sales/{m}.csv", encoding="utf-8")):
            sales.setdefault(str(int(r["店舗コード"])), {})[m] = r
    rows, excluded = [], []
    for code, ms in sales.items():
        t = area.get(code)
        name = next(iter(ms.values()))["店舗名称"].strip()
        reason = None
        if not t:
            reason = "商圏データなし"
        elif not t.get("商圏5km_人口"):
            reason = "緯度経度なし（商圏人口なし）"
        elif t["開店状態"] != "営業中":
            reason = f"開店状態: {t['開店状態']}"
        else:
            try:
                od = dt.date(*map(int, t["オープン日"].split("/")))
                if od > open_before:
                    reason = f"開店 {t['オープン日']}（開店直後の売上が高い時期を含む）"
            except ValueError:
                reason = "オープン日不明"
        if not reason and any(m not in ms for m in months):
            reason = "月のデータが欠けている"
        if not reason and any(ms[m]["稼動日数"] != ms[m]["暦日付"] or float(ms[m]["売上実績"] or 0) <= 0 for m in months):
            reason = "休業日のある月がある"
        if reason:
            excluded.append((code, name, reason))
            continue
        rows.append(dict(
            店舗コード=code, 店舗名=name, 営業部=next(iter(ms.values()))["営業部名称"].strip(), 直営FC=t["直営FC"],
            ブロック=t.get("ブロック") or "（未設定）",
            市区町村=t["市区町村名"], 席数=float(t["席数"] or 0) or None,
            人口5km=float(t["商圏5km_人口"]), 世帯5km=float(t["商圏5km_世帯数"]), 人口3km=float(t["商圏3km_人口"]),
            年収5km=float(t["商圏5km_推計平均世帯年収_万円"] or 0) or None, 高齢5km=float(t["商圏5km_65歳以上比率%"]),
            月商=sum(float(ms[m]["売上実績"]) for m in months) / len(months),
        ))
    return rows, excluded


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--months", required=True)
    ap.add_argument("--open-before", default="2025-03-31")
    ap.add_argument("--out", default="output/期待月商.xlsx")
    a = ap.parse_args()
    months = a.months.split(",")
    rows, excluded = load(months, dt.date.fromisoformat(a.open_before))

    x = np.log([r["人口5km"] for r in rows])
    y = np.array([r["月商"] for r in rows])
    b, a0 = np.polyfit(x, y, 1)
    pred = a0 + b * x
    r2 = 1 - ((y - pred) ** 2).sum() / ((y - y.mean()) ** 2).sum()
    for r, p in zip(rows, pred):
        r["期待月商"] = p
        r["差"] = r["月商"] - p
        r["達成率"] = r["月商"] / p
    print(f"対象 {len(rows)}店 / 除外 {len(excluded)}店  期待月商 = {a0:,.0f} + {b:,.0f} × ln(5km人口)  R²={r2:.3f}")
    print(f"  例: 10万人 {a0 + b * math.log(1e5):,.0f}円 / 30万人 {a0 + b * math.log(3e5):,.0f}円 / 100万人 {a0 + b * math.log(1e6):,.0f}円")
    for lo, hi in BANDS:
        g = [r for r in rows if lo <= r["人口5km"] < hi]
        if g:
            print(f"  {band_label(lo, hi):<12} {len(g):>3}店 平均人口{np.mean([r['人口5km'] for r in g]):>10,.0f} "
                  f"平均世帯{np.mean([r['世帯5km'] for r in g]):>9,.0f} 平均月商{np.mean([r['月商'] for r in g]) / 1e4:>6,.0f}万円 "
                  f"中央値{np.median([r['月商'] for r in g]) / 1e4:>6,.0f}万円 最小{min(r['月商'] for r in g) / 1e4:>5,.0f} 最大{max(r['月商'] for r in g) / 1e4:>5,.0f}")
    rs = sorted(rows, key=lambda r: r["達成率"])
    print("下位:", [(r["店舗名"], round(r["達成率"] * 100)) for r in rs[:10]])
    print("上位:", [(r["店舗名"], round(r["達成率"] * 100)) for r in rs[-10:][::-1]])
    write_xlsx(a.out, rows, excluded, months)


def style_head(ws, row):
    for c in ws[row]:
        if c.value is not None:
            c.fill = HEAD
            c.font = Font(name=FONT, bold=True, color="FFFFFF", size=9)
            c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")


def write_xlsx(path, rows, excluded, months):
    wb = Workbook()
    # ---- 店舗別（元データ＋数式）。列: A コード B 店舗名 C ブロック D AMエリア E 直営FC F 市区町村 G 席数
    # H 5km人口 I 5km世帯 J 3km人口 K 年収 L 高齢 M 月商 N 人口帯 O ln P 期待 Q 差 R 達成率 S 全店順位 T ブロック内 U AM内
    ws = wb.create_sheet("店舗別の期待月商")
    cols = ["店舗コード", "店舗名", "ブロック", "AMエリア", "直営FC", "市区町村", "席数", "5km人口", "5km世帯数", "3km人口",
            "5km推計世帯年収(万円)", "5km 65歳以上比率%", "実績月商(円)", "人口帯", "ln(5km人口)", "期待月商(円)",
            "差(実績−期待, 円)", "達成率(実績÷期待)", "達成率の順位(全店)", "ブロック内の順位", "AMエリア内の順位"]
    ws.append(cols)
    style_head(ws, 1)
    n = len(rows)
    last = n + 1
    R = lambda c: f"${c}$2:${c}${last}"  # noqa: E731
    rows_sorted = sorted(rows, key=lambda r: (r["ブロック"], r["営業部"], -r["達成率"]))
    for i, r in enumerate(rows_sorted, start=2):
        ws.append([int(r["店舗コード"]), r["店舗名"], r["ブロック"], r["営業部"], r["直営FC"], r["市区町村"], r["席数"],
                   r["人口5km"], r["世帯5km"], r["人口3km"], r["年収5km"], r["高齢5km"], round(r["月商"]),
                   f'=IF(H{i}<75000,"7.5万未満",IF(H{i}<100000,"7.5万〜10万未満",IF(H{i}<150000,"10万〜15万未満",IF(H{i}<200000,"15万〜20万未満",'
                   f'IF(H{i}<300000,"20万〜30万未満",IF(H{i}<500000,"30万〜50万未満",IF(H{i}<1000000,"50万〜100万未満","100万以上")))))))',
                   f"=LN(H{i})", f"=係数!$B$3+係数!$B$4*O{i}", f"=M{i}-P{i}", f"=M{i}/P{i}",
                   f"=RANK(R{i},{R('R')},0)",
                   f'=COUNTIFS({R("C")},C{i},{R("R")},">"&R{i})+1',
                   f'=COUNTIFS({R("D")},D{i},{R("R")},">"&R{i})+1'])
        for c, fmt in ((8, "#,##0"), (9, "#,##0"), (10, "#,##0"), (13, "#,##0"), (15, "0.000"), (16, "#,##0"),
                       (17, "#,##0;[Red]-#,##0"), (18, "0.0%")):
            ws.cell(i, c).number_format = fmt
    ws.conditional_formatting.add(f"R2:R{last}", ColorScaleRule(
        start_type="num", start_value=0.7, start_color="F8696B", mid_type="num", mid_value=1, mid_color="FFFFFF",
        end_type="num", end_value=1.3, end_color="63BE7B"))
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:U{last}"
    widths = [9, 26, 8, 14, 7, 22, 7, 11, 11, 11, 11, 10, 13, 15, 9, 13, 14, 11, 9, 9, 9]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 42
    S = "店舗別の期待月商!"

    # ---- 係数
    wc = wb.create_sheet("係数", 0)
    rng_x, rng_y = f"{S}{R('O')}", f"{S}{R('M')}"
    wc.append(["期待月商の計算式：期待月商 ＝ a ＋ b × ln(5km人口)（全店の実績に当てはめた線）"])
    wc.append([])
    wc.append(["a（切片）", f"=INTERCEPT({rng_y},{rng_x})"])
    wc.append(["b（人口の効き方）", f"=SLOPE({rng_y},{rng_x})"])
    wc.append(["当てはまり R²（0〜1。人口で説明できる割合）", f"=RSQ({rng_y},{rng_x})"])
    wc.append(["対象店数", f"=COUNT({rng_y})"])
    wc.append([])
    wc.append(["5km人口の例", "期待月商(円)"])
    style_head(wc, 8)
    for i, p in enumerate([50_000, 100_000, 150_000, 200_000, 300_000, 500_000, 1_000_000], start=9):
        wc.append([p, f"=$B$3+$B$4*LN(A{i})"])
        wc.cell(i, 1).number_format = "#,##0"
        wc.cell(i, 2).number_format = "#,##0"
    for c in ("B3", "B4"):
        wc[c].number_format = "#,##0"
    wc["B5"].number_format = "0.000"
    wc.append([])
    wc.append([f"実績月商：{', '.join(m[:4] + '年' + m[4:].lstrip('0') + '月' for m in months)} の平均（7月はコース値上げ月のため除外）。"
               "全月とも休業日のない、2025年3月末までに開店した営業中の店が対象。"])
    wc.append(["ln は自然対数。人口が2倍になると期待月商は b×0.69 だけ増える（人口が増えるほど伸びは緩やか）。"])
    wc.append(["絞り込み：「店舗別の期待月商」シートの見出しのフィルタで、ブロック・AMエリアを選ぶ。達成率の列で並べ替えると、その中で期待との差が大きい順に並ぶ。"])
    wc.column_dimensions["A"].width = 44
    wc.column_dimensions["B"].width = 18
    for row in wc.iter_rows():
        for c in row:
            if c.fill != HEAD:
                c.font = Font(name=FONT, size=10, bold=c.row == 1)

    # ---- 人口帯別の平均
    wb_ = wb.create_sheet("人口帯別の平均", 1)
    wb_.append(["5km人口の帯", "店数", "平均 5km人口", "平均 5km世帯数", "平均 月商(円)", "月商 中央値(円)", "月商 最小(円)",
                "月商 最大(円)", "平均 席数", "1席当たり月商(円)", "人口1人当たり月商(円)"])
    style_head(wb_, 1)
    labels = ["7.5万未満", "7.5万〜10万未満", "10万〜15万未満", "15万〜20万未満", "20万〜30万未満", "30万〜50万未満", "50万〜100万未満", "100万以上"]
    for i, lab in enumerate(labels, start=2):
        crit = f"{S}{R('N')},$A{i}"
        wb_.append([lab, f"=COUNTIF({crit})", f"=AVERAGEIF({crit},{S}{R('H')})",
                    f"=AVERAGEIF({crit},{S}{R('I')})", f"=AVERAGEIF({crit},{S}{R('M')})",
                    None,
                    f"=_xlfn.MINIFS({S}{R('M')},{crit})", f"=_xlfn.MAXIFS({S}{R('M')},{crit})",
                    f"=AVERAGEIF({crit},{S}{R('G')})", f"=E{i}/I{i}", f"=E{i}/C{i}"])
        wb_.cell(i, 6).value = ArrayFormula(f"F{i}", f"=MEDIAN(IF({S}{R('N')}=$A{i},{S}{R('M')}))")
    t = len(labels) + 2
    wb_.append(["全店", f"=SUM(B2:B{t - 1})", f"=AVERAGE({S}{R('H')})", f"=AVERAGE({S}{R('I')})",
                f"=AVERAGE({S}{R('M')})", f"=MEDIAN({S}{R('M')})", f"=MIN({S}{R('M')})", f"=MAX({S}{R('M')})",
                f"=AVERAGE({S}{R('G')})", f"=E{t}/I{t}", f"=E{t}/C{t}"])
    for r_ in range(2, t + 1):
        for c in range(3, 12):
            wb_.cell(r_, c).number_format = "#,##0.0" if c == 9 else "#,##0"
    wb_.cell(t, 1).font = Font(name=FONT, bold=True)
    for i, w in enumerate([16, 7, 13, 13, 14, 14, 13, 13, 9, 13, 13], 1):
        wb_.column_dimensions[get_column_letter(i)].width = w
    wb_.row_dimensions[1].height = 32

    # ---- ブロック・AM別のまとめ
    wg = wb.create_sheet("ブロック・AM別", 2)
    ghead = ["区分", "名前", "店数", "平均 5km人口", "平均 月商(円)", "平均 期待月商(円)", "平均 達成率",
             "最高の達成率", "最低の達成率", "達成率100%未満の店数", "達成率が最も高い店", "達成率が最も低い店"]
    wg.append(ghead)
    style_head(wg, 1)
    groups = [("ブロック", "C", sorted({r["ブロック"] for r in rows})), ("AMエリア", "D", sorted({r["営業部"] for r in rows}))]
    for kind, col, names in groups:
        for name in names:
            i = wg.max_row + 1
            crit = f"{S}{R(col)},$B{i}"
            wg.append([kind, name, f"=COUNTIF({crit})", f"=AVERAGEIF({crit},{S}{R('H')})",
                       f"=AVERAGEIF({crit},{S}{R('M')})", f"=AVERAGEIF({crit},{S}{R('P')})",
                       f"=AVERAGEIF({crit},{S}{R('R')})", f"=_xlfn.MAXIFS({S}{R('R')},{crit})",
                       f"=_xlfn.MINIFS({S}{R('R')},{crit})", f'=COUNTIFS({crit},{S}{R("R")},"<1")', None, None])
            wg.cell(i, 11).value = ArrayFormula(
                f"K{i}", f"=INDEX({S}{R('B')},MATCH(1,({S}{R(col)}=$B{i})*({S}{R('R')}=H{i}),0))")
            wg.cell(i, 12).value = ArrayFormula(
                f"L{i}", f"=INDEX({S}{R('B')},MATCH(1,({S}{R(col)}=$B{i})*({S}{R('R')}=I{i}),0))")
            for c, fmt in ((4, "#,##0"), (5, "#,##0"), (6, "#,##0"), (7, "0.0%"), (8, "0.0%"), (9, "0.0%")):
                wg.cell(i, c).number_format = fmt
    gl = wg.max_row
    wg.conditional_formatting.add(f"G2:G{gl}", ColorScaleRule(
        start_type="num", start_value=0.85, start_color="F8696B", mid_type="num", mid_value=1, mid_color="FFFFFF",
        end_type="num", end_value=1.15, end_color="63BE7B"))
    wg.auto_filter.ref = f"A1:L{gl}"
    wg.freeze_panes = "C2"
    for i, w in enumerate([10, 16, 6, 12, 13, 13, 9, 9, 9, 10, 24, 24], 1):
        wg.column_dimensions[get_column_letter(i)].width = w
    wg.row_dimensions[1].height = 32

    # ---- 上位・下位（全店）
    wr = wb.create_sheet("差の大きい店", 3)
    rs = sorted(rows, key=lambda r: r["達成率"])
    head = ["区分", "順位", "店舗コード", "店舗名", "ブロック", "AMエリア", "5km人口", "実績月商(円)", "期待月商(円)", "差(円)", "達成率", "席数"]
    wr.append(head)
    style_head(wr, 1)
    for title, items in (("期待より大きい（上位20）", rs[::-1][:20]), ("期待より小さい（下位20）", rs[:20])):
        for k, r in enumerate(items, 1):
            j = wr.max_row + 1
            row_ref = f"MATCH($C{j},{S}{R('A')},0)"
            wr.append([title, k, int(r["店舗コード"]), r["店舗名"], r["ブロック"], r["営業部"], r["人口5km"],
                       f"=INDEX({S}{R('M')},{row_ref})", f"=INDEX({S}{R('P')},{row_ref})",
                       f"=H{j}-I{j}", f"=H{j}/I{j}", r["席数"]])
            for c, fmt in ((7, "#,##0"), (8, "#,##0"), (9, "#,##0"), (10, "#,##0;[Red]-#,##0"), (11, "0.0%")):
                wr.cell(j, c).number_format = fmt
    wr.auto_filter.ref = f"A1:L{wr.max_row}"
    wr.freeze_panes = "E2"
    for i, w in enumerate([20, 6, 10, 26, 8, 14, 11, 13, 13, 13, 9, 7], 1):
        wr.column_dimensions[get_column_letter(i)].width = w

    # ---- 除外
    we = wb.create_sheet("対象外の店")
    we.append(["店舗コード", "店舗名", "理由"])
    style_head(we, 1)
    for e in sorted(excluded):
        we.append(list(e))
    for i, w in enumerate([10, 28, 50], 1):
        we.column_dimensions[get_column_letter(i)].width = w

    del wb["Sheet"]
    for w in wb.worksheets:
        for row in w.iter_rows():
            for c in row:
                if c.font.name != FONT:
                    c.font = Font(name=FONT, size=10, bold=c.font.bold, color=c.font.color)
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    wb.save(path)


if __name__ == "__main__":
    main()
