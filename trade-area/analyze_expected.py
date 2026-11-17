"""商圏（半径5kmまたは10km）の人口・世帯数と月商の関係をまとめ、人口から見た「期待月商」と実績の差を店ごとに出す。

入力: data/sales/YYYYMM.csv（社内データ。コミットしない）、output/stores_trade_area.csv
使い方:
  python analyze_expected.py --months 202604,202605,202606,202608,202609,202610 --km 10 --out output/期待月商_10km.xlsx
前提: python cannibal.py output/stores_trade_area.csv output/cannibal_10km.csv 10
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
KM = 5  # 商圏の半径。--km で変える
BANDS_BY_KM = {
    5: [0, 75_000, 100_000, 150_000, 200_000, 300_000, 500_000, 1_000_000],
    10: [0, 200_000, 400_000, 600_000, 1_000_000, 1_500_000, 2_500_000],
}
EFF_BANDS_BY_KM = {  # 実質人口（自社店で分け合った後）の帯
    5: [0, 75_000, 100_000, 150_000, 200_000, 300_000],
    10: [0, 150_000, 200_000, 250_000, 300_000, 400_000],
}
EFF_BANDS = []
BANDS = [(0, 75_000), (75_000, 100_000), (100_000, 150_000), (150_000, 200_000), (200_000, 300_000),
         (300_000, 500_000), (500_000, 1_000_000), (1_000_000, 10_000_000)]


def band_label(lo, hi):
    f = lambda v: f"{v / 10000:g}万"  # noqa: E731
    return f"{f(lo)}〜{f(hi)}未満" if lo and hi < 10_000_000 else (f"{f(hi)}未満" if not lo else f"{f(lo)}以上")


def band_formula(i, col="H", bands=None):
    """人口帯の名前を返す数式（H列＝商圏人口、Y列＝実質商圏人口）。"""
    bands = bands or BANDS
    f = ""
    for lo, hi in bands[:-1]:
        f += f'IF({col}{i}<{hi},"{band_label(lo, hi)}",'
    return "=" + f + f'"{band_label(*bands[-1])}"' + ")" * (len(bands) - 1)


def parse_hours(txt):
    """「17:00 - 23:00」→ (17.0, 23.0)。読めなければ None。"""
    try:
        a_, b_ = [x.strip() for x in txt.split("-")]
        return int(a_[:2]) + int(a_[3:]) / 60, int(b_[:2]) + int(b_[3:]) / 60
    except (ValueError, AttributeError):
        return None


def hours_info(wd_txt, we_txt):
    wd, we = parse_hours(wd_txt), parse_hours(we_txt)
    if not wd:
        return dict(区分="不明", 昼=0, 夕16=0, 深夜=0)
    o, c = wd
    if c == 15:  # 「11:00 - 15:00」は昼の時間帯だけの表示。夜も営業とみなす
        c = 23
    late = max(c, we[1] if we else 0) > 23
    cat = "平日 昼から" if o < 15 else ("平日 16時から" if o < 17 else "平日 17時から")
    return dict(区分=cat, 昼=int(cat == "平日 昼から"), 夕16=int(cat == "平日 16時から"), 深夜=int(late))


def load(months, open_before):
    area = {r["店舗コード"]: r for r in csv.DictReader(open("output/stores_trade_area.csv", encoding="utf-8-sig"))}
    sales = {}
    for m in months:
        for r in csv.DictReader(open(f"data/sales/{m}.csv", encoding="utf-8")):
            sales.setdefault(str(int(r["店舗コード"])), {})[m] = r
    hours = {}
    if os.path.exists("data/hours.csv"):
        hours = {r["店舗コード"]: r for r in csv.DictReader(open("data/hours.csv", encoding="utf-8"))}
    rows, excluded = [], []
    for code, ms in sales.items():
        t = area.get(code)
        name = next(iter(ms.values()))["店舗名称"].strip()
        reason = None
        if not t:
            reason = "商圏データなし"
        elif not t.get(f"商圏{KM}km_人口"):
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
            平日営業時間=hours.get(code, {}).get("平日営業時間", ""), 土日営業時間=hours.get(code, {}).get("土日営業時間", ""),
            **hours_info(hours.get(code, {}).get("平日営業時間", ""), hours.get(code, {}).get("土日営業時間", "")),
            市区町村=t["市区町村名"], 席数=float(t["席数"] or 0) or None,
            人口5km=float(t[f"商圏{KM}km_人口"]), 世帯5km=float(t[f"商圏{KM}km_世帯数"]), 人口3km=float(t["商圏3km_人口"]),
            年収5km=float(t[f"商圏{KM}km_推計平均世帯年収_万円"] or 0) or None, 高齢5km=float(t[f"商圏{KM}km_65歳以上比率%"]),
            月商=sum(float(ms[m]["売上実績"]) for m in months) / len(months),
        ))
    return rows, excluded


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--months", required=True)
    ap.add_argument("--open-before", default="2025-03-31")
    ap.add_argument("--out", default="output/期待月商.xlsx")
    ap.add_argument("--km", type=int, default=5, choices=sorted(BANDS_BY_KM))
    a = ap.parse_args()
    global KM, BANDS, EFF_BANDS
    KM = a.km
    e = BANDS_BY_KM[KM]
    BANDS = list(zip(e, e[1:] + [100_000_000]))
    e = EFF_BANDS_BY_KM[KM]
    EFF_BANDS = list(zip(e, e[1:] + [100_000_000]))
    months = a.months.split(",")
    rows, excluded = load(months, dt.date.fromisoformat(a.open_before))
    can = {r["店舗コード"]: r for r in csv.DictReader(open(f"output/cannibal_{KM}km.csv", encoding="utf-8-sig"))}
    for r in rows:
        c = can[r["店舗コード"]]
        r["最寄り距離"] = float(c["最寄りのきんぐまでの距離km"])
        r["他店数5km"] = int(c[f"{KM}km以内の他のきんぐ数"])
        r["重なり率"] = float(c["商圏で他店と重なる人口の割合%"])
        r["実質人口"] = float(c["実質商圏人口(自社店で分け合った後)"])
    xe = np.log([r["実質人口"] for r in rows])
    b2, a2 = np.polyfit(xe, np.array([r["月商"] for r in rows]), 1)
    for r, xv in zip(rows, xe):
        r["期待月商2"] = a2 + b2 * xv
        r["達成率2"] = r["月商"] / r["期待月商2"]
    rs2 = sorted(rows, key=lambda r: r["達成率2"])
    print("カニバリ考慮 下位:", [(r["店舗名"], round(r["達成率2"] * 100)) for r in rs2[:10]])
    print("カニバリ考慮 上位:", [(r["店舗名"], round(r["達成率2"] * 100)) for r in rs2[-10:][::-1]])
    # 席数込み: 月商 = a + b × ln(実質人口) + c × 席数
    X = np.column_stack([np.ones(len(rows)), xe, [r["席数"] for r in rows]])
    yv = np.array([r["月商"] for r in rows])
    coef, *_ = np.linalg.lstsq(X, yv, rcond=None)
    p3 = X @ coef
    print(f"席数込み R²={1 - ((yv - p3) ** 2).sum() / ((yv - yv.mean()) ** 2).sum():.3f} coef={coef.round(0)}")
    for r, v in zip(rows, p3):
        r["期待月商3"] = v
        r["達成率3"] = r["月商"] / v
    rs3 = sorted(rows, key=lambda r: r["達成率3"])
    print("席数込み 下位:", [(r["店舗名"], round(r["達成率3"] * 100)) for r in rs3[:12]])
    print("席数込み 上位:", [(r["店舗名"], round(r["達成率3"] * 100)) for r in rs3[-12:][::-1]])
    # 営業時間込み: ＋ d×平日昼から ＋ e×平日16時から ＋ f×24時以降まで
    X4 = np.column_stack([np.ones(len(rows)), xe, [r["席数"] for r in rows], [r["昼"] for r in rows],
                          [r["夕16"] for r in rows], [r["深夜"] for r in rows]])
    c4, *_ = np.linalg.lstsq(X4, yv, rcond=None)
    p4 = X4 @ c4
    print(f"営業時間込み R²={1 - ((yv - p4) ** 2).sum() / ((yv - yv.mean()) ** 2).sum():.3f} coef={c4.round(0)}")
    for r, v in zip(rows, p4):
        r["期待月商4"] = v
        r["達成率4"] = r["月商"] / v
    rs4 = sorted(rows, key=lambda r: r["達成率4"])
    print("営業時間込み 下位:", [(r["店舗名"], round(r["達成率4"] * 100), r["区分"]) for r in rs4[:15]])
    print("営業時間込み 上位:", [(r["店舗名"], round(r["達成率4"] * 100), r["区分"]) for r in rs4[-15:][::-1]])

    x = np.log([r["人口5km"] for r in rows])
    y = np.array([r["月商"] for r in rows])
    b, a0 = np.polyfit(x, y, 1)
    pred = a0 + b * x
    r2 = 1 - ((y - pred) ** 2).sum() / ((y - y.mean()) ** 2).sum()
    for r, p in zip(rows, pred):
        r["期待月商"] = p
        r["差"] = r["月商"] - p
        r["達成率"] = r["月商"] / p
    print(f"対象 {len(rows)}店 / 除外 {len(excluded)}店  期待月商 = {a0:,.0f} + {b:,.0f} × ln({KM}km人口)  R²={r2:.3f}")
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
    cols = ["店舗コード", "店舗名", "ブロック", "AMエリア", "直営FC", "市区町村", "席数", f"{KM}km人口", f"{KM}km世帯数", "3km人口",
            f"{KM}km推計世帯年収(万円)", f"{KM}km 65歳以上比率%", "実績月商(円)", "人口帯", f"ln({KM}km人口)", "期待月商(円)",
            "差(実績−期待, 円)", "達成率(実績÷期待)", "達成率の順位(全店)", "ブロック内の順位", "AMエリア内の順位",
            "最寄りのきんぐまでの距離km", f"{KM}km以内の他のきんぐ数", f"{KM}km商圏で他店と重なる人口の割合%",
            f"実質{KM}km人口(自社店で分け合った後)", f"ln(実質{KM}km人口)", "期待月商・カニバリ考慮(円)", "差・カニバリ考慮(円)",
            "達成率・カニバリ考慮", "同 順位(全店)", "同 ブロック内の順位", "同 AMエリア内の順位", f"実質{KM}km人口の帯",
            f"ln(実質{KM}km人口)(計算用)", "席数(計算用)", "期待月商・カニバリ＋席数(円)", "差・カニバリ＋席数(円)",
            "達成率・カニバリ＋席数", "同 順位(全店)", "同 ブロック内の順位", "同 AMエリア内の順位",
            "平日営業時間", "土日営業時間", "平日の開店", "24時以降まで営業",
            f"ln(実質{KM}km人口)(計算用2)", "席数(計算用2)", "平日昼から(計算用)", "平日16時から(計算用)", "24時以降(計算用)",
            "期待月商・カニバリ＋席数＋営業時間(円)", "差・同(円)", "達成率・カニバリ＋席数＋営業時間",
            "同 順位(全店)", "同 ブロック内の順位", "同 AMエリア内の順位"]
    ws.append(cols)
    style_head(ws, 1)
    n = len(rows)
    last = n + 1
    R = lambda c: f"${c}$2:${c}${last}"  # noqa: E731
    rows_sorted = sorted(rows, key=lambda r: (r["ブロック"], r["営業部"], -r["達成率"]))
    for i, r in enumerate(rows_sorted, start=2):
        ws.append([int(r["店舗コード"]), r["店舗名"], r["ブロック"], r["営業部"], r["直営FC"], r["市区町村"], r["席数"],
                   r["人口5km"], r["世帯5km"], r["人口3km"], r["年収5km"], r["高齢5km"], round(r["月商"]),
                   band_formula(i),
                   f"=LN(H{i})", f"=係数!$B$3+係数!$B$4*O{i}", f"=M{i}-P{i}", f"=M{i}/P{i}",
                   f"=RANK(R{i},{R('R')},0)",
                   f'=COUNTIFS({R("C")},C{i},{R("R")},">"&R{i})+1',
                   f'=COUNTIFS({R("D")},D{i},{R("R")},">"&R{i})+1',
                   r["最寄り距離"], r["他店数5km"], r["重なり率"], r["実質人口"], f"=LN(Y{i})",
                   f"=係数!$E$3+係数!$E$4*Z{i}", f"=M{i}-AA{i}", f"=M{i}/AA{i}", f"=RANK(AC{i},{R('AC')},0)",
                   f'=COUNTIFS({R("C")},C{i},{R("AC")},">"&AC{i})+1',
                   f'=COUNTIFS({R("D")},D{i},{R("AC")},">"&AC{i})+1',
                   band_formula(i, "Y", EFF_BANDS),
                   f"=Z{i}", f"=G{i}", f"=係数!$H$3+係数!$H$4*AH{i}+係数!$H$5*AI{i}", f"=M{i}-AJ{i}", f"=M{i}/AJ{i}",
                   f"=RANK(AL{i},{R('AL')},0)",
                   f'=COUNTIFS({R("C")},C{i},{R("AL")},">"&AL{i})+1',
                   f'=COUNTIFS({R("D")},D{i},{R("AL")},">"&AL{i})+1',
                   r["平日営業時間"], r["土日営業時間"], r["区分"], "○" if r["深夜"] else "",
                   f"=Z{i}", f"=G{i}", f'=IF(AR{i}="平日 昼から",1,0)', f'=IF(AR{i}="平日 16時から",1,0)',
                   f'=IF(AS{i}="○",1,0)',
                   f"=係数!$K$3+係数!$K$4*AT{i}+係数!$K$5*AU{i}+係数!$K$6*AV{i}+係数!$K$7*AW{i}+係数!$K$8*AX{i}",
                   f"=M{i}-AY{i}", f"=M{i}/AY{i}", f"=RANK(BA{i},{R('BA')},0)",
                   f'=COUNTIFS({R("C")},C{i},{R("BA")},">"&BA{i})+1',
                   f'=COUNTIFS({R("D")},D{i},{R("BA")},">"&BA{i})+1'])
        for c, fmt in ((46, "0.000"), (51, "#,##0"), (52, "#,##0;[Red]-#,##0"), (53, "0.0%")):
            ws.cell(i, c).number_format = fmt
        for c, fmt in ((34, "0.000"), (36, "#,##0"), (37, "#,##0;[Red]-#,##0"), (38, "0.0%")):
            ws.cell(i, c).number_format = fmt
        for c, fmt in ((8, "#,##0"), (9, "#,##0"), (10, "#,##0"), (13, "#,##0"), (15, "0.000"), (16, "#,##0"),
                       (17, "#,##0;[Red]-#,##0"), (18, "0.0%"), (25, "#,##0"), (26, "0.000"), (27, "#,##0"),
                       (28, "#,##0;[Red]-#,##0"), (29, "0.0%")):
            ws.cell(i, c).number_format = fmt
    ws.conditional_formatting.add(f"R2:R{last}", ColorScaleRule(
        start_type="num", start_value=0.7, start_color="F8696B", mid_type="num", mid_value=1, mid_color="FFFFFF",
        end_type="num", end_value=1.3, end_color="63BE7B"))
    ws.conditional_formatting.add(f"AC2:AC{last}", ColorScaleRule(
        start_type="num", start_value=0.7, start_color="F8696B", mid_type="num", mid_value=1, mid_color="FFFFFF",
        end_type="num", end_value=1.3, end_color="63BE7B"))
    ws.freeze_panes = "C2"
    ws.conditional_formatting.add(f"AL2:AL{last}", ColorScaleRule(
        start_type="num", start_value=0.7, start_color="F8696B", mid_type="num", mid_value=1, mid_color="FFFFFF",
        end_type="num", end_value=1.3, end_color="63BE7B"))
    ws.conditional_formatting.add(f"BA2:BA{last}", ColorScaleRule(
        start_type="num", start_value=0.7, start_color="F8696B", mid_type="num", mid_value=1, mid_color="FFFFFF",
        end_type="num", end_value=1.3, end_color="63BE7B"))
    ws.auto_filter.ref = f"A1:BD{last}"
    widths = [9, 26, 8, 14, 7, 22, 7, 11, 11, 11, 11, 10, 13, 15, 9, 13, 14, 11, 9, 9, 9,
              10, 9, 11, 12, 9, 13, 14, 11, 9, 9, 9, 15, 9, 7, 14, 14, 11, 9, 9, 9,
              13, 13, 13, 9, 9, 7, 8, 8, 8, 15, 14, 12, 9, 9, 9]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 42
    S = "店舗別の期待月商!"

    # ---- 係数
    wc = wb.create_sheet("係数", 0)
    rng_x, rng_y = f"{S}{R('O')}", f"{S}{R('M')}"
    wc.append([f"期待月商の計算式：期待月商 ＝ a ＋ b × ln({KM}km人口)（全店の実績に当てはめた線）"])
    wc.append([])
    wc.append(["a（切片）", f"=INTERCEPT({rng_y},{rng_x})"])
    wc.append(["b（人口の効き方）", f"=SLOPE({rng_y},{rng_x})"])
    wc.append(["当てはまり R²（0〜1。人口で説明できる割合）", f"=RSQ({rng_y},{rng_x})"])
    wc.append(["対象店数", f"=COUNT({rng_y})"])
    wc.append([])
    wc["D1"] = f"カニバリ考慮：期待月商 ＝ a ＋ b × ln(実質{KM}km人口)"
    wc["D3"], wc["E3"] = "a（切片）", f"=INTERCEPT({S}{R('M')},{S}{R('Z')})"
    wc["D4"], wc["E4"] = "b（人口の効き方）", f"=SLOPE({S}{R('M')},{S}{R('Z')})"
    wc["D5"], wc["E5"] = "当てはまり R²", f"=RSQ({S}{R('M')},{S}{R('Z')})"
    wc["E3"].number_format = wc["E4"].number_format = "#,##0"
    wc["E5"].number_format = "0.000"
    lx = f"{S}$AH$2:$AI${last}"
    wc["G1"] = f"カニバリ＋席数：期待月商 ＝ a ＋ b × ln(実質{KM}km人口) ＋ c × 席数"
    wc["G3"], wc["H3"] = "a（切片）", f"=INDEX(LINEST({S}{R('M')},{lx}),1,3)"
    wc["G4"], wc["H4"] = "b（人口の効き方）", f"=INDEX(LINEST({S}{R('M')},{lx}),1,2)"
    wc["G5"], wc["H5"] = "c（1席当たり）", f"=INDEX(LINEST({S}{R('M')},{lx}),1,1)"
    wc["G6"], wc["H6"] = "当てはまり R²", f"=INDEX(LINEST({S}{R('M')},{lx},TRUE,TRUE),3,1)"
    for c in ("H3", "H4", "H5"):
        wc[c].number_format = "#,##0"
    wc["H6"].number_format = "0.000"
    wc.column_dimensions["G"].width = 30
    wc.column_dimensions["H"].width = 16
    wc.append([f"{KM}km人口の例", "期待月商(円)"])
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
    wc.append([f"実質{KM}km人口：500mメッシュごとの人口を、そこを{KM}km圏に含む営業中のきんぐ全店で等分し、店ごとに足したもの。"
               f"近くに自社店がある店ほど、元の{KM}km人口より小さくなる。他業態・他社の競合店は含まない。"])
    lx4 = f"{S}$AT$2:$AX${last}"
    wc["J1"] = f"カニバリ＋席数＋営業時間：期待月商 ＝ a ＋ b×ln(実質{KM}km人口) ＋ c×席数 ＋ d×平日昼から ＋ e×平日16時から ＋ f×24時以降まで"
    for k, (lab, idx) in enumerate([("a（切片）", 6), ("b（人口の効き方）", 5), ("c（1席当たり）", 4),
                                     ("d（平日昼から営業）", 3), ("e（平日16時から営業）", 2), ("f（24時以降まで営業）", 1)], start=3):
        wc[f"J{k}"], wc[f"K{k}"] = lab, f"=INDEX(LINEST({S}{R('M')},{lx4}),1,{idx})"
        wc[f"K{k}"].number_format = "#,##0"
    wc["J9"], wc["K9"] = "当てはまり R²", f"=INDEX(LINEST({S}{R('M')},{lx4},TRUE,TRUE),3,1)"
    wc["K9"].number_format = "0.000"
    wc["J10"] = "d・e・f は「平日17時から・23時まで」の店と比べた月商の差。立地の良さも混ざるので、営業時間を変えたときの増え方そのものではない。"
    wc.column_dimensions["J"].width = 30
    wc.column_dimensions["K"].width = 16
    wc.column_dimensions["A"].width = 44
    wc.column_dimensions["B"].width = 18
    wc.column_dimensions["D"].width = 30
    wc.column_dimensions["E"].width = 16
    for row in wc.iter_rows():
        for c in row:
            if c.fill != HEAD:
                c.font = Font(name=FONT, size=10, bold=c.row == 1)

    # ---- 人口帯別の平均
    wb_ = wb.create_sheet("人口帯別の平均", 1)
    wb_.append([f"{KM}km人口の帯", "店数", f"平均 {KM}km人口", f"平均 {KM}km世帯数", "平均 月商(円)", "月商 中央値(円)", "月商 最小(円)",
                "月商 最大(円)", "平均 席数", "1席当たり月商(円)", "人口1人当たり月商(円)"])
    style_head(wb_, 1)
    labels = [band_label(lo, hi) for lo, hi in BANDS]
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
    # 実質人口（自社店で分け合った後）の帯ごとの平均
    top = wb_.max_row + 3
    wb_.cell(top - 1, 1, f"実質{KM}km人口（近くのきんぐと分け合った後の人口）の帯ごとの平均").font = Font(name=FONT, bold=True, size=11)
    hdr = [f"実質{KM}km人口の帯", "店数", f"平均 実質{KM}km人口", f"平均 {KM}km人口", "平均 月商(円)", "月商 中央値(円)",
           "月商 最小(円)", "月商 最大(円)", "平均 席数", "1席当たり月商(円)", f"平均 {KM}km以内の他のきんぐ数"]
    for j, h in enumerate(hdr, 1):
        wb_.cell(top, j, h)
    style_head(wb_, top)
    for k, (lo, hi) in enumerate(EFF_BANDS, start=top + 1):
        lab = band_label(lo, hi)
        crit = f"{S}{R('AG')},$A{k}"
        vals = [lab, f"=COUNTIF({crit})", f"=AVERAGEIF({crit},{S}{R('Y')})", f"=AVERAGEIF({crit},{S}{R('H')})",
                f"=AVERAGEIF({crit},{S}{R('M')})", None, f"=_xlfn.MINIFS({S}{R('M')},{crit})",
                f"=_xlfn.MAXIFS({S}{R('M')},{crit})", f"=AVERAGEIF({crit},{S}{R('G')})", f"=E{k}/I{k}",
                f"=AVERAGEIF({crit},{S}{R('W')})"]
        for j, v in enumerate(vals, 1):
            if v is not None:
                wb_.cell(k, j, v)
        wb_.cell(k, 6).value = ArrayFormula(f"F{k}", f"=MEDIAN(IF({S}{R('AG')}=$A{k},{S}{R('M')}))")
        for c in range(3, 12):
            wb_.cell(k, c).number_format = "#,##0.0" if c in (9, 11) else "#,##0"
    wb_.row_dimensions[top].height = 32
    # 平日の開店時間別の平均
    top2 = wb_.max_row + 3
    wb_.cell(top2 - 1, 1, "平日の開店時間別の平均").font = Font(name=FONT, bold=True, size=11)
    hdr2 = ["平日の開店", "店数", f"平均 実質{KM}km人口", "平均 席数", "平均 月商(円)", "平均 達成率・カニバリ＋席数",
            "平均 達成率・カニバリ＋席数＋営業時間"]
    for j, h_ in enumerate(hdr2, 1):
        wb_.cell(top2, j, h_)
    style_head(wb_, top2)
    for k, cat in enumerate(["平日 昼から", "平日 16時から", "平日 17時から"], start=top2 + 1):
        crit = f"{S}{R('AR')},$A{k}"
        for j, v in enumerate([cat, f"=COUNTIF({crit})", f"=AVERAGEIF({crit},{S}{R('Y')})", f"=AVERAGEIF({crit},{S}{R('G')})",
                               f"=AVERAGEIF({crit},{S}{R('M')})", f"=AVERAGEIF({crit},{S}{R('AL')})",
                               f"=AVERAGEIF({crit},{S}{R('BA')})"], 1):
            wb_.cell(k, j, v)
        for c, fmt in ((3, "#,##0"), (4, "#,##0.0"), (5, "#,##0"), (6, "0.0%"), (7, "0.0%")):
            wb_.cell(k, c).number_format = fmt
    wb_.row_dimensions[top2].height = 32

    # ---- ブロック・AM別のまとめ
    wg = wb.create_sheet("ブロック・AM別", 2)
    ghead = ["区分", "名前", "店数", f"平均 {KM}km人口", "平均 月商(円)", "平均 期待月商(円)", "平均 達成率",
             "最高の達成率", "最低の達成率", "達成率100%未満の店数", "達成率が最も高い店", "達成率が最も低い店",
             "平均 達成率・カニバリ考慮", f"平均 {KM}km以内の他のきんぐ数", "平均 達成率・カニバリ＋席数", "平均 達成率・カニバリ＋席数＋営業時間", "平日昼から営業の店数"]
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
            wg.cell(i, 13).value = f"=AVERAGEIF({crit},{S}{R('AC')})"
            wg.cell(i, 14).value = f"=AVERAGEIF({crit},{S}{R('W')})"
            wg.cell(i, 15).value = f"=AVERAGEIF({crit},{S}{R('AL')})"
            wg.cell(i, 15).number_format = "0.0%"
            wg.cell(i, 16).value = f"=AVERAGEIF({crit},{S}{R('BA')})"
            wg.cell(i, 16).number_format = "0.0%"
            wg.cell(i, 17).value = f'=COUNTIFS({crit},{S}{R("AR")},"平日 昼から")'
            for c, fmt in ((4, "#,##0"), (5, "#,##0"), (6, "#,##0"), (7, "0.0%"), (8, "0.0%"), (9, "0.0%"), (13, "0.0%"), (14, "0.0")):
                wg.cell(i, c).number_format = fmt
    gl = wg.max_row
    wg.conditional_formatting.add(f"G2:G{gl}", ColorScaleRule(
        start_type="num", start_value=0.85, start_color="F8696B", mid_type="num", mid_value=1, mid_color="FFFFFF",
        end_type="num", end_value=1.15, end_color="63BE7B"))
    wg.conditional_formatting.add(f"M2:M{gl}", ColorScaleRule(
        start_type="num", start_value=0.85, start_color="F8696B", mid_type="num", mid_value=1, mid_color="FFFFFF",
        end_type="num", end_value=1.15, end_color="63BE7B"))
    wg.conditional_formatting.add(f"O2:O{gl}", ColorScaleRule(
        start_type="num", start_value=0.85, start_color="F8696B", mid_type="num", mid_value=1, mid_color="FFFFFF",
        end_type="num", end_value=1.15, end_color="63BE7B"))
    wg.conditional_formatting.add(f"P2:P{gl}", ColorScaleRule(
        start_type="num", start_value=0.85, start_color="F8696B", mid_type="num", mid_value=1, mid_color="FFFFFF",
        end_type="num", end_value=1.15, end_color="63BE7B"))
    wg.auto_filter.ref = f"A1:Q{gl}"
    wg.freeze_panes = "C2"
    for i, w in enumerate([10, 16, 6, 12, 13, 13, 9, 9, 9, 10, 24, 24, 11, 10, 11, 12, 9], 1):
        wg.column_dimensions[get_column_letter(i)].width = w
    wg.row_dimensions[1].height = 32

    # ---- 上位・下位（全店）
    wr = wb.create_sheet("差の大きい店", 3)
    rs = sorted(rows, key=lambda r: r["達成率"])
    head = ["区分", "順位", "店舗コード", "店舗名", "ブロック", "AMエリア", f"{KM}km人口", "実績月商(円)", "期待月商(円)", "差(円)", "達成率", "席数"]
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

    # ---- 上位・下位（カニバリ考慮）
    wr2 = wb.create_sheet("差の大きい店(カニバリ考慮)", 4)
    rs2 = sorted(rows, key=lambda r: r["達成率2"])
    wr2.append(["区分", "順位", "店舗コード", "店舗名", "ブロック", "AMエリア", f"{KM}km人口", f"実質{KM}km人口", f"{KM}km以内の他のきんぐ数",
                "実績月商(円)", "期待月商・カニバリ考慮(円)", "差(円)", "達成率・カニバリ考慮", "達成率(人口のみ)", "席数"])
    style_head(wr2, 1)
    for title, items in (("期待より大きい（上位20）", rs2[::-1][:20]), ("期待より小さい（下位20）", rs2[:20])):
        for k, r in enumerate(items, 1):
            j = wr2.max_row + 1
            ref = f"MATCH($C{j},{S}{R('A')},0)"
            wr2.append([title, k, int(r["店舗コード"]), r["店舗名"], r["ブロック"], r["営業部"], r["人口5km"], r["実質人口"],
                        r["他店数5km"], f"=INDEX({S}{R('M')},{ref})", f"=INDEX({S}{R('AA')},{ref})", f"=J{j}-K{j}",
                        f"=J{j}/K{j}", f"=INDEX({S}{R('R')},{ref})", r["席数"]])
            for c, fmt in ((7, "#,##0"), (8, "#,##0"), (10, "#,##0"), (11, "#,##0"), (12, "#,##0;[Red]-#,##0"), (13, "0.0%"), (14, "0.0%")):
                wr2.cell(j, c).number_format = fmt
    wr2.auto_filter.ref = f"A1:O{wr2.max_row}"
    wr2.freeze_panes = "E2"
    for i, w in enumerate([20, 6, 10, 26, 8, 14, 11, 11, 9, 13, 13, 13, 10, 10, 7], 1):
        wr2.column_dimensions[get_column_letter(i)].width = w
    wr2.row_dimensions[1].height = 32

    # ---- 上位・下位（カニバリ＋席数＋営業時間）
    wr4 = wb.create_sheet("差の大きい店(営業時間込み)", 3)
    rs4 = sorted(rows, key=lambda r: r["達成率4"])
    wr4.append(["区分", "順位", "店舗コード", "店舗名", "ブロック", "AMエリア", f"実質{KM}km人口", "席数", "平日の開店",
                "24時以降", "実績月商(円)", "期待月商・営業時間込み(円)", "差(円)", "達成率・営業時間込み", "達成率・カニバリ＋席数"])
    style_head(wr4, 1)
    for title, items in (("期待より大きい（上位20）", rs4[::-1][:20]), ("期待より小さい（下位20）", rs4[:20])):
        for k, r in enumerate(items, 1):
            j = wr4.max_row + 1
            ref = f"MATCH($C{j},{S}{R('A')},0)"
            wr4.append([title, k, int(r["店舗コード"]), r["店舗名"], r["ブロック"], r["営業部"], r["実質人口"], r["席数"],
                        r["区分"], "○" if r["深夜"] else "", f"=INDEX({S}{R('M')},{ref})", f"=INDEX({S}{R('AY')},{ref})",
                        f"=K{j}-L{j}", f"=K{j}/L{j}", f"=INDEX({S}{R('AL')},{ref})"])
            for c, fmt in ((7, "#,##0"), (11, "#,##0"), (12, "#,##0"), (13, "#,##0;[Red]-#,##0"), (14, "0.0%"), (15, "0.0%")):
                wr4.cell(j, c).number_format = fmt
    wr4.auto_filter.ref = f"A1:O{wr4.max_row}"
    wr4.freeze_panes = "E2"
    for i, w in enumerate([20, 6, 10, 26, 8, 14, 11, 7, 12, 7, 13, 14, 13, 11, 11], 1):
        wr4.column_dimensions[get_column_letter(i)].width = w
    wr4.row_dimensions[1].height = 32

    # ---- 上位・下位（カニバリ＋席数）
    wr3 = wb.create_sheet("差の大きい店(カニバリ＋席数)", 3)
    rs3 = sorted(rows, key=lambda r: r["達成率3"])
    wr3.append(["区分", "順位", "店舗コード", "店舗名", "ブロック", "AMエリア", f"実質{KM}km人口", "席数",
                "実績月商(円)", "期待月商・カニバリ＋席数(円)", "差(円)", "達成率・カニバリ＋席数", "達成率(人口のみ)"])
    style_head(wr3, 1)
    for title, items in (("期待より大きい（上位20）", rs3[::-1][:20]), ("期待より小さい（下位20）", rs3[:20])):
        for k, r in enumerate(items, 1):
            j = wr3.max_row + 1
            ref = f"MATCH($C{j},{S}{R('A')},0)"
            wr3.append([title, k, int(r["店舗コード"]), r["店舗名"], r["ブロック"], r["営業部"], r["実質人口"], r["席数"],
                        f"=INDEX({S}{R('M')},{ref})", f"=INDEX({S}{R('AJ')},{ref})", f"=I{j}-J{j}", f"=I{j}/J{j}",
                        f"=INDEX({S}{R('R')},{ref})"])
            for c, fmt in ((7, "#,##0"), (9, "#,##0"), (10, "#,##0"), (11, "#,##0;[Red]-#,##0"), (12, "0.0%"), (13, "0.0%")):
                wr3.cell(j, c).number_format = fmt
    wr3.auto_filter.ref = f"A1:M{wr3.max_row}"
    wr3.freeze_panes = "E2"
    for i, w in enumerate([20, 6, 10, 26, 8, 14, 11, 7, 13, 14, 13, 11, 10], 1):
        wr3.column_dimensions[get_column_letter(i)].width = w
    wr3.row_dimensions[1].height = 32

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
