"""値上げ前後の売上前年比と、商圏の人口・世帯年収の関係を調べる。

入力:
  data/sales/YYYYMM.csv   月次売上（社内データ。コミットしない）
  output/stores_trade_area.csv  build.py の結果
使い方:
  python analyze_price.py --pre 202604,202605,202606 --post 202608,202609 --out output/price_analysis.csv
"""
import argparse
import csv
import datetime as dt
import os

import numpy as np
from scipy import stats

VARS = [
    ("商圏3km_人口", "3km人口"),
    ("商圏5km_人口", "5km人口"),
    ("商圏5km_世帯数", "5km世帯数"),
    ("平均世帯年収_万円_推計", "市区町村の平均世帯年収"),
    ("商圏5km_推計平均世帯年収_万円", "5km商圏の推計世帯年収"),
    ("納税者1人当たり課税対象所得_万円", "納税者1人当たり課税所得"),
    ("商圏5km_65歳以上比率%", "5km 65歳以上比率"),
    ("商圏5km_核家族世帯比率%", "5km 核家族世帯比率"),
    ("商圏5km_6歳未満のいる世帯比率%", "5km 6歳未満のいる世帯比率"),
    ("市区町村_昼夜間人口比率", "昼夜間人口比率"),
    ("席数", "席数"),
]


def f(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def load_sales(months):
    data = {}
    for m in months:
        for r in csv.DictReader(open(os.path.join("data/sales", f"{m}.csv"), encoding="utf-8")):
            code = str(int(r["店舗コード"].strip()))
            data.setdefault(code, {})[m] = r
    return data


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pre", required=True)
    ap.add_argument("--post", required=True)
    ap.add_argument("--open-before", default="2025-03-31", help="この日までに開店した店だけ使う（前年が丸1年ある店）")
    ap.add_argument("--out", default="output/price_analysis.csv")
    a = ap.parse_args()
    pre, post = a.pre.split(","), a.post.split(",")
    sales = load_sales(pre + post)
    area = {r["店舗コード"]: r for r in csv.DictReader(open("output/stores_trade_area.csv", encoding="utf-8-sig"))}
    cutoff = dt.date.fromisoformat(a.open_before)

    rows, excluded = [], []
    for code, ms in sales.items():
        t = area.get(code, {})
        name = next(iter(ms.values()))["店舗名称"].strip()
        reason = None
        od = t.get("オープン日", "")
        try:
            y, mo, d = map(int, od.split("/"))
            if dt.date(y, mo, d) > cutoff:
                reason = f"開店が{od}（前年が丸1年ない）"
        except ValueError:
            reason = "オープン日不明"
        if not reason and any(m not in ms for m in pre + post):
            reason = "月のデータが欠けている"
        if not reason and any(f(ms[m]["売上前年"]) in (None, 0) for m in pre + post):
            reason = "前年売上なし"
        if not reason and any(ms[m]["稼動日数"] != ms[m]["暦日付"] for m in pre + post):
            reason = "今年に休業日あり（改装など）"
        if not reason and t.get("開店状態") not in ("営業中",):
            reason = f"開店状態: {t.get('開店状態')}"
        if reason:
            excluded.append((code, name, reason))
            continue

        def tot(months, col):
            return sum(f(ms[m][col]) or 0 for m in months)

        s_pre, s_post = tot(pre, "売上実績"), tot(post, "売上実績")
        p_pre, p_post = tot(pre, "売上前年"), tot(post, "売上前年")
        c_pre, c_post = tot(pre, "客数実績"), tot(post, "客数実績")
        cp_pre, cp_post = tot(pre, "客数前年"), tot(post, "客数前年")
        if not (c_pre and c_post and cp_pre and cp_post):
            excluded.append((code, name, "客数データなし"))
            continue
        o = dict(店舗コード=code, 店舗名=name, 営業部=next(iter(ms.values()))["営業部名称"].strip(),
                 直営FC=t.get("直営FC"), 市区町村=t.get("市区町村名"))
        o["値上げ前_月平均売上"] = s_pre / len(pre)
        o["値上げ後_月平均売上"] = s_post / len(post)
        o["値上げ前_売上前年比"] = s_pre / p_pre
        o["値上げ後_売上前年比"] = s_post / p_post
        o["売上前年比の変化pt"] = (s_post / p_post - s_pre / p_pre) * 100
        o["値上げ前_客数前年比"] = c_pre / cp_pre if cp_pre else None
        o["値上げ後_客数前年比"] = c_post / cp_post if cp_post else None
        o["客数前年比の変化pt"] = (o["値上げ後_客数前年比"] - o["値上げ前_客数前年比"]) * 100 if cp_pre and cp_post else None
        o["値上げ前_客単価"] = s_pre / c_pre
        o["値上げ後_客単価"] = s_post / c_post
        # 異常値チェック（前年に休業があった店など）
        if not (0.6 < o["値上げ前_売上前年比"] < 1.6 and 0.6 < o["値上げ後_売上前年比"] < 1.6):
            excluded.append((code, name, f"前年比が極端（{o['値上げ前_売上前年比']:.2f}→{o['値上げ後_売上前年比']:.2f}）前年の休業などの影響"))
            continue
        for k, _ in VARS:
            o[k] = f(t.get(k))
        rows.append(o)

    os.makedirs("output", exist_ok=True)
    with open(a.out, "w", newline="", encoding="utf-8-sig") as fp:
        w = csv.DictWriter(fp, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    with open(a.out.replace(".csv", "_excluded.csv"), "w", newline="", encoding="utf-8-sig") as fp:
        w = csv.writer(fp)
        w.writerow(["店舗コード", "店舗名", "除外理由"])
        w.writerows(excluded)

    print(f"対象 {len(rows)}店 / 除外 {len(excluded)}店")
    ys = {
        "値上げ前の売上前年比": "値上げ前_売上前年比",
        "値上げ後の売上前年比": "値上げ後_売上前年比",
        "前年比の変化(後-前)": "売上前年比の変化pt",
        "客数前年比の変化": "客数前年比の変化pt",
        "値上げ後の月平均売上": "値上げ後_月平均売上",
    }
    res = []
    for k, label in VARS:
        line = {"指標": label}
        for yl, yk in ys.items():
            pairs = [(r[k], r[yk]) for r in rows if r[k] is not None and r[yk] is not None]
            x, y = np.array(pairs).T
            pr, pp = stats.pearsonr(x, y)
            sr, sp = stats.spearmanr(x, y)
            line[yl] = (round(pr, 2), round(pp, 3), round(sr, 2), len(x))
        res.append(line)
    print("\n相関係数 r（p値, 順位相関, n）")
    for line in res:
        print(line["指標"].ljust(16), " | ".join(f"{yl}: r={v[0]:+.2f} (p={v[1]:.3f}, ρ={v[2]:+.2f})" for yl, v in line.items() if yl != "指標"))

    # 年収5分位ごとの平均
    for key in ("商圏5km_推計平均世帯年収_万円", "商圏5km_人口"):
        print(f"\n{key} の5分位ごとの平均")
        vals = sorted([r for r in rows if r[key] is not None], key=lambda r: r[key])
        n = len(vals)
        for i in range(5):
            g = vals[i * n // 5:(i + 1) * n // 5]
            print(f"  第{i+1}分位 {g[0][key]:>9,.0f}〜{g[-1][key]:>9,.0f}: 店数{len(g)}  前{np.mean([r['値上げ前_売上前年比'] for r in g])*100:.1f}%"
                  f" → 後{np.mean([r['値上げ後_売上前年比'] for r in g])*100:.1f}%  変化{np.mean([r['売上前年比の変化pt'] for r in g]):+.1f}pt"
                  f"  客数変化{np.mean([r['客数前年比の変化pt'] for r in g]):+.1f}pt  後の月売上{np.mean([r['値上げ後_月平均売上'] for r in g])/1e4:,.0f}万円")

    # 重回帰（標準化係数）: 前年比の変化 ~ 年収 + 人口 + 65歳以上比率 + 核家族比率 + 昼夜間比率
    xs = ["商圏5km_推計平均世帯年収_万円", "商圏5km_人口", "商圏5km_65歳以上比率%", "商圏5km_核家族世帯比率%", "市区町村_昼夜間人口比率"]
    for yk in ("売上前年比の変化pt", "値上げ後_月平均売上"):
        d = [r for r in rows if all(r[k] is not None for k in xs)]
        X = np.array([[r[k] for k in xs] for r in d])
        y = np.array([r[yk] for r in d])
        Xz = (X - X.mean(0)) / X.std(0)
        yz = (y - y.mean()) / y.std()
        A = np.column_stack([np.ones(len(d)), Xz])
        beta, *_ = np.linalg.lstsq(A, yz, rcond=None)
        resid = yz - A @ beta
        r2 = 1 - (resid ** 2).sum() / ((yz - yz.mean()) ** 2).sum()
        dof = len(d) - A.shape[1]
        se = np.sqrt(np.diag((resid ** 2).sum() / dof * np.linalg.inv(A.T @ A)))
        tv = beta / se
        pv = 2 * stats.t.sf(np.abs(tv), dof)
        print(f"\n重回帰: {yk}（n={len(d)}, 決定係数R²={r2:.3f}）標準化係数")
        for k, b, p in zip(xs, beta[1:], pv[1:]):
            print(f"  {k:<30} β={b:+.2f}  p={p:.3f}")


if __name__ == "__main__":
    main()
