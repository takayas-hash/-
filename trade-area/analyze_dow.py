"""曜日差抜き前年比（社内シート「曜日差抜き前年比」）で、値上げ前後の変化と商圏データの関係を調べる。

入力:
  data/sales/dow_adjusted_raw.json  シートの値（get_values の結果 {range, values}）。社内データのためコミットしない
  output/stores_trade_area.csv      build.py の結果
使い方:
  python analyze_dow.py --pre 2026年4月,2026年5月,2026年6月 --post 2026年8月,2026/09/01,2026/10/01
"""
import argparse
import csv
import json

import numpy as np
from scipy import stats

from analyze_price import VARS, f


def pct(v):
    try:
        x = float(str(v).replace("%", "")) / 100
    except ValueError:
        return None
    return x if x > 0 else None  # 0.0% は営業日数の違いで除外された月


def corr_table(rows, ys):
    out = []
    for k, label in VARS:
        line = [label]
        for _, yk in ys:
            pairs = [(r[k], r[yk]) for r in rows if r.get(k) is not None and r.get(yk) is not None]
            x, y = np.array(pairs).T
            r_, p_ = stats.pearsonr(x, y)
            rho, _ = stats.spearmanr(x, y)
            line.append((r_, p_, rho, len(x)))
        out.append(line)
    return out


def quintiles(rows, key, cols):
    vals = sorted([r for r in rows if r.get(key) is not None], key=lambda r: r[key])
    n = len(vals)
    res = []
    for i in range(5):
        g = vals[i * n // 5:(i + 1) * n // 5]
        res.append((i + 1, g[0][key], g[-1][key], len(g), [np.mean([r[c] for r in g]) for c in cols]))
    return res


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default="data/sales/dow_adjusted_raw.json")
    ap.add_argument("--pre", required=True)
    ap.add_argument("--post", required=True)
    ap.add_argument("--out", default="output/price_analysis_dow.csv")
    a = ap.parse_args()
    pre, post = a.pre.split(","), a.post.split(",")
    vals = json.load(open(a.src))["values"]
    header = vals[0]
    col = {h: i for i, h in enumerate(header) if h}
    area = {r["店舗コード"]: r for r in csv.DictReader(open("output/stores_trade_area.csv", encoding="utf-8-sig"))}

    rows, excluded = [], []
    for r in vals[5:]:
        if len(r) < 46 or not r[3]:
            continue
        code, name = r[3].strip(), r[4].strip()
        s = {m: pct(r[col[m]]) for m in pre + post}
        c = {m: pct(r[col[m] + 1]) for m in pre + post}
        if any(v is None for v in list(s.values()) + list(c.values())):
            excluded.append((code, name, "前年と営業日数が違う月がある（シート上0.0%）"))
            continue
        if any(not 0.6 < v < 1.6 for v in s.values()):
            excluded.append((code, name, "前年比が極端な月がある"))
            continue
        t = area.get(code, {})
        if t.get("開店状態") != "営業中":
            excluded.append((code, name, f"開店状態: {t.get('開店状態')}"))
            continue
        o = dict(店舗コード=code, 店舗名=name, 直営FC=r[1], AM=r[5])
        o["値上げ前_売上前年比"] = np.mean([s[m] for m in pre])
        o["値上げ後_売上前年比"] = np.mean([s[m] for m in post])
        o["値上げ前_客数前年比"] = np.mean([c[m] for m in pre])
        o["値上げ後_客数前年比"] = np.mean([c[m] for m in post])
        o["売上前年比の変化pt"] = (o["値上げ後_売上前年比"] - o["値上げ前_売上前年比"]) * 100
        o["客数前年比の変化pt"] = (o["値上げ後_客数前年比"] - o["値上げ前_客数前年比"]) * 100
        o["最終月_売上前年比"] = s[post[-1]]
        o["最終月_客数前年比"] = c[post[-1]]
        for k, _ in VARS:
            o[k] = f(t.get(k))
        rows.append(o)

    with open(a.out, "w", newline="", encoding="utf-8-sig") as fp:
        w = csv.DictWriter(fp, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    with open(a.out.replace(".csv", "_excluded.csv"), "w", newline="", encoding="utf-8-sig") as fp:
        csv.writer(fp).writerows([("店舗コード", "店舗名", "除外理由")] + excluded)

    print(f"対象 {len(rows)}店 / 除外 {len(excluded)}店")
    print(f"全体平均: 売上前年比 {np.mean([r['値上げ前_売上前年比'] for r in rows])*100:.1f}% → {np.mean([r['値上げ後_売上前年比'] for r in rows])*100:.1f}%"
          f" / 客数前年比 {np.mean([r['値上げ前_客数前年比'] for r in rows])*100:.1f}% → {np.mean([r['値上げ後_客数前年比'] for r in rows])*100:.1f}%")
    ys = [("売上の変化", "売上前年比の変化pt"), ("客数の変化", "客数前年比の変化pt"), ("最終月の客数前年比", "最終月_客数前年比")]
    print("\n相関 r (p値, 順位相関ρ)")
    for line in corr_table(rows, ys):
        print(f"{line[0]:<16}", " | ".join(f"{yl}: r={v[0]:+.2f} (p={v[1]:.3f}, ρ={v[2]:+.2f})" for (yl, _), v in zip(ys, line[1:])))
    cols = ["値上げ前_売上前年比", "値上げ後_売上前年比", "売上前年比の変化pt", "値上げ前_客数前年比", "値上げ後_客数前年比", "客数前年比の変化pt"]
    for key in ("商圏5km_推計平均世帯年収_万円", "納税者1人当たり課税対象所得_万円", "商圏5km_人口", "商圏5km_65歳以上比率%"):
        print(f"\n{key} 5分位")
        for q, lo, hi, n, m in quintiles(rows, key, cols):
            print(f"  第{q} {lo:>9,.0f}〜{hi:>9,.0f} n={n}: 売上 {m[0]*100:.1f}→{m[1]*100:.1f}% ({m[2]:+.1f}pt)  客数 {m[3]*100:.1f}→{m[4]*100:.1f}% ({m[5]:+.1f}pt)")

    xs = ["商圏5km_推計平均世帯年収_万円", "商圏5km_人口", "商圏5km_65歳以上比率%", "商圏5km_核家族世帯比率%", "市区町村_昼夜間人口比率"]
    for yk in ("売上前年比の変化pt", "客数前年比の変化pt"):
        d = [r for r in rows if all(r[k] is not None for k in xs)]
        X = np.array([[r[k] for k in xs] for r in d]); y = np.array([r[yk] for r in d])
        A = np.column_stack([np.ones(len(d)), (X - X.mean(0)) / X.std(0)])
        yz = (y - y.mean()) / y.std()
        beta, *_ = np.linalg.lstsq(A, yz, rcond=None)
        resid = yz - A @ beta
        dof = len(d) - A.shape[1]
        se = np.sqrt(np.diag((resid ** 2).sum() / dof * np.linalg.inv(A.T @ A)))
        pv = 2 * stats.t.sf(np.abs(beta / se), dof)
        print(f"\n重回帰 {yk} n={len(d)} R²={1 - (resid**2).sum() / (yz**2).sum():.3f}")
        for k, b, p in zip(xs, beta[1:], pv[1:]):
            print(f"  {k:<30} β={b:+.2f} p={p:.3f}")


if __name__ == "__main__":
    main()
