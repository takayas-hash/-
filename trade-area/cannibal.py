"""自社店（きんぐ）どうしの商圏の重なり（カニバリ）を店ごとに計算する。

各500mメッシュの人口を、そのメッシュを半径5km圏に含む営業中のきんぐ全店で等分し、
店ごとに足し上げたものを「実質5km人口（自社店で分け合った後）」とする。
使い方: python cannibal.py output/stores_trade_area.csv output/cannibal_10km.csv 10   # 最後の数字が商圏の半径km
"""
import csv
import sys
from collections import defaultdict

from build import haversine, load_mesh1, meshes_within

R = float(sys.argv[3]) if len(sys.argv) > 3 else 5.0


def main(src, dst):
    stores = [r for r in csv.DictReader(open(src, encoding="utf-8-sig"))
              if r["緯度"] and r["開店状態"] == "営業中"]
    pts = [(r["店舗コード"], float(r["緯度"]), float(r["経度"])) for r in stores]
    cover = defaultdict(list)  # メッシュ → それを商圏（半径R km）に含む店
    store_meshes = {}
    for code, lat, lon in pts:
        ms = [m for m, _ in meshes_within(lat, lon, R)]
        store_meshes[code] = ms
        for m in ms:
            cover[m].append(code)
    out = []
    for code, lat, lon in pts:
        d = sorted(haversine(lat, lon, la, lo) for c, la, lo in pts if c != code)
        pop = eff = shared = 0.0
        for m in store_meshes[code]:
            rec = load_mesh1(m[:4]).get(m)
            if not rec:
                continue
            k = len(cover[m])
            pop += rec["pop"]
            eff += rec["pop"] / k
            if k > 1:
                shared += rec["pop"]
        out.append({
            "店舗コード": code,
            "最寄りのきんぐまでの距離km": round(d[0], 2) if d else "",
            "3km以内の他のきんぐ数": sum(x <= 3 for x in d),
            "5km以内の他のきんぐ数": sum(x <= 5 for x in d),
            "10km以内の他のきんぐ数": sum(x <= 10 for x in d),
            "商圏で他店と重なる人口の割合%": round(shared / pop * 100, 1) if pop else "",
            "実質商圏人口(自社店で分け合った後)": round(eff),
            "商圏人口": round(pop),
        })
    with open(dst, "w", newline="", encoding="utf-8-sig") as fp:
        w = csv.DictWriter(fp, fieldnames=list(out[0].keys()))
        w.writeheader()
        w.writerows(out)
    print(f"{len(out)}店")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
