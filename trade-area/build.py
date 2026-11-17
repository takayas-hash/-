"""店舗ごとの商圏データ（人口・世帯・年収）を作る。

前提: geocode.py で data/geocoded.csv（店舗コード・住所・緯度・経度…）を作っておく。
統計ファイルは data/src/ に置く（download_sources.sh 参照）。メッシュは自動で取得・キャッシュする。

使い方:
  python build.py data/geocoded.csv output/stores_trade_area.csv [--codes 2180,2225,...]
"""
import argparse
import csv
import glob
import io
import math
import os
import re
import unicodedata
import urllib.request
import zipfile
from collections import defaultdict

import openpyxl

SRC = os.path.join(os.path.dirname(__file__), "data", "src")
MESH_DIR = os.path.join(os.path.dirname(__file__), "data", "mesh")
MESH_URL = "https://www.e-stat.go.jp/gis/statmap-search/data?statsId=T001141&code={code}&downloadType=2"
RADII_KM = (3, 5)

# 住宅・土地統計調査 世帯の年間収入階級(10区分) の階級値（万円）。最上位は仮定値。
INCOME_CLASSES = [
    ("01", 0, 100, 50), ("02", 100, 150, 125), ("03", 150, 200, 175), ("05", 200, 300, 250),
    ("06", 300, 400, 350), ("07", 400, 500, 450), ("08", 500, 700, 600), ("09", 700, 1000, 850),
    ("10", 1000, 1500, 1250), ("11", 1500, None, 2000),
]


def num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def norm(s: str) -> str:
    return unicodedata.normalize("NFKC", s or "").replace("ヶ", "ケ").replace("ヵ", "カ").strip()


def rows(path, sheet=0):
    return openpyxl.load_workbook(path, read_only=True).worksheets[sheet].iter_rows(values_only=True)


# ---------------------------------------------------------------- 市区町村データ
def load_census_main():
    """国勢調査2020 都道府県・市区町村別の主な結果。code5 -> dict。政令市の区→市の対応も返す。"""
    out, parent, last_city = {}, {}, None
    for i, r in enumerate(rows(os.path.join(SRC, "census2020_main.xlsx"))):
        if i < 9 or not r[1]:
            continue
        code, name = r[1].split("_", 1)
        kind = r[3]
        if kind == 1:
            last_city = code
        elif kind == 0 and last_city:
            parent[code] = last_city
        gen_hh = num(r[39])
        out[code] = dict(
            name=name, pop=num(r[4]), hh=num(r[35]), gen_hh=gen_hh,
            r1564=num(r[18]), r65=num(r[19]),
            couple_child=num(r[42]),
            r_couple_child=(num(r[42]) / gen_hh * 100) if gen_hh else None,
            r_single=(num(r[45]) / gen_hh * 100) if gen_hh else None,
        )
    return out, parent


def load_daytime():
    """国勢調査2020 従業地・通学地集計 第1-1表（都道府県別ファイル）。code5 -> (夜間, 昼間)。"""
    out = {}
    for f in sorted(glob.glob(os.path.join(SRC, "daytime", "*.xlsx"))):
        for r in rows(f):
            if len(r) > 17 and r[3] == "0_総数" and r[4] == "00_総数" and r[2] and "_" in str(r[2]):
                out[r[2].split("_")[0]] = (num(r[5]), num(r[17]))
    return out


def load_juki():
    """住民基本台帳 令和8年1月1日 市区町村別。code5 -> dict, (都道府県, 名称) 一覧。"""
    out, names = {}, []
    for i, r in enumerate(rows(os.path.join(SRC, "juki_r8_city.xlsx"))):
        if i < 6 or not r[0] or r[0] == "-" or r[2] in (None, "-"):
            continue
        code = str(r[0])[:5]
        out[code] = dict(name=r[2], pop=num(r[5]), hh=num(r[6]))
        names.append((r[1], r[2], code))
    return out, names


def load_housing_income():
    """住宅・土地統計調査2023 第43-4表。code5 -> {mean, median, n}（万円）。"""
    dist = defaultdict(dict)
    for i, r in enumerate(rows(os.path.join(SRC, "hj2023_43-4.xlsx"))):
        if i < 7 or r[2] != "0_総数" or not r[1]:
            continue
        code = r[1].split("_")[0]
        cls = str(r[3]).split("_")[0]
        dist[code][cls] = num(r[4])
    out = {}
    for code, d in dist.items():
        counts = [(lo, hi, mid, d.get(c) or 0) for c, lo, hi, mid in INCOME_CLASSES]
        n = sum(x[3] for x in counts)
        if n <= 0:
            continue
        mean = sum(mid * k for lo, hi, mid, k in counts) / n
        half, cum, median = n / 2, 0, None
        for lo, hi, mid, k in counts:
            if k and cum + k >= half:
                median = lo + ((half - cum) / k) * ((hi or lo * 2) - lo) if hi or lo else mid
                break
            cum += k
        out[code] = dict(mean=mean, median=median, n=n, total=d.get("00"))
    return out


def load_tax():
    """市町村税課税状況等の調 令和7年度 第11表。code5 -> 課税対象所得/所得割納税義務者（万円）。"""
    out = {}
    for i, r in enumerate(rows(os.path.join(SRC, "kazei_r7_11b.xlsx"))):
        if i < 4 or r[4] != "市町村民税":
            continue
        n, inc = num(r[5]), num(r[13])
        if n:
            out[str(r[1])[:5]] = dict(per=inc * 1000 / n / 10000, n=n, name=r[3])
    return out


def load_mesh_city():
    """1kmメッシュ → 市区町村コード（統計局 市区町村別メッシュ・コード一覧）。"""
    m = defaultdict(list)
    for f in glob.glob(os.path.join(SRC, "meshcity", "*.csv")):
        for line in open(f, encoding="cp932").read().splitlines()[1:]:
            p = line.split(",")
            if len(p) >= 3 and p[2].strip():
                m[p[2].strip()].append(p[0].strip())
    return m


# ---------------------------------------------------------------- 住所 → 市区町村
def match_city(addr, juki_names):
    a = norm(addr)
    best = None
    for pref, name, code in juki_names:
        pref, name = norm(pref), norm(name)
        if not a.startswith(pref):
            continue
        rest = a[len(pref):]
        cands = {name, re.sub(r"^.+?郡", "", name)}
        for c in cands:
            if c and rest.startswith(c) and (best is None or len(c) > best[0]):
                best = (len(c), code, norm(pref) + name)
    return (best[1], best[2]) if best else (None, None)


# ---------------------------------------------------------------- メッシュ
def mesh_code(lat, lon):
    """点を含む4次メッシュ（500m）コード。"""
    p = int(lat * 1.5); a = lat * 60 - p * 40
    q = int(a / 5); b = a - q * 5
    r = int(b * 60 / 30); c = b * 60 - r * 30
    s = int(c / 15)
    u = int(lon - 100); f = lon - 100 - u
    v = int(f * 60 / 7.5); g = f * 60 - v * 7.5
    w = int(g * 60 / 45); h = g * 60 - w * 45
    t = int(h / 22.5)
    return f"{p}{u:02d}{q}{v}{r}{w}{s * 2 + t + 1}"


def haversine(lat1, lon1, lat2, lon2):
    R = 6371.0088
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp, dl = p2 - p1, math.radians(lon2 - lon1)
    return 2 * R * math.asin(math.sqrt(math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2))


def meshes_within(lat, lon, km):
    """中心点が半径km以内にある4次メッシュ（中心点方式）。"""
    dlat = 15 / 3600; dlon = 22.5 / 3600
    nlat = int(km / 111.0 / dlat) + 2
    nlon = int(km / (111.32 * math.cos(math.radians(lat))) / dlon) + 2
    i0, j0 = math.floor(lat / dlat), math.floor(lon / dlon)
    out = []
    for i in range(i0 - nlat, i0 + nlat + 1):
        for j in range(j0 - nlon, j0 + nlon + 1):
            clat, clon = (i + 0.5) * dlat, (j + 0.5) * dlon
            d = haversine(lat, lon, clat, clon)
            if d <= km:
                out.append((mesh_code(clat, clon), d))
    return out


MESH_FIELDS = {
    "pop": "人口（総数）", "p0014": "０～１４歳人口　総数", "p1564": "１５～６４歳人口　総数",
    "p65": "６５歳以上人口　総数", "hh": "世帯総数", "gen_hh": "一般世帯数",
    "nuclear": "核家族世帯数　一般世帯数", "u6": "６歳未満世帯員のいる世帯数　一般世帯数",
}
_mesh_cache = {}


def load_mesh1(code1):
    """1次メッシュ単位の500mメッシュ統計（国勢調査2020 T001141）。無ければ e-Stat から取得。"""
    if code1 in _mesh_cache:
        return _mesh_cache[code1]
    os.makedirs(MESH_DIR, exist_ok=True)
    path = os.path.join(MESH_DIR, f"T001141_{code1}.zip")
    if not os.path.exists(path):
        with urllib.request.urlopen(MESH_URL.format(code=code1), timeout=120) as r:
            data = r.read()
        open(path, "wb").write(data)
    data = {}
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile:  # 海上などデータの無い1次メッシュ
        _mesh_cache[code1] = data
        return data
    for n in z.namelist():
        lines = io.TextIOWrapper(z.open(n), encoding="cp932").read().splitlines()
        labels = [x.strip("　 ") for x in lines[1].split(",")]
        idx = {k: labels.index(v) for k, v in MESH_FIELDS.items()}
        for line in lines[2:]:
            p = line.split(",")
            data[p[0]] = {k: (num(p[i]) or 0.0) for k, i in idx.items()}
    _mesh_cache[code1] = data
    return data


# ---------------------------------------------------------------- 本体
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("dst")
    ap.add_argument("--codes", help="対象の店舗コード（カンマ区切り）。省略で全店")
    a = ap.parse_args()

    census, parent = load_census_main()
    daytime = load_daytime()
    juki, juki_names = load_juki()
    housing = load_housing_income()
    tax = load_tax()
    meshcity = load_mesh_city()
    name2code = {norm(v["name"]): k for k, v in census.items()}

    def stat_code(code, name):
        """2020年以降に区が再編された市（浜松市）は市全体の値で代用する。"""
        if code in census:
            return code, ""
        city = re.match(r"(.+?市)", norm(name or "").split("県")[-1].split("府")[-1].split("都")[-1])
        if city and norm(city.group(1)) in name2code:
            return name2code[norm(city.group(1))], "区の再編後のため市全体の値"
        return None, "統計に該当なし"

    def housing_for(code):
        if code in housing:
            return housing[code], "市区町村"
        if code in parent and parent[code] in housing:
            return housing[parent[code]], "政令市全体"
        pref = code[:2] + "000"
        if pref in housing:
            return housing[pref], "都道府県（町村は非公表のため代用）"
        return None, ""

    def tax_for(code):
        if code in tax:
            return tax[code], "市区町村"
        if code in parent and parent[code] in tax:
            return tax[parent[code]], "政令市全体（区別の公表なし）"
        return None, ""

    stores = list(csv.DictReader(open(a.src, encoding="utf-8")))
    if a.codes:
        want = set(a.codes.split(","))
        stores = [s for s in stores if s["店舗コード"] in want]

    out = []
    for s in stores:
        o = {k: s.get(k, "") for k in ["店舗コード", "店舗名", "直営FC", "開店状態", "ブロック", "郵便番号", "住所",
                                       "卓数", "席数", "坪数", "オープン日"]}
        o.update(緯度=s.get("緯度", ""), 経度=s.get("経度", ""), 位置の精度=s.get("一致レベル", ""))
        notes = []
        jcode, jname = match_city(s.get("住所", ""), juki_names) if s.get("住所") else (None, None)
        o["市区町村コード"] = jcode or ""
        o["市区町村名"] = jname or ""
        if jcode:
            scode, note = stat_code(jcode, jname)
            if note:
                notes.append(f"国勢調査・住調: {note}")
            c = census.get(scode) if scode else None
            if c:
                o.update({
                    "市区町村_人口_2020国調": int(c["pop"]), "市区町村_世帯数_2020国調": int(c["hh"]),
                    "市区町村_15-64歳比率%": round(c["r1564"], 1), "市区町村_65歳以上比率%": round(c["r65"], 1),
                    "市区町村_夫婦と子供世帯比率%": round(c["r_couple_child"], 1),
                    "市区町村_単独世帯比率%": round(c["r_single"], 1),
                })
            dt = daytime.get(scode) if scode else None
            if dt and dt[0]:
                o["市区町村_昼間人口_2020"] = int(dt[1])
                o["市区町村_昼夜間人口比率"] = round(dt[1] / dt[0] * 100, 1)
            j = juki.get(jcode)
            if j:
                o["市区町村_人口_2026住基"] = int(j["pop"])
                o["市区町村_世帯数_2026住基"] = int(j["hh"])
            h, unit = housing_for(scode or jcode)
            if h:
                o["平均世帯年収_万円_推計"] = round(h["mean"])
                o["世帯年収中央値_万円_推計"] = round(h["median"]) if h["median"] else ""
                o["年収の集計単位"] = unit
            t, tunit = tax_for(jcode if jcode in tax else (scode or jcode))
            if t:
                o["納税者1人当たり課税対象所得_万円"] = round(t["per"], 1)
                o["課税所得の集計単位"] = tunit
        else:
            notes.append("住所から市区町村を特定できず")

        # 商圏（半径）
        if s.get("緯度"):
            lat, lon = float(s["緯度"]), float(s["経度"])
            for km in RADII_KM:
                tot = defaultdict(float)
                inc_w = inc_sum = tax_w = tax_sum = 0.0
                for code9, _d in meshes_within(lat, lon, km):
                    rec = load_mesh1(code9[:4]).get(code9)
                    if not rec:
                        continue
                    for k, v in rec.items():
                        tot[k] += v
                    cities = meshcity.get(code9[:8]) or ([jcode] if jcode else [])
                    share = 1 / len(cities) if cities else 0
                    for cc in cities:
                        h, unit = housing_for(cc)
                        if h and not unit.startswith("都道府県"):
                            inc_w += rec["hh"] * share
                            inc_sum += rec["hh"] * share * h["mean"]
                        t, _ = tax_for(cc)
                        if t:
                            tax_w += rec["pop"] * share
                            tax_sum += rec["pop"] * share * t["per"]
                p = f"商圏{km}km_"
                pop, hh, gh = tot["pop"], tot["hh"], tot["gen_hh"]
                o[p + "人口"] = round(pop)
                o[p + "世帯数"] = round(hh)
                o[p + "15-64歳比率%"] = round(tot["p1564"] / pop * 100, 1) if pop else ""
                o[p + "65歳以上比率%"] = round(tot["p65"] / pop * 100, 1) if pop else ""
                o[p + "核家族世帯比率%"] = round(tot["nuclear"] / gh * 100, 1) if gh else ""
                o[p + "6歳未満のいる世帯比率%"] = round(tot["u6"] / gh * 100, 1) if gh else ""
                o[p + "推計平均世帯年収_万円"] = round(inc_sum / inc_w) if inc_w else ""
                o[p + "年収データのある世帯の割合%"] = round(inc_w / hh * 100) if hh else ""
                o[p + "推計課税所得_万円"] = round(tax_sum / tax_w, 1) if tax_w else ""
                if km == RADII_KM[-1] and hh and inc_w / hh < 0.7:
                    notes.append(f"{km}km商圏の3割超が年収非公表の町村（年収推計の信頼性低）")
        else:
            notes.append("緯度経度なしのため商圏集計なし")
        o["備考"] = " / ".join(notes)
        out.append(o)
        print(o["店舗コード"], o["店舗名"], o.get("市区町村名"), o.get("商圏3km_人口"), o.get("商圏5km_人口"),
              o.get("平均世帯年収_万円_推計"), o.get("商圏5km_推計平均世帯年収_万円"), flush=True)

    keys = []
    for o in out:
        for k in o:
            if k not in keys:
                keys.append(k)
    os.makedirs(os.path.dirname(a.dst) or ".", exist_ok=True)
    with open(a.dst, "w", newline="", encoding="utf-8-sig") as fp:
        w = csv.DictWriter(fp, fieldnames=keys)
        w.writeheader()
        w.writerows(out)


if __name__ == "__main__":
    main()
