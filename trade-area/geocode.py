"""店舗住所 → 緯度経度（国土地理院 住所検索API・無料・キー不要）。

使い方:
  python geocode.py data/stores.csv data/geocoded.csv            # 問い合わせて変換
  python geocode.py data/geocoded.csv data/geocoded.csv --relevel  # 問い合わせず一致レベルだけ付け直す

入力CSVは「店舗コード」「住所」列を持つこと。結果に 緯度,経度,一致住所,一致レベル を追加する。
"""
import csv
import json
import re
import sys
import time
import unicodedata
import urllib.parse
import urllib.request

API = "https://msearch.gsi.go.jp/address-search/AddressSearch?q="
KANJI = {"〇": 0, "一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}


def clean(addr: str) -> str:
    """ビル名・階数を除き、ハイフン類と「9番13」表記をそろえる。"""
    a = unicodedata.normalize("NFKC", addr).strip()
    a = re.split(r"[ 　]", a)[0]
    a = re.sub(r"(?<=\d)[‐ｰー−–](?=\d)", "-", a)
    a = re.sub(r"(\d)番(\d)", r"\1-\2", a)
    return a


def kanji2num(s: str) -> str:
    """「十三丁目」→「13丁目」のように漢数字を算用数字にする（一致レベル判定用）。"""

    def conv(m):
        total, cur = 0, 0
        for ch in m.group(0):
            if ch in KANJI:
                cur = KANJI[ch]
            else:
                total += (cur or 1) * {"十": 10, "百": 100, "千": 1000}[ch]
                cur = 0
        return str(total + cur)

    return re.sub(r"[〇一二三四五六七八九十百千]+(?=丁|番|号|条|線)", conv, s)


def level(query: str, title: str) -> str:
    """問い合わせた住所の数字（丁目・番地・号）がどこまで一致したか。"""
    query = kanji2num(query)
    t = kanji2num(unicodedata.normalize("NFKC", title))
    nq = re.findall(r"\d+", query)
    nt = re.findall(r"\d+", t)
    if not nq:
        return "町名まで一致"
    if nt[: len(nq)] == nq:
        return "番地まで一致"
    if len(nq) >= 2 and nt[:2] == nq[:2]:
        return "番地の一部まで一致"
    if nt and nt[0] == nq[0]:
        return "丁目まで一致"
    return "町名の代表点"


def candidates(q: str):
    """見つからないときに順に試す表記ゆれ（大字・字の除去、旧区名、番地なし）。"""
    out = [q]
    q2 = re.sub(r"(大字|字)", "", q)
    q2 = re.sub(r"(奥州市水沢)区", r"\1", q2)  # 例: 奥州市水沢区 → 奥州市水沢
    out.append(q2)
    out.append(re.sub(r"[\d\-]+$", "", q2))
    return list(dict.fromkeys(out))


def geocode(addr):
    with urllib.request.urlopen(API + urllib.parse.quote(addr), timeout=30) as r:
        res = json.load(r)
    if not res:
        return None
    f = res[0]
    lon, lat = f["geometry"]["coordinates"]
    return lat, lon, f["properties"]["title"]


def write(rows, dst):
    with open(dst, "w", newline="", encoding="utf-8") as fp:
        w = csv.DictWriter(fp, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)


def main(src, dst):
    out = []
    for r in csv.DictReader(open(src, encoding="utf-8")):
        r.update(緯度="", 経度="", 一致住所="", 一致レベル="")
        if r.get("住所"):
            q = clean(r["住所"])
            try:
                g = None
                for cand in candidates(q):
                    g = geocode(cand)
                    if g:
                        break
                if g:
                    r["緯度"], r["経度"], r["一致住所"] = g
                    r["一致レベル"] = level(q, g[2])
                else:
                    r["一致レベル"] = "変換不可"
            except Exception as e:  # noqa: BLE001
                r["一致レベル"] = f"エラー: {e}"
            time.sleep(0.5)  # 公共APIへの負荷を抑える
        else:
            r["一致レベル"] = "住所なし"
        out.append(r)
        print(r["店舗コード"], r["一致レベル"], r["一致住所"], flush=True)
    write(out, dst)


def relevel(src, dst):
    rows = list(csv.DictReader(open(src, encoding="utf-8")))
    for r in rows:
        if r["一致住所"]:
            r["一致レベル"] = level(clean(r["住所"]), r["一致住所"])
    write(rows, dst)


if __name__ == "__main__":
    if "--retry-failed" in sys.argv:  # 変換不可の行だけ問い合わせ直す
        rows = list(csv.DictReader(open(sys.argv[1], encoding="utf-8")))
        for r in rows:
            if r["一致レベル"] == "変換不可":
                q = clean(r["住所"])
                for cand in candidates(q):
                    g = geocode(cand)
                    time.sleep(0.5)
                    if g:
                        r["緯度"], r["経度"], r["一致住所"] = g
                        r["一致レベル"] = level(q, g[2])
                        break
                print(r["店舗コード"], r["一致レベル"], r["一致住所"])
        write(rows, sys.argv[2])
    elif "--relevel" in sys.argv:
        relevel(sys.argv[1], sys.argv[2])
    else:
        main(sys.argv[1], sys.argv[2])
