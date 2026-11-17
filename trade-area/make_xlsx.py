"""build.py の結果CSVから、Googleスプレッドシートに取り込む .xlsx（3シート）を作る。

使い方: python make_xlsx.py output/stores_trade_area.csv output/kingu_trade_area.xlsx
"""
import csv
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

FONT = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
GROUP_FILLS = {"店舗": "D9E1F2", "市区町村": "E2EFDA", "年収": "FFF2CC", "商圏3km": "FCE4D6", "商圏5km": "EDE2F6"}

SOURCES = [
    ("項目", "出典（統計名・表）", "調査年・時点", "集計単位", "計算方法・注意"),
    ("緯度・経度", "国土地理院 住所検索API（msearch.gsi.go.jp）", "2026年11月取得", "住所",
     "ビル名・階数を除いた住所で検索。「位置の精度」列に一致の度合いを記載（番地まで一致／丁目まで一致／町名の代表点など）。"),
    ("市区町村コード・名", "総務省 住民基本台帳に基づく人口・世帯数 市区町村別（令和8年1月1日）の市区町村名", "2026年1月1日", "市区町村（政令市は区）",
     "住所の先頭を市区町村名と照合。2024年に区が再編された浜松市の区は、国勢調査・住宅土地統計調査では浜松市全体の値を使用。"),
    ("市区町村_人口・世帯数_2020国調、15-64歳比率、65歳以上比率、夫婦と子供世帯比率、単独世帯比率",
     "総務省統計局 令和2年国勢調査 都道府県・市区町村別の主な結果（e-Stat statInfId=000032143614）", "2020年10月1日", "市区町村（政令市は区）",
     "年齢比率は不詳補完値。夫婦と子供世帯比率・単独世帯比率＝各世帯数÷一般世帯数。"),
    ("市区町村_昼間人口・昼夜間人口比率", "令和2年国勢調査 従業地・通学地による人口・就業状態等集計 第1-1表（都道府県別分割版）", "2020年10月1日",
     "市区町村（政令市は区）", "昼夜間人口比率＝昼間人口÷夜間人口×100。100超は働きに来る人・買い物客が多い地域。"),
    ("市区町村_人口・世帯数_2026住基", "総務省 住民基本台帳に基づく人口、人口動態及び世帯数（市区町村別・総計）", "2026年1月1日", "市区町村（政令市は区）",
     "住民票の数。国勢調査と定義が違うため、増減の比較には使わない。"),
    ("平均世帯年収・世帯年収中央値（推計）", "総務省統計局 令和5年住宅・土地統計調査 第43-4表 世帯の年間収入階級(10区分)別主世帯数（e-Stat statInfId=000040209921）",
     "2023年10月1日（収入は前年1年間）", "市区町村（政令市は区）。人口1.5万未満の町村は非公表",
     "推計値。各収入階級の主世帯数に階級の中央値（100万円未満=50万円、1500万円以上=2000万円と仮定）を掛けて平均。中央値は累積世帯数が半分になる階級内を直線で補間。年収不詳の世帯は除外。"),
    ("納税者1人当たり課税対象所得", "総務省 令和7年度 市町村税課税状況等の調 第11表（市町村別）", "2025年度課税（2024年の所得）", "市町村（政令市は市全体、東京23区は区）",
     "課税対象所得÷所得割の納税義務者数。世帯年収ではなく、税金を納める人1人当たりの所得。"),
    ("商圏3km・5km_人口・世帯数・年齢比率", "令和2年国勢調査 地域メッシュ統計 4次メッシュ（500m）人口及び世帯（e-Stat 統計GIS T001141）", "2020年10月1日",
     "500mメッシュ", "店舗の緯度経度から半径3km・5kmの円内に中心点が入るメッシュを合計（直線距離。車10分圏の代わりの目安）。秘匿メッシュの数値は隣のメッシュに合算されているため、そのまま合計。"),
    ("商圏_核家族世帯比率・6歳未満のいる世帯比率", "同上（T001141）", "2020年10月1日", "500mメッシュ",
     "メッシュ統計には「夫婦と子供世帯」が無いため、ファミリー層の目安として核家族世帯比率と6歳未満の子がいる世帯の比率を掲載（一般世帯に対する割合）。"),
    ("商圏_推計平均世帯年収", "住宅・土地統計調査2023（上記）＋総務省統計局 市区町村別メッシュ・コード一覧", "2023年", "市区町村の値をメッシュに割り当て",
     "推計値。メッシュごとに所在の市区町村の平均世帯年収を当て、メッシュの世帯数で加重平均。メッシュ内の年収のばらつきは反映されない。年収が非公表の町村のメッシュは除外し、その割合を「年収データのある世帯の割合」に記載。"),
    ("商圏_推計課税所得", "市町村税課税状況等の調（上記）＋メッシュ・コード一覧", "2025年度", "市町村の値をメッシュに割り当て",
     "推計値。メッシュの人口で加重平均。"),
    ("店舗の基本情報（店舗コード〜オープン日）", "社内スプレッドシート「店舗情報」", "2026年11月17日時点", "店舗", "業態「きんぐ」の全行。開店状態が営業中以外の店も残している。"),
]

NEXT_USE = [
    ("次の段階（売上データとの突き合わせ）での使い方",),
    ("・売上データとは「店舗コード」で結合する（VLOOKUP / XLOOKUP、またはコードでの結合）。",),
    ("・回帰で使う候補: 商圏5km_世帯数、商圏3km_人口、商圏5km_推計平均世帯年収、市区町村_昼夜間人口比率、席数、坪数、オープン日からの経過年数。",),
    ("・営業中以外の店（改装休業中・開店準備中）や、開店1年未満の店は推定から外すのが安全。",),
    ("・年収は市区町村単位の値の割り当てなので、同じ市内の店舗どうしでは差が出にくい。",),
]


def group_of(col):
    if col.startswith("商圏3km"):
        return "商圏3km"
    if col.startswith("商圏5km"):
        return "商圏5km"
    if "年収" in col or "課税" in col:
        return "年収"
    if col.startswith("市区町村"):
        return "市区町村"
    return "店舗"


def to_value(k, v):
    if v in ("", None):
        return None
    if k in ("店舗名", "直営FC", "開店状態", "ブロック", "郵便番号", "住所", "位置の精度", "市区町村名", "年収の集計単位",
             "課税所得の集計単位", "備考", "市区町村コード", "オープン日"):
        return v
    try:
        f = float(v)
        return int(f) if f.is_integer() else f
    except ValueError:
        return v


def main(src, dst):
    rows = list(csv.DictReader(open(src, encoding="utf-8-sig")))
    for r in rows:
        if r["オープン日"].startswith("1900"):
            r["オープン日"] = ""
        if r["坪数"] in ("0", "0.0"):
            r["坪数"] = ""
        if not r["開店状態"]:
            r["開店状態"] = "（店舗情報が空欄）"
    cols = list(rows[0].keys())
    wb = Workbook()

    ws = wb.active
    ws.title = "店舗別一覧"
    ws.append(cols)
    for r in rows:
        ws.append([to_value(k, r[k]) for k in cols])
    for i, c in enumerate(cols, 1):
        cell = ws.cell(1, i)
        cell.font = Font(name=FONT, bold=True, color="000000", size=9)
        cell.fill = PatternFill("solid", fgColor=GROUP_FILLS[group_of(c)])
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        w = 14
        if c in ("店舗名", "市区町村名"):
            w = 24
        elif c == "住所":
            w = 36
        elif c == "備考":
            w = 44
        elif c in ("位置の精度", "年収の集計単位", "課税所得の集計単位"):
            w = 18
        ws.column_dimensions[get_column_letter(i)].width = w
        fmt = None
        if "比率" in c or "割合" in c:
            fmt = "0.0"
        elif any(x in c for x in ("人口", "世帯数", "年収")):
            fmt = "#,##0"
        if fmt:
            for row in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                row[0].number_format = fmt
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name=FONT, size=10)
    ws.row_dimensions[1].height = 48
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions

    ws2 = wb.create_sheet("出典・定義・計算方法")
    for r in SOURCES:
        ws2.append(list(r))
    ws2.append([])
    for r in NEXT_USE:
        ws2.append(list(r))
    for i, w in enumerate([34, 52, 22, 26, 90], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    for row in ws2.iter_rows():
        for cell in row:
            cell.font = Font(name=FONT, size=10, bold=cell.row == 1 or (cell.value or "").startswith("次の段階"))
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    for cell in ws2[1]:
        cell.fill = HEAD_FILL
        cell.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
    ws2.freeze_panes = "A2"

    ws3 = wb.create_sheet("変換できなかった店・注意点")
    ws3.append(["区分", "店舗コード", "店舗名", "住所", "内容"])
    for r in rows:
        code, name, addr = r["店舗コード"], r["店舗名"], r["住所"]
        if not addr:
            ws3.append(["店舗情報が空欄", code, name, "", "店舗情報シートに店舗名・住所などが無いため、統計を付けていない。"])
        elif not r["緯度"]:
            ws3.append(["緯度経度に変換できず", code, name, addr,
                        "国土地理院の住所検索で見つからない（「埇」の字が辞書に無い）。市区町村の値のみ記載し、商圏（半径）は空欄。"])
        elif r["位置の精度"] == "町名の代表点":
            ws3.append(["位置が町名の代表点", code, name, addr,
                        "番地まで特定できず、町の代表点を使った。数百m〜1km程度ずれる可能性があり、3km・5km商圏の値にも少し影響する。"])
    for r in rows:
        if r["開店状態"] not in ("営業中", "（店舗情報が空欄）"):
            ws3.append(["営業中以外", r["店舗コード"], r["店舗名"], r["住所"], f"開店状態: {r['開店状態']}。行は残している。"])
    for r in rows:
        if "再編" in r["備考"]:
            ws3.append(["区の再編", r["店舗コード"], r["店舗名"], r["住所"],
                        "浜松市は2024年に区を再編したため、国勢調査・住宅土地統計調査は浜松市全体の値（住民基本台帳は新しい区の値）。"])
    general = [
        ("精度の限界（全体）", "", "", "", "商圏の年収は、市区町村単位の調査結果をメッシュに割り当てた推計。同じ市の中の住宅地の違い（高級住宅地か団地か等）は反映されない。"),
        ("精度の限界（全体）", "", "", "", "商圏は直線距離の円。川・山・幹線道路・競合店による実際の来店範囲とは違う。車10分圏の代わりの目安として5kmを掲載。"),
        ("精度の限界（全体）", "", "", "", "メッシュの人口・世帯は2020年国勢調査。2020年以降に大きな宅地開発があった地域では現状より少なく出る。"),
        ("精度の限界（全体）", "", "", "", "最上位の収入階級（1500万円以上）は2000万円と仮定。高所得の区（港区・中目黒など）は平均年収が実際より低めに出る可能性がある。"),
    ]
    for g in general:
        ws3.append(list(g))
    for i, w in enumerate([22, 12, 28, 40, 90], 1):
        ws3.column_dimensions[get_column_letter(i)].width = w
    for row in ws3.iter_rows():
        for cell in row:
            cell.font = Font(name=FONT, size=10)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    for cell in ws3[1]:
        cell.fill = HEAD_FILL
        cell.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
    ws3.freeze_panes = "A2"
    wb.save(dst)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
