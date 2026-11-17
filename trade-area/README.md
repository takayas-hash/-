# trade-area: きんぐ店舗の商圏データ（人口・世帯・年収）

店舗住所から、所在市区町村と半径3km・5km商圏の人口・世帯数・世帯年収（推計）をまとめる集計スクリプト。
**店舗一覧・統計の元ファイル・結果の表はコミットしない**（`.gitignore` で `data/` `output/` などを除外）。

## 手順
1. 店舗一覧を `data/stores.csv`（列: 店舗コード,直営FC,店舗名,開店状態,ブロック,郵便番号,住所,卓数,席数,坪数,オープン日）に置く
2. `./download_sources.sh` で公的統計を取得
3. `python geocode.py data/stores.csv data/geocoded.csv`（国土地理院の住所検索で緯度経度に変換）
4. `python build.py data/geocoded.csv output/stores_trade_area.csv`（500mメッシュは e-Stat から自動取得・キャッシュ。`--codes 2180,2225` で一部の店だけ）
5. `python make_xlsx.py output/stores_trade_area.csv output/kingu_trade_area.xlsx`（Googleスプレッドシートに取り込む3シートの表）

必要なPythonパッケージ: openpyxl。出典・計算方法は make_xlsx.py の SOURCES と、出来上がった表の「出典・定義・計算方法」シートに記載。
