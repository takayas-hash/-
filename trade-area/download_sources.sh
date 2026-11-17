#!/usr/bin/env bash
# 公的統計の元ファイルを data/src/ に取得する（キー不要）。
# 必要な通信先: www.e-stat.go.jp, www.soumu.go.jp, www.stat.go.jp（住所変換は msearch.gsi.go.jp）
set -euo pipefail
cd "$(dirname "$0")"; mkdir -p data/src/daytime data/src/meshcity
ES="https://www.e-stat.go.jp/stat-search/file-download?fileKind=0&statInfId="
curl -sSfo data/src/census2020_main.xlsx "${ES}000032143614"   # 国勢調査2020 市区町村別の主な結果
curl -sSfo data/src/hj2023_43-4.xlsx    "${ES}000040209921"     # 住宅・土地統計調査2023 第43-4表
curl -sSfo data/src/juki_r8_city.xlsx   https://www.soumu.go.jp/main_content/000892952.xlsx   # 住基 R8.1.1 市区町村別
curl -sSfo data/src/kazei_r7_11b.xlsx   https://www.soumu.go.jp/main_sosiki/jichi_zeisei/czaisei/czaisei_seido/xls/J51-25-b.xlsx  # 課税状況等の調 R7 第11表
for p in $(seq 1 47); do
  pp=$(printf %02d "$p")
  # 国勢調査2020 従業地・通学地集計 第1-1表（都道府県別分割版）
  if [ "$p" -le 18 ]; then id=$((32221756 + p - 1)); else id=$((32221994 + p - 19)); fi
  curl -sSfo "data/src/daytime/$pp.xlsx" "${ES}0000${id}"
  # 市区町村別メッシュ・コード一覧（統計局）
  curl -sSfo "data/src/meshcity/$pp.csv" "https://www.stat.go.jp/data/mesh/csv/$pp.csv"
done
echo done
