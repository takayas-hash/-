"""日別客数予測ツール (Excel) を生成するスクリプト。

使い方:
    python build_forecast_tool.py 出力.xlsx [先月データ.xlsx] [先々月データ.xlsx]

データファイルを渡すと、その内容を貼り付け済みの状態で出力する（動作確認用）。
渡さなければ空のテンプレートになる。
"""
import sys

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

FONT = "Meiryo UI"
F = Font(name=FONT, size=10)
FB = Font(name=FONT, size=10, bold=True)
FT = Font(name=FONT, size=14, bold=True)
FIN = Font(name=FONT, size=10, color="0000FF")
FW = Font(name=FONT, size=10, bold=True, color="FFFFFF")
FILL_IN = PatternFill("solid", fgColor="FFF2CC")
FILL_HD = PatternFill("solid", fgColor="44546A")
FILL_SUB = PatternFill("solid", fgColor="D9E1F2")
FILL_GR = PatternFill("solid", fgColor="EDEDED")
FILL_KPI = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center")
WRAP = Alignment(wrap_text=True, vertical="top")

TYPES = ["月", "火", "水", "木", "金", "土", "日", "祝"]
HOURS = list(range(11, 24))
DATA_ROWS = 800  # 31日×24時間=744行 + 余裕
DS1, DS2 = "先月データ", "先々月データ"

HOLIDAYS = [
    ("2025-01-01", "元日"), ("2025-01-13", "成人の日"), ("2025-02-11", "建国記念の日"),
    ("2025-02-23", "天皇誕生日"), ("2025-02-24", "振替休日"), ("2025-03-20", "春分の日"),
    ("2025-04-29", "昭和の日"), ("2025-05-03", "憲法記念日"), ("2025-05-04", "みどりの日"),
    ("2025-05-05", "こどもの日"), ("2025-05-06", "振替休日"), ("2025-07-21", "海の日"),
    ("2025-08-11", "山の日"), ("2025-09-15", "敬老の日"), ("2025-09-23", "秋分の日"),
    ("2025-10-13", "スポーツの日"), ("2025-11-03", "文化の日"), ("2025-11-23", "勤労感謝の日"),
    ("2025-11-24", "振替休日"),
    ("2026-01-01", "元日"), ("2026-01-12", "成人の日"), ("2026-02-11", "建国記念の日"),
    ("2026-02-23", "天皇誕生日"), ("2026-03-20", "春分の日"), ("2026-04-29", "昭和の日"),
    ("2026-05-03", "憲法記念日"), ("2026-05-04", "みどりの日"), ("2026-05-05", "こどもの日"),
    ("2026-05-06", "振替休日"), ("2026-07-20", "海の日"), ("2026-08-11", "山の日"),
    ("2026-09-21", "敬老の日"), ("2026-09-22", "国民の休日"), ("2026-09-23", "秋分の日"),
    ("2026-10-12", "スポーツの日"), ("2026-11-03", "文化の日"), ("2026-11-23", "勤労感謝の日"),
    ("2027-01-01", "元日"), ("2027-01-11", "成人の日"), ("2027-02-11", "建国記念の日"),
    ("2027-02-23", "天皇誕生日"), ("2027-03-21", "春分の日"), ("2027-03-22", "振替休日"),
    ("2027-04-29", "昭和の日"), ("2027-05-03", "憲法記念日"), ("2027-05-04", "みどりの日"),
    ("2027-05-05", "こどもの日"), ("2027-07-19", "海の日"), ("2027-08-11", "山の日"),
    ("2027-09-20", "敬老の日"), ("2027-09-23", "秋分の日"), ("2027-10-11", "スポーツの日"),
    ("2027-11-03", "文化の日"), ("2027-11-23", "勤労感謝の日"),
]


def type_formula(cell):
    """日付セルから日タイプ（月〜日 / 祝）を返す式。"""
    return (f'IF(COUNTIF(祝日!$A:$A,{cell})>0,"祝",'
            f'CHOOSE(WEEKDAY({cell},2),"月","火","水","木","金","土","日"))')


def style(c, font=F, fill=None, fmt=None, align=None, border=True):
    c.font = font
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if align:
        c.alignment = align
    if border:
        c.border = BOX


def header(ws, row, col, labels, fill=FILL_HD, font=FW):
    for i, label in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=label)
        style(c, font=font, fill=fill, align=Alignment(horizontal="center", vertical="center", wrap_text=True))


def build(out, src1=None, src2=None):
    wb = openpyxl.Workbook()
    ws_help = wb.active
    ws_help.title = "使い方"
    ws_fc = wb.create_sheet("予測")
    ws_hr = wb.create_sheet("時間帯別予測")
    ws_set = wb.create_sheet("設定")
    ws_d1 = wb.create_sheet(DS1)
    ws_d2 = wb.create_sheet(DS2)
    ws_hol = wb.create_sheet("祝日")
    ws_c = wb.create_sheet("集計")

    # ---------------- データ貼り付けシート ----------------
    for ws, src, label in ((ws_d1, src1, "先月"), (ws_d2, src2, "先々月")):
        ws.sheet_properties.tabColor = "FFC000"
        if src:
            sws = openpyxl.load_workbook(src).active
            for r in sws.iter_rows(values_only=True):
                ws.append(list(r))
            for row in ws.iter_rows(min_row=2, max_col=2):
                for c in row:
                    c.number_format = "yyyy/mm/dd"
        else:
            ws["A1"] = f"← A1セルを選択して、{label}の「時間帯別売上実績表」を見出し行ごと貼り付けてください"
            ws["A1"].font = Font(name=FONT, size=11, bold=True, color="C00000")
        ws["Z1"] = "日タイプ(自動)"
        style(ws["Z1"], font=FB, fill=FILL_GR)
        for r in range(2, DATA_ROWS + 1):
            ws.cell(row=r, column=26, value=f'=IF(ISNUMBER(A{r}),{type_formula(f"A{r}")},"")').font = Font(name=FONT, size=9, color="808080")
        ws.column_dimensions["Z"].width = 14
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 12
        ws.freeze_panes = "A2"

    # ---------------- 祝日 ----------------
    ws_hol["A1"], ws_hol["B1"] = "日付", "名称"
    header(ws_hol, 1, 1, ["日付", "名称"])
    for i, (d, name) in enumerate(HOLIDAYS, start=2):
        y, m, dd = map(int, d.split("-"))
        ws_hol.cell(row=i, column=1, value=f"=DATE({y},{m},{dd})")
        ws_hol.cell(row=i, column=2, value=name)
        style(ws_hol.cell(row=i, column=1), fmt="yyyy/mm/dd(aaa)")
        style(ws_hol.cell(row=i, column=2))
    ws_hol["D1"] = "祝日扱いにしたい日（お盆・年末年始・地域イベント等）は、この表の下に日付を追加してください。"
    ws_hol["D1"].font = FB
    ws_hol.column_dimensions["A"].width = 16
    ws_hol.column_dimensions["B"].width = 16

    # ---------------- 設定 ----------------
    s = ws_set
    s["A1"] = "設定"
    s["A1"].font = FT
    s.column_dimensions["A"].width = 30
    s.column_dimensions["B"].width = 16
    s.column_dimensions["C"].width = 70
    rows = [
        # (row, label, value, fmt, input?, note)
        (3, "先月（データから自動判定）", f'=IF(ISNUMBER(\'{DS1}\'!A2),DATE(YEAR(\'{DS1}\'!A2),MONTH(\'{DS1}\'!A2),1),"")', "yyyy年m月", False, "先月データの1行目の日付から自動判定"),
        (4, "先々月（データから自動判定）", f'=IF(ISNUMBER(\'{DS2}\'!A2),DATE(YEAR(\'{DS2}\'!A2),MONTH(\'{DS2}\'!A2),1),"")', "yyyy年m月", False, "先々月データの1行目の日付から自動判定"),
        (5, "予測する月（当月）", '=IF(B3="","",EDATE(B3,1))', "yyyy年m月", True, "通常は先月の翌月（自動）。別の月にしたい場合は 2026/11/1 のように月初日を直接入力"),
        (6, "当月の日数", '=IF(B5="","",DAY(EOMONTH(B5,0)))', "0", False, ""),
        (8, "先月の重み", 0.6, "0.00", True, "直近ほど重視。先々月と合計1でなくてもOK（比率で按分）"),
        (9, "先々月の重み", 0.4, "0.00", True, ""),
        (11, "トレンド係数", 1.0, "0.00", True, "全体の上げ下げ。例：販促・値上げで5%増を見込む→1.05。下の参考値（前年比）も判断材料に"),
        (12, "前年当月の客数合計（任意）", None, "#,##0", True, "入力すると季節係数を自動計算。空欄なら季節係数=1"),
        (13, "季節係数（自動）", None, "0.000", False, "＝前年当月の日平均 ÷ 前年の先月・先々月の日平均（重み付き）"),
        (15, "【参考】先月 客数前年比", f"=IFERROR(SUM(集計!D3:D33)/SUM(集計!F3:F33),\"\")", "0.0%", False, "当年客数 ÷ 前年客数"),
        (16, "【参考】先々月 客数前年比", f"=IFERROR(SUM(集計!I3:I33)/SUM(集計!K3:K33),\"\")", "0.0%", False, ""),
        (17, "【参考】先月 日平均客数", "=IFERROR(SUM(集計!D3:D33)/DAY(EOMONTH(B3,0)),\"\")", "#,##0.0", False, ""),
        (18, "【参考】先々月 日平均客数", "=IFERROR(SUM(集計!I3:I33)/DAY(EOMONTH(B4,0)),\"\")", "#,##0.0", False, ""),
        (19, "データチェック", '=IF(B3="","先月データ未貼付",IF(B4="","先々月データ未貼付（先月のみで予測中）",IF(EDATE(B4,1)<>B3,"⚠ 先々月が先月の前月になっていません","OK")))', None, False, ""),
    ]
    for r, label, val, fmt, is_in, note in rows:
        style(s.cell(row=r, column=1, value=label), font=FB, fill=FILL_SUB)
        c = s.cell(row=r, column=2, value=val)
        style(c, font=FIN if is_in else F, fill=FILL_IN if is_in else None, fmt=fmt, align=CENTER)
        s.cell(row=r, column=3, value=note).font = Font(name=FONT, size=9, color="595959")
    # 季節係数
    s["B13"] = ('=IFERROR(IF(B12="",1,(B12/B6)/(($B$8*IF(B3="",0,SUM(集計!F3:F33)/DAY(EOMONTH(B3,0)))'
                '+$B$9*IF(B4="",0,SUM(集計!K3:K33)/DAY(EOMONTH(B4,0))))/($B$8*(B3<>"")+$B$9*(B4<>"")))),1)')
    s["A7"] = "▼ 重み・補正"
    s["A7"].font = FB
    s["A14"] = "▼ 判断材料"
    s["A14"].font = FB
    s["A21"] = "黄色セル＝入力欄（青字）。その他は自動計算です。"
    s["A21"].font = Font(name=FONT, size=9, color="595959")

    # ---------------- 集計 ----------------
    c = ws_c
    c["A1"] = "集計（自動計算・編集不要）"
    c["A1"].font = FT
    header(c, 2, 1, ["日", "先月 日付", "タイプ", "客数", "売上", "前年客数",
                     "先々月 日付", "タイプ", "客数", "売上", "前年客数"])
    for i in range(31):
        r = 3 + i
        c.cell(row=r, column=1, value=i + 1)
        for col0, sheet, mcell in ((2, DS1, "設定!$B$3"), (7, DS2, "設定!$B$4")):
            dl = openpyxl.utils.get_column_letter(col0)
            c.cell(row=r, column=col0, value=f'=IF({mcell}="","",IF(A{r}<=DAY(EOMONTH({mcell},0)),{mcell}+A{r}-1,""))')
            c.cell(row=r, column=col0 + 1, value=f'=IF({dl}{r}="","",{type_formula(f"{dl}{r}")})')
            for k, src_col in enumerate(("Q", "L", "R")):
                c.cell(row=r, column=col0 + 2 + k,
                       value=f"=IF({dl}{r}=\"\",\"\",SUMIFS('{sheet}'!${src_col}:${src_col},'{sheet}'!$A:$A,{dl}{r}))")
        for col in range(1, 12):
            cc = c.cell(row=r, column=col)
            style(cc, fmt="m/d(aaa)" if col in (2, 7) else "#,##0", align=CENTER if col in (1, 2, 3, 7, 8) else None)

    # 日タイプ別の基準値
    TR = 37  # タイプ表の先頭行
    c.cell(row=TR - 2, column=1, value="日タイプ別 基準値（客数0の日＝休業日は除外）").font = FB
    header(c, TR - 1, 1, ["タイプ", "先月 日数", "先月 平均客数", "先々月 日数", "先々月 平均客数",
                          "重み付け平均", "採用基準客数", "先月 客単価", "先々月 客単価", "採用客単価"])
    for i, t in enumerate(TYPES):
        r = TR + i
        c.cell(row=r, column=1, value=t)
        c.cell(row=r, column=2, value=f'=COUNTIFS($C$3:$C$33,A{r},$D$3:$D$33,">0")')
        c.cell(row=r, column=3, value=f'=IF(B{r}>0,AVERAGEIFS($D$3:$D$33,$C$3:$C$33,A{r},$D$3:$D$33,">0"),"")')
        c.cell(row=r, column=4, value=f'=COUNTIFS($H$3:$H$33,A{r},$I$3:$I$33,">0")')
        c.cell(row=r, column=5, value=f'=IF(D{r}>0,AVERAGEIFS($I$3:$I$33,$H$3:$H$33,A{r},$I$3:$I$33,">0"),"")')
        c.cell(row=r, column=6, value=(f'=IF(AND(B{r}>0,D{r}>0),(設定!$B$8*C{r}+設定!$B$9*E{r})/(設定!$B$8+設定!$B$9),'
                                        f'IF(B{r}>0,C{r},IF(D{r}>0,E{r},"")))'))
        sun = TR + TYPES.index("日")
        c.cell(row=r, column=7, value=(f'=IF(F{r}<>"",F{r},IF(AND(A{r}="祝",$F${sun}<>""),$F${sun},$G${TR + 9}))'))
        c.cell(row=r, column=8, value=f'=IFERROR(SUMIFS($E$3:$E$33,$C$3:$C$33,A{r})/SUMIFS($D$3:$D$33,$C$3:$C$33,A{r}),"")')
        c.cell(row=r, column=9, value=f'=IFERROR(SUMIFS($J$3:$J$33,$H$3:$H$33,A{r})/SUMIFS($I$3:$I$33,$H$3:$H$33,A{r}),"")')
        c.cell(row=r, column=10, value=(f'=IF(AND(H{r}<>"",I{r}<>""),(設定!$B$8*H{r}+設定!$B$9*I{r})/(設定!$B$8+設定!$B$9),'
                                         f'IF(H{r}<>"",H{r},IF(I{r}<>"",I{r},$J${TR + 9})))'))
        for col in range(1, 11):
            style(c.cell(row=r, column=col), fmt="#,##0" if col in (8, 9, 10) else "#,##0.0" if col in (3, 5, 6, 7) else "0",
                  align=CENTER if col == 1 else None)
    # 全体平均（タイプが一度も出現しない場合の保険）
    r = TR + 9
    c.cell(row=r, column=1, value="全体平均")
    c.cell(row=r, column=7, value=('=IFERROR((設定!$B$8*IFERROR(AVERAGEIF($D$3:$D$33,">0"),0)+設定!$B$9*IFERROR(AVERAGEIF($I$3:$I$33,">0"),0))'
                                   '/(設定!$B$8*(COUNTIF($D$3:$D$33,">0")>0)+設定!$B$9*(COUNTIF($I$3:$I$33,">0")>0)),0)'))
    c.cell(row=r, column=10, value=('=IFERROR((SUM($E$3:$E$33)+SUM($J$3:$J$33))/(SUM($D$3:$D$33)+SUM($I$3:$I$33)),0)'))
    for col in range(1, 11):
        style(c.cell(row=r, column=col), font=FB, fill=FILL_GR, fmt="#,##0.0" if col == 7 else "#,##0")

    # 時間帯構成比
    HR = 52
    c.cell(row=HR - 2, column=1, value="日タイプ×時刻の構成比（1日の客数に占める割合）  ※時刻は見出しを書き換え可").font = FB
    style(c.cell(row=HR - 1, column=1, value="タイプ"), font=FW, fill=FILL_HD, align=CENTER)
    for j, h in enumerate(HOURS):
        cc = c.cell(row=HR - 1, column=2 + j, value=h)
        style(cc, font=FW, fill=FILL_HD, fmt='0"時"', align=CENTER)
    last_col = openpyxl.utils.get_column_letter(1 + len(HOURS))
    style(c.cell(row=HR - 1, column=2 + len(HOURS), value="合計"), font=FW, fill=FILL_HD, align=CENTER)

    # 時刻別の「1日あたり平均客数」を重み付けで作る（行 HR+10 〜 ）
    AR = HR + 11
    c.cell(row=AR - 2, column=1, value="（計算用）日タイプ×時刻の1日あたり平均客数").font = Font(name=FONT, size=9, color="808080")
    for i, t in enumerate(TYPES):
        r_share = HR + i
        r_avg = AR + i
        trow = TR + i
        c.cell(row=r_share, column=1, value=t)
        c.cell(row=r_avg, column=1, value=t)
        for j in range(len(HOURS)):
            hl = openpyxl.utils.get_column_letter(2 + j)
            a1 = (f"IF(OR(設定!$B$3=\"\",$B${trow}=0),\"\",SUMIFS('{DS1}'!$Q:$Q,'{DS1}'!$Z:$Z,$A{r_avg},'{DS1}'!$K:$K,{hl}${HR - 1},"
                  f"'{DS1}'!$A:$A,\">=\"&設定!$B$3,'{DS1}'!$A:$A,\"<=\"&EOMONTH(設定!$B$3,0))/$B${trow})")
            a2 = (f"IF(OR(設定!$B$4=\"\",$D${trow}=0),\"\",SUMIFS('{DS2}'!$Q:$Q,'{DS2}'!$Z:$Z,$A{r_avg},'{DS2}'!$K:$K,{hl}${HR - 1},"
                  f"'{DS2}'!$A:$A,\">=\"&設定!$B$4,'{DS2}'!$A:$A,\"<=\"&EOMONTH(設定!$B$4,0))/$D${trow})")
            # 1か月分しかない場合はその月、両方なければ空
            c.cell(row=r_avg, column=2 + j, value=(
                f'=IF(AND($B${trow}>0,$D${trow}>0),(設定!$B$8*{a1}+設定!$B$9*{a2})/(設定!$B$8+設定!$B$9),'
                f'IF($B${trow}>0,{a1},IF($D${trow}>0,{a2},"")))'))
            style(c.cell(row=r_avg, column=2 + j), font=Font(name=FONT, size=9, color="808080"), fmt="0.0")
        c.cell(row=r_avg, column=2 + len(HOURS), value=f"=SUM(B{r_avg}:{last_col}{r_avg})")
        style(c.cell(row=r_avg, column=2 + len(HOURS)), font=Font(name=FONT, size=9, color="808080"), fmt="0.0")
        style(c.cell(row=r_avg, column=1), font=Font(name=FONT, size=9, color="808080"))
    # 構成比（タイプにデータがなければ、祝→日、それ以外→全タイプ合算 を使う）
    sun_avg = AR + TYPES.index("日")
    tot_col = openpyxl.utils.get_column_letter(2 + len(HOURS))
    for i, t in enumerate(TYPES):
        r_share = HR + i
        r_avg = AR + i
        for j in range(len(HOURS)):
            hl = openpyxl.utils.get_column_letter(2 + j)
            c.cell(row=r_share, column=2 + j, value=(
                f'=IF(N(${tot_col}{r_avg})>0,N({hl}{r_avg})/${tot_col}{r_avg},'
                f'IF(AND($A{r_share}="祝",N(${tot_col}${sun_avg})>0),N({hl}${sun_avg})/${tot_col}${sun_avg},'
                f'IFERROR(SUM({hl}${AR}:{hl}${AR + 7})/SUM(${tot_col}${AR}:${tot_col}${AR + 7}),0)))'))
            style(c.cell(row=r_share, column=2 + j), fmt="0.0%")
        c.cell(row=r_share, column=2 + len(HOURS), value=f"=SUM(B{r_share}:{last_col}{r_share})")
        style(c.cell(row=r_share, column=2 + len(HOURS)), fmt="0.0%", font=FB)
        style(c.cell(row=r_share, column=1), align=CENTER)
    c.column_dimensions["A"].width = 10
    for col in range(2, 16):
        c.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 11
    c.sheet_properties.tabColor = "A6A6A6"

    # ---------------- 予測 ----------------
    f = ws_fc
    f.sheet_properties.tabColor = "70AD47"
    f["A1"] = '=IF(設定!B5="","日別客数予測（先月データを貼り付けてください）",TEXT(設定!B5,"yyyy年m月")&"　日別客数予測")'
    f["A1"].font = FT
    f["A2"] = '="データ: "&設定!B19&"　／　季節係数 "&TEXT(設定!B13,"0.000")&"　トレンド係数 "&TEXT(設定!B11,"0.00")'
    f["A2"].font = Font(name=FONT, size=9, color="595959")
    # KPI
    kpis = [("月間予測客数", "=SUM(H8:H38)", "#,##0"), ("月間予測売上", "=SUM(I8:I38)", "#,##0"),
            ("先月実績客数", "=SUM(集計!D3:D33)", "#,##0"), ("先月比(客数)", '=IFERROR(D3/D5,"")', "0.0%")]
    for k, (lab, fml, fmt) in enumerate(kpis):
        r = 3 + k
        for col in range(1, 6):
            style(f.cell(row=r, column=col), font=FB, fill=FILL_KPI)
        f.cell(row=r, column=1, value=lab)
        f.cell(row=r, column=4, value=fml).number_format = fmt
        f.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        f.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
    f["G3"] = "黄色の「日別補正」に、イベント・天候・販促などの影響を倍率で入力（例：雨予報0.9、地域祭り1.2）"
    f["G3"].font = Font(name=FONT, size=9, color="595959")
    hdr = ["日付", "曜日", "日タイプ", "基準客数", "季節係数", "トレンド係数", "日別補正", "予測客数", "予測売上", "メモ"]
    header(f, 7, 1, hdr)
    for i in range(31):
        r = 8 + i
        f.cell(row=r, column=1, value=f'=IF(OR(設定!$B$5="",{i + 1}>設定!$B$6),"",設定!$B$5+{i})')
        f.cell(row=r, column=2, value=f'=IF(A{r}="","",CHOOSE(WEEKDAY(A{r},2),"月","火","水","木","金","土","日"))')
        f.cell(row=r, column=3, value=f'=IF(A{r}="","",{type_formula(f"A{r}")})')
        f.cell(row=r, column=4, value=f'=IF(A{r}="","",INDEX(集計!$G${TR}:$G${TR + 7},MATCH(C{r},集計!$A${TR}:$A${TR + 7},0)))')
        f.cell(row=r, column=5, value=f'=IF(A{r}="","",設定!$B$13)')
        f.cell(row=r, column=6, value=f'=IF(A{r}="","",設定!$B$11)')
        f.cell(row=r, column=7, value=1)
        f.cell(row=r, column=8, value=f'=IF(A{r}="","",ROUND(D{r}*E{r}*F{r}*N(G{r}),0))')
        f.cell(row=r, column=9, value=f'=IF(A{r}="","",ROUND(H{r}*INDEX(集計!$J${TR}:$J${TR + 7},MATCH(C{r},集計!$A${TR}:$A${TR + 7},0)),-2))')
        fmts = ["m/d", None, None, "#,##0.0", "0.000", "0.00", "0.00", "#,##0", "#,##0", None]
        for col in range(1, 11):
            is_in = col in (7, 10)
            style(f.cell(row=r, column=col), font=FB if col == 8 else (FIN if is_in else F),
                  fill=FILL_IN if is_in else None, fmt=fmts[col - 1], align=CENTER if col in (1, 2, 3) else None)
    r = 39
    style(f.cell(row=r, column=1, value="合計"), font=FB, fill=FILL_SUB)
    for col in range(2, 11):
        style(f.cell(row=r, column=col), fill=FILL_SUB)
    f.cell(row=r, column=8, value="=SUM(H8:H38)")
    f.cell(row=r, column=9, value="=SUM(I8:I38)")
    for col in (8, 9):
        f.cell(row=r, column=col).number_format = "#,##0"
        f.cell(row=r, column=col).font = FB
    # 土日祝の色分け
    f.conditional_formatting.add("A8:C38", FormulaRule(formula=['$C8="土"'], font=Font(color="0070C0", bold=True)))
    f.conditional_formatting.add("A8:C38", FormulaRule(formula=['OR($C8="日",$C8="祝")'], font=Font(color="C00000", bold=True)))
    dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="5", allow_blank=True,
                        error="0〜5の倍率で入力してください（例：1.2）")
    f.add_data_validation(dv)
    dv.add("G8:G38")
    f.cell(row=7, column=7).comment = Comment("1.00=補正なし。1.20=2割増、0.80=2割減。空欄にすると0扱い（休業日）になります。", "tool")
    for col, w in zip("ABCDEFGHIJ", (10, 6, 9, 10, 9, 10, 9, 11, 12, 28)):
        f.column_dimensions[col].width = w
    f.freeze_panes = "A8"
    ch = BarChart()
    ch.type = "col"
    ch.title = "日別 予測客数"
    ch.y_axis.title = "客数"
    ch.height, ch.width = 9, 22
    ch.add_data(Reference(f, min_col=8, min_row=7, max_row=38), titles_from_data=True)
    ch.set_categories(Reference(f, min_col=1, min_row=8, max_row=38))
    ch.legend = None
    ch.x_axis.number_format = "d"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    f.add_chart(ch, "L7")

    # ---------------- 時間帯別予測 ----------------
    h = ws_hr
    h.sheet_properties.tabColor = "70AD47"
    h["A1"] = '=IF(設定!B5="","時間帯別 予測客数",TEXT(設定!B5,"yyyy年m月")&"　時間帯別 予測客数（シフト作成用）")'
    h["A1"].font = FT
    h["A2"] = "日別予測客数 × 日タイプ別の時刻構成比（集計シート）。小数は四捨五入表示のため、合計が日別予測と±数名ずれることがあります。"
    h["A2"].font = Font(name=FONT, size=9, color="595959")
    header(h, 4, 1, ["日付", "タイプ", "日予測"])
    for j in range(len(HOURS)):
        hl = openpyxl.utils.get_column_letter(2 + j)
        cc = h.cell(row=4, column=4 + j, value=f"=集計!{hl}{HR - 1}")
        style(cc, font=FW, fill=FILL_HD, fmt='0"時"', align=CENTER)
    for i in range(31):
        r = 5 + i
        h.cell(row=r, column=1, value=f"=予測!A{8 + i}")
        h.cell(row=r, column=2, value=f"=予測!C{8 + i}")
        h.cell(row=r, column=3, value=f"=予測!H{8 + i}")
        style(h.cell(row=r, column=1), fmt="m/d(aaa)", align=CENTER)
        style(h.cell(row=r, column=2), align=CENTER)
        style(h.cell(row=r, column=3), fmt="#,##0", font=FB)
        for j in range(len(HOURS)):
            hl = openpyxl.utils.get_column_letter(2 + j)
            h.cell(row=r, column=4 + j, value=(
                f'=IF($A{r}="","",$C{r}*INDEX(集計!{hl}${HR}:{hl}${HR + 7},MATCH($B{r},集計!$A${HR}:$A${HR + 7},0)))'))
            style(h.cell(row=r, column=4 + j), fmt="0")
    last_h = openpyxl.utils.get_column_letter(3 + len(HOURS))
    h.conditional_formatting.add(f"D5:{last_h}35", FormulaRule(formula=["AND(ISNUMBER(D5),D5>=1)"], fill=PatternFill("solid", fgColor="FCE4D6")))
    h.conditional_formatting.add(f"D5:{last_h}35", ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="F4B183"))
    h.conditional_formatting.add("A5:B35", FormulaRule(formula=['$B5="土"'], font=Font(color="0070C0", bold=True)))
    h.conditional_formatting.add("A5:B35", FormulaRule(formula=['OR($B5="日",$B5="祝")'], font=Font(color="C00000", bold=True)))
    h.column_dimensions["A"].width = 11
    h.column_dimensions["B"].width = 6
    h.column_dimensions["C"].width = 8
    for j in range(len(HOURS)):
        h.column_dimensions[openpyxl.utils.get_column_letter(4 + j)].width = 6
    h.freeze_panes = "D5"

    # ---------------- 使い方 ----------------
    u = ws_help
    u.sheet_properties.tabColor = "4472C4"
    u.column_dimensions["A"].width = 3
    u.column_dimensions["B"].width = 110
    lines = [
        ("日別客数予測ツール", FT),
        ("", F),
        ("■ 毎月の使い方（3ステップ）", FB),
        ("① 「先月データ」シートのA1を選択 → 先月の時間帯別売上実績表（見出し行ごと）を貼り付け", F),
        ("② 「先々月データ」シートのA1を選択 → 先々月の時間帯別売上実績表を貼り付け", F),
        ("   ※ 貼り付け前に、前回のデータを A〜X 列ごと削除してください（Z列の自動計算列は消さない）", F),
        ("③ 「予測」シートに当月（先月の翌月）の日別予測客数が表示されます。「時間帯別予測」はシフト作成用", F),
        ("   必要に応じて「予測」シートの黄色セル（日別補正）にイベント・天候などの倍率を入力", F),
        ("", F),
        ("■ 予測の考え方", FB),
        ("予測客数 ＝ 日タイプ別の基準客数 × 季節係数 × トレンド係数 × 日別補正", FB),
        ("・日タイプ：月〜日の曜日 ＋「祝」（祝日シートに載っている日）の8区分。焼肉業態は曜日差が非常に大きいため、まず曜日で分けるのが最重要", F),
        ("・基準客数：先月・先々月それぞれの日タイプ別平均客数を、重み（先月0.6／先々月0.4）で加重平均。客数0の日は休業日として除外", F),
        ("   祝日が過去2カ月に無い場合は日曜の値、データが全く無いタイプは全体平均で代用", F),
        ("・季節係数：「設定」に前年当月の客数合計を入れると、前年の『先月・先々月 → 当月』の伸び率で補正（お盆・年末などの季節変動を反映）", F),
        ("・トレンド係数：販促・値上げ・競合出店など全体的な上げ下げを手入力（参考に前年比を表示）", F),
        ("・日別補正：給料日後の週末、地域イベント、雨予報など、その日だけの要因", F),
        ("・時間帯別：日予測 × 同じ日タイプの時刻別構成比（過去2カ月の平均）", F),
        ("・予測売上：予測客数 × 日タイプ別の客単価（加重平均）", F),
        ("", F),
        ("■ シート一覧", FB),
        ("予測／時間帯別予測：結果　設定：重み・係数（黄色が入力欄）　先月データ／先々月データ：貼り付け先　祝日：祝日・特異日リスト　集計：内部計算", F),
        ("", F),
        ("■ 色の意味", FB),
        ("黄色背景・青字＝入力してよいセル　／　それ以外＝自動計算（触らない）", F),
    ]
    for i, (txt, font) in enumerate(lines, start=2):
        cc = u.cell(row=i, column=2, value=txt)
        cc.font = font
        cc.alignment = Alignment(wrap_text=True, vertical="top")

    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = ws.title in (DS1, DS2)
    wb.active = 1
    wb.save(out)


if __name__ == "__main__":
    build(sys.argv[1], *(sys.argv[2:4] + [None, None])[:2])
