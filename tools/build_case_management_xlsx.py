# -*- coding: utf-8 -*-
"""案件管理表（.xlsx）を生成するスクリプト。

生成されるシート:
  使い方        … 凡例・入力ルール
  案件管理      … メインの案件一覧（金額 / クォーター / 工数 / 進行状況 / 作業ステップ）
  工数明細      … 日次の工数入力。案件管理の「実績工数」に自動集計される
  ダッシュボード … 年度を指定してクォーター別・ステータス別・担当者別に集計
  マスタ        … 年度開始月などの設定値とプルダウンの選択肢

再生成: python tools/build_case_management_xlsx.py
"""

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

OUT = "案件管理表.xlsx"

FONT = "Arial"

# 案件管理シートの行範囲
CASE_FIRST = 3
CASE_LAST = 202          # 200行ぶん
# 工数明細シートの行範囲
TIME_FIRST = 3
TIME_LAST = 302          # 300行ぶん

# ---------------------------------------------------------------- 配色・書式
C_INPUT = "0000FF"        # 直接入力するセル（青文字）
C_FORMULA = "000000"      # 数式（黒文字）
C_LINK = "008000"         # 他シート参照（緑文字）
F_HEADER = PatternFill("solid", fgColor="1F3864")
F_SUBHEAD = PatternFill("solid", fgColor="D9E2F3")
F_INPUT = PatternFill("solid", fgColor="FFF9E6")   # 入力してほしい列
F_CALC = PatternFill("solid", fgColor="F2F2F2")    # 自動計算列
F_KEY = PatternFill("solid", fgColor="FFFF00")     # 重要な設定値
F_SAMPLE = PatternFill("solid", fgColor="EAF3EA")  # サンプル行
F_TITLE = PatternFill("solid", fgColor="1F3864")

THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FMT_YEN = '¥#,##0;[Red](¥#,##0);"-"'
FMT_PCT = '0.0%;(0.0%);"-"'
FMT_DAY = '#,##0.0"人日";(#,##0.0);"-"'
FMT_HOUR = '#,##0.0"h";(#,##0.0);"-"'
FMT_DATE = "yyyy/mm/dd"
FMT_FY = '0"年度"'
FMT_INT = '#,##0;(#,##0);"-"'
FMT_DELAY = '#,##0"日 遅延";#,##0"日 前倒し";"予定どおり"'


def head(cell, text, width_fill=F_HEADER):
    cell.value = text
    cell.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
    cell.fill = width_fill
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BOX


def title(ws, ref, text):
    ws.merge_cells(ref)
    c = ws[ref.split(":")[0]]
    c.value = text
    c.font = Font(name=FONT, size=14, bold=True, color="FFFFFF")
    c.fill = F_TITLE
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)


def label(cell, text, bold=False, size=10, color="000000"):
    cell.value = text
    cell.font = Font(name=FONT, size=size, bold=bold, color=color)


wb = Workbook()

# ================================================================ マスタ
ms = wb.create_sheet("マスタ")
title(ms, "A1:H1", "マスタ（設定値とプルダウンの選択肢）")
ms.row_dimensions[1].height = 24

label(ms["A2"], "■ 設定値（黄色いセルを自分の会社に合わせて変更してください）", bold=True)
settings = [
    ("年度開始月", 4, "0\"月\"", "日本の一般的な年度（4月開始）です。1月開始なら 1 に変更。"),
    ("1人日あたりの時間", 8, '0.0"h"', "工数明細の「時間」を「人日」に換算するときの分母です。"),
    ("当年度（ダッシュボード初期値）", 2026, FMT_FY, "ダッシュボードの集計対象年度の初期値。"),
]
head(ms["A3"], "設定項目")
head(ms["B3"], "値")
head(ms["C3"], "説明")
for i, (name, val, fmt, note) in enumerate(settings):
    r = 4 + i
    label(ms.cell(r, 1), name)
    c = ms.cell(r, 2, val)
    c.font = Font(name=FONT, size=10, bold=True, color=C_INPUT)
    c.fill = F_KEY
    c.number_format = fmt
    c.alignment = Alignment(horizontal="center")
    c.border = BOX
    ms.cell(r, 1).border = BOX
    label(ms.cell(r, 3), note, size=9, color="595959")
    ms.cell(r, 3).border = BOX

FY_START = "マスタ!$B$4"
HOURS_PER_DAY = "マスタ!$B$5"

# --- 選択肢リスト（E列以降）
lists = {
    "E": ("ステータス", ["未着手", "進行中", "保留", "完了", "失注"]),
    "F": ("作業ステップ", ["要件定義", "基本設計", "詳細設計", "開発・構築", "テスト",
                          "納品・導入", "検収待ち", "完了"]),
    "G": ("担当者", ["山田 太郎", "佐藤 花子", "鈴木 一郎", "（未定）"]),
    "H": ("案件区分", ["新規", "既存拡大", "保守・運用", "社内"]),
    "I": ("受注確度", [1.0, 0.8, 0.5, 0.3, 0.1, 0.0]),
}
label(ms["E2"], "■ プルダウンの選択肢（行を足せば選択肢が増えます）", bold=True)
for col, (name, values) in lists.items():
    head(ms[f"{col}3"], name)
    for i, v in enumerate(values):
        c = ms[f"{col}{4 + i}"]
        c.value = v
        c.font = Font(name=FONT, size=10)
        c.border = BOX
        c.alignment = Alignment(horizontal="center")
        if col == "I":
            c.number_format = "0%"

ms.column_dimensions["A"].width = 26
ms.column_dimensions["B"].width = 12
ms.column_dimensions["C"].width = 46
for col in lists:
    ms.column_dimensions[col].width = 14
ms.sheet_view.showGridLines = False

# 定義名（データの入力規則で使用）
named = {
    "StatusList": "マスタ!$E$4:$E$13",
    "PhaseList": "マスタ!$F$4:$F$16",
    "OwnerList": "マスタ!$G$4:$G$23",
    "TypeList": "マスタ!$H$4:$H$13",
    "ProbList": "マスタ!$I$4:$I$9",
}
for n, ref in named.items():
    wb.defined_names.add(DefinedName(n, attr_text=ref))

# ================================================================ 案件管理
cs = wb.create_sheet("案件管理", 0)

# (見出し, 幅, 種別, 表示形式, メモ)
COLS = [
    ("案件ID",            12, "in",   "@",       "重複しない番号。工数明細と紐づくキーです（例: P-001）"),
    ("案件名",            26, "in",   None,      None),
    ("顧客名",            18, "in",   None,      None),
    ("担当者",            12, "in",   None,      "マスタのリストから選択"),
    ("案件区分",          11, "in",   None,      None),
    ("受注確度",          10, "in",   "0%",      "見込み案件の確からしさ。加重金額の計算に使います"),
    ("ステータス",        11, "in",   None,      "未着手 / 進行中 / 保留 / 完了 / 失注"),
    ("現在の作業ステップ", 15, "in",  None,      "いま案件のどこをやっているか"),
    ("進捗率",            9,  "in",   FMT_PCT,   "0〜100% で入力"),
    ("受注(予定)日",      12, "in",   FMT_DATE,  "この日付から年度・クォーターを自動判定します"),
    ("着手日",            12, "in",   FMT_DATE,  None),
    ("完了予定日",        12, "in",   FMT_DATE,  None),
    ("完了日",            12, "in",   FMT_DATE,  None),
    ("年度",              10, "fx",   FMT_FY,    None),
    ("クォーター",        10, "fx",   "@",       None),
    ("受注金額",          14, "in",   FMT_YEN,   "税抜の受注（見込）金額"),
    ("加重金額",          14, "fx",   FMT_YEN,   None),
    ("原価",              14, "in",   FMT_YEN,   "外注費・仕入など"),
    ("粗利",              14, "fx",   FMT_YEN,   None),
    ("粗利率",            9,  "fx",   FMT_PCT,   None),
    ("予定工数",          11, "in",   FMT_DAY,   "人日で入力"),
    ("実績工数",          11, "fx",   FMT_DAY,   None),
    ("残工数",            11, "fx",   FMT_DAY,   None),
    ("工数消化率",        11, "fx",   FMT_PCT,   None),
    ("予実判定",          10, "fx",   "@",       None),
    ("遅延日数",          10, "fx",   FMT_DELAY, None),
    ("備考",              30, "in",   None,      None),
]
IDX = {name: i + 1 for i, (name, *_rest) in enumerate(COLS)}
L = {name: get_column_letter(i) for name, i in IDX.items()}
LAST_COL = get_column_letter(len(COLS))


def crange(name, first=CASE_FIRST, last=CASE_LAST, abs_=True):
    c = L[name]
    return f"案件管理!${c}${first}:${c}${last}" if abs_ else f"${c}${first}:${c}${last}"


title(cs, f"A1:{LAST_COL}1", "案件管理表　│　金額・クォーター・工数・進行状況をまとめて管理")
cs.row_dimensions[1].height = 26
cs.row_dimensions[2].height = 34

for i, (name, width, kind, fmt, note) in enumerate(COLS, start=1):
    c = cs.cell(2, i)
    head(c, name, F_HEADER if kind == "in" else PatternFill("solid", fgColor="375623"))
    if note:
        c.comment = Comment(note, "案件管理表")
    cs.column_dimensions[get_column_letter(i)].width = width

# --- 数式（各データ行に展開）
def formulas(r):
    a, f, j, k, l, m = L["案件ID"], L["受注確度"], L["受注(予定)日"], L["着手日"], L["完了予定日"], L["完了日"]
    amt, cost = L["受注金額"], L["原価"]
    gp, plan, act = L["粗利"], L["予定工数"], L["実績工数"]
    return {
        "年度": f'=IF(${j}{r}="","",YEAR(${j}{r})-IF(MONTH(${j}{r})<{FY_START},1,0))',
        "クォーター": f'=IF(${j}{r}="","","Q"&(INT(MOD(MONTH(${j}{r})-{FY_START}+12,12)/3)+1))',
        "加重金額": f'=IF(${amt}{r}="","",${amt}{r}*N(${f}{r}))',
        "粗利": f'=IF(${amt}{r}="","",${amt}{r}-N(${cost}{r}))',
        "粗利率": f'=IF(N(${amt}{r})=0,"",${gp}{r}/${amt}{r})',
        "実績工数": (f'=IF(${a}{r}="","",SUMIFS(工数明細!$G${TIME_FIRST}:$G${TIME_LAST},'
                     f'工数明細!$B${TIME_FIRST}:$B${TIME_LAST},${a}{r}))'),
        "残工数": f'=IF(${plan}{r}="","",${plan}{r}-N(${act}{r}))',
        "工数消化率": f'=IF(N(${plan}{r})=0,"",${act}{r}/${plan}{r})',
        "予実判定": (f'=IF(N(${plan}{r})=0,"",IF(${act}{r}>${plan}{r},"超過",'
                     f'IF(${act}{r}>=${plan}{r}*0.9,"注意","適正")))'),
        "遅延日数": (f'=IF(${l}{r}="","",IF(${m}{r}<>"",${m}{r}-${l}{r},'
                     f'IF(TODAY()>${l}{r},TODAY()-${l}{r},0)))'),
    }


SAMPLES = [
    {
        "案件ID": "P-001", "案件名": "社内ネットワーク更改", "顧客名": "株式会社アルファ商事",
        "担当者": "山田 太郎", "案件区分": "新規", "受注確度": 1.0, "ステータス": "進行中",
        "現在の作業ステップ": "開発・構築", "進捗率": 0.45, "受注(予定)日": "2026-05-20",
        "着手日": "2026-06-01", "完了予定日": "2026-08-31", "完了日": None,
        "受注金額": 4800000, "原価": 2600000, "予定工数": 45,
        "備考": "※サンプル行です。中身を書き換えるか、行ごと削除してお使いください。",
    },
    {
        "案件ID": "P-002", "案件名": "基幹システム保守（年間）", "顧客名": "ベータ工業株式会社",
        "担当者": "佐藤 花子", "案件区分": "保守・運用", "受注確度": 1.0, "ステータス": "完了",
        "現在の作業ステップ": "完了", "進捗率": 1.0, "受注(予定)日": "2026-04-10",
        "着手日": "2026-04-15", "完了予定日": "2026-06-30", "完了日": "2026-06-28",
        "受注金額": 1200000, "原価": 480000, "予定工数": 12,
        "備考": "※サンプル行です。",
    },
    {
        "案件ID": "P-003", "案件名": "Webサイトリニューアル", "顧客名": "ガンマデザイン合同会社",
        "担当者": "鈴木 一郎", "案件区分": "新規", "受注確度": 0.5, "ステータス": "未着手",
        "現在の作業ステップ": "要件定義", "進捗率": 0.1, "受注(予定)日": "2026-11-05",
        "着手日": None, "完了予定日": "2027-02-28", "完了日": None,
        "受注金額": 3000000, "原価": 1500000, "予定工数": 30,
        "備考": "※サンプル行です。提案中のため確度50%。",
    },
]

import datetime as _dt


def as_date(s):
    return _dt.date.fromisoformat(s) if s else None


for r in range(CASE_FIRST, CASE_LAST + 1):
    fx = formulas(r)
    sample = SAMPLES[r - CASE_FIRST] if r - CASE_FIRST < len(SAMPLES) else None
    for i, (name, width, kind, fmt, note) in enumerate(COLS, start=1):
        c = cs.cell(r, i)
        c.border = BOX
        c.number_format = fmt or "General"
        if kind == "fx":
            c.value = fx[name]
            c.font = Font(name=FONT, size=10, color=C_LINK if name == "実績工数" else C_FORMULA)
            c.fill = F_CALC
            c.alignment = Alignment(horizontal="center" if name in
                                    ("年度", "クォーター", "予実判定", "遅延日数") else "right")
        else:
            c.font = Font(name=FONT, size=10, color=C_INPUT)
            c.fill = F_SAMPLE if sample else F_INPUT
            if sample:
                v = sample.get(name)
                if name in ("受注(予定)日", "着手日", "完了予定日", "完了日"):
                    v = as_date(v)
                c.value = v
            if fmt == FMT_DATE or name in ("受注確度", "進捗率"):
                c.alignment = Alignment(horizontal="center")
            elif name == "備考":
                c.alignment = Alignment(horizontal="left")
    cs.row_dimensions[r].height = 18

# --- 入力規則（プルダウン）
dvs = [
    ("=StatusList", "ステータス"),
    ("=PhaseList", "現在の作業ステップ"),
    ("=OwnerList", "担当者"),
    ("=TypeList", "案件区分"),
    ("=ProbList", "受注確度"),
]
for src, colname in dvs:
    dv = DataValidation(type="list", formula1=src, allow_blank=True, showDropDown=False)
    dv.error = "マスタシートのリストから選んでください。"
    dv.errorTitle = "入力できない値です"
    cs.add_data_validation(dv)
    dv.add(f"{L[colname]}{CASE_FIRST}:{L[colname]}{CASE_LAST}")

dv_pct = DataValidation(type="decimal", operator="between", formula1="0", formula2="1",
                        allow_blank=True)
dv_pct.error = "0%〜100% の範囲で入力してください。"
dv_pct.errorTitle = "進捗率の入力"
cs.add_data_validation(dv_pct)
dv_pct.add(f"{L['進捗率']}{CASE_FIRST}:{L['進捗率']}{CASE_LAST}")

# --- 条件付き書式
def colrange(name):
    c = L[name]
    return f"{c}{CASE_FIRST}:{c}{CASE_LAST}"


cs.conditional_formatting.add(
    colrange("進捗率"),
    DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="63BE7B"),
)
status_colors = [("完了", "C6EFCE", "006100"), ("進行中", "BDD7EE", "1F3864"),
                 ("保留", "FFEB9C", "9C5700"), ("失注", "D9D9D9", "808080")]
for text, bg, fg in status_colors:
    cs.conditional_formatting.add(
        colrange("ステータス"),
        CellIsRule(operator="equal", formula=[f'"{text}"'],
                   fill=PatternFill("solid", fgColor=bg), font=Font(name=FONT, size=10, color=fg)),
    )
for text, bg, fg in [("超過", "FFC7CE", "9C0006"), ("注意", "FFEB9C", "9C5700"),
                     ("適正", "C6EFCE", "006100")]:
    cs.conditional_formatting.add(
        colrange("予実判定"),
        CellIsRule(operator="equal", formula=[f'"{text}"'],
                   fill=PatternFill("solid", fgColor=bg), font=Font(name=FONT, size=10, color=fg)),
    )
cs.conditional_formatting.add(
    colrange("遅延日数"),
    CellIsRule(operator="greaterThan", formula=["0"],
               fill=PatternFill("solid", fgColor="FFC7CE"),
               font=Font(name=FONT, size=10, bold=True, color="9C0006")),
)
# 失注した案件の行はグレーの打ち消し線にする
cs.conditional_formatting.add(
    f"A{CASE_FIRST}:{LAST_COL}{CASE_LAST}",
    FormulaRule(formula=[f'${L["ステータス"]}{CASE_FIRST}="失注"'],
                font=Font(name=FONT, size=10, color="A6A6A6", strike=True)),
)

cs.auto_filter.ref = f"A2:{LAST_COL}{CASE_LAST}"
cs.freeze_panes = "C3"
cs.sheet_view.showGridLines = False

# 凡例（表の下）
lg = CASE_LAST + 2
label(cs.cell(lg, 1), "凡例：", bold=True)
label(cs.cell(lg, 2), "青文字＝入力する欄", color=C_INPUT)
label(cs.cell(lg, 4), "黒文字（グレー背景）＝自動計算。編集しないでください", color=C_FORMULA)
label(cs.cell(lg, 8), "緑文字＝工数明細シートから自動集計", color=C_LINK)
label(cs.cell(lg + 1, 1),
      "行が足りなくなったら、最終行を選んで下にドラッグ（オートフィル）すると数式ごと増やせます。",
      size=9, color="595959")

# ================================================================ 工数明細
ts = wb.create_sheet("工数明細")
TCOLS = [("日付", 12, "in", FMT_DATE, None),
         ("案件ID", 12, "in", "@", "案件管理シートに登録済みのIDを入力/選択してください"),
         ("案件名", 26, "fx", None, None),
         ("担当者", 14, "in", None, None),
         ("作業ステップ", 15, "in", None, None),
         ("工数(時間)", 12, "in", FMT_HOUR, "その日にかけた時間を入力"),
         ("工数(人日)", 12, "fx", FMT_DAY, None),
         ("作業内容・メモ", 40, "in", None, None)]
TL = {n: get_column_letter(i + 1) for i, (n, *_r) in enumerate(TCOLS)}
title(ts, "A1:H1", "工数明細　│　ここに日々の工数を入れると案件管理シートの「実績工数」に自動で積み上がります")
ts.row_dimensions[1].height = 26
ts.row_dimensions[2].height = 30
for i, (name, width, kind, fmt, note) in enumerate(TCOLS, start=1):
    c = ts.cell(2, i)
    head(c, name, F_HEADER if kind == "in" else PatternFill("solid", fgColor="375623"))
    if note:
        c.comment = Comment(note, "案件管理表")
    ts.column_dimensions[get_column_letter(i)].width = width

TIME_SAMPLES = [
    ("2026-06-03", "P-001", "山田 太郎", "開発・構築", 8, "L3スイッチ設定・疎通確認（※サンプル行）"),
    ("2026-06-04", "P-001", "山田 太郎", "開発・構築", 6, "配線作業（※サンプル行）"),
    ("2026-06-05", "P-002", "佐藤 花子", "テスト", 4, "定期点検レポート作成（※サンプル行）"),
]
for r in range(TIME_FIRST, TIME_LAST + 1):
    s = TIME_SAMPLES[r - TIME_FIRST] if r - TIME_FIRST < len(TIME_SAMPLES) else None
    for i, (name, width, kind, fmt, note) in enumerate(TCOLS, start=1):
        c = ts.cell(r, i)
        c.border = BOX
        c.number_format = fmt or "General"
        if name == "案件名":
            c.value = (f'=IF($B{r}="","",IFERROR(INDEX({crange("案件名")},'
                       f'MATCH($B{r},{crange("案件ID")},0)),"←IDが未登録です"))')
            c.font = Font(name=FONT, size=10, color=C_LINK)
            c.fill = F_CALC
        elif name == "工数(人日)":
            c.value = f'=IF(N($F{r})=0,"",$F{r}/{HOURS_PER_DAY})'
            c.font = Font(name=FONT, size=10, color=C_FORMULA)
            c.fill = F_CALC
            c.alignment = Alignment(horizontal="right")
        else:
            c.font = Font(name=FONT, size=10, color=C_INPUT)
            c.fill = F_SAMPLE if s else F_INPUT
            if s:
                vals = {"日付": as_date(s[0]), "案件ID": s[1], "担当者": s[2],
                        "作業ステップ": s[3], "工数(時間)": s[4], "作業内容・メモ": s[5]}
                c.value = vals.get(name)
            if name in ("日付", "案件ID", "工数(時間)"):
                c.alignment = Alignment(horizontal="center")

for src, colname in [("=OwnerList", "担当者"), ("=PhaseList", "作業ステップ")]:
    dv = DataValidation(type="list", formula1=src, allow_blank=True, showDropDown=False)
    ts.add_data_validation(dv)
    dv.add(f"{TL[colname]}{TIME_FIRST}:{TL[colname]}{TIME_LAST}")

dv_id = DataValidation(type="list", formula1=f"={crange('案件ID')}", allow_blank=True,
                       showDropDown=False)
dv_id.error = "案件管理シートに登録済みの案件IDを選んでください。"
dv_id.errorTitle = "案件IDの入力"
ts.add_data_validation(dv_id)
dv_id.add(f"B{TIME_FIRST}:B{TIME_LAST}")

ts.conditional_formatting.add(
    f"C{TIME_FIRST}:C{TIME_LAST}",
    CellIsRule(operator="equal", formula=['"←IDが未登録です"'],
               fill=PatternFill("solid", fgColor="FFC7CE"),
               font=Font(name=FONT, size=10, bold=True, color="9C0006")),
)
ts.auto_filter.ref = f"A2:H{TIME_LAST}"
ts.freeze_panes = "A3"
ts.sheet_view.showGridLines = False
label(ts.cell(TIME_LAST + 2, 1), "「工数(人日)」＝ 工数(時間) ÷ マスタ!B5（1人日あたりの時間）",
      size=9, color="595959")

# ================================================================ ダッシュボード
ds = wb.create_sheet("ダッシュボード", 0)
title(ds, "B1:I1", "ダッシュボード　│　年度を選ぶと自動で集計されます")
ds.row_dimensions[1].height = 26
ds.column_dimensions["A"].width = 2
ds.column_dimensions["B"].width = 22
for col in "CDEFGH":
    ds.column_dimensions[col].width = 15
ds.column_dimensions["I"].width = 15
ds.sheet_view.showGridLines = False

label(ds["B3"], "集計対象の年度", bold=True)
c = ds["C3"]
c.value = "=マスタ!$B$6"
c.font = Font(name=FONT, size=11, bold=True, color=C_LINK)
c.fill = F_KEY
c.number_format = FMT_FY
c.alignment = Alignment(horizontal="center")
c.border = BOX
label(ds["D3"], "← マスタ!B6 の値を見ています。ここに直接 2027 などと入力して切り替えてもOKです。",
      size=9, color="595959")

FY_RANGE = crange("年度")
Q_RANGE = crange("クォーター")


def sumifs_q(col, qcell):
    return f'=SUMIFS({crange(col)},{FY_RANGE},$C$3,{Q_RANGE},{qcell})'


def countifs_q(qcell):
    return f'=COUNTIFS({FY_RANGE},$C$3,{Q_RANGE},{qcell},{crange("案件ID")},"<>")'


# --- 表1: クォーター別
r0 = 5
label(ds.cell(r0, 2), "■ クォーター別サマリ", bold=True, size=11)
r0 += 1
head(ds.cell(r0, 2), "項目", F_SUBHEAD)
ds.cell(r0, 2).font = Font(name=FONT, size=10, bold=True, color="1F3864")
for i, q in enumerate(["Q1", "Q2", "Q3", "Q4"]):
    hc = ds.cell(r0, 3 + i)
    head(hc, q, F_SUBHEAD)
    hc.font = Font(name=FONT, size=10, bold=True, color="1F3864")
head(ds.cell(r0, 7), "年度合計", F_SUBHEAD)
ds.cell(r0, 7).font = Font(name=FONT, size=10, bold=True, color="1F3864")

q_rows = [
    ("案件数", None, FMT_INT, "count"),
    ("受注金額", "受注金額", FMT_YEN, "sum"),
    ("加重金額（確度反映）", "加重金額", FMT_YEN, "sum"),
    ("原価", "原価", FMT_YEN, "sum"),
    ("粗利", "粗利", FMT_YEN, "sum"),
    ("粗利率", None, FMT_PCT, "ratio_gp"),
    ("予定工数", "予定工数", FMT_DAY, "sum"),
    ("実績工数", "実績工数", FMT_DAY, "sum"),
    ("工数消化率", None, FMT_PCT, "ratio_ws"),
]
base = r0
for j, (name, col, fmt, kind) in enumerate(q_rows):
    r = base + 1 + j
    label(ds.cell(r, 2), name)
    ds.cell(r, 2).border = BOX
    ds.cell(r, 2).fill = PatternFill("solid", fgColor="F2F2F2")
    for i in range(4):
        cc = ds.cell(r, 3 + i)
        qref = f"{get_column_letter(3 + i)}{base}"
        if kind == "count":
            cc.value = countifs_q(qref)
        elif kind == "sum":
            cc.value = sumifs_q(col, qref)
        elif kind == "ratio_gp":
            amt = f"{get_column_letter(3 + i)}{base + 2}"
            gp = f"{get_column_letter(3 + i)}{base + 5}"
            cc.value = f"=IF(N({amt})=0,\"\",{gp}/{amt})"
        else:
            plan = f"{get_column_letter(3 + i)}{base + 7}"
            act = f"{get_column_letter(3 + i)}{base + 8}"
            cc.value = f"=IF(N({plan})=0,\"\",{act}/{plan})"
        cc.number_format = fmt
        cc.font = Font(name=FONT, size=10)
        cc.border = BOX
    tc = ds.cell(r, 7)
    if kind == "ratio_gp":
        tc.value = f"=IF(N(G{base + 2})=0,\"\",G{base + 5}/G{base + 2})"
    elif kind == "ratio_ws":
        tc.value = f"=IF(N(G{base + 7})=0,\"\",G{base + 8}/G{base + 7})"
    else:
        tc.value = f"=SUM(C{r}:F{r})"
    tc.number_format = fmt
    tc.font = Font(name=FONT, size=10, bold=True)
    tc.fill = PatternFill("solid", fgColor="E2EFDA")
    tc.border = BOX

# --- 表2: ステータス別
r0 = base + len(q_rows) + 3
label(ds.cell(r0, 2), "■ ステータス別（対象年度）", bold=True, size=11)
r0 += 1
for i, h in enumerate(["ステータス", "案件数", "受注金額", "粗利", "予定工数", "実績工数"]):
    hc = ds.cell(r0, 2 + i)
    head(hc, h, F_SUBHEAD)
    hc.font = Font(name=FONT, size=10, bold=True, color="1F3864")
for j, st in enumerate(lists["E"][1]):
    r = r0 + 1 + j
    ds.cell(r, 2, st).font = Font(name=FONT, size=10)
    ds.cell(r, 2).border = BOX
    sr = crange("ステータス")
    ds.cell(r, 3).value = f'=COUNTIFS({FY_RANGE},$C$3,{sr},$B{r})'
    ds.cell(r, 4).value = f'=SUMIFS({crange("受注金額")},{FY_RANGE},$C$3,{sr},$B{r})'
    ds.cell(r, 5).value = f'=SUMIFS({crange("粗利")},{FY_RANGE},$C$3,{sr},$B{r})'
    ds.cell(r, 6).value = f'=SUMIFS({crange("予定工数")},{FY_RANGE},$C$3,{sr},$B{r})'
    ds.cell(r, 7).value = f'=SUMIFS({crange("実績工数")},{FY_RANGE},$C$3,{sr},$B{r})'
    for cix, fmt in [(3, FMT_INT), (4, FMT_YEN), (5, FMT_YEN), (6, FMT_DAY), (7, FMT_DAY)]:
        ds.cell(r, cix).number_format = fmt
        ds.cell(r, cix).font = Font(name=FONT, size=10)
        ds.cell(r, cix).border = BOX

# --- 表3: 担当者別
r0 = r0 + len(lists["E"][1]) + 3
label(ds.cell(r0, 2), "■ 担当者別（対象年度）", bold=True, size=11)
r0 += 1
for i, h in enumerate(["担当者", "案件数", "受注金額", "粗利", "予定工数", "実績工数", "工数消化率"]):
    hc = ds.cell(r0, 2 + i)
    head(hc, h, F_SUBHEAD)
    hc.font = Font(name=FONT, size=10, bold=True, color="1F3864")
for j, ow in enumerate(lists["G"][1]):
    r = r0 + 1 + j
    ds.cell(r, 2, ow).font = Font(name=FONT, size=10)
    ds.cell(r, 2).border = BOX
    orng = crange("担当者")
    ds.cell(r, 3).value = f'=COUNTIFS({FY_RANGE},$C$3,{orng},$B{r})'
    ds.cell(r, 4).value = f'=SUMIFS({crange("受注金額")},{FY_RANGE},$C$3,{orng},$B{r})'
    ds.cell(r, 5).value = f'=SUMIFS({crange("粗利")},{FY_RANGE},$C$3,{orng},$B{r})'
    ds.cell(r, 6).value = f'=SUMIFS({crange("予定工数")},{FY_RANGE},$C$3,{orng},$B{r})'
    ds.cell(r, 7).value = f'=SUMIFS({crange("実績工数")},{FY_RANGE},$C$3,{orng},$B{r})'
    ds.cell(r, 8).value = f'=IF(N($F{r})=0,"",$G{r}/$F{r})'
    for cix, fmt in [(3, FMT_INT), (4, FMT_YEN), (5, FMT_YEN), (6, FMT_DAY), (7, FMT_DAY),
                     (8, FMT_PCT)]:
        ds.cell(r, cix).number_format = fmt
        ds.cell(r, cix).font = Font(name=FONT, size=10)
        ds.cell(r, cix).border = BOX

# --- 表4: 作業ステップ別（いまどこをやっているか）
r0 = r0 + len(lists["G"][1]) + 3
label(ds.cell(r0, 2), "■ いまどの作業ステップにいるか（対象年度／失注をのぞく）", bold=True, size=11)
r0 += 1
for i, h in enumerate(["作業ステップ", "案件数", "受注金額", "実績工数"]):
    hc = ds.cell(r0, 2 + i)
    head(hc, h, F_SUBHEAD)
    hc.font = Font(name=FONT, size=10, bold=True, color="1F3864")
for j, ph in enumerate(lists["F"][1]):
    r = r0 + 1 + j
    ds.cell(r, 2, ph).font = Font(name=FONT, size=10)
    ds.cell(r, 2).border = BOX
    prng = crange("現在の作業ステップ")
    srng = crange("ステータス")
    ds.cell(r, 3).value = f'=COUNTIFS({FY_RANGE},$C$3,{prng},$B{r},{srng},"<>失注")'
    ds.cell(r, 4).value = f'=SUMIFS({crange("受注金額")},{FY_RANGE},$C$3,{prng},$B{r},{srng},"<>失注")'
    ds.cell(r, 5).value = f'=SUMIFS({crange("実績工数")},{FY_RANGE},$C$3,{prng},$B{r},{srng},"<>失注")'
    for cix, fmt in [(3, FMT_INT), (4, FMT_YEN), (5, FMT_DAY)]:
        ds.cell(r, cix).number_format = fmt
        ds.cell(r, cix).font = Font(name=FONT, size=10)
        ds.cell(r, cix).border = BOX

label(ds.cell(r0 + len(lists["F"][1]) + 2, 2),
      "※ 集計は「受注(予定)日」から判定した年度・クォーターを基準にしています（年度開始月は マスタ!B4）。",
      size=9, color="595959")

# ================================================================ 使い方
us = wb.create_sheet("使い方", 0)
title(us, "B1:F1", "使い方とルール")
us.row_dimensions[1].height = 26
us.column_dimensions["A"].width = 2
us.column_dimensions["B"].width = 20
us.column_dimensions["C"].width = 100
us.sheet_view.showGridLines = False

guide = [
    ("", ""),
    ("■ 入力する場所", ""),
    ("案件管理", "案件を1件1行で登録します。青文字＋うすい黄色の列だけ入力してください。"),
    ("", "グレー背景の列（年度・クォーター・粗利・実績工数など）は自動計算なので触らなくてOKです。"),
    ("工数明細", "日々の工数を1行ずつ入れます。案件IDを選ぶと、案件管理の「実績工数」に自動で合算されます。"),
    ("マスタ", "年度開始月（初期値4月）、1人日あたりの時間（初期値8h）、プルダウンの選択肢を変えられます。"),
    ("ダッシュボード", "C3に年度を入れると、クォーター別・ステータス別・担当者別・作業ステップ別に自動集計します。"),
    ("", ""),
    ("■ 色のルール", ""),
    ("青文字", "自分で入力するセル"),
    ("黒文字＋グレー背景", "数式による自動計算。上書きすると壊れるので編集しないでください。"),
    ("緑文字", "ほかのシートを参照している自動計算"),
    ("黄色の塗り", "設定値。会社のルールに合わせて変更してください（マスタシート）。"),
    ("うすい緑の行", "サンプル行です。書き換えるか、行ごと削除してからお使いください。"),
    ("", ""),
    ("■ 自動でやってくれること", ""),
    ("クォーター判定", "「受注(予定)日」から年度とQ1〜Q4を自動判定します（年度開始月は マスタ!B4 で変更可）。"),
    ("金額まわり", "粗利＝受注金額−原価、粗利率＝粗利÷受注金額、加重金額＝受注金額×受注確度。"),
    ("工数まわり", "実績工数は工数明細から自動集計。残工数・消化率・予実判定（適正/注意/超過）も自動です。"),
    ("進行状況", "ステータスは色分け表示、進捗率はデータバー、失注行は打ち消し線になります。"),
    ("遅延の検知", "完了予定日を過ぎて未完了の案件は「遅延日数」が赤く表示されます。"),
    ("", ""),
    ("■ 行が足りなくなったら", ""),
    ("オートフィル", "案件管理は202行目、工数明細は302行目まで数式が入っています。"),
    ("", "最終行を選択して右下を下へドラッグすると、数式・書式・プルダウンごと増やせます。"),
    ("", ""),
    ("■ 前提としている値", ""),
    ("年度開始月", "4月（日本の一般的な年度）。マスタ!B4 で変更できます。"),
    ("1人日", "8時間。マスタ!B5 で変更できます。"),
    ("金額", "すべて税抜・円で入力する前提です。"),
    ("サンプルの数値", "案件管理・工数明細のサンプル行の数値はすべて説明用の架空データです。"),
]
r = 2
for a, b in guide:
    if a.startswith("■"):
        label(us.cell(r, 2), a, bold=True, size=11, color="1F3864")
    elif a:
        label(us.cell(r, 2), a, bold=True)
        label(us.cell(r, 3), b)
    elif b:
        label(us.cell(r, 3), b, color="404040")
    r += 1

# 色見本
for rr, (txt, font_kw, fill) in enumerate([
    ("入力するセルの見た目", dict(color=C_INPUT), F_INPUT),
    ("自動計算セルの見た目", dict(color=C_FORMULA), F_CALC),
    ("他シート参照の見た目", dict(color=C_LINK), F_CALC),
    ("設定値セルの見た目", dict(color=C_INPUT, bold=True), F_KEY),
    ("サンプル行の見た目", dict(color=C_INPUT), F_SAMPLE),
]):
    cell = us.cell(r + 1 + rr, 2)
    cell.value = txt
    cell.font = Font(name=FONT, size=10, **font_kw)
    cell.fill = fill
    cell.border = BOX

del wb["Sheet"]
wb.move_sheet("使い方", offset=0)
wb._sheets = [wb["使い方"], wb["案件管理"], wb["工数明細"], wb["ダッシュボード"], wb["マスタ"]]
wb.active = 1
wb.save(OUT)
print("wrote", OUT)
