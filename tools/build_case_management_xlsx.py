# -*- coding: utf-8 -*-
"""案件管理表（.xlsx）を生成するスクリプト。

生成されるシート:
  使い方         … 凡例・入力ルール・前提条件
  案件管理       … メインの案件一覧（金額 / クォーター / 機器納品 / 現地作業 / 工数 / 進行状況）
  工数明細       … 日次の工数入力。案件管理の「実績工数」に自動集計される
  月次工数計画   … 何月に何人日を投入するかの計画と、その月次実績
  発注・納品     … 機器の発注明細。発注日〜受入日を追いかける
  現地作業       … 現地作業の予定日・実施日・作業者・場所
  ダッシュボード … 年度を指定してクォーター別・ステータス別などに集計
  マスタ         … 年度開始月などの設定値とプルダウンの選択肢

再生成: python tools/build_case_management_xlsx.py
"""

import datetime as dt

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

OUT = "案件管理表.xlsx"
FONT = "Arial"

CASE_FIRST, CASE_LAST = 3, 202     # 案件管理
TIME_FIRST, TIME_LAST = 3, 302     # 工数明細
ORD_FIRST, ORD_LAST = 3, 202       # 発注・納品
SITE_FIRST, SITE_LAST = 3, 202     # 現地作業
MONTHS = 24                        # 月次工数計画の月数

# ---------------------------------------------------------------- 配色・書式
C_INPUT, C_FORMULA, C_LINK = "0000FF", "000000", "008000"
F_HEADER = PatternFill("solid", fgColor="1F3864")
F_HEADER_FX = PatternFill("solid", fgColor="375623")
F_SUBHEAD = PatternFill("solid", fgColor="D9E2F3")
F_INPUT = PatternFill("solid", fgColor="FFF9E6")
F_CALC = PatternFill("solid", fgColor="F2F2F2")
F_KEY = PatternFill("solid", fgColor="FFFF00")
F_SAMPLE = PatternFill("solid", fgColor="EAF3EA")
F_TITLE = PatternFill("solid", fgColor="1F3864")
F_TOTAL = PatternFill("solid", fgColor="E2EFDA")

THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FMT_YEN = '¥#,##0;[Red](¥#,##0);"-"'
FMT_PCT = '0.0%;(0.0%);"-"'
FMT_DAY = '#,##0.0"人日";(#,##0.0);"-"'
FMT_HOUR = '#,##0.0"h";(#,##0.0);"-"'
FMT_DATE = "yyyy/mm/dd"
FMT_MONTH = "yyyy/mm"
FMT_TIME = "hh:mm"
FMT_FY = '0"年度"'
FMT_INT = '#,##0;(#,##0);"-"'
FMT_CNT = '#,##0"件";;"-"'
FMT_DELAY = '#,##0"日 遅延";#,##0"日 前倒し";"予定どおり"'
FMT_MD = '#,##0.0;(#,##0.0);"-"'

wb = Workbook()


def head(cell, text, fill=F_HEADER):
    cell.value = text
    cell.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
    cell.fill = fill
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BOX


def subhead(cell, text):
    head(cell, text, F_SUBHEAD)
    cell.font = Font(name=FONT, size=10, bold=True, color="1F3864")


def title(ws, ref, text):
    ws.merge_cells(ref)
    c = ws[ref.split(":")[0]]
    c.value = text
    c.font = Font(name=FONT, size=14, bold=True, color="FFFFFF")
    c.fill = F_TITLE
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[c.row].height = 26


def label(cell, text, bold=False, size=10, color="000000"):
    cell.value = text
    cell.font = Font(name=FONT, size=size, bold=bold, color=color)
    return cell


def as_date(s):
    return dt.date.fromisoformat(s) if s else None


def build_grid(ws, cols, first, last, samples, sample_key, dv_specs, row_height=18):
    """共通の表組みを作る。cols は (見出し, 幅, 'in'|'fx', 表示形式, メモ, 数式関数)。"""
    letters = {name: get_column_letter(i + 1) for i, (name, *_r) in enumerate(cols)}
    for i, (name, width, kind, fmt, note, _fx) in enumerate(cols, start=1):
        c = ws.cell(2, i)
        head(c, name, F_HEADER if kind == "in" else F_HEADER_FX)
        if note:
            c.comment = Comment(note, "案件管理表")
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[2].height = 32

    for r in range(first, last + 1):
        sample = samples[r - first] if r - first < len(samples) else None
        for i, (name, width, kind, fmt, note, fx) in enumerate(cols, start=1):
            c = ws.cell(r, i)
            c.border = BOX
            c.number_format = fmt or "General"
            if kind == "fx":
                c.value = fx(r)
                c.font = Font(name=FONT, size=10, color=C_FORMULA)
                c.fill = F_CALC
            else:
                c.font = Font(name=FONT, size=10, color=C_INPUT)
                c.fill = F_SAMPLE if sample else F_INPUT
                if sample:
                    c.value = sample_key(sample, name)
            if fmt in (FMT_DATE, FMT_MONTH, FMT_TIME) or fmt in (FMT_PCT, "0%", FMT_CNT):
                c.alignment = Alignment(horizontal="center")
        ws.row_dimensions[r].height = row_height

    for src, colname in dv_specs:
        dv = DataValidation(type="list", formula1=src, allow_blank=True, showDropDown=False)
        dv.error = "マスタシートのリストから選んでください。"
        dv.errorTitle = "入力できない値です"
        ws.add_data_validation(dv)
        dv.add(f"{letters[colname]}{first}:{letters[colname]}{last}")

    ws.auto_filter.ref = f"A2:{get_column_letter(len(cols))}{last}"
    ws.sheet_view.showGridLines = False
    return letters


# ================================================================ マスタ
ms = wb.create_sheet("マスタ")
title(ms, "A1:J1", "マスタ（設定値とプルダウンの選択肢）")
label(ms["A2"], "■ 設定値（黄色いセルを自分の会社に合わせて変更してください）", bold=True)
head(ms["A3"], "設定項目")
head(ms["B3"], "値")
head(ms["C3"], "説明")
settings = [
    ("年度開始月", 4, '0"月"', "日本の一般的な年度（4月開始）。1月開始なら 1 に変更してください。"),
    ("1人日あたりの時間", 8, '0.0"h"', "工数明細の「時間」を「人日」に換算するときの分母です。"),
    ("当年度（ダッシュボード初期値）", 2026, FMT_FY, "ダッシュボードの集計対象年度の初期値。"),
    ("月次工数計画の開始年月", dt.date(2026, 4, 1), FMT_MONTH,
     "月次工数計画シートの一番左の月。ここを変えると24ヶ月ぶんの列見出しがずれます"
     "（入力済みの数値は動かないので注意）。"),
]
for i, (name, val, fmt, note) in enumerate(settings):
    r = 4 + i
    label(ms.cell(r, 1), name).border = BOX
    c = ms.cell(r, 2, val)
    c.font = Font(name=FONT, size=10, bold=True, color=C_INPUT)
    c.fill, c.number_format = F_KEY, fmt
    c.alignment = Alignment(horizontal="center")
    c.border = BOX
    label(ms.cell(r, 3), note, size=9, color="595959").border = BOX

FY_START, HOURS_PER_DAY = "マスタ!$B$4", "マスタ!$B$5"
PLAN_START = "マスタ!$B$7"

lists = {
    "E": ("ステータス", ["未着手", "進行中", "保留", "完了", "失注"]),
    "F": ("作業ステップ", ["要件定義", "基本設計", "詳細設計", "機器手配", "開発・構築",
                          "現地作業", "テスト", "納品・導入", "検収待ち", "完了"]),
    "G": ("担当者", ["山田 太郎", "佐藤 花子", "鈴木 一郎", "（未定）"]),
    "H": ("案件区分", ["新規", "既存拡大", "保守・運用", "社内"]),
    "I": ("受注確度", [1.0, 0.8, 0.5, 0.3, 0.1, 0.0]),
    "J": ("現地作業区分", ["現地調査", "搬入・設置", "配線", "設定・調整", "疎通試験",
                          "立会い", "移行作業", "障害対応", "撤去"]),
    "K": ("作業状態", ["予定", "完了", "延期", "中止"]),
    "L": ("発注先", ["（未定）", "A商事", "Bシステム", "Cネットワークス"]),
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

ms.column_dimensions["A"].width = 28
ms.column_dimensions["B"].width = 14
ms.column_dimensions["C"].width = 60
for col in lists:
    ms.column_dimensions[col].width = 14
ms.sheet_view.showGridLines = False

for n, ref in {
    "StatusList": "マスタ!$E$4:$E$13", "PhaseList": "マスタ!$F$4:$F$18",
    "OwnerList": "マスタ!$G$4:$G$23", "TypeList": "マスタ!$H$4:$H$13",
    "ProbList": "マスタ!$I$4:$I$9", "SiteTypeList": "マスタ!$J$4:$J$18",
    "WorkStateList": "マスタ!$K$4:$K$10", "VendorList": "マスタ!$L$4:$L$23",
}.items():
    wb.defined_names.add(DefinedName(n, attr_text=ref))

# ================================================================ 案件管理（列定義）
CASE_COLS_SPEC = [
    # (見出し, 幅, 種別, 表示形式, メモ)
    ("案件ID", 11, "in", "@", "重複しない番号。他のシートと紐づくキーです（例: P-001）"),
    ("案件名", 24, "in", None, None),
    ("顧客名", 18, "in", None, None),
    ("担当者", 12, "in", None, "マスタのリストから選択"),
    ("案件区分", 11, "in", None, None),
    ("受注確度", 9, "in", "0%", "見込みの確からしさ。加重金額の計算に使います"),
    ("ステータス", 11, "in", None, "未着手 / 進行中 / 保留 / 完了 / 失注"),
    ("現在の作業ステップ", 15, "in", None, "いま案件のどこをやっているか"),
    ("進捗率", 9, "in", FMT_PCT, "0〜100% で入力"),
    ("受注(予定)日", 12, "in", FMT_DATE, "この日付から年度・クォーターを判定します"),
    ("着手日", 12, "in", FMT_DATE, None),
    ("完了予定日", 12, "in", FMT_DATE, None),
    ("完了日", 12, "in", FMT_DATE, None),
    ("年度", 10, "fx", FMT_FY, None),
    ("クォーター", 10, "fx", "@", None),
    ("遅延日数", 12, "fx", FMT_DELAY, None),
    ("受注金額", 13, "in", FMT_YEN, "税抜の受注（見込）金額"),
    ("加重金額", 13, "fx", FMT_YEN, None),
    ("原価(機器以外)", 13, "in", FMT_YEN, "外注費・交通費など。機器代は発注・納品シートから自動集計します"),
    ("機器原価", 13, "fx", FMT_YEN, None),
    ("原価計", 13, "fx", FMT_YEN, None),
    ("粗利", 13, "fx", FMT_YEN, None),
    ("粗利率", 9, "fx", FMT_PCT, None),
    ("発注件数", 10, "fx", FMT_CNT, None),
    ("未入荷件数", 10, "fx", FMT_CNT, None),
    ("最終入荷日", 12, "fx", FMT_DATE, None),
    ("納品状況", 13, "fx", "@", None),
    ("現地作業(予定)", 11, "fx", FMT_CNT, None),
    ("現地作業(完了)", 11, "fx", FMT_CNT, None),
    ("次回現地作業日", 13, "fx", FMT_DATE, None),
    ("予定工数", 11, "in", FMT_DAY, "人日で入力。月ごとの内訳は月次工数計画シートへ"),
    ("月次計画計", 11, "fx", FMT_DAY, None),
    ("実績工数", 11, "fx", FMT_DAY, None),
    ("残工数", 11, "fx", FMT_DAY, None),
    ("工数消化率", 10, "fx", FMT_PCT, None),
    ("予実判定", 10, "fx", "@", None),
    ("備考", 30, "in", None, None),
]
CL = {name: get_column_letter(i + 1) for i, (name, *_r) in enumerate(CASE_COLS_SPEC)}
CASE_LASTCOL = get_column_letter(len(CASE_COLS_SPEC))


def cr(name, first=CASE_FIRST, last=CASE_LAST):
    c = CL[name]
    return f"案件管理!${c}${first}:${c}${last}"


# 他シートの範囲（列位置は後段の定義と一致させること）
OID = f"'発注・納品'!$B${ORD_FIRST}:$B${ORD_LAST}"      # 案件ID
OAMT = f"'発注・納品'!$H${ORD_FIRST}:$H${ORD_LAST}"     # 金額
OARR = f"'発注・納品'!$M${ORD_FIRST}:$M${ORD_LAST}"     # 入荷日
OACC = f"'発注・納品'!$N${ORD_FIRST}:$N${ORD_LAST}"     # 受入日
SID = f"現地作業!$A${SITE_FIRST}:$A${SITE_LAST}"        # 案件ID
SDATE = f"現地作業!$C${SITE_FIRST}:$C${SITE_LAST}"      # 作業日
SSTATE = f"現地作業!$L${SITE_FIRST}:$L${SITE_LAST}"     # 状態
TID = f"工数明細!$B${TIME_FIRST}:$B${TIME_LAST}"        # 案件ID
TDAY = f"工数明細!$G${TIME_FIRST}:$G${TIME_LAST}"       # 工数(人日)
TYM = f"工数明細!$I${TIME_FIRST}:$I${TIME_LAST}"        # 年月キー

PLAN_FIRST = 8                      # 月次工数計画・計画ブロックの先頭行
PLAN_OFFSET = PLAN_FIRST - CASE_FIRST
PLAN_M1 = 7                         # 月列の開始（G列）
PLAN_MLAST_L = get_column_letter(PLAN_M1 + MONTHS - 1)


def case_fx(name, r):
    A, F = f"${CL['案件ID']}{r}", f"${CL['受注確度']}{r}"
    J = f"${CL['受注(予定)日']}{r}"
    LP, MD = f"${CL['完了予定日']}{r}", f"${CL['完了日']}{r}"
    amt = f"${CL['受注金額']}{r}"
    return {
        "年度": f'=IF({J}="","",YEAR({J})-IF(MONTH({J})<{FY_START},1,0))',
        "クォーター": f'=IF({J}="","","Q"&(INT(MOD(MONTH({J})-{FY_START}+12,12)/3)+1))',
        "遅延日数": (f'=IF({LP}="","",IF({MD}<>"",{MD}-{LP},'
                     f'IF(TODAY()>{LP},TODAY()-{LP},0)))'),
        "加重金額": f'=IF({amt}="","",{amt}*N({F}))',
        "機器原価": f'=IF({A}="","",SUMIFS({OAMT},{OID},{A}))',
        "原価計": f'=IF({amt}="","",N(${CL["原価(機器以外)"]}{r})+N(${CL["機器原価"]}{r}))',
        "粗利": f'=IF({amt}="","",{amt}-N(${CL["原価計"]}{r}))',
        "粗利率": f'=IF(N({amt})=0,"",${CL["粗利"]}{r}/{amt})',
        "発注件数": f'=IF({A}="","",COUNTIFS({OID},{A}))',
        "未入荷件数": f'=IF({A}="","",COUNTIFS({OID},{A},{OARR},""))',
        "最終入荷日": (f'=IF(N(${CL["発注件数"]}{r})=0,"",'
                       f'IF(_xlfn.MAXIFS({OARR},{OID},{A})=0,"",_xlfn.MAXIFS({OARR},{OID},{A})))'),
        "納品状況": (f'=IF({A}="","",IF(N(${CL["発注件数"]}{r})=0,"機器なし",'
                     f'IF(N(${CL["未入荷件数"]}{r})>0,"入荷待ち",'
                     f'IF(COUNTIFS({OID},{A},{OACC},"")>0,"受入待ち","受入完了"))))'),
        "現地作業(予定)": f'=IF({A}="","",COUNTIFS({SID},{A},{SSTATE},"予定"))',
        "現地作業(完了)": f'=IF({A}="","",COUNTIFS({SID},{A},{SSTATE},"完了"))',
        "次回現地作業日": (
            f'=IF(N(${CL["現地作業(予定)"]}{r})=0,"",'
            f'IF(_xlfn.MINIFS({SDATE},{SID},{A},{SSTATE},"予定",{SDATE},">="&TODAY())=0,"",'
            f'_xlfn.MINIFS({SDATE},{SID},{A},{SSTATE},"予定",{SDATE},">="&TODAY())))'),
        "月次計画計": (f'=IF({A}="","",SUM(月次工数計画!$G${r + PLAN_OFFSET}:'
                       f'${PLAN_MLAST_L}${r + PLAN_OFFSET}))'),
        "実績工数": f'=IF({A}="","",SUMIFS({TDAY},{TID},{A}))',
        "残工数": (f'=IF(${CL["予定工数"]}{r}="","",'
                   f'${CL["予定工数"]}{r}-N(${CL["実績工数"]}{r}))'),
        "工数消化率": (f'=IF(N(${CL["予定工数"]}{r})=0,"",'
                       f'${CL["実績工数"]}{r}/${CL["予定工数"]}{r})'),
        "予実判定": (f'=IF(N(${CL["予定工数"]}{r})=0,"",'
                     f'IF(${CL["実績工数"]}{r}>${CL["予定工数"]}{r},"超過",'
                     f'IF(${CL["実績工数"]}{r}>=${CL["予定工数"]}{r}*0.9,"注意","適正")))'),
    }[name]


CASE_SAMPLES = [
    {"案件ID": "P-001", "案件名": "本社ネットワーク更改", "顧客名": "株式会社アルファ商事",
     "担当者": "山田 太郎", "案件区分": "新規", "受注確度": 1.0, "ステータス": "進行中",
     "現在の作業ステップ": "現地作業", "進捗率": 0.45, "受注(予定)日": "2026-05-20",
     "着手日": "2026-06-01", "完了予定日": "2026-08-31", "完了日": None,
     "受注金額": 4800000, "原価(機器以外)": 600000, "予定工数": 45,
     "備考": "※サンプル行です。書き換えるか、行ごと削除してお使いください。"},
    {"案件ID": "P-002", "案件名": "基幹システム保守（年間）", "顧客名": "ベータ工業株式会社",
     "担当者": "佐藤 花子", "案件区分": "保守・運用", "受注確度": 1.0, "ステータス": "完了",
     "現在の作業ステップ": "完了", "進捗率": 1.0, "受注(予定)日": "2026-04-10",
     "着手日": "2026-04-15", "完了予定日": "2026-06-30", "完了日": "2026-06-28",
     "受注金額": 1200000, "原価(機器以外)": 480000, "予定工数": 12,
     "備考": "※サンプル行です。"},
    {"案件ID": "P-003", "案件名": "Webサイトリニューアル", "顧客名": "ガンマデザイン合同会社",
     "担当者": "鈴木 一郎", "案件区分": "新規", "受注確度": 0.5, "ステータス": "未着手",
     "現在の作業ステップ": "要件定義", "進捗率": 0.1, "受注(予定)日": "2026-11-05",
     "着手日": None, "完了予定日": "2027-02-28", "完了日": None,
     "受注金額": 3000000, "原価(機器以外)": 1500000, "予定工数": 30,
     "備考": "※サンプル行です。提案中のため確度50%。"},
]
DATE_COLS = {"受注(予定)日", "着手日", "完了予定日", "完了日"}

cs = wb.create_sheet("案件管理")
title(cs, f"A1:{CASE_LASTCOL}1",
      "案件管理表　│　金額・クォーター・機器納品・現地作業・工数・進行状況をまとめて管理")
case_cols = [(n, w, k, f, note, (lambda r, nm=n: case_fx(nm, r)))
             for (n, w, k, f, note) in CASE_COLS_SPEC]
build_grid(
    cs, case_cols, CASE_FIRST, CASE_LAST, CASE_SAMPLES,
    lambda s, n: as_date(s.get(n)) if n in DATE_COLS else s.get(n),
    [("=StatusList", "ステータス"), ("=PhaseList", "現在の作業ステップ"),
     ("=OwnerList", "担当者"), ("=TypeList", "案件区分"), ("=ProbList", "受注確度")],
)
for name in ("実績工数", "機器原価", "月次計画計", "発注件数", "未入荷件数", "最終入荷日",
             "納品状況", "現地作業(予定)", "現地作業(完了)", "次回現地作業日"):
    for r in range(CASE_FIRST, CASE_LAST + 1):
        cs[f"{CL[name]}{r}"].font = Font(name=FONT, size=10, color=C_LINK)

dv_pct = DataValidation(type="decimal", operator="between", formula1="0", formula2="1",
                        allow_blank=True)
dv_pct.error, dv_pct.errorTitle = "0%〜100% の範囲で入力してください。", "進捗率の入力"
cs.add_data_validation(dv_pct)
dv_pct.add(f"{CL['進捗率']}{CASE_FIRST}:{CL['進捗率']}{CASE_LAST}")


def ccf(name):
    c = CL[name]
    return f"{c}{CASE_FIRST}:{c}{CASE_LAST}"


cs.conditional_formatting.add(ccf("進捗率"), DataBarRule(
    start_type="num", start_value=0, end_type="num", end_value=1, color="63BE7B"))
for col, pairs in [
    ("ステータス", [("完了", "C6EFCE", "006100"), ("進行中", "BDD7EE", "1F3864"),
                    ("保留", "FFEB9C", "9C5700"), ("失注", "D9D9D9", "808080")]),
    ("予実判定", [("超過", "FFC7CE", "9C0006"), ("注意", "FFEB9C", "9C5700"),
                  ("適正", "C6EFCE", "006100")]),
    ("納品状況", [("入荷待ち", "FFEB9C", "9C5700"), ("受入待ち", "BDD7EE", "1F3864"),
                  ("受入完了", "C6EFCE", "006100"), ("機器なし", "F2F2F2", "808080")]),
]:
    for text, bg, fg in pairs:
        cs.conditional_formatting.add(ccf(col), CellIsRule(
            operator="equal", formula=[f'"{text}"'],
            fill=PatternFill("solid", fgColor=bg), font=Font(name=FONT, size=10, color=fg)))
cs.conditional_formatting.add(ccf("遅延日数"), CellIsRule(
    operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor="FFC7CE"),
    font=Font(name=FONT, size=10, bold=True, color="9C0006")))
cs.conditional_formatting.add(ccf("未入荷件数"), CellIsRule(
    operator="greaterThan", formula=["0"], font=Font(name=FONT, size=10, bold=True, color="9C5700")))
cs.conditional_formatting.add(
    f"A{CASE_FIRST}:{CASE_LASTCOL}{CASE_LAST}",
    FormulaRule(formula=[f'${CL["ステータス"]}{CASE_FIRST}="失注"'],
                font=Font(name=FONT, size=10, color="A6A6A6", strike=True)))
cs.freeze_panes = "C3"

lg = CASE_LAST + 2
label(cs.cell(lg, 1), "凡例：", bold=True)
label(cs.cell(lg, 2), "青文字＝入力する欄", color=C_INPUT)
label(cs.cell(lg, 5), "黒文字（グレー背景）＝このシート内で自動計算", color=C_FORMULA)
label(cs.cell(lg, 12), "緑文字＝他のシートから自動集計", color=C_LINK)
label(cs.cell(lg + 1, 1),
      "行が足りなくなったら、最終行を選んで下にドラッグ（オートフィル）すると数式ごと増やせます。"
      "※月次工数計画シートの行も同じだけ増やしてください。", size=9, color="595959")

# ================================================================ 工数明細
ts = wb.create_sheet("工数明細")
title(ts, "A1:I1",
      "工数明細　│　日々の工数を入れると、案件管理の「実績工数」と月次工数計画の「実績」に積み上がります")
TIME_COLS = [
    ("日付", 12, "in", FMT_DATE, None, None),
    ("案件ID", 12, "in", "@", "案件管理シートに登録済みのIDを選んでください", None),
    ("案件名", 24, "fx", None, None,
     lambda r: (f'=IF($B{r}="","",IFERROR(INDEX({cr("案件名")},'
                f'MATCH($B{r},{cr("案件ID")},0)),"←IDが未登録です"))')),
    ("担当者", 13, "in", None, None, None),
    ("作業ステップ", 15, "in", None, None, None),
    ("工数(時間)", 11, "in", FMT_HOUR, "その日にかけた時間を入力", None),
    ("工数(人日)", 11, "fx", FMT_DAY, None,
     lambda r: f'=IF(N($F{r})=0,"",$F{r}/{HOURS_PER_DAY})'),
    ("作業内容・メモ", 38, "in", None, None, None),
    ("年月", 10, "fx", "@", "月次工数計画シートの集計キーです（例: 202606）", None),
]
TIME_COLS[8] = ("年月", 10, "fx", "0", "月次工数計画シートの集計キーです（例: 202606）",
                lambda r: f'=IF($A{r}="","",YEAR($A{r})*100+MONTH($A{r}))')
TIME_SAMPLES = [
    {"日付": "2026-06-03", "案件ID": "P-001", "担当者": "山田 太郎", "作業ステップ": "現地作業",
     "工数(時間)": 8, "作業内容・メモ": "L3スイッチ設置・疎通確認（※サンプル行）"},
    {"日付": "2026-06-04", "案件ID": "P-001", "担当者": "山田 太郎", "作業ステップ": "現地作業",
     "工数(時間)": 6, "作業内容・メモ": "配線作業（※サンプル行）"},
    {"日付": "2026-07-01", "案件ID": "P-001", "担当者": "佐藤 花子", "作業ステップ": "テスト",
     "工数(時間)": 4, "作業内容・メモ": "受入試験の立会い（※サンプル行）"},
    {"日付": "2026-06-05", "案件ID": "P-002", "担当者": "佐藤 花子", "作業ステップ": "テスト",
     "工数(時間)": 4, "作業内容・メモ": "定期点検レポート作成（※サンプル行）"},
]
tl = build_grid(ts, TIME_COLS, TIME_FIRST, TIME_LAST, TIME_SAMPLES,
                lambda s, n: as_date(s.get(n)) if n == "日付" else s.get(n),
                [("=OwnerList", "担当者"), ("=PhaseList", "作業ステップ")])
for r in range(TIME_FIRST, TIME_LAST + 1):
    ts[f"C{r}"].font = Font(name=FONT, size=10, color=C_LINK)
dv_id = DataValidation(type="list", formula1=f"={cr('案件ID')}", allow_blank=True,
                       showDropDown=False)
dv_id.error, dv_id.errorTitle = "案件管理シートに登録済みの案件IDを選んでください。", "案件IDの入力"
ts.add_data_validation(dv_id)
dv_id.add(f"B{TIME_FIRST}:B{TIME_LAST}")
ts.conditional_formatting.add(f"C{TIME_FIRST}:C{TIME_LAST}", CellIsRule(
    operator="equal", formula=['"←IDが未登録です"'],
    fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FONT, size=10, bold=True,
                                                           color="9C0006")))
ts.freeze_panes = "A3"
label(ts.cell(TIME_LAST + 2, 1),
      "「工数(人日)」＝ 工数(時間) ÷ マスタ!B5（1人日あたりの時間）。"
      "「年月」は月次集計用の自動計算です。", size=9, color="595959")

# ================================================================ 発注・納品
os_ = wb.create_sheet("発注・納品")
title(os_, "A1:Q1", "発注・納品　│　機器の発注から受入までを1明細1行で管理")
ORD_COLS = [
    ("発注No", 11, "in", "@", "自由な採番でOK（例: O-001）", None),
    ("案件ID", 11, "in", "@", "案件管理シートに登録済みのIDを選んでください", None),
    ("案件名", 22, "fx", None, None,
     lambda r: (f'=IF($B{r}="","",IFERROR(INDEX({cr("案件名")},'
                f'MATCH($B{r},{cr("案件ID")},0)),"←IDが未登録です"))')),
    ("品名・型番", 26, "in", None, None, None),
    ("メーカー", 14, "in", None, None, None),
    ("数量", 8, "in", FMT_INT, None, None),
    ("仕入単価", 13, "in", FMT_YEN, "1台あたりの仕入価格（税抜）", None),
    ("金額", 13, "fx", FMT_YEN, None,
     lambda r: f'=IF(N($F{r})*N($G{r})=0,"",$F{r}*$G{r})'),
    ("発注先", 14, "in", None, None, None),
    ("発注日", 12, "in", FMT_DATE, None, None),
    ("希望納期", 12, "in", FMT_DATE, "こちらが希望する納品日", None),
    ("回答納期", 12, "in", FMT_DATE, "発注先から回答された納期。納期差異はここを基準に計算します", None),
    ("入荷日", 12, "in", FMT_DATE, "実際にモノが届いた日", None),
    ("受入日", 12, "in", FMT_DATE, "検品して受け入れた日", None),
    ("納品状況", 15, "fx", "@", None, None),
    ("納期差異", 12, "fx", FMT_DELAY, None, None),
    ("備考", 28, "in", None, None, None),
]
ORD_COLS[14] = ("納品状況", 15, "fx", "@", None, lambda r: (
    f'=IF($B{r}="","",IF($N{r}<>"","受入済",IF($M{r}<>"","入荷済（受入待ち）",'
    f'IF($J{r}<>"","発注済","未発注"))))'))
ORD_COLS[15] = ("納期差異", 12, "fx", FMT_DELAY, "回答納期に対する入荷日のズレ", lambda r: (
    f'=IF($L{r}="","",IF($M{r}<>"",$M{r}-$L{r},IF(TODAY()>$L{r},TODAY()-$L{r},0)))'))
ORD_SAMPLES = [
    {"発注No": "O-001", "案件ID": "P-001", "品名・型番": "L3スイッチ XG-4800-24P",
     "メーカー": "アライドテレシス", "数量": 4, "仕入単価": 320000, "発注先": "A商事",
     "発注日": "2026-05-25", "希望納期": "2026-06-15", "回答納期": "2026-06-18",
     "入荷日": "2026-06-18", "受入日": "2026-06-19", "備考": "※サンプル行です。"},
    {"発注No": "O-002", "案件ID": "P-001", "品名・型番": "光メディアコンバータ MC-200",
     "メーカー": "アライドテレシス", "数量": 8, "仕入単価": 45000, "発注先": "A商事",
     "発注日": "2026-05-25", "希望納期": "2026-06-15", "回答納期": "2026-07-10",
     "入荷日": None, "受入日": None, "備考": "※サンプル行です。納期回答が遅く入荷待ち。"},
    {"発注No": "O-003", "案件ID": "P-001", "品名・型番": "19インチラック 24U",
     "メーカー": "Cネットワークス", "数量": 1, "仕入単価": 180000, "発注先": "Cネットワークス",
     "発注日": "2026-05-28", "希望納期": "2026-06-20", "回答納期": "2026-06-20",
     "入荷日": "2026-06-19", "受入日": None, "備考": "※サンプル行です。入荷済み・受入待ち。"},
]
ORD_DATES = {"発注日", "希望納期", "回答納期", "入荷日", "受入日"}
build_grid(os_, ORD_COLS, ORD_FIRST, ORD_LAST, ORD_SAMPLES,
           lambda s, n: as_date(s.get(n)) if n in ORD_DATES else s.get(n),
           [("=VendorList", "発注先")])
for r in range(ORD_FIRST, ORD_LAST + 1):
    os_[f"C{r}"].font = Font(name=FONT, size=10, color=C_LINK)
dv_oid = DataValidation(type="list", formula1=f"={cr('案件ID')}", allow_blank=True,
                        showDropDown=False)
dv_oid.error, dv_oid.errorTitle = "案件管理シートに登録済みの案件IDを選んでください。", "案件IDの入力"
os_.add_data_validation(dv_oid)
dv_oid.add(f"B{ORD_FIRST}:B{ORD_LAST}")
for text, bg, fg in [("受入済", "C6EFCE", "006100"), ("入荷済（受入待ち）", "BDD7EE", "1F3864"),
                     ("発注済", "FFEB9C", "9C5700"), ("未発注", "FFC7CE", "9C0006")]:
    os_.conditional_formatting.add(f"O{ORD_FIRST}:O{ORD_LAST}", CellIsRule(
        operator="equal", formula=[f'"{text}"'],
        fill=PatternFill("solid", fgColor=bg), font=Font(name=FONT, size=10, color=fg)))
os_.conditional_formatting.add(f"P{ORD_FIRST}:P{ORD_LAST}", CellIsRule(
    operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor="FFC7CE"),
    font=Font(name=FONT, size=10, bold=True, color="9C0006")))
os_.conditional_formatting.add(f"C{ORD_FIRST}:C{ORD_LAST}", CellIsRule(
    operator="equal", formula=['"←IDが未登録です"'],
    fill=PatternFill("solid", fgColor="FFC7CE"),
    font=Font(name=FONT, size=10, bold=True, color="9C0006")))
os_.freeze_panes = "D3"
label(os_.cell(ORD_LAST + 2, 1),
      "金額（数量×仕入単価）は案件管理の「機器原価」に自動集計されます。"
      "納期差異は回答納期に対するズレで、未入荷なら今日までの経過日数を表示します。",
      size=9, color="595959")

# ================================================================ 現地作業
ss = wb.create_sheet("現地作業")
title(ss, "A1:N1", "現地作業　│　現地に行く日の予定と実施記録（工数そのものは工数明細に入れてください）")
SITE_COLS = [
    ("案件ID", 11, "in", "@", "案件管理シートに登録済みのIDを選んでください", None),
    ("案件名", 22, "fx", None, None,
     lambda r: (f'=IF($A{r}="","",IFERROR(INDEX({cr("案件名")},'
                f'MATCH($A{r},{cr("案件ID")},0)),"←IDが未登録です"))')),
    ("作業日", 12, "in", FMT_DATE, None, None),
    ("曜日", 7, "fx", "@", None,
     lambda r: f'=IF($C{r}="","",CHOOSE(WEEKDAY($C{r},1),"日","月","火","水","木","金","土"))'),
    ("開始", 8, "in", FMT_TIME, "例: 9:00", None),
    ("終了", 8, "in", FMT_TIME, "例: 18:00", None),
    ("休憩(h)", 9, "in", FMT_HOUR, None, None),
    ("実働時間", 10, "fx", FMT_HOUR, None,
     lambda r: f'=IF(OR($E{r}="",$F{r}=""),"",($F{r}-$E{r})*24-N($G{r}))'),
    ("作業場所", 22, "in", None, "現地の住所や拠点名", None),
    ("作業者", 16, "in", None, "複数名なら「山田/佐藤」のように書いてもOK", None),
    ("作業区分", 13, "in", None, None, None),
    ("状態", 9, "in", "@", "予定 / 完了 / 延期 / 中止", None),
    ("作業内容", 30, "in", None, None, None),
    ("備考", 24, "in", None, None, None),
]
SITE_SAMPLES = [
    {"案件ID": "P-001", "作業日": "2026-06-03", "開始": dt.time(9, 0), "終了": dt.time(18, 0),
     "休憩(h)": 1, "作業場所": "アルファ商事 本社 3F サーバ室", "作業者": "山田 太郎",
     "作業区分": "搬入・設置", "状態": "完了", "作業内容": "L3スイッチ設置・疎通確認",
     "備考": "※サンプル行です。"},
    {"案件ID": "P-001", "作業日": "2026-06-04", "開始": dt.time(9, 0), "終了": dt.time(16, 0),
     "休憩(h)": 1, "作業場所": "アルファ商事 本社 各フロア", "作業者": "山田 太郎",
     "作業区分": "配線", "状態": "完了", "作業内容": "フロア配線の敷設",
     "備考": "※サンプル行です。"},
    {"案件ID": "P-001", "作業日": "2026-10-14", "開始": dt.time(10, 0), "終了": dt.time(17, 0),
     "休憩(h)": 1, "作業場所": "アルファ商事 本社 3F サーバ室", "作業者": "山田/佐藤",
     "作業区分": "移行作業", "状態": "予定", "作業内容": "本番切替（夜間立会いあり）",
     "備考": "※サンプル行です。これから行く予定の作業。"},
]
build_grid(ss, SITE_COLS, SITE_FIRST, SITE_LAST, SITE_SAMPLES,
           lambda s, n: as_date(s.get(n)) if n == "作業日" else s.get(n),
           [("=SiteTypeList", "作業区分"), ("=WorkStateList", "状態"), ("=OwnerList", "作業者")])
for r in range(SITE_FIRST, SITE_LAST + 1):
    ss[f"B{r}"].font = Font(name=FONT, size=10, color=C_LINK)
dv_sid = DataValidation(type="list", formula1=f"={cr('案件ID')}", allow_blank=True,
                        showDropDown=False)
dv_sid.error, dv_sid.errorTitle = "案件管理シートに登録済みの案件IDを選んでください。", "案件IDの入力"
ss.add_data_validation(dv_sid)
dv_sid.add(f"A{SITE_FIRST}:A{SITE_LAST}")
for text, bg, fg in [("完了", "C6EFCE", "006100"), ("予定", "BDD7EE", "1F3864"),
                     ("延期", "FFEB9C", "9C5700"), ("中止", "D9D9D9", "808080")]:
    ss.conditional_formatting.add(f"L{SITE_FIRST}:L{SITE_LAST}", CellIsRule(
        operator="equal", formula=[f'"{text}"'],
        fill=PatternFill("solid", fgColor=bg), font=Font(name=FONT, size=10, color=fg)))
# 土日の作業日を色づけ
ss.conditional_formatting.add(f"C{SITE_FIRST}:D{SITE_LAST}", FormulaRule(
    formula=[f'AND($C{SITE_FIRST}<>"",OR(WEEKDAY($C{SITE_FIRST},1)=1,'
             f'WEEKDAY($C{SITE_FIRST},1)=7))'],
    fill=PatternFill("solid", fgColor="FCE4EC")))
ss.conditional_formatting.add(f"B{SITE_FIRST}:B{SITE_LAST}", CellIsRule(
    operator="equal", formula=['"←IDが未登録です"'],
    fill=PatternFill("solid", fgColor="FFC7CE"),
    font=Font(name=FONT, size=10, bold=True, color="9C0006")))
ss.freeze_panes = "C3"
label(ss.cell(SITE_LAST + 2, 1),
      "状態が「予定」で今日以降の作業日のうち、いちばん近い日が案件管理の「次回現地作業日」に出ます。"
      "土日の行はピンク色になります。", size=9, color="595959")

# ================================================================ 月次工数計画
mp = wb.create_sheet("月次工数計画")
LASTM = get_column_letter(PLAN_M1 + MONTHS - 1)
title(mp, f"A1:{LASTM}1", "月次工数計画　│　どの案件に、何月、何人日を投入するか")
label(mp["A3"], "開始年月", bold=True)
c = mp["B3"]
c.value = f"={PLAN_START}"
c.font = Font(name=FONT, size=11, bold=True, color=C_LINK)
c.fill, c.number_format, c.border = F_KEY, FMT_MONTH, BOX
c.alignment = Alignment(horizontal="center")
label(mp["C3"], "← マスタ!B7 で変更できます。変えると列見出しがずれるので、"
                "入力済みの数値も合わせて動かしてください。", size=9, color="595959")

FIX_COLS = [("案件ID", 11), ("案件名", 22), ("担当者", 12),
            ("予定工数", 11), ("月次計画計", 11), ("差異", 10)]

# 計画ブロックの記入例（案件管理の先頭3行に対応）。合計が各案件の予定工数と一致するようにしてある。
PLAN_SAMPLES = [
    {"2026-06": 15, "2026-07": 20, "2026-08": 10},                  # P-001 合計45人日
    {"2026-04": 4, "2026-05": 4, "2026-06": 4},                     # P-002 合計12人日
    {"2026-11": 5, "2026-12": 10, "2027-01": 10, "2027-02": 5},     # P-003 合計30人日
]
PLAN_START_DATE = settings[3][1]


def month_offset(ym):
    """'2026-06' が開始年月から何ヶ月目かを返す（範囲外なら None）。"""
    y, m = (int(x) for x in ym.split("-"))
    n = (y - PLAN_START_DATE.year) * 12 + (m - PLAN_START_DATE.month)
    return n if 0 <= n < MONTHS else None


def month_block(start_row, block_title, kind):
    """計画/実績ブロックを1つ作る。start_row はブロック見出しの行。"""
    label(mp.cell(start_row, 1), block_title, bold=True, size=11, color="1F3864")
    key_row, hdr_row = start_row + 1, start_row + 2
    first = start_row + 3
    last = first + (CASE_LAST - CASE_FIRST)
    headers = {"案件ID": "案件ID", "案件名": "案件名", "担当者": "担当者",
               "予定工数": "予定工数"}
    headers["月次計画計"] = "月次 計画計" if kind == "plan" else "月次 実績計"
    headers["差異"] = "差異\n(未配分)" if kind == "plan" else "差異\n(計画−実績)"
    for i, (name, width) in enumerate(FIX_COLS, start=1):
        subhead(mp.cell(hdr_row, i), headers[name])
        mp.column_dimensions[get_column_letter(i)].width = width
    for m in range(MONTHS):
        col = PLAN_M1 + m
        cl = get_column_letter(col)
        mp.column_dimensions[cl].width = 9
        h = mp.cell(hdr_row, col)
        h.value = f"=EDATE({PLAN_START},{m})"
        h.number_format = FMT_MONTH
        h.font = Font(name=FONT, size=9, bold=True, color="1F3864")
        h.fill = F_SUBHEAD
        h.alignment = Alignment(horizontal="center")
        h.border = BOX
        k = mp.cell(key_row, col)
        k.value = f"=YEAR({cl}{hdr_row})*100+MONTH({cl}{hdr_row})"
        k.number_format = "0"
        k.font = Font(name=FONT, size=8, color="BFBFBF")
        k.alignment = Alignment(horizontal="center")
    label(mp.cell(key_row, 1), "（集計キー：編集しないでください）", size=8, color="BFBFBF")
    mp.row_dimensions[key_row].height = 12
    mp.row_dimensions[hdr_row].height = 20

    for r in range(first, last + 1):
        case_r = CASE_FIRST + (r - first)
        plan_r = PLAN_FIRST + (r - first)
        vals = {
            "案件ID": f'=IF({cr("案件ID", case_r, case_r)}="","",{cr("案件ID", case_r, case_r)})',
            "案件名": f'=IF($A{r}="","",{cr("案件名", case_r, case_r)})',
            "担当者": f'=IF($A{r}="","",{cr("担当者", case_r, case_r)})',
            "予定工数": f'=IF($A{r}="","",{cr("予定工数", case_r, case_r)})',
            "月次計画計": (f'=IF($A{r}="","",SUM($G{plan_r}:${LASTM}{plan_r}))' if kind == "plan"
                           else f'=IF($A{r}="","",SUM($G{r}:${LASTM}{r}))'),
            "差異": (f'=IF($A{r}="","",N($D{r})-N($E{r}))' if kind == "plan"
                     else f'=IF($A{r}="","",N($E{plan_r})-N($E{r}))'),
        }
        fmts = {"案件ID": "@", "案件名": None, "担当者": None,
                "予定工数": FMT_DAY, "月次計画計": FMT_DAY, "差異": FMT_MD}
        for i, (name, _w) in enumerate(FIX_COLS, start=1):
            cell = mp.cell(r, i)
            cell.value = vals[name]
            cell.number_format = fmts[name] or "General"
            cell.font = Font(name=FONT, size=10, color=C_LINK)
            cell.fill = F_CALC
            cell.border = BOX
        for m in range(MONTHS):
            col = PLAN_M1 + m
            cell = mp.cell(r, col)
            cell.number_format = FMT_MD
            cell.border = BOX
            if kind == "plan":
                cell.font = Font(name=FONT, size=10, color=C_INPUT)
                idx = r - first
                sample = PLAN_SAMPLES[idx] if idx < len(PLAN_SAMPLES) else None
                cell.fill = F_SAMPLE if sample else F_INPUT
                if sample:
                    for ym, v in sample.items():
                        if month_offset(ym) == m:
                            cell.value = v
            else:
                kcl = get_column_letter(col)
                cell.value = (f'=IF($A{r}="","",SUMIFS({TDAY},{TID},$A{r},'
                              f'{TYM},{kcl}${key_row}))')
                cell.font = Font(name=FONT, size=10, color=C_LINK)
                cell.fill = F_CALC
        mp.row_dimensions[r].height = 16

    tot = last + 1
    label(mp.cell(tot, 1), "月別 合計", bold=True)
    mp.cell(tot, 1).fill = F_TOTAL
    mp.cell(tot, 1).border = BOX
    for i in range(2, 7):
        cell = mp.cell(tot, i)
        cell.fill, cell.border = F_TOTAL, BOX
        if i >= 4:
            cell.value = f"=SUM({get_column_letter(i)}{first}:{get_column_letter(i)}{last})"
            cell.number_format = FMT_DAY if i < 6 else FMT_MD
            cell.font = Font(name=FONT, size=10, bold=True)
    for m in range(MONTHS):
        cl = get_column_letter(PLAN_M1 + m)
        cell = mp.cell(tot, PLAN_M1 + m)
        cell.value = f"=SUM({cl}{first}:{cl}{last})"
        cell.number_format = FMT_MD
        cell.font = Font(name=FONT, size=10, bold=True)
        cell.fill, cell.border = F_TOTAL, BOX
    return first, last, tot, key_row, hdr_row


plan_first, plan_last, plan_tot, plan_key, plan_hdr = month_block(
    5, "■ 計画（人日）　… 青いセルに「何月に何人日」を入力してください", "plan")
act_start = plan_tot + 3
act_first, act_last, act_tot, act_key, act_hdr = month_block(
    act_start, "■ 実績（人日）　… 工数明細から自動集計されます（入力不要）", "actual")

mp.freeze_panes = f"G{plan_hdr + 1}"
mp.sheet_view.showGridLines = False
mp.conditional_formatting.add(f"F{plan_first}:F{plan_last}", CellIsRule(
    operator="notEqual", formula=["0"], fill=PatternFill("solid", fgColor="FFEB9C"),
    font=Font(name=FONT, size=10, color="9C5700")))
mp.conditional_formatting.add(f"G{plan_first}:{LASTM}{plan_last}", DataBarRule(
    start_type="min", end_type="max", color="9BC2E6"))
mp.conditional_formatting.add(f"G{act_first}:{LASTM}{act_last}", DataBarRule(
    start_type="min", end_type="max", color="A9D08E"))
label(mp.cell(act_tot + 2, 1),
      "「差異」は 計画ブロック＝予定工数−月次計画計（月への配分し忘れを検出）、"
      "実績ブロック＝月次計画計−月次実績計（計画に対する進み遅れ）です。"
      "計画ブロックの差異が0でない行は黄色くなります。",
      size=9, color="595959")

# ================================================================ ダッシュボード
ds = wb.create_sheet("ダッシュボード")
title(ds, "B1:I1", "ダッシュボード　│　年度を選ぶと自動で集計されます")
ds.column_dimensions["A"].width = 2
ds.column_dimensions["B"].width = 24
for col in "CDEFGHI":
    ds.column_dimensions[col].width = 15
ds.sheet_view.showGridLines = False

label(ds["B3"], "集計対象の年度", bold=True)
c = ds["C3"]
c.value = "=マスタ!$B$6"
c.font = Font(name=FONT, size=11, bold=True, color=C_LINK)
c.fill, c.number_format, c.border = F_KEY, FMT_FY, BOX
c.alignment = Alignment(horizontal="center")
label(ds["D3"], "← マスタ!B6 の値を見ています。ここに直接 2027 などと入力して切り替えてもOKです。",
      size=9, color="595959")

FYR, QR = cr("年度"), cr("クォーター")


def table(start, heading, headers, rows_spec, widths_fmt):
    label(ds.cell(start, 2), heading, bold=True, size=11, color="1F3864")
    hdr = start + 1
    for i, h in enumerate(headers):
        subhead(ds.cell(hdr, 2 + i), h)
    for j, (rowlabel, cells) in enumerate(rows_spec):
        r = hdr + 1 + j
        label(ds.cell(r, 2), rowlabel).border = BOX
        for i, formula in enumerate(cells):
            cell = ds.cell(r, 3 + i)
            cell.value = formula(r) if callable(formula) else formula.format(r=r)
            cell.number_format = widths_fmt[i]
            cell.font = Font(name=FONT, size=10)
            cell.border = BOX
    return hdr + len(rows_spec) + 3


# 表1: クォーター別
r0, hdr = 5, 6
label(ds.cell(r0, 2), "■ クォーター別サマリ", bold=True, size=11)
subhead(ds.cell(hdr, 2), "項目")
for i, q in enumerate(["Q1", "Q2", "Q3", "Q4"]):
    subhead(ds.cell(hdr, 3 + i), q)
subhead(ds.cell(hdr, 7), "年度合計")

q_rows = [("案件数", None, FMT_CNT, "count"), ("受注金額", "受注金額", FMT_YEN, "sum"),
          ("加重金額（確度反映）", "加重金額", FMT_YEN, "sum"),
          ("機器原価", "機器原価", FMT_YEN, "sum"), ("原価計", "原価計", FMT_YEN, "sum"),
          ("粗利", "粗利", FMT_YEN, "sum"), ("粗利率", None, FMT_PCT, "ratio_gp"),
          ("予定工数", "予定工数", FMT_DAY, "sum"), ("実績工数", "実績工数", FMT_DAY, "sum"),
          ("工数消化率", None, FMT_PCT, "ratio_ws")]
for j, (name, col, fmt, kind) in enumerate(q_rows):
    r = hdr + 1 + j
    label(ds.cell(r, 2), name).border = BOX
    ds.cell(r, 2).fill = F_CALC
    for i in range(4):
        cc = ds.cell(r, 3 + i)
        L_ = get_column_letter(3 + i)
        qref = f"{L_}{hdr}"
        if kind == "count":
            cc.value = f'=COUNTIFS({FYR},$C$3,{QR},{qref},{cr("案件ID")},"<>")'
        elif kind == "sum":
            cc.value = f'=SUMIFS({cr(col)},{FYR},$C$3,{QR},{qref})'
        elif kind == "ratio_gp":
            cc.value = f'=IF(N({L_}{hdr + 2})=0,"",{L_}{hdr + 6}/{L_}{hdr + 2})'
        else:
            cc.value = f'=IF(N({L_}{hdr + 8})=0,"",{L_}{hdr + 9}/{L_}{hdr + 8})'
        cc.number_format, cc.border = fmt, BOX
        cc.font = Font(name=FONT, size=10)
    tc = ds.cell(r, 7)
    if kind == "ratio_gp":
        tc.value = f'=IF(N(G{hdr + 2})=0,"",G{hdr + 6}/G{hdr + 2})'
    elif kind == "ratio_ws":
        tc.value = f'=IF(N(G{hdr + 8})=0,"",G{hdr + 9}/G{hdr + 8})'
    else:
        tc.value = f"=SUM(C{r}:F{r})"
    tc.number_format, tc.border, tc.fill = fmt, BOX, F_TOTAL
    tc.font = Font(name=FONT, size=10, bold=True)

nxt = hdr + len(q_rows) + 3

# 表2: ステータス別
nxt = table(nxt, "■ ステータス別（対象年度）",
            ["ステータス", "案件数", "受注金額", "粗利", "予定工数", "実績工数"],
            [(st, [f'=COUNTIFS({FYR},$C$3,{cr("ステータス")},$B{{r}})',
                   f'=SUMIFS({cr("受注金額")},{FYR},$C$3,{cr("ステータス")},$B{{r}})',
                   f'=SUMIFS({cr("粗利")},{FYR},$C$3,{cr("ステータス")},$B{{r}})',
                   f'=SUMIFS({cr("予定工数")},{FYR},$C$3,{cr("ステータス")},$B{{r}})',
                   f'=SUMIFS({cr("実績工数")},{FYR},$C$3,{cr("ステータス")},$B{{r}})'])
             for st in lists["E"][1]],
            [FMT_CNT, FMT_YEN, FMT_YEN, FMT_DAY, FMT_DAY])

# 表3: 担当者別
nxt = table(nxt, "■ 担当者別（対象年度）",
            ["担当者", "案件数", "受注金額", "粗利", "予定工数", "実績工数", "工数消化率"],
            [(ow, [f'=COUNTIFS({FYR},$C$3,{cr("担当者")},$B{{r}})',
                   f'=SUMIFS({cr("受注金額")},{FYR},$C$3,{cr("担当者")},$B{{r}})',
                   f'=SUMIFS({cr("粗利")},{FYR},$C$3,{cr("担当者")},$B{{r}})',
                   f'=SUMIFS({cr("予定工数")},{FYR},$C$3,{cr("担当者")},$B{{r}})',
                   f'=SUMIFS({cr("実績工数")},{FYR},$C$3,{cr("担当者")},$B{{r}})',
                   "=IF(N($F{r})=0,\"\",$G{r}/$F{r})"])
             for ow in lists["G"][1]],
            [FMT_CNT, FMT_YEN, FMT_YEN, FMT_DAY, FMT_DAY, FMT_PCT])

# 表4: 作業ステップ別
nxt = table(nxt, "■ いまどの作業ステップにいるか（対象年度／失注をのぞく）",
            ["作業ステップ", "案件数", "受注金額", "実績工数"],
            [(ph, [f'=COUNTIFS({FYR},$C$3,{cr("現在の作業ステップ")},$B{{r}},'
                   f'{cr("ステータス")},"<>失注")',
                   f'=SUMIFS({cr("受注金額")},{FYR},$C$3,{cr("現在の作業ステップ")},$B{{r}},'
                   f'{cr("ステータス")},"<>失注")',
                   f'=SUMIFS({cr("実績工数")},{FYR},$C$3,{cr("現在の作業ステップ")},$B{{r}},'
                   f'{cr("ステータス")},"<>失注")'])
             for ph in lists["F"][1]],
            [FMT_CNT, FMT_YEN, FMT_DAY])

# 表5: 機器の発注・納品（案件の年度ではなく発注明細そのものを集計）
label(ds.cell(nxt, 2), "■ 機器の発注・納品（全期間）", bold=True, size=11)
hdr5 = nxt + 1
for i, h in enumerate(["項目", "件数", "金額"]):
    subhead(ds.cell(hdr5, 2 + i), h)
ostat = f"'発注・納品'!$O${ORD_FIRST}:$O${ORD_LAST}"
for j, st in enumerate(["未発注", "発注済", "入荷済（受入待ち）", "受入済"]):
    r = hdr5 + 1 + j
    label(ds.cell(r, 2), st).border = BOX
    ds.cell(r, 3).value = f'=COUNTIFS({ostat},$B{r})'
    ds.cell(r, 4).value = f'=SUMIFS({OAMT},{ostat},$B{r})'
    for cix, fmt in [(3, FMT_CNT), (4, FMT_YEN)]:
        ds.cell(r, cix).number_format = fmt
        ds.cell(r, cix).font = Font(name=FONT, size=10)
        ds.cell(r, cix).border = BOX
r = hdr5 + 5
label(ds.cell(r, 2), "納期遅れ（回答納期超過）", bold=True).border = BOX
ds.cell(r, 3).value = f"=COUNTIF('発注・納品'!$P${ORD_FIRST}:$P${ORD_LAST},\">0\")"
ds.cell(r, 3).number_format = FMT_CNT
ds.cell(r, 3).font = Font(name=FONT, size=10, bold=True, color="9C0006")
ds.cell(r, 3).border = BOX
ds.cell(r, 4).border = BOX
nxt = hdr5 + 7

# 表6: 現地作業
label(ds.cell(nxt, 2), "■ 現地作業（全期間）", bold=True, size=11)
hdr6 = nxt + 1
for i, h in enumerate(["項目", "件数", "実働時間"]):
    subhead(ds.cell(hdr6, 2 + i), h)
SHOURS = f"現地作業!$H${SITE_FIRST}:$H${SITE_LAST}"
for j, st in enumerate(lists["K"][1]):
    r = hdr6 + 1 + j
    label(ds.cell(r, 2), st).border = BOX
    ds.cell(r, 3).value = f'=COUNTIFS({SSTATE},$B{r})'
    ds.cell(r, 4).value = f'=SUMIFS({SHOURS},{SSTATE},$B{r})'
    for cix, fmt in [(3, FMT_CNT), (4, FMT_HOUR)]:
        ds.cell(r, cix).number_format = fmt
        ds.cell(r, cix).font = Font(name=FONT, size=10)
        ds.cell(r, cix).border = BOX
r = hdr6 + len(lists["K"][1]) + 1
label(ds.cell(r, 2), "今後30日以内の現地作業", bold=True).border = BOX
ds.cell(r, 3).value = (f'=COUNTIFS({SSTATE},"予定",{SDATE},">="&TODAY(),'
                       f'{SDATE},"<="&TODAY()+30)')
ds.cell(r, 3).number_format = FMT_CNT
ds.cell(r, 3).font = Font(name=FONT, size=10, bold=True, color="1F3864")
ds.cell(r, 3).border = BOX
ds.cell(r, 4).border = BOX

label(ds.cell(r + 2, 2),
      "※ クォーター別・ステータス別などの集計は「受注(予定)日」から判定した年度が基準です"
      "（年度開始月は マスタ!B4）。機器と現地作業の表は全期間の集計です。",
      size=9, color="595959")

# ================================================================ 使い方
us = wb.create_sheet("使い方")
title(us, "B1:C1", "使い方とルール")
us.column_dimensions["A"].width = 2
us.column_dimensions["B"].width = 22
us.column_dimensions["C"].width = 104
us.sheet_view.showGridLines = False

guide = [
    ("", ""),
    ("■ シートの役割", ""),
    ("案件管理", "案件を1件1行で登録します。青文字＋うすい黄色の列だけ入力してください。"),
    ("", "グレー背景の列は自動計算です。緑文字は他のシートから集計した値です。"),
    ("工数明細", "日々の工数を1行ずつ入れます。案件管理の「実績工数」と月次工数計画の「実績」に反映されます。"),
    ("月次工数計画", "どの案件に何月何人日を投入するかを入力します。下半分に工数明細からの実績が並びます。"),
    ("発注・納品", "機器の発注を1明細1行で。発注日→希望納期→回答納期→入荷日→受入日 を追いかけます。"),
    ("現地作業", "現地に行く日の予定と実施記録。作業日・時間・場所・作業者・状態を入れます。"),
    ("ダッシュボード", "C3に年度を入れると、クォーター別・ステータス別・担当者別・作業ステップ別、"),
    ("", "および機器の納品状況・現地作業の件数を自動集計します。"),
    ("マスタ", "年度開始月（初期値4月）、1人日あたりの時間（初期値8h）、月次計画の開始年月、"),
    ("", "各プルダウンの選択肢を変えられます。"),
    ("", ""),
    ("■ 色のルール", ""),
    ("青文字", "自分で入力するセル"),
    ("黒文字＋グレー背景", "そのシート内の数式による自動計算。編集しないでください。"),
    ("緑文字＋グレー背景", "他のシートから集計している自動計算。編集しないでください。"),
    ("黄色の塗り", "設定値。会社のルールに合わせて変更してください（マスタシート）。"),
    ("うすい緑の行", "サンプル行です。書き換えるか、行ごと削除してからお使いください。"),
    ("", ""),
    ("■ 入力の順番（おすすめ）", ""),
    ("1", "マスタで年度開始月・1人日あたりの時間・月次計画の開始年月を決める"),
    ("2", "案件管理に案件を登録する（案件IDは重複しない値にしてください）"),
    ("3", "月次工数計画で、その案件に何月何人日を投入するか配分する"),
    ("4", "機器があれば発注・納品に明細を登録する（金額は案件管理の機器原価に自動反映）"),
    ("5", "現地作業に行く日を予定として登録しておく"),
    ("6", "作業したら工数明細に工数を、現地作業の状態を「完了」に更新する"),
    ("", ""),
    ("■ 自動でやってくれること", ""),
    ("クォーター判定", "「受注(予定)日」から年度とQ1〜Q4を自動判定します（年度開始月はマスタで変更可）。"),
    ("金額まわり", "機器原価＝発注・納品の金額合計、原価計＝原価(機器以外)＋機器原価、"),
    ("", "粗利＝受注金額−原価計、加重金額＝受注金額×受注確度。"),
    ("機器の納品状況", "発注件数・未入荷件数・最終入荷日・納品状況（機器なし/入荷待ち/受入待ち/受入完了）。"),
    ("", "回答納期を過ぎても入荷していない明細は、納期差異が赤く表示されます。"),
    ("現地作業", "予定件数・完了件数と、今日以降でいちばん近い「次回現地作業日」を表示します。"),
    ("工数まわり", "実績工数は工数明細から自動集計。残工数・消化率・予実判定（適正/注意/超過）も自動です。"),
    ("", "月次工数計画では、予定工数に対して月配分が足りているか「差異」で確認できます。"),
    ("進行状況", "ステータスは色分け、進捗率はデータバー、失注行は打ち消し線になります。"),
    ("遅延の検知", "完了予定日を過ぎて未完了の案件は「遅延日数」が赤く表示されます。"),
    ("", ""),
    ("■ 行が足りなくなったら", ""),
    ("オートフィル", "案件管理・発注・納品・現地作業は202行目、工数明細は302行目まで数式が入っています。"),
    ("", "最終行を選択して右下を下へドラッグすると、数式・書式・プルダウンごと増やせます。"),
    ("注意", "案件管理の行を増やしたら、月次工数計画の計画ブロック・実績ブロックの行も"),
    ("", "同じだけ増やしてください（行の位置が対応しています）。"),
    ("", ""),
    ("■ 前提としている値", ""),
    ("年度開始月", "4月（日本の一般的な年度）。マスタ!B4 で変更できます。"),
    ("1人日", "8時間。マスタ!B5 で変更できます。"),
    ("月次計画の開始年月", "2026年4月から24ヶ月ぶん。マスタ!B7 で変更できます。"),
    ("金額", "すべて税抜・円で入力する前提です。"),
    ("工数の入り口", "工数は工数明細だけに入れてください。現地作業シートの実働時間は参考値で、"),
    ("", "実績工数には加算されません（二重計上を避けるためです）。"),
    ("サンプルの数値", "各シートのサンプル行の数値はすべて説明用の架空データです。"),
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

for rr, (txt, kw, fill) in enumerate([
        ("入力するセルの見た目", dict(color=C_INPUT), F_INPUT),
        ("自動計算セルの見た目", dict(color=C_FORMULA), F_CALC),
        ("他シート集計セルの見た目", dict(color=C_LINK), F_CALC),
        ("設定値セルの見た目", dict(color=C_INPUT, bold=True), F_KEY),
        ("サンプル行の見た目", dict(color=C_INPUT), F_SAMPLE)]):
    cell = us.cell(r + 1 + rr, 2)
    cell.value = txt
    cell.font = Font(name=FONT, size=10, **kw)
    cell.fill, cell.border = fill, BOX

del wb["Sheet"]
wb._sheets = [wb[n] for n in ["使い方", "案件管理", "工数明細", "月次工数計画",
                              "発注・納品", "現地作業", "ダッシュボード", "マスタ"]]
wb.active = 1
wb.save(OUT)
print("wrote", OUT)
