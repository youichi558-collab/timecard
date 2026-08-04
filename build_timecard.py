#!/usr/bin/env python3
"""出勤簿ブックの改修スクリプト。

元ブック(出勤簿_2026_original.xlsx)の12シートは、体裁が微妙に揃っていなかった
(4月だけ明細の開始行が2行ずれている、見出しの文言が違う、など)。個別に直すのでは
なく、1月シートだけを土台として体裁を整え、それを11か月ぶん複製することで、
12シートの体裁を完全にそろえる。データ(出勤・退勤・理由など)は複製後にシートごと
入れ直す。

直すのは中身だけ:
  - 祝日リストを2025年から2026年に差し替え、VLOOKUPの参照範囲の取りこぼしを修正
  - 勤務時間から実働時間を自動計算し、集計に反映
  - 時刻はコロン付き(9:30、全角の9：00も可)で入力する。コロンなし(930)入力は
    実機Excelで日付として誤表示される不具合が起きたためサポート対象外にした
"""

import datetime
from copy import copy

from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# 土日・水曜・祝日をこの色で塗る(休日として同じ色でまとめる)
# 条件付き書式のdxfは、通常のセルと違いbgColor側に色を入れないと
# Excel実機では表示されない(fgColorだけだとLibreOfficeは表示するがExcelは無視する)。
HOLIDAY_FILL = PatternFill("solid", fgColor="FFFFC7CE", bgColor="FFFFC7CE")

SRC = "出勤簿_2026_original.xlsx"
DST = "出勤簿_2026.xlsx"
YEAR = 2026

# 元ブックの列構成(A〜K)はそのまま。列の追加はしない。
C_IN, C_OUT = 4, 6                            # D=出勤, F=退勤
C_OT, C_OT_N, C_OT_E, C_LATE = 7, 8, 9, 10    # G〜J=各残業・遅刻早退
C_REASON = 11                                 # K=理由
# N〜T は非表示の作業列。入力値のシリアル値変換、日ごとの実働時間、
# 曜日区分の判定に使う。集計はここを参照する（表には出さない）。
C_CONV = {C_IN: 14, C_OUT: 15, C_OT: 16, C_OT_N: 17, C_OT_E: 18, C_LATE: 19}
C_DAILY_WORK, C_KIND = 20, 21
HIDDEN = range(14, 22)

INPUT_COLS = [C_IN, C_OUT, C_OT, C_OT_N, C_OT_E, C_LATE]
COL_LABEL = {C_IN: "出勤", C_OUT: "退勤", C_OT: "通常残業", C_OT_N: "深夜残業",
             C_OT_E: "早朝残業", C_LATE: "遅刻・早退"}

# 時刻入力欄の表示形式。
#
# コロンなし(930)専用の数字分割書式 "00\":\"00" と、コロン付き(h:mm)を条件分岐
# で共存させようとしたが、実機Excelでは書式コードに日付/時刻トークンが1つでも
# 入っていると、セル全体が「日付/時刻型」として扱われてしまい、930 のような
# 整数値が日付(1902年ごろ)として表示される不具合が起きた。LibreOfficeでは
# 再現せず実機Excelでのみ発生する差異で、確実な回避策が見つからなかった。
#
# コロン付き入力(9:30)を標準にする方針にしたため、Excel純正の時刻書式 hh:mm
# だけを使う。コロンなし(930)入力はデータ検証で弾く(サポート対象外)。
TIME_FMT = "hh:mm"

HOLIDAYS_2026 = [
    ((1, 1), "元日", None),
    ((1, 12), "成人の日", None),
    ((2, 11), "建国記念の日", None),
    ((2, 23), "天皇誕生日", None),
    ((3, 20), "春分の日", None),
    ((4, 29), "昭和の日", None),
    ((5, 3), "憲法記念日", None),
    ((5, 4), "みどりの日", None),
    ((5, 5), "こどもの日", None),
    ((5, 6), "休日", "祝日法第3条第2項による休日"),
    ((7, 20), "海の日", None),
    ((8, 11), "山の日", None),
    ((9, 21), "敬老の日", None),
    ((9, 22), "国民の休日", "祝日法第3条第3項による休日"),
    ((9, 23), "秋分の日", None),
    ((10, 12), "スポーツの日", None),
    ((11, 3), "文化の日", None),
    ((11, 23), "勤労感謝の日", None),
]
HOLIDAY_NAMES = {name for _, name, _ in HOLIDAYS_2026}

# 元ブックで欄に残っていた飾り文字・見出し語。データではないので消す。
JUNK = {"：", "～", "日", "時間", "出勤日数", "普通出勤時間", "残業", "土曜出勤", "合計"}


def to_serial(ref):
    """時刻入力欄をシリアル値(1日=1)に変換する数式を返す。

    9:30 のように半角コロン付きで入れた場合、Excel は時刻値(1未満の小数)として
    持つのでそのまま使う。日本語入力の全角コロン「9：00」は Excel が数値として
    認識できず文字列のまま残るため、SUBSTITUTE で半角コロンに直してから
    TIMEVALUE で読み直す。930 のようなコロンなし入力はデータ検証で弾いている
    ため通常は来ないが、来た場合も HHMM とみなして換算する(念のための保険)。
    """
    return (f'IF(ISNUMBER({ref}),'
            f'IF({ref}<1,{ref},(INT({ref}/100)*60+MOD({ref},100))/1440),'
            f'IFERROR(TIMEVALUE(SUBSTITUTE({ref},"：",":")),""))')


WEEKDAY_JA = ["月", "火", "水", "木", "金", "土", "日"]


def days_in_month(month):
    nxt = datetime.date(YEAR + month // 12, month % 12 + 1, 1)
    return (nxt - datetime.timedelta(days=1)).day


def copy_style(dst, src):
    dst.font = copy(src.font)
    dst.border = copy(src.border)
    dst.fill = copy(src.fill)
    dst.alignment = copy(src.alignment)
    dst.number_format = src.number_format


def find_row(ws, col, text, limit=60):
    for r in range(1, limit):
        if ws.cell(r, col).value == text:
            return r
    raise ValueError(f"{ws.title}: 「{text}」の行が見つかりません")


def find_anchor(ws):
    """明細1日目の行を探す。4月だけ開始行が違うため決め打ちにしない。"""
    for r in range(1, 60):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.startswith("=DATE("):
            return r
    raise ValueError(f"{ws.title}: 明細開始行が見つかりません")


def fix_holidays(wb):
    """祝日リストを2026年のものに差し替える。書式は元の行のものを使い回す。"""
    ws = wb["祝日リスト"]
    styles = []
    for c in range(1, 5):
        src = ws.cell(2, c)
        styles.append((copy(src.font), copy(src.border), copy(src.fill),
                       copy(src.alignment), src.number_format))

    for r in range(2, ws.max_row + 1):          # 旧データ(2025年分)を消す
        for c in range(1, 5):
            ws.cell(r, c).value = None

    for i, ((m, d), name, memo) in enumerate(HOLIDAYS_2026, start=2):
        # 曜日はロケールに依存しない CHOOSE(WEEKDAY()) で組み立てる
        # （TEXT(...,"aaa") だと、開いた環境の言語設定によって英語表記になりうる）
        dow = (f'=IF(A{i}="","","（"&'
               f'CHOOSE(WEEKDAY(A{i},2),"月","火","水","木","金","土","日")&"）")')
        vals = [datetime.date(YEAR, m, d), dow, name, memo]
        for c, (v, st) in enumerate(zip(vals, styles), start=1):
            cell = ws.cell(i, c, v)
            (cell.font, cell.border, cell.fill,
             cell.alignment, cell.number_format) = st
    return len(HOLIDAYS_2026)


HOLIDAY_DATES = {datetime.date(YEAR, m, d): nm for (m, d), nm, _ in HOLIDAYS_2026}


def extract_month_records(ws, anchor, ndays):
    """既存シートから、その月に入力済みの勤務データを読み取る。

    12シートは元々レイアウトが揃っていなかった(4月だけ開始行が2行ずれている等)
    ため、複製前のこの段階で「anchorはシートごとに検出する」必要がある。
    """
    records = {}
    for day in range(1, ndays + 1):
        r = anchor + day - 1
        vals, bad = {}, []
        for c in (C_IN, C_OUT, C_OT, C_OT_N, C_OT_E, C_LATE):
            v = ws.cell(r, c).value
            if isinstance(v, datetime.time):
                vals[c] = v          # そのまま書き込む(openpyxlが時刻値に変換)
            elif isinstance(v, (int, float)):
                vals[c] = v
            elif isinstance(v, str) and v.strip() and v.strip() not in JUNK:
                bad.append((c, v.strip()))
        k = ws.cell(r, C_REASON).value
        note = None
        if (isinstance(k, str) and not k.startswith("=") and k.strip()
                and k.strip() not in JUNK and k.strip() not in HOLIDAY_NAMES):
            note = k.strip()
        if vals or bad or note:
            records[day] = {"vals": vals, "bad": bad, "note": note}
    return records


def build_template(ws):
    """1つのシートを土台として体裁を整える(数式・入力規則・印刷設定など)。

    この関数の結果を他の11か月へ複製することで、12シートの体裁を完全に
    そろえる。データ(出勤・退勤・理由など)はここでは触らない。
    """
    anchor = find_anchor(ws)
    last = anchor + 30
    ws.cell(anchor - 1, C_REASON, "理由")

    for r in range(anchor, last + 1):
        # 曜日欄。ロケール非依存の CHOOSE(WEEKDAY()) にする
        # (TEXT(...,"aaa") だと、開いた環境の言語設定で英語表記になりうる)
        ws.cell(r, 3,
                f'=IF(A{r}="","","("&'
                f'CHOOSE(WEEKDAY(A{r},2),"月","火","水","木","金","土","日")&"）")')
        # 非表示の作業列。入力値のシリアル値変換、日ごとの実働、曜日区分。
        # 表には出さず、集計だけがここを参照する。
        for src, dst in C_CONV.items():
            ws.cell(r, dst, "=" + to_serial(f"${get_column_letter(src)}{r}"))
        ws.cell(r, C_DAILY_WORK,
                f'=IF(OR($N{r}="",$O{r}=""),"",MOD($O{r}-$N{r},1))')
        # 祝日かどうかは K列の文字(手入力で書き換えられうる)ではなく、
        # 祝日リストを直接参照して判定する
        ws.cell(r, C_KIND,
                f'=IF($A{r}="","",IF(ISNUMBER(MATCH($A{r},祝日リスト!$A$2:$A$40,0)),"祝",'
                f'IF(WEEKDAY($A{r},2)=6,"土",IF(WEEKDAY($A{r},2)=7,"日","平"))))')
        for c in range(C_KIND + 1, 30):         # 作業列より右の残骸を掃除
            ws.cell(r, c).value = None

    # --- 集計。行の位置と見出しは元のまま、式だけ実績ベースに直す ---
    s = find_row(ws, 3, "出勤日数")
    work = f"$T${anchor}:$T${last}"
    kind = f"$U${anchor}:$U${last}"
    for i in range(1, 4):                        # 元は空欄だった行の書式をそろえる
        copy_style(ws.cell(s + i, 4), ws.cell(s, 4))
    ws.cell(s, 4, f"=COUNT($N${anchor}:$N${last})").number_format = "0"
    ws.cell(s + 1, 4, f'=SUM({work})-SUMIF({kind},"土",{work})')
    ws.cell(s + 2, 4,
            f"=SUM($P${anchor}:$P${last})+SUM($Q${anchor}:$Q${last})"
            f"+SUM($R${anchor}:$R${last})")
    ws.cell(s + 3, 4, f'=SUMIF({kind},"土",{work})')
    for i in range(1, 4):
        ws.cell(s + i, 4).number_format = "[h]:mm"
    ws.cell(s + 3, 10, f"=SUM({work})").number_format = "[h]:mm"

    apply_dv_and_print(ws, anchor, last, s)
    return anchor, last, s


def apply_dv_and_print(ws, anchor, last, s):
    """入力規則・印刷範囲・行の色分け。copy_worksheet ではどれもコピーされない
    ため、複製後のシートにも毎回かけ直す必要がある(テンプレート自身にも
    同じ処理でよい)。"""
    # コロン付き(9:30, 9：00)入力を標準とする。コロンなし(930)は数字分割書式が
    # 実機Excelで日付として誤表示される不具合の原因だったため対象外にした。
    dv = DataValidation(
        type="custom", allow_blank=True, showErrorMessage=True,
        formula1=f'=OR(AND(ISNUMBER(D{anchor}),D{anchor}>=0,D{anchor}<1),'
                 f'IFERROR(ISNUMBER(TIMEVALUE(SUBSTITUTE(D{anchor},'
                 f'"：",":"))),FALSE))')
    dv.errorTitle = "時刻の入力"
    dv.error = "9:30 のように、コロン付きで入力してください（全角の 9：00 でも可）。"
    ws.add_data_validation(dv)
    for c in INPUT_COLS:
        col = get_column_letter(c)
        dv.add(f"{col}{anchor}:{col}{last}")

    for c in HIDDEN:
        ws.column_dimensions[get_column_letter(c)].hidden = True
    ws.print_area = f"A1:K{s + 3}"

    # 土日・水曜・祝日の行を同じ色で塗る。祝日かどうかは K列の文字(手入力で
    # 書き換えられうる)ではなく、祝日リストを直接参照して判定する
    ws.conditional_formatting.add(
        f"A{anchor}:K{last}",
        FormulaRule(
            formula=[f'AND($A{anchor}<>"",OR(WEEKDAY($A{anchor},2)=6,'
                     f'WEEKDAY($A{anchor},2)=7,WEEKDAY($A{anchor},2)=3,'
                     f'ISNUMBER(MATCH($A{anchor},祝日リスト!$A$2:$A$40,0))))'],
            fill=HOLIDAY_FILL))


def populate_month(ws, month, anchor, last, records, name):
    """1つのシートに、その月の実データを書き込む。

    12シートとも build_template で作った同一構造の複製なので、ここでは
    値を入れるだけでよい(体裁は一切いじらない)。
    """
    ws.cell(2, 3, month)
    ws.cell(2, C_REASON, name)

    ndays = days_in_month(month)
    salvaged = []

    for r in range(anchor, last + 1):
        day = r - anchor + 1
        date_val = datetime.date(YEAR, month, day) if day <= ndays else None

        rec = records.get(day, {})
        vals = rec.get("vals", {})
        for c in (C_IN, C_OUT, C_OT, C_OT_N, C_OT_E, C_LATE):
            cell = ws.cell(r, c)
            cell.value = vals.get(c)
            cell.number_format = TIME_FMT
        for c, raw in rec.get("bad", []):
            salvaged.append((day, ws.cell(r, c).coordinate,
                              f"{COL_LABEL[c]}「{raw}」"))

        note = rec.get("note")
        row_bad = [x for d, _, x in salvaged if d == day]
        holiday_name = HOLIDAY_DATES.get(date_val)
        k = ws.cell(r, C_REASON)
        if row_bad:
            # 時刻として読めなかった入力は消したままにせず、理由欄に退避する
            flag = "要確認: " + "・".join(row_bad)
            k.value = f"{note} / {flag}" if note else flag
        elif note:
            k.value = note
        else:
            k.value = (f'=IF($A{r}="","",'
                       f'IFERROR(VLOOKUP($A{r},祝日リスト!$A$2:$C$40,3,0),""))')

    return salvaged


TEMPLATE_MONTH = 1  # このシートの体裁を12か月ぶん複製する
HIRE_DATE = datetime.date(2026, 4, 15)  # 入社日。これより前のデータは使わない
RESET_MONTHS = {7}  # 入社日以降でも、一旦データをリセットする月


def main():
    wb = load_workbook(SRC)

    name = ""
    for m in range(1, 13):
        v = wb[f"{m}月"].cell(2, C_REASON).value
        if isinstance(v, str) and v.strip() and v.strip() != "㊞":
            name = v
            break

    # 複製・削除する前に、12シート分の実データをそれぞれの元の行位置から
    # 読み取っておく(4月だけ開始行が2行ずれているなど、シートごとに構造が
    # 揃っていなかったため)。入社日(HIRE_DATE)より前のデータは使わない。
    month_records = {}
    for m in range(1, 13):
        ws = wb[f"{m}月"]
        anchor = find_anchor(ws)
        records = extract_month_records(ws, anchor, days_in_month(m))
        month_records[m] = {} if m in RESET_MONTHS else {
            day: rec for day, rec in records.items()
            if datetime.date(YEAR, m, day) >= HIRE_DATE
        }

    n = fix_holidays(wb)
    print(f"祝日リスト: {YEAR}年 {n}件に差し替え")

    # 1つのシートだけ体裁を整え、残り11か月はそれを複製して作る。
    # 個別に直すのではなく複製することで、12シートの体裁を完全にそろえる。
    template = wb[f"{TEMPLATE_MONTH}月"]
    anchor, last, s = build_template(template)

    for m in range(1, 13):
        if m != TEMPLATE_MONTH:
            del wb[f"{m}月"]

    sheets = {TEMPLATE_MONTH: template}
    for m in range(1, 13):
        if m != TEMPLATE_MONTH:
            new_ws = wb.copy_worksheet(template)
            new_ws.title = f"{m}月"
            apply_dv_and_print(new_ws, anchor, last, s)  # copy_worksheetは引き継がない
            sheets[m] = new_ws

    for m in range(1, 13):
        for day, addr, v in populate_month(sheets[m], m, anchor, last,
                                            month_records[m], name):
            print(f"  要確認 {m}月{day}日 {addr} = {v!r}（時刻として読めないため削除）")

    holiday_ws = wb["祝日リスト"]
    wb._sheets = [sheets[m] for m in range(1, 13)] + [holiday_ws]

    wb.save(DST)
    print(f"saved {DST}: {wb.sheetnames}")


if __name__ == "__main__":
    main()
