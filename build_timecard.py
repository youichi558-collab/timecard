#!/usr/bin/env python3
"""出勤簿ブックの生成スクリプト。

旧ブック(出勤簿_2026_original.xlsx)から入力済みデータを引き継ぎつつ、
使い勝手を改善した新ブックを出力する。
"""

import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

SRC = "出勤簿_2026_original.xlsx"
DST = "出勤簿_2026.xlsx"
YEAR = 2026

# 日次明細の行範囲(31日分)とレイアウト
ROW_HEAD = 7          # 見出し行
ROW_FIRST = 8         # 1日目
ROW_LAST = ROW_FIRST + 30  # 38 = 31日目
COL = {
    "date": 1, "dow": 2, "in": 3, "out": 4, "brk": 5, "work": 6,
    "ot": 7, "ot_night": 8, "ot_early": 9, "late": 10,
    "holiday": 11, "note": 12,
    # M列以降は非表示の作業列。区分の判定と、時刻入力のシリアル値変換に使う。
    "kind": 13,
    "_in": 14, "_out": 15, "_brk": 16,
    "_ot": 17, "_ot_night": 18, "_ot_early": 19, "_late": 20,
}
HIDDEN_FIRST, HIDDEN_LAST = 13, 20
LAST_COL = 12  # 印刷対象はL列まで

# 時刻入力欄の表示形式。930 という数値を 09:30 と表示する。
# 日付/時刻書式ではなく数値書式にしておくことが重要で、こうしておくと
# 930 が日付シリアル値と解釈されない。コロン付きで 9:30 と入力した場合は
# Excel 側が時刻値として取り込み、そのセルの書式を h:mm に置き換える。
TIME_FMT = '00":"00'

FONT = "Meiryo UI"

# 2026年の国民の祝日
HOLIDAYS_2026 = [
    ((1, 1), "元日", ""),
    ((1, 12), "成人の日", ""),
    ((2, 11), "建国記念の日", ""),
    ((2, 23), "天皇誕生日", ""),
    ((3, 20), "春分の日", ""),
    ((4, 29), "昭和の日", ""),
    ((5, 3), "憲法記念日", ""),
    ((5, 4), "みどりの日", ""),
    ((5, 5), "こどもの日", ""),
    ((5, 6), "休日", "祝日法第3条第2項による休日"),
    ((7, 20), "海の日", ""),
    ((8, 11), "山の日", ""),
    ((9, 21), "敬老の日", ""),
    ((9, 22), "国民の休日", "祝日法第3条第3項による休日"),
    ((9, 23), "秋分の日", ""),
    ((10, 12), "スポーツの日", ""),
    ((11, 3), "文化の日", ""),
    ((11, 23), "勤労感謝の日", ""),
]

# 色
C_HEAD = "FF1F3864"       # 見出し背景(濃紺)
C_HEAD_TXT = "FFFFFFFF"
C_INPUT = "FFFFFBE6"      # 入力欄(淡黄)
C_CALC = "FFF2F2F2"       # 自動計算欄(灰)
C_SAT = "FFE3EFFB"        # 土曜
C_SUN = "FFFDE6E6"        # 日曜・祝日
C_BLANK = "FFEDEDED"      # 存在しない日
C_TITLE = "FFDCE6F1"

thin = Side(style="thin", color="FFB0B0B0")
med = Side(style="medium", color="FF1F3864")


def box(left=thin, right=thin, top=thin, bottom=thin):
    return Border(left=left, right=right, top=top, bottom=bottom)


def hhmm(t):
    """datetime.time を、この帳簿の入力形式である HHMM の整数にする。"""
    return t.hour * 100 + t.minute


def to_serial(ref):
    """時刻入力欄をシリアル値(1日=1)に変換する数式を返す。

    コロンなしで 930 と入れた場合は 9時30分、コロン付きで 9:30 と入れた場合は
    Excel が時刻値(1未満)として持つのでそのまま使う。どちらの打ち方でも通る。
    """
    return (f'IF(NOT(ISNUMBER({ref})),"",'
            f'IF({ref}<1,{ref},(INT({ref}/100)*60+MOD({ref},100))/1440))')


HOLIDAY_NAMES = {name for _, name, _ in HOLIDAYS_2026}
# 旧ブックで欄に残っていた飾り文字・見出し語。データではないので取り込まない。
JUNK = {"：", "～", "日", "時間", "出勤日数", "普通出勤時間", "残業", "土曜出勤", "合計"}


def days_in_month(month):
    nxt = datetime.date(YEAR + month // 12, month % 12 + 1, 1)
    return (nxt - datetime.timedelta(days=1)).day


def find_anchor(ws):
    """明細1日目の行を探す。月によって開始行が違うため決め打ちにしない。"""
    for r in range(1, 60):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.startswith("=DATE("):
            return r
    raise ValueError(f"{ws.title}: 明細開始行が見つかりません")


def read_source():
    """旧ブックから入力済みの勤務データを読み取る。

    取り込むのは実際の時刻値と自由記述の備考だけ。旧ブックに残っていた
    プレースホルダ文字や、書式崩れで数式欄に入っていた文字列は捨て、
    捨てた内容は skipped に記録して呼び出し側で報告できるようにする。
    """
    wb = load_workbook(SRC)
    data, skipped, name = {}, [], None
    for m in range(1, 13):
        ws = wb[f"{m}月"]
        anchor = find_anchor(ws)
        if not name and isinstance(ws["K2"].value, str) and ws["K2"].value != "㊞":
            name = ws["K2"].value
        recs = {}
        labels = {4: "出勤", 6: "退勤", 7: "通常残業", 8: "深夜残業",
                  9: "早朝残業", 10: "遅刻・早退"}
        for day in range(1, days_in_month(m) + 1):
            r = anchor + day - 1
            times, salvaged = [], []
            for c in (4, 6, 7, 8, 9, 10):  # 出勤/退勤/通常/深夜/早朝/遅刻早退
                v = ws.cell(r, c).value
                if isinstance(v, datetime.time):
                    times.append(v)
                else:
                    if isinstance(v, str) and v.strip() and v.strip() not in JUNK:
                        skipped.append((m, day, ws.cell(r, c).coordinate, v))
                        salvaged.append(f"{labels[c]}「{v.strip()}」")
                    times.append(None)
            k = ws.cell(r, 11).value
            note = None
            if isinstance(k, str) and not k.startswith("="):
                k = k.strip()
                if k and k not in JUNK and k not in HOLIDAY_NAMES:
                    note = k
            if salvaged:
                # 時刻として読めなかった入力は失わずに備考へ退避する
                flag = "要確認: " + "・".join(salvaged)
                note = f"{note} / {flag}" if note else flag
            if any(times) or note:
                recs[day] = {"in": times[0], "out": times[1],
                             "note": note, "ot": times[2:]}
        data[m] = recs
    return data, skipped, name


def build_settings(wb, name=""):
    ws = wb.create_sheet("設定")
    ws.sheet_properties.tabColor = "FF1F3864"
    ws["A1"] = "出勤簿 共通設定"
    ws["A1"].font = Font(name=FONT, sz=14, b=True, color=C_HEAD)

    rows = [
        ("年", YEAR, "この年を全シートの日付・祝日判定に使います"),
        ("氏名", name, "各月シートの氏名欄に自動反映されます"),
        ("所属", "", "各月シートの所属欄に自動反映されます"),
        ("所定労働時間/日", 800, "1日の所定時間。超過分が「所定超過」に出ます"),
        ("休憩時間(既定)", 100, "休憩欄が空のとき、この時間を自動で差し引きます"),
        ("休憩を差し引く最低勤務時間", 600, "この時間を超えたときだけ既定休憩を引きます"),
    ]
    for i, (label, val, memo) in enumerate(rows, start=3):
        ws.cell(i, 1, label).font = Font(name=FONT, sz=11, b=True)
        ws.cell(i, 1).fill = PatternFill("solid", fgColor=C_TITLE)
        ws.cell(i, 1).border = box()
        c = ws.cell(i, 2, val)
        c.font = Font(name=FONT, sz=11)
        c.fill = PatternFill("solid", fgColor=C_INPUT)
        c.border = box()
        c.alignment = Alignment(horizontal="center")
        if i >= 6:  # 時間の設定欄。月シートと同じくコロンなしで入力できる
            c.number_format = TIME_FMT
            ws.cell(i, 5, "=" + to_serial(f"$B${i}"))
        ws.cell(i, 3, memo).font = Font(name=FONT, sz=9, color="FF666666")
    ws.column_dimensions["E"].hidden = True  # 月シートが参照する変換後の値

    ws["A11"] = "使い方"
    ws["A11"].font = Font(name=FONT, sz=12, b=True, color=C_HEAD)
    tips = [
        "1. まずこの設定シートで「氏名」「所属」「所定労働時間」を入力してください。",
        "2. 各月シートは、黄色いセル（出勤・退勤・休憩・残業・備考）だけ入力します。",
        "   時刻はコロンなしで入力できます。930 と打てば 09:30、1615 と打てば 16:15 になります。",
        "   9:30 のようにコロン付きで入力しても構いません。どちらでも正しく計算されます。",
        "   30分だけの休憩は 30、1時間なら 100 と入力します（HHMM形式のため）。",
        "3. 出勤と退勤を入れると「実働」が自動計算されます。日をまたぐ勤務にも対応しています。",
        "4. 休憩を空欄にすると、設定シートの既定休憩が自動で差し引かれます。個別に変えたい日は直接入力してください。",
        "5. 土曜は青、日曜・祝日はピンクで色分けされます。祝日名はK列に自動表示されます。",
        "6. 月次の集計はシート下部、年間の集計は「年間集計」シートに自動で出ます。",
        "7. 年が変わったら、この設定の「年」と「祝日リスト」シートを更新してください。",
    ]
    for i, t in enumerate(tips, start=12):
        ws.cell(i, 1, t).font = Font(name=FONT, sz=10)

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 52
    ws.sheet_view.showGridLines = False
    return ws


def build_holidays(wb):
    ws = wb.create_sheet("祝日リスト")
    heads = ["日付", "曜日", "名称", "備考"]
    for i, h in enumerate(heads, start=1):
        c = ws.cell(1, i, h)
        c.font = Font(name=FONT, sz=11, b=True, color=C_HEAD_TXT)
        c.fill = PatternFill("solid", fgColor=C_HEAD)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = box()
    for i, ((m, d), name, memo) in enumerate(HOLIDAYS_2026, start=2):
        ws.cell(i, 1, datetime.date(YEAR, m, d)).number_format = "yyyy/m/d"
        ws.cell(i, 2,
                f'=IF(A{i}="","",CHOOSE(WEEKDAY(A{i},2),"月","火","水","木","金","土","日"))')
        ws.cell(i, 3, name)
        ws.cell(i, 4, memo)
        for col in range(1, 5):
            c = ws.cell(i, col)
            c.font = Font(name=FONT, sz=11)
            c.border = box()
            c.alignment = Alignment(horizontal="center" if col < 3 else "left")

    ws.cell(len(HOLIDAYS_2026) + 3, 1, "※ 会社独自の休日もここに追加できます（40行目まで自動で参照されます）")
    ws.cell(len(HOLIDAYS_2026) + 3, 1).font = Font(name=FONT, sz=9, color="FF666666")

    for col, w in zip("ABCD", (13, 8, 20, 32)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    return ws


def build_month(wb, month, recs):
    ws = wb.create_sheet(f"{month}月")
    L = get_column_letter(LAST_COL)

    # --- ヘッダー ---
    ws.merge_cells(f"A2:D2")
    ws["A2"] = f'=設定!$B$4&"　"&設定!$B$3&"年"&{month}&"月度 出勤簿"'
    ws["A2"].font = Font(name=FONT, sz=16, b=True, color=C_HEAD)
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 24

    ws["A4"], ws["C4"] = "年", "月"
    ws["B4"] = "=設定!$B$3"
    ws["D4"] = month
    for a in ("A4", "C4"):
        ws[a].font = Font(name=FONT, sz=10, color="FF666666")
        ws[a].alignment = Alignment(horizontal="right")
    for a in ("B4", "D4"):
        ws[a].font = Font(name=FONT, sz=12, b=True)
        ws[a].alignment = Alignment(horizontal="center")
        ws[a].border = Border(bottom=Side(style="thin"))

    ws["G4"] = "所属"
    ws["H4"] = "=設定!$B$5"
    ws["I4"] = "氏名"
    ws.merge_cells("J4:L4")
    ws["J4"] = "=設定!$B$4"
    for a in ("G4", "I4"):
        ws[a].font = Font(name=FONT, sz=10, color="FF666666")
        ws[a].alignment = Alignment(horizontal="right")
    for a in ("H4", "J4"):
        ws[a].font = Font(name=FONT, sz=12, b=True)
        ws[a].alignment = Alignment(horizontal="center")
        ws[a].border = Border(bottom=Side(style="thin"))

    ws["A5"] = "黄色のセルに入力してください（灰色は自動計算）。時刻は 930 のようにコロンなしで入力できます。"
    ws["A5"].font = Font(name=FONT, sz=9, color="FF888888")

    # --- 見出し ---
    heads = [
        ("日", 4.5), ("曜", 4.5), ("出勤", 8.5), ("退勤", 8.5), ("休憩", 8.0),
        ("実働", 8.5), ("通常残業", 9.0), ("深夜残業", 9.0), ("早朝残業", 9.0),
        ("遅刻・早退", 10.0), ("祝日", 14.0), ("備考・理由", 26.0),
    ]
    for i, (h, w) in enumerate(heads, start=1):
        c = ws.cell(ROW_HEAD, i, h)
        c.font = Font(name=FONT, sz=10, b=True, color=C_HEAD_TXT)
        c.fill = PatternFill("solid", fgColor=C_HEAD)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(left=thin, right=thin, top=med, bottom=med)
        ws.column_dimensions[get_column_letter(i)].width = w
    for i in range(HIDDEN_FIRST, HIDDEN_LAST + 1):
        ws.column_dimensions[get_column_letter(i)].hidden = True
    ws.row_dimensions[ROW_HEAD].height = 26

    input_cols = {"in", "out", "brk", "ot", "ot_night", "ot_early", "late", "note"}

    # --- 日次明細 ---
    for r in range(ROW_FIRST, ROW_LAST + 1):
        day = r - ROW_FIRST + 1
        prev = r - 1
        if r == ROW_FIRST:
            ws.cell(r, COL["date"], "=DATE($B$4,$D$4,1)")
        else:
            ws.cell(r, COL["date"],
                    f'=IF(A{prev}="","",IF(MONTH(A{prev}+1)=$D$4,A{prev}+1,""))')
        ws.cell(r, COL["dow"],
                f'=IF($A{r}="","",CHOOSE(WEEKDAY($A{r},2),"月","火","水","木","金","土","日"))')
        # N〜T列: 入力欄をシリアル値に変換した作業列(非表示)
        for key in ("in", "out", "brk", "ot", "ot_night", "ot_early", "late"):
            src = f"${get_column_letter(COL[key])}{r}"
            ws.cell(r, COL["_" + key], "=" + to_serial(src))
        # 実働 = 退勤 - 出勤 (日跨ぎ対応) - 休憩(未入力なら既定休憩)
        ws.cell(r, COL["work"],
                f'=IF(OR($N{r}="",$O{r}=""),"",'
                f'MAX(0,MOD($O{r}-$N{r},1)'
                f'-IF($P{r}<>"",$P{r},'
                f'IF(MOD($O{r}-$N{r},1)>設定!$E$8,設定!$E$7,0))))')
        ws.cell(r, COL["holiday"],
                f'=IF($A{r}="","",IFERROR(VLOOKUP($A{r},祝日リスト!$A$2:$C$40,3,0),""))')
        # M列: 日区分(平日/土/日/祝) — 集計とCFの判定用
        ws.cell(r, COL["kind"],
                f'=IF($A{r}="","",IF($K{r}<>"","祝",'
                f'IF(WEEKDAY($A{r},2)=6,"土",IF(WEEKDAY($A{r},2)=7,"日","平"))))')

        for key, col in COL.items():
            c = ws.cell(r, col)
            c.font = Font(name=FONT, sz=11)
            c.border = box()
            c.alignment = Alignment(horizontal="center", vertical="center")
            if key in ("in", "out", "brk", "ot", "ot_night", "ot_early", "late"):
                c.number_format = TIME_FMT
            if key == "work":
                c.number_format = "[h]:mm"
            if key == "date":
                c.number_format = "d"
            if key == "note":
                c.alignment = Alignment(horizontal="left", vertical="center")
            if key == "holiday":
                c.alignment = Alignment(horizontal="left", vertical="center")
                c.font = Font(name=FONT, sz=9, color="FFC00000")
            if key in input_cols:
                c.fill = PatternFill("solid", fgColor=C_INPUT)
            elif col < HIDDEN_FIRST:
                c.fill = PatternFill("solid", fgColor=C_CALC)
        ws.cell(r, COL["date"]).border = Border(left=med, right=thin, top=thin, bottom=thin)
        ws.cell(r, LAST_COL).border = Border(left=thin, right=med, top=thin, bottom=thin)
        ws.row_dimensions[r].height = 19

        # 旧ブックの入力値を移行
        rec = recs.get(day)
        if rec:
            if rec["in"]:
                ws.cell(r, COL["in"], hhmm(rec["in"]))
            if rec["out"]:
                ws.cell(r, COL["out"], hhmm(rec["out"]))
            if rec["note"]:
                ws.cell(r, COL["note"], rec["note"])
            for k, v in zip(("ot", "ot_night", "ot_early", "late"), rec["ot"]):
                if v:
                    ws.cell(r, COL[k], hhmm(v))

    # 明細の下辺
    for i in range(1, LAST_COL + 1):
        c = ws.cell(ROW_LAST, i)
        b = c.border
        c.border = Border(left=b.left, right=b.right, top=b.top, bottom=med)

    # --- 条件付き書式 ---
    rng = f"A{ROW_FIRST}:{L}{ROW_LAST}"
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'$A{ROW_FIRST}=""'],
        fill=PatternFill("solid", start_color=C_BLANK, end_color=C_BLANK),
        stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'OR($M{ROW_FIRST}="祝",$M{ROW_FIRST}="日")'],
        fill=PatternFill("solid", start_color=C_SUN, end_color=C_SUN)))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'$M{ROW_FIRST}="土"'],
        fill=PatternFill("solid", start_color=C_SAT, end_color=C_SAT)))
    # 片方だけ入力された行は赤くして入力漏れに気付けるようにする
    alert = dict(
        fill=PatternFill("solid", start_color="FFFFC7CE", end_color="FFFFC7CE"),
        font=Font(name=FONT, sz=11, color="FF9C0006"))
    ws.conditional_formatting.add(f"D{ROW_FIRST}:D{ROW_LAST}", FormulaRule(
        formula=[f'AND($N{ROW_FIRST}<>"",$O{ROW_FIRST}="")'], **alert))
    ws.conditional_formatting.add(f"C{ROW_FIRST}:C{ROW_LAST}", FormulaRule(
        formula=[f'AND($O{ROW_FIRST}<>"",$N{ROW_FIRST}="")'], **alert))

    # --- 入力補助 ---
    dv = DataValidation(
        type="list", allow_blank=True, showErrorMessage=False,
        formula1='"有給,半休,欠勤,遅刻,早退,休日出勤,振替休日,直行,直帰,出張,特別休暇"')
    dv.prompt = "選択、または自由入力できます"
    dv.promptTitle = "備考"
    dv.showInputMessage = True
    ws.add_data_validation(dv)
    dv.add(f"L{ROW_FIRST}:L{ROW_LAST}")

    dvt = DataValidation(
        type="custom", allow_blank=True, showErrorMessage=True,
        formula1=f'=OR(AND(C{ROW_FIRST}>=0,C{ROW_FIRST}<1),'
                 f'AND(C{ROW_FIRST}=INT(C{ROW_FIRST}),C{ROW_FIRST}<=2359,'
                 f'MOD(C{ROW_FIRST},100)<60))')
    dvt.error = ("コロンなしで 930（＝9時30分）のように入力してください。"
                 "9:30 のようにコロン付きでも入力できます。")
    dvt.errorTitle = "時刻の入力"
    dvt.prompt = "930 と入力すれば 09:30 になります"
    dvt.promptTitle = "時刻の入力"
    dvt.showInputMessage = True
    ws.add_data_validation(dvt)
    for col in ("C", "D", "E", "G", "H", "I", "J"):
        dvt.add(f"{col}{ROW_FIRST}:{col}{ROW_LAST}")

    # --- 集計 ---
    F, M = f"$F${ROW_FIRST}:$F${ROW_LAST}", f"$M${ROW_FIRST}:$M${ROW_LAST}"
    top = ROW_LAST + 2
    ws.cell(top, 1, "月次集計").font = Font(name=FONT, sz=12, b=True, color=C_HEAD)
    ws.merge_cells(start_row=top, start_column=1, end_row=top, end_column=LAST_COL)
    ws.cell(top, 1).fill = PatternFill("solid", fgColor=C_TITLE)
    ws.cell(top, 1).alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[top].height = 20

    left = [
        ("出勤日数", f'=COUNT($N${ROW_FIRST}:$N${ROW_LAST})', "日", "0"),
        ("実働時間 合計", f"=SUM({F})", "", "[h]:mm"),
        ("　うち 平日", f'=SUMIF({M},"平",{F})', "", "[h]:mm"),
        ("　うち 土曜", f'=SUMIF({M},"土",{F})', "", "[h]:mm"),
        ("　うち 日曜・祝日", f'=SUMIF({M},"日",{F})+SUMIF({M},"祝",{F})', "", "[h]:mm"),
    ]
    right = [
        ("通常残業", f"=SUM($Q${ROW_FIRST}:$Q${ROW_LAST})", "[h]:mm"),
        ("深夜残業", f"=SUM($R${ROW_FIRST}:$R${ROW_LAST})", "[h]:mm"),
        ("早朝残業", f"=SUM($S${ROW_FIRST}:$S${ROW_LAST})", "[h]:mm"),
        ("遅刻・早退", f"=SUM($T${ROW_FIRST}:$T${ROW_LAST})", "[h]:mm"),
        ("所定超過(実働-所定×出勤日数)",
         f'=MAX(0,SUM({F})-設定!$E$6*COUNT($N${ROW_FIRST}:$N${ROW_LAST}))', "[h]:mm"),
    ]

    def put(row, col_label, col_val, label, formula, fmt, unit=""):
        lc = ws.cell(row, col_label, label)
        lc.font = Font(name=FONT, sz=10, b=not label.startswith("　"))
        lc.alignment = Alignment(horizontal="left", vertical="center")
        lc.border = box()
        lc.fill = PatternFill("solid", fgColor="FFF7F9FC")
        vc = ws.cell(row, col_val, formula)
        vc.font = Font(name=FONT, sz=11, b=True)
        vc.number_format = fmt
        vc.alignment = Alignment(horizontal="center", vertical="center")
        vc.border = box()
        vc.fill = PatternFill("solid", fgColor=C_CALC)
        if unit:
            u = ws.cell(row, col_val + 1, unit)
            u.font = Font(name=FONT, sz=9, color="FF666666")
            u.alignment = Alignment(horizontal="left", vertical="center")

    for i, (label, f, unit, fmt) in enumerate(left):
        r = top + 1 + i
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        put(r, 1, 4, label, f, fmt, unit)
        ws.row_dimensions[r].height = 18
    for i, (label, f, fmt) in enumerate(right):
        r = top + 1 + i
        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=10)
        put(r, 7, 11, label, f, fmt)

    note_r = top + 7
    ws.cell(note_r, 1, "※ 残業欄は実働時間の内数です。実働は「退勤－出勤－休憩」で自動計算されます。")
    ws.cell(note_r, 1).font = Font(name=FONT, sz=9, color="FF888888")

    ws.cell(note_r + 2, 7, "承認")
    ws.cell(note_r + 2, 7).font = Font(name=FONT, sz=10, color="FF666666")
    ws.cell(note_r + 2, 7).alignment = Alignment(horizontal="right")
    for c0, lab in ((8, "本人"), (10, "上長")):
        ws.merge_cells(start_row=note_r + 2, start_column=c0, end_row=note_r + 2, end_column=c0 + 1)
        cell = ws.cell(note_r + 2, c0)
        cell.border = box(top=thin, bottom=thin, left=thin, right=thin)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        h = ws.cell(note_r + 3, c0, lab)
        h.font = Font(name=FONT, sz=8, color="FF888888")
        h.alignment = Alignment(horizontal="center")
    ws.row_dimensions[note_r + 2].height = 30

    # --- 表示・印刷設定 ---
    ws.freeze_panes = f"A{ROW_FIRST + 1}"
    ws.sheet_view.showGridLines = False
    ws.print_area = f"A1:{L}{note_r + 3}"
    ws.print_title_rows = f"1:{ROW_HEAD}"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top = ws.page_margins.bottom = 0.4
    ws.sheet_properties.tabColor = "FF4472C4" if month % 2 else "FF8FAADC"
    return ws


def build_yearly(wb):
    ws = wb.create_sheet("年間集計")
    ws.sheet_properties.tabColor = "FFC00000"
    ws["A1"] = '=設定!$B$4&"　"&設定!$B$3&"年 年間集計"'
    ws["A1"].font = Font(name=FONT, sz=16, b=True, color=C_HEAD)
    ws.row_dimensions[1].height = 24

    heads = ["月", "出勤日数", "実働時間", "平日", "土曜", "日祝",
             "通常残業", "深夜残業", "早朝残業", "遅刻・早退"]
    for i, h in enumerate(heads, start=1):
        c = ws.cell(3, i, h)
        c.font = Font(name=FONT, sz=10, b=True, color=C_HEAD_TXT)
        c.fill = PatternFill("solid", fgColor=C_HEAD)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = box()
    ws.row_dimensions[3].height = 22

    base = ROW_LAST + 2  # 各月シートの集計見出し行
    src = {
        2: f"D{base + 1}", 3: f"D{base + 2}", 4: f"D{base + 3}",
        5: f"D{base + 4}", 6: f"D{base + 5}",
        7: f"K{base + 1}", 8: f"K{base + 2}", 9: f"K{base + 3}", 10: f"K{base + 4}",
    }
    for m in range(1, 13):
        r = 3 + m
        ws.cell(r, 1, f"{m}月").font = Font(name=FONT, sz=11, b=True)
        ws.cell(r, 1).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(r, 1).border = box()
        ws.cell(r, 1).fill = PatternFill("solid", fgColor="FFF7F9FC")
        for col, addr in src.items():
            c = ws.cell(r, col, f"='{m}月'!{addr}")
            c.font = Font(name=FONT, sz=11)
            c.number_format = "0" if col == 2 else "[h]:mm"
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = box()
        ws.row_dimensions[r].height = 18

    r = 16
    ws.cell(r, 1, "合計").font = Font(name=FONT, sz=11, b=True, color=C_HEAD_TXT)
    ws.cell(r, 1).fill = PatternFill("solid", fgColor=C_HEAD)
    ws.cell(r, 1).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(r, 1).border = box()
    for col in range(2, 11):
        cl = get_column_letter(col)
        c = ws.cell(r, col, f"=SUM({cl}4:{cl}15)")
        c.font = Font(name=FONT, sz=11, b=True)
        c.number_format = "0" if col == 2 else "[h]:mm"
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = Border(left=thin, right=thin, top=med, bottom=med)
        c.fill = PatternFill("solid", fgColor=C_TITLE)
    ws.row_dimensions[r].height = 22

    ws.cell(18, 1, "※ 各月シートの集計を自動で集めています。この表は直接編集しないでください。")
    ws.cell(18, 1).font = Font(name=FONT, sz=9, color="FF888888")

    ws.column_dimensions["A"].width = 7
    for col in "BCDEFGHIJ":
        ws.column_dimensions[col].width = 11
    ws.freeze_panes = "B4"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    return ws


def main():
    data, skipped, name = read_source()
    wb = Workbook()
    wb.remove(wb.active)
    build_settings(wb, name or "")
    for m in range(1, 13):
        build_month(wb, m, data[m])
    build_yearly(wb)
    build_holidays(wb)
    wb.active = wb.index(wb["1月"])
    wb.save(DST)
    print(f"saved {DST}: {wb.sheetnames}")
    print(f"氏名: {name!r}")
    total = sum(len(v) for v in data.values())
    print(f"移行した日数: {total}")
    if skipped:
        print("取り込まなかった値（旧ブックの入力ミス／書式崩れ）:")
        for m, day, addr, v in skipped:
            print(f"  {m}月{day}日 旧{addr} = {v!r}")


if __name__ == "__main__":
    main()
