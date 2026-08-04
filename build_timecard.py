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
  - 時刻をコロンなし(930 → 09:30)で入力できるようにする
"""

import datetime
import re
import xml.etree.ElementTree as ET
import zipfile
from copy import copy
from xml.sax.saxutils import escape

from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import to_excel
from openpyxl.worksheet.datavalidation import DataValidation

# 土日・水曜・祝日をこの色で塗る(休日として同じ色でまとめる)
# 条件付き書式のdxfは、通常のセルと違いbgColor側に色を入れないと
# Excel実機では表示されない(fgColorだけだとLibreOfficeは表示するがExcelは無視する)。
HOLIDAY_FILL = PatternFill("solid", fgColor="FFFFC7CE", bgColor="FFFFC7CE")

# 数式セルの計算結果(キャッシュ値)。{シート名: {セル番地: 値}}
# openpyxl は数式の文字列だけを保存し、計算済みの値は保存しない。そのため
# Excel が開いて再計算するまで、日付などが一瞬空欄に見えることがある
# (fullCalcOnLoad で通常は自動的に直るが、再計算しないビューアもある)。
# ここでは自分で書いた数式の結果を Python側でも計算しておき、保存後に
# セルのスタイルには一切触れずに <v> タグだけを差し込んで解消する。
CACHE = {}


def cache(ws, coord, value):
    if value is not None:
        CACHE.setdefault(ws.title, {})[coord] = value

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

# 時刻入力欄の表示形式。930 という数値を 09:30 と表示する。
# 日付/時刻書式ではなく数値書式にしておくのが要点で、こうしておけば 930 が
# 日付シリアル値と解釈されない。コロン付きで 9:30 と入力した場合は Excel 側が
# 時刻値として取り込み、そのセルの書式を h:mm に置き換える。
TIME_FMT = '00":"00'

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

    コロンなしで 930 と入れた場合は 9時30分として換算し、コロン付きで 9:30 と
    入れた場合は Excel が時刻値(1未満)として持つのでそのまま使う。
    """
    return (f'IF(NOT(ISNUMBER({ref})),"",'
            f'IF({ref}<1,{ref},(INT({ref}/100)*60+MOD({ref},100))/1440))')


def hhmm(t):
    """datetime.time を、この帳簿の入力形式である HHMM の整数にする。"""
    return t.hour * 100 + t.minute


WEEKDAY_JA = ["月", "火", "水", "木", "金", "土", "日"]


def days_in_month(month):
    nxt = datetime.date(YEAR + month // 12, month % 12 + 1, 1)
    return (nxt - datetime.timedelta(days=1)).day


def frac(v):
    """HHMM整数を、Excelの時刻シリアル値(1日=1)相当の小数に変換する。"""
    return None if v is None else ((v // 100) * 60 + v % 100) / 1440


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
        wd = WEEKDAY_JA[datetime.date(YEAR, m, d).weekday()]
        cache(ws, f"B{i}", f"（{wd}）")
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
                vals[c] = hhmm(v)
            elif isinstance(v, (int, float)):
                vals[c] = int(v)
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
    dv = DataValidation(
        type="custom", allow_blank=True, showErrorMessage=True,
        formula1=f'=OR(AND(D{anchor}>=0,D{anchor}<1),'
                 f'AND(D{anchor}=INT(D{anchor}),D{anchor}<=2359,'
                 f'MOD(D{anchor},100)<60))')
    dv.errorTitle = "時刻の入力"
    dv.error = ("コロンなしで 930（＝9時30分）のように入力してください。"
                "9:30 のようにコロン付きでも入力できます。")
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


def populate_month(ws, month, anchor, last, s, records, name):
    """1つのシートに、その月の実データを書き込む。

    12シートとも build_template で作った同一構造の複製なので、ここでは
    値を入れるだけでよい(体裁は一切いじらない)。
    """
    ws.cell(2, 3, month)
    ws.cell(2, C_REASON, name)

    ndays = days_in_month(month)
    salvaged = []
    count_in = 0
    work_sum = sat_work_sum = ot_sum = 0.0

    for r in range(anchor, last + 1):
        day = r - anchor + 1
        date_val = datetime.date(YEAR, month, day) if day <= ndays else None
        cache(ws, f"A{r}", to_excel(date_val) if date_val else None)
        if date_val:
            cache(ws, f"C{r}", f"({WEEKDAY_JA[date_val.weekday()]}）")

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
            cache(ws, k.coordinate, holiday_name)

        # 集計欄のキャッシュ値を作るため、この行の実働・残業を集計しておく
        in_f, out_f = frac(vals.get(C_IN)), frac(vals.get(C_OUT))
        work_f = None if in_f is None or out_f is None else (out_f - in_f) % 1
        if date_val and vals.get(C_IN) is not None:
            count_in += 1
        if work_f is not None:
            work_sum += work_f
            if date_val and holiday_name is None and date_val.weekday() == 5:
                sat_work_sum += work_f
        for c in (C_OT, C_OT_N, C_OT_E):
            fv = frac(vals.get(c))
            if fv is not None:
                ot_sum += fv

    cache(ws, ws.cell(s, 4).coordinate, count_in)
    cache(ws, ws.cell(s + 1, 4).coordinate, work_sum - sat_work_sum)
    cache(ws, ws.cell(s + 2, 4).coordinate, ot_sum)
    cache(ws, ws.cell(s + 3, 4).coordinate, sat_work_sum)
    cache(ws, ws.cell(s + 3, 10).coordinate, work_sum)
    return salvaged


NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_PREL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def fmt_value(v):
    """CACHE の値を <v> タグの中身にする。文字列かどうかも一緒に返す。"""
    if isinstance(v, str):
        return True, escape(v)
    if isinstance(v, float) and v.is_integer():
        return False, str(int(v))
    if isinstance(v, int):
        return False, str(v)
    return False, ("%.10f" % v).rstrip("0").rstrip(".")


def sheet_file_map(data):
    """シート名 → xl/worksheets/sheetN.xml のパスを workbook.xml から求める。"""
    wb_xml = ET.fromstring(data["xl/workbook.xml"])
    rels = ET.fromstring(data["xl/_rels/workbook.xml.rels"])
    rid_target = {rel.get("Id"): rel.get("Target")
                  for rel in rels.findall(f"{NS_PREL}Relationship")}
    mapping = {}
    for sheet in wb_xml.find(f"{NS_MAIN}sheets").findall(f"{NS_MAIN}sheet"):
        target = rid_target[sheet.get(f"{NS_R}id")]
        if not target.startswith("worksheets/"):
            target = "worksheets/" + target.rsplit("/", 1)[-1]
        mapping[sheet.get("name")] = "xl/" + target
    return mapping


def patch_sheet_xml(xml_text, cell_cache):
    for coord, value in cell_cache.items():
        is_str, text = fmt_value(value)
        m = re.search(rf'(<c r="{re.escape(coord)}"[^>]*>)(.*?)(</c>)',
                       xml_text, re.S)
        if not m:
            continue
        open_tag, inner, close_tag = m.groups()
        if is_str:
            open_tag = (re.sub(r' t="[^"]*"', ' t="str"', open_tag)
                        if ' t="' in open_tag else open_tag[:-1] + ' t="str">')
        inner = re.sub(r"<v\s*/>|<v>.*?</v>", "", inner, flags=re.S) + f"<v>{text}</v>"
        xml_text = xml_text[:m.start()] + open_tag + inner + close_tag + xml_text[m.end():]
    return xml_text


def inject_cache(path):
    """保存済みファイルに、数式の計算結果(CACHE)だけを直接書き込む。

    openpyxl のスタイル出力には一切触れない。<c> タグの中身に <v> を足すだけの
    XML パッチなので、見た目(フォント・罫線・列幅など)は完全にそのまま保たれる。
    """
    if not CACHE:
        return
    with zipfile.ZipFile(path, "r") as zin:
        names = zin.namelist()
        data = {n: zin.read(n) for n in names}
    mapping = sheet_file_map(data)
    for sheet, cell_cache in CACHE.items():
        file = mapping.get(sheet)
        if not file:
            continue
        data[file] = patch_sheet_xml(data[file].decode("utf-8"),
                                      cell_cache).encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, data[n])


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
        for day, addr, v in populate_month(sheets[m], m, anchor, last, s,
                                            month_records[m], name):
            print(f"  要確認 {m}月{day}日 {addr} = {v!r}（時刻として読めないため削除）")

    holiday_ws = wb["祝日リスト"]
    wb._sheets = [sheets[m] for m in range(1, 13)] + [holiday_ws]

    wb.save(DST)
    # inject_cache(DST)  # 実機Excelでは fullCalcOnLoad により自動再計算されるため
    # 不要。むしろこの生XMLパッチが原因で条件付き書式が読み込まれない不具合が
    # 疑われるため、当面は無効化して切り分ける。
    print(f"saved {DST}: {wb.sheetnames}")


if __name__ == "__main__":
    main()
