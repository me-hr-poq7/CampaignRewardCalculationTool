# -*- coding: utf-8 -*-

"""
キャンペーン参加者 集計処理

【処理内容】

1. 「集計前」シートのB列に「取消」がある行を除外
2. A列（名前）＋C列（メール）をキーにしてD列（金額）を合計
3. 同じExcelブック内に「集計後」シートを作成
4. 合計金額の降順で並べ替え
5. 集計後の合計金額を基準に付与額を判定
6. D列の付与額ごとに色分け
7. F〜G 列の付与額サマリを件数の降順で出力する
8. F1:G5に付与額ごとの件数を表示
9. 集計後シートの A, C, D, F, G 列のみ自動調整する（他列は調整しない）
10. 罫線は A〜D 列、F〜G 列に分けて、データが入っている最後の行まで引く（E列は罫線なし） 
"""

import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side


# ============================================================
# ユーザー設定
# ============================================================

INPUT_XLSX = r"キャンペーン参加者.xlsx"
INPUT_SHEET_NAME = "集計前"
OUTPUT_SHEET_NAME = "集計後"


# ============================================================
# 色分け設定
# ============================================================

COLOR_MAP = {
    "": "FFFFFF",
    "1万円": "C6EFCE",
    "2万円": "C9DAF8",
    "3万円": "FCE4D6",
    "4万円": "D9D9D9",
}


# ============================================================
# 金額を円に変換する関数
# ============================================================

def parse_money_to_yen(value):
    """
    金額の値を整数の円に変換する。
    例：
    1,234円
    1234567
    ¥1,234
    """
    if pd.isna(value):
        return 0
    if isinstance(value, (int, float)):
        return int(value)

    text = str(value).strip()
    if text == "":
        return 0

    text = re.sub(r"[,\s円¥$]", "", text)

    if text == "":
        return 0

    try:
        return int(float(text))
    except ValueError:
        return 0


# ============================================================
# 合計金額から付与額を判定する関数
# ============================================================

def compute_tier_text(amount_yen):
    """
    集計後の合計金額から付与額を判定する。

    100万円未満           → 空白
    100万円～200万円未満  → 1万円
    200万円～300万円未満  → 2万円
    300万円～400万円未満  → 3万円
    400万円以上           → 4万円
    """
    if amount_yen < 1000000:
        return ""
    if amount_yen < 2000000:
        return "1万円"
    if amount_yen < 3000000:
        return "2万円"
    if amount_yen < 4000000:
        return "3万円"
    return "4万円"


# ============================================================
# メイン処理
# ============================================================

def main():

    # 1. 「集計前」シート読み込み
    df = pd.read_excel(INPUT_XLSX, sheet_name=INPUT_SHEET_NAME, dtype=str)

    # A〜D列が存在するか確認
    if df.shape[1] < 4:
        raise ValueError("「集計前」シートにA～D列が必要です。")

    colA, colB, colC, colD = df.columns[:4]

    # 2. B列に「取消」がある行を除外
    df = df.loc[~df[colB].fillna("").astype(str).str.contains("取消")].copy()

    # 3. D列の金額を円に変換
    df["__金額_yen__"] = df[colD].apply(parse_money_to_yen)

    # 4. A＋Cでグループ化して合計
    grouped = (
        df.groupby([colA, colC], dropna=False)["__金額_yen__"]
        .sum()
        .reset_index()
        .rename(columns={colA: "名前", colC: "メール", "__金額_yen__": "合計金額"})
        .sort_values(by="合計金額", ascending=False)
        .reset_index(drop=True)
    )

    # 5. 付与額判定
    grouped["付与額"] = grouped["合計金額"].apply(compute_tier_text)

    # 6. 付与額ごとの件数集計
    count_dict = {
        "1万円": int((grouped["付与額"] == "1万円").sum()),
        "2万円": int((grouped["付与額"] == "2万円").sum()),
        "3万円": int((grouped["付与額"] == "3万円").sum()),
        "4万円": int((grouped["付与額"] == "4万円").sum()),
    }

    # 7. Excelブックを開く
    wb = load_workbook(INPUT_XLSX)

    # 8. 既存の「集計後」シートがあれば削除
    if OUTPUT_SHEET_NAME in wb.sheetnames:
        del wb[OUTPUT_SHEET_NAME]

    # 9. 新しい「集計後」シート作成
    ws = wb.create_sheet(OUTPUT_SHEET_NAME, 0)

    # 10. 見出し
    ws["A1"] = "名前"
    ws["B1"] = "メール"
    ws["C1"] = "合計金額"
    ws["D1"] = "付与額"

    # 11. 集計結果を書き込み
    for row_number, row_data in enumerate(grouped.itertuples(index=False), start=2):
        ws.cell(row=row_number, column=1, value=row_data.名前)
        ws.cell(row=row_number, column=2, value=row_data.メール)
        ws.cell(row=row_number, column=3, value=row_data.合計金額)
        ws.cell(row=row_number, column=4, value=row_data.付与額)

    # 12. D列を付与額ごとに色分け
    for row_number in range(2, len(grouped) + 2):
        tier = ws.cell(row=row_number, column=4).value or ""
        fill_color = COLOR_MAP.get(tier, "FFFFFF")
        ws.cell(row=row_number, column=4).fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

    # 13. F-G列のサマリ見出し
    ws["F1"] = "付与額"
    ws["G1"] = "件数"

    # ★ 14. 付与額降順でサマリ出力
    tier_order = ["4万円", "3万円", "2万円", "1万円", ""]
    summary_sorted = [(tier, count_dict.get(tier, 0)) for tier in tier_order if count_dict.get(tier, 0) > 0]

    for i, (tier, cnt) in enumerate(summary_sorted, start=2):
        ws[f"F{i}"] = tier
        ws[f"G{i}"] = f"{cnt}件"
        ws[f"F{i}"].fill = PatternFill(start_color=COLOR_MAP[tier], end_color=COLOR_MAP[tier], fill_type="solid")

    # ★ 15. A, C, D, F, G 列のみ自動調整
    import unicodedata

    def get_display_width(text):
        width = 0
        for ch in text:
            # 全角文字は幅2、半角は幅1として計算
            if unicodedata.east_asian_width(ch) in ("F", "W", "A"):
                width += 2
            else:
                width += 1
        return width

    target_cols = ["A", "B", "C", "D", "F", "G"]

    for col_letter in target_cols:
        max_length = 0
        for cell in ws[col_letter]:
            if cell.value:
                for line in str(cell.value).split("\n"):
                    max_length = max(max_length, get_display_width(line))
        ws.column_dimensions[col_letter].width = max_length + 2


    # ★ 16. 罫線（A-D列、F-G列のみ）
    thin_side = Side(style="thin", color="000000")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # A〜D の最後の行（集計結果）
    last_row_AD = len(grouped) + 1

    # F〜G の最後の行（サマリ）
    last_row_FG = len(summary_sorted) + 1

    # A〜D に罫線
    for row in range(1, last_row_AD + 1):
        for col in ["A", "B", "C", "D"]:
            ws[f"{col}{row}"].border = border

    # F〜G に罫線
    for row in range(1, last_row_FG + 1):
        for col in ["F", "G"]:
            ws[f"{col}{row}"].border = border

    # 17. 見出し太字
    for col in ["A", "B", "C", "D", "F", "G"]:
        ws[f"{col}1"].font = ws[f"{col}1"].font.copy(bold=True)

    # 18. 保存
    wb.save(INPUT_XLSX)

    print("処理が完了しました。")


# ============================================================
# プログラム実行
# ============================================================

if __name__ == "__main__":
    main()
