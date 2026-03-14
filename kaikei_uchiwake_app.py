import io
from dataclasses import dataclass
from typing import List, Optional, Tuple

import pandas as pd
import streamlit as st

st.set_page_config(page_title="会計内訳生成アプリ", layout="wide")

BS_DEFAULT = [
    "現金", "普通預金", "当座預金", "定期預金",
    "売掛金", "未収入金", "立替金", "前払費用", "前払金", "仮払金",
    "商品", "貯蔵品", "建物", "建物附属設備", "構築物", "機械装置", "車両運搬具", "工具器具備品", "土地", "リース資産", "ソフトウェア", "差入保証金",
    "買掛金", "未払金", "未払費用", "未払法人税等", "未払消費税等", "預り金", "前受金", "前受収益", "仮受金", "短期借入金", "長期借入金", "賞与引当金", "退職給付引当金",
    "資本金", "資本剰余金", "利益剰余金", "繰越利益剰余金",
]

HELP_TEXT = """
### このアプリでできること
- 仕訳日記帳CSV/Excelを読み込む
- 貸借対照表に残る勘定科目だけ抽出する
- 科目ごとの残高内訳表を自動作成する
- 内訳表から仕訳明細までドリルダウンできる
- Excelで出力できる

### 想定データ
1行1仕訳、または1行1明細の形式を想定しています。
借方・貸方の勘定科目と金額が同じ行にある標準的な仕訳日記帳が最も扱いやすいです。
"""


@dataclass
class Cols:
    date: str
    debit_account: str
    credit_account: str
    amount: str
    partner: Optional[str]
    memo: Optional[str]
    voucher: Optional[str]


def normalize_amount(series: pd.Series) -> pd.Series:
    s = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("¥", "", regex=False)
        .str.strip()
    )
    s = s.replace({"": None, "nan": None, "None": None})
    return pd.to_numeric(s, errors="coerce").fillna(0)


def load_file(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        for enc in ["utf-8-sig", "cp932", "utf-8"]:
            try:
                uploaded_file.seek(0)
                return pd.read_csv(uploaded_file, encoding=enc)
            except Exception:
                pass
        uploaded_file.seek(0)
        return pd.read_csv(uploaded_file)

    uploaded_file.seek(0)
    return pd.read_excel(uploaded_file)


def build_ledger(df: pd.DataFrame, cols: Cols) -> pd.DataFrame:
    work = df.copy()
    work[cols.amount] = normalize_amount(work[cols.amount])
    if cols.date:
        work[cols.date] = pd.to_datetime(work[cols.date], errors="coerce")

    common = {}
    common["日付"] = work[cols.date] if cols.date else ""
    common["相手先"] = work[cols.partner].astype(str).replace("nan", "") if cols.partner else ""
    common["摘要"] = work[cols.memo].astype(str).replace("nan", "") if cols.memo else ""
    common["伝票番号"] = work[cols.voucher].astype(str).replace("nan", "") if cols.voucher else ""

    debit = pd.DataFrame(
        {
            **common,
            "勘定科目": work[cols.debit_account].astype(str).str.strip(),
            "借方": work[cols.amount],
            "貸方": 0,
            "増減": work[cols.amount],
            "元区分": "借方",
        }
    )
    credit = pd.DataFrame(
        {
            **common,
            "勘定科目": work[cols.credit_account].astype(str).str.strip(),
            "借方": 0,
            "貸方": work[cols.amount],
            "増減": -work[cols.amount],
            "元区分": "貸方",
        }
    )
    ledger = pd.concat([debit, credit], ignore_index=True)
    ledger = ledger[ledger["勘定科目"].notna() & (ledger["勘定科目"] != "")]
    return ledger


def summarize_balances(ledger: pd.DataFrame, bs_accounts: List[str]) -> pd.DataFrame:
    bs = ledger[ledger["勘定科目"].isin(bs_accounts)].copy()
    if bs.empty:
        return pd.DataFrame(columns=["勘定科目", "借方累計", "貸方累計", "残高"])

    summary = (
        bs.groupby("勘定科目", as_index=False)
        .agg(借方累計=("借方", "sum"), 貸方累計=("貸方", "sum"), 残高=("増減", "sum"))
    )
    summary = summary[summary["残高"].round(0) != 0].copy()
    summary = summary.sort_values("残高", key=lambda s: s.abs(), ascending=False).reset_index(drop=True)
    return summary


def breakdown_table(ledger: pd.DataFrame, account: str, key_col: str) -> pd.DataFrame:
    temp = ledger[ledger["勘定科目"] == account].copy()
    if temp.empty:
        return pd.DataFrame(columns=[key_col, "借方累計", "貸方累計", "残高"])

    if key_col not in temp.columns:
        temp[key_col] = ""
    temp[key_col] = temp[key_col].fillna("").replace("nan", "")
    temp[key_col] = temp[key_col].mask(temp[key_col].eq(""), "(空欄)")

    out = (
        temp.groupby(key_col, as_index=False)
        .agg(借方累計=("借方", "sum"), 貸方累計=("貸方", "sum"), 残高=("増減", "sum"))
    )
    out = out[out["残高"].round(0) != 0].copy()
    out = out.sort_values("残高", key=lambda s: s.abs(), ascending=False).reset_index(drop=True)
    return out


def to_excel_bytes(summary: pd.DataFrame, ledger: pd.DataFrame, key_col: str) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        summary.to_excel(writer, sheet_name="BS残高一覧", index=False)
        for acc in summary["勘定科目"].tolist():
            sheet_name = (acc or "sheet")[:31]
            breakdown_table(ledger, acc, key_col).to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return output.getvalue()


def infer_columns(df: pd.DataFrame) -> Tuple[Optional[str], Optional[str], Optional[str], Optional[str], Optional[str], Optional[str], Optional[str]]:
    cols = list(df.columns)

    def find(candidates):
        for cand in candidates:
            for col in cols:
                if cand.lower() in str(col).lower():
                    return col
        return None

    return (
        find(["日付", "date"]),
        find(["借方科目", "借方勘定", "debit"]),
        find(["貸方科目", "貸方勘定", "credit"]),
        find(["金額", "amount", "税込金額", "税抜金額"]),
        find(["相手先", "取引先", "補助", "partner", "vendor", "customer"]),
        find(["摘要", "内容", "memo", "description"]),
        find(["伝票", "voucher", "journal"]),
    )


st.title("会計内訳生成アプリ")
st.caption("仕訳日記帳から、B/Sに残っている勘定科目の内訳表を自動生成します。")

with st.expander("使い方", expanded=False):
    st.markdown(HELP_TEXT)

uploaded = st.file_uploader("仕訳日記帳ファイルをアップロード", type=["csv", "xlsx", "xls"])

if uploaded:
    try:
        df_raw = load_file(uploaded)
    except Exception as e:
        st.error(f"ファイル読込エラー: {e}")
        st.stop()

    st.subheader("アップロードデータ確認")
    st.dataframe(df_raw.head(20), use_container_width=True)

    date_c, debit_c, credit_c, amount_c, partner_c, memo_c, voucher_c = infer_columns(df_raw)
    columns = df_raw.columns.tolist()

    st.subheader("列マッピング")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        date_col = st.selectbox("日付列", options=[""] + columns, index=(columns.index(date_c) + 1) if date_c in columns else 0)
        debit_col = st.selectbox("借方科目列", options=[""] + columns, index=(columns.index(debit_c) + 1) if debit_c in columns else 0)
    with c2:
        credit_col = st.selectbox("貸方科目列", options=[""] + columns, index=(columns.index(credit_c) + 1) if credit_c in columns else 0)
        amount_col = st.selectbox("金額列", options=[""] + columns, index=(columns.index(amount_c) + 1) if amount_c in columns else 0)
    with c3:
        partner_col = st.selectbox("相手先列（任意）", options=[""] + columns, index=(columns.index(partner_c) + 1) if partner_c in columns else 0)
        memo_col = st.selectbox("摘要列（任意）", options=[""] + columns, index=(columns.index(memo_c) + 1) if memo_c in columns else 0)
    with c4:
        voucher_col = st.selectbox("伝票番号列（任意）", options=[""] + columns, index=(columns.index(voucher_c) + 1) if voucher_c in columns else 0)

    bs_input = st.text_area(
        "B/S対象勘定科目（改行区切り）",
        value="\n".join(BS_DEFAULT),
        height=220,
        help="ここにある科目だけをB/S残高一覧に抽出します。病院会計や施設の勘定体系に合わせて自由に編集してください。",
    )
    bs_accounts = [x.strip() for x in bs_input.splitlines() if x.strip()]

    breakdown_key = st.radio("内訳の集計軸", ["相手先", "摘要"], horizontal=True)

    if not debit_col or not credit_col or not amount_col:
        st.warning("借方科目列・貸方科目列・金額列は必須です。")
        st.stop()

    cols = Cols(
        date=date_col or "",
        debit_account=debit_col,
        credit_account=credit_col,
        amount=amount_col,
        partner=partner_col or None,
        memo=memo_col or None,
        voucher=voucher_col or None,
    )

    try:
        ledger = build_ledger(df_raw, cols)
        summary = summarize_balances(ledger, bs_accounts)
    except Exception as e:
        st.error(f"仕訳展開エラー: {e}")
        st.stop()

    st.subheader("B/S残高一覧")
    st.dataframe(summary, use_container_width=True)

    if summary.empty:
        st.info("B/S対象勘定科目に残高がありません。対象勘定科目の設定や列マッピングを確認してください。")
        st.stop()

    selected_account = st.selectbox("詳細を見る勘定科目", summary["勘定科目"].tolist())
    key_col = breakdown_key

    bd = breakdown_table(ledger, selected_account, key_col)
    st.subheader(f"{selected_account} の内訳表")
    st.dataframe(bd, use_container_width=True)

    st.subheader(f"{selected_account} の仕訳明細")
    detail = ledger[ledger["勘定科目"] == selected_account].copy()
    st.dataframe(detail.sort_values(["日付", "伝票番号"], na_position="last"), use_container_width=True)

    try:
        excel_bytes = to_excel_bytes(summary, ledger, key_col)
        st.download_button(
            "Excel出力",
            data=excel_bytes,
            file_name="会計内訳表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Excel出力エラー: {e}")
        st.info("requirements.txt に xlsxwriter と openpyxl が入っているか確認してください。")
else:
    st.info("まずは仕訳日記帳ファイルをアップロードしてください。")
