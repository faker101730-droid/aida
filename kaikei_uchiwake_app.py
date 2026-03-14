
import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="会計内訳生成アプリ（実務版デモ）", layout="wide")

st.title("会計内訳生成アプリ（デモ）")
st.write("仕訳日記帳CSVから勘定科目残高と内訳表を自動作成します。")

uploaded = st.file_uploader("仕訳CSVをアップロード", type=["csv"])

def build_ledger(df):
    rows = []
    for _, r in df.iterrows():
        rows.append({"科目": r["借方科目"], "金額": r["金額"], "相手先": r["相手先"], "摘要": r["摘要"]})
        rows.append({"科目": r["貸方科目"], "金額": -r["金額"], "相手先": r["相手先"], "摘要": r["摘要"]})
    ledger = pd.DataFrame(rows)
    return ledger

if uploaded:
    df = pd.read_csv(uploaded)
    st.subheader("仕訳データ")
    st.dataframe(df, use_container_width=True)

    ledger = build_ledger(df)

    st.subheader("勘定科目残高")
    balance = ledger.groupby("科目")["金額"].sum().reset_index().sort_values("金額", ascending=False)
    st.dataframe(balance, use_container_width=True)

    st.subheader("科目内訳")
    account = st.selectbox("勘定科目を選択", balance["科目"].unique())

    detail = ledger[ledger["科目"] == account]
    by_partner = detail.groupby("相手先")["金額"].sum().reset_index()

    st.write("相手先別内訳")
    st.dataframe(by_partner, use_container_width=True)

    st.write("仕訳明細")
    st.dataframe(detail, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        balance.to_excel(writer, sheet_name="残高", index=False)
        by_partner.to_excel(writer, sheet_name="内訳", index=False)
        detail.to_excel(writer, sheet_name="仕訳", index=False)

    st.download_button(
        label="Excelダウンロード",
        data=output.getvalue(),
        file_name="account_detail.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
