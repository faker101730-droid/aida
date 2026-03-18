
# 改修版：中区分 / 小区分 切替対応

import io
from pathlib import Path
import pandas as pd
import streamlit as st

st.set_page_config(page_title="会計内訳アプリ Pro", layout="wide")

st.title("📘 会計内訳アプリ Pro（粒度切替版）")

# --------------------------
# 列マッピング（簡略）
# --------------------------
COL_MAP = {
    "日付": ["日付","伝票日付"],
    "借方中区分": ["借方中区分","借方科目（中区分）"],
    "貸方中区分": ["貸方中区分","貸方科目（中区分）"],
    "借方小区分": ["借方小区分","借方科目（小区分）"],
    "貸方小区分": ["貸方小区分","貸方科目（小区分）"],
    "借方金額": ["借方金額"],
    "貸方金額": ["貸方金額"],
    "摘要": ["摘要","摘要文"]
}

def normalize(df):
    df.columns = [str(c).replace(" ","").replace("　","") for c in df.columns]
    return df

def map_columns(df):
    df = normalize(df.copy())
    rename = {}
    for k, vals in COL_MAP.items():
        for v in vals:
            if v in df.columns:
                rename[v] = k
    return df.rename(columns=rename)

def load(file):
    if file.name.endswith(".csv"):
        return pd.read_csv(file, encoding="cp932")
    return pd.read_excel(file)

# --------------------------
# UI
# --------------------------
with st.sidebar:
    init_file = st.file_uploader("①初期残高")
    journal_file = st.file_uploader("②仕訳データ")
    mode = st.radio("集計粒度", ["中区分","小区分"])

# --------------------------
# 処理
# --------------------------
if init_file and journal_file:
    init = load(init_file)
    j = load(journal_file)
    j = map_columns(j)

    j["日付"] = pd.to_datetime(j["日付"])
    j["借方金額"] = pd.to_numeric(j["借方金額"], errors="coerce")
    j["貸方金額"] = pd.to_numeric(j["貸方金額"], errors="coerce")

    if mode == "中区分":
        j["借方科目"] = j["借方中区分"]
        j["貸方科目"] = j["貸方中区分"]
    else:
        j["借方科目"] = j["借方小区分"]
        j["貸方科目"] = j["貸方小区分"]

    account = st.selectbox("勘定科目", sorted(set(j["借方科目"]) | set(j["貸方科目"])))

    df = j[(j["借方科目"]==account)|(j["貸方科目"]==account)].copy()

    df["増減"] = df["借方金額"].fillna(0) - df["貸方金額"].fillna(0)

    st.subheader("仕訳明細")
    st.dataframe(df[["日付","借方科目","貸方科目","借方金額","貸方金額","摘要"]])

    st.subheader("合計")
    st.write(df["増減"].sum())

else:
    st.info("ファイルをアップロードしてください")
