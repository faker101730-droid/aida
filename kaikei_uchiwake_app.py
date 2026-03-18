
import io
from pathlib import Path
import pandas as pd
import streamlit as st

st.set_page_config(page_title="会計内訳アプリ Pro", layout="wide")
st.title("📘 会計内訳アプリ Pro")

COL_MAP = {
    "日付": ["日付","伝票日付"],
    "借方中区分": ["借方中区分","借方科目（中区分）","借方科目(中区分)"],
    "貸方中区分": ["貸方中区分","貸方科目（中区分）","貸方科目(中区分)"],
    "借方小区分": ["借方小区分","借方科目（小区分）","借方科目(小区分)"],
    "貸方小区分": ["貸方小区分","貸方科目（小区分）","貸方科目(小区分)"],
    "借方金額": ["借方金額"],
    "貸方金額": ["貸方金額"],
    "摘要": ["摘要","摘要文"]
}

def normalize(df):
    df = df.copy()
    df.columns = [str(c).replace(" ","").replace("　","").strip() for c in df.columns]
    return df

def map_columns(df):
    df = normalize(df)
    rename = {}
    for k, vals in COL_MAP.items():
        for v in vals:
            if v in df.columns:
                rename[v] = k
                break
    return df.rename(columns=rename)

def load(file):
    suffix = Path(file.name).suffix.lower()
    if suffix == ".csv":
        for enc in ["utf-8-sig", "cp932", "utf-8", "shift_jis"]:
            try:
                file.seek(0)
                return pd.read_csv(file, encoding=enc)
            except Exception:
                continue
        file.seek(0)
        return pd.read_csv(file)
    file.seek(0)
    return pd.read_excel(file)

def clean_text_series(s):
    return s.fillna("").astype(str).str.strip()

with st.sidebar:
    init_file = st.file_uploader("①初期残高", type=["csv","xlsx","xls"])
    journal_file = st.file_uploader("②仕訳データ", type=["csv","xlsx","xls"])
    mode = st.radio("集計粒度", ["中区分","小区分"])

if init_file and journal_file:
    try:
        init = load(init_file)
        j = map_columns(load(journal_file))

        required = ["日付","借方金額","貸方金額"]
        miss = [c for c in required if c not in j.columns]
        if miss:
            st.error("仕訳データに必要列がありません: " + ", ".join(miss))
            st.stop()

        j["日付"] = pd.to_datetime(j["日付"], errors="coerce")
        j["借方金額"] = pd.to_numeric(j["借方金額"].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)
        j["貸方金額"] = pd.to_numeric(j["貸方金額"].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)

        if mode == "中区分":
            if "借方中区分" not in j.columns or "貸方中区分" not in j.columns:
                st.error("中区分列が見つかりません。")
                st.stop()
            j["借方科目"] = clean_text_series(j["借方中区分"])
            j["貸方科目"] = clean_text_series(j["貸方中区分"])
        else:
            if "借方小区分" not in j.columns or "貸方小区分" not in j.columns:
                st.error("小区分列が見つかりません。")
                st.stop()
            j["借方科目"] = clean_text_series(j["借方小区分"])
            j["貸方科目"] = clean_text_series(j["貸方小区分"])

        if "摘要" not in j.columns:
            j["摘要"] = ""
        j["摘要"] = clean_text_series(j["摘要"])

        account_list = sorted(set(j["借方科目"]) | set(j["貸方科目"]))
        account_list = [x for x in account_list if x not in ["", "nan", "None"]]

        if not account_list:
            st.error("勘定科目候補が作れませんでした。")
            st.stop()

        account = st.selectbox("勘定科目", account_list)

        df = j[(j["借方科目"] == account) | (j["貸方科目"] == account)].copy()
        df["増減"] = df["借方金額"] - df["貸方金額"]

        show_cols = [c for c in ["日付","借方科目","貸方科目","借方金額","貸方金額","摘要"] if c in df.columns]
        show = df[show_cols + ["増減"]].copy()

        for c in ["借方金額","貸方金額","増減"]:
            if c in show.columns:
                show[c] = show[c].apply(lambda x: f"{int(round(x)):,}")

        st.subheader("仕訳明細")
        st.dataframe(show, use_container_width=True, hide_index=True)

        total = df["増減"].sum()
        st.subheader("合計")
        st.write(f"{int(round(total)):,}")

    except Exception as e:
        st.error(f"読み込みエラー: {e}")
else:
    st.info("ファイルをアップロードしてください")
