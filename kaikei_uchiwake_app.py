
import io
from pathlib import Path
import pandas as pd
import streamlit as st

st.set_page_config(page_title="会計内訳生成アプリ 完全版", page_icon="💼", layout="wide")

# ----------------------------
# Styling
# ----------------------------
st.markdown("""
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
.metric-card {
    background: linear-gradient(135deg, #132238 0%, #1b3554 100%);
    border-radius: 18px;
    padding: 16px 18px;
    color: white;
    border: 1px solid rgba(255,255,255,.08);
}
.metric-label {font-size: 0.85rem; opacity: .8; margin-bottom: 6px;}
.metric-value {font-size: 1.55rem; font-weight: 700;}
.soft-box {
    background: #f6f8fb;
    border: 1px solid #e5e9f0;
    border-radius: 16px;
    padding: 12px 14px;
}
.small-note {font-size: 0.85rem; color: #5b6574;}
.section-title {
    font-size: 1.1rem;
    font-weight: 700;
    margin: .3rem 0 .6rem;
}
</style>
""", unsafe_allow_html=True)

BS_DEFAULT = [
    "現金","普通預金","当座預金","定期預金",
    "売掛金","未収金","未収入金","立替金","前払費用","前払金","仮払金","貸付金","差入保証金",
    "商品","貯蔵品","建物","建物附属設備","構築物","機械装置","車両運搬具","工具器具備品","土地","リース資産","ソフトウェア",
    "買掛金","未払金","未払費用","未払法人税等","未払消費税等","預り金","前受金","前受収益","仮受金","短期借入金","長期借入金","賞与引当金","退職給付引当金",
    "資本金","資本剰余金","利益剰余金","繰越利益剰余金",
]

REQUIRED_INIT = ["基準日", "勘定科目", "相手先", "初期残高"]
REQUIRED_JOURNAL = ["日付", "借方科目", "貸方科目", "金額", "相手先", "摘要"]

# ----------------------------
# Helpers
# ----------------------------
def normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    x = df.copy()
    x.columns = [
        str(c).replace("\n", "").replace("\r", "").replace(" ", "").replace("　", "").strip()
        for c in x.columns
    ]
    return x

def normalize_text(s):
    return str(s).replace("\n", "").replace("\r", "").replace(" ", "").replace("　", "").strip()

def normalize_amount(series: pd.Series) -> pd.Series:
    s = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("¥", "", regex=False)
        .str.replace("円", "", regex=False)
        .str.replace("△", "-", regex=False)
        .str.strip()
    )
    s = s.replace({"": None, "nan": None, "None": None, "-": None})
    return pd.to_numeric(s, errors="coerce").fillna(0)

def read_csv_safely(file):
    for enc in ["utf-8-sig", "cp932", "utf-8", "shift_jis"]:
        try:
            file.seek(0)
            return pd.read_csv(file, encoding=enc)
        except Exception:
            continue
    file.seek(0)
    return pd.read_csv(file)

def detect_header_row_excel(file, expected_cols, max_rows=8):
    for header in range(max_rows):
        try:
            file.seek(0)
            df = pd.read_excel(file, header=header)
            df = normalize_headers(df)
            if set(expected_cols).issubset(set(df.columns)):
                return df
        except Exception:
            continue
    file.seek(0)
    return normalize_headers(pd.read_excel(file))

def load_any_table(file, expected_cols=None):
    name = file.name.lower()
    if name.endswith(".csv"):
        df = read_csv_safely(file)
        return normalize_headers(df)
    if expected_cols is None:
        file.seek(0)
        return normalize_headers(pd.read_excel(file))
    return detect_header_row_excel(file, expected_cols)

def detect_file_kind(df: pd.DataFrame):
    cols = set(normalize_headers(df).columns)
    if set(REQUIRED_INIT).issubset(cols):
        return "initial"
    if set(REQUIRED_JOURNAL).issubset(cols):
        return "journal"
    return "unknown"

def ensure_required(df, required, name):
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"{name}に必要列が不足しています：{', '.join(missing)}")

def build_ledger(journal_df: pd.DataFrame) -> pd.DataFrame:
    work = journal_df.copy()
    work = normalize_headers(work)
    ensure_required(work, REQUIRED_JOURNAL, "全期間仕訳マスタ")
    work["日付"] = pd.to_datetime(work["日付"], errors="coerce")
    work["金額"] = normalize_amount(work["金額"])
    for c in ["相手先", "摘要"]:
        work[c] = work[c].fillna("").astype(str).replace("nan", "")
    if "伝票番号" not in work.columns:
        work["伝票番号"] = ""

    debit = pd.DataFrame({
        "日付": work["日付"],
        "勘定科目": work["借方科目"].astype(str).str.strip(),
        "相手先": work["相手先"].astype(str).str.strip().replace("", "(空欄)"),
        "摘要": work["摘要"].astype(str).str.strip(),
        "伝票番号": work["伝票番号"].astype(str).str.strip(),
        "借方": work["金額"],
        "貸方": 0,
        "増減": work["金額"],
        "元区分": "借方",
    })
    credit = pd.DataFrame({
        "日付": work["日付"],
        "勘定科目": work["貸方科目"].astype(str).str.strip(),
        "相手先": work["相手先"].astype(str).str.strip().replace("", "(空欄)"),
        "摘要": work["摘要"].astype(str).str.strip(),
        "伝票番号": work["伝票番号"].astype(str).str.strip(),
        "借方": 0,
        "貸方": work["金額"],
        "増減": -work["金額"],
        "元区分": "貸方",
    })
    ledger = pd.concat([debit, credit], ignore_index=True)
    ledger["相手先"] = ledger["相手先"].replace("", "(空欄)")
    ledger = ledger[ledger["勘定科目"].notna() & (ledger["勘定科目"] != "")]
    return ledger

def prep_initial(initial_df: pd.DataFrame) -> pd.DataFrame:
    x = normalize_headers(initial_df)
    ensure_required(x, REQUIRED_INIT, "初期残高データ")
    x["基準日"] = pd.to_datetime(x["基準日"], errors="coerce")
    x["勘定科目"] = x["勘定科目"].astype(str).str.strip()
    x["相手先"] = x["相手先"].fillna("").astype(str).str.strip().replace("", "(空欄)")
    x["初期残高"] = normalize_amount(x["初期残高"])
    if "摘要" not in x.columns:
        x["摘要"] = ""
    return x

def month_range(target_month):
    start = pd.Timestamp(target_month).to_period("M").start_time
    end = pd.Timestamp(target_month).to_period("M").end_time
    return start.normalize(), end.normalize()

def account_month_summary(init_df, ledger_df, bs_accounts, target_month):
    start, end = month_range(target_month)
    init_bs = init_df[init_df["勘定科目"].isin(bs_accounts)].copy()
    led_bs = ledger_df[ledger_df["勘定科目"].isin(bs_accounts)].copy()

    prior = led_bs[led_bs["日付"] < start]
    current = led_bs[(led_bs["日付"] >= start) & (led_bs["日付"] <= end)]

    init_sum = init_bs.groupby("勘定科目", as_index=False)["初期残高"].sum().rename(columns={"初期残高": "初期残高"})
    prior_sum = prior.groupby("勘定科目", as_index=False)["増減"].sum().rename(columns={"増減": "期首前累計"})
    curr_inc = current.assign(当月増加=current["増減"].clip(lower=0)).groupby("勘定科目", as_index=False)["当月増加"].sum()
    curr_dec = current.assign(当月減少=(-current["増減"].clip(upper=0))).groupby("勘定科目", as_index=False)["当月減少"].sum()

    accounts = pd.DataFrame({"勘定科目": sorted(set(init_bs["勘定科目"]).union(set(led_bs["勘定科目"])))})
    out = accounts.merge(init_sum, on="勘定科目", how="left").merge(prior_sum, on="勘定科目", how="left").merge(curr_inc, on="勘定科目", how="left").merge(curr_dec, on="勘定科目", how="left")
    for col in ["初期残高", "期首前累計", "当月増加", "当月減少"]:
        out[col] = out[col].fillna(0)
    out["期首残高"] = out["初期残高"] + out["期首前累計"]
    out["当月増減"] = out["当月増加"] - out["当月減少"]
    out["期末残高"] = out["期首残高"] + out["当月増減"]
    out = out[(out["期首残高"].round(0) != 0) | (out["当月増加"].round(0) != 0) | (out["当月減少"].round(0) != 0) | (out["期末残高"].round(0) != 0)].copy()
    out = out.sort_values("期末残高", key=lambda s: s.abs(), ascending=False).reset_index(drop=True)
    return out

def partner_breakdown(init_df, ledger_df, account, target_month):
    start, end = month_range(target_month)
    init_acc = init_df[init_df["勘定科目"] == account].copy()
    prior = ledger_df[(ledger_df["勘定科目"] == account) & (ledger_df["日付"] < start)].copy()
    current = ledger_df[(ledger_df["勘定科目"] == account) & (ledger_df["日付"] >= start) & (ledger_df["日付"] <= end)].copy()

    init_sum = init_acc.groupby("相手先", as_index=False)["初期残高"].sum()
    prior_sum = prior.groupby("相手先", as_index=False)["増減"].sum().rename(columns={"増減": "期首前累計"})
    curr_inc = current.assign(当月増加=current["増減"].clip(lower=0)).groupby("相手先", as_index=False)["当月増加"].sum()
    curr_dec = current.assign(当月減少=(-current["増減"].clip(upper=0))).groupby("相手先", as_index=False)["当月減少"].sum()

    partners = pd.DataFrame({"相手先": sorted(set(init_acc["相手先"]).union(set(prior["相手先"])).union(set(current["相手先"])))})
    if partners.empty:
        return pd.DataFrame(columns=["相手先","期首残高","当月増加","当月減少","当月増減","期末残高"])
    out = partners.merge(init_sum, on="相手先", how="left").merge(prior_sum, on="相手先", how="left").merge(curr_inc, on="相手先", how="left").merge(curr_dec, on="相手先", how="left")
    for col in ["初期残高", "期首前累計", "当月増加", "当月減少"]:
        out[col] = out[col].fillna(0)
    out["期首残高"] = out["初期残高"] + out["期首前累計"]
    out["当月増減"] = out["当月増加"] - out["当月減少"]
    out["期末残高"] = out["期首残高"] + out["当月増減"]
    out = out[(out["期首残高"].round(0) != 0) | (out["当月増加"].round(0) != 0) | (out["当月減少"].round(0) != 0) | (out["期末残高"].round(0) != 0)].copy()
    out["異常フラグ"] = ""
    out.loc[out["期末残高"] < 0, "異常フラグ"] = "マイナス残高"
    out = out.sort_values("期末残高", key=lambda s: s.abs(), ascending=False).reset_index(drop=True)
    return out

def details_for_account(ledger_df, account, target_month):
    start, end = month_range(target_month)
    prior = ledger_df[(ledger_df["勘定科目"] == account) & (ledger_df["日付"] < start)].copy()
    current = ledger_df[(ledger_df["勘定科目"] == account) & (ledger_df["日付"] >= start) & (ledger_df["日付"] <= end)].copy()
    prior = prior.sort_values(["日付","相手先","摘要"])
    current = current.sort_values(["日付","相手先","摘要"])
    return prior, current

def format_df_for_display(df):
    x = df.copy()
    num_cols = x.select_dtypes(include="number").columns
    for c in num_cols:
        x[c] = x[c].round(0).astype(int)
    return x

def to_excel_bytes(summary_df, partner_df, prior_df, current_df, target_month, account):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="勘定科目残高一覧", index=False)
        partner_df.to_excel(writer, sheet_name="相手先別内訳", index=False)
        prior_df.to_excel(writer, sheet_name="期首算出用履歴", index=False)
        current_df.to_excel(writer, sheet_name="当月仕訳明細", index=False)
    output.seek(0)
    return output.getvalue()

# ----------------------------
# UI
# ----------------------------
st.title("💼 会計内訳生成アプリ 完全版")
st.caption("初期残高＋全期間仕訳マスタから、対象月の期首残高・当月増減・期末残高を自動表示します。")

left, right = st.columns([1, 2], gap="large")

with left:
    st.markdown('<div class="section-title">📂 データ読み込み</div>', unsafe_allow_html=True)
    init_file = st.file_uploader("① 初期残高データ", type=["csv", "xlsx", "xls"], key="init")
    journal_file = st.file_uploader("② 全期間仕訳マスタ", type=["csv", "xlsx", "xls"], key="journal")

    st.markdown(
        '<div class="soft-box small-note">推奨列名<br>初期残高：基準日 / 勘定科目 / 相手先 / 初期残高<br>仕訳マスタ：日付 / 借方科目 / 貸方科目 / 金額 / 相手先 / 摘要</div>',
        unsafe_allow_html=True
    )

    bs_text = st.text_area("B/S対象勘定科目（改行区切り）", value="\n".join(BS_DEFAULT), height=220)
    bs_accounts = [x.strip() for x in bs_text.splitlines() if x.strip()]

if init_file and journal_file:
    try:
        init_df_raw = load_any_table(init_file, expected_cols=REQUIRED_INIT)
        journal_df_raw = load_any_table(journal_file, expected_cols=REQUIRED_JOURNAL)

        kind1 = detect_file_kind(init_df_raw)
        kind2 = detect_file_kind(journal_df_raw)

        if kind1 == "journal" and kind2 == "initial":
            init_df_raw, journal_df_raw = journal_df_raw, init_df_raw
            st.info("アップロード順を自動で補正しました。")
        elif kind1 == "unknown" or kind2 == "unknown":
            st.warning("列名が想定と異なるファイルがあります。テンプレート列名に合わせてください。")

        init_df = prep_initial(init_df_raw)
        ledger_df = build_ledger(journal_df_raw)

        if ledger_df["日付"].notna().sum() == 0:
            st.error("仕訳マスタの日付を認識できませんでした。日付列を確認してください。")
            st.stop()

        available_months = sorted(ledger_df["日付"].dropna().dt.to_period("M").astype(str).unique().tolist())
        if not available_months:
            st.error("対象月を作れませんでした。仕訳マスタの日付列を確認してください。")
            st.stop()

        with right:
            ctl1, ctl2 = st.columns([1,1])
            with ctl1:
                target_month = st.selectbox("対象月", options=available_months, index=len(available_months)-1)
            summary_df = account_month_summary(init_df, ledger_df, bs_accounts, target_month)
            if summary_df.empty:
                st.warning("対象月に表示対象となるB/S科目がありません。")
                st.stop()
            with ctl2:
                selected_account = st.selectbox("勘定科目", options=summary_df["勘定科目"].tolist())

            selected_row = summary_df[summary_df["勘定科目"] == selected_account].iloc[0]
            partner_df = partner_breakdown(init_df, ledger_df, selected_account, target_month)
            prior_df, current_df = details_for_account(ledger_df, selected_account, target_month)

            st.markdown(
                f"""
                <div class="small-note">
                初期残高基準日: {init_df["基準日"].dropna().max().date() if init_df["基準日"].notna().any() else "-"}　
                ／　対象月: {target_month}
                </div>
                """,
                unsafe_allow_html=True
            )

            m1, m2, m3, m4 = st.columns(4)
            metrics = [
                ("期首残高", selected_row["期首残高"]),
                ("当月増加", selected_row["当月増加"]),
                ("当月減少", selected_row["当月減少"]),
                ("期末残高", selected_row["期末残高"]),
            ]
            for col, (label, val) in zip([m1,m2,m3,m4], metrics):
                with col:
                    st.markdown(
                        f'<div class="metric-card"><div class="metric-label">{label}</div><div class="metric-value">{val:,.0f}</div></div>',
                        unsafe_allow_html=True
                    )

            tab1, tab2, tab3, tab4 = st.tabs(["📊 勘定科目残高一覧", "📋 相手先別内訳", "🔎 明細ドリルダウン", "📥 Excel出力"])

            with tab1:
                show_summary = format_df_for_display(summary_df[["勘定科目","期首残高","当月増加","当月減少","当月増減","期末残高"]])
                st.dataframe(show_summary, use_container_width=True, hide_index=True)
                st.caption("対象月の勘定科目別残高一覧です。期首残高・当月増減・期末残高を表示します。")

            with tab2:
                show_partner = format_df_for_display(partner_df[["相手先","期首残高","当月増加","当月減少","当月増減","期末残高","異常フラグ"]]) if not partner_df.empty else partner_df
                st.dataframe(show_partner, use_container_width=True, hide_index=True)
                st.caption("選択した勘定科目の相手先別内訳です。マイナス残高は異常フラグを表示します。")

            with tab3:
                a, b = st.columns(2)
                with a:
                    st.markdown("**期首算出用の過去履歴**")
                    st.dataframe(format_df_for_display(prior_df), use_container_width=True, hide_index=True)
                with b:
                    st.markdown("**当月仕訳明細**")
                    st.dataframe(format_df_for_display(current_df), use_container_width=True, hide_index=True)

            with tab4:
                excel_bytes = to_excel_bytes(
                    format_df_for_display(summary_df[["勘定科目","期首残高","当月増加","当月減少","当月増減","期末残高"]]),
                    format_df_for_display(partner_df),
                    format_df_for_display(prior_df),
                    format_df_for_display(current_df),
                    target_month,
                    selected_account
                )
                st.download_button(
                    "Excelをダウンロード",
                    data=excel_bytes,
                    file_name=f"会計内訳_{target_month}_{selected_account}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with st.expander("プレビュー（読み込みデータ確認）", expanded=False):
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("**初期残高**")
                    st.dataframe(init_df.head(20), use_container_width=True)
                with c2:
                    st.markdown("**仕訳マスタ（先頭20行）**")
                    st.dataframe(journal_df_raw.head(20), use_container_width=True)

    except Exception as e:
        st.error(f"読み込みエラー：{e}")
else:
    with right:
        st.info("左側から初期残高データと全期間仕訳マスタをアップロードしてください。")
        st.markdown("""
        #### この完全版でできること
        - 初期残高＋全期間仕訳マスタによる繰越計算
        - 対象月の期首残高 / 当月増加 / 当月減少 / 期末残高の表示
        - 相手先別の内訳表
        - 期首算出用履歴と当月仕訳明細のドリルダウン
        - Excel出力
        """)
