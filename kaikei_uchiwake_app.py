import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="会計内訳アプリ Pro", page_icon="📘", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 1.5rem;}
.main-title {font-size: 2rem; font-weight: 800; color: #143a52; margin-bottom: 0.2rem;}
.sub-title {color: #4b6475; margin-bottom: 1rem;}
.section-head {font-size: 1.1rem; font-weight: 700; color: #143a52; margin-top: 0.5rem; margin-bottom: 0.4rem;}
div[data-testid="stMetric"] {background: linear-gradient(135deg, #f8fbff 0%, #eef6fb 100%); border: 1px solid #d8e6f0; padding: 10px 14px; border-radius: 16px;}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">会計内訳アプリ Pro</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">初期残高＋全期間仕訳マスタから、対象月の期首残高・当月増減・期末残高を自動表示</div>', unsafe_allow_html=True)

INITIAL_REQUIRED = ["基準日", "勘定科目", "相手先", "初期残高"]
JOURNAL_REQUIRED = ["日付", "借方科目", "貸方科目", "金額", "相手先", "摘要"]


def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [str(c).replace("\n", " ").replace("\r", " ").strip() for c in out.columns]
    return out


def normalize_text(v):
    if pd.isna(v):
        return ""
    return str(v).replace("\n", " ").replace("\r", " ").strip()


def rename_by_map(df: pd.DataFrame, colmap: dict) -> pd.DataFrame:
    df = clean_columns(df)
    renamed = {}
    for c in df.columns:
        key = normalize_text(c)
        renamed[c] = colmap.get(key, key)
    return df.rename(columns=renamed)


def detect_header_row(df: pd.DataFrame, expected_cols: list[str]) -> int:
    # 既定は0行目見出し。ダメなら先頭5行から探す
    score_best = -1
    best_idx = 0
    check_limit = min(len(df), 5)
    expected_set = set(expected_cols)
    for i in range(check_limit):
        row_vals = {normalize_text(v) for v in df.iloc[i].tolist()}
        score = len(expected_set & row_vals)
        if score > score_best:
            score_best = score
            best_idx = i
    return best_idx


def read_excel_smart(uploaded_file) -> pd.DataFrame:
    xls = pd.ExcelFile(uploaded_file)
    best_df = None
    best_score = -1
    for sheet in xls.sheet_names:
        raw = pd.read_excel(xls, sheet_name=sheet, header=None)
        for expected in [JOURNAL_REQUIRED, INITIAL_REQUIRED]:
            idx = detect_header_row(raw, expected)
            header = [normalize_text(v) for v in raw.iloc[idx].tolist()]
            data = raw.iloc[idx+1:].copy()
            data.columns = header
            data = data.dropna(how="all").reset_index(drop=True)
            score = len(set(expected) & set(data.columns))
            if score > best_score:
                best_df = data
                best_score = score
    if best_df is None:
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file)
    return clean_columns(best_df)


def read_csv_smart(uploaded_file) -> pd.DataFrame:
    raw = uploaded_file.getvalue()
    for enc in ["utf-8-sig", "cp932", "utf-8"]:
        try:
            return clean_columns(pd.read_csv(io.BytesIO(raw), encoding=enc))
        except Exception:
            pass
    raise ValueError("CSVの読み込みに失敗しました。UTF-8 BOM付き または CP932 を推奨します。")


def read_any(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith((".xlsx", ".xlsm", ".xls")):
        uploaded_file.seek(0)
        return read_excel_smart(uploaded_file)
    if name.endswith(".csv"):
        uploaded_file.seek(0)
        return read_csv_smart(uploaded_file)
    raise ValueError("対応ファイルは CSV / Excel のみです。")


def normalize_initial(df):
    colmap = {
        "勘定科目":"勘定科目", "科目":"勘定科目", "account":"勘定科目",
        "相手先":"相手先", "補助科目":"相手先", "partner":"相手先",
        "初期残高":"初期残高", "前月末残高":"初期残高", "残高":"初期残高", "opening_balance":"初期残高",
        "基準日":"基準日", "日付":"基準日", "date":"基準日",
    }
    df = rename_by_map(df, colmap)
    missing = [c for c in INITIAL_REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"初期残高データに必要列が不足しています: {', '.join(missing)}")
    df = df[INITIAL_REQUIRED].copy()
    df["基準日"] = pd.to_datetime(df["基準日"], errors="coerce")
    df["相手先"] = df["相手先"].fillna("（相手先未設定）").astype(str).str.strip()
    df["勘定科目"] = df["勘定科目"].astype(str).str.strip()
    df["初期残高"] = pd.to_numeric(df["初期残高"].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)
    df = df.dropna(subset=["基準日"])
    return df


def normalize_journal(df):
    colmap = {
        "日付":"日付", "date":"日付",
        "借方科目":"借方科目", "借方勘定":"借方科目", "debit_account":"借方科目",
        "貸方科目":"貸方科目", "貸方勘定":"貸方科目", "credit_account":"貸方科目",
        "金額":"金額", "amount":"金額",
        "相手先":"相手先", "補助科目":"相手先", "partner":"相手先",
        "摘要":"摘要", "description":"摘要",
        "仕訳ID":"仕訳ID", "id":"仕訳ID",
    }
    df = rename_by_map(df, colmap)
    missing = [c for c in JOURNAL_REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"仕訳マスタに必要列が不足しています: {', '.join(missing)}")
    work = df.copy()
    work["日付"] = pd.to_datetime(work["日付"], errors="coerce")
    work["金額"] = pd.to_numeric(work["金額"].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)
    work["相手先"] = work["相手先"].fillna("（相手先未設定）").astype(str).str.strip()
    work["摘要"] = work["摘要"].fillna("").astype(str).str.strip()
    work["借方科目"] = work["借方科目"].astype(str).str.strip()
    work["貸方科目"] = work["貸方科目"].astype(str).str.strip()
    if "仕訳ID" not in work.columns:
        work["仕訳ID"] = range(1, len(work) + 1)
    work = work.dropna(subset=["日付"])
    return work


def auto_classify(initial_df: pd.DataFrame | None, journal_df: pd.DataFrame | None):
    # 両方埋まっていても誤投入のときは入れ替える
    if initial_df is not None and journal_df is not None:
        init_cols = set(initial_df.columns)
        jour_cols = set(journal_df.columns)
        init_journal_score = len(set(JOURNAL_REQUIRED) & init_cols)
        init_initial_score = len(set(INITIAL_REQUIRED) & init_cols)
        jour_journal_score = len(set(JOURNAL_REQUIRED) & jour_cols)
        jour_initial_score = len(set(INITIAL_REQUIRED) & jour_cols)
        if init_journal_score > init_initial_score and jour_initial_score > jour_journal_score:
            return journal_df, initial_df, True
    return initial_df, journal_df, False


def build_movements(journal_df):
    debit = journal_df[["日付", "借方科目", "金額", "相手先", "摘要", "仕訳ID"]].copy()
    debit.columns = ["日付", "勘定科目", "金額", "相手先", "摘要", "仕訳ID"]
    debit["符号区分"] = "借方"
    debit["増加額"] = debit["金額"].clip(lower=0)
    debit["減少額"] = 0

    credit = journal_df[["日付", "貸方科目", "金額", "相手先", "摘要", "仕訳ID"]].copy()
    credit.columns = ["日付", "勘定科目", "金額", "相手先", "摘要", "仕訳ID"]
    credit["金額"] = -credit["金額"]
    credit["符号区分"] = "貸方"
    credit["増加額"] = 0
    credit["減少額"] = (-credit["金額"]).clip(lower=0)

    return pd.concat([debit, credit], ignore_index=True)


def month_bounds(month_period):
    start = month_period.to_timestamp()
    end = (month_period + 1).to_timestamp()
    return start, end


def fmt_amount(x):
    try:
        x = float(x)
    except Exception:
        return "-"
    if abs(x) < 0.000001:
        return "-"
    return f"{x:,.0f}"


def make_export(selected_month_label, all_balance, partner_summary, month_detail, pre_opening_detail):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        all_balance.to_excel(writer, sheet_name="月末残高一覧", index=False)
        partner_summary.to_excel(writer, sheet_name="相手先別内訳", index=False)
        month_detail.to_excel(writer, sheet_name="当月仕訳明細", index=False)
        pre_opening_detail.to_excel(writer, sheet_name="期首算出用履歴", index=False)
    output.seek(0)
    return output.getvalue()


with st.sidebar:
    st.header("📂 データ読み込み")
    initial_file = st.file_uploader("① 初期残高データ", type=["csv", "xlsx", "xls"], key="initial")
    journal_file = st.file_uploader("② 全期間仕訳マスタ", type=["csv", "xlsx", "xls"], key="journal")
    st.caption("推奨列名")
    st.code("初期残高: 基準日 / 勘定科目 / 相手先 / 初期残高\n仕訳: 日付 / 借方科目 / 貸方科目 / 金額 / 相手先 / 摘要 / 仕訳ID", language="text")
    st.caption("※ 先頭に説明行が1行あるExcelでも自動判定します。誤って逆にアップした場合も自動で入れ替えます。")

if not journal_file and not initial_file:
    st.info("まずはファイルをアップロードしてください。初期残高は未設定でも動きますが、期首残高は0始まりになります。")
    st.stop()

try:
    initial_raw = read_any(initial_file) if initial_file else None
    journal_raw = read_any(journal_file) if journal_file else None
    initial_raw, journal_raw, swapped = auto_classify(initial_raw, journal_raw)
    if journal_raw is None:
        raise ValueError("全期間仕訳マスタが未アップロードです。")
    journal = normalize_journal(journal_raw)
    initial = normalize_initial(initial_raw) if initial_raw is not None else None
except Exception as e:
    st.error(f"読み込みエラー: {e}")
    st.stop()

if swapped:
    st.warning("アップロードされた2ファイルの中身を判定し、自動で『初期残高データ』と『全期間仕訳マスタ』を入れ替えて処理しました。")

movements = build_movements(journal)

if initial is None:
    min_date = journal["日付"].min()
    initial = pd.DataFrame({
        "基準日": [min_date - pd.Timedelta(days=1)],
        "勘定科目": ["初期設定なし"],
        "相手先": ["（相手先未設定）"],
        "初期残高": [0]
    })
    no_initial_flag = True
else:
    no_initial_flag = False

initial_base_date = initial["基準日"].max()
months = sorted(journal["日付"].dt.to_period("M").unique())
selected_period = st.selectbox("対象月", months, index=len(months)-1, format_func=lambda p: p.strftime("%Y-%m"))
account_options = sorted(pd.unique(pd.concat([movements["勘定科目"], initial["勘定科目"]]).astype(str)))
selected_account = st.selectbox("勘定科目", account_options)
month_start, next_month_start = month_bounds(selected_period)
selected_month_label = selected_period.strftime("%Y-%m")

initial_clean = initial[initial["勘定科目"] != "初期設定なし"].copy()
opening_initial = initial_clean.groupby(["勘定科目", "相手先"], as_index=False)["初期残高"].sum()
pre_month_movements = movements[movements["日付"] < month_start].groupby(["勘定科目", "相手先"], as_index=False)["金額"].sum().rename(columns={"金額": "前月まで累計増減"})
opening_partner = opening_initial.merge(pre_month_movements, on=["勘定科目", "相手先"], how="outer")
opening_partner["初期残高"] = opening_partner["初期残高"].fillna(0)
opening_partner["前月まで累計増減"] = opening_partner["前月まで累計増減"].fillna(0)
opening_partner["期首残高"] = opening_partner["初期残高"] + opening_partner["前月まで累計増減"]
month_movement_detail = movements[(movements["日付"] >= month_start) & (movements["日付"] < next_month_start)].copy()
month_partner = month_movement_detail.groupby(["勘定科目", "相手先"], as_index=False).agg(当月増加=("増加額", "sum"), 当月減少=("減少額", "sum"), 当月増減=("金額", "sum"))
partner_summary = opening_partner.merge(month_partner, on=["勘定科目", "相手先"], how="outer")
for col in ["初期残高", "前月まで累計増減", "期首残高", "当月増加", "当月減少", "当月増減"]:
    if col in partner_summary.columns:
        partner_summary[col] = partner_summary[col].fillna(0)
partner_summary["期末残高"] = partner_summary["期首残高"] + partner_summary["当月増減"]
all_balance = partner_summary.groupby("勘定科目", as_index=False).agg(期首残高=("期首残高", "sum"), 当月増加=("当月増加", "sum"), 当月減少=("当月減少", "sum"), 当月増減=("当月増減", "sum"), 期末残高=("期末残高", "sum")).sort_values("期末残高", ascending=False)
account_partner = partner_summary[partner_summary["勘定科目"] == selected_account].copy().sort_values("期末残高", ascending=False)
pre_opening_detail = movements[(movements["日付"] < month_start) & (movements["勘定科目"] == selected_account)].copy()
month_detail = month_movement_detail[month_movement_detail["勘定科目"] == selected_account].copy().sort_values(["日付", "仕訳ID"])
account_total = all_balance[all_balance["勘定科目"] == selected_account]
if len(account_total) == 0:
    account_total = pd.DataFrame([{"勘定科目": selected_account, "期首残高": 0, "当月増加": 0, "当月減少": 0, "当月増減": 0, "期末残高": 0}])
account_total = account_total.iloc[0]

note_cols = st.columns([1.2, 1.8])
with note_cols[0]:
    if no_initial_flag:
        st.warning("初期残高データ未設定のため、期首残高は0起点で計算しています。")
    else:
        st.success(f"初期残高基準日: {initial_base_date.strftime('%Y-%m-%d')}")
with note_cols[1]:
    st.info(f"対象月は {selected_month_label}。この月の期首残高・当月増減・期末残高を表示しています。")

m1, m2, m3, m4 = st.columns(4)
m1.metric("期首残高", fmt_amount(account_total["期首残高"]))
m2.metric("当月増加", fmt_amount(account_total["当月増加"]))
m3.metric("当月減少", fmt_amount(account_total["当月減少"]))
m4.metric("期末残高", fmt_amount(account_total["期末残高"]))

tab1, tab2, tab3, tab4 = st.tabs(["📊 残高一覧", "🧾 相手先別内訳", "🔎 明細ドリルダウン", "⬇️ Excel出力"])
with tab1:
    st.markdown('<div class="section-head">対象月の勘定科目別残高一覧</div>', unsafe_allow_html=True)
    st.dataframe(all_balance, use_container_width=True, height=420)
with tab2:
    st.markdown(f'<div class="section-head">「{selected_account}」の相手先別内訳</div>', unsafe_allow_html=True)
    show_cols = ["相手先", "期首残高", "当月増加", "当月減少", "当月増減", "期末残高"]
    st.dataframe(account_partner[show_cols] if len(account_partner) else pd.DataFrame(columns=show_cols), use_container_width=True, height=420)
with tab3:
    c_left, c_right = st.columns(2)
    with c_left:
        st.markdown('<div class="section-head">期首残高算出用の過去履歴</div>', unsafe_allow_html=True)
        st.dataframe(pre_opening_detail, use_container_width=True, height=360)
    with c_right:
        st.markdown('<div class="section-head">当月仕訳明細</div>', unsafe_allow_html=True)
        st.dataframe(month_detail, use_container_width=True, height=360)
with tab4:
    show_cols = ["相手先", "期首残高", "当月増加", "当月減少", "当月増減", "期末残高"]
    excel_bytes = make_export(selected_month_label, all_balance, account_partner[show_cols] if len(account_partner) else pd.DataFrame(columns=show_cols), month_detail, pre_opening_detail)
    st.download_button("Excelダウンロード", data=excel_bytes, file_name=f"kaikei_uchiwake_{selected_month_label}_{selected_account}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with st.expander("読み込みデータ確認"):
    st.markdown("**初期残高データ（正規化後）**")
    st.dataframe(initial_clean if len(initial_clean) else pd.DataFrame(columns=INITIAL_REQUIRED), use_container_width=True)
    st.markdown("**全期間仕訳マスタ（正規化後）**")
    st.dataframe(journal, use_container_width=True)
