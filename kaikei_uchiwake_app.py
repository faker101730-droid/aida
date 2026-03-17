
import io
from pathlib import Path
import pandas as pd
import streamlit as st

st.set_page_config(page_title="会計内訳アプリ Pro", page_icon="📘", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
.kpi-card {
    background: linear-gradient(135deg, #0f172a 0%, #111827 100%);
    border: 1px solid #334155;
    border-radius: 18px;
    padding: 18px 18px 14px 18px;
    min-height: 118px;
    box-shadow: 0 8px 20px rgba(0,0,0,0.18);
}
.kpi-label {font-size: 0.9rem; color: #cbd5e1; margin-bottom: 6px;}
.kpi-value {font-size: 2rem; font-weight: 700; color: white; line-height: 1.1;}
.kpi-note {font-size: 0.8rem; color: #94a3b8; margin-top: 6px;}
.info-box {
    background: #0b3b2f;
    border: 1px solid #14532d;
    border-radius: 14px;
    padding: 12px 14px;
    color: #dcfce7;
    margin-bottom: 0.5rem;
}
.subtle-box {
    background: #0f172a;
    border: 1px solid #1e293b;
    border-radius: 14px;
    padding: 12px 14px;
    color: #cbd5e1;
    margin-bottom: 0.5rem;
}
.small-caption {font-size: 0.82rem; color: #94a3b8;}
</style>
""", unsafe_allow_html=True)

INIT_REQUIRED = ["基準日", "勘定科目", "相手先", "初期残高"]
JOURNAL_REQUIRED = ["日付", "借方科目", "貸方科目", "金額", "相手先", "摘要"]

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [
        str(c).replace("\n", "").replace("\r", "").replace(" ", "").replace("　", "").strip()
        for c in df.columns
    ]
    return df

def detect_header_row(raw_df: pd.DataFrame, required_cols: list[str]) -> int:
    max_scan = min(len(raw_df), 10)
    need = set(required_cols)
    for i in range(max_scan):
        vals = set(
            str(v).replace("\n", "").replace("\r", "").replace(" ", "").replace("　", "").strip()
            for v in raw_df.iloc[i].tolist()
        )
        if need.issubset(vals):
            return i
    return 0

def read_any_table(uploaded_file):
    suffix = Path(uploaded_file.name).suffix.lower()
    if suffix in [".xlsx", ".xlsm", ".xls"]:
        raw = pd.read_excel(uploaded_file, header=None)
        uploaded_file.seek(0)
        return raw
    elif suffix == ".csv":
        last_err = None
        for enc in ["utf-8-sig", "cp932", "utf-8"]:
            try:
                raw = pd.read_csv(uploaded_file, header=None, encoding=enc)
                uploaded_file.seek(0)
                return raw
            except Exception as e:
                last_err = e
                uploaded_file.seek(0)
        raise ValueError(f"CSV読込に失敗しました: {last_err}")
    else:
        raise ValueError("対応形式は CSV / XLSX / XLS です。")

def parse_file(uploaded_file):
    raw = read_any_table(uploaded_file)
    init_row = detect_header_row(raw, INIT_REQUIRED)
    journal_row = detect_header_row(raw, JOURNAL_REQUIRED)
    parsed_init = None
    parsed_journal = None

    try:
        uploaded_file.seek(0)
        if Path(uploaded_file.name).suffix.lower() in [".xlsx", ".xlsm", ".xls"]:
            parsed_init = pd.read_excel(uploaded_file, header=init_row)
        else:
            for enc in ["utf-8-sig", "cp932", "utf-8"]:
                try:
                    uploaded_file.seek(0)
                    parsed_init = pd.read_csv(uploaded_file, header=init_row, encoding=enc)
                    break
                except Exception:
                    continue
        if parsed_init is not None:
            parsed_init = normalize_columns(parsed_init)
    except Exception:
        parsed_init = None

    try:
        uploaded_file.seek(0)
        if Path(uploaded_file.name).suffix.lower() in [".xlsx", ".xlsm", ".xls"]:
            parsed_journal = pd.read_excel(uploaded_file, header=journal_row)
        else:
            for enc in ["utf-8-sig", "cp932", "utf-8"]:
                try:
                    uploaded_file.seek(0)
                    parsed_journal = pd.read_csv(uploaded_file, header=journal_row, encoding=enc)
                    break
                except Exception:
                    continue
        if parsed_journal is not None:
            parsed_journal = normalize_columns(parsed_journal)
    except Exception:
        parsed_journal = None

    def has_cols(df, cols):
        return df is not None and set(cols).issubset(df.columns)

    if has_cols(parsed_init, INIT_REQUIRED) and not has_cols(parsed_journal, JOURNAL_REQUIRED):
        return "initial", parsed_init
    if has_cols(parsed_journal, JOURNAL_REQUIRED) and not has_cols(parsed_init, INIT_REQUIRED):
        return "journal", parsed_journal
    if has_cols(parsed_init, INIT_REQUIRED) and has_cols(parsed_journal, JOURNAL_REQUIRED):
        name = uploaded_file.name.lower()
        if "initial" in name or "balance" in name or "残高" in name:
            return "initial", parsed_init
        if "journal" in name or "仕訳" in name:
            return "journal", parsed_journal
        return ("initial", parsed_init) if init_row <= journal_row else ("journal", parsed_journal)

    raise ValueError("必要列を判定できませんでした。初期残高か仕訳マスタの形式を確認してください。")

def prepare_initial(df: pd.DataFrame) -> pd.DataFrame:
    missing = [c for c in INIT_REQUIRED if c not in df.columns]
    if missing:
        raise ValueError("初期残高に必要列が不足しています: " + ", ".join(missing))
    out = df[INIT_REQUIRED].copy()
    out["基準日"] = pd.to_datetime(out["基準日"], errors="coerce")
    out["初期残高"] = out["初期残高"].astype(str).str.replace(",", "", regex=False).str.replace(" ", "", regex=False)
    out["初期残高"] = pd.to_numeric(out["初期残高"], errors="coerce").fillna(0)
    out["勘定科目"] = out["勘定科目"].astype(str).str.strip()
    out["相手先"] = out["相手先"].astype(str).str.strip()
    if out["基準日"].isna().any():
        raise ValueError("初期残高の基準日に日付変換できない値があります。")
    return out

def prepare_journal(df: pd.DataFrame) -> pd.DataFrame:
    missing = [c for c in JOURNAL_REQUIRED if c not in df.columns]
    if missing:
        raise ValueError("仕訳マスタに必要列が不足しています: " + ", ".join(missing))
    out = df[JOURNAL_REQUIRED].copy()
    out["日付"] = pd.to_datetime(out["日付"], errors="coerce")
    out["金額"] = out["金額"].astype(str).str.replace(",", "", regex=False).str.replace(" ", "", regex=False)
    out["金額"] = pd.to_numeric(out["金額"], errors="coerce")
    for col in ["借方科目", "貸方科目", "相手先", "摘要"]:
        out[col] = out[col].astype(str).str.strip()
    if out["日付"].isna().any():
        raise ValueError("仕訳マスタの日付に日付変換できない値があります。")
    if out["金額"].isna().any():
        raise ValueError("仕訳マスタの金額に数値変換できない値があります。")
    return out

def choose_accounts(initial_df, journal_df):
    accounts = set(initial_df["勘定科目"].dropna().unique().tolist())
    accounts.update(journal_df["借方科目"].dropna().unique().tolist())
    accounts.update(journal_df["貸方科目"].dropna().unique().tolist())
    return sorted([a for a in accounts if str(a).strip() and str(a) != "nan"])

def month_options(initial_df, journal_df):
    dates = pd.concat([initial_df["基準日"], journal_df["日付"]]).dropna()
    return sorted(dates.dt.to_period("M").astype(str).unique().tolist())

def build_summary(initial_df, journal_df, account, target_month):
    target_period = pd.Period(target_month)
    month_start = target_period.start_time
    month_end = target_period.end_time
    init_base = initial_df["基準日"].max()

    if month_start <= init_base:
        raise ValueError("対象月が初期残高基準日以前です。基準日より後の月を選択してください。")

    init_acc = initial_df[initial_df["勘定科目"] == account].copy()
    before = journal_df[(journal_df["日付"] > init_base) & (journal_df["日付"] < month_start)].copy()
    current = journal_df[(journal_df["日付"] >= month_start) & (journal_df["日付"] <= month_end)].copy()

    init_partner = init_acc.groupby("相手先", dropna=False)["初期残高"].sum().rename("期首初期残高").reset_index()
    inc_before = before[before["借方科目"] == account].groupby("相手先", dropna=False)["金額"].sum().rename("前月まで増加").reset_index()
    dec_before = before[before["貸方科目"] == account].groupby("相手先", dropna=False)["金額"].sum().rename("前月まで減少").reset_index()
    inc_cur = current[current["借方科目"] == account].groupby("相手先", dropna=False)["金額"].sum().rename("当月増加").reset_index()
    dec_cur = current[current["貸方科目"] == account].groupby("相手先", dropna=False)["金額"].sum().rename("当月減少").reset_index()

    summary = init_partner.merge(inc_before, on="相手先", how="outer") \
                         .merge(dec_before, on="相手先", how="outer") \
                         .merge(inc_cur, on="相手先", how="outer") \
                         .merge(dec_cur, on="相手先", how="outer")
    if summary.empty:
        summary = pd.DataFrame({"相手先": []})
    for col in ["期首初期残高", "前月まで増加", "前月まで減少", "当月増加", "当月減少"]:
        if col not in summary.columns:
            summary[col] = 0
        summary[col] = pd.to_numeric(summary[col], errors="coerce").fillna(0)

    summary["期首残高"] = summary["期首初期残高"] + summary["前月まで増加"] - summary["前月まで減少"]
    summary["期末残高"] = summary["期首残高"] + summary["当月増加"] - summary["当月減少"]
    summary["マイナス残高"] = summary["期末残高"] < 0
    summary["当月動きなし"] = (summary["当月増加"].abs() + summary["当月減少"].abs()) == 0
    summary["長期残存候補"] = (summary["期首残高"] != 0) & (summary["当月動きなし"])

    summary = summary[["相手先", "期首残高", "当月増加", "当月減少", "期末残高", "長期残存候補", "マイナス残高"]].copy()
    summary = summary.sort_values(["期末残高", "相手先"], ascending=[False, True]).reset_index(drop=True)

    history_detail = before[(before["借方科目"] == account) | (before["貸方科目"] == account)].copy()
    current_detail = current[(current["借方科目"] == account) | (current["貸方科目"] == account)].copy()
    return summary, history_detail, current_detail, init_base, month_start, month_end

def fmt_yen(x):
    try:
        return f"{int(round(x, 0)):,}"
    except Exception:
        return x

def card(label, value, note):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{fmt_yen(value)}</div>
            <div class="kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

def to_excel(summary, current_detail, history_detail, account, month_text):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="相手先別残高", index=False)
        current_detail.to_excel(writer, sheet_name="当月仕訳明細", index=False)
        history_detail.to_excel(writer, sheet_name="期首算出用履歴", index=False)

        workbook = writer.book
        currency_fmt = workbook.add_format({"num_format": "#,##0"})
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9E2F3", "border": 1})

        for sheet_name, df in {"相手先別残高": summary, "当月仕訳明細": current_detail, "期首算出用履歴": history_detail}.items():
            ws = writer.sheets[sheet_name]
            for col_num, value in enumerate(df.columns.values):
                ws.write(0, col_num, value, header_fmt)
                width = max(12, min(28, len(str(value)) + 4))
                ws.set_column(col_num, col_num, width)
            for idx, col in enumerate(df.columns):
                if col in ["期首残高", "当月増加", "当月減少", "期末残高", "金額"]:
                    ws.set_column(idx, idx, 14, currency_fmt)

        pd.DataFrame({"項目": ["対象月", "勘定科目"], "内容": [month_text, account]}).to_excel(writer, sheet_name="出力条件", index=False)
    return output.getvalue()

def style_summary(df: pd.DataFrame):
    def row_style(row):
        if row.get("マイナス残高", False):
            return ["background-color: #4c0519; color: #ffe4e6;" for _ in row.index]
        if row.get("長期残存候補", False):
            return ["background-color: #3b2f0b; color: #fef3c7;" for _ in row.index]
        return ["" for _ in row.index]

    return df.style.apply(row_style, axis=1).format({
        "期首残高": "{:,.0f}",
        "当月増加": "{:,.0f}",
        "当月減少": "{:,.0f}",
        "期末残高": "{:,.0f}",
    })

def style_detail(df: pd.DataFrame):
    format_map = {}
    for col in ["金額", "初期残高", "期首残高", "当月増加", "当月減少", "期末残高"]:
        if col in df.columns:
            format_map[col] = "{:,.0f}"
    if "日付" in df.columns:
        try:
            x = df.copy()
            x["日付"] = pd.to_datetime(x["日付"], errors="coerce").dt.strftime("%Y-%m-%d")
            return x.style.format(format_map)
        except Exception:
            return df.style.format(format_map)
    return df.style.format(format_map)

st.title("📘 会計内訳アプリ Pro")
st.caption("初期残高 + 全期間仕訳マスタ から、対象月の期首残高・当月増減・期末残高を相手先別に可視化します。")

with st.sidebar:
    st.header("📂 データ読み込み")
    with st.expander("アップロードと入力ルール", expanded=False):
        init_up = st.file_uploader("① 初期残高データ", type=["csv", "xlsx", "xls"])
        journal_up = st.file_uploader("② 全期間仕訳マスタ", type=["csv", "xlsx", "xls"])
        st.markdown("<div class='small-caption'>推奨列名</div>", unsafe_allow_html=True)
        st.code("初期残高: 基準日 / 勘定科目 / 相手先 / 初期残高\n仕訳マスタ: 日付 / 借方科目 / 貸方科目 / 金額 / 相手先 / 摘要")

if not init_up or not journal_up:
    st.info("左のサイドバーから初期残高データと全期間仕訳マスタをアップロードしてください。")
    st.stop()

try:
    type1, df1 = parse_file(init_up)
    type2, df2 = parse_file(journal_up)
    parsed = {type1: df1, type2: df2}
    if "initial" not in parsed or "journal" not in parsed:
        raise ValueError("アップロード2ファイルから初期残高と仕訳マスタを判定できませんでした。")
    initial_df = prepare_initial(parsed["initial"])
    journal_df = prepare_journal(parsed["journal"])
except Exception as e:
    st.error(f"読み込みエラー：{e}")
    st.stop()

accounts = choose_accounts(initial_df, journal_df)
months = month_options(initial_df, journal_df)
if not months:
    st.warning("対象月候補がありません。日付データを確認してください。")
    st.stop()

c1, c2 = st.columns([1, 1])
with c1:
    target_month = st.selectbox("対象月", options=months, index=len(months)-1)
with c2:
    default_idx = accounts.index("未払金") if "未払金" in accounts else 0
    account = st.selectbox("勘定科目", options=accounts, index=default_idx)

try:
    summary, history_detail, current_detail, init_base, month_start, month_end = build_summary(initial_df, journal_df, account, target_month)
except Exception as e:
    st.error(f"集計エラー：{e}")
    st.stop()

view_summary = summary.copy()

col_info1, col_info2 = st.columns([1, 2])
with col_info1:
    st.markdown(f"<div class='info-box'><b>初期残高基準日:</b> {init_base.strftime('%Y-%m-%d')}</div>", unsafe_allow_html=True)
with col_info2:
    st.markdown(f"<div class='subtle-box'><b>対象月:</b> {target_month}・この月の <b>期首残高 / 当月増加 / 当月減少 / 期末残高</b> を表示しています。</div>", unsafe_allow_html=True)

k1, k2, k3, k4 = st.columns(4)
with k1:
    card("期首残高", view_summary["期首残高"].sum(), "対象月開始時点の残高")
with k2:
    card("当月増加", view_summary["当月増加"].sum(), "借方側で増えた金額")
with k3:
    card("当月減少", view_summary["当月減少"].sum(), "貸方側で減った金額")
with k4:
    card("期末残高", view_summary["期末残高"].sum(), "対象月終了時点の残高")

with st.sidebar:
    with st.expander("表示条件", expanded=False):
        hide_zero = st.checkbox("期末残高ゼロの相手先を非表示", value=True)
        only_warning = st.checkbox("注意行のみ表示", value=False)
        partner_kw = st.text_input("相手先検索", value="")

if hide_zero:
    view_summary = view_summary[view_summary["期末残高"] != 0]
if only_warning:
    view_summary = view_summary[(view_summary["長期残存候補"]) | (view_summary["マイナス残高"])]
if partner_kw.strip():
    view_summary = view_summary[view_summary["相手先"].astype(str).str.contains(partner_kw.strip(), case=False, na=False)]

tab1, tab2, tab3, tab4 = st.tabs(["📊 相手先別残高", "🧾 当月仕訳明細", "🕰️ 期首算出用履歴", "📥 Excel出力"])
with tab1:
    st.subheader(f"{target_month} {account} 相手先別残高")
    st.caption("黄色は長期残存候補、赤はマイナス残高です。")
    st.dataframe(style_summary(view_summary), use_container_width=True, height=460)
    st.caption(f"表示件数: {len(view_summary)}件")
with tab2:
    st.subheader(f"{target_month} 当月仕訳明細")
    st.dataframe(style_detail(current_detail.sort_values(["日付", "相手先"]).reset_index(drop=True)), use_container_width=True, height=460)
with tab3:
    st.subheader("期首算出用履歴")
    st.caption(f"{init_base.strftime('%Y-%m-%d')} の初期残高から {month_start.strftime('%Y-%m-%d')} の前日までの履歴")
    st.dataframe(style_detail(history_detail.sort_values(["日付", "相手先"]).reset_index(drop=True)), use_container_width=True, height=460)
with tab4:
    st.subheader("Excel出力")
    excel_bytes = to_excel(view_summary.reset_index(drop=True), current_detail, history_detail, account, target_month)
    st.download_button(
        "📥 Excelダウンロード",
        data=excel_bytes,
        file_name=f"会計内訳_{account}_{target_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption("出力内容: 相手先別残高 / 当月仕訳明細 / 期首算出用履歴 / 出力条件")
