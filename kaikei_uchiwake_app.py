
import io
from pathlib import Path
import pandas as pd
import streamlit as st

st.set_page_config(page_title="会計内訳アプリ Pro 福祉の森対応版", page_icon="📘", layout="wide")

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
FUKUSHI_REQUIRED = ["日付", "借方中区分", "貸方中区分", "借方金額", "貸方金額", "摘要文"]

BS_DEFAULT = [
    "現金","現金預金","普通預金","小口現金","立替金","未収金","未収補助金","前払費用","前払金","仮払金","貸付金",
    "未払金","未払費用","預り金","前受金","仮受金","借入金","賞与引当金","退職給付引当金","繰越利益剰余金"
]

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [
        str(c).replace("\n", "").replace("\r", "").replace(" ", "").replace("　", "").strip()
        for c in df.columns
    ]
    return df

def detect_header_row(raw_df: pd.DataFrame, required_cols):
    need = set(required_cols)
    max_scan = min(len(raw_df), 10)
    for i in range(max_scan):
        vals = set(str(v).replace("\n", "").replace("\r", "").replace(" ", "").replace("　", "").strip() for v in raw_df.iloc[i].tolist())
        if need.issubset(vals):
            return i
    return 0

def read_any_table(uploaded_file):
    suffix = Path(uploaded_file.name).suffix.lower()
    if suffix in [".xlsx", ".xlsm", ".xls"]:
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file, header=None)
    for enc in ["utf-8-sig", "cp932", "utf-8", "shift_jis"]:
        try:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, header=None, encoding=enc)
        except Exception:
            continue
    uploaded_file.seek(0)
    return pd.read_csv(uploaded_file, header=None)

def read_with_header(uploaded_file, header_row):
    suffix = Path(uploaded_file.name).suffix.lower()
    if suffix in [".xlsx", ".xlsm", ".xls"]:
        uploaded_file.seek(0)
        return normalize_columns(pd.read_excel(uploaded_file, header=header_row))
    for enc in ["utf-8-sig", "cp932", "utf-8", "shift_jis"]:
        try:
            uploaded_file.seek(0)
            return normalize_columns(pd.read_csv(uploaded_file, header=header_row, encoding=enc))
        except Exception:
            continue
    uploaded_file.seek(0)
    return normalize_columns(pd.read_csv(uploaded_file, header=header_row))

def parse_file(uploaded_file):
    raw = read_any_table(uploaded_file)
    init_row = detect_header_row(raw, INIT_REQUIRED)
    journal_row = detect_header_row(raw, JOURNAL_REQUIRED)
    fukushi_row = detect_header_row(raw, FUKUSHI_REQUIRED)

    parsed_init = None
    parsed_journal = None
    parsed_fukushi = None

    try:
        parsed_init = read_with_header(uploaded_file, init_row)
    except Exception:
        parsed_init = None
    try:
        parsed_journal = read_with_header(uploaded_file, journal_row)
    except Exception:
        parsed_journal = None
    try:
        parsed_fukushi = read_with_header(uploaded_file, fukushi_row)
    except Exception:
        parsed_fukushi = None

    def has_cols(df, cols):
        return df is not None and set(cols).issubset(df.columns)

    if has_cols(parsed_init, INIT_REQUIRED) and not (has_cols(parsed_journal, JOURNAL_REQUIRED) or has_cols(parsed_fukushi, FUKUSHI_REQUIRED)):
        return "initial", parsed_init

    if has_cols(parsed_journal, JOURNAL_REQUIRED) and not has_cols(parsed_init, INIT_REQUIRED):
        return "journal", parsed_journal

    if has_cols(parsed_fukushi, FUKUSHI_REQUIRED) and not has_cols(parsed_init, INIT_REQUIRED):
        return "journal", parsed_fukushi

    name = uploaded_file.name.lower()
    if has_cols(parsed_init, INIT_REQUIRED) and ("initial" in name or "balance" in name or "残高" in name):
        return "initial", parsed_init
    if "仕訳" in name or "journal" in name:
        if has_cols(parsed_journal, JOURNAL_REQUIRED):
            return "journal", parsed_journal
        if has_cols(parsed_fukushi, FUKUSHI_REQUIRED):
            return "journal", parsed_fukushi

    raise ValueError("必要列を判定できませんでした。初期残高または仕訳データの形式を確認してください。")

def clean_account_name(s):
    if pd.isna(s):
        return ""
    s = str(s).strip()
    s = pd.Series([s]).str.replace(r"^\([^\)]*\)", "", regex=True).iloc[0].strip()
    return s

def choose_account_label(mid, small, sub):
    for v in [small, mid, sub]:
        vv = clean_account_name(v)
        if vv:
            return vv
    return ""

def choose_partner(primary, secondary):
    p = "" if pd.isna(primary) else str(primary).strip()
    s = "" if pd.isna(secondary) else str(secondary).strip()
    return p if p else (s if s else "(空欄)")

def prepare_initial(df):
    missing = [c for c in INIT_REQUIRED if c not in df.columns]
    if missing:
        raise ValueError("初期残高に必要列が不足しています: " + ", ".join(missing))
    out = df.copy()
    out["基準日"] = pd.to_datetime(out["基準日"], errors="coerce")
    out["初期残高"] = pd.to_numeric(out["初期残高"].astype(str).str.replace(",", "", regex=False), errors="coerce")
    out["勘定科目"] = out["勘定科目"].astype(str).str.strip()
    out["相手先"] = out["相手先"].fillna("(空欄)").astype(str).str.strip().replace("", "(空欄)")
    if out["基準日"].isna().any():
        raise ValueError("初期残高の基準日に日付変換できない値があります。")
    if out["初期残高"].isna().any():
        raise ValueError("初期残高に数値変換できない値があります。")
    return out

def prepare_journal(df):
    cols = set(df.columns)

    if set(JOURNAL_REQUIRED).issubset(cols):
        out = df.copy()
        out["日付"] = pd.to_datetime(out["日付"], errors="coerce")
        out["金額"] = pd.to_numeric(out["金額"].astype(str).str.replace(",", "", regex=False), errors="coerce")
        for c in ["借方科目", "貸方科目", "相手先", "摘要"]:
            out[c] = out[c].fillna("").astype(str).str.strip()
        out["相手先_借方"] = out["相手先"].replace("", "(空欄)")
        out["相手先_貸方"] = out["相手先"].replace("", "(空欄)")
        if out["日付"].isna().any():
            raise ValueError("仕訳データの日付に日付変換できない値があります。")
        if out["金額"].isna().any():
            raise ValueError("仕訳データの金額に数値変換できない値があります。")
        return out

    if set(FUKUSHI_REQUIRED).issubset(cols):
        out = df.copy()
        out["日付"] = pd.to_datetime(out["日付"], errors="coerce")
        out["借方金額"] = pd.to_numeric(out["借方金額"].astype(str).str.replace(",", "", regex=False), errors="coerce")
        out["貸方金額"] = pd.to_numeric(out["貸方金額"].astype(str).str.replace(",", "", regex=False), errors="coerce")
        out["金額"] = out["借方金額"].where(out["借方金額"].notna(), out["貸方金額"])
        out["借方科目"] = [choose_account_label(m, s, b) for m, s, b in zip(out.get("借方中区分", ""), out.get("借方小区分", ""), out.get("借方補助区分", ""))]
        out["貸方科目"] = [choose_account_label(m, s, b) for m, s, b in zip(out.get("貸方中区分", ""), out.get("貸方小区分", ""), out.get("貸方補助区分", ""))]
        out["相手先_借方"] = [choose_partner(a, b) for a, b in zip(out.get("借方取引先", ""), out.get("貸方取引先", ""))]
        out["相手先_貸方"] = [choose_partner(a, b) for a, b in zip(out.get("貸方取引先", ""), out.get("借方取引先", ""))]
        out["相手先"] = out["相手先_借方"]
        out["摘要"] = out["摘要文"].fillna("").astype(str).str.strip()
        if out["日付"].isna().any():
            raise ValueError("福祉の森データの日付に日付変換できない値があります。")
        if out["金額"].isna().any():
            raise ValueError("福祉の森データの金額に数値変換できない値があります。")
        return out

    raise ValueError("仕訳データに必要列が不足しています。")

def choose_accounts(initial_df, journal_df, manual_list):
    manual = [x.strip() for x in manual_list.splitlines() if x.strip()]
    all_accounts = sorted(set(initial_df["勘定科目"].dropna().tolist()) | set(journal_df["借方科目"].dropna().tolist()) | set(journal_df["貸方科目"].dropna().tolist()))
    if manual:
        chosen = [a for a in all_accounts if a in manual]
        return chosen if chosen else all_accounts
    return all_accounts

def month_options(initial_df, journal_df):
    base = initial_df["基準日"].max()
    months = sorted(journal_df[journal_df["日付"] > base]["日付"].dt.to_period("M").astype(str).unique().tolist())
    return months

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
    inc_before = before[before["借方科目"] == account].groupby("相手先_借方", dropna=False)["金額"].sum().rename("前月まで増加").reset_index().rename(columns={"相手先_借方": "相手先"})
    dec_before = before[before["貸方科目"] == account].groupby("相手先_貸方", dropna=False)["金額"].sum().rename("前月まで減少").reset_index().rename(columns={"相手先_貸方": "相手先"})
    inc_cur = current[current["借方科目"] == account].groupby("相手先_借方", dropna=False)["金額"].sum().rename("当月増加").reset_index().rename(columns={"相手先_借方": "相手先"})
    dec_cur = current[current["貸方科目"] == account].groupby("相手先_貸方", dropna=False)["金額"].sum().rename("当月減少").reset_index().rename(columns={"相手先_貸方": "相手先"})

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
    for df in [history_detail, current_detail]:
        df["表示相手先"] = df["相手先_借方"]
        df.loc[df["貸方科目"] == account, "表示相手先"] = df.loc[df["貸方科目"] == account, "相手先_貸方"]

    return summary, history_detail, current_detail, init_base, month_start, month_end

def fmt_yen(x):
    try:
        return f"{int(round(float(x), 0)):,}"
    except Exception:
        return x

def card(label, value, note):
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{fmt_yen(value)}</div>
        <div class="kpi-note">{note}</div>
    </div>
    """, unsafe_allow_html=True)

def style_summary(df):
    show = df.copy()
    for c in ["期首残高", "当月増加", "当月減少", "期末残高"]:
        if c in show.columns:
            show[c] = show[c].apply(fmt_yen)
    show["長期残存候補"] = show["長期残存候補"].map(lambda x: "⚠" if x else "")
    show["マイナス残高"] = show["マイナス残高"].map(lambda x: "⚠" if x else "")
    return show.rename(columns={"長期残存候補": "長期残存", "マイナス残高": "マイナス"})

def style_detail(df):
    cols = [c for c in ["日付", "表示相手先", "借方科目", "貸方科目", "金額", "摘要", "伝票番号"] if c in df.columns]
    show = df[cols].copy() if cols else df.copy()
    if "表示相手先" in show.columns:
        show = show.rename(columns={"表示相手先": "相手先"})
    if "金額" in show.columns:
        show["金額"] = show["金額"].apply(fmt_yen)
    return show

def to_excel(summary, current_detail, history_detail, account, month_text):
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    output = io.BytesIO()
    condition_df = pd.DataFrame({"項目": ["対象月", "勘定科目"], "内容": [month_text, account]})

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="相手先別残高", index=False)
        current_detail.to_excel(writer, sheet_name="当月仕訳明細", index=False)
        history_detail.to_excel(writer, sheet_name="期首算出用履歴", index=False)
        condition_df.to_excel(writer, sheet_name="出力条件", index=False)

        header_fill = PatternFill(fill_type="solid", fgColor="D9E2F3")
        header_font = Font(bold=True)
        center = Alignment(horizontal="center", vertical="center")

        for sheet_name, df in {
            "相手先別残高": summary,
            "当月仕訳明細": current_detail,
            "期首算出用履歴": history_detail,
            "出力条件": condition_df,
        }.items():
            ws = writer.sheets[sheet_name]
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center

            for col_idx, col_name in enumerate(df.columns, start=1):
                max_len = len(str(col_name))
                for row in range(2, ws.max_row + 1):
                    value = ws.cell(row=row, column=col_idx).value
                    if value is not None:
                        max_len = max(max_len, len(str(value)))
                    if col_name in ["期首残高", "当月増加", "当月減少", "期末残高", "金額"]:
                        ws.cell(row=row, column=col_idx).number_format = '#,##0'
                ws.column_dimensions[get_column_letter(col_idx)].width = max(12, min(28, max_len + 2))

    output.seek(0)
    return output.getvalue()

st.title("📘 会計内訳アプリ Pro 福祉の森対応版")

with st.sidebar:
    with st.expander("データ読み込み", expanded=True):
        init_file = st.file_uploader("① 初期残高データ（CSV / Excel）", type=["csv", "xlsx", "xlsm", "xls"], key="init")
        journal_file = st.file_uploader("② 全期間仕訳マスタ（CSV / Excel / 福祉の森CSV）", type=["csv", "xlsx", "xlsm", "xls"], key="journal")

    with st.expander("表示条件", expanded=True):
        bs_text = st.text_area("対象勘定科目（改行区切り）", value="\n".join(BS_DEFAULT), height=220)
        hide_zero_partner = st.checkbox("期末残高ゼロの相手先を非表示", value=True)
        only_flags = st.checkbox("注意行のみ表示", value=False)
        partner_search = st.text_input("相手先検索")

    with st.expander("入力ルール", expanded=False):
        st.markdown("""
        - 初期残高：`基準日 / 勘定科目 / 相手先 / 初期残高`
        - 仕訳データ：標準形式 または 福祉の森出力CSVに対応
        - 福祉の森CSVは `借方中区分 / 借方小区分 / 借方取引先 / 借方金額 / 貸方... / 摘要文` を自動変換
        """)

if init_file and journal_file:
    try:
        kind1, df1 = parse_file(init_file)
        kind2, df2 = parse_file(journal_file)

        if kind1 == "journal" and kind2 == "initial":
            kind1, kind2 = kind2, kind1
            df1, df2 = df2, df1
            st.info("アップロード順を自動補正しました。")

        if kind1 != "initial" or kind2 != "journal":
            raise ValueError("初期残高と仕訳データの組み合わせを確認してください。")

        initial_df = prepare_initial(df1)
        journal_df = prepare_journal(df2)

        accounts = choose_accounts(initial_df, journal_df, bs_text)
        months = month_options(initial_df, journal_df)
        if not months:
            raise ValueError("初期残高基準日より後の仕訳月がありません。")

        head1, head2 = st.columns([1, 1])
        with head1:
            target_month = st.selectbox("対象月", months, index=len(months)-1)
        with head2:
            account = st.selectbox("勘定科目", accounts)

        summary, history_detail, current_detail, init_base, month_start, month_end = build_summary(initial_df, journal_df, account, target_month)

        if hide_zero_partner:
            summary = summary[summary["期末残高"] != 0].copy()
        if only_flags:
            summary = summary[(summary["長期残存候補"]) | (summary["マイナス残高"])].copy()
        if partner_search:
            summary = summary[summary["相手先"].astype(str).str.contains(partner_search, case=False, na=False)].copy()

        selected = summary.iloc[0] if not summary.empty else pd.Series({"期首残高":0, "当月増加":0, "当月減少":0, "期末残高":0})

        st.markdown(
            f'<div class="info-box">基準日: {init_base.date()} / 対象月: {target_month} / 勘定科目: {account}</div>',
            unsafe_allow_html=True
        )

        c1, c2, c3, c4 = st.columns(4)
        with c1: card("期首残高", selected["期首残高"], "前月末時点の残高")
        with c2: card("当月増加", selected["当月増加"], "当月借方計上")
        with c3: card("当月減少", selected["当月減少"], "当月貸方計上")
        with c4: card("期末残高", selected["期末残高"], "対象月末の残高")

        tab1, tab2, tab3, tab4 = st.tabs(["相手先別残高", "当月仕訳明細", "期首算出用履歴", "Excel出力"])

        with tab1:
            st.dataframe(style_summary(summary), use_container_width=True, hide_index=True)

        with tab2:
            current_display = current_detail.copy()
            current_display = current_display[(current_display["借方科目"] == account) | (current_display["貸方科目"] == account)]
            st.dataframe(style_detail(current_display), use_container_width=True, hide_index=True)

        with tab3:
            hist_display = history_detail.copy()
            hist_display = hist_display[(hist_display["借方科目"] == account) | (hist_display["貸方科目"] == account)]
            st.dataframe(style_detail(hist_display), use_container_width=True, hide_index=True)

        with tab4:
            excel_bytes = to_excel(summary, style_detail(current_detail), style_detail(history_detail), account, target_month)
            st.download_button(
                "Excelダウンロード",
                data=excel_bytes,
                file_name=f"会計内訳_{target_month}_{account}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"読み込みエラー: {e}")
else:
    st.markdown(
        '<div class="subtle-box">左のサイドバーから初期残高データと全期間仕訳マスタをアップロードしてください。福祉の森CSVはそのまま投入できます。</div>',
        unsafe_allow_html=True
    )
