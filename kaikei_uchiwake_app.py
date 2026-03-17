
import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="会計内訳アプリ Pro", page_icon="📘", layout="wide")

# ------------------------
# Style
# ------------------------
st.markdown("""
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 1.5rem;}
.main-title {
    font-size: 2rem;
    font-weight: 800;
    color: #143a52;
    margin-bottom: 0.2rem;
}
.sub-title {
    color: #4b6475;
    margin-bottom: 1rem;
}
.kpi-card {
    background: linear-gradient(135deg, #f8fbff 0%, #eef6fb 100%);
    border: 1px solid #d8e6f0;
    border-radius: 16px;
    padding: 14px 16px;
}
.kpi-label {
    color: #5d7482;
    font-size: 0.9rem;
    margin-bottom: 0.2rem;
}
.kpi-value {
    color: #143a52;
    font-size: 1.6rem;
    font-weight: 800;
}
.section-head {
    font-size: 1.1rem;
    font-weight: 700;
    color: #143a52;
    margin-top: 0.5rem;
    margin-bottom: 0.4rem;
}
.small-note {
    color: #6b7c88;
    font-size: 0.85rem;
}
div[data-testid="stMetric"] {
    background: linear-gradient(135deg, #f8fbff 0%, #eef6fb 100%);
    border: 1px solid #d8e6f0;
    padding: 10px 14px;
    border-radius: 16px;
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">会計内訳アプリ Pro</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">初期残高＋全期間仕訳マスタから、対象月の期首残高・当月増減・期末残高を自動表示</div>', unsafe_allow_html=True)

# ------------------------
# Utility
# ------------------------
def read_any(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith(".xlsx") or name.endswith(".xlsm") or name.endswith(".xls"):
        return pd.read_excel(uploaded_file)
    if name.endswith(".csv"):
        raw = uploaded_file.getvalue()
        for enc in ["utf-8-sig", "cp932", "utf-8"]:
            try:
                return pd.read_csv(io.BytesIO(raw), encoding=enc)
            except Exception:
                pass
        raise ValueError("CSVの読み込みに失敗しました。UTF-8 BOM付き または CP932 を推奨します。")
    raise ValueError("対応ファイルは CSV / Excel のみです。")

def normalize_initial(df):
    colmap = {
        "勘定科目":"勘定科目", "科目":"勘定科目", "account":"勘定科目",
        "相手先":"相手先", "補助科目":"相手先", "partner":"相手先",
        "初期残高":"初期残高", "前月末残高":"初期残高", "残高":"初期残高", "opening_balance":"初期残高",
        "基準日":"基準日", "日付":"基準日", "date":"基準日",
    }
    df = df.rename(columns={c: colmap.get(c, c) for c in df.columns})
    required = ["基準日", "勘定科目", "相手先", "初期残高"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"初期残高データに必要列が不足しています: {', '.join(missing)}")
    df = df[required].copy()
    df["基準日"] = pd.to_datetime(df["基準日"], errors="coerce")
    df["相手先"] = df["相手先"].fillna("（相手先未設定）").astype(str)
    df["勘定科目"] = df["勘定科目"].astype(str)
    df["初期残高"] = pd.to_numeric(df["初期残高"], errors="coerce").fillna(0)
    df = df.dropna(subset=["基準日"])
    return df

def normalize_journal(df):
    colmap = {
        "日付":"日付", "date":"日付",
        "借方科目":"借方科目", "debit_account":"借方科目",
        "貸方科目":"貸方科目", "credit_account":"貸方科目",
        "金額":"金額", "amount":"金額",
        "相手先":"相手先", "補助科目":"相手先", "partner":"相手先",
        "摘要":"摘要", "description":"摘要",
        "仕訳ID":"仕訳ID", "id":"仕訳ID",
    }
    df = df.rename(columns={c: colmap.get(c, c) for c in df.columns})
    required = ["日付", "借方科目", "貸方科目", "金額", "相手先", "摘要"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"仕訳マスタに必要列が不足しています: {', '.join(missing)}")
    work = df.copy()
    work["日付"] = pd.to_datetime(work["日付"], errors="coerce")
    work["金額"] = pd.to_numeric(work["金額"], errors="coerce").fillna(0)
    work["相手先"] = work["相手先"].fillna("（相手先未設定）").astype(str)
    work["摘要"] = work["摘要"].fillna("").astype(str)
    work["借方科目"] = work["借方科目"].astype(str)
    work["貸方科目"] = work["貸方科目"].astype(str)
    if "仕訳ID" not in work.columns:
        work["仕訳ID"] = range(1, len(work) + 1)
    work = work.dropna(subset=["日付"])
    return work

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

        workbook = writer.book
        money_fmt = workbook.add_format({'num_format': '#,##0;[Red](#,##0)'})
        head_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#143A52', 'font_color': 'white',
            'border': 0, 'align': 'center', 'valign': 'vcenter'
        })
        note_fmt = workbook.add_format({'font_color': '#5D7482'})
        for sheet_name, df in {
            "月末残高一覧": all_balance,
            "相手先別内訳": partner_summary,
            "当月仕訳明細": month_detail,
            "期首算出用履歴": pre_opening_detail
        }.items():
            ws = writer.sheets[sheet_name]
            ws.freeze_panes(1, 0)
            ws.set_row(0, 24, head_fmt)
            for i, col in enumerate(df.columns):
                width = max(12, min(28, int(max(df[col].astype(str).map(len).max() if len(df) else 0, len(str(col))) + 2)))
                ws.set_column(i, i, width)
            for i, col in enumerate(df.columns):
                if any(k in col for k in ["残高", "増加", "減少", "金額"]):
                    ws.set_column(i, i, 14, money_fmt)

        cover = workbook.add_worksheet("README")
        cover.write("A1", "会計内訳アプリ Pro 出力", workbook.add_format({'bold': True, 'font_size': 14}))
        cover.write("A3", f"対象月: {selected_month_label}")
        cover.write("A5", "シート説明", workbook.add_format({'bold': True}))
        cover.write("A6", "月末残高一覧: 期首残高・当月増減・期末残高を勘定科目別に集計")
        cover.write("A7", "相手先別内訳: 選択勘定科目の相手先別内訳")
        cover.write("A8", "当月仕訳明細: 対象月の選択勘定科目に関する仕訳")
        cover.write("A9", "期首算出用履歴: 対象月の期首残高計算に使った過去明細")
        cover.set_column("A:A", 72, note_fmt)
    output.seek(0)
    return output.getvalue()

# ------------------------
# Sidebar
# ------------------------
with st.sidebar:
    st.header("📂 データ読み込み")
    initial_file = st.file_uploader("① 初期残高データ", type=["csv", "xlsx", "xls"])
    journal_file = st.file_uploader("② 全期間仕訳マスタ", type=["csv", "xlsx", "xls"])
    st.caption("推奨列名")
    st.code("初期残高: 基準日 / 勘定科目 / 相手先 / 初期残高\n仕訳: 日付 / 借方科目 / 貸方科目 / 金額 / 相手先 / 摘要 / 仕訳ID", language="text")

if not journal_file:
    st.info("まずは「全期間仕訳マスタ」をアップロードしてください。初期残高は未設定でも動きますが、期首残高は0始まりになります。")
    st.stop()

# ------------------------
# Load data
# ------------------------
try:
    journal_raw = read_any(journal_file)
    journal = normalize_journal(journal_raw)
    initial = None
    if initial_file:
        initial_raw = read_any(initial_file)
        initial = normalize_initial(initial_raw)
except Exception as e:
    st.error(f"読み込みエラー: {e}")
    st.stop()

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

# ------------------------
# Period selection
# ------------------------
months = sorted(journal["日付"].dt.to_period("M").unique())
default_month = months[-1]
c1, c2, c3 = st.columns([1.2, 1, 1.2])
with c1:
    selected_period = st.selectbox("対象月", months, index=len(months)-1, format_func=lambda p: p.strftime("%Y-%m"))
with c2:
    account_options = sorted(pd.unique(pd.concat([movements["勘定科目"], initial["勘定科目"]]).astype(str)))
    selected_account = st.selectbox("勘定科目", account_options)
with c3:
    view_mode = st.selectbox("表示モード", ["月次残高管理", "明細確認"])

month_start, next_month_start = month_bounds(selected_period)
selected_month_label = selected_period.strftime("%Y-%m")

# ------------------------
# Calculations
# ------------------------
initial_clean = initial[initial["勘定科目"] != "初期設定なし"].copy()

opening_initial = initial_clean[initial_clean["基準日"] <= initial_base_date].groupby(["勘定科目", "相手先"], as_index=False)["初期残高"].sum()

pre_month_movements = movements[movements["日付"] < month_start].groupby(["勘定科目", "相手先"], as_index=False)["金額"].sum()
pre_month_movements = pre_month_movements.rename(columns={"金額": "前月まで累計増減"})

opening_partner = opening_initial.merge(pre_month_movements, on=["勘定科目", "相手先"], how="outer")
opening_partner["初期残高"] = opening_partner["初期残高"].fillna(0)
opening_partner["前月まで累計増減"] = opening_partner["前月まで累計増減"].fillna(0)
opening_partner["期首残高"] = opening_partner["初期残高"] + opening_partner["前月まで累計増減"]

month_movement_detail = movements[(movements["日付"] >= month_start) & (movements["日付"] < next_month_start)].copy()
month_partner = month_movement_detail.groupby(["勘定科目", "相手先"], as_index=False).agg(
    当月増加=("増加額", "sum"),
    当月減少=("減少額", "sum"),
    当月増減=("金額", "sum")
)

partner_summary = opening_partner.merge(month_partner, on=["勘定科目", "相手先"], how="outer")
for col in ["初期残高", "前月まで累計増減", "期首残高", "当月増加", "当月減少", "当月増減"]:
    if col in partner_summary.columns:
        partner_summary[col] = partner_summary[col].fillna(0)
partner_summary["期末残高"] = partner_summary["期首残高"] + partner_summary["当月増減"]

all_balance = partner_summary.groupby("勘定科目", as_index=False).agg(
    期首残高=("期首残高", "sum"),
    当月増加=("当月増加", "sum"),
    当月減少=("当月減少", "sum"),
    当月増減=("当月増減", "sum"),
    期末残高=("期末残高", "sum"),
)
all_balance = all_balance.sort_values("期末残高", ascending=False)

account_partner = partner_summary[partner_summary["勘定科目"] == selected_account].copy()
account_partner = account_partner.sort_values("期末残高", ascending=False)

pre_opening_detail = movements[(movements["日付"] < month_start) & (movements["勘定科目"] == selected_account)].copy()
month_detail = month_movement_detail[month_movement_detail["勘定科目"] == selected_account].copy()
month_detail = month_detail.sort_values(["日付", "仕訳ID"])

account_total = all_balance[all_balance["勘定科目"] == selected_account]
if len(account_total) == 0:
    account_total = pd.DataFrame([{"勘定科目": selected_account, "期首残高": 0, "当月増加": 0, "当月減少": 0, "当月増減": 0, "期末残高": 0}])
account_total = account_total.iloc[0]

# ------------------------
# Top notes
# ------------------------
note_cols = st.columns([1.1, 1.6])
with note_cols[0]:
    if no_initial_flag:
        st.warning("初期残高データ未設定のため、期首残高は0起点で計算しています。")
    else:
        st.success(f"初期残高基準日: {initial_base_date.strftime('%Y-%m-%d')}")

with note_cols[1]:
    st.info(f"対象月は {selected_month_label}。この月の期首残高・当月増減・期末残高を表示しています。")

# ------------------------
# KPI row
# ------------------------
m1, m2, m3, m4 = st.columns(4)
m1.metric("期首残高", fmt_amount(account_total["期首残高"]))
m2.metric("当月増加", fmt_amount(account_total["当月増加"]))
m3.metric("当月減少", fmt_amount(account_total["当月減少"]))
m4.metric("期末残高", fmt_amount(account_total["期末残高"]))

# ------------------------
# Main tabs
# ------------------------
tab1, tab2, tab3, tab4 = st.tabs(["📊 残高一覧", "🧾 相手先別内訳", "🔎 明細ドリルダウン", "⬇️ Excel出力"])

with tab1:
    st.markdown('<div class="section-head">対象月の勘定科目別残高一覧</div>', unsafe_allow_html=True)
    st.dataframe(
        all_balance.style.format({
            "期首残高": "{:,.0f}",
            "当月増加": "{:,.0f}",
            "当月減少": "{:,.0f}",
            "当月増減": "{:,.0f}",
            "期末残高": "{:,.0f}",
        }),
        use_container_width=True,
        height=460
    )

with tab2:
    st.markdown(f'<div class="section-head">「{selected_account}」の相手先別内訳</div>', unsafe_allow_html=True)
    show_cols = ["相手先", "期首残高", "当月増加", "当月減少", "当月増減", "期末残高"]
    if len(account_partner) == 0:
        st.warning("この勘定科目には該当データがありません。")
    else:
        st.dataframe(
            account_partner[show_cols].style.format({
                "期首残高": "{:,.0f}",
                "当月増加": "{:,.0f}",
                "当月減少": "{:,.0f}",
                "当月増減": "{:,.0f}",
                "期末残高": "{:,.0f}",
            }),
            use_container_width=True,
            height=460
        )

with tab3:
    c_left, c_right = st.columns(2)
    with c_left:
        st.markdown('<div class="section-head">期首残高算出用の過去履歴</div>', unsafe_allow_html=True)
        st.dataframe(pre_opening_detail, use_container_width=True, height=380)
        st.caption("対象月より前の、この勘定科目に関する履歴です。")
    with c_right:
        st.markdown('<div class="section-head">当月仕訳明細</div>', unsafe_allow_html=True)
        st.dataframe(month_detail, use_container_width=True, height=380)
        st.caption("対象月の、この勘定科目に関する仕訳です。")

with tab4:
    st.markdown('<div class="section-head">Excel出力</div>', unsafe_allow_html=True)
    st.write("対象月・選択勘定科目の内容を整理したExcelをダウンロードできます。")
    excel_bytes = make_export(selected_month_label, all_balance, account_partner[show_cols] if len(account_partner) else pd.DataFrame(columns=show_cols), month_detail, pre_opening_detail)
    st.download_button(
        "Excelダウンロード",
        data=excel_bytes,
        file_name=f"kaikei_uchiwake_{selected_month_label}_{selected_account}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ------------------------
# Raw data expanders
# ------------------------
with st.expander("読み込みデータ確認"):
    st.markdown("**初期残高データ**")
    st.dataframe(initial_clean if len(initial_clean) else pd.DataFrame(columns=["基準日","勘定科目","相手先","初期残高"]), use_container_width=True)
    st.markdown("**全期間仕訳マスタ**")
    st.dataframe(journal, use_container_width=True)
