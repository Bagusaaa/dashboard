
import io
from typing import List, Optional, Dict, Tuple

import pandas as pd
import streamlit as st
import plotly.express as px


# =========================
# PAGE CONFIG
# =========================
st.set_page_config(
    
    page_title="Ticket Sales Dashboard",
    page_icon="🎟️",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================
# STYLE
# =========================
st.markdown(
    """
    <style>
    .main {
        background-color: #f7f8fb;
    }

    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }

    .dashboard-title {
        font-size: 34px;
        font-weight: 800;
        color: #111827;
        margin-bottom: 2px;
    }

    .dashboard-subtitle {
        font-size: 15px;
        color: #6b7280;
        margin-bottom: 20px;
    }

    .kpi-card {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        padding: 18px 18px;
        border-radius: 18px;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.04);
        height: 118px;
    }

    .kpi-label {
        color: #6b7280;
        font-size: 13px;
        font-weight: 600;
        margin-bottom: 8px;
    }

    .kpi-value {
        color: #111827;
        font-size: 30px;
        font-weight: 800;
        line-height: 1.1;
    }

    .kpi-note {
        color: #9ca3af;
        font-size: 12px;
        margin-top: 7px;
    }

    .section-card {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 18px;
        padding: 18px;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.04);
        margin-bottom: 16px;
    }

    .section-title {
        font-size: 18px;
        font-weight: 800;
        color: #111827;
        margin-bottom: 4px;
    }

    .section-subtitle {
        font-size: 13px;
        color: #6b7280;
        margin-bottom: 12px;
    }

    div[data-testid="stMetric"] {
        background: #ffffff;
        border-radius: 16px;
        padding: 12px;
        border: 1px solid #e5e7eb;
    }

    .small-muted {
        color: #6b7280;
        font-size: 13px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================
# HELPER FUNCTIONS
# =========================
def normalize_col_name(col: object) -> str:
    return str(col).strip()


def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [normalize_col_name(c) for c in df.columns]
    return df


def find_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """
    Find column by exact normalized match first, then contains match.
    Compatible with Python 3.8+.
    """
    cols = list(df.columns)
    lower_map = {c.lower().strip(): c for c in cols}

    for cand in candidates:
        key = cand.lower().strip()
        if key in lower_map:
            return lower_map[key]

    for cand in candidates:
        key = cand.lower().strip()
        for c in cols:
            if key in c.lower().strip():
                return c

    return None


def safe_series(df: pd.DataFrame, col: Optional[str], default: str = "") -> pd.Series:
    if col and col in df.columns:
        return df[col]
    return pd.Series([default] * len(df), index=df.index)


def format_number(num: object) -> str:
    try:
        return f"{int(num):,}".replace(",", ".")
    except Exception:
        return "0"


def build_kpi_card(label: str, value: object, note: str = "") -> None:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{format_number(value)}</div>
            <div class="kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def to_excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, data in sheets.items():
            safe_sheet_name = sheet_name[:31]
            data.to_excel(writer, index=False, sheet_name=safe_sheet_name)
    return output.getvalue()


def prepare_data(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, Optional[str]]]:
    df = clean_columns(df)

    col_status = find_column(df, ["Status", "Payment Status", "Order Status"])
    col_order_id = find_column(df, ["Order ID", "Order Id", "ID Order", "Transaction ID", "Invoice ID"])
    col_order_date = find_column(df, ["Order Date", "Created Date", "Created At", "Transaction Date", "Tanggal Order"])
    col_ticket_name = find_column(df, ["Ticket Name", "Ticket", "Nama Ticket", "Nama Tiket", "Ticket Type"])
    col_voucher = find_column(df, ["Voucher", "Coupon", "Coupon Code", "Promo Code", "Kode Voucher"])
    col_payment_method = find_column(df, ["Payment Method", "Payment Type", "Metode Pembayaran"])
    col_full_name = find_column(df, ["Full Name", "Name", "Nama", "Guest Name"])
    col_email = find_column(df, ["Email"])
    col_phone = find_column(df, ["Phone", "Phone Number", "Mobile", "Whatsapp", "WhatsApp"])
    col_company = find_column(df, ["Company Name", "Company", "Nama Perusahaan"])
    col_company_type = find_column(df, ["Company Type", "Business Type"])
    col_job_title = find_column(df, ["Job Title", "Position", "Jabatan"])
    col_checked_in = find_column(df, ["Checked-In", "Checked In", "Checkin", "Check-in"])
    col_country = find_column(df, ["Country", "Negara"])
    col_province = find_column(df, ["Province", "Provinsi"])
    col_city = find_column(df, ["City", "Kota"])
    col_source = find_column(df, ["How did you find out", "Source", "Channel", "Info Source"])

    cols = {
        "status": col_status,
        "order_id": col_order_id,
        "order_date": col_order_date,
        "ticket_name": col_ticket_name,
        "voucher": col_voucher,
        "payment_method": col_payment_method,
        "full_name": col_full_name,
        "email": col_email,
        "phone": col_phone,
        "company": col_company,
        "company_type": col_company_type,
        "job_title": col_job_title,
        "checked_in": col_checked_in,
        "country": col_country,
        "province": col_province,
        "city": col_city,
        "source": col_source,
    }

    # Standard fields
    df["_status"] = safe_series(df, col_status).astype(str).str.strip()
    df["_order_id"] = safe_series(df, col_order_id).astype(str).str.strip()
    df["_ticket_name"] = safe_series(df, col_ticket_name).fillna("").astype(str).str.strip()
    df["_voucher"] = safe_series(df, col_voucher).fillna("").astype(str).str.strip()
    df["_payment_method"] = safe_series(df, col_payment_method).fillna("").astype(str).str.strip()
    df["_source"] = safe_series(df, col_source).fillna("").astype(str).str.strip()

    # Empty / null labeling
    df["_ticket_name_clean"] = df["_ticket_name"].replace({"": "VIP", "nan": "VIP", "None": "VIP"})
    df["_voucher_clean"] = df["_voucher"].replace({"": "(Kosong)", "nan": "(Kosong)", "None": "(Kosong)"})
    df["_payment_method_clean"] = df["_payment_method"].replace({"": "(Kosong)", "nan": "(Kosong)", "None": "(Kosong)"})
    df["_source_clean"] = df["_source"].replace({"": "(Kosong)", "nan": "(Kosong)", "None": "(Kosong)"})

    # Date
    if col_order_date and col_order_date in df.columns:
        df["_order_datetime"] = pd.to_datetime(df[col_order_date], errors="coerce")
    else:
        df["_order_datetime"] = pd.NaT

    df["_order_date"] = df["_order_datetime"].dt.date

    # Business logic
    df["_is_approve"] = df["_status"].str.lower().eq("approve")
    df["_is_waiting"] = df["_status"].str.lower().eq("checkout to payment")
    df["_is_coupon"] = ~df["_voucher_clean"].eq("(Kosong)")

    return df, cols


def plot_bar(data: pd.DataFrame, x: str, y: str, title: str, color: Optional[str] = None):
    fig = px.bar(data, x=x, y=y, color=color, text=y, title=title)
    fig.update_layout(
        title_font_size=16,
        height=390,
        margin=dict(l=10, r=10, t=50, b=10),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        legend_title_text="",
    )
    fig.update_traces(textposition="outside")
    return fig


def plot_line(data: pd.DataFrame, x: str, y: str, title: str, color: Optional[str] = None):
    fig = px.line(data, x=x, y=y, color=color, markers=True, title=title)
    fig.update_layout(
        title_font_size=16,
        height=390,
        margin=dict(l=10, r=10, t=50, b=10),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        legend_title_text="",
    )
    return fig


def plot_donut(data: pd.DataFrame, names: str, values: str, title: str):
    fig = px.pie(data, names=names, values=values, hole=0.55, title=title)
    fig.update_layout(
        title_font_size=16,
        height=390,
        margin=dict(l=10, r=10, t=50, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        legend_title_text="",
    )
    return fig


# =========================
# HEADER
# =========================
st.markdown('<div class="dashboard-title">🎟️ Ticket Sales Dashboard</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="dashboard-subtitle">Upload raw ticketing Excel, lalu dashboard otomatis membaca status, ticket type, coupon, dan transaksi per hari.</div>',
    unsafe_allow_html=True,
)


# =========================
# SIDEBAR UPLOAD
# =========================
with st.sidebar:
    st.header("Upload Data")
    uploaded_file = st.file_uploader("Upload raw Excel ticketing", type=["xlsx", "xls", "csv"])

    st.markdown("---")
    st.caption("Business logic:")
    st.caption("• Ticket sold = Status Approve")
    st.caption("• Waiting = Checkout to Payment")
    st.caption("• Jumlah tiket = count row")
    st.caption("• Jumlah transaksi = distinct Order ID")


if uploaded_file is None:
    st.info("Upload file Excel/CSV raw ticketing dulu untuk menampilkan dashboard.")
    st.stop()


# =========================
# LOAD DATA
# =========================
try:
    if uploaded_file.name.lower().endswith(".csv"):
        raw_df = pd.read_csv(uploaded_file)
    else:
        raw_df = pd.read_excel(uploaded_file)
except Exception as e:
    st.error(f"Gagal membaca file: {e}")
    st.stop()

if raw_df.empty:
    st.warning("File berhasil dibaca, tapi datanya kosong.")
    st.stop()

df, cols = prepare_data(raw_df)


# =========================
# SIDEBAR FILTER
# =========================
with st.sidebar:
    st.header("Filter")

    # Date filter
    valid_dates = df["_order_datetime"].dropna()
    if not valid_dates.empty:
        min_date = valid_dates.min().date()
        max_date = valid_dates.max().date()
        date_range = st.date_input(
            "Tanggal transaksi",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
        )
    else:
        date_range = None
        st.warning("Kolom tanggal transaksi tidak terdeteksi / tidak valid.")

    status_options = sorted([x for x in df["_status"].dropna().unique().tolist() if str(x).strip() != ""])
    selected_status = st.multiselect("Status", status_options, default=status_options)

    ticket_options = sorted(df["_ticket_name_clean"].dropna().unique().tolist())
    selected_tickets = st.multiselect("Ticket Type", ticket_options, default=ticket_options)

    voucher_options = sorted(df["_voucher_clean"].dropna().unique().tolist())
    selected_vouchers = st.multiselect("Coupon / Voucher", voucher_options, default=voucher_options)

    payment_options = sorted(df["_payment_method_clean"].dropna().unique().tolist())
    selected_payments = st.multiselect("Payment Method", payment_options, default=payment_options)

    source_options = sorted(df["_source_clean"].dropna().unique().tolist())
    selected_sources = st.multiselect("Source / Channel", source_options, default=source_options)


filtered = df.copy()

if date_range and len(date_range) == 2:
    start_date, end_date = date_range
    filtered = filtered[
        (filtered["_order_date"] >= start_date) &
        (filtered["_order_date"] <= end_date)
    ]

if selected_status:
    filtered = filtered[filtered["_status"].isin(selected_status)]

if selected_tickets:
    filtered = filtered[filtered["_ticket_name_clean"].isin(selected_tickets)]

if selected_vouchers:
    filtered = filtered[filtered["_voucher_clean"].isin(selected_vouchers)]

if selected_payments:
    filtered = filtered[filtered["_payment_method_clean"].isin(selected_payments)]

if selected_sources:
    filtered = filtered[filtered["_source_clean"].isin(selected_sources)]


# =========================
# KPI
# =========================
total_ticket = len(filtered)
sold_ticket = int(filtered["_is_approve"].sum())
waiting_ticket = int(filtered["_is_waiting"].sum())
unique_order = filtered["_order_id"].nunique() if cols["order_id"] else 0
coupon_ticket = int(filtered["_is_coupon"].sum())

kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
with kpi1:
    build_kpi_card("Total Ticket Row", total_ticket, "Jumlah tiket dari row data")
with kpi2:
    build_kpi_card("Sold / Approve", sold_ticket, "Status = Approve")
with kpi3:
    build_kpi_card("Waiting Payment", waiting_ticket, "Status = Checkout to Payment")
with kpi4:
    build_kpi_card("Unique Order", unique_order, "Distinct Order ID")
with kpi5:
    build_kpi_card("Coupon Used", coupon_ticket, "Voucher tidak kosong")


st.markdown("<br>", unsafe_allow_html=True)


# =========================
# TABS
# =========================
tab_overview, tab_daily, tab_coupon, tab_ticket, tab_source, tab_raw = st.tabs(
    [
        "Overview",
        "Daily Sales",
        "Coupon Analysis",
        "Ticket Type",
        "Source & Buyer",
        "Raw Data & Export",
    ]
)


# =========================
# OVERVIEW TAB
# =========================
with tab_overview:
    c1, c2 = st.columns([1.2, 1])

    with c1:
        daily_status = (
            filtered.dropna(subset=["_order_date"])
            .groupby(["_order_date", "_status"])
            .size()
            .reset_index(name="Total Ticket")
            .sort_values("_order_date")
        )
        if not daily_status.empty:
            st.plotly_chart(
                plot_bar(daily_status, "_order_date", "Total Ticket", "Daily Ticket by Status", "_status"),
                use_container_width=True,
            )
        else:
            st.info("Data tanggal belum tersedia untuk chart daily status.")

    with c2:
        status_summary = (
            filtered.groupby("_status")
            .size()
            .reset_index(name="Total Ticket")
            .sort_values("Total Ticket", ascending=False)
        )
        if not status_summary.empty:
            st.plotly_chart(
                plot_donut(status_summary, "_status", "Total Ticket", "Status Breakdown"),
                use_container_width=True,
            )

    c3, c4 = st.columns(2)

    with c3:
        ticket_summary = (
            filtered.groupby("_ticket_name_clean")
            .size()
            .reset_index(name="Total Ticket")
            .sort_values("Total Ticket", ascending=False)
            .head(10)
        )
        if not ticket_summary.empty:
            st.plotly_chart(
                plot_bar(ticket_summary, "_ticket_name_clean", "Total Ticket", "Top Ticket Type"),
                use_container_width=True,
            )

    with c4:
        promo_summary = pd.DataFrame({
            "Category": ["Coupon Used", "No Coupon"],
            "Total Ticket": [
                int(filtered["_is_coupon"].sum()),
                int((~filtered["_is_coupon"]).sum())
            ]
        })
        st.plotly_chart(
            plot_donut(promo_summary, "Category", "Total Ticket", "Coupon vs No Coupon"),
            use_container_width=True,
        )


# =========================
# DAILY TAB
# =========================
with tab_daily:
    st.markdown("### Daily Transaction Summary")

    daily_summary = (
        filtered.dropna(subset=["_order_date"])
        .groupby("_order_date")
        .agg(
            Total_Ticket=("_order_id", "size"),
            Sold_Approve=("_is_approve", "sum"),
            Waiting_Payment=("_is_waiting", "sum"),
            Coupon_Used=("_is_coupon", "sum"),
            Unique_Order=("_order_id", "nunique"),
        )
        .reset_index()
        .sort_values("_order_date")
    )

    if not daily_summary.empty:
        st.plotly_chart(
            plot_line(daily_summary, "_order_date", "Sold_Approve", "Approved Ticket per Day"),
            use_container_width=True,
        )
        st.dataframe(daily_summary, use_container_width=True)
    else:
        st.info("Tidak ada data daily yang bisa ditampilkan.")


# =========================
# COUPON TAB
# =========================
with tab_coupon:
    st.markdown("### Coupon / Voucher Analysis")

    coupon_only = filtered[filtered["_is_coupon"]].copy()

    c1, c2 = st.columns(2)

    with c1:
        coupon_summary = (
            coupon_only.groupby("_voucher_clean")
            .size()
            .reset_index(name="Total Ticket")
            .sort_values("Total Ticket", ascending=False)
        )

        if not coupon_summary.empty:
            st.plotly_chart(
                plot_bar(coupon_summary.head(15), "_voucher_clean", "Total Ticket", "Top 15 Coupon Usage"),
                use_container_width=True,
            )
        else:
            st.info("Belum ada coupon/voucher yang terpakai pada filter ini.")

    with c2:
        coupon_ticket = (
            coupon_only.groupby(["_voucher_clean", "_ticket_name_clean"])
            .size()
            .reset_index(name="Total Ticket")
            .sort_values("Total Ticket", ascending=False)
            .head(20)
        )

        if not coupon_ticket.empty:
            st.plotly_chart(
                plot_bar(coupon_ticket, "_voucher_clean", "Total Ticket", "Coupon by Ticket Type", "_ticket_name_clean"),
                use_container_width=True,
            )
        else:
            st.info("Coupon by ticket type belum tersedia.")

    st.markdown("#### Coupon Summary Table")
    st.dataframe(coupon_summary if not coupon_only.empty else pd.DataFrame(), use_container_width=True)


# =========================
# TICKET TYPE TAB
# =========================
with tab_ticket:
    st.markdown("### Ticket Type Breakdown")

    ticket_type_summary = (
        filtered.groupby("_ticket_name_clean")
        .agg(
            Total_Ticket=("_order_id", "size"),
            Sold_Approve=("_is_approve", "sum"),
            Waiting_Payment=("_is_waiting", "sum"),
            Unique_Order=("_order_id", "nunique"),
            Coupon_Used=("_is_coupon", "sum"),
        )
        .reset_index()
        .sort_values("Total_Ticket", ascending=False)
    )

    if not ticket_type_summary.empty:
        st.plotly_chart(
            plot_bar(ticket_type_summary, "_ticket_name_clean", "Total_Ticket", "Ticket Type Performance"),
            use_container_width=True,
        )

    st.dataframe(ticket_type_summary, use_container_width=True)


# =========================
# SOURCE & BUYER TAB
# =========================
with tab_source:
    st.markdown("### Source / Channel & Buyer Insight")

    c1, c2 = st.columns(2)

    with c1:
        source_summary = (
            filtered.groupby("_source_clean")
            .size()
            .reset_index(name="Total Ticket")
            .sort_values("Total Ticket", ascending=False)
            .head(15)
        )

        if not source_summary.empty:
            st.plotly_chart(
                plot_bar(source_summary, "_source_clean", "Total Ticket", "Top Source / Channel"),
                use_container_width=True,
            )
        else:
            st.info("Source/channel tidak tersedia.")

    with c2:
        payment_summary = (
            filtered.groupby("_payment_method_clean")
            .size()
            .reset_index(name="Total Ticket")
            .sort_values("Total Ticket", ascending=False)
            .head(15)
        )

        if not payment_summary.empty:
            st.plotly_chart(
                plot_donut(payment_summary, "_payment_method_clean", "Total Ticket", "Payment Method"),
                use_container_width=True,
            )
        else:
            st.info("Payment method tidak tersedia.")

    st.markdown("#### Geographic Summary")

    geo_cols = []
    if cols["country"]:
        geo_cols.append(cols["country"])
    if cols["province"]:
        geo_cols.append(cols["province"])
    if cols["city"]:
        geo_cols.append(cols["city"])

    if geo_cols:
        geo_summary = (
            filtered.groupby(geo_cols)
            .size()
            .reset_index(name="Total Ticket")
            .sort_values("Total Ticket", ascending=False)
            .head(50)
        )
        st.dataframe(geo_summary, use_container_width=True)
    else:
        st.info("Kolom country/province/city tidak terdeteksi.")


# =========================
# RAW DATA & EXPORT TAB
# =========================
with tab_raw:
    st.markdown("### Raw Data Filtered")

    st.caption(f"Menampilkan {format_number(len(filtered))} row dari total {format_number(len(df))} row.")

    display_cols = []
    for key in [
        "status",
        "order_id",
        "order_date",
        "ticket_name",
        "voucher",
        "payment_method",
        "full_name",
        "email",
        "phone",
        "company",
        "company_type",
        "job_title",
        "checked_in",
        "country",
        "province",
        "city",
        "source",
    ]:
        col = cols.get(key)
        if col and col in filtered.columns and col not in display_cols:
            display_cols.append(col)

    if not display_cols:
        display_cols = list(filtered.columns)

    st.dataframe(filtered[display_cols], use_container_width=True, height=500)

    daily_export = (
        filtered.dropna(subset=["_order_date"])
        .groupby("_order_date")
        .agg(
            Total_Ticket=("_order_id", "size"),
            Sold_Approve=("_is_approve", "sum"),
            Waiting_Payment=("_is_waiting", "sum"),
            Coupon_Used=("_is_coupon", "sum"),
            Unique_Order=("_order_id", "nunique"),
        )
        .reset_index()
    )

    status_export = (
        filtered.groupby("_status")
        .size()
        .reset_index(name="Total Ticket")
        .sort_values("Total Ticket", ascending=False)
    )

    ticket_export = (
        filtered.groupby("_ticket_name_clean")
        .size()
        .reset_index(name="Total Ticket")
        .sort_values("Total Ticket", ascending=False)
    )

    coupon_export = (
        filtered.groupby("_voucher_clean")
        .size()
        .reset_index(name="Total Ticket")
        .sort_values("Total Ticket", ascending=False)
    )

    export_bytes = to_excel_bytes({
        "Filtered Raw": filtered[display_cols],
        "Daily Summary": daily_export,
        "Status Summary": status_export,
        "Ticket Summary": ticket_export,
        "Coupon Summary": coupon_export,
    })

    st.download_button(
        label="⬇️ Download Excel Report",
        data=export_bytes,
        file_name="ticket_dashboard_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# =========================
# COLUMN MAPPING INFO
# =========================
with st.expander("Detected Column Mapping"):
    mapping_df = pd.DataFrame(
        [{"Field": k, "Detected Column": v if v else "-"} for k, v in cols.items()]
    )
    st.dataframe(mapping_df, use_container_width=True)
