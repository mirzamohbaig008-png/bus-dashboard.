"""
BusOps Command Center  |  SOLVATHON 2026 - PS-3
Bus ticket analytics for city service operations.

Run:  py -m streamlit run app.py
Keep the Excel file in the same folder as this script (or upload it from the sidebar).
"""

import io
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

DATA_FILE = "Ticket data_1 Day_City Service.xlsx"
REQUIRED = ["Waybill Number", "Ticket Number", "From Stop", "To Stop", "Time",
            "Date", "Total Psgr", "Coll Amt", "Uncoll Amt"]

# Palette: transit-map inspired
INK, TEAL, AMBER, RED, BLUE, MIST = "#12263a", "#0f9d8a", "#e8a317", "#d64545", "#2f6fb5", "#e6ecf1"

st.set_page_config(page_title="BusOps Command Center", page_icon="🚌",
                   layout="wide", initial_sidebar_state="expanded")

# ─────────────────────────────────────────────────────────────
# STYLE
# ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=DM+Sans:wght@400;500;700&display=swap');
.stMarkdown, .stCaption, label, p { font-family: 'DM Sans', sans-serif; }
h1, h2, h3, h4 { font-family: 'Space Grotesk', sans-serif !important; color: #12263a; }
.block-container { padding-top: 1.6rem; max-width: 1400px; }
#MainMenu, footer { visibility: hidden; }

/* Hero with route line */
.hero { background: #12263a; color: #fff; border-radius: 14px; padding: 1.5rem 1.8rem 1.3rem; margin-bottom: 1.1rem; }
.hero-title { font-family: 'Space Grotesk', sans-serif; font-size: 2rem; font-weight: 700; letter-spacing: -0.02em; }
.hero-sub { color: #a9bccd; font-size: 0.95rem; margin-top: 0.15rem; }
.routeline { position: relative; display: flex; margin-top: 1.5rem; }
.routeline::before { content: ""; position: absolute; top: 7px; left: 8%; right: 8%; height: 4px; background: #0f9d8a; border-radius: 2px; }
.stop { flex: 1; text-align: center; position: relative; padding-top: 26px; }
.stop::before { content: ""; position: absolute; top: 0; left: 50%; transform: translateX(-50%);
  width: 18px; height: 18px; border-radius: 50%; background: #12263a; border: 4px solid #e8a317; }
.stop-name { font-weight: 700; font-size: 0.92rem; }
.stop-val { color: #a9bccd; font-size: 0.8rem; }

/* KPI strip */
.kpi { border-left: 4px solid #0f9d8a; padding: 0.2rem 0 0.2rem 0.8rem; }
.kpi.bad { border-left-color: #d64545; }
.kpi.warn { border-left-color: #e8a317; }
.kpi-label { color: #5b6b7b; font-size: 0.82rem; }
.kpi-value { font-family: 'Space Grotesk', sans-serif; color: #12263a; font-size: 1.65rem; font-weight: 700; line-height: 1.2; }
.kpi-note { color: #5b6b7b; font-size: 0.78rem; }
.kpi.bad .kpi-note { color: #d64545; font-weight: 500; }

/* Insight strip */
.insight { background: #f1f5f8; border-radius: 10px; padding: 0.7rem 0.9rem; height: 100%; }
.insight-label { color: #5b6b7b; font-size: 0.78rem; }
.insight-text { color: #12263a; font-size: 0.95rem; font-weight: 500; margin-top: 0.1rem; }

.note { color: #5b6b7b; font-size: 0.85rem; }
button[data-baseweb="tab"] { font-weight: 500; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────
def short_inr(n) -> str:
    n = float(n)
    a = abs(n)
    if a >= 1e7:
        return f"₹{n / 1e7:.2f} Cr"
    if a >= 1e5:
        return f"₹{n / 1e5:.2f} L"
    return f"₹{n:,.0f}"


def kpi(label, value, note="", tone=""):
    return (f'<div class="kpi {tone}"><div class="kpi-label">{label}</div>'
            f'<div class="kpi-value">{value}</div><div class="kpi-note">{note}</div></div>')


def insight(label, text):
    return (f'<div class="insight"><div class="insight-label">{label}</div>'
            f'<div class="insight-text">{text}</div></div>')


def style(fig, h=380):
    fig.update_layout(
        template="plotly_white", height=h, margin=dict(l=8, r=8, t=56, b=8),
        font=dict(family="DM Sans, sans-serif", size=12, color=INK),
        title=dict(font=dict(family="Space Grotesk, sans-serif", size=16), x=0),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="right", x=1),
    )
    return fig


def clean_stop(s: pd.Series) -> pd.Series:
    s = s.astype("string").str.strip().fillna("Unknown")
    return s.replace({"nan": "Unknown", "": "Unknown", "<NA>": "Unknown", "None": "Unknown"}).astype(str)


@st.cache_data(show_spinner="Reading ticket data…")
def load_data(raw: bytes) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(raw))
    df.columns = df.columns.astype(str).str.strip()
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError("Missing columns: " + ", ".join(missing))

    df["From Stop"] = clean_stop(df["From Stop"])
    df["To Stop"] = clean_stop(df["To Stop"])
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Hour"] = pd.to_datetime(df["Time"].astype(str), errors="coerce").dt.hour
    for c in ["Total Psgr", "Coll Amt", "Uncoll Amt"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df = df.dropna(subset=["Hour"]).copy()
    df["Hour"] = df["Hour"].astype(int)
    df["Total Amt"] = df["Coll Amt"] + df["Uncoll Amt"]
    df["Route"] = df["From Stop"] + " → " + df["To Stop"]
    df["IsZeroColl"] = df["Coll Amt"] == 0
    return df


# ─────────────────────────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────────────────────────
default_path = Path(__file__).with_name(DATA_FILE)
upload = st.sidebar.file_uploader("Use a different Excel file", type=["xlsx"])

if upload is not None:
    raw = upload.getvalue()
elif default_path.exists():
    raw = default_path.read_bytes()
else:
    st.error(f"Could not find **{DATA_FILE}** next to app.py. Put the file in the same folder, "
             "or upload it from the sidebar.")
    st.stop()

try:
    df = load_data(raw)
except Exception as e:
    st.error(f"The file could not be read. {e}")
    st.stop()

HMIN, HMAX = int(df["Hour"].min()), int(df["Hour"].max())
ALL_WB = sorted(df["Waybill Number"].unique())
ALL_FROM = sorted(df["From Stop"].unique())
ALL_TO = sorted(df["To Stop"].unique())
VIEWS = ["All tickets", "Zero-collection tickets only", "Fully collected tickets only"]

dates = df["Date"].dropna().dt.date.unique()
if len(dates) == 1:
    DATE_LABEL = dates[0].strftime("%d %b %Y")
elif len(dates) > 1:
    DATE_LABEL = f"{min(dates):%d %b} – {max(dates):%d %b %Y}"
else:
    DATE_LABEL = "Date not available"


# ─────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────
def reset_filters():
    st.session_state.f_view = VIEWS[0]
    st.session_state.f_hours = (HMIN, HMAX)
    st.session_state.f_wb = []
    st.session_state.f_from = []
    st.session_state.f_to = []


st.sidebar.markdown("### Filters")
st.sidebar.caption("Leave a list empty to include everything.")
view = st.sidebar.radio("Ticket view", VIEWS, key="f_view")
hours = st.sidebar.slider("Hours of the day", HMIN, HMAX, (HMIN, HMAX), key="f_hours") if HMIN < HMAX \
    else (HMIN, HMAX)
sel_wb = st.sidebar.multiselect("Waybill", ALL_WB, key="f_wb", placeholder="All waybills")
sel_from = st.sidebar.multiselect("Boarding stop", ALL_FROM, key="f_from", placeholder="All stops")
sel_to = st.sidebar.multiselect("Destination stop", ALL_TO, key="f_to", placeholder="All stops")
st.sidebar.button("Reset filters", on_click=reset_filters, width="stretch")

st.sidebar.markdown("---")
target = st.sidebar.slider("Collection target (%)", 10, 100, 50, 5, key="f_target",
                           help="Used to colour the KPIs and in the scenario simulator.")

mask = df["Hour"].between(hours[0], hours[1])
if sel_wb:
    mask &= df["Waybill Number"].isin(sel_wb)
if sel_from:
    mask &= df["From Stop"].isin(sel_from)
if sel_to:
    mask &= df["To Stop"].isin(sel_to)
if view == VIEWS[1]:
    mask &= df["IsZeroColl"]
elif view == VIEWS[2]:
    mask &= ~df["IsZeroColl"]
fdf = df[mask].copy()

if fdf.empty:
    st.warning("No tickets match the current filters. Widen the hour range or click **Reset filters** in the sidebar.")
    st.stop()

st.sidebar.download_button("Download filtered data (CSV)", fdf.to_csv(index=False).encode("utf-8"),
                           file_name="filtered_tickets.csv", mime="text/csv", width="stretch")


# ─────────────────────────────────────────────────────────────
# METRICS
# ─────────────────────────────────────────────────────────────
tickets = len(fdf)
pax = int(fdf["Total Psgr"].sum())
coll = float(fdf["Coll Amt"].sum())
uncoll = float(fdf["Uncoll Amt"].sum())
potential = coll + uncoll
rate = coll / potential * 100 if potential > 0 else 0.0
zero_n = int(fdf["IsZeroColl"].sum())
zero_pct = zero_n / tickets * 100
zero_pax = int(fdf.loc[fdf["IsZeroColl"], "Total Psgr"].sum())
avg_fare = potential / pax if pax else 0.0

# Hourly table
hourly = fdf.groupby("Hour").agg(Tickets=("Ticket Number", "count"), Pax=("Total Psgr", "sum"),
                                 Coll=("Coll Amt", "sum"), Uncoll=("Uncoll Amt", "sum"),
                                 Zero=("IsZeroColl", "sum"))
hourly = hourly.reindex(range(int(fdf["Hour"].min()), int(fdf["Hour"].max()) + 1), fill_value=0)
hourly.index.name = "Hour"
hourly = hourly.reset_index()
hourly["Total"] = hourly["Coll"] + hourly["Uncoll"]
hourly["Rate"] = np.where(hourly["Total"] > 0, hourly["Coll"] / hourly["Total"] * 100, np.nan)
hourly["ZeroPct"] = np.where(hourly["Tickets"] > 0, hourly["Zero"] / hourly["Tickets"] * 100, 0)

# Busiest 3-hour window
pax_by_hour = hourly.set_index("Hour")["Pax"]
roll = pax_by_hour.rolling(3, min_periods=1).sum()
w_end = int(roll.idxmax())
w_start = max(w_end - 2, int(pax_by_hour.index.min()))
w_share = roll.max() / pax_by_hour.sum() * 100 if pax_by_hour.sum() else 0
peak_window = f"{w_start:02d}:00 – {w_end + 1:02d}:00"
peak_hour = int(pax_by_hour.idxmax())

# Origin-destination table
od = fdf.groupby(["From Stop", "To Stop"]).agg(
    Tickets=("Ticket Number", "count"), Passengers=("Total Psgr", "sum"),
    Collected=("Coll Amt", "sum"), Uncollected=("Uncoll Amt", "sum"), TotalAmt=("Total Amt", "sum"),
).reset_index()
od["Route"] = od["From Stop"] + " → " + od["To Stop"]
od["CollRate"] = np.where(od["TotalAmt"] > 0, od["Collected"] / od["TotalAmt"] * 100, 0.0)
od = od.sort_values("Passengers", ascending=False)

od_known = od[(od["To Stop"] != "Unknown") & (od["From Stop"] != od["To Stop"])].copy()
a, b = od_known["From Stop"], od_known["To Stop"]
od_known["Corridor"] = np.where(a < b, a + " ⇄ " + b, b + " ⇄ " + a)
corr = (od_known.groupby("Corridor")
        .agg(Passengers=("Passengers", "sum"), Potential=("TotalAmt", "sum"), Directions=("Route", "nunique"))
        .reset_index().sort_values("Passengers", ascending=False))

# Waybills
wb = fdf.groupby("Waybill Number").agg(
    Tickets=("Ticket Number", "count"), Passengers=("Total Psgr", "sum"),
    Collected=("Coll Amt", "sum"), Uncollected=("Uncoll Amt", "sum"),
    TotalAmt=("Total Amt", "sum"), ZeroTickets=("IsZeroColl", "sum"),
).reset_index()
wb["CollRate"] = np.where(wb["TotalAmt"] > 0, wb["Collected"] / wb["TotalAmt"] * 100, 0.0)
wb["ZeroPct"] = wb["ZeroTickets"] / wb["Tickets"] * 100
wb["Waybill"] = wb["Waybill Number"].astype(str)
wb = wb.sort_values("CollRate")

# Stops
from_stats = fdf.groupby("From Stop").agg(Boardings=("Total Psgr", "sum")).reset_index().rename(
    columns={"From Stop": "Stop"})
to_stats = (fdf[fdf["To Stop"] != "Unknown"].groupby("To Stop").agg(Alightings=("Total Psgr", "sum"))
            .reset_index().rename(columns={"To Stop": "Stop"}))
stops = pd.merge(from_stats, to_stats, on="Stop", how="outer").fillna(0)
stops = stops[stops["Stop"] != "Unknown"]
stops["Activity"] = stops["Boardings"] + stops["Alightings"]
stops["NetFlow"] = stops["Boardings"] - stops["Alightings"]
stops["Balance"] = np.where(stops["Activity"] > 0, stops["NetFlow"] / stops["Activity"], 0)
stops["Role"] = np.select([stops["Balance"] > 0.2, stops["Balance"] < -0.2],
                          ["Generator", "Attractor"], default="Balanced")
stops = stops.sort_values("Activity", ascending=False)


# ─────────────────────────────────────────────────────────────
# HERO + KPIs + INSIGHTS
# ─────────────────────────────────────────────────────────────
top5 = from_stats[from_stats["Stop"] != "Unknown"].nlargest(5, "Boardings")
stops_html = "".join(
    f'<div class="stop"><div class="stop-name">{r.Stop}</div><div class="stop-val">{int(r.Boardings):,} boarded</div></div>'
    for r in top5.itertuples()
)
st.markdown(
    f'<div class="hero"><div class="hero-title">BusOps Command Center</div>'
    f'<div class="hero-sub">City bus ticket analytics • {DATE_LABEL} • {tickets:,} of {len(df):,} tickets shown</div>'
    f'<div class="routeline">{stops_html}</div></div>',
    unsafe_allow_html=True,
)

below = rate < target
cols = st.columns(6)
cols[0].markdown(kpi("Tickets", f"{tickets:,}"), unsafe_allow_html=True)
cols[1].markdown(kpi("Passengers", f"{pax:,}"), unsafe_allow_html=True)
cols[2].markdown(kpi("Potential revenue", short_inr(potential), f"{short_inr(avg_fare)} per passenger"
                     if avg_fare < 1e5 else ""), unsafe_allow_html=True)
cols[3].markdown(kpi("Collected", short_inr(coll)), unsafe_allow_html=True)
cols[4].markdown(kpi("Collection rate", f"{rate:.1f}%", f"{rate - target:+.1f} pts vs {target}% target",
                     "bad" if below else ""), unsafe_allow_html=True)
cols[5].markdown(kpi("Zero-collection tickets", f"{zero_pct:.1f}%", f"{zero_n:,} tickets",
                     "warn" if zero_pct > 25 else ""), unsafe_allow_html=True)

worst_route = od.loc[od["Uncollected"].idxmax()]
worst_wb = wb.iloc[0]
top_stop = stops.iloc[0]["Stop"] if not stops.empty else "n/a"
st.write("")
i = st.columns(4)
i[0].markdown(insight("Busiest window", f"{peak_window} carries {w_share:.0f}% of passengers"), unsafe_allow_html=True)
i[1].markdown(insight("Biggest revenue leak", f"{worst_route['Route']} ({short_inr(worst_route['Uncollected'])} uncollected)"),
              unsafe_allow_html=True)
i[2].markdown(insight("Lowest-collecting waybill", f"{worst_wb['Waybill']} at {worst_wb['CollRate']:.1f}%"),
              unsafe_allow_html=True)
i[3].markdown(insight("Busiest stop", f"{top_stop}"), unsafe_allow_html=True)
st.write("")


# ─────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────
tab_over, tab_route, tab_time, tab_wb, tab_stop, tab_act = st.tabs(
    ["Revenue leakage", "Routes", "Time of day", "Waybills", "Stops", "Actions & scenarios"])

# ── Revenue leakage ──────────────────────────────────────────
with tab_over:
    c1, c2 = st.columns([1.5, 1])
    with c1:
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_bar(x=hourly["Hour"], y=hourly["Coll"], name="Collected", marker_color=TEAL)
        fig.add_bar(x=hourly["Hour"], y=hourly["Uncoll"], name="Uncollected", marker_color=RED)
        fig.add_trace(go.Scatter(x=hourly["Hour"], y=hourly["Rate"], name="Collection rate %",
                                 mode="lines+markers", line=dict(color=AMBER, width=3)), secondary_y=True)
        fig.update_layout(barmode="stack", title="Collected vs uncollected revenue by hour")
        fig.update_xaxes(title_text="Hour of day", dtick=1)
        fig.update_yaxes(title_text="Amount (₹)", secondary_y=False)
        fig.update_yaxes(title_text="Collection rate (%)", range=[0, 100], secondary_y=True, showgrid=False)
        st.plotly_chart(style(fig, 400), width="stretch")
    with c2:
        fig = px.bar(hourly, x="Hour", y="ZeroPct", title="Share of tickets with zero collection",
                     labels={"ZeroPct": "% of tickets"}, color_discrete_sequence=[RED])
        fig.update_xaxes(dtick=1)
        st.plotly_chart(style(fig, 400), width="stretch")

    m = st.columns(4)
    m[0].markdown(kpi("Total leakage", short_inr(uncoll), "uncollected amount", "bad"), unsafe_allow_html=True)
    m[1].markdown(kpi("Zero-collection tickets", f"{zero_n:,}"), unsafe_allow_html=True)
    m[2].markdown(kpi("Passengers on those tickets", f"{zero_pax:,}"), unsafe_allow_html=True)
    m[3].markdown(kpi("Average fare per passenger", f"₹{avg_fare:.1f}"), unsafe_allow_html=True)
    st.write("")

    l1, l2 = st.columns(2)
    with l1:
        top_leak = od.nlargest(10, "Uncollected")
        fig = px.bar(top_leak, x="Uncollected", y="Route", orientation="h",
                     title="Top 10 routes by uncollected revenue", color_discrete_sequence=[RED])
        fig.update_yaxes(autorange="reversed", title="")
        st.plotly_chart(style(fig, 400), width="stretch")
    with l2:
        top_wb_leak = wb.nlargest(10, "Uncollected")
        fig = px.bar(top_wb_leak, x="Uncollected", y="Waybill", orientation="h",
                     title="Top 10 waybills by uncollected revenue", color_discrete_sequence=[RED])
        fig.update_yaxes(autorange="reversed", title="", type="category")
        st.plotly_chart(style(fig, 400), width="stretch")

# ── Routes ───────────────────────────────────────────────────
with tab_route:
    n = st.slider("Routes to show", 5, 30, 15, key="n_routes")
    top = od.head(n)
    c1, c2 = st.columns([1.3, 1])
    with c1:
        fig = px.bar(top, x="Passengers", y="Route", orientation="h", color="CollRate",
                     color_continuous_scale="RdYlGn", range_color=[0, max(50, float(top["CollRate"].max()))],
                     title=f"Top {len(top)} routes by passengers",
                     labels={"CollRate": "Collection %"}, text_auto=True)
        fig.update_yaxes(autorange="reversed", title="")
        fig.update_traces(textposition="outside")
        st.plotly_chart(style(fig, max(380, 26 * len(top) + 90)), width="stretch")
    with c2:
        st.markdown("#### Route table")
        st.dataframe(
            top[["Route", "Tickets", "Passengers", "TotalAmt", "Collected", "CollRate"]],
            hide_index=True, width="stretch", height=max(380, 26 * len(top) + 90),
            column_config={
                "TotalAmt": st.column_config.NumberColumn("Potential ₹", format="₹%.0f"),
                "Collected": st.column_config.NumberColumn("Collected ₹", format="₹%.0f"),
                "CollRate": st.column_config.ProgressColumn("Collection %", min_value=0, max_value=100,
                                                            format="%.1f%%"),
            })

    st.markdown("#### Corridors (both directions combined)")
    if corr.empty:
        st.info("No two-stop corridors in the current selection.")
    else:
        cc = corr.head(15)
        fig = px.bar(cc, x="Passengers", y="Corridor", orientation="h", color="Potential",
                     color_continuous_scale="Blues", text_auto=True,
                     labels={"Potential": "Potential ₹"}, title="Top corridors by total passengers")
        fig.update_yaxes(autorange="reversed", title="")
        fig.update_traces(textposition="outside")
        st.plotly_chart(style(fig, max(360, 26 * len(cc) + 90)), width="stretch")

# ── Time of day ──────────────────────────────────────────────
with tab_time:
    p1, p2, p3 = st.columns(3)
    p1.markdown(kpi("Peak hour", f"{peak_hour:02d}:00 – {peak_hour + 1:02d}:00",
                    f"{int(pax_by_hour.max()):,} passengers"), unsafe_allow_html=True)
    p2.markdown(kpi("Busiest 3-hour window", peak_window, f"{w_share:.0f}% of passengers"), unsafe_allow_html=True)
    quiet = int(pax_by_hour.idxmin())
    p3.markdown(kpi("Quietest hour", f"{quiet:02d}:00 – {quiet + 1:02d}:00",
                    f"{int(pax_by_hour.min()):,} passengers", "warn"), unsafe_allow_html=True)
    st.write("")

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_bar(x=hourly["Hour"], y=hourly["Pax"], name="Passengers", marker_color=BLUE, opacity=0.85)
    fig.add_trace(go.Scatter(x=hourly["Hour"], y=hourly["Total"], name="Potential revenue",
                             mode="lines+markers", line=dict(color=AMBER, width=3)), secondary_y=True)
    fig.update_layout(title="Passenger volume vs potential revenue by hour")
    fig.update_xaxes(title_text="Hour of day", dtick=1)
    fig.update_yaxes(title_text="Passengers", secondary_y=False)
    fig.update_yaxes(title_text="Potential revenue (₹)", secondary_y=True, showgrid=False)
    st.plotly_chart(style(fig, 400), width="stretch")

    st.markdown("#### Where and when demand peaks")
    k = st.slider("Boarding stops in heatmap", 5, 25, 12, key="n_heat")
    top_b = fdf["From Stop"].value_counts().head(k).index
    heat = (fdf[fdf["From Stop"].isin(top_b)].groupby(["From Stop", "Hour"])["Total Psgr"].sum()
            .unstack(fill_value=0))
    heat = heat.loc[heat.sum(axis=1).sort_values(ascending=False).index]
    fig = px.imshow(heat, aspect="auto", color_continuous_scale="YlGnBu", text_auto=True,
                    labels=dict(x="Hour", y="Boarding stop", color="Passengers"))
    st.plotly_chart(style(fig, max(380, 30 * len(heat) + 90)), width="stretch")

# ── Waybills ─────────────────────────────────────────────────
with tab_wb:
    best, worst = wb.nlargest(3, "CollRate"), wb.nsmallest(3, "CollRate")
    b1, b2 = st.columns(2)
    b1.markdown("**Strongest collection**")
    for r in best.itertuples():
        b1.markdown(f"Waybill `{r.Waybill}`: **{r.CollRate:.1f}%** on {int(r.Passengers):,} passengers")
    b2.markdown("**Needs attention**")
    for r in worst.itertuples():
        b2.markdown(f"Waybill `{r.Waybill}`: **{r.CollRate:.1f}%** on {int(r.Passengers):,} passengers")
    st.write("")

    c1, c2 = st.columns([1, 1])
    with c1:
        fig = px.bar(wb, x="CollRate", y="Waybill", orientation="h", color="CollRate",
                     color_continuous_scale="RdYlGn", range_color=[0, max(50, float(wb["CollRate"].max()))],
                     title="Collection rate by waybill (lowest first)", text_auto=".1f",
                     labels={"CollRate": "Collection %"})
        fig.update_yaxes(autorange="reversed", title="", type="category")
        fig.update_traces(textposition="outside")
        st.plotly_chart(style(fig, max(420, 24 * len(wb) + 90)), width="stretch")
    with c2:
        fig = px.scatter(wb, x="Passengers", y="CollRate", size="TotalAmt", color="ZeroPct",
                         color_continuous_scale="Reds", hover_name="Waybill",
                         title="Volume vs collection rate (bubble = potential revenue)",
                         labels={"CollRate": "Collection %", "ZeroPct": "Zero-coll %"})
        fig.add_hline(y=target, line_dash="dash", line_color=AMBER, annotation_text=f"{target}% target")
        st.plotly_chart(style(fig, 420), width="stretch")

    st.dataframe(
        wb.sort_values("CollRate", ascending=False)[["Waybill", "Tickets", "Passengers", "TotalAmt",
                                                      "Collected", "CollRate", "ZeroPct"]],
        hide_index=True, width="stretch",
        column_config={
            "TotalAmt": st.column_config.NumberColumn("Potential ₹", format="₹%.0f"),
            "Collected": st.column_config.NumberColumn("Collected ₹", format="₹%.0f"),
            "CollRate": st.column_config.ProgressColumn("Collection %", min_value=0, max_value=100, format="%.1f%%"),
            "ZeroPct": st.column_config.ProgressColumn("Zero-coll tickets %", min_value=0, max_value=100,
                                                       format="%.1f%%"),
        })

# ── Stops ────────────────────────────────────────────────────
with tab_stop:
    c1, c2 = st.columns(2)
    with c1:
        tf = stops.nlargest(15, "Boardings")
        fig = px.bar(tf, x="Boardings", y="Stop", orientation="h", title="Top boarding stops",
                     color_discrete_sequence=[BLUE], text_auto=True)
        fig.update_yaxes(autorange="reversed", title="")
        fig.update_traces(textposition="outside")
        st.plotly_chart(style(fig, 480), width="stretch")
    with c2:
        tt = stops.nlargest(15, "Alightings")
        fig = px.bar(tt, x="Alightings", y="Stop", orientation="h", title="Top alighting stops",
                     color_discrete_sequence=[TEAL], text_auto=True)
        fig.update_yaxes(autorange="reversed", title="")
        fig.update_traces(textposition="outside")
        st.plotly_chart(style(fig, 480), width="stretch")

    st.markdown("#### Stop roles")
    st.caption("Generator: more people board than alight (net flow above 20% of activity). "
               "Attractor: the opposite. Balanced: in between.")
    sc = stops.head(40)
    fig = px.scatter(sc, x="Boardings", y="Alightings", size="Activity", color="Role", hover_name="Stop",
                     color_discrete_map={"Generator": BLUE, "Attractor": TEAL, "Balanced": "#9aa9b8"},
                     title="Boardings vs alightings (top 40 stops by activity)")
    mx = float(max(sc["Boardings"].max(), sc["Alightings"].max()))
    fig.add_trace(go.Scatter(x=[0, mx], y=[0, mx], mode="lines", line=dict(color=MIST, dash="dash"),
                             showlegend=False, hoverinfo="skip"))
    st.plotly_chart(style(fig, 440), width="stretch")
    st.dataframe(stops.head(25)[["Stop", "Boardings", "Alightings", "Activity", "NetFlow", "Role"]],
                 hide_index=True, width="stretch")

# ── Actions & scenarios ──────────────────────────────────────
with tab_act:
    st.markdown("### Key findings")
    findings = [
        f"**Collection rate is {rate:.1f}%.** {short_inr(uncoll)} of {short_inr(potential)} potential revenue was not collected.",
        f"**{zero_pct:.1f}% of tickets recorded zero collection**, covering {zero_pax:,} passengers.",
        f"**Peak demand runs {peak_window}**, carrying {w_share:.0f}% of passengers.",
        f"**Waybill collection rates range from {wb['CollRate'].min():.1f}% to {wb['CollRate'].max():.1f}%.**",
    ]
    if not corr.empty:
        findings.insert(2, f"**Strongest corridor:** {corr.iloc[0]['Corridor']} ({int(corr.iloc[0]['Passengers']):,} passengers).")
    if len(stops) >= 2:
        findings.insert(3, f"**Busiest hubs:** {stops.iloc[0]['Stop']} and {stops.iloc[1]['Stop']}.")
    for f in findings:
        st.markdown(f"- {f}")

    st.markdown("---")
    st.markdown("### What-if simulator")
    st.caption("Illustrative only. These numbers come from this data and are not forecasts.")
    goal = st.slider("If the collection rate reached (%)", 10, 100, max(target, int(rate) + 5), 5, key="goal")
    goal_coll = potential * goal / 100
    uplift = goal_coll - coll
    med = float(wb["CollRate"].median())
    lift_low = float((wb.loc[wb["CollRate"] < med, "TotalAmt"] * (med - wb.loc[wb["CollRate"] < med, "CollRate"]) / 100).sum())

    s = st.columns(4)
    s[0].markdown(kpi("Collected now", short_inr(coll)), unsafe_allow_html=True)
    s[1].markdown(kpi(f"Collected at {goal}%", short_inr(goal_coll)), unsafe_allow_html=True)
    s[2].markdown(kpi("Extra per day", short_inr(max(uplift, 0)),
                      f"{short_inr(max(uplift, 0) * 30)} over 30 days"), unsafe_allow_html=True)
    s[3].markdown(kpi("If weaker waybills reach the median", short_inr(lift_low),
                      f"median rate {med:.1f}%"), unsafe_allow_html=True)

    st.markdown("---")
    a1, a2 = st.columns(2)
    with a1:
        st.markdown("#### Next two weeks")
        st.markdown(f"- Review the lowest-collecting waybills, starting with {worst_wb['Waybill']}\n"
                    f"- Audit {worst_route['Route']}, the route with the largest uncollected amount\n"
                    "- Promote UPI and digital payments on the busiest routes\n"
                    "- Show live collection rate to supervisors")
    with a2:
        st.markdown("#### Next one to three months")
        st.markdown(f"- Plan duty allocation around the {peak_window} peak\n"
                    "- Link conductor incentives to collection rate\n"
                    "- Pilot QR ticketing on the top corridors\n"
                    "- Collect several weeks of data before building forecasts")

    with st.expander("About this data"):
        st.markdown(
            f"- Covers **{DATE_LABEL}** with {len(df):,} tickets.\n"
            "- *Total Psgr* is passengers on a ticket, not unique people.\n"
            "- Uncollected amounts may include concessions or free passes.\n"
            "- Reliable forecasting needs multi-week history.")

st.caption(f"BusOps Command Center • SOLVATHON 2026 PS-3 • {tickets:,} tickets • {pax:,} passengers shown")