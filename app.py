"""
BusOps Command Center  |  SOLVATHON 2026 - PS-3
City bus ticket dashboard for department officials.

Everything shown is calculated from the Excel data. Nothing is assumed.

Run:  py -m streamlit run app.py
Data: keep "Ticket data*.xlsx" file(s) in the same folder as this script
      (add more days as extra files and the date comparison and forecast switch on),
      or upload files from the sidebar.
"""

import io
import re
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

REQUIRED = ["Waybill Number", "Ticket Number", "From Stop", "To Stop", "Time",
            "Date", "Total Psgr", "Coll Amt", "Uncoll Amt"]
INK, TEAL, AMBER, RED, BLUE, GREY = "#12263a", "#0f9d8a", "#e8a317", "#d64545", "#2f6fb5", "#9aa9b8"

st.set_page_config(page_title="BusOps Command Center", page_icon="🚌",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=DM+Sans:wght@400;500;700&display=swap');
.stMarkdown, .stCaption, label, p { font-family: 'DM Sans', sans-serif; }
h1, h2, h3, h4 { font-family: 'Space Grotesk', sans-serif !important; color: #12263a; }
.block-container { padding-top: 1.6rem; max-width: 1400px; }
#MainMenu, footer { visibility: hidden; }
.hero { background: #12263a; color: #fff; border-radius: 14px; padding: 1.4rem 1.8rem 1.2rem; margin-bottom: 1.1rem; }
.hero-title { font-family: 'Space Grotesk', sans-serif; font-size: 1.9rem; font-weight: 700; letter-spacing: -0.02em; }
.hero-sub { color: #a9bccd; font-size: 0.95rem; margin-top: 0.15rem; }
.routeline { position: relative; display: flex; margin-top: 1.4rem; }
.routeline::before { content: ""; position: absolute; top: 7px; left: 8%; right: 8%; height: 4px; background: #0f9d8a; border-radius: 2px; }
.stop { flex: 1; text-align: center; position: relative; padding-top: 26px; }
.stop::before { content: ""; position: absolute; top: 0; left: 50%; transform: translateX(-50%);
  width: 18px; height: 18px; border-radius: 50%; background: #12263a; border: 4px solid #e8a317; }
.stop-name { font-weight: 700; font-size: 0.9rem; }
.stop-val { color: #a9bccd; font-size: 0.8rem; }
.kpi { border-left: 4px solid #0f9d8a; padding: 0.2rem 0 0.2rem 0.8rem; }
.kpi.warn { border-left-color: #e8a317; }
.kpi-label { color: #5b6b7b; font-size: 0.82rem; }
.kpi-value { font-family: 'Space Grotesk', sans-serif; color: #12263a; font-size: 1.6rem; font-weight: 700; line-height: 1.2; }
.kpi-note { color: #5b6b7b; font-size: 0.78rem; }
.headline { background: #f1f5f8; border-radius: 10px; padding: 0.9rem 1.1rem; }
.headline li { margin-bottom: 0.25rem; }
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


def fmt_day(d) -> str:
    return pd.Timestamp(d).strftime("%d %b %Y")


def style(fig, h=380):
    fig.update_layout(
        template="plotly_white", height=h, margin=dict(l=8, r=8, t=56, b=8),
        font=dict(family="DM Sans, sans-serif", size=12, color=INK),
        title=dict(font=dict(family="Space Grotesk, sans-serif", size=16), x=0),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="right", x=1),
    )
    return fig


def hbar(d, x, y, title, color=BLUE):
    fig = px.bar(d, x=x, y=y, orientation="h", title=title, color_discrete_sequence=[color], text_auto=True)
    fig.update_yaxes(autorange="reversed", title="", type="category")
    fig.update_traces(textposition="outside", cliponaxis=False)
    return style(fig, max(320, 26 * len(d) + 90))


def clean_stop(s: pd.Series) -> pd.Series:
    s = s.astype("string").str.strip().fillna("Unknown")
    return s.replace({"nan": "Unknown", "": "Unknown", "<NA>": "Unknown", "None": "Unknown"}).astype(str)


def pct(a, b):
    return a / b * 100 if b else 0.0


@st.cache_data(show_spinner="Reading ticket data…")
def load_data(raws: tuple) -> pd.DataFrame:
    frames = []
    for raw in raws:
        part = pd.read_excel(io.BytesIO(raw))
        part.columns = part.columns.astype(str).str.strip()
        missing = [c for c in REQUIRED if c not in part.columns]
        if missing:
            raise ValueError("A file is missing these columns: " + ", ".join(missing))
        frames.append(part)
    df = pd.concat(frames, ignore_index=True)

    df["From Stop"] = clean_stop(df["From Stop"])
    df["To Stop"] = clean_stop(df["To Stop"])
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    t = df["Time"].astype(str).str.strip().str[-8:]
    hour = pd.to_datetime(t, format="%H:%M:%S", errors="coerce")
    if hour.isna().mean() > 0.5:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            hour = pd.to_datetime(df["Time"].astype(str), errors="coerce")
    df["Hour"] = hour.dt.hour
    df["MinOfDay"] = hour.dt.hour * 60 + hour.dt.minute

    for c in ["Total Psgr", "Coll Amt", "Uncoll Amt"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df = df.dropna(subset=["Hour", "Date"]).copy()
    df["Hour"] = df["Hour"].astype(int)
    df["MinOfDay"] = df["MinOfDay"].astype(int)
    df["Day"] = df["Date"].dt.date
    df["Total Amt"] = df["Coll Amt"] + df["Uncoll Amt"]
    df["Route"] = df["From Stop"] + " → " + df["To Stop"]
    df["IsZeroColl"] = df["Coll Amt"] == 0
    return df


def summarise(d: pd.DataFrame) -> dict:
    coll, unc = float(d["Coll Amt"].sum()), float(d["Uncoll Amt"].sum())
    return dict(Tickets=len(d), Passengers=int(d["Total Psgr"].sum()), Collected=coll, Uncollected=unc,
                Total=coll + unc, Rate=pct(coll, coll + unc),
                Zero=pct(int(d["IsZeroColl"].sum()), len(d)))


# ─────────────────────────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────────────────────────
here = Path(__file__).parent
uploads = st.sidebar.file_uploader("Upload Excel file(s)", type=["xlsx"], accept_multiple_files=True,
                                   help="Optional. Without an upload, the app reads every 'Ticket data*.xlsx' "
                                        "in the same folder as app.py.")
if uploads:
    raws, sources = tuple(u.getvalue() for u in uploads), [u.name for u in uploads]
else:
    paths = sorted(p for p in here.glob("Ticket data*.xlsx") if not p.name.startswith("~$"))
    raws, sources = tuple(p.read_bytes() for p in paths), [p.name for p in paths]

if not raws:
    st.error("No ticket data found. Put the Excel file (named like **Ticket data…xlsx**) in the same folder as "
             "app.py, or upload it from the sidebar.")
    st.stop()

try:
    df = load_data(raws)
except Exception as e:
    st.error(f"The data could not be read. {e}")
    st.stop()

if df.empty:
    st.error("The file was read but no rows had a valid date and time.")
    st.stop()

st.sidebar.caption("Data: " + ", ".join(sources))

HMIN, HMAX = int(df["Hour"].min()), int(df["Hour"].max())
DMIN, DMAX = df["Day"].min(), df["Day"].max()
ALL_DAYS = sorted(df["Day"].unique())


# ─────────────────────────────────────────────────────────────
# FILTERS
# ─────────────────────────────────────────────────────────────
def reset_filters():
    st.session_state.f_dates = (DMIN, DMAX)
    st.session_state.f_hours = (HMIN, HMAX)
    st.session_state.f_wb = []
    st.session_state.f_from = []
    st.session_state.f_to = []


st.sidebar.markdown("### Filters")
st.sidebar.caption("Leave a list empty to include everything.")

if len(ALL_DAYS) > 1:
    picked = st.sidebar.date_input("Dates", value=(DMIN, DMAX), min_value=DMIN, max_value=DMAX, key="f_dates")
    d0, d1 = (picked[0], picked[-1]) if isinstance(picked, (tuple, list)) and len(picked) else (DMIN, DMAX)
else:
    d0, d1 = DMIN, DMAX

hours = st.sidebar.slider("Hours of the day", HMIN, HMAX, (HMIN, HMAX), key="f_hours") if HMIN < HMAX else (HMIN, HMAX)
sel_wb = st.sidebar.multiselect("Waybill", sorted(df["Waybill Number"].unique()), key="f_wb", placeholder="All waybills")
sel_from = st.sidebar.multiselect("Boarding stop", sorted(df["From Stop"].unique()), key="f_from", placeholder="All stops")
sel_to = st.sidebar.multiselect("Destination stop", sorted(df["To Stop"].unique()), key="f_to", placeholder="All stops")
st.sidebar.button("Reset filters", on_click=reset_filters, width="stretch")

mask = df["Day"].between(d0, d1) & df["Hour"].between(hours[0], hours[1])
if sel_wb:
    mask &= df["Waybill Number"].isin(sel_wb)
if sel_from:
    mask &= df["From Stop"].isin(sel_from)
if sel_to:
    mask &= df["To Stop"].isin(sel_to)
fdf = df[mask].copy()

if fdf.empty:
    st.warning("No tickets match these filters. Widen the selection or click **Reset filters** in the sidebar.")
    st.stop()

st.sidebar.download_button("Download filtered data (CSV)", fdf.to_csv(index=False).encode("utf-8"),
                           file_name="filtered_tickets.csv", mime="text/csv", width="stretch")


# ─────────────────────────────────────────────────────────────
# CALCULATIONS (all from the filtered data)
# ─────────────────────────────────────────────────────────────
S = summarise(fdf)
days_in_view = sorted(fdf["Day"].unique())
n_days = len(days_in_view)
DAY_LABEL = fmt_day(days_in_view[0]) if n_days == 1 else f"{fmt_day(days_in_view[0])} to {fmt_day(days_in_view[-1])}"

hour_range = range(int(fdf["Hour"].min()), int(fdf["Hour"].max()) + 1)
hourly = (fdf.groupby("Hour").agg(Tickets=("Ticket Number", "count"), Pax=("Total Psgr", "sum"),
                                  Coll=("Coll Amt", "sum"), Uncoll=("Uncoll Amt", "sum"),
                                  Zero=("IsZeroColl", "sum"))
          .reindex(hour_range, fill_value=0).rename_axis("Hour").reset_index())
hourly["Total"] = hourly["Coll"] + hourly["Uncoll"]
hourly["Rate"] = np.where(hourly["Total"] > 0, hourly["Coll"] / hourly["Total"] * 100, np.nan)
hourly["ZeroPct"] = np.where(hourly["Tickets"] > 0, hourly["Zero"] / hourly["Tickets"] * 100, 0)

pax_h = hourly.set_index("Hour")["Pax"]
roll = pax_h.rolling(3, min_periods=1).sum()
w_end = int(roll.idxmax())
w_start = max(w_end - 2, int(pax_h.index.min()))
peak_window = f"{w_start:02d}:00 – {w_end + 1:02d}:00"
peak_share = pct(roll.max(), pax_h.sum())

daily = (fdf.groupby("Day").agg(Tickets=("Ticket Number", "count"), Pax=("Total Psgr", "sum"),
                                Coll=("Coll Amt", "sum"), Uncoll=("Uncoll Amt", "sum")).sort_index())
daily["Total"] = daily["Coll"] + daily["Uncoll"]
daily["Rate"] = np.where(daily["Total"] > 0, daily["Coll"] / daily["Total"] * 100, np.nan)

origins = (fdf.groupby("From Stop").agg(Passengers=("Total Psgr", "sum"), Tickets=("Ticket Number", "count"),
                                        Collected=("Coll Amt", "sum")).reset_index().rename(columns={"From Stop": "Stop"}))
dests = (fdf[fdf["To Stop"] != "Unknown"].groupby("To Stop")
         .agg(Passengers=("Total Psgr", "sum"), Tickets=("Ticket Number", "count"), Collected=("Coll Amt", "sum"))
         .reset_index().rename(columns={"To Stop": "Stop"}))

od = fdf.groupby("Route").agg(Tickets=("Ticket Number", "count"), Passengers=("Total Psgr", "sum"),
                              Collected=("Coll Amt", "sum"), Uncollected=("Uncoll Amt", "sum"),
                              Total=("Total Amt", "sum")).reset_index()
od["Rate"] = np.where(od["Total"] > 0, od["Collected"] / od["Total"] * 100, 0.0)
od = od.sort_values("Passengers", ascending=False)

wb = fdf.groupby("Waybill Number").agg(Tickets=("Ticket Number", "count"), Passengers=("Total Psgr", "sum"),
                                       Collected=("Coll Amt", "sum"), Uncollected=("Uncoll Amt", "sum"),
                                       Total=("Total Amt", "sum"), Zero=("IsZeroColl", "sum")).reset_index()
wb["Rate"] = np.where(wb["Total"] > 0, wb["Collected"] / wb["Total"] * 100, 0.0)
wb["ZeroPct"] = wb["Zero"] / wb["Tickets"] * 100
wb["Waybill"] = wb["Waybill Number"].astype(str)
wb = wb.sort_values("Rate")


# ─────────────────────────────────────────────────────────────
# HEADER, SUMMARY NUMBERS, PLAIN-LANGUAGE HEADLINES
# ─────────────────────────────────────────────────────────────
top5 = origins[origins["Stop"] != "Unknown"].nlargest(5, "Passengers")
stops_html = "".join(f'<div class="stop"><div class="stop-name">{r.Stop}</div>'
                     f'<div class="stop-val">{int(r.Passengers):,} passengers</div></div>' for r in top5.itertuples())
st.markdown(
    f'<div class="hero"><div class="hero-title">BusOps Command Center</div>'
    f'<div class="hero-sub">City bus ticket dashboard • {DAY_LABEL} • {S["Tickets"]:,} of {len(df):,} tickets shown'
    f' • busiest boarding stops below</div><div class="routeline">{stops_html}</div></div>',
    unsafe_allow_html=True)

k = st.columns(6)
k[0].markdown(kpi("Tickets issued", f"{S['Tickets']:,}"), unsafe_allow_html=True)
k[1].markdown(kpi("Passengers", f"{S['Passengers']:,}"), unsafe_allow_html=True)
k[2].markdown(kpi("Total fare value", short_inr(S["Total"]), "collected + uncollected"), unsafe_allow_html=True)
k[3].markdown(kpi("Collection recorded", short_inr(S["Collected"])), unsafe_allow_html=True)
k[4].markdown(kpi("Uncollected (as recorded)", short_inr(S["Uncollected"]), f"{100 - S['Rate']:.1f}% of fare value",
                   "warn"), unsafe_allow_html=True)
k[5].markdown(kpi("Collection rate", f"{S['Rate']:.1f}%", "collected ÷ total fare value"), unsafe_allow_html=True)
st.write("")

top_route = od.iloc[0]
top_dest = dests.nlargest(1, "Passengers").iloc[0] if not dests.empty else None
top_orig = origins[origins["Stop"] != "Unknown"].nlargest(1, "Passengers").iloc[0]
low_wb, high_wb = wb.iloc[0], wb.iloc[-1]

lines = [
    f"**{S['Tickets']:,} tickets** carried **{S['Passengers']:,} passengers**. Recorded collection is "
    f"**{short_inr(S['Collected'])}** out of {short_inr(S['Total'])} total fare value (**{S['Rate']:.1f}%**).",
    f"**Busiest time:** {peak_window}, which carries {peak_share:.0f}% of passengers.",
    f"**Busiest route:** {top_route['Route']} ({int(top_route['Passengers']):,} passengers).",
    f"**Busiest boarding stop:** {top_orig['Stop']}"
    + (f". **Busiest destination:** {top_dest['Stop']}." if top_dest is not None else "."),
    f"**{S['Zero']:.1f}% of tickets** have zero collection recorded.",
    f"**Waybill collection rate** ranges from {low_wb['Rate']:.1f}% (waybill {low_wb['Waybill']}) "
    f"to {high_wb['Rate']:.1f}% (waybill {high_wb['Waybill']}).",
]
st.markdown('<div class="headline"><b>Key points from the data</b><ul>' +
            "".join(f"<li>{re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', x)}</li>" for x in lines) +
            "</ul></div>", unsafe_allow_html=True)
st.write("")


# ─────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────
t_sum, t_live, t_route, t_demand, t_rev, t_cmp, t_pred = st.tabs([
    "Summary", "Live tracking", "Routes & destinations", "Passenger demand", "Revenue & collection",
    "Compare dates", "Predictions"])

# ── Summary ──────────────────────────────────────────────────
with t_sum:
    c1, c2 = st.columns([1.4, 1])
    with c1:
        fig = go.Figure()
        fig.add_bar(x=hourly["Hour"], y=hourly["Pax"], name="Passengers", marker_color=BLUE)
        fig.update_layout(title="Passengers by hour of day")
        fig.update_xaxes(title_text="Hour of day", dtick=1)
        fig.update_yaxes(title_text="Passengers")
        st.plotly_chart(style(fig, 360), width="stretch")
    with c2:
        fig = go.Figure(go.Pie(labels=["Collected", "Uncollected"], values=[S["Collected"], S["Uncollected"]],
                               hole=0.55, marker=dict(colors=[TEAL, RED]), sort=False))
        fig.update_layout(title="Share of total fare value")
        st.plotly_chart(style(fig, 360), width="stretch")

    c3, c4 = st.columns(2)
    with c3:
        st.plotly_chart(hbar(od.head(5), "Passengers", "Route", "Top 5 routes by passengers"), width="stretch")
    with c4:
        st.plotly_chart(hbar(dests.nlargest(5, "Passengers"), "Passengers", "Stop", "Top 5 destinations by passengers",
                             TEAL), width="stretch")

# ── Routes & destinations ────────────────────────────────────
with t_route:
    metric = st.radio("Rank by", ["Passengers", "Tickets", "Collected"], horizontal=True,
                      format_func=lambda m: {"Collected": "Collected amount"}.get(m, m))
    n = st.slider("How many to show", 5, 25, 10, key="n_show")

    a, b = st.columns(2)
    with a:
        st.plotly_chart(hbar(origins[origins["Stop"] != "Unknown"].nlargest(n, metric), metric, "Stop",
                             f"Top {n} boarding stops", BLUE), width="stretch")
    with b:
        st.plotly_chart(hbar(dests.nlargest(n, metric), metric, "Stop", f"Top {n} destinations", TEAL),
                        width="stretch")

    st.markdown("#### Routes (boarding stop → destination)")
    top_od = od.nlargest(n, metric)
    r1, r2 = st.columns([1, 1.1])
    with r1:
        st.plotly_chart(hbar(top_od, metric, "Route", f"Top {n} routes", AMBER), width="stretch")
    with r2:
        st.dataframe(top_od[["Route", "Tickets", "Passengers", "Total", "Collected", "Rate"]], hide_index=True,
                     width="stretch", height=max(320, 26 * len(top_od) + 90),
                     column_config={
                         "Total": st.column_config.NumberColumn("Fare value ₹", format="₹%.0f"),
                         "Collected": st.column_config.NumberColumn("Collected ₹", format="₹%.0f"),
                         "Rate": st.column_config.ProgressColumn("Collection %", min_value=0, max_value=100,
                                                                 format="%.1f%%")})

    st.markdown("#### Explore one stop")
    pick = st.selectbox("Choose a stop", ["(none)"] + sorted(set(origins["Stop"]) | set(dests["Stop"])),
                        help="See where passengers from this stop travel to, and where its passengers come from.")
    if pick != "(none)":
        out = (fdf[fdf["From Stop"] == pick].groupby("To Stop")["Total Psgr"].sum().nlargest(10)
               .reset_index().rename(columns={"To Stop": "Stop", "Total Psgr": "Passengers"}))
        inc = (fdf[fdf["To Stop"] == pick].groupby("From Stop")["Total Psgr"].sum().nlargest(10)
               .reset_index().rename(columns={"From Stop": "Stop", "Total Psgr": "Passengers"}))
        e1, e2 = st.columns(2)
        with e1:
            if out.empty:
                st.info("No passengers board here in the current selection.")
            else:
                st.plotly_chart(hbar(out, "Passengers", "Stop", f"Where passengers from {pick} go", TEAL),
                                width="stretch")
        with e2:
            if inc.empty:
                st.info("No passengers travel to this stop in the current selection.")
            else:
                st.plotly_chart(hbar(inc, "Passengers", "Stop", f"Where passengers to {pick} come from", BLUE),
                                width="stretch")

# ── Passenger demand ─────────────────────────────────────────
with t_demand:
    p = st.columns(3)
    p[0].markdown(kpi("Peak hour", f"{int(pax_h.idxmax()):02d}:00 – {int(pax_h.idxmax()) + 1:02d}:00",
                      f"{int(pax_h.max()):,} passengers"), unsafe_allow_html=True)
    p[1].markdown(kpi("Busiest 3-hour window", peak_window, f"{peak_share:.0f}% of passengers"), unsafe_allow_html=True)
    p[2].markdown(kpi("Quietest hour", f"{int(pax_h.idxmin()):02d}:00 – {int(pax_h.idxmin()) + 1:02d}:00",
                      f"{int(pax_h.min()):,} passengers", "warn"), unsafe_allow_html=True)
    st.write("")

    d1c, d2c = st.columns([1.5, 1])
    with d1c:
        fig = px.bar(hourly, x="Hour", y="Pax", title="Passengers by hour", color_discrete_sequence=[BLUE],
                     labels={"Pax": "Passengers"})
        fig.update_xaxes(dtick=1)
        st.plotly_chart(style(fig, 360), width="stretch")
    with d2c:
        size = fdf["Total Psgr"].astype(int).clip(lower=0, upper=6).value_counts().sort_index()
        gs = pd.DataFrame({"Passengers on ticket": [str(i) if i < 6 else "6+" for i in size.index],
                           "Tickets": size.values})
        fig = px.bar(gs, x="Passengers on ticket", y="Tickets", title="Tickets by group size",
                     color_discrete_sequence=[TEAL], text_auto=True)
        st.plotly_chart(style(fig, 360), width="stretch")

    st.markdown("#### When each boarding stop is busy")
    kk = st.slider("Number of stops", 5, 25, 12, key="n_heat")
    top_b = fdf["From Stop"].value_counts().head(kk).index
    heat = (fdf[fdf["From Stop"].isin(top_b)].groupby(["From Stop", "Hour"])["Total Psgr"].sum().unstack(fill_value=0))
    heat = heat.loc[heat.sum(axis=1).sort_values(ascending=False).index]
    fig = px.imshow(heat, aspect="auto", color_continuous_scale="YlGnBu", text_auto=True,
                    labels=dict(x="Hour", y="Boarding stop", color="Passengers"))
    st.plotly_chart(style(fig, max(380, 30 * len(heat) + 90)), width="stretch")

# ── Revenue & collection ─────────────────────────────────────
with t_rev:
    if n_days > 1:
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_bar(x=daily.index, y=daily["Coll"], name="Collected", marker_color=TEAL)
        fig.add_bar(x=daily.index, y=daily["Uncoll"], name="Uncollected", marker_color=RED)
        fig.add_trace(go.Scatter(x=daily.index, y=daily["Rate"], name="Collection rate %", mode="lines+markers",
                                 line=dict(color=AMBER, width=3)), secondary_y=True)
        fig.update_layout(barmode="stack", title="Daily collection trend")
        fig.update_yaxes(title_text="Amount (₹)", secondary_y=False)
        fig.update_yaxes(title_text="Collection rate (%)", range=[0, 100], secondary_y=True, showgrid=False)
        st.plotly_chart(style(fig, 380), width="stretch")

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_bar(x=hourly["Hour"], y=hourly["Coll"], name="Collected", marker_color=TEAL)
    fig.add_bar(x=hourly["Hour"], y=hourly["Uncoll"], name="Uncollected", marker_color=RED)
    fig.add_trace(go.Scatter(x=hourly["Hour"], y=hourly["Rate"], name="Collection rate %", mode="lines+markers",
                             line=dict(color=AMBER, width=3)), secondary_y=True)
    fig.update_layout(barmode="stack", title="Collection by hour of day")
    fig.update_xaxes(title_text="Hour of day", dtick=1)
    fig.update_yaxes(title_text="Amount (₹)", secondary_y=False)
    fig.update_yaxes(title_text="Collection rate (%)", range=[0, 100], secondary_y=True, showgrid=False)
    st.plotly_chart(style(fig, 380), width="stretch")

    m = st.columns(4)
    m[0].markdown(kpi("Collected", short_inr(S["Collected"])), unsafe_allow_html=True)
    m[1].markdown(kpi("Uncollected (as recorded)", short_inr(S["Uncollected"]), "", "warn"), unsafe_allow_html=True)
    m[2].markdown(kpi("Tickets with zero collection", f"{int(fdf['IsZeroColl'].sum()):,}", f"{S['Zero']:.1f}% of tickets"),
                  unsafe_allow_html=True)
    m[3].markdown(kpi("Average fare value per passenger", f"₹{S['Total'] / S['Passengers']:.1f}" if S["Passengers"] else "n/a"),
                  unsafe_allow_html=True)
    st.caption("“Uncollected” is the Uncoll Amt column as recorded in the data. It may include concessions or free "
               "passes; the data does not say.")

    st.markdown("#### Collection by waybill")
    w1, w2 = st.columns([1, 1])
    with w1:
        fig = px.bar(wb, x="Rate", y="Waybill", orientation="h", color="Rate", color_continuous_scale="RdYlGn",
                     title="Collection rate by waybill (lowest first)", text_auto=".1f",
                     labels={"Rate": "Collection %"})
        fig.update_yaxes(autorange="reversed", title="", type="category")
        fig.update_traces(textposition="outside", cliponaxis=False)
        st.plotly_chart(style(fig, max(380, 24 * len(wb) + 90)), width="stretch")
    with w2:
        st.dataframe(wb.sort_values("Rate", ascending=False)[["Waybill", "Tickets", "Passengers", "Total",
                                                              "Collected", "Rate", "ZeroPct"]],
                     hide_index=True, width="stretch", height=max(380, 24 * len(wb) + 90),
                     column_config={
                         "Total": st.column_config.NumberColumn("Fare value ₹", format="₹%.0f"),
                         "Collected": st.column_config.NumberColumn("Collected ₹", format="₹%.0f"),
                         "Rate": st.column_config.ProgressColumn("Collection %", min_value=0, max_value=100,
                                                                 format="%.1f%%"),
                         "ZeroPct": st.column_config.ProgressColumn("Zero-collection tickets %", min_value=0,
                                                                    max_value=100, format="%.1f%%")})

# ── Compare dates ────────────────────────────────────────────
def compare_block(label_a, da, label_b, db):
    sa, sb = summarise(da), summarise(db)
    cols = st.columns(5)
    items = [("Tickets", "Tickets", "{:,.0f}"), ("Passengers", "Passengers", "{:,.0f}"),
             ("Collected", "Collected", None), ("Uncollected", "Uncollected", None), ("Collection rate", "Rate", "{:.1f}%")]
    for col, (name, key, f) in zip(cols, items):
        va, vb = sa[key], sb[key]
        shown = f.format(vb) if f else short_inr(vb)
        if key == "Rate":
            delta = f"{vb - va:+.1f} pts"
        else:
            delta = f"{(vb - va) / va * 100:+.1f}%" if va else None
        col.metric(f"{name} ({label_b})", shown, delta, help=f"Change compared with {label_a}")
    return sa, sb


with t_cmp:
    if len(ALL_DAYS) >= 2:
        st.caption("Pick any two days from the data. This tab uses your hour, waybill and stop filters, "
                   "but not the date filter.")
        base = df[df["Hour"].between(*hours)]
        if sel_wb:
            base = base[base["Waybill Number"].isin(sel_wb)]
        if sel_from:
            base = base[base["From Stop"].isin(sel_from)]
        if sel_to:
            base = base[base["To Stop"].isin(sel_to)]
        avail = sorted(base["Day"].unique())
        if len(avail) < 2:
            st.info("Fewer than two days have tickets with these filters. Widen the filters.")
        else:
            ca, cb = st.columns(2)
            day_a = ca.selectbox("Day A (baseline)", avail, index=0, format_func=fmt_day)
            day_b = cb.selectbox("Day B (compared)", avail, index=len(avail) - 1, format_func=fmt_day)
            da, db = base[base["Day"] == day_a], base[base["Day"] == day_b]
            compare_block(fmt_day(day_a), da, fmt_day(day_b), db)

            ha = da.groupby("Hour")["Total Psgr"].sum()
            hb = db.groupby("Hour")["Total Psgr"].sum()
            idx = range(int(base["Hour"].min()), int(base["Hour"].max()) + 1)
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=list(idx), y=ha.reindex(idx, fill_value=0), name=fmt_day(day_a),
                                     mode="lines+markers", line=dict(color=GREY, width=3)))
            fig.add_trace(go.Scatter(x=list(idx), y=hb.reindex(idx, fill_value=0), name=fmt_day(day_b),
                                     mode="lines+markers", line=dict(color=BLUE, width=3)))
            fig.update_layout(title="Passengers by hour: Day A vs Day B")
            fig.update_xaxes(title_text="Hour of day", dtick=1)
            st.plotly_chart(style(fig, 380), width="stretch")

            ra = da.groupby("Route")["Total Psgr"].sum().rename(fmt_day(day_a))
            rb = db.groupby("Route")["Total Psgr"].sum().rename(fmt_day(day_b))
            rt = pd.concat([ra, rb], axis=1).fillna(0)
            rt["Change"] = rt.iloc[:, 1] - rt.iloc[:, 0]
            rt = rt.reindex(rt["Change"].abs().sort_values(ascending=False).index).head(10).reset_index()
            st.markdown("#### Routes with the biggest change in passengers")
            st.dataframe(rt, hide_index=True, width="stretch")
    else:
        only = fmt_day(ALL_DAYS[0])
        st.info(f"The data covers one day ({only}), so two dates cannot be compared yet. "
                "Add more days as extra Excel files (same columns) in the app folder, or upload them in the sidebar. "
                "This tab then lets you compare any two dates. Meanwhile, you can compare two parts of the day below.")
        s1, s2 = st.columns(2)
        ra_ = s1.slider("Period A (hours)", HMIN, HMAX, (HMIN, max(HMIN, (HMIN + HMAX) // 2 - 1)), key="pa") \
            if HMIN < HMAX else (HMIN, HMAX)
        rb_ = s2.slider("Period B (hours)", HMIN, HMAX, (min(HMAX, (HMIN + HMAX) // 2), HMAX), key="pb") \
            if HMIN < HMAX else (HMIN, HMAX)
        base = df[df["Day"].between(d0, d1)]
        da, db = base[base["Hour"].between(*ra_)], base[base["Hour"].between(*rb_)]
        if da.empty or db.empty:
            st.info("One of the periods has no tickets. Adjust the sliders.")
        else:
            compare_block(f"{ra_[0]:02d}–{ra_[1] + 1:02d}h", da, f"{rb_[0]:02d}–{rb_[1] + 1:02d}h", db)

# ── Predictions ──────────────────────────────────────────────
with t_pred:
    st.caption("Simple estimates built only from the data. They are guides for planning, not guarantees.")
    all_daily = (df.groupby("Day").agg(Pax=("Total Psgr", "sum"), Tickets=("Ticket Number", "count"),
                                       Coll=("Coll Amt", "sum"), Total=("Total Amt", "sum")).sort_index())
    if len(all_daily) >= 3:
        recent = all_daily.tail(7)
        nxt = pd.Timestamp(all_daily.index.max()) + pd.Timedelta(days=1)
        st.markdown(f"#### Expected for {fmt_day(nxt)}")
        st.caption(f"Based on the average of the last {len(recent)} days in the data. The range is the lowest "
                   "and highest of those days.")
        f = st.columns(3)
        f[0].markdown(kpi("Expected passengers", f"{recent['Pax'].mean():,.0f}",
                          f"range {recent['Pax'].min():,.0f} – {recent['Pax'].max():,.0f}"), unsafe_allow_html=True)
        f[1].markdown(kpi("Expected fare value", short_inr(recent["Total"].mean()),
                          f"range {short_inr(recent['Total'].min())} – {short_inr(recent['Total'].max())}"),
                      unsafe_allow_html=True)
        f[2].markdown(kpi("Expected collection", short_inr(recent["Coll"].mean()),
                          f"range {short_inr(recent['Coll'].min())} – {short_inr(recent['Coll'].max())}"),
                      unsafe_allow_html=True)
        st.write("")

        fig = go.Figure()
        fig.add_bar(x=list(all_daily.index), y=all_daily["Pax"], name="Actual", marker_color=BLUE)
        fig.add_bar(x=[nxt.date()], y=[recent["Pax"].mean()], name="Expected", marker_color=AMBER)
        fig.update_layout(title="Daily passengers: actual and expected")
        st.plotly_chart(style(fig, 360), width="stretch")

        prof = df[df["Day"].isin(recent.index)].groupby(["Day", "Hour"])["Total Psgr"].sum().unstack(fill_value=0).mean()
        fig = px.bar(x=prof.index, y=prof.values, title="Expected passengers by hour",
                     labels={"x": "Hour of day", "y": "Passengers"}, color_discrete_sequence=[AMBER])
        fig.update_xaxes(dtick=1)
        st.plotly_chart(style(fig, 340), width="stretch")

        if len(all_daily) >= 5:
            slope = np.polyfit(np.arange(len(all_daily)), all_daily["Pax"].values, 1)[0]
            direction = "rising" if slope > 0 else "falling"
            st.info(f"Over {len(all_daily)} days, daily passengers are {direction} by about {abs(slope):,.0f} per day.")
    else:
        st.info(f"The data has {len(all_daily)} day(s). Day-ahead forecasts need at least 3 days. Add more days as "
                "extra Excel files and this tab will switch on automatically. Until then, here is what today’s "
                "data already tells us.")

        thr = pax_h.quantile(0.75)
        hi_hours = [int(h) for h in pax_h.index if pax_h[h] >= thr]
        st.markdown("#### High-demand hours to plan duties around")
        st.write("Hours in the top quarter of passenger volume: "
                 + ", ".join(f"{h:02d}:00–{h + 1:02d}:00" for h in hi_hours))
        fig = go.Figure(go.Bar(x=pax_h.index, y=pax_h.values,
                               marker_color=[AMBER if h in hi_hours else BLUE for h in pax_h.index]))
        fig.update_layout(title="Passengers by hour (high-demand hours highlighted)", showlegend=False)
        fig.update_xaxes(title_text="Hour of day", dtick=1)
        st.plotly_chart(style(fig, 340), width="stretch")

        st.markdown("#### What if weaker hours matched the day’s overall collection rate?")
        weak = hourly[(hourly["Total"] > 0) & (hourly["Rate"] < S["Rate"])]
        gain = float((weak["Total"] * (S["Rate"] - weak["Rate"]) / 100).sum())
        if weak.empty:
            st.write("No hour is below the overall collection rate.")
        else:
            g = st.columns(3)
            g[0].markdown(kpi("Hours below the overall rate", f"{len(weak)}"), unsafe_allow_html=True)
            g[1].markdown(kpi("Extra collection possible", short_inr(gain),
                              f"{pct(gain, S['Collected']):.1f}% more than recorded"), unsafe_allow_html=True)
            g[2].markdown(kpi("Overall rate used", f"{S['Rate']:.1f}%"), unsafe_allow_html=True)
            st.caption("This is a what-if using today’s own numbers, not a forecast.")

# ── Live tracking ────────────────────────────────────────────
def hhmm(m) -> str:
    m = int(m)
    return f"{m // 60:02d}:{m % 60:02d}"


with t_live:
    st.markdown("#### Live bus tracking")
    gps = st.file_uploader(
        "Real GPS feed (optional CSV)", type=["csv"], key="gps_csv",
        help="Needs columns: Waybill Number, Latitude, Longitude. Optional: Timestamp, Speed. "
             "Upload a fresh export to refresh the map.")

    feed_ready = False
    if gps is not None:
        try:
            feed = pd.read_csv(gps)
            feed.columns = feed.columns.str.strip()
            lack = [c for c in ["Waybill Number", "Latitude", "Longitude"] if c not in feed.columns]
            if lack:
                raise ValueError("Missing columns: " + ", ".join(lack))
            feed["Latitude"] = pd.to_numeric(feed["Latitude"], errors="coerce")
            feed["Longitude"] = pd.to_numeric(feed["Longitude"], errors="coerce")
            feed = feed.dropna(subset=["Latitude", "Longitude"])
            if "Timestamp" in feed.columns:
                feed["Timestamp"] = pd.to_datetime(feed["Timestamp"], errors="coerce")
                feed = feed.sort_values("Timestamp", kind="stable")
            latest = feed.groupby("Waybill Number").tail(1).copy()
            if latest.empty:
                raise ValueError("The file has no valid positions.")
            latest["Waybill"] = latest["Waybill Number"].astype(str)
            st.success(f"Live feed loaded: {len(latest)} buses (latest position of each).")
            fig = px.scatter_map(latest, lat="Latitude", lon="Longitude", hover_name="Waybill", zoom=11, height=520)
            fig.update_traces(marker=dict(size=14, color=AMBER))
            fig.update_layout(map_style="open-street-map", margin=dict(l=0, r=0, t=0, b=0))
            st.plotly_chart(fig, width="stretch")
            st.dataframe(latest.drop(columns="Waybill"), hide_index=True, width="stretch")
            feed_ready = True
        except Exception as e:
            st.error(f"The GPS file could not be used. {e}")

    if not feed_ready:
        st.warning("**No GPS feed connected.** Bus positions below are estimated from ticket records: the stop "
                   "where each bus last issued a ticket. This is not real-time tracking. Upload a GPS CSV above "
                   "to see actual positions on a map.")

        day_pick = (st.selectbox("Day", ALL_DAYS, index=len(ALL_DAYS) - 1, format_func=fmt_day)
                    if len(ALL_DAYS) > 1 else ALL_DAYS[0])
        base = df[df["Day"] == day_pick]
        lo, hi = int(base["MinOfDay"].min() // 5 * 5), int(base["MinOfDay"].max() // 5 * 5 + 5)
        opts = list(range(lo, hi + 1, 5))
        c1, c2 = st.columns([2, 1])
        now = c1.select_slider("Time of day", options=opts, value=opts[len(opts) // 2], format_func=hhmm)
        idle_after = c2.slider("Idle after (minutes without a ticket)", 10, 120, 30, 5)

        upto = base[base["MinOfDay"] <= now].sort_values("MinOfDay", kind="stable")
        last = upto.groupby("Waybill Number").tail(1).set_index("Waybill Number")
        sofar = upto.groupby("Waybill Number").agg(Tickets=("Ticket Number", "count"), Pax=("Total Psgr", "sum"),
                                                   Coll=("Coll Amt", "sum"))
        rows = []
        for w in sorted(base["Waybill Number"].unique()):
            if w in last.index:
                r = last.loc[w]
                gap = now - int(r["MinOfDay"])
                rows.append({"Waybill": str(w), "Status": "🟢 Running" if gap <= idle_after else "🟡 Idle",
                             "Last known stop": r["From Stop"], "Heading to": r["To Stop"],
                             "Last ticket": hhmm(r["MinOfDay"]), "Minutes ago": gap,
                             "Tickets so far": int(sofar.loc[w, "Tickets"]),
                             "Passengers so far": int(sofar.loc[w, "Pax"]),
                             "Collected so far ₹": float(sofar.loc[w, "Coll"])})
            else:
                rows.append({"Waybill": str(w), "Status": "⚪ Not started", "Last known stop": "-",
                             "Heading to": "-", "Last ticket": "-", "Minutes ago": None, "Tickets so far": 0,
                             "Passengers so far": 0, "Collected so far ₹": 0.0})
        fleet = pd.DataFrame(rows)
        order = {"🟢 Running": 0, "🟡 Idle": 1, "⚪ Not started": 2}
        fleet = fleet.sort_values("Status", key=lambda s: s.map(order), kind="stable").reset_index(drop=True)

        q = st.columns(4)
        q[0].markdown(kpi("Time shown", hhmm(now)), unsafe_allow_html=True)
        q[1].markdown(kpi("Running", f"{(fleet['Status'] == '🟢 Running').sum()}"), unsafe_allow_html=True)
        q[2].markdown(kpi("Idle", f"{(fleet['Status'] == '🟡 Idle').sum()}", "", "warn"), unsafe_allow_html=True)
        q[3].markdown(kpi("Not started", f"{(fleet['Status'] == '⚪ Not started').sum()}"), unsafe_allow_html=True)
        st.write("")

        st.dataframe(fleet, hide_index=True, width="stretch",
                     column_config={"Collected so far ₹": st.column_config.NumberColumn(format="₹%.0f")})

        st.markdown("#### Follow one bus")
        sel = st.selectbox("Waybill", fleet["Waybill"].tolist())
        wbd = base[base["Waybill Number"].astype(str) == sel]
        info = fleet[fleet["Waybill"] == sel].iloc[0]
        seq = (wbd[wbd["From Stop"] != "Unknown"].groupby("From Stop")
               .agg(T=("MinOfDay", "median"), N=("Ticket Number", "count"), First=("MinOfDay", "min")).reset_index())
        keep = seq.nlargest(14, "N")
        if info["Last known stop"] in set(seq["From Stop"]) and info["Last known stop"] not in set(keep["From Stop"]):
            keep = pd.concat([keep, seq[seq["From Stop"] == info["Last known stop"]]])
        keep = keep.sort_values("T").reset_index(drop=True)

        if keep.empty:
            st.info("This bus has no ticket records with a known boarding stop.")
        else:
            visited = keep["First"] <= now
            pos = [TEAL if v else GREY for v in visited]
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=list(keep.index), y=[0] * len(keep), mode="lines",
                                     line=dict(color=GREY, width=4), hoverinfo="skip", showlegend=False))
            fig.add_trace(go.Scatter(
                x=list(keep.index), y=[0] * len(keep), mode="markers+text", text=keep["From Stop"],
                textposition=["top center" if i % 2 == 0 else "bottom center" for i in range(len(keep))],
                marker=dict(size=16, color=pos, line=dict(color=INK, width=2)),
                customdata=np.stack([keep["N"], keep["T"].map(hhmm)], axis=-1),
                hovertemplate="%{text}<br>%{customdata[0]} tickets, typical time %{customdata[1]}<extra></extra>",
                showlegend=False))
            bus_idx = keep.index[keep["From Stop"] == info["Last known stop"]]
            if len(bus_idx):
                fig.add_trace(go.Scatter(x=[int(bus_idx[0])], y=[0], mode="markers", name="Bus (last known stop)",
                                         marker=dict(size=28, symbol="diamond", color=AMBER,
                                                     line=dict(color=INK, width=2))))
            fig.update_yaxes(visible=False, range=[-1, 1])
            fig.update_xaxes(visible=False, range=[-0.7, len(keep) - 0.3])
            fig.update_layout(title=f"Waybill {sel}: stops in typical order (teal = tickets already issued by {hhmm(now)})")
            st.plotly_chart(style(fig, 320), width="stretch")

        recent = wbd[wbd["MinOfDay"] <= now].sort_values("MinOfDay", kind="stable").tail(10)
        if recent.empty:
            st.caption("No tickets issued yet on this waybill at the chosen time.")
        else:
            st.markdown("**Latest tickets on this bus**")
            st.dataframe(recent.assign(Time=recent["MinOfDay"].map(hhmm))[
                ["Time", "From Stop", "To Stop", "Total Psgr", "Coll Amt"]].rename(
                columns={"Total Psgr": "Passengers", "Coll Amt": "Collected ₹"}),
                hide_index=True, width="stretch")


st.caption(f"BusOps Command Center • SOLVATHON 2026 PS-3 • {S['Tickets']:,} tickets • {S['Passengers']:,} passengers shown"),