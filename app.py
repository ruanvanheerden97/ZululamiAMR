import glob, os, io, json, datetime, html, zipfile, re
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="Estate Smart Metering Progress", page_icon="⚡", layout="wide", initial_sidebar_state="auto")

# ---------- Styling (same palette as the Lake Michelle / Sitari apps) ----------
st.markdown("""
<style>
.block-container {padding-top: 1.2rem; max-width: 1500px;}
div[data-testid="stMetric"] {background:#FBF9F3 !important;border:1px solid #DCD6C4;border-radius:10px;padding:10px 14px;}
div[data-testid="stMetric"] * {color:#152B45 !important;}
div[data-testid="stMetricLabel"], div[data-testid="stMetricLabel"] * {font-size:12px !important;text-transform:uppercase;letter-spacing:.05em;color:#3E5066 !important;}
div[data-testid="stMetricDelta"], div[data-testid="stMetricDelta"] * {color:#3F7D5C !important;fill:#3F7D5C !important;}
h1, h2, h3 {color:#152B45;}
@media (max-width: 640px) {
  .block-container {padding: 3.2rem 0.7rem 2rem !important;}
  div[data-testid="stHorizontalBlock"] {flex-wrap: wrap !important; gap: 0.5rem !important;}
  div[data-testid="stColumn"], div[data-testid="column"] {min-width: calc(50% - 0.5rem) !important; flex: 1 1 calc(50% - 0.5rem) !important; width: calc(50% - 0.5rem) !important;}
  h1 {font-size: 1.45rem !important; line-height: 1.25 !important;}
  h2 {font-size: 1.2rem !important;} h3 {font-size: 1.05rem !important;}
  div[data-testid="stMetric"] {padding: 6px 10px;}
  div[data-testid="stMetricValue"], div[data-testid="stMetricValue"] * {font-size: 1.35rem !important;}
  div[data-testid="stMetricLabel"], div[data-testid="stMetricLabel"] * {font-size: 10.5px !important;}
}
</style>""", unsafe_allow_html=True)

HERE = os.path.dirname(os.path.abspath(__file__))
# ---------- Estates ----------
ESTATES = {
    "Zululami": dict(code="ZLM", wb="ZLM_Meter_Hierarchy_*.xlsx", elec_export="ZLM_AMR_Export_*.xlsx",
                     water_list="ZLM_ALL_Water_meters_*.csv", water_export="ZLM_Water_AMR_Export_*.xlsx",
                     stands="Zululami Stands", roads_kmz="ZLM Minisubs*.kmz", blurb="Zululami and Coral Cove"),
    "Seaton": dict(code="SEA", wb="SEA_Meter_Hierarchy_*.xlsx", elec_export=None, device_list="SEA_Device_List_*.csv",
                   stands="Seaton polygons", roads_kmz=None, blurb="Seaton"),
}
ESTATE = st.sidebar.selectbox("🏘️ Estate", list(ESTATES), key="estate")
EST = ESTATES[ESTATE]
SITE_NAME = f"{ESTATE} Estate"
FILE_PATTERN = EST["wb"]
AMR_EXPORT_PATTERN = EST["elec_export"] or "__no_export__"   # optional: latest platform export, matched on meter serial
STALE_DAYS = 7
COL = {"Installed": "#2E7D32", "Ordered": "#E3A008", "Outstanding": "#C0392B", "Deferred": "#8A8F98"}
STATUS_ORDER = ["Installed", "Ordered", "Outstanding", "Deferred"]


# ---------- Helpers ----------
def latest(pattern):
    m = glob.glob(os.path.join(HERE, pattern))
    return max(m, key=os.path.getmtime) if m else None


def fmt_sn(v):
    if v is None or (isinstance(v, float) and pd.isna(v)) or v == "":
        return ""
    try:
        return str(int(float(v)))
    except (ValueError, OverflowError):
        return str(v).strip()


def to_bool(v):
    return v if isinstance(v, bool) else str(v).strip().upper() in ("TRUE", "YES", "1")


def parse_comms(v):
    try:
        return pd.to_datetime(str(v), format="%m/%d/%Y, %I:%M %p")
    except Exception:
        return pd.to_datetime(v, errors="coerce")


@st.cache_data(show_spinner=False)
def load_workbook(path, _mtime):
    x = pd.ExcelFile(path)
    elec = x.parse("Elec").fillna("")
    elec["Meter Serial"] = elec["Meter Serial"].apply(fmt_sn)
    elec["AMR Installed"] = elec["AMR Installed"].apply(to_bool)
    for col in ["Stand / Supply", "Breaker", "Phase", "Comments", "AMR Order", "AMR Label", "AMR Last Comms", "AMR Device", "Occupancy"]:
        elec[col] = elec[col].astype(str).replace("nan", "")
    elec["AMR Port"] = elec["AMR Port"].apply(lambda v: "" if v == "" else fmt_sn(v))
    kiosks = x.parse("Kiosks").fillna("")
    mss = x.parse("Minisubs").fillna("")
    mss["Bulk Meter Serial"] = mss["Bulk Meter Serial"].apply(fmt_sn)
    orders = x.parse("Orders").fillna("") if "Orders" in x.sheet_names else pd.DataFrame()
    issues = x.parse("Data Issues").fillna("") if "Data Issues" in x.sheet_names else pd.DataFrame()
    return elec, kiosks, mss, orders, issues


@st.cache_data(show_spinner=False)
def load_amr_export(path, _mtime):
    """Platform export: SNR | ADDRESS | METER NUMBER | LASTCOMMS | READING."""
    df = pd.read_excel(path)
    df.columns = [str(c).strip().upper() for c in df.columns]
    df = df[df["SNR"].notna()].copy()
    df["serial"] = df["METER NUMBER"].apply(fmt_sn)
    df["device"] = df["SNR"].astype(str).str.split("_").str[0]
    df["port"] = df["SNR"].astype(str).str.split("_").str[1].fillna("")
    df["comms"] = df["LASTCOMMS"].apply(parse_comms)
    return df


def apply_export(elec, mss, exp):
    """Serial-keyed overlay of the latest AMR export onto the workbook."""
    e = elec.copy(); notes = []
    by_sn = {r["serial"]: r for _, r in exp.iterrows()}
    for i, r in e.iterrows():
        a = by_sn.get(r["Meter Serial"])
        if a is not None and r["Meter Serial"]:
            if not r["AMR Installed"]:
                notes.append({"Type": "Newly on AMR", "Minisub": r["Minisub"], "Kiosk": r["Kiosk"], "Stand": r["Stand / Supply"], "Serial": r["Meter Serial"], "Detail": a["ADDRESS"]})
            e.at[i, "AMR Installed"] = True; e.at[i, "AMR Device"] = a["device"]; e.at[i, "AMR Port"] = a["port"]
            e.at[i, "AMR Label"] = a["ADDRESS"]; e.at[i, "AMR Last Comms"] = a["LASTCOMMS"]
    # minisub bulk meters: mark as on AMR when their serial is in the export
    for i, m in mss.iterrows():
        a = by_sn.get(m["Bulk Meter Serial"])
        if a is not None and m["Bulk Meter Serial"]:
            if m["Bulk on AMR"] != "Yes":
                notes.append({"Type": "Bulk meter newly on AMR", "Minisub": m["Minisub"], "Kiosk": "", "Stand": "Bulk", "Serial": m["Bulk Meter Serial"], "Detail": a["ADDRESS"]})
            mss.at[i, "Bulk on AMR"] = "Yes"; mss.at[i, "Bulk AMR Device"] = a["device"]; mss.at[i, "Bulk Last Comms"] = a["LASTCOMMS"]
    known = set(e["Meter Serial"]) | set(mss["Bulk Meter Serial"])
    for _, a in exp.iterrows():
        if a["serial"] not in known:
            notes.append({"Type": "In export, not in workbook", "Minisub": "", "Kiosk": "", "Stand": "", "Serial": a["serial"], "Detail": a["ADDRESS"]})
    return e, pd.DataFrame(notes)


def status_of(r):
    if r["AMR Installed"]:
        return "Installed"
    if str(r.get("Comments", "")).startswith("Deferred"):
        return "Deferred"
    if str(r.get("AMR Order", "")).strip():
        return "Ordered"
    return "Outstanding"


def bar_html(counts, total, h=12):
    if not total:
        return ""
    seg = "".join(f'<span title="{k} {counts.get(k,0)}" style="display:inline-block;height:{h}px;width:{100*counts.get(k,0)/total:.3f}%;background:{COL[k]}"></span>'
                  for k in STATUS_ORDER if counts.get(k, 0))
    return f'<div style="width:100%;background:#eee;border-radius:3px;overflow:hidden;line-height:0">{seg}</div>'


# ---------- Load ----------
data_path = latest(FILE_PATTERN)
if not data_path:
    st.error(f"No spreadsheet matching `{FILE_PATTERN}` next to app.py. Push the workbook to the repo.")
    st.stop()
elec, kiosks, mss, orders, issues = load_workbook(data_path, os.path.getmtime(data_path))
exp_path = latest(AMR_EXPORT_PATTERN)
exp_notes = pd.DataFrame()
exp = None
if exp_path:
    exp = load_amr_export(exp_path, os.path.getmtime(exp_path))
    mss = mss.copy()
    elec, exp_notes = apply_export(elec, mss, exp)
# Seaton: AMR flags come from the site device list (IsAMR / IsAMR1), matched on serial
dev_list = None
if EST.get("device_list"):
    _dl = latest(EST["device_list"])
    if _dl:
        dev_list = pd.read_csv(_dl, dtype=str).fillna("")
        dev_list["esn"] = dev_list["SerialNumber"].str.replace(r"\s|^M\.E\.S\.", "", regex=True)
        _on = set(dev_list.loc[dev_list["IsAMR"].str.upper() == "TRUE", "esn"])
        elec["AMR Installed"] = elec["AMR Installed"] | elec["Meter Serial"].isin(_on)
        mss = mss.copy(); mss.loc[mss["Bulk Meter Serial"].isin(_on), "Bulk on AMR"] = "Yes"
elec["Status"] = elec.apply(status_of, axis=1)
elec["comms_dt"] = elec["AMR Last Comms"].apply(lambda v: parse_comms(v) if v else pd.NaT)
ref_now = (exp["comms"].max() if exp is not None else elec["comms_dt"].max())
if pd.isna(ref_now):
    ref_now = pd.Timestamp.now()
elec["Stale"] = elec["AMR Installed"] & elec["comms_dt"].notna() & ((ref_now - elec["comms_dt"]).dt.days > STALE_DAYS)
MS_LIST = sorted(mss["Minisub"].tolist(), key=lambda m: int(m[3:]))

# =====================================================================
# SLD TAB
# =====================================================================
def build_tree(ms):
    k = kiosks[(kiosks["Minisub"] == ms) & (kiosks["Kiosk"] != "MS LV board")]
    nodes = {r["Kiosk"]: dict(row=r, children=[]) for _, r in k.iterrows()}
    roots, unconf = [], []
    for name, n in nodes.items():
        p = n["row"]["Fed From"]
        if p in nodes: nodes[p]["children"].append(n)
        elif p == "MS LV board": roots.append(n)
        else: unconf.append(n)
    return roots, unconf


def kiosk_card(n, ms, search):
    r = n["row"]; name = r["Kiosk"]
    sup = elec[(elec["Minisub"] == ms) & (elec["Kiosk"] == name)]
    cc = sup["Status"].value_counts().to_dict(); total = len(sup)
    if not total: colr = "#9AA5B1"
    elif cc.get("Installed", 0) == total: colr = COL["Installed"]
    elif cc.get("Ordered", 0): colr = COL["Ordered"]
    elif cc.get("Installed", 0): colr = COL["Installed"]
    else: colr = COL["Outstanding"]
    hit = search and (search.lower() in name.lower() or sup["Stand / Supply"].astype(str).str.lower().str.contains(search.lower(), regex=False).any()
                      or sup["Meter Serial"].str.contains(search, regex=False).any())
    rows = "".join(
        f'<tr><td><span class="dot" style="background:{COL[s["Status"]]}"></span>{html.escape(str(s["Stand / Supply"]))}</td>'
        f'<td class="mono">{s["Meter Serial"] or "–"}</td><td>{html.escape(str(s["Phase"]))}</td><td>{ {"Occupied":"O","Under construction":"C","Vacant":"V"}.get(s["Occupancy"],"?") }</td>'
        f'<td class="mono">{("P" + str(s["AMR Port"])) if s["AMR Port"] else ""}{" ⚠" if s["Stale"] else ""}</td></tr>' for _, s in sup.iterrows())
    link = r["Feed Link Basis"]
    style = "dashed" if ("Inferred" in link or "Assumed" in link) else ("dotted" if "Unconfirmed" in link else "solid")
    feed = html.escape(str(r["Feed Breaker / Cable"]).split("·")[0])
    meta = " · ".join(x for x in [f"Main {r['Kiosk Main Breaker']}" if r["Kiosk Main Breaker"] else "", str(r["Planned AMR Unit"]) if total else "LV distribution",
                                  f"dev …{str(r['AMR Device(s)'])[-4:]}" if r["AMR Device(s)"] else ""] if x)
    return f'''<div class="node"><div class="wire" style="border-top-style:{style}"><span>{feed}</span></div>
      <details class="card{' hit' if hit else ''}" style="border-color:{colr}" {'open' if hit or total <= 6 else ''}>
      <summary style="background:{colr}22"><b>{html.escape(name)}</b><span class="cnt" style="color:{colr}">{cc.get("Installed",0)}/{total if total else "LV"}</span>
      <div class="meta">{html.escape(meta)}</div></summary>
      {f'<table>{rows}</table>' if total else ''}</details></div>'''


def render_tree(n, ms, search):
    kids = "".join(render_tree(c, ms, search) for c in n["children"])
    return f'<div class="branch">{kiosk_card(n, ms, search)}{f"<div class=kids>{kids}</div>" if kids else ""}</div>'


SLD_CSS = """<style>
body{font-family:-apple-system,Segoe UI,Arial,sans-serif;color:#152B45;margin:0}
.wrap{display:flex;gap:0;align-items:flex-start;overflow-x:auto;padding:6px}
.ms{min-width:170px;border:2px solid #152B45;background:#EEF2F6;border-radius:6px;padding:10px;margin-right:0}
.ms h2{margin:0 0 4px;font-size:20px}.ms div{font-size:12px}
.bus{border-left:5px solid #152B45;padding-left:0;margin-left:14px}
.feeder{display:flex;align-items:flex-start;margin:6px 0}
.branch{display:flex;align-items:flex-start}
.kids{display:flex;flex-direction:column}
.node{display:flex;align-items:flex-start}
.wire{width:44px;border-top:2px solid #152B45;margin-top:16px;position:relative}
.wire span{position:absolute;top:-15px;left:2px;font-size:10px;white-space:nowrap}
.card{width:236px;border:2px solid;border-radius:6px;background:#fff;margin:2px 0 8px;font-size:11px}
.card.hit{box-shadow:0 0 0 3px #2F80ED}
summary{list-style:none;padding:5px 7px;cursor:pointer;border-radius:4px 4px 0 0}
summary::-webkit-details-marker{display:none}
.cnt{float:right;font-weight:700}.meta{font-size:10px;color:#3E5066;margin-top:2px}
table{width:100%;border-collapse:collapse;margin:2px 0 4px}td{padding:1px 5px;border-top:1px solid #f0f0f0;white-space:nowrap}
.mono{font-family:Consolas,monospace;font-size:10.5px}
.dot{display:inline-block;width:8px;height:8px;border-radius:2px;margin-right:5px}
.direct{font-size:11px;margin-top:8px}.unc{margin-top:16px;color:#C0392B;font-weight:700;font-size:12px}
.legend{font-size:11px;margin:4px 0 8px}.legend span{margin-right:14px}
</style>"""

def page_sld():
    st.header("⚡ Electricity — minisub reticulation")
    st.caption("Pick a minisub to see its kiosks, every supply and its smart-meter (AMR) status. Tap a kiosk to open its meter list.")
    l, r_ = st.columns([1, 3])
    ms = l.selectbox("Minisub", MS_LIST, index=MS_LIST.index("MS 07"))
    search = r_.text_input("🔍 Highlight stand, kiosk or meter serial", "")
    msr = mss[mss["Minisub"] == ms].iloc[0]
    sel = elec[elec["Minisub"] == ms]; sc = sel["Status"].value_counts()
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Supplies", len(sel)); k2.metric("On AMR", int(sc.get("Installed", 0)), f"{round(100*sc.get('Installed',0)/len(sel)) if len(sel) else 0}%")
    k3.metric("Ordered", int(sc.get("Ordered", 0))); k4.metric("Outstanding", int(sc.get("Outstanding", 0)) + int(sc.get("Deferred", 0)))
    k5.metric("Bulk meter", msr["Bulk Meter Serial"], "on AMR" if msr["Bulk on AMR"] == "Yes" else "not on AMR")
    roots, unconf = build_tree(ms)
    direct = sel[sel["Kiosk"] == "MS LV board"]
    dlist = "".join(f'<div><span class="dot" style="background:{COL[d["Status"]]}"></span>{html.escape(str(d["Stand / Supply"]))} · {d["Breaker"]} · {d["Meter Serial"] or "–"}</div>' for _, d in direct.iterrows())
    body = f'''<div class="legend">{"".join(f'<span><span class="dot" style="background:{COL[k]}"></span>{k}</span>' for k in STATUS_ORDER)}
      <span>── feed per sheet</span><span>- - inferred from sheet layout</span><span>··· not confirmed</span><span>row: stand · serial · phase · occupancy · AMR port (⚠ stale comms)</span></div>
      <div class="wrap"><div><div class="ms"><h2>{ms}</h2><div>{html.escape(str(msr["Manufacturer"]))}</div><div>{html.escape(str(msr["Rating"]))}</div>
      <div>Bulk {msr["Bulk Meter Serial"]} · {"on AMR" if msr["Bulk on AMR"]=="Yes" else "<b style='color:#C0392B'>not on AMR</b>"}</div></div>
      {f'<div class="direct"><b>Direct from LV board</b>{dlist}</div>' if len(direct) else ''}</div>
      <div class="bus">{"".join(f'<div class="feeder">{render_tree(n, ms, search)}</div>' for n in roots)}</div></div>
      {f'<div class="unc">Feed not confirmed in sheet</div><div class="wrap">{"".join(render_tree(n, ms, search) for n in unconf)}</div>' if unconf else ''}'''
    height = 160 + sum(110 + 22 * min(len(elec[(elec["Minisub"] == ms) & (elec["Kiosk"] == k)]), 16) for k in kiosks[kiosks["Minisub"] == ms]["Kiosk"]) // 2
    h = min(max(height, 520), 2200)
    if hasattr(st, "iframe"):
        st.iframe(SLD_CSS + body, height=h)
    else:
        components.html(SLD_CSS + body, height=h, scrolling=True)

    # PDF drawing for this minisub (or full set)
    try:
        import make_sld
        cpdf1, cpdf2 = st.columns(2)
        if cpdf1.button(f"📄 Build A3 SLD PDF for {ms}"):
            st.session_state["pdf_one"] = make_sld.build_pdf_bytes(data_path, only=ms)
        if "pdf_one" in st.session_state:
            cpdf1.download_button("Download", st.session_state["pdf_one"], file_name=f"ZLM_SLD_{ms.replace(' ', '')}.pdf", mime="application/pdf")
        if cpdf2.button("📚 Build full drawing set (overview + 14 sheets)"):
            st.session_state["pdf_all"] = make_sld.build_pdf_bytes(data_path)
        if "pdf_all" in st.session_state:
            cpdf2.download_button("Download set", st.session_state["pdf_all"], file_name="ZLM_Reticulation_SLD_AMR.pdf", mime="application/pdf")
    except Exception as ex:  # drawings are optional in the cloud app
        st.caption(f"PDF drawings unavailable: {ex}")

# =====================================================================
# AMR PROGRESS TAB
# =====================================================================
def page_amr():
    st.header("📡 Electricity AMR progress")
    st.subheader("Progress per minisub")
    rows = []
    for m in MS_LIST:
        e = elec[elec["Minisub"] == m]; c2 = e["Status"].value_counts().to_dict()
        rows.append(dict(Minisub=m, Supplies=len(e), Installed=c2.get("Installed", 0), Ordered=c2.get("Ordered", 0), Outstanding=c2.get("Outstanding", 0),
                         Deferred=c2.get("Deferred", 0), Occupied=int((e["Occupancy"] == "Occupied").sum()),
                         bar=bar_html(c2, len(e)), Bulk="✅" if mss.loc[mss["Minisub"] == m, "Bulk on AMR"].iloc[0] == "Yes" else "❌"))
    t = pd.DataFrame(rows)
    hdr = "".join(f"<th style='text-align:left;padding:4px 8px'>{h}</th>" for h in ["Minisub", "Bulk on AMR", "Supplies", "Occupied", "On AMR", "Ordered", "Outstanding", "Deferred", "% on AMR", "Progress"])
    body = "".join(f"<tr><td style='padding:4px 8px'><b>{r['Minisub']}</b></td><td>{r['Bulk']}</td><td>{r['Supplies']}</td><td>{r['Occupied']}</td><td>{r['Installed']}</td><td>{r['Ordered']}</td>"
                   f"<td>{r['Outstanding']}</td><td>{r['Deferred']}</td><td>{round(100*r['Installed']/r['Supplies']) if r['Supplies'] else 0}%</td><td style='width:35%'>{r['bar']}</td></tr>" for r in rows)
    st.markdown(f"<table style='width:100%;border-collapse:collapse;font-size:14px'><tr>{hdr}</tr>{body}</table>", unsafe_allow_html=True)

    st.subheader("AMR devices on site")
    inst = elec[elec["AMR Installed"]]
    dev = inst.groupby(["AMR Device"]).agg(Minisub=("Minisub", "first"), Location=("Kiosk", lambda s: ", ".join(sorted(set(s)))), Channels=("Stand / Supply", "count"),
                                           Highest_port=("AMR Port", lambda s: max([int(x) for x in s if str(x).isdigit()] or [0])),
                                           Last_comms=("comms_dt", "max")).reset_index()
    st.dataframe(dev, width="stretch", hide_index=True)
    stale = elec[elec["Stale"]]
    st.subheader(f"Stale comms (> {STALE_DAYS} days before {ref_now:%Y-%m-%d})")
    if stale.empty: st.success("All installed devices reported recently.")
    else: st.dataframe(stale[["Minisub", "Kiosk", "Stand / Supply", "Meter Serial", "AMR Device", "AMR Port", "AMR Last Comms"]], width="stretch", hide_index=True)
    bulk_off = mss[mss["Bulk on AMR"] != "Yes"]
    if not bulk_off.empty:
        st.warning("Minisub bulk meters not on AMR: " + ", ".join(f"{r['Minisub']} ({r['Bulk Meter Serial']})" for _, r in bulk_off.iterrows()))

# =====================================================================
# ORDERS TAB
# =====================================================================
def page_orders():
    st.header("📅 Planned installations")
    te, tw = st.tabs(["⚡ Electricity", "💧 Water"])
    with te:
        page_orders_elec()
    with tw:
        page_orders_water()


def page_orders_elec():
    st.subheader("Orders")
    if not orders.empty:
        st.dataframe(orders, width="stretch", hide_index=True)
    st.subheader("Work allocated to open orders")
    op = elec[elec["Status"] == "Ordered"]
    alloc = op.groupby(["AMR Order", "Minisub", "Kiosk"]).size().reset_index(name="Probes")
    alloc = alloc.merge(kiosks[["Minisub", "Kiosk", "Planned AMR Unit", "AMR Device(s)"]], on=["Minisub", "Kiosk"], how="left")
    alloc["Unit"] = alloc.apply(lambda r: "add to existing device" if r["AMR Device(s)"] else r["Planned AMR Unit"], axis=1)
    st.dataframe(alloc.drop(columns=["AMR Device(s)", "Planned AMR Unit"]), width="stretch", hide_index=True)
    a1, a2, a3 = st.columns(3)
    a1.metric("16-probe units needed", int((alloc["Unit"] == "16-probe").sum())); a2.metric("8-probe units needed", int((alloc["Unit"] == "8-probe").sum()))
    a3.metric("Probes needed", int(alloc["Probes"].sum()))

    st.subheader("Next minisub to tackle")
    st.caption("Kiosks with nothing on AMR and not on an open order, sized 2–8 supplies = 8-probe, 9–16 = 16-probe, 1 = standalone. "
               "Ranked by occupancy, then by occupied supplies brought online per unit.")
    rk = []
    for m in MS_LIST:
        e = elec[(elec["Minisub"] == m) & (elec["Status"] == "Outstanding")]
        if e.empty: continue
        units = {"16-probe": 0, "8-probe": 0, "Standalone": 0}
        for k, g in e.groupby("Kiosk"):
            n = len(elec[(elec["Minisub"] == m) & (elec["Kiosk"] == k)])
            u = "Standalone" if n == 1 else ("8-probe" if n <= 8 else "16-probe")
            units[u] += 1
        allm = elec[elec["Minisub"] == m]
        occ = (e["Occupancy"] == "Occupied").sum()
        rk.append(dict(Minisub=m, **{"Occupancy %": round(100 * (allm["Occupancy"] == "Occupied").mean())}, **units, Units=sum(units.values()), Probes=len(e),
                       **{"Occupied supplies": occ, "Occupied per unit": round(occ / max(1, sum(units.values())), 1)}))
    if rk:
        st.dataframe(pd.DataFrame(rk).sort_values(["Occupancy %", "Occupied per unit"], ascending=False), width="stretch", hide_index=True)

# =====================================================================
# ALL SUPPLIES TAB
# =====================================================================
def page_all_elec():
    f1, f2, f3 = st.columns(3)
    fm = f1.multiselect("Minisub", MS_LIST)
    fs = f2.multiselect("Status", STATUS_ORDER)
    fq = f3.text_input("Search stand / serial / kiosk")
    v = elec.copy()
    if fm: v = v[v["Minisub"].isin(fm)]
    if fs: v = v[v["Status"].isin(fs)]
    if fq:
        q = fq.lower()
        v = v[v.apply(lambda r: q in f"{r['Stand / Supply']} {r['Meter Serial']} {r['Kiosk']}".lower(), axis=1)]
    cols = ["Supply ID", "Minisub", "Kiosk", "Stand / Supply", "Supply Type", "Breaker", "Phase", "Meter Serial", "Occupancy", "Status", "AMR Device", "AMR Port",
            "AMR Last Comms", "AMR Order", "Comments", "Source"]
    st.dataframe(v[cols], width="stretch", hide_index=True, height=600)
    st.download_button("⬇️ CSV of this view", v[cols].to_csv(index=False).encode(), "zlm_supplies.csv", "text/csv")

# =====================================================================
# MAP TAB
# =====================================================================
# =====================================================================
# CHECKS TAB
# =====================================================================
def page_checks_elec():
    if not exp_notes.empty:
        st.subheader("AMR export vs workbook"); st.dataframe(exp_notes, width="stretch", hide_index=True)
    st.subheader("Kiosk feeds to confirm")
    st.dataframe(kiosks[kiosks["Feed Link Basis"].astype(str).str.contains("Unconfirmed|Inferred|Assumed")][["Minisub", "Kiosk", "Fed From", "Feed Link Basis", "Source"]],
                 width="stretch", hide_index=True)
    st.subheader("Kiosks without a KMZ point")
    st.dataframe(kiosks[(kiosks["Latitude"] == "") & (kiosks["Supplies"] > 0)][["Minisub", "Kiosk", "Supplies"]], width="stretch", hide_index=True)
    dup = elec[(elec["Meter Serial"] != "") & elec.duplicated("Meter Serial", keep=False)].sort_values("Meter Serial")
    st.subheader("Duplicate meter serials")
    st.dataframe(dup[["Meter Serial", "Minisub", "Kiosk", "Stand / Supply", "Source"]], width="stretch", hide_index=True)
    st.subheader("Known data issues (from workbook)")
    if not issues.empty: st.dataframe(issues, width="stretch", hide_index=True)


@st.cache_data(show_spinner=False)
def kmz_lines(path, _mtime):
    K = "{http://www.opengis.net/kml/2.2}"
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read([n for n in z.namelist() if n.endswith(".kml")][0]))
    out = []
    def walk(e, p):
        for ch in e:
            if ch.tag in (K + "Folder", K + "Document"):
                nm = ch.find(K + "name"); walk(ch, p + [nm.text if nm is not None else ""])
            elif ch.tag == K + "Placemark":
                ls = ch.find(".//" + K + "LineString")
                if ls is not None and any("ROADS" in x or "CADASTRAL" in x for x in p):
                    out.append({"path": [list(map(float, c.split(",")[:2])) for c in ls.find(K + "coordinates").text.split()]})
    walk(root, [])
    return out


# =====================================================================
# WATER + STAND POLYGONS
# =====================================================================
from zlm_keys import stand_key, development

WATER_LIST_PATTERN = "ZLM_ALL_Water_meters_*.csv"
WATER_EXPORT_PATTERN = "ZLM_Water_AMR_Export_*.xlsx"
SCOL = {"On AMR": "#2E7D32", "Partly on AMR": "#8BC34A", "Planned (ordered)": "#E3A008", "Not on AMR yet": "#C0392B", "No meter yet": "#9AA5B1"}


def norm_water_sn(v):
    s = str(v).strip().upper().replace(" ", "").replace("-", "")
    s = re.sub(r"^8SEN", "SEN", s)
    return re.sub(r"^SN", "", s)


@st.cache_data(show_spinner=False)
def load_water(list_path, list_mtime, exp_path, exp_mtime):
    site = pd.read_csv(list_path, dtype=str).fillna("")
    site.columns = [c.strip() for c in site.columns]
    site = site[site["Stand"].str.strip() != ""].copy()
    site["Stand"] = site["Stand"].str.strip()
    site["Serial"] = site["SerialNumber1"].str.strip()
    site["n"] = site["Serial"].map(norm_water_sn)
    site["Key"] = site["Stand"].map(stand_key)
    site["Development"] = site["Key"].map(development)
    ex = pd.DataFrame(columns=["SNR", "ADDRESS", "METER NUMBER", "LASTCOMMS", "READING", "n", "Key"])
    if exp_path:
        ex = pd.read_excel(exp_path, dtype=str).fillna("")
        ex.columns = [str(c).strip().upper() for c in ex.columns]
        ex = ex[ex["SNR"] != ""].copy()
        ex["n"] = ex["METER NUMBER"].map(norm_water_sn)
        ex["Key"] = ex["ADDRESS"].map(stand_key)
    by = {r["n"]: r for _, r in ex.iterrows()}
    # On AMR = active on the site list (IsAMR1, covers Sensus meters not in the LoRaWAN export)
    #          OR present in the LoRaWAN AMR export (matched on serial)
    amr_col = next((c for c in site.columns if c.lower().startswith("isamr")), None)
    flag = site[amr_col].astype(str).str.strip().str.upper().isin(["TRUE", "YES", "1"]) if amr_col else False
    in_exp = site["n"].isin(by)
    site["AMR Installed"] = flag | in_exp
    site["AMR Type"] = [("LoRaWAN" if e else ("Sensus / other (site list)" if f else "")) for e, f in zip(in_exp, flag if amr_col else [False] * len(site))]
    site["AMR Device"] = site["n"].map(lambda n: by[n]["SNR"] if n in by else "")
    site["AMR Label"] = site["n"].map(lambda n: by[n]["ADDRESS"] if n in by else "")
    site["AMR Last Comms"] = site["n"].map(lambda n: by[n]["LASTCOMMS"] if n in by else "")
    site["Reading"] = site["n"].map(lambda n: by[n]["READING"] if n in by else "")
    # discrepancy: on AMR (export) but serial not on the site list
    known = set(site["n"])
    disc = []
    for _, r in ex[~ex["n"].isin(known)].iterrows():
        same = site[site["Key"] == r["Key"]]
        disc.append({"AMR label": r["ADDRESS"], "Serial on AMR": r["METER NUMBER"], "Stand (from label)": r["Key"],
                     "Serial on site list for this stand": ", ".join(same["Serial"]) or "— stand not on site list",
                     "Site list stand": ", ".join(same["Stand"]), "Last comms": r["LASTCOMMS"],
                     "Likely issue": ("Serial differs from site list (meter swapped or typo)" if len(same) else "Stand missing from site list")})
    return site, ex, pd.DataFrame(disc)


@st.cache_data(show_spinner=False)
def load_stand_polygons(path, _mtime, folder="Zululami Stands"):
    """Stand polygons from one estate's folder of the KMZ ('Zululami Stands' or 'Seaton polygons')."""
    K = "{http://www.opengis.net/kml/2.2}"
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read([n for n in z.namelist() if n.endswith(".kml")][0]))
    cont = (K + "Folder", K + "Document")
    top = next((e for e in root.iter() if e.tag in cont and (e.findtext(K + "name") or "").strip() == folder), None)
    out, notes = [], []
    if top is None:
        return out, notes
    seen = {}
    groups = [g for g in top if g.tag in cont] or [top]
    for grp in groups:
        gname = (grp.findtext(K + "name") or "").strip()
        for pm in grp.iter(K + "Placemark"):
            name = (pm.findtext(K + "name") or "").strip()
            rings = []
            for pg in pm.iter(K + "Polygon"):
                c = pg.find(".//" + K + "outerBoundaryIs//" + K + "coordinates")
                if c is not None:
                    rings.append([[round(float(p.split(",")[1]), 6), round(float(p.split(",")[0]), 6)] for p in c.text.split()])
            if not rings: continue
            if gname.startswith("Marula"): key = "BLOCK:1350"
            elif gname.startswith("Highline"): key = "BLOCK:1356"
            elif gname.startswith("Husk"): key = "BLOCK:0511_" + name[-1].upper()
            elif folder != "Zululami Stands": key = seaton_key(name)
            else: key = stand_key(name, coral=gname.startswith("Coral"))
            if name == "16628": notes.append("Stand polygon named 16628 – treated as 1628.")
            if key in seen: notes.append(f"Stand polygon {name} appears twice in '{gname}' (a neighbouring stand may be mislabelled).")
            seen[key] = 1
            out.append(dict(key=key, name=name, group=gname, rings=rings))
    return out, notes


def block_members(key, keys):
    if key == "BLOCK:1350": return [k for k in keys if k.startswith("1350_")]
    if key == "BLOCK:1356": return [k for k in keys if k.startswith("1356_")]
    if key.startswith("BLOCK:0511_"):
        rng = range(1, 11) if key.endswith("A") else range(11, 25)
        return [f"0511_{i:02d}" for i in rng]
    return [key]


def elec_stand_status(rows):
    st_ = list(rows["Status"])
    if not st_: return "No meter yet"
    if all(s == "Installed" for s in st_): return "On AMR"
    if any(s == "Installed" for s in st_): return "Partly on AMR"
    if any(s == "Ordered" for s in st_): return "Planned (ordered)"
    if (rows["Meter Serial"] != "").any(): return "Not on AMR yet"
    return "No meter yet"


def water_stand_status(rows):
    if rows.empty: return "No meter yet"
    if rows["AMR Installed"].all(): return "On AMR"
    if rows["AMR Installed"].any(): return "Partly on AMR"
    return "Not on AMR yet"


def seaton_key(s):
    s = str(s or "").strip().upper()
    if not s or s in ("???", "-"): return ""
    m = re.fullmatch(r"(\d+)\s*(/\s*DEV)?", s)
    return str(int(m.group(1))) if m else s


def seaton_dev(stand):
    s = str(stand).upper()
    if re.match(r"^\d", s): return "Freestanding"
    if s.startswith(("CONDO", "VILLAS", "EXT H")): return "Condos & villas"
    return "Common / bulk"


def load_water_seaton(dl):
    w = dl[dl["SerialNumber1"].str.strip() != ""].copy()
    w["Stand"] = w["Stand"].str.strip(); w["Serial"] = w["SerialNumber1"].str.replace(" ", "")
    w["AMR Installed"] = w["IsAMR1"].str.upper() == "TRUE"
    w["AMR Type"] = w["AMR Installed"].map(lambda b: "Sensus (Atom)" if b else "")
    w["AMR Label"] = ""; w["AMR Last Comms"] = ""; w["Reading"] = ""; w["AMR Device"] = ""
    w["Key"] = w["Stand"].map(seaton_key); w["Development"] = w["Stand"].map(seaton_dev)
    w["Occupancy"] = w["Stand-Type"].str.strip()
    return w


wdisc = pd.DataFrame()
if ESTATE == "Seaton":
    water = load_water_seaton(dev_list) if dev_list is not None else pd.DataFrame()
    elec["Key"] = [seaton_key(s) for s in elec["Stand / Supply"]]
else:
    wl_path, wx_path = latest(EST["water_list"]), latest(EST["water_export"])
    water, wexp, wdisc = (load_water(wl_path, os.path.getmtime(wl_path), wx_path, os.path.getmtime(wx_path) if wx_path else 0)
                          if wl_path else (pd.DataFrame(), pd.DataFrame(), pd.DataFrame()))
    elec["Key"] = [stand_key(s, coral=m in ("MS 04", "MS 05")) if s not in ("???", "-") else "" for s, m in zip(elec["Stand / Supply"], elec["Minisub"])]
stand_kmz = None
for p in sorted(glob.glob(os.path.join(HERE, "*.kmz")), key=os.path.getmtime, reverse=True):
    try:
        with zipfile.ZipFile(p) as z:
            if EST["stands"].encode() in z.read([n for n in z.namelist() if n.endswith(".kml")][0]):
                stand_kmz = p; break
    except Exception:
        pass
polys, poly_notes = load_stand_polygons(stand_kmz, os.path.getmtime(stand_kmz), EST["stands"]) if stand_kmz else ([], [])


def pct(a, b):
    return f"{round(100 * a / b)}%" if b else "–"


# =====================================================================
# PAGES
# =====================================================================
def page_overview():
    st.title(f"{ESTATE} — Smart Metering Progress")
    src = [os.path.basename(exp_path or data_path)]
    if ESTATE == "Zululami" and latest(EST["water_export"]): src.append(os.path.basename(latest(EST["water_export"])))
    if EST.get("device_list") and latest(EST["device_list"]): src.append(os.path.basename(latest(EST["device_list"])))
    st.caption(f"Progress of automatic meter reading (AMR) for electricity and water across {EST['blurb']}. "
               f"Updated from: {', '.join(src)}.")
    cnt = elec["Status"].value_counts(); tot = len(elec)
    st.subheader("⚡ Electricity")
    c = st.columns(4)
    c[0].metric("Meter points", tot)
    c[1].metric("On AMR", int(cnt.get("Installed", 0)), pct(cnt.get("Installed", 0), tot))
    c[2].metric("Planned (ordered)", int(cnt.get("Ordered", 0)))
    c[3].metric("Still to do", int(cnt.get("Outstanding", 0) + cnt.get("Deferred", 0)))
    st.markdown(bar_html(cnt.to_dict(), tot, 14), unsafe_allow_html=True)
    if not water.empty:
        st.subheader("💧 Water")
        wt = len(water); wi = int(water["AMR Installed"].sum())
        c = st.columns(4)
        c[0].metric("Water meters", wt); c[1].metric("On AMR", wi, pct(wi, wt)); c[2].metric("Not on AMR yet", wt - wi)
        c[3].metric("Developments", water["Development"].nunique())
        st.markdown(bar_html({"Installed": wi, "Outstanding": wt - wi}, wt, 14), unsafe_allow_html=True)
    st.subheader("Electricity per minisub")
    rows = []
    for m in MS_LIST:
        e = elec[elec["Minisub"] == m]; c2 = e["Status"].value_counts().to_dict()
        rows.append(f"<tr><td><b>{m}</b></td><td>{c2.get('Installed',0)}/{len(e)}</td><td style='width:55%'>{bar_html(c2, len(e))}</td></tr>")
    st.markdown("<table style='width:100%;font-size:14px;border-collapse:collapse'>" + "".join(rows) + "</table>", unsafe_allow_html=True)
    if not water.empty:
        st.subheader("Water per development")
        g = water.groupby("Development").agg(total=("Serial", "count"), on=("AMR Installed", "sum")).sort_values("total", ascending=False)
        rows = [f"<tr><td><b>{d}</b></td><td>{int(r.on)}/{int(r.total)}</td><td style='width:55%'>{bar_html({'Installed': int(r.on), 'Outstanding': int(r.total - r.on)}, int(r.total))}</td></tr>" for d, r in g.iterrows()]
        st.markdown("<table style='width:100%;font-size:14px;border-collapse:collapse'>" + "".join(rows) + "</table>", unsafe_allow_html=True)
    st.caption("Green = on AMR · amber = ordered / planned · red = still to do · grey = deferred.")


def page_water():
    st.header("💧 Water AMR")
    if water.empty:
        st.info(f"No water meter list found. Add `{WATER_LIST_PATTERN}` to the repo."); return
    st.caption("A water meter is on AMR when the site meter list marks it AMR-active (Sensus meters) or it reports on the LoRaWAN AMR platform.")
    devs = ["All"] + sorted(water["Development"].unique())
    d = st.selectbox("Development", devs)
    v = water if d == "All" else water[water["Development"] == d]
    wi = int(v["AMR Installed"].sum())
    c = st.columns(3); c[0].metric("Water meters", len(v)); c[1].metric("On AMR", wi, pct(wi, len(v))); c[2].metric("Not on AMR yet", len(v) - wi)
    q = st.text_input("🔍 Search stand or serial")
    if q: v = v[v.apply(lambda r: q.lower() in f"{r['Stand']} {r['Serial']}".lower(), axis=1)]
    out = v[["Stand", "Development", "Serial", "AMR Installed", "AMR Type", "AMR Label", "AMR Last Comms", "Reading"]].rename(columns={"AMR Installed": "On AMR"})
    st.dataframe(out.sort_values(["Development", "Stand"]), width="stretch", hide_index=True, height=520)


def _vendor(name):
    try:
        return open(os.path.join(HERE, "vendor", name), encoding="utf-8").read()
    except OSError:
        return ""


LEAFLET_JS, LEAFLET_CSS = _vendor("leaflet.js"), _vendor("leaflet.css")


def page_map():
    st.header("🗺️ Estate map")
    kp = latest(EST["roads_kmz"]) if EST["roads_kmz"] else None
    lines = kmz_lines(kp, os.path.getmtime(kp)) if kp else []
    ekeys = set(elec["Key"]); wkeys = set(water["Key"]) if not water.empty else set()
    feats = []
    for p in polys:
        mem = block_members(p["key"], ekeys | wkeys)
        er = elec[elec["Key"].isin(mem)]
        wr = water[water["Key"].isin(mem)] if not water.empty else pd.DataFrame()
        es, ws = elec_stand_status(er), water_stand_status(wr)
        label = p["name"] if not p["key"].startswith("BLOCK:") else f"{p['group']} {p['name']}"
        einfo = "<br>".join(f"{html.escape(str(r['Stand / Supply']))}: {r['Status']}" for _, r in er.head(14).iterrows()) or "No electricity meter yet"
        winfo = "<br>".join(f"{html.escape(r['Stand'])}: {'on AMR' if r['AMR Installed'] else 'not on AMR'}" for _, r in wr.head(14).iterrows()) if len(wr) else "No water meter on record"
        feats.append(dict(r=p["rings"], n=label, e=SCOL[es], w=SCOL[ws], es=es, ws=ws, ei=einfo, wi=winfo))
    kio = []
    for _, r in kiosks.iterrows():
        if str(r["Latitude"]).strip() in ("", "nan", "None") or r["Supplies"] == 0 or r["Kiosk"] == "MS LV board": continue
        e = elec[(elec["Minisub"] == r["Minisub"]) & (elec["Kiosk"] == r["Kiosk"])]
        kio.append(dict(lat=float(r["Latitude"]), lon=float(r["Longitude"]), n=f"{r['Minisub']} · {r['Kiosk']}",
                        i=f"{(e['Status']=='Installed').sum()}/{len(e)} on AMR"))
    msp = [dict(lat=float(m["Latitude"]), lon=float(m["Longitude"]), n=m["Minisub"]) for _, m in mss.iterrows()
           if str(m["Latitude"]).strip() not in ("", "nan", "None")]
    page = """
<style>__LCSS__</style><script>__LJS__</script>
<style>body{margin:0;font-family:Arial,sans-serif}html,body{height:100%}#map{height:100%;min-height:420px;border-radius:8px}
.leaflet-control-layers-toggle{background-image:none!important;width:36px!important;height:36px!important;display:flex!important;align-items:center;justify-content:center;font-size:20px;color:#152B45;text-decoration:none}
.leaflet-control-layers-toggle::after{content:'☰'}
.tg{display:flex;gap:0;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.35)}
.tg button{border:0;padding:9px 14px;font:600 14px Arial;background:#fff;color:#152B45;cursor:pointer}
.tg button.on{background:#152B45;color:#fff}
.lg{background:#fff;padding:7px 9px;border-radius:6px;font:12px Arial;line-height:19px;box-shadow:0 1px 4px rgba(0,0,0,.3)}
.lg i{display:inline-block;width:12px;height:12px;border-radius:2px;margin-right:6px;vertical-align:-2px}
.ms{background:#152B45;color:#fff;font:700 11px Arial;padding:2px 5px;border-radius:3px;white-space:nowrap;border:1px solid #fff}</style>
<div id="map"></div><script>
const P=__P__, KI=__K__, MS=__M__, L_=__L__, SC=__SC__;
const sat=L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',{maxZoom:20,maxNativeZoom:19,attribution:'Imagery © Esri'});
const street=L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{maxZoom:20,maxNativeZoom:19,attribution:'© OpenStreetMap'});
const map=L.map('map',{layers:[sat],maxZoom:20,tap:true});
let mode='e'; const stands=L.layerGroup().addTo(map);
function draw(){stands.clearLayers();P.forEach(p=>{const c=mode==='e'?p.e:p.w, s=mode==='e'?p.es:p.ws;
  L.polygon(p.r,{color:c,weight:1.2,fillColor:c,fillOpacity:s==='No meter yet'?0.12:0.55})
   .bindPopup(`<b>${p.n}</b><br><i>${s}</i><hr style="margin:4px 0">${mode==='e'?p.ei:p.wi}`).addTo(stands)});
  }
const kios=L.layerGroup(KI.map(k=>L.circleMarker([k.lat,k.lon],{radius:5,color:'#fff',weight:1.5,fillColor:'#152B45',fillOpacity:.95}).bindPopup(`<b>${k.n}</b><br>${k.i}`)));
const roads=L.layerGroup(L_.map(p=>L.polyline(p.path.map(q=>[q[1],q[0]]),{color:'#fff',weight:1,opacity:.5})));
const msl=L.layerGroup(MS.map(m=>L.marker([m.lat,m.lon],{icon:L.divIcon({className:'',html:`<div class="ms">${m.n}</div>`,iconAnchor:[0,10]})}))).addTo(map);
L.control.layers({'Satellite':sat,'Street map':street},{'Kiosks':kios,'Minisubs':msl,'Roads (KMZ)':roads},{collapsed:true}).addTo(map);
const tg=L.control({position:'topleft'});tg.onAdd=()=>{const d=L.DomUtil.create('div','tg');
  d.innerHTML='<button id="be" class="on">⚡ Electricity</button><button id="bw">💧 Water</button>';
  L.DomEvent.disableClickPropagation(d);
  d.querySelector('#be').onclick=()=>{mode='e';d.querySelector('#be').classList.add('on');d.querySelector('#bw').classList.remove('on');draw()};
  d.querySelector('#bw').onclick=()=>{mode='w';d.querySelector('#bw').classList.add('on');d.querySelector('#be').classList.remove('on');draw()};return d};tg.addTo(map);
const lg=L.control({position:'bottomleft'});lg.onAdd=()=>{const d=L.DomUtil.create('div','lg');d.innerHTML=Object.entries(SC).map(([k,c])=>`<i style="background:${c}"></i>${k}`).join('<br>');return d};lg.addTo(map);
draw(); const b=L.latLngBounds(P.flatMap(p=>p.r[0])); map.fitBounds(b,{padding:[10,10]});
</script>"""
    page = (page.replace("__P__", json.dumps(feats)).replace("__K__", json.dumps(kio)).replace("__M__", json.dumps(msp))
            .replace("__L__", json.dumps(lines)).replace("__SC__", json.dumps(SCOL)))
    page = page.replace("__LCSS__", LEAFLET_CSS).replace("__LJS__", LEAFLET_JS)
    if hasattr(st, "iframe"): st.iframe(page, height=600)
    else: components.html(page, height=600)
    st.caption("Use the ⚡ / 💧 buttons on the map to switch between electricity and water. Tap a stand for its meters. Kiosks, minisub labels, roads and street map can be switched on with the ☰ button. "
               "Blocks of flats (Marula, Highline, Husk) are coloured by the share of units on AMR. Satellite imagery © Esri.")


def page_lookup():
    st.header("🔎 Meter lookup")
    t1, t2 = st.tabs(["⚡ Electricity", "💧 Water"])
    with t1:
        page_all_elec()
    with t2:
        if water.empty: st.info("No water list loaded.")
        else:
            q = st.text_input("Search stand or serial", key="wq")
            v = water if not q else water[water.apply(lambda r: q.lower() in f"{r['Stand']} {r['Serial']}".lower(), axis=1)]
            st.dataframe(v[["Stand", "Development", "Serial", "AMR Installed", "AMR Type", "AMR Last Comms"]], width="stretch", hide_index=True, height=520)


def page_checks():
    st.header("⚠️ Data checks (staff)")
    t1, t2, t3 = st.tabs(["💧 Water", "⚡ Electricity", "🗺️ Stand map"])
    with t1:
        st.subheader("Water meters on AMR whose serial is not on the site list")
        st.caption("Matched on meter serial. These meters report on the AMR platform, but the serial is not in the site water-meter list.")
        if wdisc.empty: st.success("None — every AMR water serial is on the site list.")
        else: st.dataframe(wdisc, width="stretch", hide_index=True)
    with t2:
        page_checks_elec()
    with t3:
        st.subheader("Stand map issues")
        for n in poly_notes: st.write("• " + n)
        pk = {p["key"] for p in polys}
        bl = ("1350_", "1356_", "0511_")
        miss_w = sorted({k for k in water["Key"] if k and k not in pk and not k.startswith(bl)}) if not water.empty else []
        miss_e = sorted({k for k in elec["Key"] if k and k not in pk and not k.startswith(bl)})
        st.write(f"Water stands with no polygon ({len(miss_w)}):", ", ".join(miss_w))
        st.write(f"Electricity supplies with no polygon ({len(miss_e)}):", ", ".join(miss_e))




# =====================================================================
# WATER PLANNING
# =====================================================================
def poly_centroids():
    cen = {}
    for p in polys:
        r = p["rings"][0]
        cen.setdefault(p["key"], (sum(q[0] for q in r) / len(r), sum(q[1] for q in r) / len(r)))
    return cen


def nearest_ms(lat, lon):
    best, bd = None, 1e9
    for _, m in mss.iterrows():
        try:
            d = (float(m["Latitude"]) - lat) ** 2 + ((float(m["Longitude"]) - lon) * 0.87) ** 2
        except (TypeError, ValueError):
            continue
        if d < bd: best, bd = m["Minisub"], d
    return best


def water_sections():
    if water.empty: return water
    cen = poly_centroids(); w = water.copy()
    split = {"Freestanding", "Coral Cove"}
    def sec(r):
        if r["Development"] not in split:
            return r["Development"]
        c = cen.get(r["Key"])
        if not c: return f"{r['Development']} – no map location"
        return f"{r['Development']} – near {nearest_ms(*c)}"
    w["Section"] = w.apply(sec, axis=1)
    return w


def page_orders_water():
    if water.empty:
        st.info("No water meter list loaded for this estate."); return
    st.caption("One LoRaWAN water AMR device per water meter that is not on AMR yet. Complexes are their own section; freestanding "
               "stands" + (" and Coral Cove are" if ESTATE == "Zululami" else " are") + " grouped by the nearest minisub. Sections are "
               "ranked so the ones closest to finished come first.")
    w = water_sections()
    g = w.groupby("Section").agg(Meters=("Serial", "count"), On_AMR=("AMR Installed", "sum")).reset_index()
    g["To do"] = g["Meters"] - g["On_AMR"]
    if "Occupancy" in w:
        occ = w[(~w["AMR Installed"]) & (w["Occupancy"].str.startswith("Occupied"))].groupby("Section").size()
        g["Occupied to do"] = g["Section"].map(occ).fillna(0).astype(int)
    g["% done"] = (100 * g["On_AMR"] / g["Meters"]).round(0).astype(int)
    g = g[g["To do"] > 0].sort_values(["% done", "To do"], ascending=[False, True]).reset_index(drop=True)
    tot = int(g["To do"].sum())
    c = st.columns(3)
    c[0].metric("Water meters still to put on AMR", tot); c[1].metric("Sections to finish", len(g))
    n = c[2].number_input("Devices in the next order", min_value=0, value=min(100, tot), step=10, key="w_order")
    g["Cumulative devices"] = g["To do"].cumsum()
    g["Next order"] = g.apply(lambda r: "✅ Finish" if r["Cumulative devices"] <= n else ("◐ Part" if r["Cumulative devices"] - r["To do"] < n else ""), axis=1)
    g.insert(0, "Order", range(1, len(g) + 1))
    st.dataframe(g.rename(columns={"On_AMR": "On AMR"}), width="stretch", hide_index=True, height=min(60 + 35 * len(g), 640))
    pick = st.selectbox("Show stands still to do in a section", list(g["Section"]))
    v = w[(w["Section"] == pick) & (~w["AMR Installed"])]
    cols = [c_ for c_ in ["Stand", "Serial", "Occupancy"] if c_ in v]
    st.dataframe(v[cols].sort_values("Stand"), width="stretch", hide_index=True)
    st.download_button("⬇️ Water installation plan (CSV)", w[~w["AMR Installed"]].merge(g[["Section", "Order"]], on="Section")
                       .sort_values(["Order", "Stand"])[["Order", "Section"] + cols].to_csv(index=False).encode(), f"{EST['code']}_water_plan.csv", "text/csv")


# ---------- Navigation (sidebar) ----------
def staff_ok():
    try:
        pw = st.secrets.get("staff_password", "")
    except Exception:
        pw = ""
    if not pw:
        return False
    return st.session_state.get("staff_pw", "") == pw


if os.path.exists(os.path.join(HERE, "voltano_logo.png")):
    st.logo(os.path.join(HERE, "voltano_logo.png"), size="large")

pages = [st.Page(page_overview, title="Overview", icon="🏠", default=True),
         st.Page(page_map, title="Estate map", icon="🗺️", url_path="map"),
         st.Page(page_sld, title="Electricity – minisubs", icon="⚡", url_path="electricity"),
         st.Page(page_amr, title="Electricity – AMR progress", icon="📡", url_path="electricity-progress"),
         st.Page(page_water, title="Water", icon="💧", url_path="water"),
         st.Page(page_orders, title="Planned installations", icon="📅", url_path="planned"),
         st.Page(page_lookup, title="Meter lookup", icon="🔎", url_path="lookup")]
if staff_ok():
    pages.append(st.Page(page_checks, title="Data checks", icon="⚠️", url_path="checks"))
nav = st.navigation(pages, position="sidebar")
with st.sidebar:
    st.divider()
    st.text_input("Staff password", type="password", key="staff_pw", help="Voltano staff only – unlocks data checks.")
    if staff_ok(): st.success("Staff pages unlocked")
nav.run()
