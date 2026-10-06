import glob, os, io, json, datetime, html, zipfile, re
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="Zululami — Reticulation & AMR", page_icon="⚡", layout="wide")

# ---------- Styling (same palette as the Lake Michelle / Sitari apps) ----------
st.markdown("""
<style>
.block-container {padding-top: 1.2rem; max-width: 1500px;}
div[data-testid="stMetric"] {background:#FBF9F3 !important;border:1px solid #DCD6C4;border-radius:10px;padding:10px 14px;}
div[data-testid="stMetric"] * {color:#152B45 !important;}
div[data-testid="stMetricLabel"], div[data-testid="stMetricLabel"] * {font-size:12px !important;text-transform:uppercase;letter-spacing:.05em;color:#3E5066 !important;}
div[data-testid="stMetricDelta"], div[data-testid="stMetricDelta"] * {color:#3F7D5C !important;fill:#3F7D5C !important;}
h1, h2, h3 {color:#152B45;}
</style>""", unsafe_allow_html=True)

HERE = os.path.dirname(os.path.abspath(__file__))
SITE_NAME = "Zululami Estate"
FILE_PATTERN = "ZLM_Meter_Hierarchy_*.xlsx"
AMR_EXPORT_PATTERN = "ZLM_AMR_Export_*.xlsx"     # optional: latest platform export, matched on meter serial
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
    elec, exp_notes = apply_export(elec, mss, exp)
elec["Status"] = elec.apply(status_of, axis=1)
elec["comms_dt"] = elec["AMR Last Comms"].apply(lambda v: parse_comms(v) if v else pd.NaT)
ref_now = (exp["comms"].max() if exp is not None else elec["comms_dt"].max())
if pd.isna(ref_now):
    ref_now = pd.Timestamp.now()
elec["Stale"] = elec["AMR Installed"] & elec["comms_dt"].notna() & ((ref_now - elec["comms_dt"]).dt.days > STALE_DAYS)
MS_LIST = sorted(mss["Minisub"].tolist(), key=lambda m: int(m[3:]))

st.title(f"⚡ {SITE_NAME} — Reticulation & Smart Metering (AMR)")
st.caption(f"Source: `{os.path.basename(data_path)}`" + (f" · AMR export: `{os.path.basename(exp_path)}`" if exp_path else " · no AMR export loaded (using workbook AMR columns)")
           + " · 14 minisubs · meters keyed on serial number")

# ---------- KPI strip ----------
cnt = elec["Status"].value_counts()
tot = len(elec)
done_ms = sum(1 for m in MS_LIST if len(elec[elec["Minisub"] == m]) and (elec[elec["Minisub"] == m]["Status"].isin(["Installed", "Deferred"])).all())
c = st.columns(6)
c[0].metric("Supplies (meter points)", tot, f"{(kiosks['Supplies'] > 0).sum()} kiosks / boards")
c[1].metric("On AMR", int(cnt.get("Installed", 0)), f"{round(100 * cnt.get('Installed', 0) / tot)}%")
c[2].metric("Ordered (open orders)", int(cnt.get("Ordered", 0)))
c[3].metric("Outstanding", int(cnt.get("Outstanding", 0)))
c[4].metric("Minisubs complete", f"{done_ms} / {len(MS_LIST)}")
c[5].metric("Bulk meters on AMR", f"{(mss['Bulk on AMR'] == 'Yes').sum()} / {len(mss)}")
st.divider()

tab_sld, tab_amr, tab_orders, tab_all, tab_map, tab_checks = st.tabs(
    ["🗼 Minisub SLD", "📡 AMR Progress", "📦 Orders & Next Minisub", "📋 All Supplies", "🗺️ Estate Map", "⚠️ Data Checks"])


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

with tab_sld:
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
with tab_amr:
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
with tab_orders:
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
with tab_all:
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

with tab_map:
    import math
    kp = latest("*.kmz")
    lines = kmz_lines(kp, os.path.getmtime(kp)) if kp else []
    lat0 = float(mss["Latitude"].astype(float).mean()); kx = math.cos(math.radians(lat0)) * 111320; ky = 110540
    allpts = [(p[0], p[1]) for l in lines for p in l["path"]] + list(zip(kiosks.loc[kiosks["Longitude"] != "", "Longitude"].astype(float), kiosks.loc[kiosks["Latitude"] != "", "Latitude"].astype(float)))
    lo0 = min(p[0] for p in allpts); la1 = max(p[1] for p in allpts)
    P = lambda lo, la: ((lo - lo0) * kx, (la1 - la) * ky)
    Wm = max(P(p[0], p[1])[0] for p in allpts) + 20; Hm = max(P(p[0], p[1])[1] for p in allpts) + 20
    paths = "".join('<path d="M' + " L".join(f"{P(*q)[0]:.1f},{P(*q)[1]:.1f}" for q in l["path"]) + '"/>' for l in lines)
    marks = []
    for _, r in kiosks.iterrows():
        if r["Latitude"] == "" or r["Supplies"] == 0 or r["Kiosk"] == "MS LV board": continue
        e = elec[(elec["Minisub"] == r["Minisub"]) & (elec["Kiosk"] == r["Kiosk"])]; c3 = e["Status"].value_counts().to_dict()
        s_ = "Installed" if c3.get("Installed", 0) == len(e) else ("Ordered" if c3.get("Ordered", 0) else ("Installed" if c3.get("Installed", 0) else ("Deferred" if c3.get("Deferred", 0) == len(e) else "Outstanding")))
        x, y = P(float(r["Longitude"]), float(r["Latitude"]))
        marks.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="{3 + len(e) ** 0.5 * 1.6:.1f}" fill="{COL[s_]}"><title>{html.escape(r["Minisub"] + " · " + r["Kiosk"])}\n{c3.get("Installed",0)}/{len(e)} on AMR · {s_}</title></circle>'
                     f'<text class="kl" x="{x + 6:.1f}" y="{y + 3:.1f}">{html.escape(str(r["Kiosk"]).replace("K ", ""))[:14]}</text>')
    for _, m in mss.iterrows():
        e = elec[elec["Minisub"] == m["Minisub"]]; x, y = P(float(m["Longitude"]), float(m["Latitude"]))
        pct = round(100 * (e["Status"] == "Installed").mean()) if len(e) else 0
        marks.append(f'<rect x="{x-6:.1f}" y="{y-6:.1f}" width="12" height="12" fill="#152B45"><title>{m["Minisub"]}: {pct}% on AMR</title></rect><text class="ml" x="{x+9:.1f}" y="{y-7:.1f}">{m["Minisub"]} · {pct}%</text>')
    svg = f"""<style>body{{margin:0;font-family:Arial}}svg{{width:100%;height:640px;background:#FBFCFD;border:1px solid #DCD6C4;cursor:grab}}
    path{{fill:none;stroke:#B9C3CC;stroke-width:1.2;vector-effect:non-scaling-stroke}}circle{{stroke:#fff;stroke-width:1}}
    .kl{{font-size:7px;fill:#3E5066;paint-order:stroke;stroke:#fff;stroke-width:2px}}.ml{{font-size:13px;font-weight:700;fill:#152B45;paint-order:stroke;stroke:#fff;stroke-width:3px}}
    .lg{{font-size:12px;color:#152B45;margin:4px 0}}.lg span{{margin-right:14px}}.lg i{{display:inline-block;width:10px;height:10px;border-radius:50%;margin-right:4px}}</style>
    <div class="lg">{"".join(f'<span><i style="background:{COL[k]}"></i>{k}</span>' for k in STATUS_ORDER)}<span>■ minisub</span><span>scroll to zoom · drag to pan · hover a kiosk</span></div>
    <svg id="m" viewBox="0 0 {Wm:.0f} {Hm:.0f}"><g>{paths}</g><g>{"".join(marks)}</g></svg>
    <script>const s=document.getElementById('m');let v=s.getAttribute('viewBox').split(' ').map(Number);const set=()=>s.setAttribute('viewBox',v.join(' '));
    s.addEventListener('wheel',e=>{{e.preventDefault();const r=s.getBoundingClientRect(),f=e.deltaY>0?1.2:1/1.2,
    sc=Math.max(v[2]/r.width,v[3]/r.height),cx=v[0]+(e.clientX-r.left)*sc,cy=v[1]+(e.clientY-r.top)*sc;v=[cx-(cx-v[0])*f,cy-(cy-v[1])*f,v[2]*f,v[3]*f];set()}},{{passive:false}});
    let d=null;s.addEventListener('pointerdown',e=>{{d=[e.clientX,e.clientY,v.slice()]}});window.addEventListener('pointerup',()=>d=null);
    s.addEventListener('pointermove',e=>{{if(!d)return;const r=s.getBoundingClientRect(),sc=Math.max(v[2]/r.width,v[3]/r.height);v=[d[2][0]-(e.clientX-d[0])*sc,d[2][1]-(e.clientY-d[1])*sc,v[2],v[3]];set()}});</script>"""
    if hasattr(st, "iframe"): st.iframe(svg, height=700)
    else: components.html(svg, height=700)
    st.caption("Kiosk colour = AMR status. Kiosks without a KMZ point are not shown — see Data Checks. Update the map by pushing a newer *.kmz to the repo.")

# =====================================================================
# CHECKS TAB
# =====================================================================
with tab_checks:
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
