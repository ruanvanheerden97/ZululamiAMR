"""Zululami – minisub reticulation SLDs with AMR status (A3 landscape PDF).
Reads the master workbook (latest ZLM_Meter_Hierarchy_*.xlsx next to this script)
and writes ZLM_Reticulation_SLD_AMR_<date>.pdf. Re-run whenever the workbook changes."""
import glob, os, sys, datetime, re
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.patches import Rectangle, FancyBboxPatch

HERE = os.path.dirname(os.path.abspath(__file__))
REV = "A"
AUTHOR = "Ruan van Heerden, Electrical Engineer"
COL = {"Installed": "#2E7D32", "Ordered": "#E3A008", "Outstanding": "#C0392B", "Deferred": "#8A8F98"}
INK = "#152B45"; GRID = "#9AA5B1"; PAPER = "white"
W, H = 420, 297            # A3 landscape, mm
M = 6                      # frame margin
TB_H = 30                  # title block height


def latest_workbook():
    files = glob.glob(os.path.join(HERE, "ZLM_Meter_Hierarchy_*.xlsx"))
    if not files:
        sys.exit("No ZLM_Meter_Hierarchy_*.xlsx found next to make_sld.py")
    return max(files, key=os.path.getmtime)


def fmt_sn(v):
    if pd.isna(v) or v == "": return ""
    try: return str(int(float(v)))
    except Exception: return str(v)


def supply_status(r):
    if bool(r["AMR Installed"]): return "Installed"
    if str(r.get("Comments", "")).startswith("Deferred"): return "Deferred"
    if str(r.get("AMR Order", "")).strip(): return "Ordered"
    return "Outstanding"


def load(path):
    elec = pd.read_excel(path, "Elec")
    elec["AMR Installed"] = elec["AMR Installed"].map(lambda v: v if isinstance(v, bool) else str(v).upper() == "TRUE")
    elec = elec.fillna("")
    elec["status"] = elec.apply(supply_status, axis=1)
    kiosks = pd.read_excel(path, "Kiosks").fillna("")
    mss = pd.read_excel(path, "Minisubs").fillna("")
    return elec, kiosks, mss


# ---------------------------------------------------------------- layout
ROW = 2.55; HDR = 7.2; BOXW = 47; GAPX = 9; GAPY = 3.0


def box_h(n):
    return HDR + max(n, 0) * ROW + (1.2 if n else 0)


def build_tree(ms, kiosks, elec):
    k = kiosks[(kiosks["Minisub"] == ms) & (kiosks["Kiosk"] != "MS LV board")]
    nodes = {}
    for _, r in k.iterrows():
        sup = elec[(elec["Minisub"] == ms) & (elec["Kiosk"] == r["Kiosk"])]
        nodes[r["Kiosk"]] = dict(name=r["Kiosk"], parent=r["Fed From"], basis=r["Feed Link Basis"], brk=str(r["Feed Breaker / Cable"]).strip(),
                                 main=str(r["Kiosk Main Breaker"]), unit=r["Planned AMR Unit"], dev=r["AMR Device(s)"], sup=sup, children=[])
    roots, unconf = [], []
    for n in nodes.values():
        p = n["parent"]
        if p in nodes: nodes[p]["children"].append(n)
        elif p == "MS LV board": roots.append(n)
        else: unconf.append(n)
    return roots, unconf


def measure(n):
    n["h"] = box_h(len(n["sup"]))
    if not n["children"]:
        n["sub"] = n["h"]; n["depth"] = 1; return
    for c in n["children"]: measure(c)
    ch = sum(c["sub"] for c in n["children"]) + GAPY * (len(n["children"]) - 1)
    n["sub"] = max(n["h"], ch); n["depth"] = 1 + max(c["depth"] for c in n["children"])


def place(n, x, y):          # y = top (units, downward)
    n["x"], n["y"] = x, y
    cy = y
    for c in n["children"]:
        place(c, x + BOXW + GAPX, cy); cy += c["sub"] + GAPY


GAPV = 7


def measure_v(n):
    n["h"] = box_h(len(n["sup"]))
    for c in n["children"]: measure_v(c)
    if n["children"]:
        n["subw"] = max(BOXW, sum(c["subw"] for c in n["children"]) + GAPX * (len(n["children"]) - 1))
        n["subh"] = n["h"] + GAPV + max(c["subh"] for c in n["children"])
    else:
        n["subw"], n["subh"] = BOXW, n["h"]


def place_v(n, x, y):
    n["x"], n["y"] = x, y
    cx = x
    for c in n["children"]:
        place_v(c, cx, y + n["h"] + GAPV); cx += c["subw"] + GAPX


def walk(n):
    yield n
    for c in n["children"]:
        yield from walk(c)


# ---------------------------------------------------------------- drawing
class Pg:
    def __init__(self, pdf):
        self.fig = plt.figure(figsize=(W / 25.4, H / 25.4)); self.pdf = pdf
        self.ax = self.fig.add_axes([0, 0, 1, 1]); self.ax.set_xlim(0, W); self.ax.set_ylim(0, H); self.ax.axis("off")
    def save(self):
        self.pdf.savefig(self.fig); plt.close(self.fig)


def frame(pg, title, sub, dwg, sheet, nsheets, date, stats):
    ax = pg.ax
    ax.add_patch(Rectangle((M, M), W - 2 * M, H - 2 * M, fill=False, lw=1.2, ec=INK))
    y0 = M; x0 = M
    ax.add_patch(Rectangle((x0, y0), W - 2 * M, TB_H, fill=False, lw=1.0, ec=INK))
    # logo / company
    logo = os.path.join(HERE, "voltano_logo.png")
    cw = 62
    if os.path.exists(logo):
        img = plt.imread(logo); ih, iw = img.shape[:2]; lw = min(cw - 6, (TB_H - 4) * iw / ih); lh = lw * ih / iw
        lx = x0 + (cw - lw) / 2
        ax.imshow(img, extent=(lx, lx + lw, y0 + (TB_H - lh) / 2, y0 + (TB_H + lh) / 2), zorder=5, interpolation="antialiased")
    else:
        ax.text(x0 + cw / 2, y0 + TB_H / 2 + 3, "VOLTANO", ha="center", va="center", fontsize=15, weight="bold", color=INK)
        ax.text(x0 + cw / 2, y0 + TB_H / 2 - 5, "METERING", ha="center", va="center", fontsize=9, color=INK, family="DejaVu Sans")
    ax.plot([x0 + cw] * 2, [y0, y0 + TB_H], color=INK, lw=0.8)
    # progress box
    px = x0 + cw; pw = 120
    ax.text(px + 3, y0 + TB_H - 4, "AMR PROGRESS (supplies)", fontsize=6.5, weight="bold", color=INK, va="center")
    tot = sum(stats.values()) or 1; bx = px + 3; by = y0 + 13; bw = pw - 6
    cx = bx
    for k in ["Installed", "Ordered", "Outstanding", "Deferred"]:
        v = stats.get(k, 0)
        if v: ax.add_patch(Rectangle((cx, by), bw * v / tot, 5, color=COL[k], lw=0)); cx += bw * v / tot
    ax.add_patch(Rectangle((bx, by), bw, 5, fill=False, ec=INK, lw=0.5))
    for i, k in enumerate(["Installed", "Ordered", "Outstanding", "Deferred"]):
        ax.add_patch(Rectangle((bx + i * 29, y0 + 4.5), 3, 3, color=COL[k], lw=0))
        ax.text(bx + i * 29 + 4.2, y0 + 6, f"{k} {stats.get(k, 0)}", fontsize=5.6, va="center", color=INK)
    ax.text(bx + bw, y0 + TB_H - 4, f"{round(100 * stats.get('Installed', 0) / tot)}% on AMR", fontsize=6.5, ha="right", va="center", color=COL["Installed"], weight="bold")
    ax.plot([px + pw] * 2, [y0, y0 + TB_H], color=INK, lw=0.8)
    # title
    tx = px + pw; tw = W - M - tx - 120
    ax.text(tx + 4, y0 + TB_H - 6, "ZULULAMI ESTATE", fontsize=8, color=GRID, weight="bold")
    ax.text(tx + 4, y0 + TB_H - 14, title, fontsize=12, color=INK, weight="bold")
    ax.text(tx + 4, y0 + 9, sub, fontsize=5.6, color=INK)
    ly = y0 + 3.5
    ax.plot([tx + 4, tx + 10], [ly, ly], color=INK, lw=0.7); ax.text(tx + 11, ly, "feed per sheet", fontsize=4.8, va="center", color=INK)
    ax.plot([tx + 32, tx + 38], [ly, ly], color=INK, lw=0.7, ls=(0, (3, 2))); ax.text(tx + 39, ly, "inferred from layout", fontsize=4.8, va="center", color=INK)
    ax.text(tx + 66, ly, "Rows: status · stand · S/N · ph · O/C/V · port", fontsize=4.8, va="center", color=GRID)
    ax.plot([tx + tw] * 2, [y0, y0 + TB_H], color=INK, lw=0.8)
    # fields: 3 rows x 2 cols
    fx = tx + tw; fw = 120; ch = TB_H / 3
    cells = [[("DRAWING No.", dwg), ("REVISION", REV)], [("DATE", date), ("SCALE", "NTS")], [("DRAWN", AUTHOR), ("SHEET", f"{sheet} of {nsheets}")]]
    for r, row in enumerate(cells):
        yy = y0 + TB_H - (r + 1) * ch
        widths = [fw * 0.7, fw * 0.3] if r == 2 else [fw / 2, fw / 2]
        cx = fx
        for (a, b), wc in zip(row, widths):
            ax.add_patch(Rectangle((cx, yy), wc, ch, fill=False, ec=INK, lw=0.5))
            ax.text(cx + 1.5, yy + ch - 2.3, a, fontsize=4.6, color=GRID)
            ax.text(cx + 1.5, yy + 2.2, b, fontsize=6.4 if r < 2 else 5.8, color=INK, weight="bold")
            cx += wc


def draw_kiosk(ax, n, X, Y, s):
    """X(u),Y(u) -> mm via closures; s = scale."""
    x, ytop = X(n["x"]), Y(n["y"]); w = BOXW * s; h = n["h"] * s
    sup = n["sup"]
    st = [r["status"] for _, r in sup.iterrows()]
    if not st: kcol = GRID
    elif all(v == "Installed" for v in st): kcol = COL["Installed"]
    elif any(v == "Ordered" for v in st): kcol = COL["Ordered"]
    elif any(v == "Installed" for v in st): kcol = COL["Installed"]
    else: kcol = COL["Outstanding"]
    ax.add_patch(Rectangle((x, ytop - h), w, h, fc="white", ec=kcol, lw=0.9 * s + 0.2, zorder=3))
    ax.add_patch(Rectangle((x, ytop - HDR * s), w, HDR * s, fc=kcol, ec=kcol, lw=0, alpha=0.16, zorder=3))
    fs = 6.2 * s
    ax.text(x + 1.2 * s, ytop - 2.6 * s, n["name"], fontsize=fs * 1.08, weight="bold", color=INK, va="center", zorder=4)
    meta = []
    if n["main"]: meta.append(f"Main {n['main']}")
    if len(sup): meta.append(f"{len(sup)} sup · {n['unit']}" if n["unit"] else f"{len(sup)} sup")
    if n["dev"]: meta.append("dev " + str(n["dev"])[-4:])
    ax.text(x + 1.2 * s, ytop - 5.7 * s, " · ".join(meta), fontsize=fs * 0.78, color=INK, va="center", zorder=4)
    ax.text(x + w - 1.2 * s, ytop - 2.6 * s, f"{sum(v == 'Installed' for v in st)}/{len(st)}" if st else "LV", fontsize=fs, color=kcol,
            ha="right", va="center", weight="bold", zorder=4)
    for i, (_, r) in enumerate(sup.iterrows()):
        yy = ytop - (HDR + 0.6 + (i + 0.5) * ROW) * s
        ax.add_patch(Rectangle((x + 1.0 * s, yy - 0.95 * s), 1.9 * s, 1.9 * s, color=COL[r["status"]], lw=0, zorder=4))
        occ = {"Occupied": "O", "Under construction": "C", "Vacant": "V"}.get(r["Occupancy"], "?")
        ax.text(x + 3.7 * s, yy, str(r["Stand / Supply"])[:14], fontsize=fs * 0.86, va="center", color=INK, zorder=4)
        ax.text(x + 22 * s, yy, fmt_sn(r["Meter Serial"]) or "–", fontsize=fs * 0.82, va="center", color=INK, family="DejaVu Sans Mono", zorder=4)
        ax.text(x + 35.5 * s, yy, str(r["Phase"])[:3], fontsize=fs * 0.8, va="center", color=INK, zorder=4)
        ax.text(x + 40 * s, yy, occ, fontsize=fs * 0.8, va="center", color=INK, zorder=4)
        port = r["AMR Port"]
        ax.text(x + w - 1.0 * s, yy, f"P{int(port)}" if str(port) not in ("", "nan") else "", fontsize=fs * 0.72, va="center", ha="right", color=COL["Installed"], zorder=4)


def lstyle(basis):
    return (0, (3, 2)) if ("Inferred" in basis or "Assumed" in basis) else ("-" if "sheet" in basis else (0, (1, 1.5)))


def short(label):
    return str(label).split("·")[0].split("(")[0].strip()[:8]


def link(ax, x1, y1, x2, y2, basis, label, s):
    xm = x1 + (x2 - x1) * 0.45
    ax.plot([x1, xm, xm, x2], [y1, y1, y2, y2], color=INK, lw=0.6 * s + 0.15, ls=lstyle(basis), zorder=2)
    if label:
        ax.text(x2 - 0.5 * s, y2 + 0.9 * s, short(label), fontsize=4.6 * s, color=INK, ha="right", zorder=4)


def link_v(ax, x1, y1, x2, y2, basis, label, s):
    ym = y2 + GAPV * 0.45 * s
    ax.plot([x1, x1, x2, x2], [y1, ym, ym, y2], color=INK, lw=0.6 * s + 0.15, ls=lstyle(basis), zorder=2)
    if label:
        ax.text(x2 + 0.8 * s, y2 + 1.2 * s, short(label), fontsize=4.6 * s, color=INK, zorder=4)


def ms_sheet(pdf, ms, elec, kiosks, mss, sheet, nsheets, date):
    pg = Pg(pdf); ax = pg.ax
    msr = mss[mss["Minisub"] == ms].iloc[0]
    msel = elec[elec["Minisub"] == ms]
    stats = msel["status"].value_counts().to_dict()
    frame(pg, f"{ms} – LV RETICULATION & AMR STATUS", f"{msr['Manufacturer']} · {msr['Rating']} · mfd {msr['Manufacture Date']} · bulk meter {fmt_sn(msr['Bulk Meter Serial'])}"
          f" ({'on AMR' if msr['Bulk on AMR'] == 'Yes' else 'not on AMR'})", f"ZLM-SLD-{ms[3:]}", sheet, nsheets, date, stats)
    roots, unconf = build_tree(ms, kiosks, elec)
    direct = msel[msel["Kiosk"] == "MS LV board"]
    MSW = 40
    ax_x0 = M + 6; aw = W - 2 * M - 12; ah = H - 2 * M - TB_H - 14
    # --- option A: horizontal (depth to the right)
    for r in roots + unconf: measure(r)
    y = 0
    for r in roots: place(r, MSW + 14, y); y += r["sub"] + GAPY * 2
    uyA = y + 6; y = uyA
    for r in unconf: place(r, MSW + 14, y); y += r["sub"] + GAPY * 2
    depth = max([r["depth"] for r in roots + unconf] or [1])
    hA = max(y, 30 + len(direct) * ROW * 1.2 + 10); wA = MSW + 14 + depth * BOXW + (depth - 1) * GAPX + 2
    sA = min(1.4, aw / wA, ah / hA)
    # --- option B: vertical (depth downward)
    for r in roots + unconf: measure_v(r)
    top = 36; x = 0
    for r in roots: x += 0; r["_x"] = x; x += r["subw"] + GAPX
    uxB = x + (8 if unconf else 0); x = uxB
    for r in unconf: r["_x"] = x; x += r["subw"] + GAPX
    wB = max(x, MSW + 80); hB = top + max([r["subh"] for r in roots + unconf] or [0]) + 2
    sB = min(1.4, aw / wB, ah / hB)
    vertical = sB > sA * 1.08
    s = sB if vertical else sA
    if vertical:
        for r in roots + unconf: measure_v(r); place_v(r, r["_x"], top)
    else:
        for r in roots + unconf: measure(r)
        y = 0
        for r in roots: place(r, MSW + 14, y); y += r["sub"] + GAPY * 2
        y = uyA
        for r in unconf: place(r, MSW + 14, y); y += r["sub"] + GAPY * 2
    X = lambda u: ax_x0 + u * s; Y = lambda u: H - M - 6 - u * s
    # MS block
    mx, my = X(0), Y(0)
    ax.add_patch(Rectangle((mx, my - 22 * s), MSW * s, 22 * s, fc="#EEF2F6", ec=INK, lw=1.2, zorder=3))
    ax.text(mx + MSW * s / 2, my - 6 * s, ms, ha="center", va="center", fontsize=11 * s, weight="bold", color=INK, zorder=4)
    ax.text(mx + MSW * s / 2, my - 12 * s, str(msr["Rating"]), ha="center", va="center", fontsize=5.6 * s, color=INK, zorder=4)
    ax.text(mx + MSW * s / 2, my - 16.5 * s, f"Bulk {fmt_sn(msr['Bulk Meter Serial'])}", ha="center", va="center", fontsize=5.6 * s, color=INK, zorder=4)
    bc = COL["Installed"] if msr["Bulk on AMR"] == "Yes" else COL["Outstanding"]
    ax.text(mx + MSW * s / 2, my - 20 * s, "Bulk on AMR" if msr["Bulk on AMR"] == "Yes" else "Bulk NOT on AMR", ha="center", va="center", fontsize=5.4 * s, color=bc, weight="bold", zorder=4)
    lw_bus = 2.2 * s + 0.4
    if vertical:
        busy = Y(29); xs_ = [X(r["x"] + 4) for r in roots] + [mx + MSW * s / 2]
        ax.plot([min(xs_), max(xs_)], [busy, busy], color=INK, lw=lw_bus, zorder=2)
        ax.plot([mx + MSW * s / 2] * 2, [my - 22 * s, busy], color=INK, lw=1.2, zorder=2)
        ax.text(max(xs_) + 1 * s, busy, "LV BOARD", fontsize=5 * s, va="center", color=INK)
        dx, dy = mx + MSW * s + 4 * s, my - 3 * s
    else:
        busx = X(MSW + 6)
        ys_ = [Y(r["y"] + 3.6) for r in roots] + [my - 11 * s]
        ax.plot([busx, busx], [min(ys_), my - 3 * s], color=INK, lw=lw_bus, zorder=2)
        ax.plot([mx + MSW * s, busx], [my - 11 * s] * 2, color=INK, lw=1.2, zorder=2)
        ax.text(busx, my - 1.5 * s, "LV BOARD", fontsize=5 * s, ha="center", color=INK)
        dx, dy = mx, my - 26 * s
    if len(direct):
        ax.text(dx, dy, "Direct from LV board:", fontsize=5.4 * s, color=INK, weight="bold")
        for i, (_, r) in enumerate(direct.iterrows()):
            yy = dy - (i + 1) * ROW * 1.15 * s
            ax.add_patch(Rectangle((dx, yy - 0.95 * s), 1.9 * s, 1.9 * s, color=COL[r["status"]], lw=0))
            ax.text(dx + 2.8 * s, yy, f"{str(r['Stand / Supply'])[:22]} · {r['Breaker']} · {fmt_sn(r['Meter Serial']) or '–'}", fontsize=5 * s, va="center", color=INK)
    for grp, is_root in ((roots, True), (unconf, False)):
        for r in grp:
            for n in walk(r):
                draw_kiosk(ax, n, X, Y, s)
                for c in n["children"]:
                    if vertical: link_v(ax, X(n["x"] + 4), Y(n["y"] + n["h"]), X(c["x"] + 4), Y(c["y"]), c["basis"], c["brk"], s)
                    else: link(ax, X(n["x"] + BOXW), Y(n["y"] + 3.6), X(c["x"]), Y(c["y"] + 3.6), c["basis"], c["brk"], s)
            if is_root:
                if vertical: link_v(ax, X(r["x"] + 4), Y(29), X(r["x"] + 4), Y(r["y"]), r["basis"], r["brk"], s)
                else: link(ax, X(MSW + 6), Y(r["y"] + 3.6), X(r["x"]), Y(r["y"] + 3.6), r["basis"], r["brk"], s)
    used = (wB if vertical else wA) * s
    if aw - used > 92:
        kiosk_table(ax, ms, kiosks, elec, W - M - 90, H - M - 8)
    if unconf:
        if vertical: ax.text(X(uxB), Y(top - 4), "FEED NOT CONFIRMED IN SHEET", fontsize=5.6 * s, color=COL["Outstanding"], weight="bold")
        else: ax.text(X(MSW + 14), Y(uyA - 2), "FEED NOT CONFIRMED IN SHEET – placement only", fontsize=5.6 * s, color=COL["Outstanding"], weight="bold")
    pg.save()


def kiosk_table(ax, ms, kiosks, elec, x0, y0):
    k = kiosks[(kiosks["Minisub"] == ms) & (kiosks["Supplies"] > 0)]
    cols = [("KIOSK", 0), ("SUP", 38), ("ON AMR", 47), ("ORD", 60), ("OUT", 69), ("UNIT", 77)]
    ax.text(x0, y0, "KIOSK AMR SCHEDULE", fontsize=7, weight="bold", color=INK)
    for t, dx in cols: ax.text(x0 + dx, y0 - 5, t, fontsize=4.8, color=GRID, weight="bold")
    y = y0 - 9
    for _, r in k.iterrows():
        e = elec[(elec["Minisub"] == ms) & (elec["Kiosk"] == r["Kiosk"])]; c = e["status"].value_counts()
        vals = [str(r["Kiosk"])[:24], len(e), c.get("Installed", 0), c.get("Ordered", 0), c.get("Outstanding", 0) + c.get("Deferred", 0),
                "existing" if c.get("Installed", 0) == len(e) else ("deferred" if c.get("Deferred", 0) == len(e) else str(r["Planned AMR Unit"]))]
        for (t, dx), v in zip(cols, vals):
            col = INK
            if t == "ON AMR" and v: col = COL["Installed"]
            if t == "ORD" and v: col = COL["Ordered"]
            if t == "OUT" and v: col = COL["Outstanding"]
            ax.text(x0 + dx, y, str(v), fontsize=5.2, color=col, va="center")
        ax.plot([x0, x0 + 88], [y - 1.8] * 2, color="#E3E7EC", lw=0.4)
        y -= 3.7


def read_kmz():
    import zipfile, math
    import xml.etree.ElementTree as ET
    files = glob.glob(os.path.join(HERE, "*.kmz"))
    if not files: return None
    K = "{http://www.opengis.net/kml/2.2}"
    with zipfile.ZipFile(max(files, key=os.path.getmtime)) as z:
        root = ET.fromstring(z.read([n for n in z.namelist() if n.endswith(".kml")][0]))
    lines, pts = [], {}
    def walk(e, path):
        for ch in e:
            if ch.tag in (K + "Folder", K + "Document"):
                nm = ch.find(K + "name"); walk(ch, path + [nm.text if nm is not None else ""])
            elif ch.tag == K + "Placemark":
                ls = ch.find(".//" + K + "LineString"); pt = ch.find(".//" + K + "Point"); nm = ch.find(K + "name")
                if ls is not None and any("ROADS" in p or "CADASTRAL" in p for p in path):
                    lines.append([tuple(map(float, c.split(",")[:2])) for c in ls.find(K + "coordinates").text.split()])
                elif pt is not None and nm is not None and re.fullmatch(r"MS \d+", nm.text or ""):
                    pts[nm.text] = tuple(map(float, pt.find(K + "coordinates").text.strip().split(",")[:2]))
    walk(root, [])
    return lines, pts


def keyplan(ax, elec, mss, x0, y0, x1, y1):
    data = read_kmz()
    if not data: return
    import math
    lines, pts = data
    allp = [p for l in lines for p in l] + list(pts.values())
    lo0, lo1 = min(p[0] for p in allp), max(p[0] for p in allp); la0, la1 = min(p[1] for p in allp), max(p[1] for p in allp)
    kx = math.cos(math.radians((la0 + la1) / 2))
    sc = min((x1 - x0) / ((lo1 - lo0) * kx), (y1 - y0) / (la1 - la0))
    P = lambda lo, la: (x0 + (lo - lo0) * kx * sc, y0 + (la - la0) * sc)
    for l in lines:
        xy = [P(*p) for p in l]; ax.plot([a for a, b in xy], [b for a, b in xy], color="#C9D1D9", lw=0.5)
    for ms, (lo, la) in pts.items():
        e = elec[elec["Minisub"] == ms]; pct = (e["status"] == "Installed").mean() if len(e) else 0
        ordered = (e["status"] == "Ordered").any()
        c = COL["Installed"] if pct >= 0.999 else (COL["Ordered"] if ordered else (COL["Installed"] if pct > 0 else COL["Outstanding"]))
        x, y = P(lo, la)
        ax.add_patch(Rectangle((x - 1.3, y - 1.3), 2.6, 2.6, color=c, zorder=5))
        ax.text(x + 2, y + 0.3, f"{ms} {round(pct * 100)}%", fontsize=4.8, va="center", color=INK, weight="bold", zorder=6)
    ax.text(x0, y1 + 2, "KEY PLAN (from KMZ) – square colour: green = AMR started/complete, amber = in current order, red = not started", fontsize=5.6, color=INK, weight="bold")


def overview(pdf, elec, kiosks, mss, nsheets, date):
    pg = Pg(pdf); ax = pg.ax
    st = elec["status"].value_counts().to_dict()
    frame(pg, "SITE OVERVIEW – AMR PROGRESS PER MINISUB", "Supplies per minisub from the master hierarchy workbook. Order 1 = MS 07 + MS 06.", "ZLM-SLD-00", 1, nsheets, date, st)
    cols = ["Minisub", "Rating", "Bulk on AMR", "Kiosks", "Supplies", "Installed", "Ordered", "Outstanding", "Deferred", "% on AMR"]
    xs = [M + 10, 40, 82, 104, 124, 146, 170, 194, 220, 244]
    y = H - M - 20
    ax.text(M + 10, y + 8, "AMR PROGRESS PER MINISUB", fontsize=11, weight="bold", color=INK)
    for c, x in zip(cols, xs): ax.text(x, y, c.upper(), fontsize=6, color=GRID, weight="bold")
    for i, (_, m) in enumerate(mss.iterrows()):
        ms = m["Minisub"]; e = elec[elec["Minisub"] == ms]; c = e["status"].value_counts()
        yy = y - 8 - i * 9.5; tot = len(e)
        vals = [ms, str(m["Rating"])[:16], m["Bulk on AMR"], kiosks[(kiosks["Minisub"] == ms) & (kiosks["Supplies"] > 0) & (kiosks["Kiosk"] != "MS LV board")].shape[0], tot,
                c.get("Installed", 0), c.get("Ordered", 0), c.get("Outstanding", 0), c.get("Deferred", 0), f"{round(100 * c.get('Installed', 0) / tot) if tot else 0}%"]
        for v, x in zip(vals, xs): ax.text(x, yy, str(v), fontsize=7, color=INK, va="center", weight="bold" if x == xs[0] else "normal")
        bx = 262; bw = 140; cx = bx
        for k in ["Installed", "Ordered", "Outstanding", "Deferred"]:
            v = c.get(k, 0)
            if v and tot: ax.add_patch(Rectangle((cx, yy - 2.2), bw * v / tot, 4.4, color=COL[k], lw=0)); cx += bw * v / tot
        ax.add_patch(Rectangle((bx, yy - 2.2), bw, 4.4, fill=False, ec=GRID, lw=0.4))
        ax.plot([M + 8, W - M - 8], [yy - 4.7] * 2, color="#E3E7EC", lw=0.5)
    yy = y - 8 - len(mss) * 9.5
    tot = len(elec)
    vals = ["TOTAL", "", "", "", tot, st.get("Installed", 0), st.get("Ordered", 0), st.get("Outstanding", 0), st.get("Deferred", 0), f"{round(100 * st.get('Installed', 0) / tot)}%"]
    for v, x in zip(vals, xs): ax.text(x, yy, str(v), fontsize=7.5, color=INK, va="center", weight="bold")
    keyplan(ax, elec, mss, 150, M + TB_H + 22, 255, yy - 10)
    ax.text(M + 10, M + TB_H + 14, "Notes: 1) Supply = one metered connection (1 probe). 511 Husk (MS 02) added from the AMR export – not in the reticulation sheet.  "
            "2) Kiosk feeds not stated in the sheet are inferred from the sheet layout (dashed) – verify against as-built SLDs.", fontsize=6, color=INK)
    ax.text(M + 10, M + TB_H + 8, "3) Status colours: green = on AMR, amber = in Order 1, red = outstanding, grey = deferred (streetlight on MS 06/07 board).", fontsize=6, color=INK)
    pg.save()


def write_set(target, path, only=None):
    elec, kiosks, mss = load(path)
    date = datetime.date.today().strftime("%Y-%m-%d")
    order = sorted(mss["Minisub"].tolist(), key=lambda m: int(m[3:]))
    n = len(order) + 1
    with PdfPages(target) as pdf:
        if only is None:
            overview(pdf, elec, kiosks, mss, n, date)
        for i, ms in enumerate(order):
            if only is None or ms == only:
                ms_sheet(pdf, ms, elec, kiosks, mss, i + 2, n, date)
        d = pdf.infodict(); d["Title"] = "Zululami – Reticulation SLDs & AMR status"; d["Author"] = AUTHOR


def build_pdf_bytes(path=None, only=None):
    import io
    buf = io.BytesIO(); write_set(buf, path or latest_workbook(), only); return buf.getvalue()


def main():
    path = latest_workbook()
    out = os.path.join(HERE, f"ZLM_Reticulation_SLD_AMR_{datetime.date.today():%Y-%m-%d}.pdf")
    write_set(out, path)
    print(out)


if __name__ == "__main__":
    main()
