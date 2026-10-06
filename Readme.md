# Zululami Estate — Reticulation & Smart Metering (AMR) Tracker

Streamlit app + A3 drawing generator tracking AMR (smart metering) progress across
Zululami's 14 minisubs, per kiosk and per supply. Same approach as EVG Lake Michelle.

## Repo contents
- `app.py` — the Streamlit app
- `make_sld.py` — builds the A3 landscape SLD drawing set (overview + one sheet per minisub)
- `ZLM_Meter_Hierarchy_YYYY-MM-DD.xlsx` — **master data** (the file you keep updating)
- `ZLM_AMR_Export_YYYY-MM-DD.xlsx` — optional: latest export from the AMR platform
- `ZLM Minisubs & kiosks.kmz` — kiosk / minisub positions and roads for the map
- `requirements.txt`

## Master workbook sheets
| Sheet | What it holds |
|---|---|
| `Elec` | One row per supply (meter point): minisub, kiosk, stand, breaker, phase, meter serial, occupancy, AMR installed / device / port / last comms, AMR order, source cell |
| `Kiosks` | Topology: which kiosk feeds which (`Fed From`), feed breaker, how the link is known (`Feed Link Basis`), planned AMR unit, KMZ point |
| `Minisubs` | Make, rating, bulk meter serial, bulk AMR device, LV board outgoing circuits |
| `AMR Devices` | Device (SNR) → location, channels used |
| `Orders` | Units and probes ordered, allocation |
| `Data Issues` | Known anomalies to resolve |

Built 06 Oct 2026 from `Zululami - Reticulation - 28 Sep 2026.xlsx` + the AMR export of 05 Oct 2026.
511 Husk (24 meters, MS 02) was added from the AMR export — it is not in the reticulation sheet.

## Updating
- **New AMR installs:** either set `AMR Installed` = TRUE (+ device/port) on the `Elec` rows, or
  just push a new `ZLM_AMR_Export_YYYY-MM-DD.xlsx` (same columns as the platform export:
  SNR, ADDRESS, METER NUMBER, LASTCOMMS, READING). The app matches the export to the workbook
  **on meter serial** and lists anything that doesn't match under ⚠️ Data Checks.
- **New meters / stands:** add or edit rows in `Elec` (keep the `Supply ID` unique).
- **New order:** add rows to `Orders` and put the order name (e.g. `Order 2`) in `AMR Order`
  for the supplies it covers.
- The app always uses the most recently modified matching workbook / export.

## Status colours
Installed (on AMR) · Ordered (on an open order) · Outstanding · Deferred (streetlights on the
MS 06 / MS 07 boards, excluded from Order 1).

## Topology
`Fed From` is taken from the reticulation sheet where it is explicit (MS LV board circuits,
feeder rows inside kiosks). Elsewhere it is inferred from the vertical order of kiosks in the
sheet (same column = daisy chain) and drawn dashed. Four kiosks could not be placed
(MS 03 K 3,3; MS 05 K 2,1, K 2,3, K 2,21). Correct `Fed From` in the `Kiosks` sheet once
confirmed against the as-built SLDs and set `Feed Link Basis` to "As-built SLD".

## Drawings
`python make_sld.py` writes `ZLM_Reticulation_SLD_AMR_<date>.pdf` (A3 landscape, sheet 1 =
site overview + key plan, sheets 2–15 = MS 01–MS 14). The same drawings can be downloaded from
the 🗼 Minisub SLD tab. Put `voltano_logo.png` next to the script to show the logo in the title block.
