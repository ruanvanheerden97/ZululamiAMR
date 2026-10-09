# Zululami, Coral Cove & Seaton — Smart Metering (AMR) Progress App

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


## HOA app layout (Oct 2026)
Sidebar navigation (collapses to the ☰ menu on phones): Overview · Estate map · Electricity – minisubs ·
Electricity – AMR progress · Water · Planned installations · Meter lookup. **Data checks** is staff-only.

### Staff password
In Streamlit Cloud: *App → Settings → Secrets* and add
```
staff_password = "choose-a-password"
```
Locally, copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` (git-ignored).
Without a password set, the Data checks page stays hidden.

### Water
- `ZLM_ALL_Water_meters_YYYY-MM-DD.csv` — site list of all water meters (Stand, SerialNumber1, IsAMR1)
- `ZLM_Water_AMR_Export_YYYY-MM-DD.xlsx` — AMR platform export (SNR, ADDRESS, METER NUMBER, LASTCOMMS, READING)
A water meter counts as **on AMR** when its serial appears in the latest export (8SEN/SEN and SN
prefixes are treated as the same serial). Export serials that are not on the site list are listed
under Data checks → Water.

### Estate map
Stand polygons come from the KMZ that contains a `Zululami Stands` folder (`ZululamiSeatonCoralCove.kmz`).
Seaton is ignored. Stands are matched across the reticulation sheet, water list, exports and KMZ by
`zlm_keys.py` (e.g. `505-01` = `0505_01`, `REM 501/01` = `Rem 501_01`, Coral Cove `38` = `038`,
water `BL A/00/01` = Husk 511-01, `BL A/00/001` = Coral Cove A001). Use the ⚡ / 💧 buttons on the map
to switch utility. Leaflet is bundled in `vendor/` so the map does not depend on a CDN.

## Estates (Oct 2026)
Three estates, each with its own HOA: **Zululami**, **Coral Cove** (MS 04 and MS 05 of the Zululami network – it shares the `ZLM_` files and is split out by minisub, and for water by the `CC-` stand key) and **Seaton**. Each estate's files are found by prefix:

| | Zululami (`ZLM_`) | Seaton (`SEA_`) |
|---|---|---|
| Electricity master | `ZLM_Meter_Hierarchy_*.xlsx` | `SEA_Meter_Hierarchy_*.xlsx` |
| Electricity AMR | `ZLM_AMR_Export_*.xlsx` (by serial) | `SEA_Device_List_*.csv` → `IsAMR` (by serial) |
| Water meters / AMR | `ZLM_ALL_Water_meters_*.csv` (`IsAMR1`) + `ZLM_Water_AMR_Export_*.xlsx` (LoRaWAN) | `SEA_Device_List_*.csv` → `IsAMR1` |
| Stand polygons | KMZ folder "Zululami Stands" | KMZ folder "Seaton polygons" |

To update Seaton, drop a newer `SEA_Device_List_YYYY-MM-DD.csv` in the folder (the latest date is used), push, and reboot.
Seaton PUD rows (future developments) are not counted as supplies.

## Water planned installations
Planned installations → 💧 Water splits the remaining water meters into sections:
complexes are their own section; freestanding stands (and Coral Cove) are grouped by their nearest minisub.
Sections are ranked closest-to-finished first, and the **Devices in the next order** box shows which sections
that order would finish (✅) or partly cover (◐). Assumes one LoRaWAN device per water meter not yet on AMR.

## Drawings
`python make_sld.py ZLM` or `python make_sld.py SEA` regenerates that estate's A3 SLD PDF from the latest master workbook.

## Logins
The app opens on a login screen. Each HOA password shows only that estate; the staff password shows all three
(with an Estate selector) plus Data checks. Set them in Streamlit Cloud → app → Settings → **Secrets**
(locally: `.streamlit/secrets.toml`, which is git-ignored – never commit passwords):

```toml
staff_password = "..."

[hoa_passwords]
"Zululami" = "..."
"Coral Cove" = "..."
"Seaton" = "..."
```

To change or revoke an HOA's access, edit its line in Secrets and save; the app picks it up straight away.
A login lasts for the browser session (closing the tab or refreshing logs out).

Drawings: `python make_sld.py ZLM` (12 minisubs), `python make_sld.py CC` (MS 04, MS 05), `python make_sld.py SEA`.
