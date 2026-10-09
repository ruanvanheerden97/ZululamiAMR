"""Stand-key normalisation shared by the app: maps the many spellings used in the
reticulation sheet, water site list, AMR exports and the stand KMZ to one key."""
import re

ALIAS = {"PAVILLION": "PAVILION", "WELLNESS CEN": "WELLNESS CENTRE", "16628": "1628"}
CORAL_NAMES = {"CARE CENTRE", "PAVILION", "WELLNESS CENTRE"}
COMMON = {"BULK WATER", "MUNIC BULK WATER", "GATEHOUSE", "CONTRACTORS GATEHOUSE", "SALES OFFICE", "UTILITYS", "UTILITIES"}
DEV_NAMES = {"0498": "Moyana", "0503": "The Reserve", "0505": "Aura", "0511": "Husk", "0512": "512", "0513": "Azure",
             "1336": "Savili", "1350": "Marula", "1356": "Highline", "1705": "Intaba Terrace", "1706": "Mbali Loop",
             "1707": "Sea La Vie", "REM501": "Valley Views", "PTN501": "North Shore", "CC": "Coral Cove"}


def _suffix(s):
    s = s.upper().replace(" ", "").strip("_-/")
    m = re.fullmatch(r"(\d+)", s)
    if m:
        return f"{int(m.group(1)):02d}"
    m = re.fullmatch(r"([A-Z])(\d+)", s)          # F1 -> F01
    if m and m.group(1) in "F":
        return f"{m.group(1)}{int(m.group(2)):02d}"
    m = re.fullmatch(r"(\d+)([A-Z])", s)          # 8A -> 08A
    if m:
        return f"{int(m.group(1)):02d}{m.group(2)}"
    return s


def stand_key(text, coral=False):
    """coral=True when the context says the stand is in Coral Cove (MS 04/05, Coral KMZ group)."""
    if text is None:
        return ""
    s = str(text).strip()
    if not s or s in ("-", "???", "nan"):
        return ""
    up = s.upper()
    tag = "REM" if "[VALLEY VIEW" in up else ("PTN" if "[NORTH SHORE" in up else "")
    u = re.sub(r"\s*\[.*?\]", "", up).strip()
    u = re.sub(r"\s+-\s+(NOT ON EMS|AMR BEHIND.*|NO USAGE.*|AURA|METER REMOVED|-AURA)$", "", u).strip()
    u = re.sub(r"\s*-AURA$", "", u)
    u = re.sub(r"^ZLM\s+", "", u)
    if u.startswith("CORAL COVE"):
        coral = True
        u = u.replace("CORAL COVE", "").strip(" -")
    u = re.sub(r"_NEW$", "", u).strip()
    u = re.sub(r"\s+(HUSK|HIGHLINE|MBALI LOOP|MARULA|MOYANA|SAVILI|AURA|AZURE|THE RESERVE)$", "", u).strip()
    if re.match(r"PTN_?501_BIN", u):
        return "PTN501_BC"
    u = ALIAS.get(u, u)
    if u in CORAL_NAMES:
        return f"CC-{u}"
    if tag and re.fullmatch(r"501\s*[-_/]\s*\w+", u):
        u = ("REM " if tag == "REM" else "PTN 4/") + u
    # Coral Cove apartment blocks / Husk blocks in the water list
    m = re.fullmatch(r"BL\s*([AB])/(\d\d)/(\d+)", u)
    if m:
        if len(m.group(3)) == 3:
            return f"CC-{m.group(1)}{m.group(3)}"
        return f"0511_{int(m.group(3)):02d}"
    if re.fullmatch(r"[AB]\d{3}", u):
        return f"CC-{u}"
    m = re.fullmatch(r"REM\s*(?:4/)?501\s*(?:[/_-]\s*(\w+))?\s*(BC)?", u)
    if m:
        return "REM501_" + _suffix(m.group(1) or m.group(2) or "BC")
    m = (re.fullmatch(r"PTN\s*/\s*501\s*(BC)", u) and re.fullmatch(r"PTN\s*/\s*501\s*()(BC)", u)) or re.fullmatch(r"PTN\s*(?:4\s*/)?\s*501\s*[/_-]\s*(\w+)()", u)
    if m:
        return "PTN501_" + _suffix(m.group(1) or "BC")
    m = re.fullmatch(r"0?(\d{3,4})\s*[-_/ ]\s*(.+)", u)
    if m and not re.fullmatch(r"\d+", m.group(2)) or (m and len(m.group(1)) in (3, 4)):
        base = f"{int(m.group(1)):04d}"
        suf = m.group(2).strip()
        if re.fullmatch(r"(BEACH|SPORTS)\s+PAVIL+ION", suf):
            return str(int(m.group(1)))
        return f"{base}_{_suffix(suf)}"
    m = re.fullmatch(r"(\d+)([A-Z]?)", u)
    if m:
        n = int(m.group(1))
        if coral or n <= 205:
            return f"CC-{n:03d}{m.group(2)}"
        return str(n)
    if coral:
        return f"CC-{u}"
    return u


def development(key):
    if not key:
        return "Unassigned"
    if key.startswith("CC-"):
        return "Coral Cove"
    head = key.split("_")[0]
    if head in DEV_NAMES:
        return DEV_NAMES[head]
    if key in COMMON or key.startswith("BULK"):
        return "Common / bulk"
    if re.fullmatch(r"\d{3,4}", key):
        return "Freestanding"
    return "Common / bulk"
