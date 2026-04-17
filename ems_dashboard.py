"""
Energy Management System (EMS) Dashboard
Run:  streamlit run ems_dashboard.py
Deps: pip install streamlit plotly pandas openpyxl
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io

st.set_page_config(page_title="EMS Dashboard", layout="wide", page_icon="⚡")

# ══════════════════════════════════════════════════════════════
#  BUILT-IN SAMPLE DATA
# ══════════════════════════════════════════════════════════════

SAMPLE_INPUT_ROWS = [
    ("00:00:00", 65,   0,    0.95, 1, 0), ("00:10:00", 65,   0,    1.12, 1, 0),
    ("00:20:00", 65,   0,    1.00, 1, 0), ("00:30:00", 65,   0,    1.08, 1, 0),
    ("00:40:00", 65,   0,    1.12, 1, 0), ("00:50:00", 65,   0,    1.17, 1, 0),
    ("01:00:00", 65,   0,    1.02, 1, 0), ("01:30:00", 65,   0,    0.94, 1, 0),
    ("02:00:00", 65,   0,    0.96, 1, 0), ("03:00:00", 65,   0,    1.14, 1, 0),
    ("04:00:00", 65,   0,    1.00, 1, 0), ("05:00:00", 65,   0,    1.11, 1, 0),
    ("06:00:00", 65,   0,    2.79, 1, 0), ("06:10:00", 65,   0.22, 1.83, 1, 0),
    ("06:30:00", 65,   0.65, 2.16, 1, 0), ("07:00:00", 65,   1.29, 1.93, 1, 0),
    ("07:30:00", 66.2, 1.91, 1.85, 1, 0), ("08:00:00", 67.4, 2.50, 2.59, 1, 0),
    ("08:30:00", 71,   3.04, 2.28, 1, 0), ("09:00:00", 74.6, 3.54, 1.81, 1, 0),
    ("09:30:00", 78.2, 3.97, 2.04, 1, 0), ("10:00:00", 81.8, 4.33, 1.94, 1, 1),
    ("10:30:00", 85.4, 4.62, 2.35, 1, 1), ("11:00:00", 89,   4.83, 2.03, 1, 1),
    ("11:30:00", 92.6, 4.96, 2.19, 1, 1), ("12:00:00", 96.2, 5.00, 3.18, 1, 1),
    ("12:10:00", 97.4, 5.00, 3.49, 1, 1), ("12:40:00", 100,  4.92, 3.21, 1, 1),
    ("13:00:00", 100,  4.83, 3.35, 1, 1), ("13:10:00", 100,  4.77, 2.95, 1, 0),
    ("14:00:00", 100,  4.33, 3.34, 1, 0), ("15:00:00", 100,  3.54, 3.20, 1, 0),
    ("16:00:00", 100,  2.50, 2.67, 1, 0), ("16:50:00", 100,  1.50, 3.29, 1, 0),
    ("17:00:00", 98.5, 1.29, 4.29, 0, 0), ("17:30:00", 94,   0.65, 4.46, 0, 0),
    ("18:00:00", 89.5, 0,    4.31, 0, 1), ("18:10:00", 88,   0,    4.89, 0, 1),
    ("19:00:00", 80.5, 0,    4.89, 0, 1), ("20:00:00", 71.5, 0,    4.57, 0, 1),
    ("20:10:00", 71.5, 0,    4.11, 1, 0), ("21:00:00", 71.5, 0,    1.63, 1, 0),
    ("22:00:00", 71.5, 0,    1.75, 1, 0), ("23:00:00", 71.5, 0,    2.24, 1, 0),
    ("23:50:00", 71.5, 0,    2.45, 1, 0),
]

SAMPLE_RULES_ROWS = [
    (1,  0,1,0,0, 0,"","",0, 1,"","","", "RE",      "Idle",      "No Export", "Use RE with reduced load"),
    (2,  0,1,0,0, 0,"","",0, "",1,"","", "RE",      "Idle",      "No Export", "Use RE with reduced load"),
    (3,  0,1,0,0, 0,"","",0, "","",1,"", "RE",      "Idle",      "No Export", "Use RE with reduced load"),
    (4,  0,1,0,0, 0,"","",0, "","","",1, "RE",      "Idle",      "No Export", "Use RE with reduced load"),
    (5,  0,1,0,0, 1,1,"", 0, 1,"","","", "RE",      "Idle",      "No Export", "Use RE with reduced load"),
    (6,  0,1,0,0, 1,1,"", 0, "",1,"","", "RE",      "Idle",      "No Export", "Use RE with reduced load"),
    (7,  0,1,0,0, 1,1,"", 0, "","",1,"", "RE",      "Idle",      "No Export", "Use RE with reduced load"),
    (8,  0,1,0,0, 1,1,"", 0, "","","",1, "RE",      "Idle",      "No Export", "Use RE with reduced load"),
    (9,  0,1,0,0, 1,"",1, 0, 1,"","","", "Grid",    "Charge",    "No Export", "Grid primary, charge batt"),
    (10, 0,1,0,0, 1,"",1, 0, "",1,"","", "Grid",    "Charge",    "Export",    "RE Export"),
    (11, 0,1,0,0, 1,"",1, 0, "","",1,"", "Grid",    "Charge",    "Export",    "RE Export"),
    (12, 0,1,0,0, 1,"",1, 0, "","","",1, "Grid",    "Charge",    "No Export", "Grid primary"),
    (13, 0,0,1,0, 0,"","",0, 1,"","","", "Battery", "Discharge", "No Export", "Batt till 30%"),
    (14, 0,0,1,0, 0,"","",0, "",1,"","", "Battery", "Discharge", "No Export", "Batt till 30%"),
    (15, 0,0,1,0, 0,"","",0, "","",1,"", "Battery", "Discharge", "No Export", "Batt till 30%"),
    (16, 0,0,1,0, 0,"","",0, "","","",1, "Battery", "Discharge", "No Export", "Batt till 30%"),
    (17, 0,0,1,0, 1,1,"", 0, 1,"","","", "Battery", "Discharge", "No Export", "Batt till 30%, peak tariff"),
    (18, 0,0,1,0, 1,1,"", 0, "",1,"","", "Battery", "Discharge", "No Export", "Batt till 30%"),
    (19, 0,0,1,0, 1,1,"", 0, "","",1,"", "Battery", "Discharge", "No Export", "Batt till 30%"),
    (20, 0,0,1,0, 1,1,"", 0, "","","",1, "Battery", "Discharge", "No Export", "Batt till 30%, peak"),
    (21, 0,0,1,0, 1,"",1, 0, 1,"","","", "Grid",    "Charge",    "No Export", "Grid primary"),
    (22, 0,0,1,0, 1,"",1, 0, "",1,"","", "Grid",    "Charge",    "Export",    "Grid + RE Export"),
    (23, 0,0,1,0, 1,"",1, 0, "","",1,"", "Grid",    "Charge",    "Export",    "Grid + RE Export"),
    (24, 0,0,1,0, 1,"",1, 0, "","","",1, "Grid",    "Charge",    "No Export", "Grid primary"),
    (25, 0,0,0,1, 0,"","",0, 1,"","","", "Battery", "Discharge", "No Export", "Battery primary"),
    (26, 0,0,0,1, 0,"","",0, "",1,"","", "Battery", "Discharge", "Export",    "Battery + RE Export"),
    (27, 0,0,0,1, 0,"","",0, "","",1,"", "Battery", "Discharge", "Export",    "Battery + RE Export"),
    (28, 0,0,0,1, 0,"","",0, "","","",1, "Battery", "Discharge", "No Export", "Battery primary"),
    (29, 0,0,0,1, 1,1,"", 0, 1,"","","", "Battery", "Discharge", "No Export", "Battery primary, peak tariff"),
    (30, 0,0,0,1, 1,1,"", 0, "",1,"","", "Battery", "Discharge", "Export",    "Battery + RE Export"),
    (31, 0,0,0,1, 1,1,"", 0, "","",1,"", "Battery", "Discharge", "Export",    "Battery + RE Export"),
    (32, 0,0,0,1, 1,1,"", 0, "","","",1, "Battery", "Discharge", "No Export", "Battery primary"),
    (33, 0,0,0,1, 1,"",1, 0, 1,"","","", "Grid",    "Charge",    "No Export", "Grid primary, charge from RE"),
    (34, 0,0,0,1, 1,"",1, 0, "",1,"","", "Grid",    "Charge",    "No Export", "Grid primary"),
    (35, 0,0,0,1, 1,"",1, 0, "","",1,"", "Grid",    "Charge",    "No Export", "Grid primary"),
    (36, 0,0,0,1, 1,"",1, 0, "","","",1, "Grid",    "Charge",    "No Export", "Grid primary"),
    (37, 1,1,0,0, 0,"","",0, 1,"","","", "RE",      "Charge",    "No Export", "RE primary, charge batt"),
    (38, 1,1,0,0, 0,"","",0, "",1,"","", "RE",      "Charge",    "No Export", "RE primary, charge batt"),
    (39, 1,1,0,0, 0,"","",0, "","",1,"", "RE",      "Charge",    "No Export", "RE primary, charge batt"),
    (40, 1,1,0,0, 0,"","",0, "","","",1, "RE",      "Charge",    "No Export", "RE primary, charge batt"),
    (41, 1,1,0,0, 1,1,"", 0, 1,"","","", "RE",      "Charge",    "Export",    "RE primary + export"),
    (42, 1,1,0,0, 1,1,"", 0, "",1,"","", "RE",      "Charge",    "Export",    "RE primary + export"),
    (43, 1,1,0,0, 1,1,"", 0, "","",1,"", "RE",      "Charge",    "Export",    "RE primary + export"),
    (44, 1,1,0,0, 1,1,"", 0, "","","",1, "RE",      "Charge",    "Export",    "RE primary + export"),
    (45, 1,1,0,0, 1,"",1, 0, 1,"","","", "RE",      "Charge",    "Export",    "Battery Charge - Grid"),
    (46, 1,1,0,0, 1,"",1, 0, "",1,"","", "RE",      "Charge",    "Export",    "Battery Charge - Grid"),
    (47, 1,1,0,0, 1,"",1, 0, "","",1,"", "RE",      "Charge",    "Export",    "Battery Charge - Grid"),
    (48, 1,1,0,0, 1,"",1, 0, "","","",1, "RE",      "Charge",    "Export",    "Battery Charge - Grid"),
    (49, 1,0,1,0, 0,"","",0, 1,"","","", "RE",      "Charge",    "No Export", "RE primary"),
    (50, 1,0,1,0, 0,"","",0, "",1,"","", "RE",      "Charge",    "No Export", "RE primary"),
    (51, 1,0,1,0, 0,"","",0, "","",1,"", "RE",      "Charge",    "No Export", "RE primary"),
    (52, 1,0,1,0, 0,"","",0, "","","",1, "RE",      "Charge",    "No Export", "RE primary"),
    (53, 1,0,1,0, 1,1,"", 0, 1,"","","", "RE",      "Charge",    "Export",    "RE + Export"),
    (54, 1,0,1,0, 1,1,"", 0, "",1,"","", "RE",      "Charge",    "Export",    "RE + Export"),
    (55, 1,0,1,0, 1,1,"", 0, "","",1,"", "RE",      "Charge",    "Export",    "RE + Export"),
    (56, 1,0,1,0, 1,1,"", 0, "","","",1, "RE",      "Charge",    "Export",    "RE + Export"),
    (57, 1,0,1,0, 1,"",1, 0, 1,"","","", "RE",      "Charge",    "Export",    "RE + Export, Batt Grid"),
    (58, 1,0,1,0, 1,"",1, 0, "",1,"","", "RE",      "Charge",    "Export",    "RE + Export"),
    (59, 1,0,1,0, 1,"",1, 0, "","",1,"", "RE",      "Charge",    "Export",    "RE + Export"),
    (60, 1,0,1,0, 1,"",1, 0, "","","",1, "RE",      "Charge",    "Export",    "RE + Export"),
    (61, 1,0,0,1, 0,"","",0, 1,"","","", "RE",      "Idle",      "Export",    "RE primary, Batt full"),
    (62, 1,0,0,1, 0,"","",0, "",1,"","", "RE",      "Idle",      "Export",    "RE primary + export"),
    (63, 1,0,0,1, 0,"","",0, "","",1,"", "RE",      "Idle",      "Export",    "RE primary + export"),
    (64, 1,0,0,1, 0,"","",0, "","","",1, "RE",      "Idle",      "Export",    "RE primary + export"),
    (65, 1,0,0,1, 1,1,"", 0, 1,"","","", "RE",      "Idle",      "Export",    "RE + export"),
    (66, 1,0,0,1, 1,1,"", 0, "",1,"","", "RE",      "Idle",      "Export",    "RE + export"),
    (67, 1,0,0,1, 1,1,"", 0, "","",1,"", "RE",      "Idle",      "Export",    "RE + export"),
    (68, 1,0,0,1, 1,1,"", 0, "","","",1, "RE",      "Idle",      "Export",    "RE + export"),
    (69, 1,0,0,1, 1,"",1, 0, 1,"","","", "RE",      "Idle",      "Export",    "RE + export"),
    (70, 1,0,0,1, 1,"",1, 0, "",1,"","", "RE",      "Idle",      "Export",    "RE + export"),
    (71, 1,0,0,1, 1,"",1, 0, "","",1,"", "RE",      "Idle",      "Export",    "RE + export"),
    (72, 1,0,0,1, 1,"",1, 0, "","","",1, "RE",      "Idle",      "Export",    "RE + export"),
]

RULES_COLS = ["Rule No","RE","Batt_L","Batt_M","Batt_H","Grid","Grid_T_H","Grid_T_L",
              "DG","TOD1","TOD2","TOD3","TOD4","Load","Batt","Grid_out","Note"]
INPUT_COLS  = ["Time","Battery SOC (%)","RE (kW)","Load (kW)","Grid Available","Tariff"]


def _make_excel(df, sheet_name):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=sheet_name)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════
#  RULES PARSER
# ══════════════════════════════════════════════════════════════

def _flag(val):
    if val is None: return None
    if isinstance(val, float) and pd.isna(val): return None
    s = str(val).strip()
    if s in ("","x","X","nan","None"): return None
    try: return int(float(s))
    except: return None


def parse_rules_df(df):
    rules = []
    for _, row in df.iterrows():
        rno = _flag(row.get("Rule No"))
        if rno is None: continue
        rules.append({
            "rule": rno,
            "RE":   _flag(row.get("RE")),
            "BL":   _flag(row.get("Batt_L")),
            "BM":   _flag(row.get("Batt_M")),
            "BH":   _flag(row.get("Batt_H")),
            "Grid": _flag(row.get("Grid")),
            "GTH":  _flag(row.get("Grid_T_H")),
            "GTL":  _flag(row.get("Grid_T_L")),
            "DG":   _flag(row.get("DG")),
            "TOD1": _flag(row.get("TOD1")),
            "TOD2": _flag(row.get("TOD2")),
            "TOD3": _flag(row.get("TOD3")),
            "TOD4": _flag(row.get("TOD4")),
            "Load":     str(row.get("Load","Grid")).strip(),
            "Batt":     str(row.get("Batt","Idle")).strip(),
            "Grid_out": str(row.get("Grid_out","No Export")).strip(),
            "note":     str(row.get("Note","")).strip(),
        })
    return rules


_DEFAULT_RULES = parse_rules_df(pd.DataFrame(SAMPLE_RULES_ROWS, columns=RULES_COLS))


# ══════════════════════════════════════════════════════════════
#  EMS LOGIC
# ══════════════════════════════════════════════════════════════

def get_tod(ts):
    """Extract hour from either 'YYYY-MM-DD HH:MM' / 'YYYY-MM-DD HH:MM:SS'
    or legacy 'HH:MM:SS' / 'HH:MM' formats."""
    s = str(ts).strip()
    # If there is a space, the part after it is the time component
    if " " in s:
        s = s.split(" ", 1)[1]
    # If there is a 'T' (ISO 8601), split on that
    elif "T" in s:
        s = s.split("T", 1)[1]
    h = int(s.split(":")[0])
    return (1 if 0<=h<6 else 0, 1 if 6<=h<12 else 0,
            1 if 12<=h<18 else 0, 1 if 18<=h<24 else 0)


def get_batt_range(soc):
    return (1 if soc<30 else 0, 1 if 30<=soc<70 else 0, 1 if soc>=70 else 0)


def match_rule(re_f,bl,bm,bh,grid,gth,gtl,dg,t1,t2,t3,t4,rules):
    for r in rules:
        if r["RE"]   is not None and r["RE"]  !=re_f:  continue
        if r["BL"]   is not None and r["BL"]  !=bl:    continue
        if r["BM"]   is not None and r["BM"]  !=bm:    continue
        if r["BH"]   is not None and r["BH"]  !=bh:    continue
        if r["Grid"] is not None and r["Grid"]!=grid:   continue
        if r["GTH"]  is not None and r["GTH"] !=gth:    continue
        if r["GTL"]  is not None and r["GTL"] !=gtl:    continue
        if r["DG"]   is not None and r["DG"]  !=dg:     continue
        if not any(r[k]==1 and v==1 for k,v in
                   [("TOD1",t1),("TOD2",t2),("TOD3",t3),("TOD4",t4)]): continue
        return r["Load"],r["Batt"],r["Grid_out"],r["rule"],r["note"]
    return "Grid","Idle","No Export",0,"Fallback: Grid"


def compute_power(load_kw, re_kw, load_src, batt_mode, grid_out):
    if load_src == "RE":
        excess  = re_kw - load_kw
        batt_kw = max(0, excess) if batt_mode=="Charge" else 0
        grid_kw = (-max(0,excess-batt_kw) if grid_out=="Export" and excess>0 else 0)
        if excess < 0: grid_kw = -excess
    elif load_src == "Battery":
        batt_kw = -load_kw
        grid_kw = -re_kw if grid_out=="Export" and re_kw>0 else 0
    elif load_src == "Grid":
        grid_kw = load_kw
        batt_kw = re_kw if batt_mode=="Charge" else 0
        if grid_out=="Export" and re_kw>0: grid_kw -= re_kw
    else:
        batt_kw, grid_kw = 0, load_kw
    return round(load_kw,3), round(batt_kw,3), round(grid_kw,3)


def _normalise_time(val):
    """Return a clean string for a Time value.

    Accepts any of:
      - pandas Timestamp / datetime objects
      - 'YYYY-MM-DD HH:MM'  or  'YYYY-MM-DD HH:MM:SS'
      - 'HH:MM:SS'  or  'HH:MM'   (legacy, no date)
    """
    if hasattr(val, "strftime"):          # Timestamp / datetime
        return val.strftime("%Y-%m-%d %H:%M:%S")
    s = str(val).strip()
    try:
        parsed = pd.to_datetime(s, dayfirst=False)
        if parsed.year < 1970:            # Excel time-only epoch
            return parsed.strftime("%H:%M:%S")
        return parsed.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return s


@st.cache_data
def process_dataframe(df, rules):
    out = []
    for _, row in df.iterrows():
        ts  = _normalise_time(row["Time"])
        soc = float(row["Battery SOC (%)"])
        re  = float(row["RE (kW)"])
        ld  = float(row["Load (kW)"])
        ga  = int(row["Grid Available"])
        tar = int(row["Tariff"])
        t1,t2,t3,t4 = get_tod(ts)
        bl,bm,bh     = get_batt_range(soc)
        re_f = 1 if re>=ld else 0
        gth  = 1 if tar==1 else 0
        gtl  = 1 if tar==0 else 0
        src,bmode,gmode,rno,note = match_rule(re_f,bl,bm,bh,ga,gth,gtl,0,t1,t2,t3,t4,rules)
        lv,bv,gv = compute_power(ld,re,src,bmode,gmode)
        grid_import = round(max(0,  gv), 3)   # positive grid_out = importing
        grid_export = round(max(0, -gv), 3)   # negative grid_out = exporting
        out.append({"Time":ts,"Battery SOC (%)":soc,"RE (kW)":re,"Load (kW)":ld,
                    "Grid Available":ga,"Tariff":tar,"Rule No":rno,
                    "Load Source":src,"Batt Mode":bmode,"Grid Mode":gmode,
                    "Load_out (kW)":lv,"Batt_out (kW)":bv,"Grid_out (kW)":gv,
                    "Grid Import (kW)":grid_import,"Grid Export (kW)":grid_export,
                    "Note":note})
    return pd.DataFrame(out)


# ══════════════════════════════════════════════════════════════
#  UI
# ══════════════════════════════════════════════════════════════

st.title("⚡ Energy Management System Dashboard")
st.caption("Upload your rules & input data — or download the sample templates and use them to get started instantly.")

# ── Sidebar ────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Configuration")

    # Sample downloads
    st.subheader("📥 Download Sample Files")
    st.caption("Fill these templates and re-upload below.")

    sample_input_xl = _make_excel(pd.DataFrame(SAMPLE_INPUT_ROWS, columns=INPUT_COLS), "Sample Input")
    sample_rules_xl = _make_excel(pd.DataFrame(SAMPLE_RULES_ROWS, columns=RULES_COLS), "Rules")

    st.download_button("⬇️ Sample Input Data (.xlsx)", sample_input_xl,
                       "sample_input_data.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)
    st.download_button("⬇️ Sample Rules File (.xlsx)", sample_rules_xl,
                       "sample_rules.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)

    st.divider()

    # Rules upload
    st.subheader("📐 Rules File")
    rules_file = st.file_uploader("Upload custom rules (.xlsx/.csv)", type=["xlsx","csv"],
                                  help=f"Columns: {', '.join(RULES_COLS)}")
    if rules_file:
        try:
            df_r = (pd.read_csv(rules_file) if rules_file.name.endswith(".csv")
                    else pd.read_excel(rules_file))
            df_r.columns = [c.strip() for c in df_r.columns]
            active_rules = parse_rules_df(df_r)
            st.success(f"✅ {len(active_rules)} custom rules loaded")
        except Exception as e:
            st.error(f"Rules error: {e}")
            active_rules = _DEFAULT_RULES
    else:
        active_rules = _DEFAULT_RULES
        st.info(f"Using {len(_DEFAULT_RULES)} built-in default rules")

    st.divider()

    # Input data
    st.subheader("📂 Input Data")
    input_mode = st.radio("Source", ["Use Sample Data", "Upload File", "Manual Entry"], index=0)

    df_input = None

    if input_mode == "Use Sample Data":
        df_input = pd.DataFrame(SAMPLE_INPUT_ROWS, columns=INPUT_COLS)
        st.success(f"✅ {len(df_input)} sample rows ready")

    elif input_mode == "Upload File":
        data_file = st.file_uploader("Upload input data (.xlsx/.csv)", type=["xlsx","csv"],
                                     help=f"Columns: {', '.join(INPUT_COLS)}")
        if data_file:
            try:
                df_raw = (pd.read_csv(data_file) if data_file.name.endswith(".csv")
                          else pd.read_excel(data_file))
                df_raw.columns = [c.strip() for c in df_raw.columns]
                missing = [c for c in INPUT_COLS if c not in df_raw.columns]
                if missing:
                    st.error(f"Missing columns: {missing}")
                else:
                    df_input = df_raw[INPUT_COLS].dropna()
                    st.success(f"✅ {len(df_input)} rows loaded")
            except Exception as e:
                st.error(f"File error: {e}")

    else:  # Manual Entry
        n = st.slider("Rows", 1, 50, 5)
        df_edit = st.data_editor(
            pd.DataFrame({c:[""]*n for c in INPUT_COLS}),
            use_container_width=True, num_rows="dynamic",
            column_config={
                "Time": st.column_config.TextColumn("Time (HH:MM:SS or YYYY-MM-DD HH:MM:SS)"),
                "Battery SOC (%)": st.column_config.NumberColumn(min_value=0, max_value=100),
                "RE (kW)":         st.column_config.NumberColumn(min_value=0.0),
                "Load (kW)":       st.column_config.NumberColumn(min_value=0.0),
                "Grid Available":  st.column_config.SelectboxColumn(options=[0,1]),
                "Tariff":          st.column_config.SelectboxColumn(options=[0,1]),
            })
        try:
            for c in INPUT_COLS[1:]:
                df_edit[c] = pd.to_numeric(df_edit[c], errors="coerce")
            df_input = df_edit.dropna(subset=["Time"]).dropna()
        except Exception:
            df_input = None

    st.divider()

    # Generate button
    can_run = df_input is not None and not df_input.empty
    generate = st.button("🚀 Generate EMS Output", type="primary",
                         use_container_width=True, disabled=not can_run)

# ══════════════════════════════════════════════════════════════
#  MAIN PANEL
# ══════════════════════════════════════════════════════════════

# Auto-run ONCE on first ever load when sample data is selected.
# Use a session flag so subsequent re-runs (widget interactions) don't
# re-trigger this and wipe the stored result.
auto_run = (input_mode == "Use Sample Data") and can_run and ("df_result" not in st.session_state)

should_run = generate or auto_run

if should_run and can_run:
    # On explicit Generate click, bust the @cache_data cache so new
    # data/rules are always recomputed; auto_run reuses cache if inputs unchanged.
    if generate:
        process_dataframe.clear()
    with st.spinner("Running EMS logic…"):
        st.session_state["df_result"]    = process_dataframe(df_input, active_rules)
        st.session_state["rules_file"]   = rules_file
        st.session_state["active_rules"] = active_rules

if "df_result" in st.session_state:
    df_result    = st.session_state["df_result"]
    rules_file   = st.session_state.get("rules_file", rules_file)
    active_rules = st.session_state.get("active_rules", active_rules)

    rules_src = "custom" if rules_file else "default"
    st.success(
        f"✅ Processed **{len(df_result)} rows** using **{rules_src} rules** "
        f"— **{df_result['Rule No'].nunique()}** unique rules fired"
    )

    # KPIs
    total_load   = df_result['Load_out (kW)'].sum()
    re_gen       = df_result['RE (kW)'].sum()
    grid_import  = df_result[df_result['Grid_out (kW)']>0]['Grid_out (kW)'].sum()
    grid_export  = abs(df_result[df_result['Grid_out (kW)']<0]['Grid_out (kW)'].sum())
    batt_net     = df_result['Batt_out (kW)'].sum()
    re_pct       = (re_gen / total_load * 100) if total_load > 0 else 0
    self_suf     = max(0, min(100, (total_load - grid_import) / total_load * 100)) if total_load > 0 else 0

    kpi_html = f"""
    <style>
    .kpi-grid {{
        display: grid;
        grid-template-columns: repeat(5, 1fr);
        gap: 12px;
        margin-bottom: 8px;
    }}
    .kpi-card {{
        background: #1e2130;
        border: 1px solid #2e3250;
        border-radius: 10px;
        padding: 14px 16px;
        text-align: center;
        position: relative;
        cursor: default;
        transition: border-color 0.2s, transform 0.15s;
    }}
    .kpi-card:hover {{
        border-color: #4a90e2;
        transform: translateY(-2px);
    }}
    .kpi-card:hover .kpi-tooltip {{
        opacity: 1;
        pointer-events: auto;
    }}
    .kpi-icon  {{ font-size: 20px; margin-bottom: 4px; }}
    .kpi-label {{ font-size: 11px; color: #8a9bb0; text-transform: uppercase;
                  letter-spacing: .06em; margin-bottom: 6px; }}
    .kpi-value {{ font-size: 22px; font-weight: 700; color: #e8edf5;
                  white-space: nowrap; overflow: visible; }}
    .kpi-sub   {{ font-size: 11px; color: #5a7a9a; margin-top: 4px; }}
    .kpi-tooltip {{
        opacity: 0;
        pointer-events: none;
        position: absolute;
        bottom: calc(100% + 8px);
        left: 50%; transform: translateX(-50%);
        background: #0d1117;
        border: 1px solid #4a90e2;
        border-radius: 8px;
        padding: 10px 14px;
        font-size: 12px;
        color: #cdd9e5;
        white-space: nowrap;
        z-index: 999;
        transition: opacity 0.2s;
        text-align: left;
        line-height: 1.7;
        box-shadow: 0 4px 20px rgba(0,0,0,0.5);
    }}
    .kpi-tooltip::after {{
        content: '';
        position: absolute;
        top: 100%; left: 50%; transform: translateX(-50%);
        border: 6px solid transparent;
        border-top-color: #4a90e2;
    }}
    </style>

    <div class="kpi-grid">

      <div class="kpi-card">
        <div class="kpi-icon">⚡</div>
        <div class="kpi-label">Total Load</div>
        <div class="kpi-value">{total_load:.1f} kWh</div>
        <div class="kpi-sub">Site consumption</div>
        <div class="kpi-tooltip">
          <b>Total Load</b><br>
          Exact: {total_load:.3f} kWh<br>
          Rows: {len(df_result)}<br>
          Avg per row: {total_load/len(df_result):.3f} kW
        </div>
      </div>

      <div class="kpi-card">
        <div class="kpi-icon">☀️</div>
        <div class="kpi-label">RE Generation</div>
        <div class="kpi-value">{re_gen:.1f} kWh</div>
        <div class="kpi-sub">{re_pct:.1f}% of load</div>
        <div class="kpi-tooltip">
          <b>Renewable Generation</b><br>
          Exact: {re_gen:.3f} kWh<br>
          RE coverage: {re_pct:.2f}%<br>
          Self-sufficiency: {self_suf:.1f}%
        </div>
      </div>

      <div class="kpi-card">
        <div class="kpi-icon">🔌</div>
        <div class="kpi-label">Grid Import</div>
        <div class="kpi-value">{grid_import:.1f} kWh</div>
        <div class="kpi-sub">{grid_import/total_load*100 if total_load else 0:.1f}% of load</div>
        <div class="kpi-tooltip">
          <b>Grid Import</b><br>
          Exact: {grid_import:.3f} kWh<br>
          Share of load: {grid_import/total_load*100 if total_load else 0:.2f}%<br>
          Net (import−export): {grid_import - grid_export:+.2f} kWh
        </div>
      </div>

      <div class="kpi-card">
        <div class="kpi-icon">📤</div>
        <div class="kpi-label">Grid Export</div>
        <div class="kpi-value">{grid_export:.1f} kWh</div>
        <div class="kpi-sub">Excess RE to grid</div>
        <div class="kpi-tooltip">
          <b>Grid Export</b><br>
          Exact: {grid_export:.3f} kWh<br>
          Net grid position: {grid_import - grid_export:+.3f} kWh<br>
          {'Net importer ↑' if grid_import > grid_export else 'Net exporter ↓'}
        </div>
      </div>

      <div class="kpi-card">
        <div class="kpi-icon">🔋</div>
        <div class="kpi-label">Batt Net</div>
        <div class="kpi-value" style="color:{'#4CAF50' if batt_net>=0 else '#FF7043'}">{batt_net:+.1f} kWh</div>
        <div class="kpi-sub">{'Net charging ↑' if batt_net>=0 else 'Net discharging ↓'}</div>
        <div class="kpi-tooltip">
          <b>Battery Net Energy</b><br>
          Exact: {batt_net:+.3f} kWh<br>
          + = net charge &nbsp; − = net discharge<br>
          Charge rows: {(df_result['Batt Mode']=='Charge').sum()}<br>
          Discharge rows: {(df_result['Batt Mode']=='Discharge').sum()}
        </div>
      </div>

    </div>
    """
    st.html(kpi_html)

    st.divider()

    # ── Shared time-range slider ───────────────────────────────
    time_vals = df_result["Time"].tolist()
    n_pts     = len(time_vals)

    # Build slider only when there are enough points
    if n_pts > 1:
        sl_col1, sl_col2 = st.columns([1, 3])
        with sl_col1:
            st.markdown("**🕐 Time Window**")
        with sl_col2:
            t_range = st.slider(
                "Time range",
                min_value=0, max_value=n_pts - 1,
                value=(0, n_pts - 1),
                key="time_slider",
                label_visibility="collapsed",
            )
        t_start, t_end = t_range
        # Show human-readable labels
        st.caption(f"Showing **{time_vals[t_start]}** → **{time_vals[t_end]}**  ({t_end - t_start + 1} of {n_pts} points)")
    else:
        t_start, t_end = 0, n_pts - 1

    # Slice data to selected window
    df_slice = df_result.iloc[t_start : t_end + 1]
    x_slice  = df_slice["Time"].tolist()

    st.divider()

    # ── Chart 1 ────────────────────────────────────────────────
    st.subheader("📊 Interactive Energy Chart")
    SIGNALS1 = ["Load (kW)","RE (kW)","Load_out (kW)","Batt_out (kW)","Grid_out (kW)","Battery SOC (%)"]
    COLORS   = ["#2196F3","#4CAF50","#FF9800","#E91E63","#9C27B0","#00BCD4","#F44336","#8BC34A"]

    cc1, cc2 = st.columns(2)
    with cc1:
        pri = st.multiselect("Primary Y-axis (kW)", SIGNALS1,
                             default=["Load (kW)","RE (kW)","Load_out (kW)","Grid_out (kW)"],
                             key="chart1_pri")
    with cc2:
        sec = st.multiselect("Secondary Y-axis", SIGNALS1, default=["Battery SOC (%)"],
                             key="chart1_sec")

    if pri or sec:
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        ci = 0
        for s in pri:
            if s in df_slice.columns:
                fig.add_trace(go.Scatter(x=x_slice, y=df_slice[s].tolist(), name=s,
                    line=dict(color=COLORS[ci % len(COLORS)], width=2), mode="lines"),
                    secondary_y=False); ci += 1
        for s in sec:
            if s in df_slice.columns:
                fig.add_trace(go.Scatter(x=x_slice, y=df_slice[s].tolist(), name=s + " (R)",
                    line=dict(color=COLORS[ci % len(COLORS)], width=2, dash="dot"), mode="lines"),
                    secondary_y=True); ci += 1
        fig.update_layout(height=500, template="plotly_dark", hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(l=40, r=40, t=40, b=60),
            xaxis=dict(title="Time", tickangle=-45, tickmode="auto", nticks=24))
        fig.update_yaxes(title_text="Power (kW)", secondary_y=False)
        fig.update_yaxes(title_text="SOC (%)",    secondary_y=True)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Select at least one signal above to plot.")

    st.divider()

    # ── Chart 2 ────────────────────────────────────────────────
    st.subheader("📊 Interactive Energy Chart 2")
    SIGNALS2 = ["Load (kW)","RE (kW)","Load_out (kW)","Batt_out (kW)","Grid_out (kW)","Grid Available","Tariff"]
    COLORS2  = ["#2196F3","#4CAF50","#FF9800","#E91E63","#9C27B0","#00BCD4","#F44336","#8BC34A","#E91E63"]

    cc1, cc2 = st.columns(2)
    with cc1:
        pri = st.multiselect("Primary Y-axis (kW)", SIGNALS2,
                             default=["Grid Available","Tariff"],
                             key="chart2_pri")
    with cc2:
        sec = st.multiselect("Secondary Y-axis", SIGNALS2,
                             default=["Load (kW)","RE (kW)","Load_out (kW)","Batt_out (kW)","Grid_out (kW)"],
                             key="chart2_sec")

    if pri or sec:
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        ci = 0
        for s in pri:
            if s in df_slice.columns:
                fig.add_trace(go.Scatter(x=x_slice, y=df_slice[s].tolist(), name=s,
                    line=dict(color=COLORS2[ci % len(COLORS2)], width=2), mode="lines"),
                    secondary_y=False); ci += 1
        for s in sec:
            if s in df_slice.columns:
                fig.add_trace(go.Scatter(x=x_slice, y=df_slice[s].tolist(), name=s + " (R)",
                    line=dict(color=COLORS2[ci % len(COLORS2)], width=2, dash="dot"), mode="lines"),
                    secondary_y=True); ci += 1
        fig.update_layout(height=500, template="plotly_dark", hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(l=40, r=40, t=40, b=60),
            xaxis=dict(title="Time", tickangle=-45, tickmode="auto", nticks=24))
        fig.update_yaxes(title_text="0-False/1-True", secondary_y=False)
        fig.update_yaxes(title_text="Power (kW)",     secondary_y=True)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Select at least one signal above to plot.")

    st.divider()
    
    # Output table
    st.subheader("📋 EMS Output Table")
    DCOLS = ["Time","Battery SOC (%)","RE (kW)","Load (kW)","Grid Available","Tariff",
             "Rule No","Load Source","Batt Mode","Grid Mode",
             "Load_out (kW)","Batt_out (kW)","Grid_out (kW)",
             "Grid Import (kW)","Grid Export (kW)","Note"]

    # Row colours keyed on Load Source
    _ROW_COLORS = {"Grid": "#fff0f0", "Battery": "#fff7ec", "RE": "#f0faf2"}

    df_display = df_result[DCOLS].copy()
    for _col in ["Batt_out (kW)", "Grid_out (kW)", "Load_out (kW)",
                  "Grid Import (kW)", "Grid Export (kW)"]:
        df_display[_col] = df_display[_col].round(3)

    # Build HTML table with row-level background colours
    def _build_table_html(df, row_colors):
        col_widths = {
            "Time": "140px", "Battery SOC (%)": "90px", "RE (kW)": "72px",
            "Load (kW)": "72px", "Grid Available": "72px", "Tariff": "60px",
            "Rule No": "60px", "Load Source": "80px", "Batt Mode": "82px",
            "Grid Mode": "82px", "Load_out (kW)": "90px", "Batt_out (kW)": "90px",
            "Grid_out (kW)": "90px", "Grid Import (kW)": "100px",
            "Grid Export (kW)": "100px", "Note": "200px",
        }
        header_cells = "".join(
            f'<th style="width:{col_widths.get(c,"90px")};min-width:{col_widths.get(c,"90px")};'
            f'padding:6px 8px;background:#1e2130;color:#8a9bb0;font-size:11px;'
            f'text-transform:uppercase;letter-spacing:.05em;border-bottom:2px solid #2e3250;'
            f'white-space:nowrap;position:sticky;top:0;z-index:1;">{c}</th>'
            for c in df.columns
        )
        rows_html = ""
        for _, row in df.iterrows():
            bg = row_colors.get(str(row.get("Load Source", "")), "#1a1d2e")
            # darken text slightly for readability on coloured bg
            fg = "#1a1a1a"
            cells = ""
            for c in df.columns:
                val = row[c]
                if c == "Battery SOC (%)":
                    pct = min(100, max(0, float(val)))
                    bar_color = "#4CAF50" if pct >= 70 else "#FF9800" if pct >= 30 else "#F44336"
                    cells += (
                        f'<td style="padding:5px 8px;font-size:12px;color:{fg};">'
                        f'<div style="background:#ddd;border-radius:4px;height:10px;width:100%;margin-bottom:2px;">'
                        f'<div style="background:{bar_color};width:{pct:.0f}%;height:10px;border-radius:4px;"></div></div>'
                        f'<span style="font-size:10px;">{val:.1f}%</span></td>'
                    )
                else:
                    if isinstance(val, float):
                        display = f"{val:+.3f}" if c in ("Batt_out (kW)","Grid_out (kW)") else f"{val:.3f}" if "(kW)" in c else f"{val}"
                    else:
                        display = str(val)
                    cells += f'<td style="padding:5px 8px;font-size:12px;color:{fg};white-space:nowrap;">{display}</td>'
            rows_html += f'<tr style="background:{bg};">{cells}</tr>'

        return f"""
        <div style="overflow:auto;max-height:440px;border:1px solid #2e3250;border-radius:8px;">
          <table style="border-collapse:collapse;width:100%;font-family:Arial,sans-serif;">
            <thead><tr>{header_cells}</tr></thead>
            <tbody>{rows_html}</tbody>
          </table>
        </div>
        """

    st.html(_build_table_html(df_display, _ROW_COLORS))

    # Download results
    st.divider()
    d1,d2 = st.columns(2)
    csv_buf = io.StringIO()
    df_result[DCOLS].to_csv(csv_buf, index=False)
    d1.download_button("⬇️ Download Results CSV", csv_buf.getvalue(),
                       "ems_output.csv","text/csv", use_container_width=True)
    xl_buf = io.BytesIO()
    with pd.ExcelWriter(xl_buf, engine="openpyxl") as w:
        df_result[DCOLS].to_excel(w, index=False, sheet_name="EMS Output")
    d2.download_button("⬇️ Download Results Excel", xl_buf.getvalue(),
                       "ems_output.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)

    # Rule distribution
    st.divider()
    st.subheader("🔢 Rule Usage Distribution")
    rc = df_result["Rule No"].value_counts().reset_index()
    rc.columns = ["Rule No","Count"]
    fig2 = go.Figure(go.Bar(x=rc["Rule No"].astype(str), y=rc["Count"],
        marker_color="#42A5F5", text=rc["Count"], textposition="outside"))
    fig2.update_layout(height=300, template="plotly_dark",
        xaxis_title="Rule Number", yaxis_title="Count",
        margin=dict(l=40,r=20,t=20,b=40))
    st.plotly_chart(fig2, use_container_width=True)

    # Active rules preview
    with st.expander("🔍 View Active Rules Table"):
        df_rv = pd.DataFrame(active_rules).rename(columns={
            "rule":"Rule No","BL":"Batt_L","BM":"Batt_M","BH":"Batt_H",
            "GTH":"Grid_T_H","GTL":"Grid_T_L","Load":"Load Source",
            "Batt":"Batt Mode","Grid_out":"Grid Mode","note":"Note"})
        st.dataframe(df_rv, use_container_width=True, height=300)

else:
    # Welcome screen
    st.info("👈 Use the sidebar to configure inputs, then click **🚀 Generate EMS Output**.")

    c1,c2 = st.columns(2)
    with c1:
        st.markdown("""
### 🗂️ Input Data Columns
| Column | Description |
|---|---|
| `Time` | `HH:MM:SS` **or** `YYYY-MM-DD HH:MM(:SS)` — both formats accepted, full datetime recommended for multi-day / full-year data |
| `Battery SOC (%)` | State of Charge 0–100 |
| `RE (kW)` | Solar + Wind generation |
| `Load (kW)` | Site load demand |
| `Grid Available` | 1 = available, 0 = outage |
| `Tariff` | 1 = Peak (high), 0 = Off-peak |
""")
    with c2:
        st.markdown("""
### 📐 Rules File Columns
| Column | Description |
|---|---|
| `Rule No` | Unique rule number |
| `RE` | 1 = RE≥Load, 0 = RE<Load |
| `Batt_L/M/H` | SOC bands (0–30 / 30–70 / 70–100%) |
| `Grid`, `Grid_T_H`, `Grid_T_L` | Grid & tariff flags |
| `TOD1–TOD4` | Time-of-day bands (6 hr each) |
| `Load`, `Batt`, `Grid_out` | Output actions |
| `Note` | Description |

> **Empty cells = "don't care"** (match any value)
""")
    st.markdown("""
### ⏰ Time of Day Bands
| Band | Window |
|---|---|
| TOD1 | 00:00 – 06:00 |
| TOD2 | 06:00 – 12:00 |
| TOD3 | 12:00 – 18:00 |
| TOD4 | 18:00 – 24:00 |
""")
