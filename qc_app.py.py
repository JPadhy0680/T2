# qc_twofile_compare.py
import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime, date
import io, re, calendar
from typing import Optional, Dict, Any, List, Tuple, Set

# ---------------- UI setup ----------------
st.set_page_config(page_title="E2B_R3 Two-File Vertical Comparator", layout="wide")
st.title("🧪📄📄 E2B_R3 Two‑File Vertical Comparator (Vertical Triage View)")

# ---------------- Utilities ----------------
NS = {'hl7': 'urn:hl7-org:v3', 'xsi': 'http://www.w3.org/2001/XMLSchema-instance'}
UNKNOWN_TOKENS = {"unk", "asku", "unknown"}

def _digits_only(s: str) -> str:
    return re.sub(r"\D", "", (s or "").strip())

def format_date(date_str: str) -> str:
    if not date_str: return ""
    digits = _digits_only(date_str)
    try:
        if len(digits) >= 8:
            return datetime.strptime(digits[:8], "%Y%m%d").strftime("%d-%b-%Y")
        elif len(digits) >= 6:
            return datetime.strptime(digits[:6], "%Y%m").strftime("%b-%Y")
        elif len(digits) >= 4:
            return digits[:4]
    except Exception:
        pass
    return ""

def parse_date_obj(date_str: str) -> Optional[date]:
    if not date_str: return None
    digits = _digits_only(date_str)
    try:
        if len(digits) >= 8:
            return datetime.strptime(digits[:8], "%Y%m%d").date()
        elif len(digits) >= 6:
            y, m = int(digits[:4]), int(digits[4:6])
            last = calendar.monthrange(y, m)[1]
            return date(y, m, last)
        elif len(digits) >= 4:
            y = int(digits[:4]); return date(y, 12, 31)
    except Exception:
        pass
    return None

def clean_value(v: Any) -> str:
    if v is None: return ""
    s = str(v).strip()
    return "" if (not s or s.lower() in UNKNOWN_TOKENS) else s

def normalize_text(s: str) -> str:
    s = (s or "").lower()
    s = re.sub(r'[^a-z0-9\s\+\-]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()

def map_gender(code: str) -> str:
    return {"1":"Male", "2":"Female", "M":"Male", "F":"Female"}.get(code, "Unknown")

def get_text(elem) -> str:
    return clean_value(elem.text) if (elem is not None and elem.text) else ""

def find_first(root, xpath) -> Optional[ET.Element]:
    return root.find(xpath, NS)

def findall(root, xpath) -> List[ET.Element]:
    return root.findall(xpath, NS)

# ---------------- Canonical extraction ----------------
SENDER_ID_OID = "2.16.840.1.113883.3.989.2.1.3.1"  # same as in your existing app

TD_PATHS = [
    './/hl7:transmissionWrapper/hl7:creationTime',
    './/hl7:ControlActProcess/hl7:effectiveTime',
    './/hl7:ClinicalDocument/hl7:effectiveTime',
    './/hl7:creationTime',
]

def extract_sender_id(root: ET.Element) -> str:
    e = find_first(root, f'.//hl7:id[@root="{SENDER_ID_OID}"]')
    return clean_value(e.attrib.get('extension', '')) if e is not None else ""

def extract_td_frd_lrd(root: ET.Element) -> Dict[str, str]:
    out = {"TD_raw":"", "TD":"", "FRD_raw":"", "FRD":"", "LRD_raw":"", "LRD":""}
    # TD by priority paths
    for p in TD_PATHS:
        e = find_first(root, p)
        if e is not None:
            val = e.attrib.get('value') or get_text(e)
            if val:
                out["TD_raw"] = val; out["TD"] = format_date(val)
                break
    # LRD from explicit availabilityTime (first encountered)
    for el in root.iter():
        ln = el.tag.split('}')[-1] if '}' in el.tag else el.tag
        if ln == 'availabilityTime':
            v = el.attrib.get('value')
            if v:
                out["LRD_raw"] = v; out["LRD"] = format_date(v); break
    # FRD heuristic: earliest <low> in the document (safer than "last low")
    lows = []
    for el in root.iter():
        ln = el.tag.split('}')[-1] if '}' in el.tag else el.tag
        if ln == 'low':
            v = el.attrib.get('value')
            if v: lows.append(v)
    if lows:
        pairs = [(parse_date_obj(v), v) for v in lows if parse_date_obj(v)]
        if pairs:
            pairs.sort(key=lambda t: t[0])
            out["FRD_raw"] = pairs[0][1]; out["FRD"] = format_date(pairs[0][1])
    return out

def extract_patient(root: ET.Element) -> Dict[str, str]:
    gender_elem = find_first(root, './/hl7:administrativeGenderCode')
    gender_code = gender_elem.attrib.get('code', '') if gender_elem is not None else ''
    gender = clean_value(map_gender(gender_code))

    age_elem = find_first(root, './/hl7:code[@displayName="age"]/../hl7:value')
    age_val = age_elem.attrib.get('value','') if age_elem is not None else ''
    age_unit_raw = age_elem.attrib.get('unit','') if age_elem is not None else ''
    unit = {'a':'year', 'b':'month'}.get(str(age_unit_raw).lower(), age_unit_raw)
    age = f"{clean_value(age_val)}{(' ' + clean_value(unit)) if clean_value(age_val) and clean_value(unit) else ''}".strip()

    age_group_map = {"0":"Foetus","1":"Neonate","2":"Infant","3":"Child","4":"Adolescent","5":"Adult","6":"Elderly"}
    ag_elem = find_first(root, './/hl7:code[@displayName="ageGroup"]/../hl7:value')
    age_group = ""
    if ag_elem is not None:
        c = ag_elem.attrib.get('code','')
        nf = ag_elem.attrib.get('nullFlavor','')
        age_group = age_group_map.get(c, "[Masked/Unknown]" if (c in ["MSK","UNK","ASKU","NI"] or nf in ["MSK","UNK","ASKU","NI"]) else "")

    weight_elem = find_first(root, './/hl7:code[@displayName="bodyWeight"]/../hl7:value')
    weight = ""
    if weight_elem is not None:
        wv = clean_value(weight_elem.attrib.get('value',''))
        wu = clean_value(weight_elem.attrib.get('unit',''))
        weight = f"{wv}{(' ' + wu) if wv and wu else ''}"

    height_elem = find_first(root, './/hl7:code[@displayName="height"]/../hl7:value')
    height = ""
    if height_elem is not None:
        hv = clean_value(height_elem.attrib.get('value',''))
        hu = clean_value(height_elem.attrib.get('unit',''))
        height = f"{hv}{(' ' + hu) if hv and hu else ''}"

    # Initials (with mask handling)
    initials = ""
    nm = find_first(root, './/hl7:player1/hl7:name')
    if nm is not None:
        if nm.attrib.get('nullFlavor') == 'MSK':
            initials = "Masked"
        else:
            parts = []
            for g in nm.findall('hl7:given', NS):
                if g.text and g.text.strip(): parts.append(g.text.strip()[0].upper())
            fam = nm.find('hl7:family', NS)
            if fam is not None and fam.text and fam.text.strip(): parts.append(fam.text.strip()[0].upper())
            initials = "".join(parts) or clean_value(get_text(nm))
    return {
        "gender": clean_value(gender), "age": clean_value(age),
        "age_group": clean_value(age_group), "height": clean_value(height),
        "weight": clean_value(weight), "initials": clean_value(initials),
    }

def extract_suspect_ids(root: ET.Element) -> Set[str]:
    out = set()
    for c in findall(root, './/hl7:causalityAssessment'):
        v = find_first(c, './/hl7:value')
        if v is not None and v.attrib.get('code') == '1':
            sid = find_first(c, './/hl7:subject2/hl7:productUseReference/hl7:id')
            if sid is not None:
                out.add(sid.attrib.get('root',''))
    return out

def extract_products(root: ET.Element) -> List[Dict[str, str]]:
    suspects = extract_suspect_ids(root)
    products = []
    for drug in findall(root, './/hl7:substanceAdministration'):
        id_elem = find_first(drug, './/hl7:id')
        drug_id = id_elem.attrib.get('root','') if id_elem is not None else ''
        if drug_id in suspects:  # only suspects (as per your triage)
            nm = find_first(drug, './/hl7:kindOfProduct/hl7:name')
            raw_name = ""
            if nm is not None:
                raw_name = (nm.text or "").strip() or clean_value(nm.attrib.get('displayName', ''))
                if not raw_name:
                    ot = nm.find('hl7:originalText', NS)
                    raw_name = get_text(ot)
            if not raw_name:
                alt = find_first(drug, './/hl7:manufacturedProduct/hl7:name')
                raw_name = get_text(alt)

            txt = get_text(find_first(drug, './/hl7:text'))
            dq = find_first(drug, './/hl7:doseQuantity')
            dose_v = dq.attrib.get('value','') if dq is not None else ''
            dose_u = dq.attrib.get('unit','') if dq is not None else ''

            low = find_first(drug, './/hl7:low'); high = find_first(drug, './/hl7:high')
            sd_raw = low.attrib.get('value','') if low is not None else ''
            ed_raw = high.attrib.get('value','') if high is not None else ''

            form = get_text(find_first(drug, './/hl7:formCode/hl7:originalText'))
            lot = get_text(find_first(drug, './/hl7:lotNumberText'))

            mah = ""
            for p in [
                './/hl7:playingOrganization/hl7:name',
                './/hl7:manufacturerOrganization/hl7:name',
                './/hl7:asManufacturedProduct/hl7:manufacturerOrganization/hl7:name',
            ]:
                node = find_first(drug, p)
                if node is not None and get_text(node):
                    mah = get_text(node); break

            products.append({
                "name": clean_value(raw_name),
                "dose_text": clean_value(txt),
                "dose_value": clean_value(dose_v),
                "dose_unit": clean_value(dose_u),
                "start_raw": sd_raw, "start": format_date(sd_raw),
                "end_raw": ed_raw, "end": format_date(ed_raw),
                "form": clean_value(form),
                "lot": clean_value(lot),
                "mah": clean_value(mah)
            })
    return products

def extract_events(root: ET.Element) -> List[Dict[str, Any]]:
    out = []
    for rxn in findall(root, './/hl7:observation'):
        code = find_first(rxn, 'hl7:code')
        if code is not None and code.attrib.get('displayName') == 'reaction':
            val = find_first(rxn, 'hl7:value')
            llt_code = val.attrib.get('code','') if val is not None else ''
            llt_term = val.attrib.get('displayName','') if val is not None else ''

            ser_map = {
                "resultsInDeath":"Death",
                "isLifeThreatening":"LT",
                "requiresInpatientHospitalization":"Hospital",
                "resultsInPersistentOrSignificantDisability":"Disability",
                "congenitalAnomalyBirthDefect":"Congenital",
                "otherMedicallyImportantCondition":"IME"
            }
            flags = []
            for k, lbl in ser_map.items():
                crit = find_first(rxn, f'.//hl7:code[@displayName="{k}"]/../hl7:value')
                if crit is not None and crit.attrib.get('value') == 'true':
                    flags.append(lbl)

            outcome_elem = find_first(rxn, './/hl7:code[@displayName="outcome"]/../hl7:value')
            outcome_code = outcome_elem.attrib.get('code','') if outcome_elem is not None else ''
            outcome_map = {
                "1":"Recovered/Resolved","2":"Recovering/Resolving","3":"Not recovered/Ongoing",
                "4":"Recovered with sequelae","5":"Fatal","0":"Unknown"
            }
            outcome = outcome_map.get(outcome_code, "Unknown")

            low = find_first(rxn, './/hl7:effectiveTime/hl7:low')
            high = find_first(rxn, './/hl7:effectiveTime/hl7:high')
            sd_raw = low.attrib.get('value','') if low is not None else ''
            ed_raw = high.attrib.get('value','') if high is not None else ''

            out.append({
                "llt_code": clean_value(llt_code),
                "llt_term": clean_value(llt_term),
                "seriousness": "Non-serious" if not flags else ", ".join(sorted(set(flags))),
                "outcome": clean_value(outcome),
                "start_raw": sd_raw, "start": format_date(sd_raw),
                "end_raw": ed_raw, "end": format_date(ed_raw),
            })
    return out

def extract_narrative(root: ET.Element) -> str:
    narrative_elem = root.find('.//hl7:code[@code="PAT_ADV_EVNT"]/../hl7:text', NS)
    return clean_value(narrative_elem.text if narrative_elem is not None else '')

def extract_model(xml_bytes: bytes) -> Dict[str, Any]:
    try:
        root = ET.fromstring(xml_bytes)
    except Exception as e:
        return {"_error": f"XML parse error: {e}"}
    model: Dict[str, Any] = {}
    model["sender_id"] = extract_sender_id(root)
    model.update(extract_td_frd_lrd(root))          # TD/FRD/LRD
    model["patient"] = extract_patient(root)        # patient block
    model["products"] = extract_products(root)      # suspects only
    model["events"] = extract_events(root)          # reaction observations
    model["narrative"] = extract_narrative(root)    # clinical narrative (optional)
    return model

# ---------------- Diff helpers (for icons, Excel, and small severity tags) ----------------
def same_day(a_raw: str, b_raw: str) -> bool:
    da, db = parse_date_obj(a_raw or ""), parse_date_obj(b_raw or "")
    return (da is not None and db is not None and da == db)

def differ(a: Any, b: Any) -> bool:
    return (str(a) or "") != (str(b) or "")

def mismatch_icon(a: Any, b: Any) -> str:
    return " 🔴" if differ(a, b) else ""

def safe_display(v: str) -> str:
    return v if v else "—"

def add_diff_row(diff_rows: List[Dict[str, str]], path: str, left: str, right: str, severity: str):
    if differ(left, right):
        diff_rows.append({"field_path": path, "source": left, "processed": right, "severity": severity})

# ---------------- UI: Upload & Parse ----------------
st.markdown("### 📤 Upload the two XML files you want to compare (no case ID pairing used)")
c1, c2 = st.columns(2)
with c1:
    src_file = st.file_uploader("Source XML", type=["xml"], key="src_xml")
with c2:
    prc_file = st.file_uploader("Processed XML", type=["xml"], key="prc_xml")

if not (src_file and prc_file):
    st.info("Please upload **both** Source and Processed XML files to view the vertical triage comparison.")
    st.stop()

# Parse both files
with st.spinner("Parsing Source..."):
    src_model = extract_model(src_file.read())
with st.spinner("Parsing Processed..."):
    prc_model = extract_model(prc_file.read())

if src_model.get("_error") or prc_model.get("_error"):
    st.error(f"Source error: {src_model.get('_error','-')}\nProcessed error: {prc_model.get('_error','-')}")
    st.stop()

# ---------------- Build Diffs (for export) ----------------
diff_rows: List[Dict[str, str]] = []

# Header diffs
add_diff_row(diff_rows, "header.sender_id", src_model.get("sender_id",""), prc_model.get("sender_id",""), "MEDIUM")
for k, rid, sev_day, sev_fmt in [
    ("TD_raw","header.TD", "LOW", "HIGH"),
    ("FRD_raw","header.FRD", "LOW", "HIGH"),
    ("LRD_raw","header.LRD", "LOW", "HIGH"),
]:
    s_raw, p_raw = src_model.get(k,""), prc_model.get(k,"")
    if differ(s_raw, p_raw):
        sev = sev_day if same_day(s_raw, p_raw) else sev_fmt
        add_diff_row(diff_rows, rid, format_date(s_raw), format_date(p_raw), sev)

# Patient diffs
for fld, sev in [
    ("gender","MEDIUM"), ("age","MEDIUM"), ("age_group","LOW"),
    ("height","LOW"), ("weight","LOW"), ("initials","LOW")
]:
    add_diff_row(diff_rows, f"patient.{fld}", src_model["patient"].get(fld,""), prc_model["patient"].get(fld,""), sev)

# Product diffs (row-by-row alignment by index, since we are not case-ID matching; show counts)
max_prod = max(len(src_model["products"]), len(prc_model["products"]))
for i in range(max_prod):
    srec = src_model["products"][i] if i < len(src_model["products"]) else {}
    prec = prc_model["products"][i] if i < len(prc_model["products"]) else {}
    for fld, sev, fmt in [
        ("name","LOW", None),
        ("dose_text","LOW", None),
        ("dose_value","MEDIUM", None),
        ("dose_unit","LOW", None),
        ("start_raw","HIGH", "date"),
        ("end_raw","MEDIUM", "date"),
        ("form","LOW", None),
        ("lot","LOW", None),
        ("mah","HIGH", None),
    ]:
        sv = srec.get(fld,"")
        pv = prec.get(fld,"")
        if fmt == "date":
            # compare display dates (day equivalence lowers severity)
            sev_eff = "LOW" if parse_date_obj(sv) == parse_date_obj(pv) else sev
            add_diff_row(diff_rows, f"products[{i+1}].{fld.replace('_raw','')}", format_date(sv), format_date(pv), sev_eff)
        else:
            add_diff_row(diff_rows, f"products[{i+1}].{fld}", sv, pv, sev)

# Event diffs (row-by-row alignment by index)
max_evt = max(len(src_model["events"]), len(prc_model["events"]))
for i in range(max_evt):
    se = src_model["events"][i] if i < len(src_model["events"]) else {}
    pe = prc_model["events"][i] if i < len(prc_model["events"]) else {}
    # Basic fields
    for fld, sev in [
        ("llt_code","MEDIUM"),
        ("llt_term","LOW"),
        ("seriousness","HIGH"),
        ("outcome","MEDIUM")
    ]:
        add_diff_row(diff_rows, f"events[{i+1}].{fld}", se.get(fld,""), pe.get(fld,""), sev)
    # Dates
    for fld, rid, sev in [("start_raw","start","HIGH"), ("end_raw","end","MEDIUM")]:
        sv, pv = se.get(fld,""), pe.get(fld,"")
        sev_eff = "LOW" if parse_date_obj(sv) == parse_date_obj(pv) else sev
        add_diff_row(diff_rows, f"events[{i+1}].{rid}", format_date(sv), format_date(pv), sev_eff)

# ---------------- Side-by-side vertical rendering ----------------
st.markdown("### 👀 Side‑by‑side vertical triage view (left = Source, right = Processed)")
L, R = st.columns(2)

def section_header(col, title: str):
    col.markdown(f"#### {title}")

def show_kv(col, label: str, left_v: str, right_v: str, right_side: bool = False):
    """Render a key/value with mismatch hint. When right_side=True, compare right vs left to mark on the right."""
    if right_side:
        icon = mismatch_icon(left_v, right_v)
        col.markdown(f"**{label}:** {safe_display(right_v)}{icon}")
    else:
        icon = ""  # only annotate right column
        col.markdown(f"**{label}:** {safe_display(left_v)}")

# Header
section_header(L, "Header"); section_header(R, "Header")
show_kv(L, "Sender ID", src_model["sender_id"], prc_model["sender_id"])
show_kv(R, "Sender ID", src_model["sender_id"], prc_model["sender_id"], right_side=True)

for lab, raw_key in [("TD", "TD_raw"), ("FRD", "FRD_raw"), ("LRD", "LRD_raw")]:
    s_disp, p_disp = format_date(src_model.get(raw_key,"")), format_date(prc_model.get(raw_key,""))
    show_kv(L, lab, s_disp, p_disp)
    show_kv(R, lab, s_disp, p_disp, right_side=True)

# Patient
section_header(L, "Patient"); section_header(R, "Patient")
for fld, lab in [
    ("gender","Gender"), ("age","Age"), ("age_group","Age Group"),
    ("height","Height"), ("weight","Weight"), ("initials","Initials")
]:
    s_v = src_model["patient"].get(fld,"")
    p_v = prc_model["patient"].get(fld,"")
    show_kv(L, lab, s_v, p_v)
    show_kv(R, lab, s_v, p_v, right_side=True)

# Products
section_header(L, "Products (suspects)"); section_header(R, "Products (suspects)")
max_prod_rows = max(len(src_model["products"]), len(prc_model["products"]))
if max_prod_rows == 0:
    L.write("_No suspect products found_"); R.write("_No suspect products found_")
for i in range(max_prod_rows):
    srec = src_model["products"][i] if i < len(src_model["products"]) else {}
    prec = prc_model["products"][i] if i < len(prc_model["products"]) else {}
    L.markdown(f"**{i+1})**")
    R.markdown(f"**{i+1})**")
    for fld, lab, fmt in [
        ("name","Drug","text"),
        ("dose_text","Dosage","text"),
        ("dose_value","Dose Value","text"),
        ("dose_unit","Dose Unit","text"),
        ("start_raw","Start Date","date"),
        ("end_raw","Stop Date","date"),
        ("form","Formulation","text"),
        ("lot","Lot No","text"),
        ("mah","MAH","text"),
    ]:
        sv = srec.get(fld,"")
        pv = prec.get(fld,"")
        s_disp = format_date(sv) if fmt == "date" else sv
        p_disp = format_date(pv) if fmt == "date" else pv
        show_kv(L, lab, s_disp, p_disp)
        show_kv(R, lab, s_disp, p_disp, right_side=True)

# Events
section_header(L, "Events"); section_header(R, "Events")
max_evt_rows = max(len(src_model["events"]), len(prc_model["events"]))
if max_evt_rows == 0:
    L.write("_No events found_"); R.write("_No events found_")
for i in range(max_evt_rows):
    se = src_model["events"][i] if i < len(src_model["events"]) else {}
    pe = prc_model["events"][i] if i < len(prc_model["events"]) else {}
    L.markdown(f"**Event {i+1}**")
    R.markdown(f"**Event {i+1}**")
    for fld, lab, fmt in [
        ("llt_code","LLT Code","text"),
        ("llt_term","LLT Term","text"),
        ("seriousness","Seriousness","text"),
        ("outcome","Outcome","text"),
        ("start_raw","Event Start","date"),
        ("end_raw","Event End","date"),
    ]:
        sv = se.get(fld,"")
        pv = pe.get(fld,"")
        s_disp = format_date(sv) if fmt == "date" else sv
        p_disp = format_date(pv) if fmt == "date" else pv
        show_kv(L, lab, s_disp, p_disp)
        show_kv(R, lab, s_disp, p_disp, right_side=True)

# Narrative (toggle full view)
section_header(L, "Narrative"); section_header(R, "Narrative")
show_full_narrative = st.checkbox("Show full narrative (may be long)", value=True)
max_len = None if show_full_narrative else 1000
src_narr = src_model.get("narrative","")[:max_len] if max_len else src_model.get("narrative","")
prc_narr = prc_model.get("narrative","")[:max_len] if max_len else prc_model.get("narrative","")

L.code(src_narr or "—")
R.code(prc_narr or "—" + (" 🔴" if differ(src_narr, prc_narr) else ""))

# ---------------- Excel export ----------------
st.markdown("---")
st.markdown("### ⬇️ Download QC Results (Excel)")

# Side-by-side flatten for export
side_rows = []

def side_add(path: str, s_val: str, p_val: str):
    side_rows.append({"field_path": path, "source": s_val, "processed": p_val})

# Header
side_add("header.sender_id", src_model["sender_id"], prc_model["sender_id"])
for lab, raw_key in [("TD","TD_raw"), ("FRD","FRD_raw"), ("LRD","LRD_raw")]:
    side_add(f"header.{lab}", format_date(src_model.get(raw_key,"")), format_date(prc_model.get(raw_key,"")))

# Patient
for fld in ["gender","age","age_group","height","weight","initials"]:
    side_add(f"patient.{fld}", src_model["patient"].get(fld,""), prc_model["patient"].get(fld,""))

# Products
for i in range(max_prod_rows):
    srec = src_model["products"][i] if i < len(src_model["products"]) else {}
    prec = prc_model["products"][i] if i < len(prc_model["products"]) else {}
    for fld, lab, fmt in [
        ("name","Drug","text"), ("dose_text","Dosage","text"),
        ("dose_value","Dose Value","text"), ("dose_unit","Dose Unit","text"),
        ("start_raw","Start Date","date"), ("end_raw","Stop Date","date"),
        ("form","Formulation","text"), ("lot","Lot No","text"), ("mah","MAH","text"),
    ]:
        s_disp = format_date(srec.get(fld,"")) if fmt == "date" else srec.get(fld,"")
        p_disp = format_date(prec.get(fld,"")) if fmt == "date" else prec.get(fld,"")
        side_add(f"products[{i+1}].{lab}", s_disp, p_disp)

# Events
for i in range(max_evt_rows):
    se = src_model["events"][i] if i < len(src_model["events"]) else {}
    pe = prc_model["events"][i] if i < len(prc_model["events"]) else {}
    for fld, lab, fmt in [
        ("llt_code","LLT Code","text"), ("llt_term","LLT Term","text"),
        ("seriousness","Seriousness","text"), ("outcome","Outcome","text"),
        ("start_raw","Event Start","date"), ("end_raw","Event End","date"),
    ]:
        s_disp = format_date(se.get(fld,"")) if fmt == "date" else se.get(fld,"")
        p_disp = format_date(pe.get(fld,"")) if fmt == "date" else pe.get(fld,"")
        side_add(f"events[{i+1}].{lab}", s_disp, p_disp)

# Extract overview for audit
def flat_extract(model: Dict[str,Any], role: str):
    base = {
        "role": role,
        "sender_id": model.get("sender_id",""),
        "TD": format_date(model.get("TD_raw","")),
        "FRD": format_date(model.get("FRD_raw","")),
        "LRD": format_date(model.get("LRD_raw","")),
        "pat_gender": model["patient"].get("gender",""),
        "pat_age": model["patient"].get("age",""),
        "pat_age_group": model["patient"].get("age_group",""),
        "pat_height": model["patient"].get("height",""),
        "pat_weight": model["patient"].get("weight",""),
        "pat_initials": model["patient"].get("initials",""),
    }
    return base

excel_buffer = io.BytesIO()
with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
    pd.DataFrame(side_rows).to_excel(writer, index=False, sheet_name="SideBySide")
    if diff_rows:
        pd.DataFrame(diff_rows).to_excel(writer, index=False, sheet_name="Diffs")
    pd.DataFrame([flat_extract(src_model, "source")]).to_excel(writer, index=False, sheet_name="Extract_Source")
    pd.DataFrame([flat_extract(prc_model, "processed")]).to_excel(writer, index=False, sheet_name="Extract_Processed")

st.download_button("Download qc_twofile_compare.xlsx", excel_buffer.getvalue(), "qc_twofile_compare.xlsx")
