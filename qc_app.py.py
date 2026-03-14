# qc_app.py
import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime, date
import io, re, calendar
from typing import Optional, Dict, Any, List, Tuple, Set

st.set_page_config(page_title="E2B_R3 XML Quality Comparator", layout="wide")
st.title("🧪🆚 E2B_R3 XML Quality Comparator")

# ---------------------------- Helpers (borrowed & trimmed from your app) ----------------------------
UNKNOWN_TOKENS = {"unk", "asku", "unknown"}
NS = {'hl7': 'urn:hl7-org:v3', 'xsi': 'http://www.w3.org/2001/XMLSchema-instance'}

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

# Configurable OIDs / Paths
CASE_ID_OIDS = [
    "2.16.840.1.113883.3.989.2.1.3.22",  # your sample case-id system
    "2.16.840.1.113883.3.989.2.1.3.7",   # (example) patient record number in your current app
]
SENDER_ID_OID = "2.16.840.1.113883.3.989.2.1.3.1"

TD_PATHS = [
    './/hl7:transmissionWrapper/hl7:creationTime',
    './/hl7:ControlActProcess/hl7:effectiveTime',
    './/hl7:ClinicalDocument/hl7:effectiveTime',
    './/hl7:creationTime',
]

# ---------------------------- Extraction to a canonical dict ----------------------------
def get_text(elem) -> str:
    return clean_value(elem.text) if (elem is not None and elem.text) else ""

def find_first(root, xpath) -> Optional[ET.Element]:
    return root.find(xpath, NS)

def findall(root, xpath) -> List[ET.Element]:
    return root.findall(xpath, NS)

def extract_case_id(root: ET.Element) -> str:
    # Try configured OIDs first
    for id_elem in findall(root, './/hl7:id'):
        r = id_elem.attrib.get('root')
        if r in CASE_ID_OIDS:
            ext = id_elem.attrib.get('extension') or ""
            nf = id_elem.attrib.get('nullFlavor') or ""
            if nf == 'MSK': return "Masked"
            if ext: return ext.strip()
    # Fallback: first ClinicalDocument/id
    cd_id = find_first(root, './/hl7:ClinicalDocument/hl7:id')
    if cd_id is not None:
        return cd_id.attrib.get('extension') or cd_id.attrib.get('root') or ""
    return ""

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
    # FRD heuristic: earliest of any <low> values (more stable than “last low”)
    lows = []
    for el in root.iter():
        ln = el.tag.split('}')[-1] if '}' in el.tag else el.tag
        if ln == 'low':
            v = el.attrib.get('value')
            if v: lows.append(v)
    if lows:
        # pick earliest by parsed date
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

    # Initials (masked handling)
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
        "gender": gender, "age": clean_value(age),
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
        if drug_id in suspects:  # only suspects
            # name
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

            # dose & text
            txt = get_text(find_first(drug, './/hl7:text'))
            dq = find_first(drug, './/hl7:doseQuantity')
            dose_v = dq.attrib.get('value','') if dq is not None else ''
            dose_u = dq.attrib.get('unit','') if dq is not None else ''

            # dates
            low = find_first(drug, './/hl7:low'); high = find_first(drug, './/hl7:high')
            sd_raw = low.attrib.get('value','') if low is not None else ''
            ed_raw = high.attrib.get('value','') if high is not None else ''

            # form / lot / MAH
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
                "name_norm": normalize_text(raw_name),
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
            # seriousness flags
            ser_map = {
                "resultsInDeath":"Death",
                "isLifeThreatening":"LT",
                "requiresInpatientHospitalization":"Hospital",
                "resultsInPersistentOrSignificantDisability":"Disability",
                "congenitalAnomalyBirthDefect":"Congenital",
                "otherMedicallyImportantCondition":"IME"
            }
            flags = set()
            for k, lbl in ser_map.items():
                crit = find_first(rxn, f'.//hl7:code[@displayName="{k}"]/../hl7:value')
                if crit is not None and crit.attrib.get('value') == 'true':
                    flags.add(lbl)
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
                "key": llt_code or normalize_text(llt_term),  # event matching key
                "llt_code": clean_value(llt_code),
                "llt_term": clean_value(llt_term),
                "serious_flags": sorted(list(flags)),
                "outcome": clean_value(outcome),
                "start_raw": sd_raw, "start": format_date(sd_raw),
                "end_raw": ed_raw, "end": format_date(ed_raw),
            })
    return out

def extract_model(xml_bytes: bytes) -> Dict[str, Any]:
    try:
        root = ET.fromstring(xml_bytes)
    except Exception as e:
        return {"_error": f"XML parse error: {e}"}
    model = {}
    model["case_id"] = extract_case_id(root)
    model["sender_id"] = extract_sender_id(root)
    model.update(extract_td_frd_lrd(root))
    model["patient"] = extract_patient(root)
    model["products"] = extract_products(root)
    model["events"] = extract_events(root)
    return model

# ---------------------------- Comparator ----------------------------
Severity = ["CRITICAL", "HIGH", "MEDIUM", "LOW"]

def add_diff(diffs: List[Dict[str, Any]], case_id: str, path: str, src, proc, sev: str, rule_id: str):
    if str(src) == str(proc): return
    diffs.append({
        "case_id": case_id, "field_path": path,
        "source": "" if src is None else src,
        "processed": "" if proc is None else proc,
        "severity": sev, "rule_id": rule_id
    })

def compare_scalars(case_id: str, s: Dict[str,Any], p: Dict[str,Any]) -> List[Dict[str,Any]]:
    diffs = []
    add_diff(diffs, case_id, "sender_id", s.get("sender_id"), p.get("sender_id"), "MEDIUM", "SENDER_ID")
    # Dates: treat format-only changes as LOW if same day
    def same_day(a_raw, b_raw):
        da, db = parse_date_obj(a_raw or ""), parse_date_obj(b_raw or "")
        return (da is not None and db is not None and da == db)
    for k, rid, sev_on_day, sev_on_str in [
        ("TD_raw","TD", "LOW", "HIGH"),
        ("FRD_raw","FRD", "LOW", "HIGH"),
        ("LRD_raw","LRD", "LOW", "HIGH"),
    ]:
        s_raw, p_raw = s.get(k,""), p.get(k,"")
        if s_raw != p_raw:
            sev = sev_on_day if same_day(s_raw, p_raw) else sev_on_str
            add_diff(diffs, case_id, rid, format_date(s_raw), format_date(p_raw), sev, f"DATE_{rid}")
    # Patient
    sp, pp = s.get("patient",{}), p.get("patient",{})
    for fld, sev, rid in [
        ("gender","MEDIUM","PAT_GENDER"),
        ("age","MEDIUM","PAT_AGE"),
        ("age_group","LOW","PAT_AGEGROUP"),
        ("height","LOW","PAT_HEIGHT"),
        ("weight","LOW","PAT_WEIGHT"),
        ("initials","LOW","PAT_INITIALS"),
    ]:
        add_diff(diffs, case_id, f"patient.{fld}", sp.get(fld,""), pp.get(fld,""), sev, rid)
    return diffs

def list_to_index(products: List[Dict[str,Any]], key: str) -> Dict[str, Dict[str,Any]]:
    out = {}
    for x in products:
        k = x.get(key,"")
        if k:
            out[k] = x
    return out

def compare_products(case_id: str, s: Dict[str,Any], p: Dict[str,Any]) -> List[Dict[str,Any]]:
    diffs = []
    s_idx = list_to_index(s.get("products",[]), "name_norm")
    p_idx = list_to_index(p.get("products",[]), "name_norm")

    # Add/remove
    for k in sorted(set(s_idx) - set(p_idx)):
        add_diff(diffs, case_id, f"products[{k}]", "PRESENT", "MISSING", "HIGH", "PROD_REMOVED")
    for k in sorted(set(p_idx) - set(s_idx)):
        add_diff(diffs, case_id, f"products[{k}]", "MISSING", "PRESENT", "HIGH", "PROD_ADDED")

    # Field compares
    common = sorted(set(s_idx) & set(p_idx))
    fields = [
        ("name","LOW","PROD_NAME"),
        ("dose_text","LOW","PROD_DOSE_TEXT"),
        ("dose_value","MEDIUM","PROD_DOSE_VALUE"),
        ("dose_unit","LOW","PROD_DOSE_UNIT"),
        ("start_raw","HIGH","PROD_START_DATE"),
        ("end_raw","MEDIUM","PROD_END_DATE"),
        ("form","LOW","PROD_FORM"),
        ("lot","LOW","PROD_LOT"),
        ("mah","HIGH","PROD_MAH"),
    ]
    for k in common:
        srec, prec = s_idx[k], p_idx[k]
        for fld, sev, rid in fields:
            sv = srec.get(fld,""); pv = prec.get(fld,"")
            if fld.endswith("_raw"):
                # compare day-equivalence
                sev_eff = "LOW" if parse_date_obj(sv) == parse_date_obj(pv) else sev
                add_diff(diffs, case_id, f"products[{k}].{fld}", format_date(sv), format_date(pv), sev_eff, rid)
            else:
                add_diff(diffs, case_id, f"products[{k}].{fld}", sv, pv, sev, rid)
    return diffs

def compare_events(case_id: str, s: Dict[str,Any], p: Dict[str,Any]) -> List[Dict[str,Any]]:
    diffs = []
    s_idx = list_to_index(s.get("events",[]), "key")
    p_idx = list_to_index(p.get("events",[]), "key")

    for k in sorted(set(s_idx) - set(p_idx)):
        add_diff(diffs, case_id, f"events[{k}]", "PRESENT", "MISSING", "HIGH", "EVT_REMOVED")
    for k in sorted(set(p_idx) - set(s_idx)):
        add_diff(diffs, case_id, f"events[{k}]", "MISSING", "PRESENT", "HIGH", "EVT_ADDED")

    common = sorted(set(s_idx) & set(p_idx))
    for k in common:
        se, pe = s_idx[k], p_idx[k]
        # seriousness flags as sets
        if set(se.get("serious_flags",[])) != set(pe.get("serious_flags",[])):
            add_diff(diffs, case_id, f"events[{k}].serious_flags",
                     ", ".join(se.get("serious_flags",[])),
                     ", ".join(pe.get("serious_flags",[])),
                     "HIGH", "EVT_SERIOUSNESS")
        # outcome
        add_diff(diffs, case_id, f"events[{k}].outcome", se.get("outcome",""), pe.get("outcome",""), "MEDIUM", "EVT_OUTCOME")
        # dates
        for fld, rid, sev in [("start_raw","EVT_START","HIGH"), ("end_raw","EVT_END","MEDIUM")]:
            sv, pv = se.get(fld,""), pe.get(fld,"")
            sev_eff = "LOW" if parse_date_obj(sv) == parse_date_obj(pv) else sev
            add_diff(diffs, case_id, f"events[{k}].{fld}", format_date(sv), format_date(pv), sev_eff, rid)
        # term/code (low severity if only label/case toggles)
        add_diff(diffs, case_id, f"events[{k}].llt_code", se.get("llt_code",""), pe.get("llt_code",""), "MEDIUM", "EVT_LLT_CODE")
        add_diff(diffs, case_id, f"events[{k}].llt_term", se.get("llt_term",""), pe.get("llt_term",""), "LOW", "EVT_LLT_TERM")
    return diffs

def compare_models(src: Dict[str,Any], proc: Dict[str,Any]) -> List[Dict[str,Any]]:
    case_id = src.get("case_id") or proc.get("case_id") or "(unknown)"
    diffs: List[Dict[str,Any]] = []
    # Case id itself mismatch
    if src.get("case_id","") != proc.get("case_id",""):
        add_diff(diffs, case_id, "case_id", src.get("case_id",""), proc.get("case_id",""), "CRITICAL", "CASE_ID")
    diffs += compare_scalars(case_id, src, proc)
    diffs += compare_products(case_id, src, proc)
    diffs += compare_events(case_id, src, proc)
    return diffs

# ---------------------------- UI ----------------------------
tab1, tab2 = st.tabs(["Upload & Pair", "Compare & Export"])

with tab1:
    st.markdown("### 📤 Upload XMLs")
    col1, col2 = st.columns(2)
    with col1:
        src_files = st.file_uploader("Source XML(s)", type=["xml"], accept_multiple_files=True, key="src")
    with col2:
        prc_files = st.file_uploader("Processed XML(s)", type=["xml"], accept_multiple_files=True, key="prc")

    st.markdown("### 🔗 Auto-Pair by Case ID")
    pairs = []
    unmatched_src, unmatched_prc = [], []
    if src_files and prc_files:
        # Extract models and index by case_id
        src_models, prc_models = {}, {}
        with st.spinner("Parsing source XMLs..."):
            for f in src_files:
                m = extract_model(f.read())
                src_models[m.get("case_id") or f.name] = (f.name, m)
        with st.spinner("Parsing processed XMLs..."):
            for f in prc_files:
                m = extract_model(f.read())
                prc_models[m.get("case_id") or f.name] = (f.name, m)

        src_keys, prc_keys = set(src_models.keys()), set(prc_models.keys())
        common = sorted(src_keys & prc_keys)
        only_src = sorted(src_keys - prc_keys)
        only_prc = sorted(prc_keys - src_keys)

        st.success(f"Auto-paired {len(common)} case(s).")
        if only_src: st.warning(f"Unmatched Source: {len(only_src)}")
        if only_prc: st.warning(f"Unmatched Processed: {len(only_prc)}")

        if only_src or only_prc:
            st.write("You can still compare by selecting manual pairs below.")

        # Show simple manual pairing controls (optional)
        manual = st.expander("Manual pairing")
        with manual:
            left = st.multiselect("Pick a Source case", only_src, [])
            right = st.multiselect("Pick a Processed case", only_prc, [])
            manual_pairs = list(zip(left, right))
        # Build final pair list
        pairs = [(k, src_models[k], prc_models[k]) for k in common] + \
                [("manual", src_models[a], prc_models[b]) for a,b in manual_pairs]

        st.session_state["__qc_pairs__"] = pairs

with tab2:
    st.markdown("### 🧮 Compare")
    pairs = st.session_state.get("__qc_pairs__", [])
    if not pairs:
        st.info("Upload and pair files in the first tab.")
    else:
        all_diffs: List[Dict[str,Any]] = []
        extract_rows_src, extract_rows_prc = [], []
        for key, (src_name, src_model), (prc_name, prc_model) in pairs:
            diffs = compare_models(src_model, prc_model)
            # Attach filename & verdict
            verdict = "OK" if not diffs else "Issues"
            for d in diffs:
                d["src_file"] = src_name; d["proc_file"] = prc_name; d["verdict"] = verdict
            all_diffs.extend(diffs)
            # capture extracts (flatten top-level)
            def flat(model: Dict[str,Any], role: str):
                base = {
                    "case_id": model.get("case_id",""),
                    "sender_id": model.get("sender_id",""),
                    "TD": model.get("TD",""), "FRD": model.get("FRD",""), "LRD": model.get("LRD",""),
                    "role": role
                }
                pat = model.get("patient",{})
                base.update({f"pat_{k}": pat.get(k,"") for k in ["gender","age","age_group","height","weight","initials"]})
                return base
            extract_rows_src.append(flat(src_model, "source"))
            extract_rows_prc.append(flat(prc_model, "processed"))

        # Summary bar
        total = len(pairs)
        cases_with_issues = len({d["case_id"] for d in all_diffs})
        colA, colB, colC, colD = st.columns(4)
        colA.metric("Paired cases", total)
        colB.metric("Cases with issues", cases_with_issues)
        sev_counts = {s:0 for s in Severity}
        for d in all_diffs: sev_counts[d["severity"]] = sev_counts.get(d["severity"],0)+1
        colC.metric("High/Critical issues", sev_counts.get("HIGH",0)+sev_counts.get("CRITICAL",0))
        colD.metric("All diffs", len(all_diffs))

        # Case-wise view
        if all_diffs:
            df = pd.DataFrame(all_diffs).sort_values(["case_id","severity","field_path"])
            st.dataframe(df, use_container_width=True, hide_index=True)
        else:
            st.success("No differences detected across paired cases. ✅")

        # Export
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            if all_diffs:
                pd.DataFrame(all_diffs).to_excel(writer, index=False, sheet_name="Diffs")
            pd.DataFrame(extract_rows_src).to_excel(writer, index=False, sheet_name="Extract_Source")
            pd.DataFrame(extract_rows_prc).to_excel(writer, index=False, sheet_name="Extract_Processed")
        st.download_button("⬇️ Download QC Results (Excel)", excel_buffer.getvalue(), "qc_results.xlsx")