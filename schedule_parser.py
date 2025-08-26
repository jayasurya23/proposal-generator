import sys
from datetime import datetime
from tkinter import messagebox

import pandas as pd

# --- MODIFICATION: Added import for openpyxl ---
try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
except ImportError:
    messagebox.showerror("Missing Dependency", "The 'openpyxl' library is required to work with Excel files. Please install it using: pip install openpyxl")
    sys.exit()

from proposal_generator import ProposalItem, ProposalGenerator
# Parsing & Build Rules
# ===================

EXACT_SUBTOTAL_LABELS = {
    "civil engineering", "electrical engineering", "structural engineering", "substation engineering"
}

PHASES = ["30%", "60%", "90%", "IFC"]


def _pairs(ncols: int):
    """(task, price) column pairs on Proposal Page. Includes D/E explicitly."""
    return [(0, 1), (3, 4), (6, 7), (9, 10), (10, 11)]


def _categorize(text: str) -> str:
    t = (text or "").lower().strip()
    if t.startswith("civil"): return "Civil"
    if t.startswith("electrical"): return "Electrical"
    if t.startswith("structural"): return "Structural"
    if t.startswith("substation"): return "Substation"
    # NEW: BESS detection (handles "BESS", "BESS Engineering", "Battery Energy Storage")
    if t.startswith("bess") or "battery energy storage" in t:
        return "BESS"
    return ""



def _infer_phase(text: str):
    t = (text or "").lower()
    if "30% design" in t or "30%" in t: return "30%"
    if "60% design" in t or "60%" in t: return "60%"
    if "90% design" in t or "90%" in t: return "90%"
    if "ifc design" in t or "record drawings" in t or "ifc" in t: return "IFC"
    return None


def load_proposal_page_rows(path: str):
    pp = pd.read_excel(path, sheet_name="Proposal Page", header=None, engine="openpyxl")
    props = []
    ncols = pp.shape[1]

    def normalize(s):
        return (s or "").strip()

    for (txt_col, price_col) in _pairs(ncols):
        current_phase = None
        current_category = None  # tracks the most recent category header

        for i in range(pp.shape[0]):
            txt = pp.iat[i, txt_col] if txt_col < ncols else None
            val = pp.iat[i, price_col] if price_col < ncols else None

            # Phase detection
            if isinstance(txt, str):
                maybe = _infer_phase(txt)
                if maybe:
                    current_phase = maybe

                # Category header detection (now includes Substation & BESS)
                low = txt.lower().strip()
                if low in (
                    "civil engineering",
                    "electrical engineering",
                    "structural engineering",
                    "substation engineering",
                    "bess",
                    "bess engineering",
                    "battery energy storage",
                    "battery energy storage system",
                ):
                    if "bess" in low or "battery energy storage" in low:
                        current_category = "BESS"
                    elif low.startswith("substation"):
                        current_category = "Substation"
                    else:
                        # "Civil Engineering" -> "Civil", etc.
                        current_category = txt.split()[0].capitalize()
                    continue  # header row itself isn't a task

            # Candidate task row with a numeric price
            if isinstance(txt, str) and normalize(txt) and pd.notna(val):
                lower = txt.lower().strip()
                # Skip totals/headers/etc.
                if any(k in lower for k in [
                    "total", "milestone", "summary of services", "engineering proposal", "insurance adder"
                ]):
                    continue
                if lower in EXACT_SUBTOTAL_LABELS:
                    continue

                # price
                try:
                    price = float(val)
                except Exception:
                    continue

                # Category assignment: explicit on the row, else the last seen header
                cat = _categorize(txt) or current_category or ""
                if not cat:
                    continue

                props.append({
                    "category": cat,
                    "task": normalize(txt),
                    "proposal_price": price,
                    "phase": current_phase or _infer_phase(txt) or "30%",
                })
    return props




def _load_detail_map(path: str, sheet: str):
    """Return {Description: {hours, price}}; dedupe by keeping first priced/max price."""
    try:
        df = pd.read_excel(path, sheet_name=sheet, header=None, engine="openpyxl")
    except Exception:
        return {}
    df2 = df.iloc[12:].reset_index(drop=True)  # after header row (index 11)
    DESC, HRS, COST = 2, 11, 12
    out = {}
    for _, row in df2.iterrows():
        desc = row.iloc[DESC] if len(row) > DESC else None
        hours = row.iloc[HRS] if len(row) > HRS else None
        price = row.iloc[COST] if len(row) > COST else None
        if isinstance(desc, str) and desc.strip():
            name = desc.strip()
            h = float(hours) if pd.notna(hours) else 0.0
            p = 0.0
            if pd.notna(price) and str(price).strip().lower() != "not included":
                try:
                    p = float(price)
                except Exception:
                    p = 0.0
            if name not in out:
                out[name] = {"hours": h, "price": p}
            else:
                prev_p = float(out[name].get("price") or 0.0)
                if (prev_p <= 0 and p > 0) or (p > prev_p):
                    out[name] = {"hours": h, "price": p}
    return out

def _load_structural_from_electrical(path: str, sheet: str = "Electrical"):
    """
    Parse the 'Structural Engineering' section that lives inside the Electrical sheet.
    Returns: {task_name: {"hours": float, "price": float}}
    """
    try:
        df = pd.read_excel(path, sheet_name=sheet, header=None, engine="openpyxl")
    except Exception:
        return {}

    out = {}
    DESC, HRS, COST = 2, 11, 12

    # Find the section header (handles misspelling 'Structrural Engineering')
    header_row = None
    for i in range(len(df)):
        row = df.iloc[i]
        if any(isinstance(v, str) and re.search(r"\bstruct\w*\s+engineering\b", v, re.I) for v in row.tolist()):
            header_row = i
            break
    if header_row is None:
        return {}

    # Collect rows until a new major section / stage total
    for i in range(header_row + 1, len(df)):
        row = df.iloc[i]
        desc = row.iloc[DESC] if len(row) > DESC else None
        if isinstance(desc, str) and desc.strip():
            low = desc.lower().strip()
            if "stage total" in low or any(k in low for k in ["substation", "bess", "additional services"]):
                break
            # skip echoed header lines like "Structural Engineering"
            if re.search(r"\bstruct\w*\s+engineering\b", low):
                continue

            hours = row.iloc[HRS] if len(row) > HRS else None
            price = row.iloc[COST] if len(row) > COST else None
            h = float(hours) if pd.notna(hours) else 0.0
            p = 0.0
            if pd.notna(price) and str(price).strip().lower() != "not included":
                try:
                    p = float(price)
                except Exception:
                    p = 0.0
            out[desc.strip()] = {"hours": h, "price": p}

    # Ensure we also capture a standalone "Structural Plan Set" if it appears outside the block
    for _, row in df.iterrows():
        desc = row.iloc[DESC] if len(row) > DESC else None
        if isinstance(desc, str) and "structural plan set" in desc.lower():
            hours = row.iloc[HRS] if len(row) > HRS else None
            price = row.iloc[COST] if len(row) > COST else None
            h = float(hours) if pd.notna(hours) else 0.0
            p = 0.0
            if pd.notna(price) and str(price).strip().lower() != "not included":
                try:
                    p = float(price)
                except Exception:
                    p = 0.0
            out[desc.strip()] = {"hours": h, "price": p}

    return out
def _load_design_phase_rows(path: str, prefix: str, sheet: str = "Electrical"):
    """
    Pull phase-level rows like 'Substation 60% - Design', 'Substation IFC - Design',
    or 'BESS 60% - Design' from the Electrical sheet.
    Returns: {task_name: {"hours": float, "price": float}}
    """
    try:
        df = pd.read_excel(path, sheet_name=sheet, header=None, engine="openpyxl")
    except Exception:
        return {}

    out = {}
    DESC, HRS, COST = 2, 11, 12
    pref = prefix.lower() + " "

    for _, row in df.iterrows():
        desc = row.iloc[DESC] if len(row) > DESC else None
        if isinstance(desc, str) and desc.strip():
            name = desc.strip()
            low = name.lower()
            # e.g., "Substation 60% - Design", "Substation IFC - Design", "BESS 60% - Design"
            if low.startswith(pref) and "- design" in low:
                hours = row.iloc[HRS] if len(row) > HRS else None
                price = row.iloc[COST] if len(row) > COST else None
                h = float(hours) if pd.notna(hours) else 0.0
                p = 0.0
                if pd.notna(price) and str(price).strip().lower() != "not included":
                    try:
                        p = float(price)
                    except Exception:
                        p = 0.0
                out[name] = {"hours": h, "price": p}
    return out
def enrich_with_details(path: str, rows):
    """
    Attach 'hours' and 'detail_price' to each row by looking up the detail sheets.
    - Electrical tasks: from Electrical sheet
    - Civil tasks:      from Civil sheet
    - Structural tasks: from Structural section inside Electrical sheet
    - Substation tasks: primarily from Civil sheet (e.g., 'Substation Pad Design - Civ. ...')
    - BESS tasks:       primarily from Electrical sheet (e.g., 'BESS 60% - Design')
    """
    xl = pd.ExcelFile(path, engine="openpyxl")
    civil_map = _load_detail_map(path, "Civil") if "Civil" in xl.sheet_names else {}
    elec_map  = _load_detail_map(path, "Electrical") if "Electrical" in xl.sheet_names else {}
    structural_from_elec = _load_structural_from_electrical(path)

    # Helper: tolerant key lookup (handles stray spaces, minor punctuation)
    import re
    def _norm(s: str) -> str:
        return re.sub(r"\s+", " ", (s or "").strip().lower())

    # Precompute normalized maps for fallback matching
    civil_norm = { _norm(k): v for k, v in civil_map.items() }
    elec_norm  = { _norm(k): v for k, v in elec_map.items() }
    struc_norm = { _norm(k): v for k, v in structural_from_elec.items() }

    for r in rows:
        cat  = (r.get("category") or "").strip()
        task = (r.get("task") or "").strip()
        key  = _norm(task)

        # Primary source by category
        if cat == "Electrical":
            d = elec_map.get(task) or elec_map.get(task.strip()) or elec_norm.get(key, {})
        elif cat == "Civil":
            d = civil_map.get(task) or civil_map.get(task.strip()) or civil_norm.get(key, {})
        elif cat == "Structural":
            d = (structural_from_elec.get(task) or structural_from_elec.get(task.strip())
                 or struc_norm.get(key, {}))
        elif cat == "Substation":
            # Substation Pad Design rows live on the Civil sheet
            d = (civil_map.get(task) or civil_map.get(task.strip()) or civil_norm.get(key, {})
                 or elec_map.get(task) or elec_map.get(task.strip()) or elec_norm.get(key, {}))
        elif cat == "BESS":
            # BESS 60% - Design lives on the Electrical sheet
            d = (elec_map.get(task) or elec_map.get(task.strip()) or elec_norm.get(key, {})
                 or civil_map.get(task) or civil_map.get(task.strip()) or civil_norm.get(key, {}))
        else:
            d = {}

        r["hours"] = float(d.get("hours")) if d.get("hours") is not None else None
        r["detail_price"] = float(d.get("price")) if d.get("price") is not None else None

    return rows


def extract_project_info(path: str):
    """
    Robustly scan the 'Proposal Page' for:
      - date
      - client
      - project
      - location
      - state
      - size_mw  (numeric if possible; otherwise raw string)
    """
    df = pd.read_excel(path, sheet_name="Proposal Page", header=None, engine="openpyxl")
    info = {"date": None, "client": None, "project": None, "location": None, "state": None, "size_mw": None}

    # Keep the old fixed-cell fallbacks if they still apply
    try:
        info["date"] = df.iat[0, 1]
    except Exception:
        pass

    # Scan top-left area for label:value pairs (avoid picking up "Client Review" task text)
    max_rows = min(40, df.shape[0])
    max_cols = min(12, df.shape[1])

    label_map = {
        "date": ["date", "proposal date"],
        "client": ["client", "client name"],
        "project": ["project", "project name"],
        "location": ["location", "site location", "project location"],
        "state": ["state"],
        "size_mw": [
            "size(mw)", "sizemw", "size mw", "project size (mw)", "project size", "mw",
            "size (mwac)", "size (mw dc)", "size (mwac/mwdc)", "size (mwac/mw dc)"
        ],
    }

    def norm_label(s: str) -> str:
        return re.sub(r"[\s:()/_-]+", "", s.lower())

    variants = {k: {norm_label(v) for v in vals} for k, vals in label_map.items()}
    found = {}

    for r in range(max_rows):
        for c in range(max_cols - 1):
            cell = df.iat[r, c]
            if not isinstance(cell, str):
                continue
            key = norm_label(cell)
            for field, opts in variants.items():
                if key in opts:
                    val = df.iat[r, c + 1] if (c + 1) < df.shape[1] else None
                    if pd.notna(val) and field not in found:
                        found[field] = val

    # Merge discovered values
    for k in ("date", "client", "project", "location", "state"):
        if k in found and (found[k] is not None and str(found[k]).strip()):
            info[k] = found[k]

    # Parse size_mw numerically when possible
    if "size_mw" in found and found["size_mw"] is not None:
        raw = str(found["size_mw"]).strip()
        m = re.search(r"([\d.,]+)", raw)
        if m:
            try:
                info["size_mw"] = float(m.group(1).replace(",", ""))
            except Exception:
                info["size_mw"] = raw  # fallback to raw
        else:
            info["size_mw"] = raw

    return info



def build_model_rows(path: str):
    rows = load_proposal_page_rows(path)
    rows = enrich_with_details(path, rows)

    # Normalize phases
    for r in rows:
        r["phase"] = r.get("phase") or _infer_phase(r["task"]) or "30%"

    # Keep rows with either proposal or detail pricing (> 0)
    filtered = []
    for r in rows:
        pp = r.get("proposal_price") or 0
        dp = r.get("detail_price") or 0
        if (pp > 0) or (dp > 0):
            filtered.append(r)
    rows = filtered

    # Bucket by category/phase (now includes Substation & BESS)
    buckets = {
        "Civil": {p: [] for p in PHASES},
        "Electrical": {p: [] for p in PHASES},
        "Structural": {p: [] for p in PHASES},
        "Substation": {p: [] for p in PHASES},
        "BESS": {p: [] for p in PHASES},
    }
    for r in rows:
        cat, ph = r["category"], r["phase"]
        if cat in buckets and ph in buckets[cat]:
            buckets[cat][ph].append(r)

    info = extract_project_info(path)
    return buckets, info



    # Bucket by category/phase
    buckets = {
        "Civil": {p: [] for p in PHASES},
        "Electrical": {p: [] for p in PHASES},
        "Structural": {p: [] for p in PHASES},
    }
    for r in rows:
        cat, ph = r["category"], r["phase"]
        if cat in buckets and ph in buckets[cat]:
            buckets[cat][ph].append(r)

    info = extract_project_info(path)
    return buckets, info


def flatten_to_template_rows(buckets, hours_per_day: float, price_source: str, review_pairs: set):
    """
    Flatten into the table format ProposalGenerator expects:
      - Project Initiation (with Civil/Electrical Due Diligence 1d)
      - Civil, Electrical, Structural, Substation, BESS (30/60/90/IFC)
      - Prices:
          * "proposal" → Proposal Page price
          * "detail"   → Detail price (fallback to Proposal)
      - Project Closeout (top-level only)
    """
    rows_out = []
    next_id = 1

    # cross-category predecessor trackers
    e60_first_task_id = None
    structural_first_task_applied = False

    def add_item(item_id, name, duration, price, is_milestone, indent, enabled, pred_id, lag, pinned, parent_id):
        rows_out.append({
            "ID": item_id,
            "Name": name,
            "Duration": int(duration or 0),
            "Price": int(price or 0),
            "Is Milestone": bool(is_milestone),
            "Indent Level": int(indent),
            "Enabled": bool(enabled),
            "Predecessor ID": pred_id,
            "Lag": int(lag or 0),
            "Is Start Pinned": bool(pinned),
            "Parent ID": parent_id,
        })

    # ---- Project Initiation ----
    pi_id = next_id; next_id += 1
    add_item(pi_id, "Project Initiation", 0, 0, True, 0, True, None, 0, False, None)

    last_pi_child = None
    civil_dd_id = None
    electrical_dd_id = None
    for name, dur in [
        ("Deposit & Contract Signed", 0),
        ("Notice to Proceed", 0),
        ("Civil Start - Civil Due Diligence", 1),
        ("Electrical Start - Electrical Due Diligence", 1),
    ]:
        tid = next_id; next_id += 1
        add_item(tid, name, dur, 0, False, 1, True, last_pi_child, 0, False, pi_id)
        last_pi_child = tid
        if name.startswith("Civil Start"): civil_dd_id = tid
        if name.startswith("Electrical Start"): electrical_dd_id = tid

    def build_category(cat_key, cat_label):
        nonlocal next_id, e60_first_task_id, structural_first_task_applied
        cat_id = next_id; next_id += 1
        # disabled until we actually add children
        add_item(cat_id, cat_label, 0, 0, True, 0, False, None, 0, False, None)

        added_any = False

        def _price_of(t):
            if cat_key == "Structural":
                if price_source == "detail":
                    return (t.get("detail_price") or t.get("proposal_price") or 0)
                return (t.get("proposal_price") or t.get("detail_price") or 0)
            if price_source == "detail":
                return (t.get("detail_price") or t.get("proposal_price") or 0)
            return (t.get("proposal_price") or t.get("detail_price") or 0)

        def _reorder_for_30(tasks, phase):
            if phase != "30%":
                return list(tasks)
            def score(t):
                return 0 if "plan set" in (t.get("task", "").lower()) else 1
            return sorted(tasks, key=score)

        prev_phase_last_id = None

        for phase in PHASES:
            raw = buckets.get(cat_key, {}).get(phase, [])
            if not raw:
                continue

            added_any = True
            tasks = _reorder_for_30(raw, phase)

            # Phase milestone
            ms_id = next_id; next_id += 1
            add_item(ms_id, f"{cat_label} — {phase} Design", 0, 0, True, 1, True, None, 0, False, cat_id)

            first_task_id = None
            last_task_id = None

            for t in tasks:
                # Duration: use hours if available (now extended to Substation & BESS)
                if t.get("hours") is not None and cat_key in ("Civil", "Electrical", "Structural", "Substation", "BESS"):
                    dur_days = math.ceil((t.get("hours") or 0) / float(hours_per_day))
                else:
                    dur_days = 0

                tid = next_id; next_id += 1
                add_item(
                    tid,
                    t["task"],
                    dur_days,
                    _price_of(t),
                    False,
                    2,              # under phase milestone
                    True,
                    last_task_id,   # sequential within phase
                    0,
                    False,
                    ms_id
                )
                if first_task_id is None:
                    first_task_id = tid
                last_task_id = tid

            # Client Review for the selected pairs (unchanged policy)
            if (cat_key, phase) in review_pairs:
                cr_id = next_id; next_id += 1
                add_item(cr_id, "Client Review", 10, 0, False, 2, True, last_task_id, 0, False, ms_id)
                last_task_id = cr_id

            # Capture first task of Electrical 60% to link Structural’s first task later
            if cat_key == "Electrical" and phase == "60%" and first_task_id is not None and e60_first_task_id is None:
                e60_first_task_id = first_task_id
            elif cat_key == "Electrical" and phase == "90%" and first_task_id is not None and e60_first_task_id is None:
                e60_first_task_id = first_task_id  # fallback if 60% missing

            # Wire the first task's predecessor
            if first_task_id is not None:
                idx = next(i for i in range(len(rows_out)) if rows_out[i]["ID"] == first_task_id)
                if cat_key == "Structural" and (not structural_first_task_applied) and e60_first_task_id:
                    rows_out[idx]["Predecessor ID"] = e60_first_task_id
                    structural_first_task_applied = True
                elif prev_phase_last_id is not None:
                    rows_out[idx]["Predecessor ID"] = prev_phase_last_id
                else:
                    # For first phase, tie to Due Diligence where applicable; otherwise last Project Initiation child
                    if phase == "30%":
                        if cat_key == "Civil" and civil_dd_id:
                            rows_out[idx]["Predecessor ID"] = civil_dd_id
                        elif cat_key == "Electrical" and electrical_dd_id:
                            rows_out[idx]["Predecessor ID"] = electrical_dd_id
                        else:
                            rows_out[idx]["Predecessor ID"] = last_pi_child
                    else:
                        rows_out[idx]["Predecessor ID"] = last_pi_child

            prev_phase_last_id = last_task_id

        # enable the top-level only if we added content
        if added_any:
            for i in range(len(rows_out) - 1, -1, -1):
                if rows_out[i]["ID"] == cat_id:
                    rows_out[i]["Enabled"] = True
                    break

        return cat_id

    # Default review pairs policy remains the same
    review_pairs = {("Civil", "30%"), ("Civil", "60%"), ("Electrical", "30%"), ("Electrical", "60%")}

    # Order: Civil → Electrical → Structural → Substation → BESS
    build_category("Civil", "Civil Engineering")
    build_category("Electrical", "Electrical Engineering")
    build_category("Structural", "Structural Engineering")
    build_category("Substation", "Substation Engineering")  # NEW
    build_category("BESS", "BESS")                          # NEW

    # ---- Project Closeout ----
    closeout_id = next_id; next_id += 1
    add_item(closeout_id, "Project Closeout", 0, 0, True, 0, True, None, 0, False, None)

    return rows_out



    # Default (non-optional) review pairs per spec
    review_pairs = {("Civil", "30%"), ("Civil", "60%"), ("Electrical", "30%"), ("Electrical", "60%")}

    # Order: Civil → Electrical → Structural
    build_category("Civil", "Civil Engineering")
    build_category("Electrical", "Electrical Engineering")
    build_category("Structural", "Structural Engineering")

    # ---- Project Closeout (top-level only, empty, 0d) ----
    closeout_id = next_id; next_id += 1
    add_item(closeout_id, "Project Closeout", 0, 0, True, 0, True, None, 0, False, None)

    return rows_out



def push_into_generator(gen: ProposalGenerator, project_info, rows_out):
    """Replace any existing task tree with the new one and refresh the UI."""
    # Try to clear any existing Treeview if present
    try:
        tree = getattr(gen, "tree", None) or getattr(gen, "treeview", None)
        if tree is not None and hasattr(tree, "get_children"):
            for iid in tree.get_children(""):
                tree.delete(iid)
    except Exception:
        pass

    # Reset internal containers
    gen.template_items = []
    gen.item_id_map = {}
    gen.task_counter = 0

    # Rebuild from rows
    id_to_item = {}

    def mk_item(row):
        return ProposalItem(
            name=row["Name"],
            duration=int(row["Duration"] or 0),
            price=int(row["Price"] or 0),
            is_milestone=bool(row["Is Milestone"]),
            indent_level=int(row["Indent Level"] or 0),
            item_id=int(row["ID"]),
        )

    for r in rows_out:
        id_to_item[r["ID"]] = mk_item(r)

    for r in rows_out:
        item = id_to_item[r["ID"]]
        pid = r["Parent ID"]
        if pid:
            parent = id_to_item.get(pid)
            if parent:
                item.parent = parent
                parent.children.append(item)
        pred = r["Predecessor ID"]
        if pred:
            item.predecessor_id = int(pred)
            item.predecessor_type = 'FS'
            item.lag = int(r.get("Lag") or 0)
        item.enabled.set(bool(r["Enabled"]))

    gen.template_items = [it for it in id_to_item.values() if it.parent is None]
    gen.item_id_map = {it.id: it for it in id_to_item.values()}
    gen.task_counter = max(gen.item_id_map.keys()) if gen.item_id_map else 0

    # Project info
    if project_info.get("project"):
        gen.project_name.set(str(project_info["project"]))
    if project_info.get("client"):
        gen.company_name.set(str(project_info["client"]))
    date_cell = project_info.get("date")
    try:
        if isinstance(date_cell, (pd.Timestamp, datetime)):
            dt = pd.to_datetime(date_cell).to_pydatetime()
            gen.project_start_date.set(dt.strftime("%m/%d/%y"))
        elif isinstance(date_cell, str) and date_cell.strip():
            dt = pd.to_datetime(date_cell)
            gen.project_start_date.set(pd.to_datetime(dt).strftime("%m/%d/%y"))
    except Exception:
        pass

    # If ProposalGenerator exposes a UI refresh method, call it
    for meth in ("rebuild_tree", "refresh_tree", "render_tree", "draw_tree"):
        if hasattr(gen, meth):
            try:
                getattr(gen, meth)()
                break
            except Exception:
                pass


