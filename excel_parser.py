import math
from datetime import datetime

import pandas as pd

from models import ProposalItem

# Parsing & Build Rules

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
        current_category = None  # NEW

        for i in range(pp.shape[0]):
            txt = pp.iat[i, txt_col] if txt_col < ncols else None
            val = pp.iat[i, price_col] if price_col < ncols else None

            # Phase detection
            if isinstance(txt, str):
                maybe = _infer_phase(txt)
                if maybe:
                    current_phase = maybe

                # Category header detection (NEW)
                low = txt.lower().strip()
                if low in ("civil engineering", "electrical engineering", "structural engineering", "substation engineering"):
                    current_category = txt.split()[0].capitalize()  # "Civil", "Electrical", etc.
                    continue  # category header row itself isn't a task

            # Candidate task
            if isinstance(txt, str) and normalize(txt) and pd.notna(val):
                lower = txt.lower().strip()
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

                # Category assignment:
                cat = _categorize(txt) or current_category or ""
                if not cat:
                    # If still unknown, skip as before
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


def _load_structural_from_electrical(path: str):
    """Structural hours/prices live under 'Structural Engineering' section in Electrical sheet."""
    try:
        df = pd.read_excel(path, sheet_name="Electrical", header=None, engine="openpyxl")
    except Exception:
        return {}
    DESC, HRS, COST = 2, 11, 12

    start_idx = None
    for i in range(len(df)):
        if any(isinstance(v, str) and "structural engineering" in v.lower() for v in df.iloc[i].tolist()):
            start_idx = i + 1
            break
    if start_idx is None:
        return {}

    out = {}
    for _, row in df.iloc[start_idx:].iterrows():
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
            out[name] = {"hours": h, "price": p}
    return out


def enrich_with_details(path: str, rows):
    xl = pd.ExcelFile(path, engine="openpyxl")
    civil_map = _load_detail_map(path, "Civil") if "Civil" in xl.sheet_names else {}
    elec_map = _load_detail_map(path, "Electrical") if "Electrical" in xl.sheet_names else {}
    structural_from_elec = _load_structural_from_electrical(path)

    for r in rows:
        if r["category"] == "Electrical":
            d = elec_map.get(r["task"], {})
        elif r["category"] == "Civil":
            d = civil_map.get(r["task"], {})
        elif r["category"] == "Structural":
            d = structural_from_elec.get(r["task"], {})
        else:
            d = {}
        r["hours"] = float(d.get("hours")) if d.get("hours") is not None else None
        r["detail_price"] = float(d.get("price")) if d.get("price") is not None else None
    return rows


def extract_project_info(path: str):
    df = pd.read_excel(path, sheet_name="Proposal Page", header=None, engine="openpyxl")
    info = {"date": None, "client": None, "project": None, "location": None}
    try:
        info["date"] = df.iat[0, 1]
        info["client"] = df.iat[1, 1] if isinstance(df.iat[1, 0], str) and "client" in str(df.iat[1, 0]).lower() else None
        info["project"] = df.iat[2, 1] if isinstance(df.iat[2, 0], str) and "project" in str(df.iat[2, 0]).lower() else None
        info["location"] = df.iat[0, 4] if isinstance(df.iat[0, 3], str) and "location" in str(df.iat[0, 3]).lower() else None
    except Exception:
        pass
    return info


def build_model_rows(path: str):
    rows = load_proposal_page_rows(path)
    rows = enrich_with_details(path, rows)

    # Normalize phases
    for r in rows:
        r["phase"] = r.get("phase") or _infer_phase(r["task"]) or "30%"

    def is_electrical_study(name: str) -> bool:
        n = (name or "").lower()
        return "study" in n and (n.startswith("electrical") or "electrical study" in n)

    # Inclusion filter
    # Inclusion filter (RELAXED):
    filtered = []
    for r in rows:
        cat = r["category"]
        pp = r.get("proposal_price") or 0
        dp = r.get("detail_price") or 0

        keep = False
        # If either proposal or detail price is present, keep it.
        if pp > 0 or dp > 0:
            keep = True

        # (Optional) for Electrical "study" heuristic can still be kept, but is redundant now
        if keep:
            filtered.append(r)
    rows = filtered

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
    Flatten into the table format ProposalGenerator expects (via ProposalItem build below).
    - Project Initiation (w/ Civil/Electrical Due Diligence 1d)
    - Civil, Electrical, Structural (30/60/90/IFC) with rules:
        * 30% & 60% phases get Client Review (always on)
        * First task of 30% depends on matching Due Diligence (Civil/Electrical)
        * Durations from hours (Civil/Electrical), Structural 0d
    - Prices:
        * "proposal" → Proposal Page price
        * "detail"   → Detail price (fallback to Proposal)
    - Project Closeout top-level milestone only (empty, 0d)
    """
    rows_out = []
    next_id = 1

    # NEW: cross-category predecessor trackers
    e60_first_task_id = None               # first task ID of Electrical 60%
    structural_first_task_applied = False  # ensure we only set Structural's first task once

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
        # Create the top-level milestone for the category, disabled by default.
        cat_id = next_id; next_id += 1
        add_item(cat_id, cat_label, 0, 0, True, 0, False, None, 0, False, None)

        added_any = False  # only enable the top-level if we actually add tasks/milestones

        def _price_of(t):
            # Price selection policy by source, with Structural special-casing preserved.
            if cat_key == "Structural":
                if price_source == "detail":
                    return (t.get("detail_price") or t.get("proposal_price") or 0)
                return (t.get("proposal_price") or t.get("detail_price") or 0)
            if price_source == "detail":
                return (t.get("detail_price") or t.get("proposal_price") or 0)
            return (t.get("proposal_price") or t.get("detail_price") or 0)

        def _reorder_for_30(tasks, phase):
            # Nudge the "plan set" to the top at 30% like your previous heuristic.
            if phase != "30%":
                return list(tasks)
            def score(t):
                return 0 if "plan set" in (t.get("task", "").lower()) else 1
            return sorted(tasks, key=score)

        def _client_review_needed(cat_key, phase):
            return (cat_key, phase) in review_pairs

        for phase in PHASES:
            tasks = list(buckets.get(cat_key, {}).get(phase, []))
            if not tasks:
                continue

            tasks = _reorder_for_30(tasks, phase)
            prev_phase_last_id = None
            last_task_id = None
            added_phase_any = False
            for t in tasks:
                dur_days = math.ceil((t.get("hours") or 0) / float(hours_per_day))
                price = _price_of(t)
                tid = next_id; next_id += 1
                add_item(
                    tid,
                    t.get("task"),
                    0 if t.get("is_milestone") else dur_days,
                    price,
                    bool(t.get("is_milestone")),
                    1,
                    True,
                    prev_phase_last_id,
                    0,
                    False,
                    cat_id,
                )
                last_task_id = tid
                added_any = True
                added_phase_any = True

            if added_phase_any and _client_review_needed(cat_key, phase):
                review_id = next_id; next_id += 1
                add_item(
                    review_id,
                    f"{cat_key} {phase} Review",
                    0,
                    0,
                    True,
                    1,
                    True,
                    last_task_id,
                    0,
                    False,
                    cat_id,
                )
                last_task_id = review_id

            if added_phase_any:
                if cat_key == "Electrical" and phase == "60%" and e60_first_task_id is None:
                    e60_first_task_id = last_task_id
                if cat_key == "Structural" and (not structural_first_task_applied) and e60_first_task_id:
                    rows_out[-1]["Predecessor ID"] = e60_first_task_id
                    structural_first_task_applied = True
                elif prev_phase_last_id is not None:
                    rows_out[-1]["Predecessor ID"] = prev_phase_last_id
                else:
                    if phase == "30%":
                        if cat_key == "Civil" and civil_dd_id:
                            rows_out[-1]["Predecessor ID"] = civil_dd_id
                        elif cat_key == "Electrical" and electrical_dd_id:
                            rows_out[-1]["Predecessor ID"] = electrical_dd_id
                        else:
                            rows_out[-1]["Predecessor ID"] = last_pi_child
                    else:
                        rows_out[-1]["Predecessor ID"] = last_pi_child
                prev_phase_last_id = last_task_id

        if added_any:
            for i in range(len(rows_out) - 1, -1, -1):
                if rows_out[i]["ID"] == cat_id:
                    rows_out[i]["Enabled"] = True
                    break

        return cat_id

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


def push_into_generator(gen, project_info, rows_out):
    """Replace any existing task tree with the new one and refresh the UI."""
    try:
        tree = getattr(gen, "tree", None) or getattr(gen, "treeview", None)
        if tree is not None and hasattr(tree, "get_children"):
            for iid in tree.get_children(""):
                tree.delete(iid)
    except Exception:
        pass

    gen.template_items = []
    gen.item_id_map = {}
    gen.task_counter = 0

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

    for meth in ("rebuild_tree", "refresh_tree", "render_tree", "draw_tree"):
        if hasattr(gen, meth):
            try:
                getattr(gen, meth)()
                break
            except Exception:
                pass
