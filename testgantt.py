import argparse
import re
import textwrap
from datetime import datetime, timedelta
import xml.etree.ElementTree as ET
import os
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as mticker
from matplotlib.gridspec import GridSpec
from matplotlib.patches import Rectangle
from matplotlib.dates import MonthLocator, DateFormatter, WeekdayLocator, MO
from matplotlib.backends.backend_pdf import PdfPages
# Font Sizes
FONTSIZE_TABLE = 8
FONTSIZE_CHART_REGULAR = 7.5
FONTSIZE_CHART_SUMMARY = 7.5 # Used for summary tasks and milestones
FONTSIZE_XTICK = 7.0
MAX_ROWS_PER_PAGE = 35  # Max tasks per page before paginating

PRIMARY = "#991f2b"
SECONDARY = "black"

# Layout (adjusted to give the chart most of the width)
# New Lines
LEFT_RIGHT_WIDTHS = [1.6, 2.4]  # Relative widths of left (table) and right (chart) panels
COL_EDGES = [0.00, 0.65, 0.75, 0.87, 1.00]
HEADERS = ["Task", "Duration", "Start", "Finish"]
HEADER_H = 0.6
MAX_NAME_LENGTH = 50  # Increased to allow longer names in table
ROW_HEIGHT = 1.0  # Standard row height for perfect alignment
MIN_WIDTH_DAYS = 0.2  # visual width for 0-day tasks (milestones)

def _nsmap(root):
    """Detect default namespace from the root tag and return a dict for XPath."""
    m = re.match(r"\{(.*)\}", root.tag)
    uri = m.group(1) if m else ""
    return {"msp": uri} if uri else {}

def _txt(el, tag, ns):
    """Extract text from element, handling both namespaced and non-namespaced tags."""
    # Try with namespace first
    if ns and "msp" in ns:
        t = el.find(f"msp:{tag}", ns)
        if t is not None and t.text:
            return t.text.strip()
    
    # Try without namespace
    t = el.find(tag)
    if t is not None and t.text:
        return t.text.strip()
    
    return None

def _parse_dt(s):
    """Parse ISO-like datetime without external deps."""
    if not s:
        return None
    # Handles strings like "2025-08-01T08:00:00" (ignore any trailing Z)
    try:
        return datetime.fromisoformat(s.replace("Z", ""))
    except ValueError:
        # If that fails, try other common formats
        try:
            return datetime.strptime(s, "%Y-%m-%dT%H:%M:%S")
        except ValueError:
            return None

def parse_msproject_xml(xml_path):
    """
    Return a list of rows in file order:
      dict(name, start: datetime|None, finish: datetime|None, kind: 'task'|'summary')
    Skips IsNull=1 rows. Preserves original order.
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()
    ns = _nsmap(root)
    rows = []
    
    # Find all tasks - try both with and without namespace
    tasks = []
    if ns and "msp" in ns:
        tasks = root.findall(".//msp:Task", ns)
    if not tasks:
        tasks = root.findall(".//Task")
    
    for task in tasks:
        is_null = _txt(task, "IsNull", ns)
        if is_null == "1":
            continue
            
        # Try multiple possible name tags
        name = _txt(task, "Name", ns) or _txt(task, "n", ns) or ""
        
        # Handle HTML entities in names
        name = name.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
        
        summary = (_txt(task, "Summary", ns) == "1")
        start_s = _txt(task, "Start", ns)
        finish_s = _txt(task, "Finish", ns)
        
        start = _parse_dt(start_s) if start_s else None
        finish = _parse_dt(finish_s) if finish_s else None
        
        # Some XMLs include project-level summary; we still show as a summary row.
        kind = "summary" if summary else "task"
        
        rows.append({
            "name": name,
            "start": start,
            "finish": finish,
            "kind": kind,
        })
    
    return rows

def compute_duration_days(start, finish):
    """Integer calendar-day span; same-day shows as 0."""
    if not start or not finish:
        return None
    return max(0, (finish.date() - start.date()).days)

def estimate_text_width_in_days(text, fontsize, chart_span_days, chart_width_inches):
    """
    Estimate the width of text in chart coordinate system (days).
    Balanced estimation for proper gaps without overlap.
    """
    # Balanced estimate: each character is about 0.6 * fontsize points wide
    char_width_points = fontsize * 0.6
    char_width_inches = char_width_points / 72
    text_width_inches = len(text) * char_width_inches
    
    # Convert inches to days based on chart proportions
    days_per_inch = chart_span_days / chart_width_inches
    text_width_days = text_width_inches * days_per_inch
    
    # Keep 10% safety margin for readability
    return text_width_days * 1.1
    # Define how many data rows fit on one page
    ROWS_PER_PAGE = 30

    # Filter dates from all rows to find the global bounds for a consistent X-axis
    starts = [r["start"] for r in rows if r["start"]]
    finishes = [r["finish"] for r in rows if r["finish"]]
    if not starts or not finishes:
        raise ValueError("No dated tasks found.")
    true_start = min(starts)
    true_finish = max(finishes)

    # Calculate total project duration in months to adjust timeline
    duration_months = (true_finish.year - true_start.year) * 12 + true_finish.month - true_start.month

    # Calculate consistent X-axis range for all pages
    left_pad_days = 1
    span_days = max(1, (true_finish.date() - true_start.date()).days)
    right_pad_days = max(1, int(span_days * 0.05))
    x_min = true_start - timedelta(days=left_pad_days)
    x_max = true_finish + timedelta(days=right_pad_days)

    # Precompute durations for all rows
    for r in rows:
        r["dur"] = compute_duration_days(r["start"], r["finish"])

    # Split the data rows into chunks for each page
    row_chunks = [rows[i:i + ROWS_PER_PAGE] for i in range(0, len(rows), ROWS_PER_PAGE)]

    with PdfPages(out_pdf) as pdf:
        # Loop through each chunk of rows and create a page for it
        for page_num, page_data_rows in enumerate(row_chunks):
            
            # For each page, create a new figure
            fig = plt.figure(figsize=(16, 10))
            gs = GridSpec(1, 2, width_ratios=LEFT_RIGHT_WIDTHS, wspace=0.0)

            # --- LEFT PANEL: TABLE ---
            ax_left = fig.add_subplot(gs[0, 0])
            ax_left.set_xlim(0, 1)
            # Y-limit is based on rows on this page, plus one for the header
            ax_left.set_ylim(-0.5, len(page_data_rows) + 0.5) 
            ax_left.invert_yaxis()
            ax_left.axis("off")

            # --- Draw Table Header for current page ---
            for c in range(4):
                x0, w = COL_EDGES[c], COL_EDGES[c + 1] - COL_EDGES[c]
                ax_left.add_patch(Rectangle((x0, -0.5), w, 1.0, facecolor=PRIMARY, edgecolor=SECONDARY, linewidth=1.2))
                ax_left.text((x0 + COL_EDGES[c+1])/2, 0, HEADERS[c], va="center", ha="center",
                             fontsize=6.5, fontweight="bold", color="white")

            # --- Draw Table Rows for current page ---
            for idx, r in enumerate(page_data_rows, 1): # Start index at 1 for data rows
                is_summary = r["kind"] == "summary"
                text_color = PRIMARY if is_summary else SECONDARY
                font_weight = "bold" if is_summary else "normal"

                # Draw grid cells
                for c in range(4):
                    x0, w = COL_EDGES[c], COL_EDGES[c + 1] - COL_EDGES[c]
                    ax_left.add_patch(Rectangle((x0, idx - 0.5), w, 1.0, fill=False, edgecolor=SECONDARY, linewidth=0.6))

                # Draw text content for the row
                name_text = r["name"].replace('\n', ' ').replace('\r', ' ')
                if len(name_text) > MAX_NAME_LENGTH:
                    name_text = name_text[:MAX_NAME_LENGTH-3] + "..."
                ax_left.text(COL_EDGES[0] + 0.008, idx, name_text, va="center", ha="left",
                             fontsize=6.5, color=text_color, fontweight=font_weight)
                
                if not is_summary:
                    ax_left.text((COL_EDGES[1] + COL_EDGES[2])/2, idx, ("" if r["dur"] is None else f"{r['dur']}d"),
                                 va="center", ha="center", fontsize=6.5, color=SECONDARY)
                    if r["start"]:
                        ax_left.text((COL_EDGES[2] + COL_EDGES[3])/2, idx, r["start"].strftime("%m/%d/%y"),
                                     va="center", ha="center", fontsize=6.5, color=SECONDARY)
                    if r["finish"]:
                        ax_left.text((COL_EDGES[3] + COL_EDGES[4])/2, idx, r["finish"].strftime("%m/%d/%y"),
                                     va="center", ha="center", fontsize=6.5, color=SECONDARY)

            # --- RIGHT PANEL: CHART ---
            ax_right = fig.add_subplot(gs[0, 1], sharey=ax_left)
            ax_right.set_ylim(ax_left.get_ylim())
            ax_right.set_xlim(mdates.date2num(x_min), mdates.date2num(x_max)) # Use global X-limits
            ax_right.xaxis.tick_top()
            ax_right.xaxis.set_label_position('top')
            ax_right.set_yticks([])

            # --- Configure X-axis ticks and grid for each page ---
            # Set locator interval based on total project duration
            if duration_months > 24:
                ax_right.xaxis.set_major_locator(MonthLocator(interval=2))
            else:
                ax_right.xaxis.set_major_locator(MonthLocator(interval=1))
            
            ax_right.xaxis.set_major_formatter(DateFormatter("%b %Y"))

            # Set size and bounding box for labels
            plt.setp(
                ax_right.get_xticklabels(),
                size=7,
                bbox=dict(
                    boxstyle="square,pad=0.2",
                    facecolor='white',
                    edgecolor=SECONDARY,
                    alpha=0.9
                )
            )
            
            ax_right.grid(which="major", axis="x", linestyle="--", linewidth=0.8, color=SECONDARY, alpha=0.6)
            for y in range(len(page_data_rows) + 2):
                ax_right.axhline(y=y - 0.5, linestyle="--", linewidth=0.6, color=SECONDARY, alpha=0.5)
            
            # --- Draw Bars and Milestones for the current page ---
            for idx, r in enumerate(page_data_rows, 1): # Start index at 1
                if not r["start"] or not r["finish"]:
                    continue

                start_num, finish_num = mdates.date2num(r["start"]), mdates.date2num(r["finish"])
                span = max(finish_num - start_num, MIN_WIDTH_DAYS)
                name_text = r["name"].replace('\n', ' ').replace('\r', ' ')
                
                if r["kind"] == "summary":
                    ax_right.broken_barh([(start_num, span)], (idx - 0.05, 0.1), facecolors=SECONDARY)
                    ax_right.plot([start_num, start_num], [idx - 0.2, idx + 0.2], color=SECONDARY, lw=2)
                    ax_right.plot([finish_num, finish_num], [idx - 0.2, idx + 0.2], color=SECONDARY, lw=2)
                else: # Task or Milestone
                    is_milestone = (finish_num - start_num) < 0.01
                    if is_milestone:
                        ax_right.plot(start_num, idx, marker='D', markersize=6, color=PRIMARY)
                    else:
                        ax_right.broken_barh([(start_num, span)], (idx - 0.25, 0.5), facecolors=PRIMARY,
                                             edgecolors=SECONDARY, linewidth=0.8)

            # Add titles and project info ONLY on the first page
            if page_num == 0:
                fig.text(0.5, 0.96, title, ha='center', va='bottom', fontsize=16, fontweight='bold', color=PRIMARY)
                info_y_start = 0.98
                if project_title:
                    fig.text(0.03, info_y_start, f"Project: {project_title}",
                             ha='left', va='top', fontsize=8, fontweight='bold', color=PRIMARY)
                if customer_name:
                    customer_y = info_y_start - 0.025 if project_title else info_y_start
                    fig.text(0.03, customer_y, f"Customer: {customer_name}",
                             ha='left', va='top', fontsize=8, color=PRIMARY)

            # Adjust layout and save the current figure to the PDF
            plt.subplots_adjust(top=0.88, left=0.03, right=0.99, bottom=0.08, wspace=0.0)
            pdf.savefig(fig, bbox_inches="tight")
            plt.close(fig) # Close the figure to free memory
        
def _build_one_page(pdf, page_rows, page_num, total_pages, x_min, x_max, true_start, 
                    total_span_days, chart_width_inches, title, project_title, customer_name, logo_path):
    """Renders a single page of the Gantt chart."""
    header_row = {
        "name": "Task", "start": None, "finish": None, "kind": "header", "dur": "Duration"
    }
    all_rows_for_page = [header_row] + page_rows

    fig = plt.figure(figsize=(16, 10))

    # Manual Axes Creation for Fixed Row Height
    left_margin, right_margin, bottom_margin, top_margin = 0.03, 0.99, 0.10, 0.88
    plot_area_w = right_margin - left_margin
    full_plot_area_h = top_margin - bottom_margin
    rows_on_full_page = MAX_ROWS_PER_PAGE + 1
    rows_on_this_page = len(all_rows_for_page)
    height_ratio = rows_on_this_page / rows_on_full_page
    current_plot_h = full_plot_area_h * height_ratio
    current_bottom = top_margin - current_plot_h
    table_w = plot_area_w * (LEFT_RIGHT_WIDTHS[0] / sum(LEFT_RIGHT_WIDTHS))
    chart_w = plot_area_w * (LEFT_RIGHT_WIDTHS[1] / sum(LEFT_RIGHT_WIDTHS))
    table_l, chart_l = left_margin, left_margin + table_w
    ax_left = fig.add_axes([table_l, current_bottom, table_w, current_plot_h])
    ax_right = fig.add_axes([chart_l, current_bottom, chart_w, current_plot_h])
    ax_right.sharey(ax_left)

    # LEFT: table
    ax_left.set_xlim(0, 1)
    ax_left.set_ylim(-0.5, len(all_rows_for_page) - 0.5)
    ax_left.invert_yaxis()
    ax_left.axis("off")

    for idx, r in enumerate(all_rows_for_page):
        is_header_row = r["kind"] == "header"
        is_summary_row = r["kind"] == "summary"
        for c in range(4):
            x0 = COL_EDGES[c]
            w = COL_EDGES[c + 1] - COL_EDGES[c]
            if is_header_row or is_summary_row:
                ax_left.add_patch(Rectangle((x0, idx - 0.5), w, 1.0, facecolor=PRIMARY, 
                                            edgecolor=SECONDARY, linewidth=1.2 if is_header_row else 0.6))
            else:
                ax_left.add_patch(Rectangle((x0, idx - 0.5), w, 1.0, fill=False, 
                                            edgecolor=SECONDARY, linewidth=0.6))
        if is_header_row:
            for c, header_text in enumerate(HEADERS):
                center_x = (COL_EDGES[c] + COL_EDGES[c + 1]) / 2
                ax_left.text(center_x, idx, header_text, va="center", ha="center", 
                             fontsize=FONTSIZE_TABLE, fontweight="bold", color="white")
        else:
            name_text = r["name"].replace('\n', ' ').replace('\r', ' ')
            if len(name_text) > MAX_NAME_LENGTH:
                name_text = name_text[:MAX_NAME_LENGTH-3] + "..."
            text_color = "white" if is_summary_row else SECONDARY
            font_weight = "bold" if is_summary_row else "normal"
            ax_left.text(COL_EDGES[0] + 0.008, idx, name_text, va="center", ha="left", 
                         fontsize=FONTSIZE_TABLE, fontweight=font_weight, color=text_color)
            ax_left.text((COL_EDGES[1] + COL_EDGES[2]) / 2, idx, 
                         ("" if r["dur"] is None else f"{r['dur']}d"), 
                         va="center", ha="center", fontsize=FONTSIZE_TABLE, color=text_color)
            if r["start"]:
                ax_left.text((COL_EDGES[2] + COL_EDGES[3]) / 2, idx, 
                             r["start"].strftime("%m/%d/%y"), va="center", ha="center", fontsize=FONTSIZE_TABLE, color=text_color)
            if r["finish"]:
                ax_left.text((COL_EDGES[3] + COL_EDGES[4]) / 2, idx, 
                             r["finish"].strftime("%m/%d/%y"), va="center", ha="center", fontsize=FONTSIZE_TABLE, color=text_color)
    
    ax_left.axhline(y=-0.5, linestyle="-", linewidth=2.0, color=SECONDARY)
    ax_left.axhline(y=len(all_rows_for_page) - 0.5, linestyle="-", linewidth=2.0, color=SECONDARY)
    ax_left.axvline(x=0, linestyle="-", linewidth=2.0, color=SECONDARY)

    # RIGHT: chart
    chart_right_edge = mdates.date2num(x_max)
    ax_right.set_xlim(mdates.date2num(x_min), chart_right_edge)
    ax_right.xaxis.tick_top()
    ax_right.xaxis.set_label_position('top')
    project_duration_months = (x_max.year - x_min.year) * 12 + (x_max.month - x_min.month)
    ax_right.xaxis.set_major_locator(MonthLocator(interval=2 if project_duration_months > 24 else 1))
    ax_right.xaxis.set_major_formatter(DateFormatter("%b %Y"))
    ax_right.xaxis.set_minor_locator(mticker.NullLocator())
    ax_right.tick_params(axis='x', which='major', labelsize=FONTSIZE_XTICK, pad=2)
    for label in ax_right.get_xticklabels():
        label.set_bbox(dict(boxstyle="round,pad=0.3", facecolor="white", edgecolor=SECONDARY, linewidth=0.8))
    ax_right.grid(which="major", axis="x", linestyle="--", linewidth=0.8, color=SECONDARY, alpha=0.6)
    for y in range(len(all_rows_for_page) + 1):
        ax_right.axhline(y=y - 0.5, linestyle="--", linewidth=0.6, color=SECONDARY, alpha=0.5)
    ax_right.axhline(y=-0.5, linestyle="-", linewidth=2.0, color=SECONDARY)
    ax_right.axhline(y=len(all_rows_for_page) - 0.5, linestyle="-", linewidth=2.0, color=SECONDARY)
    ax_right.axvline(x=chart_right_edge, linestyle="-", linewidth=2.0, color=SECONDARY)
    ax_right.set_yticks([])
    ax_right.tick_params(axis='y', which='both', length=0)
    chart_left_edge = mdates.date2num(x_min)
    
    for idx, r in enumerate(all_rows_for_page):
        if idx == 0: continue
            
        if r["kind"] != "summary" and r["start"] and r["finish"]:
            start_num, finish_num = mdates.date2num(r["start"]), mdates.date2num(r["finish"])
            is_milestone = (finish_num - start_num) < 0.01
            
            if is_milestone:
                diamond_height = 0.3
                diamond_width_days = max(3, total_span_days * 0.005) 
                diamond_x, diamond_y = start_num, idx
                diamond_verts = [(diamond_x, diamond_y + diamond_height), (diamond_x + diamond_width_days, diamond_y), (diamond_x, diamond_y - diamond_height), (diamond_x - diamond_width_days, diamond_y), (diamond_x, diamond_y + diamond_height)]
                diamond = plt.Polygon(diamond_verts, facecolor=PRIMARY, edgecolor=SECONDARY, linewidth=1.5)
                ax_right.add_patch(diamond)
                name_text = r["name"].replace('\n', ' ').replace('\r', ' ')
                text_width_days = estimate_text_width_in_days(name_text, FONTSIZE_CHART_SUMMARY, total_span_days, chart_width_inches)
                text_start_x = diamond_x + diamond_width_days + (total_span_days * 0.005)
                if chart_right_edge - text_start_x > text_width_days:
                    ax_right.text(text_start_x, diamond_y, name_text, va="center", ha="left", fontsize=FONTSIZE_CHART_SUMMARY, color=SECONDARY)
            else:
                span = max(finish_num - start_num, MIN_WIDTH_DAYS)
                bar_height, bar_y = 0.5, idx - 0.25
                ax_right.broken_barh([(start_num, span)], (bar_y, bar_height), facecolors=PRIMARY, edgecolors=SECONDARY, linewidth=0.8)
                name_text = r["name"].replace('\n', ' ').replace('\r', ' ')
                text_width_days = estimate_text_width_in_days(name_text, FONTSIZE_CHART_REGULAR, total_span_days, chart_width_inches)
                margin_days = max(1.5, total_span_days * 0.015)
                space_on_right = chart_right_edge - finish_num - margin_days
                space_on_left = start_num - chart_left_edge - margin_days
                if space_on_right >= text_width_days:
                    ax_right.text(finish_num + margin_days, idx, name_text, va="center", ha="left", fontsize=FONTSIZE_CHART_REGULAR, color=SECONDARY)
                elif space_on_left >= text_width_days:
                    ax_right.text(start_num - margin_days, idx, name_text, va="center", ha="right", fontsize=FONTSIZE_CHART_REGULAR, color=SECONDARY)
        elif r["kind"] == "summary" and r["start"] and r["finish"]:
            start_num, finish_num = mdates.date2num(r["start"]), mdates.date2num(r["finish"])
            span = max(finish_num - start_num, MIN_WIDTH_DAYS)
            bar_height, bar_y, cap_height = 0.1, idx - 0.05, 0.4
            name_text = r["name"].replace('\n', ' ').replace('\r', ' ')
            text_width_days = estimate_text_width_in_days(name_text, FONTSIZE_CHART_SUMMARY, total_span_days, chart_width_inches)
            margin_days = max(1.5, total_span_days * 0.015)
            space_on_right = chart_right_edge - (start_num + span) - margin_days
            space_on_left = start_num - chart_left_edge - margin_days
            if space_on_right >= text_width_days:
                ax_right.broken_barh([(start_num, span)], (bar_y, bar_height), facecolors=SECONDARY, edgecolors=SECONDARY, linewidth=1.0)
                ax_right.plot([start_num, start_num], [idx - cap_height/2, idx + cap_height/2], color=SECONDARY, linewidth=2)
                ax_right.plot([start_num + span, start_num + span], [idx - cap_height/2, idx + cap_height/2], color=SECONDARY, linewidth=2)
                ax_right.text(start_num + span + margin_days, idx, name_text, va="center", ha="left", fontsize=FONTSIZE_CHART_SUMMARY, color=SECONDARY)
            elif space_on_left >= text_width_days:
                ax_right.broken_barh([(start_num, span)], (bar_y, bar_height), facecolors=SECONDARY, edgecolors=SECONDARY, linewidth=1.0)
                ax_right.plot([start_num, start_num], [idx - cap_height/2, idx + cap_height/2], color=SECONDARY, linewidth=2)
                ax_right.plot([start_num + span, start_num + span], [idx - cap_height/2, idx + cap_height/2], color=SECONDARY, linewidth=2)
                ax_right.text(start_num - margin_days, idx, name_text, va="center", ha="right", fontsize=FONTSIZE_CHART_SUMMARY, color=SECONDARY)
            else:
                bar_center = start_num + span / 2
                text_space_needed = min(text_width_days + 0.2, span * 0.9)
                left_split_end = bar_center - text_space_needed / 2
                right_split_start = bar_center + text_space_needed / 2
                min_segment_width = max(0.5, total_span_days * 0.005)
                if left_split_end > start_num and (left_split_end - start_num) >= min_segment_width:
                    ax_right.broken_barh([(start_num, left_split_end - start_num)], (bar_y, bar_height), facecolors=SECONDARY, edgecolors=SECONDARY, linewidth=1.0)
                    ax_right.plot([start_num, start_num], [idx - cap_height/2, idx + cap_height/2], color=SECONDARY, linewidth=2)
                if right_split_start < start_num + span and ((start_num + span) - right_split_start) >= min_segment_width:
                    ax_right.broken_barh([(right_split_start, (start_num + span) - right_split_start)], (bar_y, bar_height), facecolors=SECONDARY, edgecolors=SECONDARY, linewidth=1.0)
                    ax_right.plot([start_num + span, start_num + span], [idx - cap_height/2, idx + cap_height/2], color=SECONDARY, linewidth=2)
                ax_right.text(bar_center, idx, name_text, va="center", ha="center", fontsize=FONTSIZE_CHART_SUMMARY, color=SECONDARY)

    ax_right.axvline(mdates.date2num(true_start), linestyle="--", linewidth=1.0, color=SECONDARY, alpha=0.4)
    fig.text(0.5, 0.96, title, ha='center', va='bottom', fontsize=16, fontweight='bold', color=PRIMARY)
    info_y_start = 0.98
    if project_title:
        fig.text(0.03, info_y_start, f"Project: {project_title}", ha='left', va='top', fontsize=11, fontweight='bold', color=PRIMARY)
    if customer_name:
        customer_y = info_y_start - 0.025 if project_title else info_y_start
        fig.text(0.03, customer_y, f"Customer: {customer_name}", ha='left', va='top', fontsize=11, color=PRIMARY)
    if logo_path:
        try:
            logo_img = plt.imread(logo_path)
            logo_height_fig, right_edge, top_edge = 0.06, 0.99, 0.98
            aspect_ratio = logo_img.shape[1] / logo_img.shape[0]
            fig_width_in, fig_height_in = fig.get_size_inches()
            logo_height_in = logo_height_fig * fig_height_in
            logo_width_in = logo_height_in * aspect_ratio
            logo_width_fig = logo_width_in / fig_width_in
            left_pos = right_edge - logo_width_fig
            bottom_pos = top_edge - logo_height_fig
            ax_logo = fig.add_axes([left_pos, bottom_pos, logo_width_fig, logo_height_fig], anchor='NE', zorder=10)
            ax_logo.imshow(logo_img)
            ax_logo.axis('off')
        except Exception as e:
            print(f"Warning: Could not load or place logo. Error: {e}")
    if total_pages > 1:
        fig.text(0.99, 0.01, f'Page {page_num} of {total_pages}', ha='right', va='bottom', fontsize=7, color='gray')

    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def build_gantt(rows, out_pdf, title="Project Schedule", project_title="", customer_name="", logo_path=""):
    """Render the Gantt to a single-page or multi-page PDF."""
    if not rows:
        print("Warning: No tasks to plot.")
        return
        
    starts = [r["start"] for r in rows if r["start"]]
    finishes = [r["finish"] for r in rows if r["finish"]]
    if not starts or not finishes:
        raise ValueError("No dated tasks found.")
    true_start, true_finish = min(starts), max(finishes)
    
    span_days = max(1, (true_finish.date() - true_start.date()).days)
    diamond_width_days = max(3, span_days * 0.005)
    left_pad_days = diamond_width_days + 2
    right_pad_days = max(1, int(span_days * 0.05))
    x_min = true_start - timedelta(days=left_pad_days)
    x_max = true_finish + timedelta(days=right_pad_days)
    total_span_days = (x_max.date() - x_min.date()).days
    
    for r in rows:
        r["dur"] = compute_duration_days(r["start"], r["finish"])

    with PdfPages(out_pdf) as pdf:
        total_pages = (len(rows) + MAX_ROWS_PER_PAGE - 1) // MAX_ROWS_PER_PAGE
        chart_width_inches = 16 * LEFT_RIGHT_WIDTHS[1] / sum(LEFT_RIGHT_WIDTHS)

        for i in range(0, len(rows), MAX_ROWS_PER_PAGE):
            page_rows_chunk = rows[i : i + MAX_ROWS_PER_PAGE]
            page_num = (i // MAX_ROWS_PER_PAGE) + 1
            print(f"Generating page {page_num} of {total_pages}...")
            _build_one_page(
                pdf=pdf, page_rows=page_rows_chunk, page_num=page_num, total_pages=total_pages,
                x_min=x_min, x_max=x_max, true_start=true_start, total_span_days=total_span_days,
                chart_width_inches=chart_width_inches, title=title, project_title=project_title,
                customer_name=customer_name, logo_path=logo_path
            )
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xml", required=True, help="Path to Microsoft Project XML file")
    ap.add_argument("--out", required=True, help="Output PDF path")
    ap.add_argument("--title", default="Project Schedule", help="Chart title")
    ap.add_argument("--project", default="", help="Project title")
    ap.add_argument("--customer", default="", help="Customer name")
    ap.add_argument("--logo", default="", help="Path to logo image file for top-right corner")
    args = ap.parse_args()

    rows = parse_msproject_xml(args.xml)

    # Keep order as in XML; (optional) you can filter out a top-level project summary if you wish:
    # if rows and rows[0]['kind'] == 'summary' and not rows[0]['start'] and not rows[0]['finish']:
    #     pass  # we currently keep it as a label row for visual structure

    build_gantt(rows, args.out, title=args.title, project_title=args.project, customer_name=args.customer, logo_path=args.logo)


if __name__ == "__main__":
    main()