import argparse
import re
import textwrap
from datetime import datetime, timedelta
import xml.etree.ElementTree as ET

import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as mticker
from matplotlib.gridspec import GridSpec
from matplotlib.patches import Rectangle
from matplotlib.dates import MonthLocator, DateFormatter, WeekdayLocator, MO
from matplotlib.backends.backend_pdf import PdfPages


PRIMARY = "#991f2b"
SECONDARY = "black"

# Layout (adjusted to give the chart most of the width)
LEFT_RIGHT_WIDTHS = [1.2, 2.8]  # table | chart (increased table width)
# Table column edges (fractions of left panel width): Task | Duration | Start | Finish
COL_EDGES = [0.00, 0.52, 0.64, 0.82, 1.00]  # Task | Duration | Start | Finish
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

def build_gantt(rows, out_pdf, title="Project Schedule", project_title="", customer_name=""):
    """Render the Gantt to a single-page PDF."""
    # Add header row as the first row
    header_row = {
        "name": "Task",
        "start": None,
        "finish": None,
        "kind": "header",
        "dur": "Duration"
    }
    
    # Insert header at the beginning
    all_rows = [header_row] + rows
    
    # Filter dates to find bounds (skip header row)
    starts = [r["start"] for r in rows if r["start"]]
    finishes = [r["finish"] for r in rows if r["finish"]]
    if not starts or not finishes:
        raise ValueError("No dated tasks found.")
    true_start = min(starts)
    true_finish = max(finishes)

    # Left pad exactly 1 day; right pad ~5% of span (min 1 day)
    left_pad_days = 1
    span_days = max(1, (true_finish.date() - true_start.date()).days)
    right_pad_days = max(1, int(span_days * 0.05))
    x_min = true_start - timedelta(days=left_pad_days)
    x_max = true_finish + timedelta(days=right_pad_days)
    
    total_span_days = (x_max.date() - x_min.date()).days

    # Precompute durations (display) - skip header row
    for r in rows:
        r["dur"] = compute_duration_days(r["start"], r["finish"])

    with PdfPages(out_pdf) as pdf:
        # Increase figure size for better visibility
        fig = plt.figure(figsize=(16, 10))
        gs = GridSpec(1, 2, width_ratios=LEFT_RIGHT_WIDTHS, wspace=0.0)  # No gap between table and chart

        # LEFT: table
        ax_left = fig.add_subplot(gs[0, 0])
        ax_left.set_xlim(0, 1)
        # Adjust ylim to accommodate header row at top (index 0)
        ax_left.set_ylim(-0.5, len(all_rows) - 0.5)
        ax_left.invert_yaxis()
        ax_left.axis("off")

        # Unified rows rendering (header + data rows)
        for idx, r in enumerate(all_rows):
            # Determine row type
            is_header_row = r["kind"] == "header"
            is_summary_row = r["kind"] == "summary"
            
            # Grid cells - create background for each column
            for c in range(4):
                x0 = COL_EDGES[c]
                w = COL_EDGES[c + 1] - COL_EDGES[c]
                
                if is_header_row or is_summary_row:
                    # Red background for header and summary rows
                    ax_left.add_patch(Rectangle(
                        (x0, idx - 0.5), w, 1.0,
                        facecolor=PRIMARY, edgecolor=SECONDARY, linewidth=1.2 if is_header_row else 0.6
                    ))
                else:
                    # Normal cell for regular tasks
                    ax_left.add_patch(Rectangle(
                        (x0, idx - 0.5), w, 1.0,
                        fill=False, edgecolor=SECONDARY, linewidth=0.6
                    ))

            if is_header_row:
                # Header row text - centered
                headers_text = ["Task", "Duration", "Start", "Finish"]
                for c, header_text in enumerate(headers_text):
                    # Center the text in each column
                    center_x = (COL_EDGES[c] + COL_EDGES[c + 1]) / 2
                    ax_left.text(
                        center_x, idx,
                        header_text, va="center", ha="center",
                        fontsize=6.5, fontweight="bold", color="white"
                    )
            else:
                # Regular row content (summary or task)
                name_text = r["name"].replace('\n', ' ').replace('\r', ' ')
                if len(name_text) > MAX_NAME_LENGTH:
                    name_text = name_text[:MAX_NAME_LENGTH-3] + "..."
                
                if is_summary_row:
                    # White text on red background for summary rows
                    ax_left.text(
                        COL_EDGES[0] + 0.008, idx, name_text,
                        va="center", ha="left", fontsize=6.5,
                        fontweight="bold", color="white"
                    )
                    # Add dates/duration for summary rows too
                    ax_left.text((COL_EDGES[1] + COL_EDGES[2]) / 2, idx,
                                 ("" if r["dur"] is None else f"{r['dur']}d"),
                                 va="center", ha="center", fontsize=6.5, color="white")
                    if r["start"]:
                        ax_left.text((COL_EDGES[2] + COL_EDGES[3]) / 2, idx,
                                     r["start"].strftime("%m/%d/%y"),
                                     va="center", ha="center", fontsize=6.5, color="white")
                    if r["finish"]:
                        ax_left.text((COL_EDGES[3] + COL_EDGES[4]) / 2, idx,
                                     r["finish"].strftime("%m/%d/%y"),
                                     va="center", ha="center", fontsize=6.5, color="white")
                else:
                    # Regular black text for normal tasks
                    ax_left.text(
                        COL_EDGES[0] + 0.008, idx, name_text,
                        va="center", ha="left", fontsize=6.5, color=SECONDARY
                    )
                    # Duration / Start / Finish columns for tasks only
                    ax_left.text((COL_EDGES[1] + COL_EDGES[2]) / 2, idx,
                                 ("" if r["dur"] is None else f"{r['dur']}d"),
                                 va="center", ha="center", fontsize=6.5, color=SECONDARY)
                    if r["start"]:
                        ax_left.text((COL_EDGES[2] + COL_EDGES[3]) / 2, idx,
                                     r["start"].strftime("%m/%d/%y"),
                                     va="center", ha="center", fontsize=6.5, color=SECONDARY)
                    if r["finish"]:
                        ax_left.text((COL_EDGES[3] + COL_EDGES[4]) / 2, idx,
                                     r["finish"].strftime("%m/%d/%y"),
                                     va="center", ha="center", fontsize=6.5, color=SECONDARY)

        # Add thick outside border for table (except right side where it meets chart)
        # Top border
        ax_left.axhline(y=-0.5, linestyle="-", linewidth=2.0, color=SECONDARY)
        # Bottom border  
        ax_left.axhline(y=len(all_rows) - 0.5, linestyle="-", linewidth=2.0, color=SECONDARY)
        # Left border
        ax_left.axvline(x=0, linestyle="-", linewidth=2.0, color=SECONDARY)
        # No right border - let it connect seamlessly to chart

        # RIGHT: chart
        ax_right = fig.add_subplot(gs[0, 1], sharey=ax_left)
        ax_right.set_ylim(ax_left.get_ylim())
        ax_right.set_xlim(mdates.date2num(x_min), mdates.date2num(x_max))

        # Move x-axis to top
        ax_right.xaxis.tick_top()
        ax_right.xaxis.set_label_position('top')

        # X ticks: adjust based on project duration
        project_duration_months = (true_finish.year - true_start.year) * 12 + (true_finish.month - true_start.month)
        
        if project_duration_months > 24:
            # For projects longer than 24 months, show every 2 months
            ax_right.xaxis.set_major_locator(MonthLocator(interval=2))
        else:
            # For projects 24 months or less, show every month
            ax_right.xaxis.set_major_locator(MonthLocator())
            
        ax_right.xaxis.set_major_formatter(DateFormatter("%b %Y"))
        ax_right.xaxis.set_minor_locator(mticker.NullLocator())  # Remove minor ticks
        ax_right.xaxis.set_minor_formatter(mticker.NullFormatter())

        # Style the x-axis labels - smaller font and add boxes
        ax_right.tick_params(axis='x', which='major', labelsize=7, pad=2)
        
        # Add boxes around x-axis labels
        for label in ax_right.get_xticklabels():
            label.set_bbox(dict(boxstyle="round,pad=0.3", facecolor="white", edgecolor=SECONDARY, linewidth=0.8))

        # Grids (dashed lines) - only major
        ax_right.grid(which="major", axis="x", linestyle="--", linewidth=0.8, color=SECONDARY, alpha=0.6)

        # Horizontal row separators - align exactly with table grid (dashed lines)
        for y in range(len(all_rows) + 1):  # Include bottom border, account for header
            ax_right.axhline(y=y - 0.5, linestyle="--", linewidth=0.6, color=SECONDARY, alpha=0.5)

        # Add thick outside border for chart (except left side where it meets table)
        # Top border
        ax_right.axhline(y=-0.5, linestyle="-", linewidth=2.0, color=SECONDARY)
        # Bottom border  
        ax_right.axhline(y=len(all_rows) - 0.5, linestyle="-", linewidth=2.0, color=SECONDARY)
        # No left border - let it connect seamlessly to table
        # Right border
        ax_right.axvline(x=mdates.date2num(x_max), linestyle="-", linewidth=2.0, color=SECONDARY)

        # Hide y-axis labels; the table is the y-axis
        ax_right.set_yticks([])
        ax_right.tick_params(axis='y', which='both', length=0)

        # Calculate chart width for text width estimation
        chart_width_inches = 16 * LEFT_RIGHT_WIDTHS[1] / sum(LEFT_RIGHT_WIDTHS)  # Approximate chart width
        
        # Calculate the actual left boundary of the chart area in data coordinates
        # The chart starts where the table ends, which is at the boundary between the two GridSpec columns
        # We need to map this to the chart's data coordinate system
        total_data_span = mdates.date2num(x_max) - mdates.date2num(x_min)
        
        # The chart area starts at x_min in data coordinates, but visually it should respect
        # a much larger margin from the absolute left edge of the chart panel to prevent table overlap
        # Use a larger percentage of the total span as the visual left margin
        visual_left_margin_days = total_data_span * 0.08  # 8% margin from true left edge (increased from 2%)
        chart_visual_left_boundary = mdates.date2num(x_min) + visual_left_margin_days

        # Draw bars and milestones for all non-summary, non-header tasks
        # Start from index 1 to skip header row
        for idx, r in enumerate(all_rows):
            if idx == 0:  # Skip header row
                continue
                
            if r["kind"] != "summary" and r["start"] and r["finish"]:
                start_num = mdates.date2num(r["start"])
                finish_num = mdates.date2num(r["finish"])
                
                # Determine if this is a milestone (0-day task)
                is_milestone = (finish_num - start_num) < 0.01
                
                if is_milestone:
                    # Draw diamond milestone marker
                    diamond_size = 0.15
                    diamond_x = start_num
                    diamond_y = idx
                    
                    # Create diamond shape (rotated square)
                    diamond_verts = [
                        (diamond_x, diamond_y + diamond_size),  # top
                        (diamond_x + diamond_size*2, diamond_y),  # right
                        (diamond_x, diamond_y - diamond_size),  # bottom
                        (diamond_x - diamond_size*2, diamond_y),  # left
                        (diamond_x, diamond_y + diamond_size)   # close
                    ]
                    
                    diamond = plt.Polygon(diamond_verts, facecolor=PRIMARY, 
                                          edgecolor=SECONDARY, linewidth=1.0)
                    ax_right.add_patch(diamond)
                    
                    # Add task name to the right of milestone
                    name_text = r["name"].replace('\n', ' ').replace('\r', ' ')
                    # Don't truncate milestone text
                    
                    # For milestones, only place text to the right to avoid table overlap
                    text_x = diamond_x + diamond_size*3
                    ax_right.text(text_x, diamond_y, name_text,
                                  va="center", ha="left", fontsize=5.5, color=SECONDARY)
                    
                else:
                    # Draw regular task bar with smart label placement
                    span = max(finish_num - start_num, MIN_WIDTH_DAYS)
                    bar_height = 0.5
                    bar_y = idx - bar_height/2
                    
                    # Prepare task name for labeling - don't truncate
                    name_text = r["name"].replace('\n', ' ').replace('\r', ' ')

                    # Define edges and margins very conservatively to prevent table overlap
                    chart_right_edge = mdates.date2num(x_max)
                    chart_left_edge = chart_visual_left_boundary  # Use the visual left boundary with large margin
                    
                    # Larger margin from boundaries to be extra safe
                    text_margin_days = max(2, total_span_days * 0.03)  # 3% margin or minimum 2 days
                    
                    # Estimate text width in days with very conservative calculation
                    text_width_days = estimate_text_width_in_days(name_text, 5, total_span_days, chart_width_inches)
                    
                    # Check available space very strictly
                    space_on_right = chart_right_edge - finish_num - text_margin_days
                    space_on_left = start_num - chart_left_edge - text_margin_days
                    
                    if space_on_right >= text_width_days:
                        # Case 1: Enough space to the right of the bar
                        ax_right.broken_barh(
                            [(start_num, span)],
                            (bar_y, bar_height),
                            facecolors=PRIMARY,
                            edgecolors=SECONDARY,
                            linewidth=0.8
                        )
                        text_x = finish_num + text_margin_days
                        text_ha = "left"
                        text_color = SECONDARY
                        
                    elif space_on_left >= text_width_days:
                        # Case 2: Enough space to the left of the bar
                        # Verify the text will fit within chart bounds
                        text_end_position = start_num - text_margin_days
                        text_start_position = text_end_position - text_width_days
                        
                        if text_start_position >= chart_left_edge:
                            ax_right.broken_barh(
                                [(start_num, span)],
                                (bar_y, bar_height),
                                facecolors=PRIMARY,
                                edgecolors=SECONDARY,
                                linewidth=0.8
                            )
                            text_x = text_end_position
                            text_ha = "right"
                            text_color = SECONDARY
                        else:
                            # No space on left either - draw bar without label
                            ax_right.broken_barh(
                                [(start_num, span)],
                                (bar_y, bar_height),
                                facecolors=PRIMARY,
                                edgecolors=SECONDARY,
                                linewidth=0.8
                            )
                            # Skip text placement
                            continue
                    else:
                        # No space on either side - draw bar without label
                        ax_right.broken_barh(
                            [(start_num, span)],
                            (bar_y, bar_height),
                            facecolors=PRIMARY,
                            edgecolors=SECONDARY,
                            linewidth=0.8
                        )
                        # Skip text placement
                        continue
                    
                    # Add the text
                    ax_right.text(text_x, idx, name_text,
                                  va="center", ha=text_ha, fontsize=5, color=text_color)
            elif r["kind"] == "summary":
                # Draw summary bar (like MS Project)
                if r["start"] and r["finish"]:
                    start_num = mdates.date2num(r["start"])
                    finish_num = mdates.date2num(r["finish"])
                    span = max(finish_num - start_num, MIN_WIDTH_DAYS)
                    
                    # Summary bar - much thinner
                    bar_height = 0.1  # Made much thinner
                    bar_y = idx - bar_height/2
                    cap_height = 0.4
                    
                    # Add summary task name - try right, then split if no space
                    name_text = r["name"].replace('\n', ' ').replace('\r', ' ')
                    # Don't truncate summary text
                    
                    # Check if there's space to the right
                    chart_right_edge = mdates.date2num(x_max)
                    text_margin_days = max(2, total_span_days * 0.03)
                    text_width_days = estimate_text_width_in_days(name_text, 5.5, total_span_days, chart_width_inches)
                    space_on_right = chart_right_edge - (start_num + span) - text_margin_days
                    
                    if space_on_right >= text_width_days:
                        # Draw regular summary bar
                        ax_right.broken_barh(
                            [(start_num, span)],
                            (bar_y, bar_height),
                            facecolors=SECONDARY,
                            edgecolors=SECONDARY,
                            linewidth=1.0
                        )
                        
                        # Add end caps (vertical lines)
                        ax_right.plot([start_num, start_num], [idx - cap_height/2, idx + cap_height/2], 
                                      color=SECONDARY, linewidth=2)
                        ax_right.plot([start_num + span, start_num + span], [idx - cap_height/2, idx + cap_height/2], 
                                      color=SECONDARY, linewidth=2)
                        
                        # Place text to the right
                        text_x = start_num + span + 0.5
                        ax_right.text(text_x, idx, name_text,
                                      va="center", ha="left", fontsize=5.5, color=SECONDARY, 
                                      fontweight="bold")
                    else:
                        # Split the summary bar and place text in the middle
                        bar_center = start_num + span / 2
                        # Use just the text width with absolutely minimal padding
                        text_space_needed = text_width_days + 0.2  # Just 0.2 days padding total
                        
                        # Ensure we don't make the text space larger than the bar
                        text_space_needed = min(text_space_needed, span * 0.9)  # Max 90% of bar
                        
                        left_split_end = bar_center - text_space_needed / 2
                        right_split_start = bar_center + text_space_needed / 2
                        
                        # Only draw bar segments if they're meaningful
                        min_segment_width = max(0.5, total_span_days * 0.005)
                        
                        # Draw left part of split summary bar
                        if left_split_end > start_num and (left_split_end - start_num) >= min_segment_width:
                            left_span = left_split_end - start_num
                            ax_right.broken_barh(
                                [(start_num, left_span)],
                                (bar_y, bar_height),
                                facecolors=SECONDARY,
                                edgecolors=SECONDARY,
                                linewidth=1.0
                            )
                            # Add left end cap
                            ax_right.plot([start_num, start_num], [idx - cap_height/2, idx + cap_height/2], 
                                          color=SECONDARY, linewidth=2)
                        
                        # Draw right part of split summary bar
                        if right_split_start < start_num + span and ((start_num + span) - right_split_start) >= min_segment_width:
                            right_span = (start_num + span) - right_split_start
                            ax_right.broken_barh(
                                [(right_split_start, right_span)],
                                (bar_y, bar_height),
                                facecolors=SECONDARY,
                                edgecolors=SECONDARY,
                                linewidth=1.0
                            )
                            # Add right end cap
                            ax_right.plot([start_num + span, start_num + span], [idx - cap_height/2, idx + cap_height/2], 
                                          color=SECONDARY, linewidth=2)
                        
                        # Place text in the center gap
                        ax_right.text(bar_center, idx, name_text,
                                      va="center", ha="center", fontsize=5.5, color=SECONDARY, 
                                      fontweight="bold")

        # Faint vertical line at true project start (dashed)
        ax_right.axvline(mdates.date2num(true_start), linestyle="--", linewidth=1.0, color=SECONDARY, alpha=0.4)

        # Add titles and project info using fig.text for precise placement outside the plot axes

        # Main Title (top center)
        fig.text(0.5, 0.96, title, ha='center', va='bottom', fontsize=16, fontweight='bold', color=PRIMARY)

        # Project Information (top left)
        info_y_start = 0.98 
        if project_title:
            fig.text(0.03, info_y_start, f"Project: {project_title}",
                     ha='left', va='top', fontsize=8, fontweight='bold', color=PRIMARY)
        if customer_name:
            # Place customer name below project title if both exist, otherwise place at the top
            customer_y = info_y_start - 0.025 if project_title else info_y_start
            fig.text(0.03, customer_y, f"Customer: {customer_name}",
                     ha='left', va='top', fontsize=8, color=PRIMARY)
        
        # Spacing - no gap between table and chart. Increased top/bottom to move plot up.
        plt.subplots_adjust(top=0.88, left=0.03, right=0.99, bottom=0.10, wspace=0.0)

        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xml", required=True, help="Path to Microsoft Project XML file")
    ap.add_argument("--out", required=True, help="Output PDF path")
    ap.add_argument("--title", default="Project Schedule", help="Chart title")
    ap.add_argument("--project", default="", help="Project title")
    ap.add_argument("--customer", default="", help="Customer name")
    args = ap.parse_args()

    rows = parse_msproject_xml(args.xml)

    # Keep order as in XML; (optional) you can filter out a top-level project summary if you wish:
    # if rows and rows[0]['kind'] == 'summary' and not rows[0]['start'] and not rows[0]['finish']:
    #     pass  # we currently keep it as a label row for visual structure

    build_gantt(rows, args.out, title=args.title, project_title=args.project, customer_name=args.customer)

if __name__ == "__main__":
    main()