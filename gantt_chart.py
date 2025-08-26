from datetime import datetime, timedelta
import re
from reportlab.platypus import Flowable
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor, black, gray, lightgrey


class GanttChartFlowable(Flowable):
    """A ReportLab Flowable to draw a high-quality, vector-based Gantt chart."""

    def __init__(self, tasks, start_date, end_date, width=10.2 * inch, height=7.0 * inch):
        super().__init__()
        # CORRECTED: Removed sorting to preserve the tree view order
        self.tasks = tasks
        self.start_date = start_date
        self.end_date = end_date
        self.width = width
        self.height = height

        # --- Styling ---
        self.bar_color = HexColor("#991f2b")
        self.y_margin = 0.2 * inch
        self.timeline_height = 0.3 * inch
        self.x_margin = 0.1 * inch
        self.label_width = 3.2 * inch
        self.font_name = 'Helvetica'
        self.font_size = 8
        self.header_font_size = 9

        # --- Dynamic Row Height Calculation ---
        if self.tasks:
            # CORRECTED: Account for both top and bottom margins for accurate available height
            available_height = self.height - (self.y_margin * 2) - self.timeline_height
            if available_height > 0 and len(self.tasks) > 0:
                self.row_height = available_height / len(self.tasks)
                # Cap row height to prevent it from being excessively large for few tasks
                self.row_height = min(self.row_height, 0.4 * inch)
            else:
                self.row_height = 0  # No space for rows
        else:
            self.row_height = 0.25 * inch  # Default value

        # --- Calculated dimensions ---
        self.chart_width = self.width - self.label_width - self.x_margin
        self.total_days = (self.end_date - self.start_date).days + 1
        self.px_per_day = self.chart_width / self.total_days if self.total_days > 0 else 0

    def wrap(self, availWidth, availHeight):
        """Specifies the size of the flowable."""
        return self.width, self.height

    def _draw_timeline(self):
        """Draws the timeline with centered month labels at the bottom, preventing overlap."""
        self.canv.setFont(self.font_name, self.header_font_size)

        timeline_text_y = self.y_margin
        line_top_y = self.height - self.y_margin
        line_bottom_y = self.y_margin + self.timeline_height

        month_positions = {}
        for day in range(self.total_days):
            date = self.start_date + timedelta(days=day)
            month_key = (date.year, date.month)
            x_pos = self.label_width + day * self.px_per_day
            if month_key not in month_positions:
                month_positions[month_key] = {'start_x': x_pos, 'end_x': x_pos}
            else:
                month_positions[month_key]['end_x'] = x_pos

        last_label_end_x = -1
        padding = 4  # Minimum pixels between labels

        sorted_month_keys = sorted(month_positions.keys())

        # Draw labels first, checking for collisions
        for month_key in sorted_month_keys:
            pos = month_positions[month_key]
            month_width_px = pos['end_x'] - pos['start_x']

            month_name = datetime(month_key[0], month_key[1], 1).strftime('%b-%y')
            text_width = pdfmetrics.stringWidth(month_name, self.font_name, self.header_font_size)

            if month_width_px > text_width + padding:
                center_x = (pos['start_x'] + pos['end_x']) / 2
                label_start_x = center_x - (text_width / 2)

                if label_start_x > last_label_end_x:
                    self.canv.drawCentredString(center_x, timeline_text_y, month_name)
                    last_label_end_x = center_x + (text_width / 2)

        # Draw vertical lines separately
        self.canv.setStrokeColor(gray)
        self.canv.setLineWidth(0.5)
        for month_key in sorted_month_keys:
            pos = month_positions[month_key]
            if pos['start_x'] > self.label_width:
                self.canv.line(pos['start_x'], line_top_y, pos['start_x'], line_bottom_y)

    def _draw_tasks(self):
        """Draws the task labels, bars, and horizontal gridlines."""
        self.canv.setFont(self.font_name, self.font_size)
        chart_area_x_start = self.label_width
        chart_area_x_end = self.width - self.x_margin

        y_start_for_tasks = self.height - self.y_margin

        # Draw horizontal lines for all rows
        self.canv.setStrokeColor(lightgrey)
        self.canv.setLineWidth(0.5)
        for i in range(len(self.tasks) + 1):
            grid_y = y_start_for_tasks - self.row_height * i
            self.canv.line(chart_area_x_start, grid_y, chart_area_x_end, grid_y)

        for i, task in enumerate(self.tasks):
            y_base = y_start_for_tasks - self.row_height * (i + 1)
            bar_height = self.row_height * 0.6
            y_pos = y_base + (self.row_height - bar_height) / 2

            # Draw task name
            self.canv.setFillColor(black)
            task_name = re.sub(r'^\(\d+\)\s*', '', task["name"].split(" > ")[-1])
            text_y = y_pos + (bar_height / 2) - (self.font_size / 2)
            self.canv.drawString(self.x_margin, text_y, task_name[:65])

            # Draw task bar
            days_from_start = (task['start'] - self.start_date).days
            duration_days = (task['end'] - task['start']).days + 1
            bar_x = self.label_width + days_from_start * self.px_per_day
            bar_width = duration_days * self.px_per_day

            self.canv.setFillColor(self.bar_color)
            self.canv.setStrokeColor(black)
            self.canv.setLineWidth(0.5)
            self.canv.rect(bar_x, y_pos, bar_width, bar_height, stroke=1, fill=1)

    def draw(self):
        """The main drawing method called by ReportLab."""
        if not self.tasks or self.px_per_day == 0:
            self.canv.setFont(self.font_name, 12)
            self.canv.drawCentredString(self.width / 2, self.height / 2, "No task data.")
            return

        self.canv.saveState()
        self._draw_tasks()
        self._draw_timeline()

        # --- Draw bounding box ---
        box_top_y = self.height - self.y_margin
        box_bottom_y = self.y_margin + self.timeline_height
        box_height = box_top_y - box_bottom_y
        self.canv.setStrokeColor(gray)
        self.canv.setLineWidth(0.5)
        self.canv.rect(self.label_width, box_bottom_y, self.chart_width, box_height)

        self.canv.restoreState()
