import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime, timedelta
import json
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate, NextPageTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
import os
import sys
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import uuid
import io
import math

# --- HELPER FUNCTION FOR PYINSTALLER ---
# This function helps find bundled files (like fonts and logos) when running as an .exe
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class ProposalItem:
    """Represents a single task or milestone in the project."""
    def __init__(self, name, duration=0, price=0, start_date="", is_milestone=False, indent_level=0):
        self.name = name
        self.duration = duration
        self.price = price
        self.start_date = start_date
        self.end_date = ""
        self.is_milestone = is_milestone
        self.indent_level = indent_level
        self.enabled = tk.BooleanVar(value=True)
        self.children = []
        self.parent = None
        # --- Fields for unique ID and predecessor tracking ---
        self.id = str(uuid.uuid4())
        self.predecessor_id = None
        self.predecessor_type = 'FS' # Finish-to-Start
        self.lag = 0 # Lag in days

class ProposalGenerator:
    """The main application class for the PDF Proposal Generator."""
    def __init__(self, root):
        self.root = root
        self.root.title("Castillo Engineering: Proposal Generator")
        self.root.geometry("1600x900")
        self.root.state('zoomed')

        # --- Initialize data ---
        self.project_name = tk.StringVar(value="Sample Project")
        self.company_name = tk.StringVar(value="Sample Company LLC")
        self.project_start_date = tk.StringVar(value="08/11/25")
        
        # --- MODIFICATION: Store the default logo path for later comparison ---
        self.default_logo_path = resource_path("logo.png")
        self.logo_path = tk.StringVar(value=self.default_logo_path)
        self.client_logo_path = tk.StringVar(value="")
        self.include_gantt = tk.BooleanVar(value=False)
        
        self.template_items = self.create_template_structure()
        self.current_editor = None
        self.drag_data = {"item": None, "index": 0}
        self.item_id_map = {}
        # Data for drag-and-drop linking
        self.link_drag_data = {"start_item_id": None, "last_hover_id": None}
        self.column_drag_data = {}

        self.setup_ui()
        self.populate_tree()
        self.expand_all_items()

    def create_template_structure(self):
        """Create the default template structure with sequential predecessors."""
        items = []
        all_tasks = []

        def collect_all_tasks(item_list):
            for item in item_list:
                if not item.is_milestone:
                    all_tasks.append(item)
                if item.children:
                    collect_all_tasks(item.children)

        # Project Initiation
        proj_init = ProposalItem("Project Initiation", 0, 0, "", True, 0)
        proj_init.children = [
            ProposalItem("Deposit & Contract Signed", 0, 0, "", False, 1),
            ProposalItem("Notice to Proceed", 0, 0, "", False, 1),
            ProposalItem("Civil Start - Civil Due Diligence", 1, 0, "", False, 1),
            ProposalItem("Electrical Start - Electrical Due Diligence", 1, 0, "", False, 1),
        ]
        items.append(proj_init)

        # Civil Engineering
        civil_eng = ProposalItem("Civil Engineering", 0, 0, "", True, 0)
        design_30_civil = ProposalItem("30% Design", 0, 0, "", True, 1)
        design_30_civil.children = [
            ProposalItem("30% - Planset/ Basis of Design", 20, 20000, "", False, 2),
            ProposalItem("Pre-Development Hydrology Study", 10, 10000, "", False, 2),
            ProposalItem("Client Review", 10, 0, "", False, 2),
        ]
        design_60_civil = ProposalItem("60% Design", 0, 0, "", True, 1)
        design_60_civil.children = [
            ProposalItem("60% - Planset", 25, 110000, "", False, 2),
            ProposalItem("Stormwater Pollution Prevention Plan", 10, 6000, "", False, 2),
            ProposalItem("Post-Development Hydrology Study", 10, 15000, "", False, 2),
            ProposalItem("Stormwater Management Report", 15, 12000, "", False, 2),
            ProposalItem("Client Review", 10, 0, "", False, 2),
        ]
        design_90_civil = ProposalItem("90% Design", 0, 0, "", True, 1)
        design_90_civil.children = [
            ProposalItem("90% - Planset", 5, 35000, "", False, 2),
            ProposalItem("Client Review", 10, 0, "", False, 2),
        ]
        ifc_design_civil = ProposalItem("IFC Design", 0, 0, "", True, 1)
        ifc_design_civil.children = [ProposalItem("IFC - Planset", 15, 56500, "", False, 2)]
        Studies_update = ProposalItem("Studies Updates", 5, 6500, "", True, 1)
        Studies_update.children = [
            ProposalItem("Stormwater Pollution Prevention Plan", 5, 1000, "", False, 2),
            ProposalItem("Post-Development Hydrology Study", 5, 2500, "", False, 2),
            ProposalItem("Stormwater Management Report", 5, 3000, "", False, 2),
        ]
        civil_eng.children = [design_30_civil, design_60_civil, design_90_civil, ifc_design_civil,Studies_update]
        items.append(civil_eng)

        # Electrical Engineering
        elec_eng = ProposalItem("Electrical Engineering", 0, 0, "", True, 0)
        design_30_elec = ProposalItem("30% Design", 0, 0, "", True, 1)
        design_30_elec.children = [
            ProposalItem("30% - Planset/Basis of Design", 11, 40000, "", False, 2),
            ProposalItem("Reactive Power Study", 6, 18500, "", False, 2),
            ProposalItem("MV - Short Circuit Study", 5, 6500, "", False, 2),
            ProposalItem("SAM Model", 3, 5000, "", False, 2),
            ProposalItem("PV SYST", 3, 5000, "", False, 2),
            ProposalItem("Client Review", 10, 0, "", False, 2),
        ]
        design_60_elec = ProposalItem("60% Design", 0, 0, "", True, 1)
        design_60_elec.children = [
            ProposalItem("60% - Planset", 14, 80000, "", False, 2),
            ProposalItem("DC - Short Circuit Study", 3, 6500, "", False, 2),
            ProposalItem("Under Ground Cable Thermal Study", 8, 10000, "", False, 2),
            ProposalItem("Grounding Study", 8, 13000, "", False, 2),
            ProposalItem("Client Review", 10, 0, "", False, 2),
        ]
        design_90_elec = ProposalItem("90% Design", 0, 0, "", True, 1)
        design_90_elec.children = [
            ProposalItem("90% - Planset", 13, 63500, "", False, 2),
            ProposalItem("Load Flow Study", 2, 13000, "", False, 2),
            ProposalItem("Coordination Study", 2, 9500, "", False, 2),
            ProposalItem("Arc Flash Study", 5, 13000, "", False, 2),
            ProposalItem("Client Review", 10, 0, "", False, 2),
        ]
        ifc_design_elec = ProposalItem("IFC Design", 0, 0, "", True, 1)
        ifc_design_elec.children = [
            ProposalItem("IFC - Planset", 10, 13000, "", False, 2),
        ]
        elec_eng.children = [design_30_elec, design_60_elec, design_90_elec, ifc_design_elec]
        items.append(elec_eng)

        # Structural Engineering
        struct_eng = ProposalItem("Structural Engineering", 0, 0, "", True, 0)
        struct_eng.children = [
            ProposalItem("Structural Engineering (Except racking foundation design)", 10, 25000, "", False, 1),
        ]
        items.append(struct_eng)

        # Project Closeout
        closeout = ProposalItem("Project Closeout", 0, 0, "", True, 0)
        items.append(closeout)

        # --- Set parent relationships ---
        def set_parents(item_list):
            for item in item_list:
                for child in item.children:
                    child.parent = item
                    if child.children:
                        set_parents(child.children)
        set_parents(items)
        
        # --- Set default sequential predecessors ---
        collect_all_tasks(items)
        for i in range(1, len(all_tasks)):
            all_tasks[i].predecessor_id = all_tasks[i-1].id

        return items

    def setup_ui(self):
        """Sets up the main graphical user interface."""
        # Create main frames
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill=tk.X, padx=10, pady=5)

        top_button_frame = ttk.Frame(self.root)
        top_button_frame.pack(fill=tk.X, pady=5)

        content_frame = ttk.Frame(self.root)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        bottom_button_frame = ttk.Frame(self.root)
        bottom_button_frame.pack(fill=tk.X, padx=10, pady=10)


        # --- MODIFIED: Header Section with Title ---
        title_font = ("Arial", 16, "bold")
        title_label = ttk.Label(header_frame, text="Castillo Engineering Proposal Generator", font=title_font, anchor="center")
        title_label.pack(pady=(10, 5), fill=tk.X)

        info_container = ttk.Frame(header_frame)
        info_container.pack(pady=10)
        
        ttk.Label(info_container, text="Project Name:").grid(row=1, column=0, sticky=tk.E, padx=5)
        ttk.Entry(info_container, textvariable=self.project_name, width=40).grid(row=1, column=1, padx=5, sticky=tk.W)
        
        ttk.Label(info_container, text="Company Name:").grid(row=1, column=2, sticky=tk.E, padx=5)
        ttk.Entry(info_container, textvariable=self.company_name, width=40).grid(row=1, column=3, padx=5, sticky=tk.W)
        
        ttk.Label(info_container, text="Project Start Date:").grid(row=2, column=0, sticky=tk.E, padx=5)
        ttk.Entry(info_container, textvariable=self.project_start_date, width=15).grid(row=2, column=1, padx=5, sticky=tk.W)
        
        ttk.Label(info_container, text="Company Logo:").grid(row=3, column=0, sticky=tk.E, padx=5, pady=(5,0))
        logo_entry = ttk.Entry(info_container, textvariable=self.logo_path, width=50, state='readonly')
        logo_entry.grid(row=3, column=1, columnspan=2, padx=5, pady=(5,0), sticky=tk.W)
        ttk.Button(info_container, text="Change Logo", command=self.change_logo).grid(row=3, column=3, padx=5, pady=(5,0), sticky=tk.W)
        
        # --- MODIFIED: Added client logo upload ---
        ttk.Label(info_container, text="Client Logo:").grid(row=4, column=0, sticky=tk.E, padx=5, pady=(5,0))
        client_logo_entry = ttk.Entry(info_container, textvariable=self.client_logo_path, width=50, state='readonly')
        client_logo_entry.grid(row=4, column=1, columnspan=2, padx=5, pady=(5,0), sticky=tk.W)
        ttk.Button(info_container, text="Change Logo", command=self.change_client_logo).grid(row=4, column=3, padx=5, pady=(5,0), sticky=tk.W)
        
        # --- MODIFIED: Added Gantt chart checkbox ---
        gantt_check = ttk.Checkbutton(info_container, text="Include Gantt Chart", variable=self.include_gantt)
        gantt_check.grid(row=5, column=1, columnspan=2, pady=(10,0), sticky=tk.W)


        # --- Top Button Container ---
        top_button_container = ttk.Frame(top_button_frame)
        top_button_container.pack()
        ttk.Button(top_button_container, text="Reset Predecessors", command=self.reset_predecessors).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_button_container, text="Clear Predecessors", command=self.clear_all_predecessors).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_button_container, text="Clear Prices", command=self.clear_all_prices).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_button_container, text="Load Template", command=self.load_template).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_button_container, text="Save Template", command=self.save_template).pack(side=tk.LEFT, padx=5)

        # --- Bottom Button Container ---
        ttk.Button(bottom_button_frame, text="Generate PDF", command=self.generate_pdf).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bottom_button_frame, text="Calculate Dates", command=self.calculate_all_dates).pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom_button_frame, text="Delete Item", command=self.delete_item).pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom_button_frame, text="Add Custom Item", command=self.add_custom_item).pack(side=tk.LEFT, padx=5)


        # Content section with treeview
        tree_frame = ttk.Frame(content_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        column_ids = ('Predecessor', 'Type', 'Enabled', 'Duration', 'Price', 'Start Date', 'End Date')
        self.tree = ttk.Treeview(tree_frame, columns=column_ids, displaycolumns=column_ids, show='tree headings')
        
        self.tree.heading('Predecessor', text='Predecessor')
        self.tree.heading('Type', text='Type')
        self.tree.heading('Enabled', text='Enabled')
        self.tree.heading('Duration', text='Duration (days)')
        self.tree.heading('Price', text='Price ($)')
        self.tree.heading('Start Date', text='Start Date')
        self.tree.heading('End Date', text='End Date')

        self.tree.column('Predecessor', width=150)
        self.tree.column('Type', width=60, anchor='center')
        self.tree.column('Enabled', width=60, anchor='center')
        self.tree.column('Duration', width=100)
        self.tree.column('Price', width=100)
        self.tree.column('Start Date', width=100)
        self.tree.column('End Date', width=100)
        
        v_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        self.tree.tag_configure('milestone', background='#E8E8E8', font=('Arial', 9, 'bold'))
        self.tree.tag_configure('task', font=('Arial', 9))
        self.tree.tag_configure('predecessor_highlight', background='lightgreen')
        self.tree.tag_configure('successor_highlight', background='lightcoral')
        self.tree.tag_configure('linking_highlight', background='lightblue')
        
        # Event Bindings
        self.tree.bind('<Double-1>', self.on_item_double_click)
        self.tree.bind('<Button-1>', self.on_item_click, add='+')
        self.tree.bind('<ButtonPress-1>', self.on_drag_start, add='+')
        self.tree.bind('<B1-Motion>', self.on_drag_motion, add='+')
        self.tree.bind('<ButtonRelease-1>', self.on_drag_release, add='+')
        self.tree.bind('<Control-ButtonPress-1>', self.on_link_start)
        self.tree.bind('<Control-B1-Motion>', self.on_link_drag)
        self.tree.bind('<Control-ButtonRelease-1>', self.on_link_drop)

    def clear_all_prices(self):
        """Sets the price of all items to zero."""
        if messagebox.askyesno("Confirm Clear", "Are you sure you want to clear all prices?"):
            for item in self.item_id_map.values():
                item.price = 0
            self.populate_tree()
            self.expand_all_items()

    def clear_all_predecessors(self):
        """Removes all predecessor links from all tasks."""
        if messagebox.askyesno("Confirm Clear", "Are you sure you want to clear all predecessor links?"):
            for item in self.item_id_map.values():
                item.predecessor_id = None
            self.populate_tree()
            self.expand_all_items()
    
    def reset_predecessors(self):
        """Resets all tasks to have a sequential predecessor link."""
        if messagebox.askyesno("Confirm Reset", "Are you sure you want to reset all predecessors to the default sequential order?"):
            ordered_tasks = []
            def get_tasks_in_order(tree_item_ids):
                for tree_id in tree_item_ids:
                    task_obj = self.tree_item_map.get(tree_id)
                    if task_obj and not task_obj.is_milestone:
                        ordered_tasks.append(task_obj)
                    children = self.tree.get_children(tree_id)
                    if children:
                        get_tasks_in_order(children)

            get_tasks_in_order(self.tree.get_children())
            
            if ordered_tasks:
                ordered_tasks[0].predecessor_id = None
                for i in range(1, len(ordered_tasks)):
                    ordered_tasks[i].predecessor_id = ordered_tasks[i-1].id
            
            self.populate_tree()
            self.expand_all_items()

    def change_logo(self):
        """Open a file dialog to select a new logo file."""
        filepath = filedialog.askopenfilename(
            title="Select Logo File",
            filetypes=[("Image Files", "*.png *.jpg *.jpeg"), ("All files", "*.*")]
        )
        if filepath:
            self.logo_path.set(filepath)

    def change_client_logo(self):
        """Open a file dialog to select a new client logo file."""
        filepath = filedialog.askopenfilename(
            title="Select Client Logo File",
            filetypes=[("Image Files", "*.png *.jpg *.jpeg"), ("All files", "*.*")]
        )
        if filepath:
            self.client_logo_path.set(filepath)

    def on_drag_start(self, event):
        """Prepares for reordering an item or a column."""
        region = self.tree.identify("region", event.x, event.y)

        if region == "heading":
            self.column_drag_data["start_x"] = event.x
            self.column_drag_data["col_id"] = self.tree.identify_column(event.x)
        else:
            if self.tree.identify_column(event.x) != '#0': return 
            item_id = self.tree.identify_row(event.y)
            if item_id and self.tree.parent(item_id):
                self.drag_data["item"] = item_id
                self.drag_data["index"] = self.tree.index(item_id)

    def on_drag_motion(self, event):
        """Moves the dragged item or provides visual feedback for column drag."""
        if self.column_drag_data.get("col_id"):
            self.root.config(cursor="sb_h_double_arrow")
            return

        if not self.drag_data.get("item"): return
        drag_item = self.drag_data["item"]
        parent = self.tree.parent(drag_item)
        self.tree.move(drag_item, parent, self.tree.index(self.tree.identify_row(event.y)))

    def on_drag_release(self, event):
        """Finalizes the item's new position or reorders columns."""
        if self.column_drag_data.get("col_id"):
            self.root.config(cursor="")
            dragged_col_id = self.column_drag_data["col_id"]
            target_col_id = self.tree.identify_column(event.x)
            
            if dragged_col_id and target_col_id:
                dragged_index = int(dragged_col_id.replace('#','')) - 1
                target_index = int(target_col_id.replace('#','')) - 1
                
                cols = list(self.tree['displaycolumns'])
                cols.insert(target_index, cols.pop(dragged_index))
                self.tree['displaycolumns'] = tuple(cols)
            self.column_drag_data = {}
            return

        if not self.drag_data.get("item"): return
        dragged_id = self.drag_data["item"]
        parent_id = self.tree.parent(dragged_id)
        if not parent_id: 
            self.drag_data = {"item": None, "index": 0}
            return
        new_index = self.tree.index(dragged_id)
        dragged_item_obj = self.tree_item_map[dragged_id]
        parent_item_obj = self.tree_item_map.get(parent_id)
        if parent_item_obj and dragged_item_obj in parent_item_obj.children:
            parent_item_obj.children.remove(dragged_item_obj)
            parent_item_obj.children.insert(new_index, dragged_item_obj)
        self.drag_data = {"item": None, "index": 0}
        
    def on_item_click(self, event):
        """Handle single click for checkbox toggle, opening type dropdown, and highlighting."""
        item_id = self.tree.identify_row(event.y)

        if not item_id:
            self.clear_highlights()
            return

        region = self.tree.identify("region", event.x, event.y)
        if region != "cell": 
            self.highlight_dependencies(item_id)
            return

        column_id = self.tree.identify_column(event.x)
        item = self.tree_item_map.get(item_id)
        if not item: return

        # Get current column display order to find correct index
        display_cols = list(self.tree['displaycolumns'])
        
        if column_id == f"#{display_cols.index('Enabled') + 1}":
            item.enabled.set(not item.enabled.get())
            self.update_item_display(item_id, item)
        
        elif column_id == f"#{display_cols.index('Type') + 1}" and not item.is_milestone:
            if item.predecessor_id:
                self.edit_type_cell(item_id, item, column_id)
        
        self.highlight_dependencies(item_id)

    def on_link_start(self, event):
        """Starts a predecessor link drag operation."""
        item_id = self.tree.identify_row(event.y)
        item = self.tree_item_map.get(item_id)
        if item and not item.is_milestone:
            self.link_drag_data["start_item_id"] = item_id

    def on_link_drag(self, event):
        """Updates the visual highlight while dragging."""
        if not self.link_drag_data.get("start_item_id"):
            return

        last_hover_id = self.link_drag_data.get("last_hover_id")
        if last_hover_id and self.tree.exists(last_hover_id):
            tags = list(self.tree.item(last_hover_id, 'tags'))
            if 'linking_highlight' in tags:
                tags.remove('linking_highlight')
                self.tree.item(last_hover_id, tags=tuple(tags))
        
        current_hover_id = self.tree.identify_row(event.y)
        start_item_id = self.link_drag_data["start_item_id"]
        
        if current_hover_id and current_hover_id != start_item_id:
            item = self.tree_item_map.get(current_hover_id)
            if item and not item.is_milestone:
                tags = list(self.tree.item(current_hover_id, 'tags'))
                if 'linking_highlight' not in tags:
                    tags.append('linking_highlight')
                    self.tree.item(current_hover_id, tags=tuple(tags))
                self.link_drag_data["last_hover_id"] = current_hover_id
            else:
                self.link_drag_data["last_hover_id"] = None
        else:
            self.link_drag_data["last_hover_id"] = None

    def on_link_drop(self, event):
        """Finalizes the predecessor link."""
        start_item_id = self.link_drag_data.get("start_item_id")
        if not start_item_id:
            return

        last_hover_id = self.link_drag_data.get("last_hover_id")
        if last_hover_id and self.tree.exists(last_hover_id):
            tags = list(self.tree.item(last_hover_id, 'tags'))
            if 'linking_highlight' in tags:
                tags.remove('linking_highlight')
                self.tree.item(last_hover_id, tags=tuple(tags))

        end_item_id = self.tree.identify_row(event.y)
        start_item = self.tree_item_map.get(start_item_id)
        end_item = self.tree_item_map.get(end_item_id)

        if (start_item and end_item and
            not start_item.is_milestone and not end_item.is_milestone and
            start_item.id != end_item.id):
            
            start_item.predecessor_id = end_item.id
            start_item.predecessor_type = 'FS'
            start_item.lag = 0
            self.update_item_display(start_item_id, start_item)
            self.highlight_dependencies(start_item_id)

        self.link_drag_data = {"start_item_id": None, "last_hover_id": None}

    def clear_highlights(self):
        """Removes all dependency highlighting from the tree."""
        for item_id in self.tree_item_map:
            if self.tree.exists(item_id):
                current_tags = list(self.tree.item(item_id, 'tags'))
                if 'predecessor_highlight' in current_tags:
                    current_tags.remove('predecessor_highlight')
                if 'successor_highlight' in current_tags:
                    current_tags.remove('successor_highlight')
                self.tree.item(item_id, tags=tuple(current_tags))

    def highlight_dependencies(self, selected_item_id):
        """Highlights the predecessor and successors of the selected item."""
        self.clear_highlights()
        
        selected_item = self.tree_item_map.get(selected_item_id)
        if not selected_item or selected_item.is_milestone:
            return

        if selected_item.predecessor_id:
            for tree_id, item_obj in self.tree_item_map.items():
                if item_obj.id == selected_item.predecessor_id:
                    if self.tree.exists(tree_id):
                        current_tags = list(self.tree.item(tree_id, 'tags'))
                        if 'predecessor_highlight' not in current_tags:
                            current_tags.append('predecessor_highlight')
                        self.tree.item(tree_id, tags=tuple(current_tags))
                    break
        
        for tree_id, item_obj in self.tree_item_map.items():
            if item_obj.predecessor_id == selected_item.id:
                if self.tree.exists(tree_id):
                    current_tags = list(self.tree.item(tree_id, 'tags'))
                    if 'successor_highlight' not in current_tags:
                        current_tags.append('successor_highlight')
                    self.tree.item(tree_id, tags=tuple(current_tags))

    def populate_tree(self):
        """Populate the treeview with template items and build ID maps."""
        expanded_items = []
        for item_id in self.tree.get_children():
            if self.tree.item(item_id, 'open'):
                expanded_items.append(self.tree.item(item_id, 'text'))
            expanded_items.extend(self.get_expanded_children(item_id))
        
        self.tree.delete(*self.tree.get_children())
        self.tree_item_map = {}
        self.item_id_map = {}
        
        def build_id_map(items):
            for item in items:
                self.item_id_map[item.id] = item
                if item.children:
                    build_id_map(item.children)
        build_id_map(self.template_items)
        
        for item in self.template_items:
            item_id = self.add_item_to_tree(item, '')
            if any(expanded in self.tree.item(item_id, 'text') for expanded in expanded_items):
                self.tree.item(item_id, open=True)
    
    def get_expanded_children(self, item_id):
        """Get all expanded children recursively."""
        expanded = []
        for child_id in self.tree.get_children(item_id):
            if self.tree.item(child_id, 'open'):
                expanded.append(self.tree.item(child_id, 'text'))
                expanded.extend(self.get_expanded_children(child_id))
        return expanded
    
    def expand_all_items(self):
        """Expand all items in the tree by default."""
        def expand_children(item_id):
            self.tree.item(item_id, open=True)
            for child_id in self.tree.get_children(item_id):
                expand_children(child_id)
        for item_id in self.tree.get_children():
            expand_children(item_id)
    
    def add_item_to_tree(self, item, parent_id):
        """Recursively add items to treeview."""
        display_name = f"{'  ' * item.indent_level}{item.name}"
        enabled_text = "✓" if item.enabled.get() else "✗"
        
        predecessor_text = ""
        predecessor_type_text = ""
        if item.predecessor_id and item.predecessor_id in self.item_id_map:
            pred_item = self.item_id_map[item.predecessor_id]
            lag_str = f" +{item.lag}d" if item.lag > 0 else f" {item.lag}d" if item.lag < 0 else ""
            predecessor_text = f"{pred_item.name[:15]}{lag_str}"
            predecessor_type_text = item.predecessor_type
            
        item_id = self.tree.insert(parent_id, 'end', text=display_name, 
                                     values=(predecessor_text, predecessor_type_text, enabled_text, item.duration, f"${item.price:,}", item.start_date, item.end_date),
                                     tags=('milestone' if item.is_milestone else 'task',))
        self.tree_item_map[item_id] = item
        for child in item.children:
            self.add_item_to_tree(child, item_id)
        return item_id

    def on_item_double_click(self, event):
        """Handle double-click for inline editing or opening predecessor dialog."""
        item_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        if not item_id or not column_id: return
        item = self.tree_item_map.get(item_id)
        if not item: return

        display_cols = list(self.tree['displaycolumns'])
        
        if column_id == f"#{display_cols.index('Enabled') + 1}":
            item.enabled.set(not item.enabled.get())
            self.update_item_display(item_id, item)
        elif not item.is_milestone:
            if column_id == f"#{display_cols.index('Duration') + 1}": self.edit_cell(item_id, item, 'duration', column_id)
            elif column_id == f"#{display_cols.index('Price') + 1}": self.edit_cell(item_id, item, 'price', column_id)
            elif column_id == f"#{display_cols.index('Predecessor') + 1}": self.edit_predecessor(item_id)
            elif column_id == f"#{display_cols.index('Start Date') + 1}": self.edit_cell(item_id, item, 'start_date', column_id)

    def edit_cell(self, item_id, item, attribute, column_id):
        """Create inline editor for a cell."""
        if self.current_editor: self.current_editor.destroy()
        bbox = self.tree.bbox(item_id, column_id)
        if not bbox: return
        x, y, w, h = bbox
        
        current_value = getattr(item, attribute)
        entry = tk.Entry(self.tree, font=('Arial', 9))
        self.current_editor = entry
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, str(current_value))
        entry.select_range(0, tk.END)
        entry.focus()
        
        def save_edit(event=None):
            try:
                new_value = entry.get()
                if attribute == 'duration':
                    # --- MODIFICATION: Round duration up ---
                    setattr(item, attribute, math.ceil(float(new_value)) if new_value else 0)
                elif attribute == 'price':
                    setattr(item, attribute, int(new_value.replace('$', '').replace(',', '')) if new_value else 0)
                else:
                    setattr(item, attribute, new_value)
                self.update_item_display(item_id, item)
            except (ValueError, tk.TclError): pass
            finally:
                if entry: entry.destroy()
                self.current_editor = None
        
        entry.bind('<Return>', save_edit)
        entry.bind('<KP_Enter>', save_edit)
        entry.bind('<Escape>', lambda e: entry.destroy())
        entry.bind('<FocusOut>', save_edit)

    def edit_type_cell(self, item_id, item, column_id):
        """Create an inline dropdown for the predecessor type."""
        if self.current_editor: self.current_editor.destroy()
        bbox = self.tree.bbox(item_id, column_id)
        if not bbox: return
        x, y, w, h = bbox

        combo = ttk.Combobox(self.tree, values=['FS', 'SS', 'FF', 'SF'], state='readonly', font=('Arial', 9))
        self.current_editor = combo
        combo.place(x=x, y=y, width=w, height=h)
        combo.set(item.predecessor_type)
        combo.focus()

        def save_type_edit(event=None):
            try:
                new_value = combo.get()
                if new_value:
                    item.predecessor_type = new_value
                self.update_item_display(item_id, item)
            except (ValueError, tk.TclError): pass
            finally:
                if combo: combo.destroy()
                self.current_editor = None
        
        combo.bind('<<ComboboxSelected>>', save_type_edit)
        combo.bind('<FocusOut>', save_type_edit)

    def edit_predecessor(self, item_id):
        """Open a dialog to set an item's predecessor."""
        item_to_edit = self.tree_item_map.get(item_id)
        if not item_to_edit: return

        dialog = tk.Toplevel(self.root)
        dialog.title(f"Set Predecessor for '{item_to_edit.name}'")
        dialog.geometry("450x200")
        dialog.transient(self.root)
        dialog.grab_set()

        possible_preds = {i.name: i.id for i in self.item_id_map.values() if not i.is_milestone and i.id != item_to_edit.id}
        
        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Predecessor Task:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        pred_var = tk.StringVar()
        pred_combo = ttk.Combobox(frame, textvariable=pred_var, values=list(possible_preds.keys()), width=40)
        pred_combo.grid(row=0, column=1, columnspan=2, padx=5, pady=5)
        
        ttk.Label(frame, text="Lag (days):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        lag_var = tk.IntVar(value=item_to_edit.lag)
        lag_entry = ttk.Entry(frame, textvariable=lag_var, width=8)
        lag_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)

        if item_to_edit.predecessor_id:
            for name, p_id in possible_preds.items():
                if p_id == item_to_edit.predecessor_id:
                    pred_combo.set(name)
                    break
        
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=2, column=0, columnspan=3, pady=20)

        def save_predecessor():
            selected_name = pred_var.get()
            if selected_name in possible_preds:
                item_to_edit.predecessor_id = possible_preds[selected_name]
                item_to_edit.lag = lag_var.get()
            self.update_item_display(item_id, item_to_edit)
            self.highlight_dependencies(item_id)
            dialog.destroy()
        
        def clear_predecessor():
            item_to_edit.predecessor_id = None
            self.update_item_display(item_id, item_to_edit)
            self.highlight_dependencies(item_id)
            dialog.destroy()

        ttk.Button(button_frame, text="Save", command=save_predecessor).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Clear", command=clear_predecessor).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=10)
    
    def update_item_display(self, item_id, item):
        """Update a single item's display without refreshing entire tree."""
        enabled_text = "✓" if item.enabled.get() else "✗"
        predecessor_text = ""
        predecessor_type_text = ""
        if item.predecessor_id and item.predecessor_id in self.item_id_map:
            pred_item = self.item_id_map[item.predecessor_id]
            lag_str = f" +{item.lag}d" if item.lag > 0 else f" {item.lag}d" if item.lag < 0 else ""
            predecessor_text = f"{pred_item.name[:15]}{lag_str}"
            predecessor_type_text = item.predecessor_type
        
        value_map = {
            'Predecessor': predecessor_text,
            'Type': predecessor_type_text,
            'Enabled': enabled_text,
            'Duration': item.duration,
            'Price': f"${item.price:,}",
            'Start Date': item.start_date,
            'End Date': item.end_date
        }
        
        display_cols = self.tree['displaycolumns']
        values_tuple = tuple(value_map[col_id] for col_id in display_cols)

        if self.tree.exists(item_id):
            self.tree.item(item_id, values=values_tuple)
    
    def add_custom_item(self):
        """Add a custom item to the project."""
        selection = self.tree.selection()
        parent_item = self.tree_item_map.get(selection[0]) if selection else None
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Custom Item")
        dialog.geometry("400x280") # Increased height for new checkbox
        
        ttk.Label(dialog, text="Name:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        name_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=name_var, width=40).grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(dialog, text="Duration (days):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        duration_var = tk.DoubleVar()
        ttk.Entry(dialog, textvariable=duration_var, width=20).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Label(dialog, text="Price ($):").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        price_var = tk.IntVar()
        ttk.Entry(dialog, textvariable=price_var, width=20).grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Label(dialog, text="Is Milestone:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        is_milestone_var = tk.BooleanVar()
        ttk.Checkbutton(dialog, variable=is_milestone_var).grid(row=3, column=1, sticky=tk.W, padx=5, pady=5)

        # --- MODIFICATION: Add "New Section" checkbox ---
        is_new_section_var = tk.BooleanVar()
        ttk.Checkbutton(dialog, text="Add as new section", variable=is_new_section_var).grid(row=4, column=1, sticky=tk.W, padx=5, pady=5)
        
        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=5, column=0, columnspan=2, pady=20)
        
        def add_item():
            is_new_section = is_new_section_var.get()
            
            # --- MODIFICATION: Round duration up ---
            duration = math.ceil(duration_var.get())
            
            if is_new_section:
                # Add as a new top-level section
                new_item = ProposalItem(name_var.get(), duration, price_var.get(), "", True, 0)
                self.template_items.append(new_item)
            else:
                # Add as a child of the selected item (or as a top-level item if nothing is selected)
                indent_level = parent_item.indent_level + 1 if parent_item else 0
                new_item = ProposalItem(name_var.get(), duration, price_var.get(), "", is_milestone_var.get(), indent_level)
                if parent_item:
                    new_item.parent = parent_item
                    parent_item.children.append(new_item)
                else:
                    self.template_items.append(new_item)

            self.populate_tree()
            self.expand_all_items()
            dialog.destroy()
        
        ttk.Button(button_frame, text="Add", command=add_item).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def delete_item(self):
        """Delete selected item."""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select an item to delete.")
            return
        item_id = selection[0]
        item = self.tree_item_map.get(item_id)
        if not item: return
        
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{item.name}' and all its children?"):
            if item.parent: item.parent.children.remove(item)
            else: self.template_items.remove(item)
            self.populate_tree()
            self.expand_all_items()
    
    def _add_business_days(self, start_date_str, days_to_add):
        """Adds or subtracts business days from a given date string."""
        if not start_date_str: return ""
        try:
            current_date = datetime.strptime(start_date_str, "%m/%d/%y")
        except ValueError:
            return ""

        step = timedelta(days=1) if days_to_add >= 0 else timedelta(days=-1)
        days_counted = 0
        
        days_to_add = int(days_to_add)

        if days_to_add > 0: days_to_add -= 1

        while current_date.weekday() >= 5:
            current_date += timedelta(days=1)

        while days_counted < abs(days_to_add):
            current_date += step
            if current_date.weekday() < 5:
                days_counted += 1
        
        while current_date.weekday() >= 5:
            current_date += step

        return current_date.strftime("%m/%d/%y")

    def calculate_all_dates(self):
        """Calculate all dates based on dependencies and durations using topological sort."""
        self.populate_tree()
        all_tasks = [item for item in self.item_id_map.values() if item.enabled.get() and not item.is_milestone]

        graph = {item.id: [] for item in all_tasks}
        in_degree = {item.id: 0 for item in all_tasks}
        for item in all_tasks:
            if item.predecessor_id and item.predecessor_id in self.item_id_map:
                pred = self.item_id_map[item.predecessor_id]
                if pred.enabled.get():
                    graph[item.predecessor_id].append(item.id)
                    in_degree[item.id] += 1

        queue = [item_id for item_id in in_degree if in_degree[item_id] == 0]
        sorted_order = []
        while queue:
            u_id = queue.pop(0)
            sorted_order.append(u_id)
            for v_id in graph.get(u_id, []):
                in_degree[v_id] -= 1
                if in_degree[v_id] == 0:
                    queue.append(v_id)

        if len(sorted_order) != len(all_tasks):
            messagebox.showerror("Calculation Error", "A circular dependency (e.g., Task A -> B -> A) was detected. Please fix the predecessors and try again.")
            return

        project_start = self.project_start_date.get()
        for item_id in sorted_order:
            item = self.item_id_map[item_id]
            pred_item = self.item_id_map.get(item.predecessor_id) if item.predecessor_id else None

            if pred_item and pred_item.enabled.get() and pred_item.end_date:
                if item.predecessor_type == 'FS':
                    item.start_date = self._add_business_days(pred_item.end_date, item.lag + 1)
                elif item.predecessor_type == 'SS':
                    item.start_date = self._add_business_days(pred_item.start_date, item.lag)
                elif item.predecessor_type == 'FF':
                    finish_date = self._add_business_days(pred_item.end_date, item.lag)
                    item.start_date = self._add_business_days(finish_date, -item.duration + 1)
                elif item.predecessor_type == 'SF':
                    finish_date = self._add_business_days(pred_item.start_date, item.lag)
                    item.start_date = self._add_business_days(finish_date, -item.duration + 1)
            else:
                item.start_date = item.start_date or project_start
            
            item.end_date = self._add_business_days(item.start_date, item.duration)

        def calculate_milestone_rollup(items):
            for item in items:
                if item.enabled.get() and item.is_milestone and item.children:
                    calculate_milestone_rollup(item.children)
                    enabled_children = [c for c in item.children if c.enabled.get()]
                    if enabled_children:
                        valid_starts = [datetime.strptime(c.start_date, "%m/%d/%y") for c in enabled_children if c.start_date]
                        valid_ends = [datetime.strptime(c.end_date, "%m/%d/%y") for c in enabled_children if c.end_date]
                        if valid_starts: item.start_date = min(valid_starts).strftime("%m/%d/%y")
                        if valid_ends: item.end_date = max(valid_ends).strftime("%m/%d/%y")
                        item.duration = sum(c.duration for c in enabled_children)
                        item.price = sum(c.price for c in enabled_children)
        calculate_milestone_rollup(self.template_items)

        self.populate_tree()
        self.expand_all_items()

    def generate_pdf(self):
        """Generate a single PDF proposal with the Gantt chart on page 2."""
        filename = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save Proposal As"
        )
        if not filename: return
            
        try:
            self.create_pdf(filename)
            messagebox.showinfo("Success", f"Successfully generated proposal:\n{os.path.basename(filename)}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate PDF: {str(e)}")

    def create_gantt_chart_image(self):
        """Generates a Gantt chart and returns it as an in-memory image buffer."""
        tasks = []
        def collect_tasks(items, parent_name=""):
            for item in items:
                if item.enabled.get():
                    if not item.is_milestone and item.start_date and item.end_date and item.duration > 0:
                        try:
                            start = datetime.strptime(item.start_date, "%m/%d/%y")
                            end = datetime.strptime(item.end_date, "%m/%d/%y")
                            # Add timedelta(days=1) to the end date so the bar visually covers the entire last day.
                            tasks.append({
                                "name": f"{' ' * item.indent_level * 2}{item.name}", 
                                "start": start, 
                                "end": end + timedelta(days=1)
                            })
                        except ValueError: 
                            continue # Skip tasks with invalid date formats
                    collect_tasks(item.children, item.name if item.indent_level == 0 else parent_name)
        
        collect_tasks(self.template_items)
        if not tasks: 
            return None

        tasks.sort(key=lambda x: x["start"])
        
        fig, ax = plt.subplots(figsize=(10.5, 7.5))
        y_labels = [task["name"] for task in tasks]
        
        for i, task in enumerate(tasks):
            start_date = mdates.date2num(task["start"])
            end_date = mdates.date2num(task["end"])
            duration = end_date - start_date
            ax.barh(i, duration, left=start_date, height=0.6, align='center', color='#991f2b', edgecolor='black')

        ax.set_yticks(range(len(y_labels)))
        ax.set_yticklabels(y_labels, fontsize=8, fontfamily='sans-serif')
        ax.invert_yaxis()

        # --- Meticulous X-axis Timeline Configuration ---
        ax.set_xlabel('Date', fontsize=11, fontfamily='sans-serif')
        
        # Set major ticks to display the first of every month.
        ax.xaxis.set_major_locator(mdates.MonthLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %Y'))
        
        # Set minor ticks to display every Monday.
        ax.xaxis.set_minor_locator(mdates.WeekdayLocator(byweekday=mdates.MO))
        
        # Add a grid for both major (monthly) and minor (weekly) ticks for better readability.
        ax.grid(True, which='major', axis='x', linestyle='-', color='gray', linewidth=0.7)
        ax.grid(True, which='minor', axis='x', linestyle='--', color='lightgray', linewidth=0.5)

        # Set the x-axis limits to show a bit of padding before the first task and after the last one.
        if tasks:
            min_date = min(task['start'] for task in tasks)
            max_date = max(task['end'] for task in tasks)
            date_padding = timedelta(days=15) # Add ~2 weeks of padding on each side
            ax.set_xlim(min_date - date_padding, max_date + date_padding)

        # Rotate date labels for a better fit and adjust bottom margin.
        fig.autofmt_xdate(rotation=45, ha='right')
        fig.subplots_adjust(left=0.3, right=0.95, top=0.95, bottom=0.15)
        
        buf = io.BytesIO()
        plt.savefig(buf, format='png', dpi=300)
        plt.close(fig)
        buf.seek(0)
        return buf

    def create_pdf(self, filename):
        """
        MODIFIED: Create the multi-page PDF document with dynamic sizing for the table.
        """
        
        font_name = 'Jost'
        font_name_bold = 'Jost-Bold'
        
        try:
            jost_regular_path = resource_path('Jost-Regular.ttf')
            jost_bold_path = resource_path('Jost-Bold.ttf')
            pdfmetrics.registerFont(TTFont(font_name, jost_regular_path))
            pdfmetrics.registerFont(TTFont(font_name_bold, jost_bold_path))
        except Exception as e:
            print(f"Could not load custom fonts, falling back to Helvetica. Error: {e}")
            font_name = 'Helvetica'
            font_name_bold = 'Helvetica-Bold'
        
        doc = BaseDocTemplate(filename, topMargin=0.5*inch, bottomMargin=0.4*inch, leftMargin=0.3*inch, rightMargin=0.3*inch)
        
        portrait_frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id='portrait_frame')
        l_width, l_height = landscape(letter)
        landscape_frame = Frame(doc.leftMargin, doc.bottomMargin, l_width - doc.leftMargin - doc.rightMargin, l_height - doc.bottomMargin - doc.topMargin, id='landscape_frame')

        doc.addPageTemplates([
            PageTemplate(id='PortraitPage', frames=[portrait_frame], pagesize=letter),
            PageTemplate(id='LandscapePage', frames=[landscape_frame], pagesize=landscape(letter)),
        ])

        elements = []
        styles = getSampleStyleSheet()
        
        # --- DYNAMIC SIZING LOGIC ---
        
        # 1. First, build the table data to count the rows
        all_table_data = []
        
        # Header and Summary rows are always present
        header_row = ['Project Milestones', 'Days', 'Start', 'Finish', 'Price']
        summary_row_content = [self.project_name.get(), '', '', '', '']
        all_table_data.append(header_row)
        all_table_data.append(summary_row_content)

        def count_enabled_items(items):
            count = 0
            for item in items:
                if item.enabled.get():
                    count += 1
                    if item.children:
                        count += count_enabled_items(item.children)
            return count

        num_rows = count_enabled_items(self.template_items) + 2 # Add 2 for header and summary

        # 2. Choose style parameters based on the number of rows
        if num_rows <= 35: # Large size
            font_size = 9
            leading = 11
            header_font_size = 10
            header_leading = 13
            col_widths = [3.5*inch, 0.8*inch, 1.0*inch, 1.0*inch, 1.6*inch]
            header_padding = 5
            row_padding = 3
        elif num_rows <= 45: # Normal size
            font_size = 8
            leading = 10
            header_font_size = 9
            header_leading = 12
            col_widths = [3.4*inch, 0.7*inch, 1.0*inch, 1.0*inch, 1.3*inch]
            header_padding = 4
            row_padding = 2
        else: # Smallest size
            font_size = 6.5
            leading = 8
            header_font_size = 8
            header_leading = 11
            col_widths = [3.3*inch, 0.7*inch, 1.0*inch, 1.0*inch, 1.2*inch]
            header_padding = 3
            row_padding = 1

        # 3. Define ParagraphStyles using the dynamic sizes
        header_project_style = ParagraphStyle('header_project_style', parent=styles['Normal'], fontName=font_name_bold, fontSize=14, alignment=0)
        
        table_text_style = ParagraphStyle('table_text_style', parent=styles['Normal'], fontName=font_name, fontSize=font_size, leading=leading, alignment=0)
        table_bold_style = ParagraphStyle('table_bold_style', parent=styles['Normal'], fontName=font_name_bold, fontSize=font_size, leading=leading, alignment=0)
        table_bold_white_style = ParagraphStyle('table_bold_white_style', parent=styles['Normal'], fontName=font_name_bold, fontSize=font_size, leading=leading, textColor=colors.white, alignment=0)
        
        table_header_style_left = ParagraphStyle('table_header_style_left', parent=styles['Normal'], fontName=font_name_bold, fontSize=header_font_size, leading=header_leading, alignment=0, textColor=colors.white)
        table_header_style_right = ParagraphStyle('table_header_style_right', parent=table_header_style_left, alignment=2)
        
        # --- END DYNAMIC SIZING LOGIC ---

        logo_path_val = self.logo_path.get()
        logo = None
        if logo_path_val and os.path.exists(logo_path_val):
            try:
                logo = Image(logo_path_val, width=2.0*inch, height=1.0*inch, kind='proportional')
                logo.hAlign = 'RIGHT'
            except Exception as e:
                print(f"Error loading company logo: {e}")
        
        client_logo_path_val = self.client_logo_path.get()
        client_logo = None
        if client_logo_path_val and os.path.exists(client_logo_path_val):
            try:
                client_logo = Image(client_logo_path_val, width=2.0*inch, height=1.0*inch, kind='proportional')
                client_logo.hAlign = 'CENTER'
            except Exception as e:
                print(f"Error loading client logo: {e}")
        else:
            client_logo = Paragraph("", ParagraphStyle('placeholder', alignment=1))

        header_left_text = f"<font color='#991f2b'>{self.company_name.get()}<br/>{self.project_name.get()}</font>"
        header_left_para = Paragraph(header_left_text, header_project_style)
        
        header_table = Table([[header_left_para, client_logo, logo]], colWidths=[3.0*inch, 2.9*inch, 2.0*inch])
        header_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (0, 0), 'LEFT'),
            ('ALIGN', (1, 0), (1, 0), 'CENTER'),
            ('ALIGN', (2, 0), (2, 0), 'RIGHT'),
            ('TOPPADDING', (0, 0), (-1, -1), 15),
        ]))
        
        elements.append(header_table)
        elements.append(Spacer(1, 0.2*inch))

        # 4. Now, rebuild the table data with formatted Paragraph objects
        all_table_data = [] # Clear the list
        
        table_header_row_formatted = [
            Paragraph('Project Milestones', table_header_style_left),
            Paragraph('Days', table_header_style_left),
            Paragraph('Start', table_header_style_left),
            Paragraph('Finish', table_header_style_left),
            Paragraph('Price', table_header_style_right)
        ]
        all_table_data.append(table_header_row_formatted)

        total_price = sum(item.price for item in self.template_items if item.enabled.get() and item.indent_level == 0)
        valid_dates = [datetime.strptime(dt, "%m/%d/%y") for item in self.template_items if item.enabled.get() for dt in (item.start_date, item.end_date) if dt]
        earliest_start = min(valid_dates).strftime("%m/%d/%y") if valid_dates else ""
        latest_end = max(valid_dates).strftime("%m/%d/%y") if valid_dates else ""
        total_duration = 0
        if earliest_start and latest_end:
            try:
                start_dt, end_dt = datetime.strptime(earliest_start, "%m/%d/%y"), datetime.strptime(latest_end, "%m/%d/%y")
                total_duration = sum(1 for d in range((end_dt - start_dt).days + 1) if (start_dt + timedelta(days=d)).weekday() < 5)
            except ValueError: total_duration = 0

        summary_row_formatted = [
            Paragraph(f"<b>{self.project_name.get()}</b>", table_bold_white_style),
            Paragraph(f"{total_duration}", table_bold_white_style),
            Paragraph(earliest_start, table_bold_white_style),
            Paragraph(latest_end, table_bold_white_style),
            Paragraph(f"${total_price:,}", ParagraphStyle('summary_price', parent=table_bold_white_style, alignment=2)),
        ]
        all_table_data.append(summary_row_formatted)

        def build_table_rows_recursive(items):
            for item in items:
                if item.enabled.get():
                    is_main_milestone = item.is_milestone and item.indent_level == 0
                    
                    if is_main_milestone:
                         current_style = table_bold_white_style
                    elif item.is_milestone:
                         current_style = table_bold_style
                    else:
                         current_style = table_text_style

                    price_style = ParagraphStyle('price_style', parent=current_style, alignment=2)
                    name_para_style = current_style

                    if is_main_milestone:
                        name_text = f"<b>{'&nbsp;' * 4 * item.indent_level}{item.name}</b>"
                        name_para_style = ParagraphStyle('main_milestone_name', parent=table_bold_white_style)
                    elif item.is_milestone:
                        name_text = f"<b>{'&nbsp;' * 4 * item.indent_level}{item.name}</b>"
                        name_para_style = ParagraphStyle('sub_milestone_name', parent=table_bold_style)
                    else:
                        name_text = f"{'&nbsp;' * 4 * item.indent_level}{item.name}"

                    name_para = Paragraph(name_text, name_para_style)

                    row_data = [
                        name_para,
                        Paragraph(f"{item.duration}", current_style),
                        Paragraph(item.start_date, current_style),
                        Paragraph(item.end_date, current_style),
                        Paragraph(f"${item.price:,}" if item.price > 0 else ("$0" if item.is_milestone else ""), price_style),
                    ]
                    all_table_data.append(row_data)
                    
                    if item.children:
                        build_table_rows_recursive(item.children)
        
        build_table_rows_recursive(self.template_items)

        full_table = Table(all_table_data, colWidths=col_widths, repeatRows=1)
        
        table_style_commands = [
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,0), header_padding), 
            ('BOTTOMPADDING', (0,0), (-1,0), header_padding),
            ('TOPPADDING', (0,1), (-1,-1), row_padding),
            ('BOTTOMPADDING', (0,1), (-1,-1), row_padding),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#991f2b")),
            ('LINEBELOW', (0, 0), (-1, 0), 0.5, colors.black),
            ('BACKGROUND', (0, 1), (-1, 1), colors.black),
            ('LINEBELOW', (0, 1), (-1, 1), 0.5, colors.black),
            ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ]
        
        row_idx_offset = 2
        
        def find_and_style_milestones(items, current_row_idx):
            for item in items:
                if item.enabled.get():
                    if item.is_milestone:
                        if item.indent_level == 0:
                            bg_color = colors.HexColor("#991f2b")
                        else:
                            bg_color = colors.HexColor("#D3D3D3")
                        
                        table_style_commands.extend([
                            ('BACKGROUND', (0, current_row_idx), (-1, current_row_idx), bg_color),
                        ])
                    current_row_idx += 1
                    if item.children:
                        current_row_idx = find_and_style_milestones(item.children, current_row_idx)
            return current_row_idx

        find_and_style_milestones(self.template_items, row_idx_offset)
        
        full_table.setStyle(TableStyle(table_style_commands))
        elements.append(full_table)
        
        if self.include_gantt.get():
            gantt_chart_buffer = self.create_gantt_chart_image()
            if gantt_chart_buffer:
                elements.append(NextPageTemplate('LandscapePage'))
                elements.append(PageBreak())
                
                schedule_title = Paragraph("Project Schedule", ParagraphStyle('gantt_title', parent=styles['h2'], alignment=1))
                elements.append(schedule_title)
                elements.append(Spacer(1, 0.1*inch))
                
                gantt_image = Image(gantt_chart_buffer, width=9*inch, height=6.5*inch, kind='proportional')
                gantt_image.hAlign = 'CENTER'
                elements.append(gantt_image)

        doc.build(elements)
    
    def save_template(self):
        """Save current template to a JSON file."""
        filename = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")], title="Save Template As")
        if not filename: return
        try:
            # --- MODIFICATION: Save a placeholder for the default logo ---
            logo_path_to_save = self.logo_path.get()
            if logo_path_to_save == self.default_logo_path:
                logo_path_to_save = "DEFAULT_LOGO"

            template_data = {
                'project_name': self.project_name.get(), 'company_name': self.company_name.get(),
                'project_start_date': self.project_start_date.get(), 
                'logo_path': logo_path_to_save,
                'client_logo_path': self.client_logo_path.get(),
                'items': self.serialize_items(self.template_items)
            }
            with open(filename, 'w') as f:
                json.dump(template_data, f, indent=2)
            messagebox.showinfo("Success", "Template saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save template: {str(e)}")
    
    def serialize_items(self, items):
        """Convert item objects to a serializable dictionary format."""
        return [{
            'name': item.name, 'duration': item.duration, 'price': item.price,
            'start_date': item.start_date, 'end_date': item.end_date,
            'is_milestone': item.is_milestone, 'indent_level': item.indent_level,
            'enabled': item.enabled.get(), 'id': item.id,
            'predecessor_id': item.predecessor_id, 'predecessor_type': item.predecessor_type,
            'lag': item.lag, 'children': self.serialize_items(item.children)
        } for item in items]
    
    def load_template(self):
        """Load a template from a JSON file."""
        filename = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")], title="Load Template")
        if not filename: return
        try:
            with open(filename, 'r') as f:
                template_data = json.load(f)
            
            self.project_name.set(template_data.get('project_name', ''))
            self.company_name.set(template_data.get('company_name', ''))
            self.project_start_date.set(template_data.get('project_start_date', ''))
            
            # --- MODIFICATION: Handle the logo path correctly ---
            saved_logo_path = template_data.get('logo_path', '')
            if saved_logo_path == "DEFAULT_LOGO" or not os.path.exists(saved_logo_path):
                self.logo_path.set(self.default_logo_path)
            else:
                self.logo_path.set(saved_logo_path)

            self.client_logo_path.set(template_data.get('client_logo_path', ''))
            # --- END MODIFICATION ---

            self.template_items = self.deserialize_items(template_data.get('items', []))
            self.populate_tree()
            self.expand_all_items()
            messagebox.showinfo("Success", "Template loaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load template: {str(e)}")
    
    def deserialize_items(self, items_data):
        """Convert serialized dictionary data back into ProposalItem objects."""
        items = []
        for item_data in items_data:
            # --- MODIFICATION: Round duration up from float to int ---
            duration = math.ceil(float(item_data.get('duration', 0)))
            item = ProposalItem(item_data['name'], duration, item_data['price'],
                                  item_data['start_date'], item_data['is_milestone'], item_data['indent_level'])
            item.end_date = item_data.get('end_date', '')
            item.enabled.set(item_data.get('enabled', True))
            item.id = item_data.get('id', str(uuid.uuid4()))
            item.predecessor_id = item_data.get('predecessor_id')
            item.predecessor_type = item_data.get('predecessor_type', 'FS')
            item.lag = item_data.get('lag', 0)
            item.children = self.deserialize_items(item_data.get('children', []))
            for child in item.children:
                child.parent = item
            items.append(item)
        return items

def main():
    """Main function to run the application."""
    root = tk.Tk()
    app = ProposalGenerator(root)
    root.mainloop()

if __name__ == "__main__":
    main()
