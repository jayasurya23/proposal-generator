import tkinter as tk

class ProposalItem:
    """Represents a single task or milestone in the project."""
    def __init__(self, name, duration=0, price=0, start_date="", is_milestone=False, indent_level=0, item_id=None):
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
        self.id = item_id  # New sequential integer ID
        self.predecessor_id = None
        self.predecessor_type = 'FS'  # Finish-to-Start
        self.lag = 0  # Lag in days
        # --- MODIFICATION: Add flag for manually set start dates ---
        self.is_start_pinned = False
