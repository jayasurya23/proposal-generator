import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from proposal_generator import ProposalGenerator
from schedule_parser import build_model_rows, flatten_to_template_rows, push_into_generator
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Proposal Builder (Project Schedule Import)")
        self.geometry("980x720")

        # Backing engine/UI
        self.pg = ProposalGenerator(self)
        self.pg.root = self

        # State
        self.xlsx_path = tk.StringVar(value="")
        self.price_source = tk.StringVar(value="proposal")  # or "detail"

        # Top controls
        ctrl = ttk.Frame(self)
        ctrl.pack(side=tk.TOP, fill=tk.X, padx=8, pady=8)

        ttk.Label(ctrl, text="Excel (Project Schedule):").grid(row=0, column=0, sticky="w")
        ttk.Entry(ctrl, textvariable=self.xlsx_path, width=60).grid(row=0, column=1, sticky="we", padx=6)
        ttk.Button(ctrl, text="Upload Project Schedule", command=self.on_upload_xlsx).grid(row=0, column=2, padx=6)

        ttk.Label(ctrl, text="Pricing:").grid(row=1, column=0, sticky="w")
        ttk.Radiobutton(ctrl, text="Proposal Page", variable=self.price_source, value="proposal").grid(row=1, column=1, sticky="w")
        ttk.Radiobutton(ctrl, text="Detail (Civil/Elect/Structural)", variable=self.price_source, value="detail").grid(row=1, column=1, sticky="e")



        ctrl.columnconfigure(1, weight=1)

    # ---- UI handlers ----

    def on_upload_xlsx(self):
        path = filedialog.askopenfilename(
            title="Select Project Schedule (Excel)",
            filetypes=[
                ("Excel Macro-Enabled", "*.xlsm"),
                ("Excel Workbook", "*.xlsx"),
            ],
        )
        if path:
            self.xlsx_path.set(path)
            # Auto-parse and overhaul immediately after upload
            self.on_parse_and_populate()

    def on_parse_and_populate(self):
        xlsx = self.xlsx_path.get().strip()
        if not xlsx or not os.path.exists(xlsx):
            messagebox.showerror("Missing file", "Please choose a Project Schedule workbook (.xlsm or .xlsx).")
            return

        try:
            buckets, info = build_model_rows(xlsx)
            review_pairs = {("Civil", "30%"), ("Civil", "60%"), ("Electrical", "30%"), ("Electrical", "60%")}

            rows_out = flatten_to_template_rows(
                buckets=buckets,
                hours_per_day=8.0,
                price_source=self.price_source.get().strip().lower(),
                review_pairs=review_pairs,
            )

            push_into_generator(self.pg, info, rows_out)
            self.pg.calculate_all_dates()

            messagebox.showinfo("Success", "Project Schedule parsed and task tree replaced.")
        except Exception as e:
            messagebox.showerror("Error parsing workbook", str(e))


if __name__ == "__main__":
    App().mainloop()
