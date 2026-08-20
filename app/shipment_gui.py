
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import threading
import traceback

from shipment_processor import process, ValidationFailure

class App:
    def __init__(self, root):
        self.root = root
        root.title("Shipment Template Automation")
        root.geometry("780x520")
        root.minsize(720, 480)

        self.sources = []
        self.template = ""
        self.output = ""

        tk.Label(root, text="Shipment Template Automation",
                 font=("Segoe UI", 18, "bold")).pack(pady=(18, 4))
        tk.Label(root, text="Portable Python • Microsoft Excel • Final .xls",
                 font=("Segoe UI", 10)).pack(pady=(0, 15))

        self.src_label = tk.Label(root, text="No source files selected", anchor="w")
        self.src_label.pack(fill="x", padx=30, pady=4)
        tk.Button(root, text="1. Select Multiple Source Excel Files",
                  command=self.select_sources, height=2).pack(fill="x", padx=30)

        self.tpl_label = tk.Label(root, text="No template selected", anchor="w")
        self.tpl_label.pack(fill="x", padx=30, pady=4)
        tk.Button(root, text="2. Select Output Template",
                  command=self.select_template, height=2).pack(fill="x", padx=30)

        self.out_label = tk.Label(root, text="3. Output folder: choose folder",
                                  anchor="w")
        self.out_label.pack(fill="x", padx=30, pady=4)
        tk.Button(root, text="3. Choose Output Folder",
                  command=self.select_output, height=2).pack(fill="x", padx=30)

        self.run_btn = tk.Button(root, text="4. RUN",
                                 command=self.run, height=2,
                                 font=("Segoe UI", 11, "bold"))
        self.run_btn.pack(fill="x", padx=30, pady=(14, 8))

        self.status = tk.StringVar(value="Ready.")
        tk.Label(root, textvariable=self.status, relief="sunken",
                 anchor="w").pack(fill="x", padx=30, pady=4)

        self.logbox = tk.Text(root, height=11, state="disabled")
        self.logbox.pack(fill="both", expand=True, padx=30, pady=(4, 20))

    def log(self, msg):
        self.logbox.config(state="normal")
        self.logbox.insert("end", msg + "\n")
        self.logbox.see("end")
        self.logbox.config(state="disabled")
        self.status.set(msg)

    def select_sources(self):
        p = filedialog.askopenfilenames(
            title="Select source files",
            filetypes=[("Excel", "*.xls *.xlsx *.xlsm"), ("All", "*.*")]
        )
        if p:
            self.sources = list(p)
            self.src_label.config(text=f"{len(p)} source file(s) selected")
            self.log("Source files selected.")

    def select_template(self):
        p = filedialog.askopenfilename(
            title="Select output template",
            filetypes=[("Excel", "*.xls *.xlsx *.xlsm"), ("All", "*.*")]
        )
        if p:
            self.template = p
            self.tpl_label.config(text=f"Template: {Path(p).name}")
            self.log("Template selected.")

    def select_output(self):
        p = filedialog.askdirectory(title="Choose output folder")
        if p:
            self.output = p
            self.out_label.config(text=f"Output: {p}")

    def run(self):
        if not self.sources:
            messagebox.showerror("Missing", "Select source Excel files.")
            return
        if not self.template:
            messagebox.showerror("Missing", "Select the template.")
            return
        if not self.output:
            self.output = str(Path(self.template).parent)

        self.run_btn.config(state="disabled")
        self.logbox.config(state="normal")
        self.logbox.delete("1.0", "end")
        self.logbox.config(state="disabled")
        self.log("Starting...")
        threading.Thread(target=self.worker, daemon=True).start()

    def worker(self):
        try:
            result = process(self.sources, self.template, self.output, self.log)
            self.root.after(0, lambda: messagebox.showinfo(
                "Completed", f"Completed successfully.\n\nFinal .xls:\n{result}"
            ))
        except ValidationFailure as e:
            details = "\n".join(
                f"{x['file']} row {x['row']}: {x['reason']} | {x['contents']}"
                for x in e.errors[:20]
            )
            self.root.after(0, lambda: messagebox.showerror(
                "Validation failed",
                "No output was created.\n\n" + details
            ))
            self.root.after(0, lambda: self.status.set("Validation failed."))
        except Exception as e:
            detail = traceback.format_exc()
            self.log(detail)
            self.root.after(0, lambda: messagebox.showerror(
                "Processing failed", str(e)
            ))
            self.root.after(0, lambda: self.status.set("Processing failed."))
        finally:
            self.root.after(0, lambda: self.run_btn.config(state="normal"))

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
