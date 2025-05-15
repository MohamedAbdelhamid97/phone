import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

class PhoneMatcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Phone Number Matcher")
        self.root.geometry("600x500")  # Set larger window size
        self.root.resizable(True, True)

        self.file1_path = ""
        self.file2_path = ""
        self.sheet_vars = {}
        self.sheets_to_process = []

        self.build_gui()

    def build_gui(self):
        # Title
        title = tk.Label(self.root, text="Excel Phone Number Matcher", font=("Arial", 16, "bold"))
        title.pack(pady=10)

        # Buttons for file selection
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=5)

        tk.Button(btn_frame, text="📂 Select File 1 (Phone List)", command=self.load_file1, width=40).pack(pady=5)
        tk.Button(btn_frame, text="📂 Select File 2 (Sheets to Match)", command=self.load_file2, width=40).pack(pady=5)

        # Sheet selection with scrollable frame
        sheet_label = tk.Label(self.root, text="✅ Select Sheets to Process:", font=("Arial", 12, "bold"))
        sheet_label.pack(pady=(10, 0))

        self.sheet_canvas = tk.Canvas(self.root, borderwidth=1)
        self.sheet_frame = tk.Frame(self.sheet_canvas)
        self.vsb = tk.Scrollbar(self.root, orient="vertical", command=self.sheet_canvas.yview)
        self.sheet_canvas.configure(yscrollcommand=self.vsb.set)

        self.vsb.pack(side="right", fill="y", padx=(0, 5))
        self.sheet_canvas.pack(fill="both", expand=True, padx=10)
        self.sheet_canvas.create_window((0, 0), window=self.sheet_frame, anchor="nw")

        self.sheet_frame.bind("<Configure>", lambda event: self.sheet_canvas.configure(scrollregion=self.sheet_canvas.bbox("all")))

        # Process button
        tk.Button(self.root, text="✅ Process & Save", command=self.process_files, bg="green", fg="white", font=("Arial", 12, "bold"), height=2).pack(pady=15)

    def load_file1(self):
        self.file1_path = filedialog.askopenfilename(title="Select File 1 (Phone List)", filetypes=[("Excel files", "*.xlsx")])
        if self.file1_path:
            messagebox.showinfo("File Selected", f"Loaded File 1:\n{self.file1_path}")

    def load_file2(self):
        self.file2_path = filedialog.askopenfilename(title="Select File 2 (Multiple Sheets)", filetypes=[("Excel files", "*.xlsx")])
        if self.file2_path:
            try:
                sheets = pd.ExcelFile(self.file2_path).sheet_names
                self.populate_sheet_checkboxes(sheets)
                messagebox.showinfo("Sheets Loaded", "Select the sheets you want to process.")
            except Exception as e:
                messagebox.showerror("Error", f"Could not load sheets:\n{e}")

    def populate_sheet_checkboxes(self, sheets):
        # Clear previous checkboxes
        for widget in self.sheet_frame.winfo_children():
            widget.destroy()

        self.sheet_vars.clear()
        for sheet in sheets:
            var = tk.BooleanVar()
            cb = tk.Checkbutton(self.sheet_frame, text=sheet, variable=var, font=("Arial", 10))
            cb.pack(anchor="w", padx=10, pady=2)
            self.sheet_vars[sheet] = var

    def process_files(self):
        if not self.file1_path or not self.file2_path:
            messagebox.showwarning("Missing Files", "Please select both File 1 and File 2.")
            return

        selected_sheets = [sheet for sheet, var in self.sheet_vars.items() if var.get()]
        if not selected_sheets:
            messagebox.showwarning("No Sheets Selected", "Please select at least one sheet.")
            return

        try:
            # Load phone numbers from File 1
            df_phones = pd.read_excel(self.file1_path, sheet_name="9-5 get work order", usecols=["ACUD_UNIT_ID", "MOBILE_NUMBER"])
            df_phones.drop_duplicates(subset="ACUD_UNIT_ID", inplace=True)

            updated_sheets = {}
            for sheet in selected_sheets:
                df = pd.read_excel(self.file2_path, sheet_name=sheet)
                df_renamed = df.rename(columns={"كود الوحدة": "ACUD_UNIT_ID"})
                df_merged = df_renamed.merge(df_phones, on="ACUD_UNIT_ID", how="left")
                df_merged.rename(columns={"MOBILE_NUMBER": "Matched Mobile", "ACUD_UNIT_ID": "كود الوحدة"}, inplace=True)
                updated_sheets[sheet] = df_merged

            # Save output
            output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", title="Save Output File", filetypes=[("Excel files", "*.xlsx")])
            if output_path:
                with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                    for sheet_name, df in updated_sheets.items():
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                messagebox.showinfo("Success", f"Updated file saved:\n{output_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Processing failed:\n{e}")


# Run the enhanced UI
if __name__ == "__main__":
    root = tk.Tk()
    app = PhoneMatcherApp(root)
    root.mainloop()
