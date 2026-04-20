import threading
import tkinter as tk
from pathlib import Path
from tkinter import font
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from script import process_pdf


class PdfSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Drawing PDF Exporter")
        self.root.geometry("1020x720")
        self.root.minsize(900, 620)

        self.input_pdf_var = tk.StringVar()
        self.output_folder_var = tk.StringVar()
        self.purchasing_folder_var = tk.StringVar()

        self._configure_styles()
        self._build_ui()

    def _configure_styles(self):
        style = ttk.Style()
        available_themes = style.theme_names()
        if "vista" in available_themes:
            style.theme_use("vista")
        elif "clam" in available_themes:
            style.theme_use("clam")

        default_font = font.nametofont("TkDefaultFont")
        default_font.configure(size=10)
        text_font = font.nametofont("TkTextFont")
        text_font.configure(size=10)

        style.configure("Card.TLabelframe", padding=10)
        style.configure("Card.TLabelframe.Label", font=("Segoe UI", 10, "bold"))
        style.configure("Header.TLabel", font=("Segoe UI", 13, "bold"))
        style.configure("Subtle.TLabel", foreground="#3a3a3a")
        style.configure("Run.TButton", padding=(12, 6))

    def _build_ui(self):
        root_frame = ttk.Frame(self.root, padding=16)
        root_frame.pack(fill=tk.BOTH, expand=True)
        root_frame.columnconfigure(0, weight=1)
        root_frame.rowconfigure(2, weight=1)

        ttk.Label(root_frame, text="Drawing PDF Exporter", style="Header.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 8)
        )
        ttk.Label(
            root_frame,
            text="Import one PDF, export unique single-page part PDFs, and skip duplicates/existing files.",
            style="Subtle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(0, 10))

        source_card = ttk.LabelFrame(root_frame, text="Input and Output", style="Card.TLabelframe")
        source_card.grid(row=2, column=0, sticky="nsew")
        source_card.columnconfigure(1, weight=1)
        source_card.rowconfigure(4, weight=1)

        ttk.Label(source_card, text="Input PDF:", width=15, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(0, 8), pady=(2, 8)
        )
        ttk.Entry(source_card, textvariable=self.input_pdf_var).grid(
            row=0, column=1, sticky="ew", pady=(2, 8)
        )
        ttk.Button(source_card, text="Browse...", command=self.browse_input_pdf).grid(
            row=0, column=2, sticky="e", padx=(8, 0), pady=(2, 8)
        )

        ttk.Label(source_card, text="Output Folder:", width=15, anchor="w").grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=(0, 8)
        )
        ttk.Entry(source_card, textvariable=self.output_folder_var).grid(
            row=1, column=1, sticky="ew", pady=(0, 8)
        )
        ttk.Button(source_card, text="Browse...", command=self.browse_output_folder).grid(
            row=1, column=2, sticky="e", padx=(8, 0), pady=(0, 8)
        )

        ttk.Label(source_card, text="Purchasing Folder:", width=15, anchor="w").grid(
            row=2, column=0, sticky="w", padx=(0, 8), pady=(0, 8)
        )
        ttk.Entry(source_card, textvariable=self.purchasing_folder_var).grid(
            row=2, column=1, sticky="ew", pady=(0, 8)
        )
        ttk.Button(source_card, text="Browse...", command=self.browse_purchasing_folder).grid(
            row=2, column=2, sticky="e", padx=(8, 0), pady=(0, 8)
        )

        self.run_button = ttk.Button(
            source_card, text="Run Split", command=self.run_split, style="Run.TButton"
        )
        self.run_button.grid(row=3, column=0, columnspan=3, sticky="w", pady=(2, 10))

        ttk.Label(source_card, text="Results").grid(row=4, column=0, columnspan=3, sticky="w")
        self.results_text = ScrolledText(source_card, wrap="word", height=24, relief="solid", borderwidth=1)
        self.results_text.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=(6, 0))
        source_card.rowconfigure(5, weight=1)

        self.results_text.configure(state=tk.DISABLED)

    def browse_input_pdf(self):
        selected = filedialog.askopenfilename(
            title="Select Input PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if selected:
            self.input_pdf_var.set(selected)

    def browse_output_folder(self):
        selected = filedialog.askdirectory(title="Select Output Folder")
        if selected:
            self.output_folder_var.set(selected)

    def browse_purchasing_folder(self):
        selected = filedialog.askdirectory(title="Select Purchasing Folder")
        if selected:
            self.purchasing_folder_var.set(selected)

    def run_split(self):
        input_pdf = self.input_pdf_var.get().strip()
        output_folder = self.output_folder_var.get().strip()
        purchasing_folder = self.purchasing_folder_var.get().strip()

        if not input_pdf:
            messagebox.showerror("Missing Input", "Please select an input PDF.")
            return
        if not output_folder:
            messagebox.showerror("Missing Output", "Please select an output folder.")
            return
        if not Path(input_pdf).exists():
            messagebox.showerror("Invalid Input", "Selected input PDF does not exist.")
            return

        self.run_button.configure(state=tk.DISABLED)
        self._set_results_text("Processing, please wait...\n")

        worker = threading.Thread(
            target=self._run_split_worker,
            args=(input_pdf, output_folder, purchasing_folder),
            daemon=True,
        )
        worker.start()

    def _run_split_worker(self, input_pdf, output_folder, purchasing_folder):
        try:
            result = process_pdf(input_pdf, output_folder, purchasing_folder or None)
            self.root.after(0, self._on_run_success, result)
        except Exception as exc:
            self.root.after(0, self._on_run_error, str(exc))

    def _on_run_success(self, result):
        summary_lines = [
            f"Input PDF: {result['input_pdf']}",
            f"Output Folder: {result['output_folder']}",
            f"Purchasing Folder: {result['purchasing_folder'] or '(not selected)'}",
            f"Total Pages: {result['total_pages']}",
            "",
            f"Imported: {result['imported_count']}",
            f"Not Imported: {result['not_imported_count']}",
            f"Duplicates (same run): {result['duplicate_in_input_count']}",
            f"Already Exists in Folder: {result['already_exists_count']}",
            f"No Part Number: {result['no_part_number_count']}",
            f"Purchasing Part Skips (no folder): {result['purchasing_folder_missing_count']}",
            "",
            "Details:",
        ]

        detail_lines = []
        for item in result["details"]:
            part = item["part_number"] if item["part_number"] else "N/A"
            detail_lines.append(
                f"Page {item['page']}: {item['status']} | Part: {part} | Reason: {item['reason']}"
            )

        full_report = "\n".join(summary_lines + detail_lines)
        self._set_results_text(full_report)
        self.run_button.configure(state=tk.NORMAL)

    def _on_run_error(self, error_text):
        self._set_results_text(f"Error:\n{error_text}")
        self.run_button.configure(state=tk.NORMAL)
        messagebox.showerror("Run Failed", error_text)

    def _set_results_text(self, text):
        self.results_text.configure(state=tk.NORMAL)
        self.results_text.delete("1.0", tk.END)
        self.results_text.insert("1.0", text)
        self.results_text.configure(state=tk.DISABLED)


def main():
    root = tk.Tk()
    app = PdfSplitterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
