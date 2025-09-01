import os
import json
from io import BytesIO
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import tkinter.font as tkFont
import threading
import tempfile

# Font mapping from Tkinter names to ReportLab names
FONT_MAPPING = {
    'Arial': 'Helvetica',
    'Times New Roman': 'Times-Roman',
    'Courier New': 'Courier',
    'Verdana': 'Helvetica',  # Fallback to Helvetica
    'Tahoma': 'Helvetica',   # Fallback to Helvetica
    'Georgia': 'Times-Roman', # Fallback to Times-Roman
    'Calibri': 'Helvetica',  # Fallback to Helvetica
    'Preeti': 'Helvetica',   # Fallback to Helvetica
    'Ganesh': 'Helvetica',   # Fallback to Helvetica
    'Kantipur': 'Helvetica'  # Fallback to Helvetica
}

# Register any custom fonts if available
try:
    # Try to register custom fonts if they exist
    if os.path.exists('preeti.ttf'):
        pdfmetrics.registerFont(TTFont('Preeti', 'preeti.ttf'))
        FONT_MAPPING['Preeti'] = 'Preeti'
except:
    pass

try:
    import win32com.client
    HAVE_WIN32 = True
except:
    HAVE_WIN32 = False

# ---------------- PDF & Excel Utilities ----------------
def export_excel_to_pdf(excel_path, pdf_path, sheet_index=1):
    if not HAVE_WIN32:
        raise RuntimeError("pywin32 not available. Install pywin32 to use Excel automation.")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(os.path.abspath(excel_path))
    try:
        ws = wb.Worksheets(sheet_index)
        ws.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
    finally:
        wb.Close(False)
        excel.Quit()

def make_footer_overlay_stream(page_width, page_height, footer_items,
                               page_num, total_pages,
                               left_margin=36, right_margin=36, bottom_margin=24,
                               font='Helvetica', size1=13, size2=13, line_gap=4):
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_width, page_height))
    ncols = len(footer_items)
    avail_w = page_width - left_margin - right_margin
    col_w = avail_w / ncols
    y_line2 = bottom_margin + size2
    y_line1 = y_line2 + size2 + line_gap

    # Use the mapped font
    for i, (line1, line2) in enumerate(footer_items):
        x_center = left_margin + (i * col_w) + col_w / 2.0
        line1 = line1.replace("{page}", str(page_num)).replace("{total}", str(total_pages))
        line2 = line2.replace("{page}", str(page_num)).replace("{total}", str(total_pages))
        c.setFont(font, size1)
        c.drawCentredString(x_center, y_line1, line1)
        c.setFont(font, size2)
        c.drawCentredString(x_center, y_line2, line2)

    c.save()
    packet.seek(0)
    return packet

def add_footer_to_pdf(input_pdf_path, output_pdf_path, footer_items, font_name, font_size):
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()
    total = len(reader.pages)

    # Map the font name for ReportLab
    mapped_font = FONT_MAPPING.get(font_name, 'Helvetica')

    for idx, page in enumerate(reader.pages):
        mediabox = page.mediabox
        page_w = float(mediabox[2]) - float(mediabox[0])
        page_h = float(mediabox[3]) - float(mediabox[1])
        overlay_stream = make_footer_overlay_stream(page_w, page_h, footer_items, idx+1, total,
                                                    font=mapped_font, size1=font_size, size2=font_size)
        overlay_pdf = PdfReader(overlay_stream)
        overlay_page = overlay_pdf.pages[0]
        page.merge_page(overlay_page)
        writer.add_page(page)

    with open(output_pdf_path, "wb") as f_out:
        writer.write(f_out)

# ---------------- Tkinter GUI ----------------
class FooterApp:
    def __init__(self, root):
        self.root = root
        root.title("PDF Footer Automation")
        root.geometry("800x600")

        # Get draft file path in user's temp directory
        self.draft_file_path = os.path.join(tempfile.gettempdir(), "pdf_footer_draft.json")
        
        # Get available system fonts for Tkinter
        self.available_fonts = list(tkFont.families())
        
        # Use your specified common fonts that are available
        common_fonts = ['Arial', 'Times New Roman', 'Courier New', 'Verdana', 
                       'Tahoma', 'Georgia', 'Preeti', 'Calibri', 'Ganesh', 'Kantipur']
        
        # Filter to only include fonts that are actually available
        self.fonts_available = [font for font in common_fonts if font in self.available_fonts]
        if not self.fonts_available:
            self.fonts_available = self.available_fonts[:10]  # Fallback to first 10 available fonts

        # Default font - use first available font
        self.font_var = tk.StringVar(value=self.fonts_available[0])
        self.font_size_var = tk.IntVar(value=16)
        self.col_count_var = tk.IntVar(value=5)

        # File selection
        tk.Label(root, text="Source PDF / Excel:").pack(pady=2)
        self.src_var = tk.StringVar()
        tk.Entry(root, textvariable=self.src_var, width=80).pack()
        tk.Button(root, text="Browse", command=self.browse_src).pack(pady=2)

        # Font name & size
        frame_font = tk.Frame(root)
        frame_font.pack(pady=5)
        tk.Label(frame_font, text="Font Name:").grid(row=0, column=0)
        font_dropdown = ttk.Combobox(frame_font, textvariable=self.font_var,
                                     values=self.fonts_available, state="readonly", width=30)
        font_dropdown.grid(row=0, column=1)
        tk.Label(frame_font, text="Font Size:").grid(row=0, column=2, padx=(10,0))
        tk.Entry(frame_font, textvariable=self.font_size_var, width=5).grid(row=0, column=3)

        # Footer columns
        frame_col = tk.Frame(root)
        frame_col.pack(pady=5)
        tk.Label(frame_col, text="Number of Footer Columns:").grid(row=0, column=0)
        tk.Entry(frame_col, textvariable=self.col_count_var, width=5).grid(row=0, column=1)
        tk.Button(frame_col, text="Set Columns", command=self.set_columns).grid(row=0, column=2, padx=5)

        # Footer entry fields placeholder
        self.entries_frame = None
        self.footer_entries = []

        # Run button
        self.run_button = tk.Button(root, text="Add Footer", command=self.run, bg="green", fg="white")
        self.run_button.pack(pady=15)

        # Auto-load draft
        self.load_draft(auto=True)

    def browse_src(self):
        file = filedialog.askopenfilename(filetypes=[("PDF and Excel files", "*.pdf;*.xlsx")])
        if file:
            self.src_var.set(file)

    def set_columns(self):
        self.footer_entries = []
        col_count = self.col_count_var.get()
        if hasattr(self, 'entries_frame') and self.entries_frame:
            self.entries_frame.destroy()

        self.entries_frame = tk.Frame(self.root)
        self.entries_frame.pack(pady=10)

        # Get the selected font, fallback to Arial if not available
        selected_font = self.font_var.get()
        if selected_font not in self.available_fonts:
            selected_font = 'Arial'

        font_obj = tkFont.Font(family=selected_font, size=self.font_size_var.get())

        for i in range(col_count):
            tk.Label(self.entries_frame, text=f"Column {i+1} Line 1:").grid(row=i, column=0, sticky="e")
            line1_var = tk.StringVar()
            entry1 = tk.Entry(self.entries_frame, textvariable=line1_var, width=25, font=font_obj)
            entry1.grid(row=i, column=1)

            tk.Label(self.entries_frame, text=f"Column {i+1} Line 2:").grid(row=i, column=2, sticky="e")
            line2_var = tk.StringVar()
            entry2 = tk.Entry(self.entries_frame, textvariable=line2_var, width=25, font=font_obj)
            entry2.grid(row=i, column=3)

            self.footer_entries.append((line1_var, line2_var))

    def save_draft(self):
        data = {
            "src": self.src_var.get(),
            "columns": self.col_count_var.get(),
            "font_name": self.font_var.get(),
            "font_size": self.font_size_var.get(),
            "footers": [(l1.get(), l2.get()) for l1, l2 in self.footer_entries]
        }
        try:
            with open(self.draft_file_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            # Don't show error for draft saving - it's not critical
            print(f"Could not save draft: {e}")

    def load_draft(self, auto=False):
        try:
            if not os.path.exists(self.draft_file_path):
                return
            with open(self.draft_file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.src_var.set(data.get("src",""))
            self.col_count_var.set(data.get("columns",5))
            
            # Load font name with fallback
            saved_font = data.get("font_name", self.fonts_available[0])
            if saved_font in self.available_fonts:
                self.font_var.set(saved_font)
            else:
                self.font_var.set(self.fonts_available[0])
                
            self.font_size_var.set(data.get("font_size",16))
            self.set_columns()
            for (l1_var,l2_var),(val1,val2) in zip(self.footer_entries,data.get("footers",[])):
                l1_var.set(val1)
                l2_var.set(val2)
            if not auto:
                messagebox.showinfo("Draft Loaded","Footer draft loaded successfully!")
        except Exception as e:
            # Don't show error for draft loading - it's not critical
            if not auto:
                print(f"Could not load draft: {e}")

    def set_ui_state(self, enabled):
        """Enable or disable UI elements during processing"""
        state = "normal" if enabled else "disabled"
        self.run_button.config(state=state)
        # You can add more UI elements here if needed

    def process_footer(self, src, dest):
        """Process the footer in a separate thread"""
        try:
            footer_items = [(l1.get(), l2.get()) for l1, l2 in self.footer_entries]
            temp_pdf = None
            if src.lower().endswith(".xlsx"):
                temp_pdf = os.path.splitext(dest)[0]+"_temp_input.pdf"
                export_excel_to_pdf(src, temp_pdf)
                input_pdf=temp_pdf
            else:
                input_pdf=src

            add_footer_to_pdf(input_pdf, dest, footer_items,
                              font_name=self.font_var.get(),
                              font_size=self.font_size_var.get())

            if temp_pdf and os.path.exists(temp_pdf):
                os.remove(temp_pdf)

            # save draft after success (silently - don't show errors)
            try:
                self.save_draft()
            except:
                pass  # Silently ignore draft saving errors
            
            # Show success message in main thread
            self.root.after(0, lambda: messagebox.showinfo("Success", f"Footer added successfully!\nOutput: {dest}"))
            
        except Exception as e:
            # Show error message in main thread
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            # Re-enable UI in main thread
            self.root.after(0, lambda: self.set_ui_state(True))
            self.root.after(0, lambda: self.run_button.config(text="Add Footer", bg="green"))

    def run(self):
        src = self.src_var.get()
        if not src:
            messagebox.showerror("Error", "Please select source file")
            return
        if not self.footer_entries:
            messagebox.showerror("Error", "Please set number of columns and fill footer values")
            return

        # Generate suggested file name
        base_name, ext = os.path.splitext(os.path.basename(src))
        suggested_name = f"{base_name}_footer.pdf"
        initial_dir = os.path.dirname(src) if src else ""

        # Ask user for destination with suggested file name
        dest = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            initialfile=suggested_name,
            initialdir=initial_dir,
            filetypes=[("PDF files", "*.pdf")],
            title="Save Output PDF As"
        )
        
        if not dest:  # User cancelled
            return

        # Change button to "Wait..." and disable UI
        self.run_button.config(text="Wait...", bg="gray", state="disabled")
        self.set_ui_state(False)
        
        # Process in a separate thread to keep UI responsive
        thread = threading.Thread(target=self.process_footer, args=(src, dest))
        thread.daemon = True
        thread.start()

# ---------------- Run App ----------------
if __name__ == "__main__":
    root = tk.Tk()
    app = FooterApp(root)
    root.mainloop()