import tkinter as tk
import sys
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
from reportlab.lib.units import cm
from reportlab.lib.pagesizes import A4
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas
import json
import os

# handle lokasi file saat dijalankan di .exe
if getattr(sys, 'frozen', False):  
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(__file__)

font_path = os.path.join(base_path, "SourceHanSansSC-Bold.ttf")

pdfmetrics.registerFont(TTFont("SourceHanSans-Bold", font_path))

# === KAPASITAS TABLE ===
capacity_table = {
    ("FOAM", "5", "CLK"): 32, ("FOAM", "6", "CLK"): 32, ("FOAM", "8", "CLK"): 32, ("FOAM", "10", "CLK"): 32,
    ("FOAM", "12", "CLK"): 32, ("FOAM", "14", "CLK"): 32, ("FOAM", "16", "CLK"): 32,
    ("FOAM", "5", "K"): 32, ("FOAM", "6", "K"): 32, ("FOAM", "8", "K"): 32, ("FOAM", "10", "K"): 32,
    ("FOAM", "12", "K"): 32, ("FOAM", "14", "K"): 32, ("FOAM", "16", "K"): 32,
    ("FOAM", "5", "Q"): 40, ("FOAM", "6", "Q"): 40, ("FOAM", "8", "Q"): 40, ("FOAM", "10", "Q"): 40,
    ("FOAM", "12", "Q"): 40, ("FOAM", "14", "Q"): 32, ("FOAM", "16", "Q"): 32,
    ("FOAM", "5", "F"): 50, ("FOAM", "6", "F"): 50, ("FOAM", "8", "F"): 50, ("FOAM", "10", "F"): 50,
    ("FOAM", "12", "F"): 50, ("FOAM", "14", "F"): 35, ("FOAM", "16", "F"): 35,
    ("FOAM", "5", "T"): 60, ("FOAM", "6", "T"): 60, ("FOAM", "8", "T"): 60, ("FOAM", "10", "T"): 60,
    ("FOAM", "12", "T"): 60, ("FOAM", "14", "T"): 50, ("FOAM", "16", "T"): 50,
    ("FOAM", "5", "TXL"): 60, ("FOAM", "6", "TXL"): 60, ("FOAM", "8", "TXL"): 60, ("FOAM", "10", "TXL"): 60,
    ("FOAM", "12", "TXL"): 60, ("FOAM", "14", "TXL"): 50, ("FOAM", "16", "TXL"): 50,
    ("SPRING", "5", "CLK"): 20, ("SPRING", "6", "CLK"): 20, ("SPRING", "8", "CLK"): 20, ("SPRING", "10", "CLK"): 20,
    ("SPRING", "12", "CLK"): 20, ("SPRING", "14", "CLK"): 20, ("SPRING", "16", "CLK"): 20,
    ("SPRING", "5", "K"): 20, ("SPRING", "6", "K"): 20, ("SPRING", "8", "K"): 20, ("SPRING", "10", "K"): 20,
    ("SPRING", "12", "K"): 20, ("SPRING", "14", "K"): 20, ("SPRING", "16", "K"): 20,
    ("SPRING", "5", "Q"): 30, ("SPRING", "6", "Q"): 30, ("SPRING", "8", "Q"): 30, ("SPRING", "10", "Q"): 30,
    ("SPRING", "12", "Q"): 30, ("SPRING", "14", "Q"): 30, ("SPRING", "16", "Q"): 30,
    ("SPRING", "5", "F"): 30, ("SPRING", "6", "F"): 30, ("SPRING", "8", "F"): 30, ("SPRING", "10", "F"): 30,
    ("SPRING", "12", "F"): 30, ("SPRING", "14", "F"): 30, ("SPRING", "16", "F"): 30,
    ("SPRING", "5", "T"): 50, ("SPRING", "6", "T"): 50, ("SPRING", "8", "T"): 50, ("SPRING", "10", "T"): 50,
    ("SPRING", "12", "T"): 50, ("SPRING", "14", "T"): 50, ("SPRING", "16", "T"): 50,
    ("SPRING", "5", "TXL"): 50, ("SPRING", "6", "TXL"): 50, ("SPRING", "8", "TXL"): 50, ("SPRING", "10", "TXL"): 50,
    ("SPRING", "12", "TXL"): 50, ("SPRING", "14", "TXL"): 50, ("SPRING", "16", "TXL"): 50,
}

def parse_sku(sku):
    parts = sku.split("-")
    inch = None
    size = None
    for p in parts:
        if p.isdigit():
            inch = p
        elif p in ["K", "Q", "F", "T", "TXL", "CLK"]:
            size = p
    return inch, size

def draw_page(c, pallet_data):
    """
    c : canvas
    pallet_data : list of 2 element [data_top, data_bottom]
    """
    width, height = A4

    for i, data in enumerate(pallet_data):
        table = Table(data, colWidths=[150, 445], rowHeights=70)
        table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('FONTNAME', (0,0), (-1,-1), 'STSong-Light'),
            ('FONTSIZE', (0,0), (-1,-1), 18),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))

        # kalau i=0 → taruh di atas, kalau i=1 → taruh di bawah
        if i == 0:
            y_pos = height/2
        else:
            y_pos = 0

        table.wrapOn(c, width, height)
        table.drawOn(c, 0, y_pos)

    # garis putus-putus di tengah
    c.setDash(6, 4)
    c.line(0, height/2, width, height/2)
    c.setDash()

def generate_pdf(data, output_file="pallet_labels.pdf"):
    c = canvas.Canvas(output_file, pagesize=A4)
    width, height = A4

    label_height = height / 2
    margin_x = 1.5 * cm
    margin_y = 1 * cm
    col_width = (width - 2 * margin_x) / 2
    row_height = 2 * cm   # diperbesar biar font besar pas

    label_index = 0
    for row in data:
        label_pos = label_index % 2  # 0=atas, 1=bawah
        y_start = height - (label_pos * label_height) - margin_y - (row_height * 6)

        rows = [
            ("P / I", row["PO"]),
            ("INVOICE NUMBER", row["INV"]),
            ("ORDER NUMBER", row["ORDER"]),
            ("SKU", row["SKU"]),
            ("PALLET ORDER", row["PALLET"]),
            ("QUANTITY", f'{row["PALLET_QTY"]} PCS ({row["TOTAL_QTY"]})')
        ]

        # pake 1 font untuk semua
        c.setFont("SourceHanSans-Bold", 20)

        for i, (k, v) in enumerate(rows):
            y_row = y_start + (len(rows)-i-1) * row_height

            # kotak kiri
            c.rect(margin_x, y_row, col_width, row_height)
            c.drawCentredString(margin_x + col_width/2, y_row+0.6*cm, k)

            # kotak kanan
            c.rect(margin_x + col_width, y_row, col_width, row_height)
            c.drawCentredString(margin_x + 1.5*col_width, y_row+0.6*cm, str(v))

        # kalau sudah 2 label (atas+ bawah), bikin garis putus2 lalu halaman baru
        if label_index % 2 == 1:
            c.setDash(6, 6)
            c.line(margin_x, height/2, width-margin_x, height/2)
            c.setDash()
            c.showPage()

        label_index += 1

    c.save()

class PalletApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Pallet Label Generator | By Hendry黄恒利")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # Style configuration
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.configure_styles()
        
        # Variables
        self.input_file = tk.StringVar()
        self.output_folder = tk.StringVar(value=os.getcwd())
        self.output_filename = tk.StringVar(value="pallet_labels")
        self.capacity_data = capacity_table.copy()
        
        # Load settings if exist
        self.load_settings()
        
        self.create_widgets()
        
    def configure_styles(self):
        self.style.configure('Title.TLabel', font=('Arial', 16, 'bold'), foreground='#2c3e50')
        self.style.configure('Subtitle.TLabel', font=('Arial', 12, 'bold'), foreground='#34495e')
        self.style.configure('TButton', font=('Arial', 10), padding=6)
        self.style.configure('Header.TFrame', background='#3498db')
        self.style.configure('TFrame', background='#ecf0f1')
        self.style.configure('TLabel', background='#ecf0f1')
        self.style.map('Action.TButton', 
                      foreground=[('pressed', 'white'), ('active', 'white')],
                      background=[('pressed', '#2980b9'), ('active', '#2980b9')])
        self.style.configure('Action.TButton', foreground='white', background='#3498db')
        
    def create_widgets(self):
        # Main container
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Pallet Label Generator", style='Title.TLabel')
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Label(file_frame, text="Input Excel File:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(file_frame, textvariable=self.input_file).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_input_file).grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(file_frame, text="Output Folder:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(file_frame, textvariable=self.output_folder).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_output_folder).grid(row=1, column=2, padx=5, pady=5)
        
        ttk.Label(file_frame, text="Output Filename:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(file_frame, textvariable=self.output_filename).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # Capacity table section
        capacity_frame = ttk.LabelFrame(main_frame, text="Pallet Capacity Settings", padding="10")
        capacity_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 10))
        capacity_frame.columnconfigure(1, weight=1)
        capacity_frame.rowconfigure(1, weight=1)
        
        # Create treeview for capacity table
        columns = ('Type', 'Inch', 'Size', 'Capacity')
        self.capacity_tree = ttk.Treeview(capacity_frame, columns=columns, show='headings', height=10)
        
        # Define headings
        self.capacity_tree.heading('Type', text='Type')
        self.capacity_tree.heading('Inch', text='Inch')
        self.capacity_tree.heading('Size', text='Size')
        self.capacity_tree.heading('Capacity', text='Capacity')
        
        # Define columns
        self.capacity_tree.column('Type', width=100)
        self.capacity_tree.column('Inch', width=80)
        self.capacity_tree.column('Size', width=80)
        self.capacity_tree.column('Capacity', width=100)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(capacity_frame, orient=tk.VERTICAL, command=self.capacity_tree.yview)
        self.capacity_tree.configure(yscrollcommand=scrollbar.set)
        
        # Grid treeview and scrollbar
        self.capacity_tree.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        scrollbar.grid(row=0, column=3, sticky=(tk.N, tk.S), pady=(0, 10))
        
        # Buttons for capacity table
        ttk.Button(capacity_frame, text="Edit Selected", command=self.edit_capacity).grid(row=1, column=0, padx=5, pady=5)
        ttk.Button(capacity_frame, text="Add New", command=self.add_capacity).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(capacity_frame, text="Delete Selected", command=self.delete_capacity).grid(row=1, column=2, padx=5, pady=5)
        
        # Action buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=20)
        
        ttk.Button(button_frame, text="Generate", command=self.generate_pdf, style='Action.TButton').grid(row=0, column=0, padx=10)
        ttk.Button(button_frame, text="Save Settings", command=self.save_settings).grid(row=0, column=1, padx=10)
        ttk.Button(button_frame, text="Reset to Default", command=self.reset_capacity).grid(row=0, column=2, padx=10)
        ttk.Button(button_frame, text="Exit", command=self.root.quit).grid(row=0, column=3, padx=10)
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Populate capacity table
        self.populate_capacity_tree()
        
    def populate_capacity_tree(self):
        # Clear existing items
        for item in self.capacity_tree.get_children():
            self.capacity_tree.delete(item)
            
        # Add items from capacity_data
        for (type_, inch, size), capacity in self.capacity_data.items():
            self.capacity_tree.insert('', 'end', values=(type_, inch, size, capacity))
    
    def browse_input_file(self):
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.input_file.set(filename)
    
    def browse_output_folder(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder.set(folder)
    
    def edit_capacity(self):
        selected = self.capacity_tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select an item to edit.")
            return
            
        item = selected[0]
        values = self.capacity_tree.item(item, 'values')
        
        # Create dialog for editing
        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Capacity")
        dialog.geometry("300x250")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Type:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        type_var = tk.StringVar(value=values[0])
        type_entry = ttk.Entry(dialog, textvariable=type_var, state='readonly')
        type_entry.grid(row=0, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))
        
        ttk.Label(dialog, text="Inch:").grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
        inch_var = tk.StringVar(value=values[1])
        inch_entry = ttk.Entry(dialog, textvariable=inch_var, state='readonly')
        inch_entry.grid(row=1, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))
        
        ttk.Label(dialog, text="Size:").grid(row=2, column=0, padx=10, pady=10, sticky=tk.W)
        size_var = tk.StringVar(value=values[2])
        size_entry = ttk.Entry(dialog, textvariable=size_var, state='readonly')
        size_entry.grid(row=2, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))
        
        ttk.Label(dialog, text="Capacity:").grid(row=3, column=0, padx=10, pady=10, sticky=tk.W)
        capacity_var = tk.StringVar(value=values[3])
        capacity_entry = ttk.Entry(dialog, textvariable=capacity_var)
        capacity_entry.grid(row=3, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))
        
        def save_changes():
            try:
                new_capacity = int(capacity_var.get())
                # Update the data
                key = (values[0], values[1], values[2])
                self.capacity_data[key] = new_capacity
                # Update the treeview
                self.capacity_tree.item(item, values=(values[0], values[1], values[2], new_capacity))
                dialog.destroy()
                self.status_var.set("Capacity updated successfully")
            except ValueError:
                messagebox.showerror("Invalid Input", "Capacity must be a number.")
        
        ttk.Button(dialog, text="Save", command=save_changes).grid(row=4, column=0, columnspan=2, pady=20)
        
        dialog.columnconfigure(1, weight=1)
    
    def add_capacity(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Add New Capacity")
        dialog.geometry("300x250")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Type (FOAM/SPRING):").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        type_var = tk.StringVar()
        type_combo = ttk.Combobox(dialog, textvariable=type_var, values=["FOAM", "SPRING"])
        type_combo.grid(row=0, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))
        
        ttk.Label(dialog, text="Inch:").grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
        inch_var = tk.StringVar()
        inch_combo = ttk.Combobox(dialog, textvariable=inch_var, values=["5", "6", "8", "10", "12", "14", "16"])
        inch_combo.grid(row=1, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))
        
        ttk.Label(dialog, text="Size:").grid(row=2, column=0, padx=10, pady=10, sticky=tk.W)
        size_var = tk.StringVar()
        size_combo = ttk.Combobox(dialog, textvariable=size_var, values=["CLK", "K", "Q", "F", "T", "TXL"])
        size_combo.grid(row=2, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))
        
        ttk.Label(dialog, text="Capacity:").grid(row=3, column=0, padx=10, pady=10, sticky=tk.W)
        capacity_var = tk.StringVar()
        capacity_entry = ttk.Entry(dialog, textvariable=capacity_var)
        capacity_entry.grid(row=3, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))
        
        def add_new():
            try:
                type_val = type_var.get().upper()
                inch_val = inch_var.get()
                size_val = size_var.get()
                capacity_val = int(capacity_var.get())
                
                if not all([type_val, inch_val, size_val]):
                    messagebox.showerror("Missing Fields", "Please fill all fields.")
                    return
                
                # Check if already exists
                key = (type_val, inch_val, size_val)
                if key in self.capacity_data:
                    messagebox.showerror("Duplicate Entry", "This combination already exists.")
                    return
                
                # Add to data and treeview
                self.capacity_data[key] = capacity_val
                self.capacity_tree.insert('', 'end', values=(type_val, inch_val, size_val, capacity_val))
                dialog.destroy()
                self.status_var.set("New capacity added successfully")
            except ValueError:
                messagebox.showerror("Invalid Input", "Capacity must be a number.")
        
        ttk.Button(dialog, text="Add", command=add_new).grid(row=4, column=0, columnspan=2, pady=20)
        
        dialog.columnconfigure(1, weight=1)
    
    def delete_capacity(self):
        selected = self.capacity_tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select an item to delete.")
            return
            
        item = selected[0]
        values = self.capacity_tree.item(item, 'values')
        
        if messagebox.askyesno("Confirm Delete", f"Delete capacity for {values[0]} {values[1]} {values[2]}?"):
            key = (values[0], values[1], values[2])
            if key in self.capacity_data:
                del self.capacity_data[key]
                self.capacity_tree.delete(item)
                self.status_var.set("Capacity deleted successfully")
    
    def reset_capacity(self):
        if messagebox.askyesno("Confirm Reset", "Reset all capacities to default values?"):
            self.capacity_data = capacity_table.copy()
            self.populate_capacity_tree()
            self.status_var.set("Capacities reset to default")
    
    def save_settings(self):
        settings = {
            'input_file': self.input_file.get(),
            'output_folder': self.output_folder.get(),
            'output_filename': self.output_filename.get(),
            'capacity_data': {f"{k[0]}-{k[1]}-{k[2]}": v for k, v in self.capacity_data.items()}
        }
        
        try:
            with open('pallet_settings.json', 'w') as f:
                json.dump(settings, f, indent=4)
            self.status_var.set("Settings saved successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {str(e)}")
    
    def load_settings(self):
        try:
            if os.path.exists('pallet_settings.json'):
                with open('pallet_settings.json', 'r') as f:
                    settings = json.load(f)
                
                self.input_file.set(settings.get('input_file', ''))
                self.output_folder.set(settings.get('output_folder', os.getcwd()))
                
                # Load filename tanpa perlu tambah .pdf
                filename = settings.get('output_filename', 'pallet_labels')
                self.output_filename.set(filename)
                
                # Load capacity data
                capacity_data = settings.get('capacity_data', {})
                for k, v in capacity_data.items():
                    parts = k.split('-')
                    if len(parts) == 3:
                        self.capacity_data[(parts[0], parts[1], parts[2])] = v
        except:
            # If loading fails, keep default values
            pass
    
    def generate_pdf(self):
        # Validate input
        if not self.input_file.get():
            messagebox.showerror("Error", "Please select an input Excel file.")
            return
            
        if not os.path.exists(self.input_file.get()):
            messagebox.showerror("Error", "Input file does not exist.")
            return
            
        # Prepare output path - TAMBAHKAN .pdf OTOMATIS
        filename = self.output_filename.get()
        if not filename.lower().endswith('.pdf'):
            filename += '.pdf'
        output_path = os.path.join(self.output_folder.get(), filename)
        
        try:
            # Read Excel file
            df = pd.read_excel(self.input_file.get())
            
            all_labels = []
            for _, row in df.iterrows():
                inch, size = parse_sku(row["SKU"])
                capacity = self.capacity_data.get((row["TYPE"].upper(), str(inch), size), None)
                if capacity is None:
                    print(f"⚠️ SKU {row['SKU']} tidak ditemukan kapasitasnya, dilewati")
                    continue

                qty = int(row["QTY"])
                
                # Hitung jumlah pallet dengan toleransi 3 pcs
                if qty <= capacity + 3:  # Jika jumlah <= kapasitas + toleransi
                    total_pallets = 1
                    pallet_qtys = [qty]
                else:
                    total_pallets = qty // capacity
                    remainder = qty % capacity
                    
                    if remainder <= 3 and total_pallets > 0:  # Toleransi 3 pcs
                        # Gabung sisa dengan pallet terakhir
                        pallet_qtys = [capacity] * (total_pallets - 1) + [capacity + remainder]
                    else:
                        # Buat pallet tambahan untuk sisa
                        total_pallets += 1
                        pallet_qtys = [capacity] * (total_pallets - 1) + [remainder]

                # Generate labels untuk setiap pallet
                for i, pallet_qty in enumerate(pallet_qtys, 1):
                    all_labels.append({
                        "PO": row["PO"],
                        "INV": row["INV"],
                        "ORDER": row["ORDER"],
                        "SKU": row["SKU"],
                        "PALLET": f"{i} / {len(pallet_qtys)} PALLET",
                        "PALLET_QTY": pallet_qty,
                        "TOTAL_QTY": qty
                    })

            # Generate PDF
            generate_pdf(all_labels, output_path)
            
            self.status_var.set(f"PDF generated successfully: {output_path}")
            messagebox.showinfo("Success", f"PDF generated successfully:\n{output_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate PDF: {str(e)}")
            self.status_var.set("Error generating PDF")

def main():
    root = tk.Tk()
    app = PalletApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()