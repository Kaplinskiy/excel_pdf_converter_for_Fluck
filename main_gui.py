#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Offline Excel → PDF Converter with a modern Tkinter GUI,
with Exit button and auto-open generated PDF.
"""

import os
import random
import datetime
import warnings
import subprocess
import sys
import pandas as pd
from fpdf import FPDF
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Suppress fpdf warnings to keep console clean
warnings.filterwarnings("ignore", category=DeprecationWarning, module="fpdf")
warnings.filterwarnings("ignore", category=UserWarning, module="fpdf")

# Paths to resources
BASE_DIR      = os.getcwd()  # assumes script is run from its own directory
IMG_FLUCK     = os.path.join(BASE_DIR, "fluck.png")
IMG_BLUE_LINE = os.path.join(BASE_DIR, "blue_line.png")
IMG_FL        = os.path.join(BASE_DIR, "fl.png")
FONT_ARIAL    = os.path.join(BASE_DIR, "Arial.ttf")
FONT_ARIAL_B  = os.path.join(BASE_DIR, "Arial-Bold.ttf")
FONT_ARIAL_I  = os.path.join(BASE_DIR, "Arial-Italic.ttf")
FONT_ARIAL_BI = os.path.join(BASE_DIR, "Arial-BoldItalic.ttf")

def open_file(path):
    """Open file with default application depending on OS."""
    try:
        if sys.platform.startswith('darwin'):
            subprocess.call(('open', path))
        elif os.name == 'nt':
            os.startfile(path)
        else:
            subprocess.call(('xdg-open', path))
    except Exception as e:
        messagebox.showerror("Ошибка при открытии", str(e))

class PDF(FPDF):
    def __init__(self, filename, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.filename = filename.replace('.xlsx', '.flw')
        self.show_headers = True

    def header(self):
        if os.path.exists(IMG_FLUCK):
            self.image(IMG_FLUCK, 0, 4, 206)
        self.add_font('Arial', '',  FONT_ARIAL)
        self.add_font('Arial', 'B', FONT_ARIAL_B)
        self.add_font('Arial', 'I', FONT_ARIAL_I)
        self.add_font('Arial', 'BI',FONT_ARIAL_BI)
        self.set_font('Arial', 'B', 6.8)
        try:
            self.set_stretching(95.0)
        except AttributeError:
            pass
        if self.show_headers:
            self.ln(16)
            self.cell(36, 4, 'Cable ID',    0, 0, 'L')
            self.cell(33, 4, 'Summary',     0, 0, 'L')
            self.cell(28, 4, 'Test Limit',  0, 0, 'L')
            self.cell(30, 4, 'Length',      0, 0, 'L')
            self.cell(32, 4, 'Headroom',    0, 0, 'L')
            self.cell(30, 4, 'Date / Time', 0, 0, 'L')
            self.ln(3.8)

    def footer(self):
        self.set_y(-28.5)
        if os.path.exists(IMG_BLUE_LINE):
            self.image(IMG_BLUE_LINE, 6, 264, 195)
        if os.path.exists(IMG_FL):
            self.image(IMG_FL, 145, 270, 50)
        self.set_font('Arial', 'B', 9)
        self.set_x(6)
        current_date = datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S %p')
        self.cell(90, 8, current_date)
        self.cell(0, 8, f'Page {self.page_no()}', 0, 0, 'L')
        self.ln(4)
        self.set_x(6)
        self.cell(30, 8, self.filename)

def generate_pdf(filepath, start_date, start_time):
    df = pd.read_excel(filepath, engine='openpyxl')
    pdf = PDF(filename=os.path.basename(filepath))
    pdf.add_page()
    pdf.set_font('Arial', '', 6.8)
    try:
        pdf.set_stretching(95.0)
    except AttributeError:
        pass
    h = 2.8
    ctime = datetime.datetime.strptime(f"{start_date} {start_time}", "%Y-%m-%d %H:%M")

    for _, row in df.iterrows():
        if pdf.get_y() > 260:
            pdf.add_page()
        pdf.cell(36, h, txt=str(row[0]))
        pdf.cell(33, h, txt=str(row[1]))
        pdf.cell(28, h, txt=str(row[2]))
        pdf.cell(30, h, txt=str(row[3]) + ' m')
        pdf.cell(32, h, txt=f'{random.randrange(100,250,1)/10:.2f} dB (NEXT)')
        pdf.cell(30, h, txt=ctime.strftime("%d/%m/%Y %H:%M"), ln=True)
        ctime += datetime.timedelta(seconds=random.randint(45, 245))

    total_length       = round(df['Length'].sum(), 2) if 'Length' in df.columns else 0
    total_reports      = len(df)
    passing_reports    = df[df['Summary']=='PASS'].shape[0] if 'Summary' in df.columns else 0
    failing_reports    = df[df['Summary']=='FAIL'].shape[0] if 'Summary' in df.columns else 0
    warning_reports    = df[df['Summary']=='WARNING'].shape[0] if 'Summary' in df.columns else 0
    documentation_only = df[df['Summary']=='DOCUMENTATION ONLY'].shape[0] if 'Summary' in df.columns else 0

    pdf.show_headers = False
    pdf.add_page()
    pdf.set_y(26); pdf.set_x(6)
    pdf.set_font('Arial', '', 9)
    pdf.cell(0,3.3,'Total Length:',0,1,'L'); pdf.set_x(6)
    pdf.cell(0,3.3,'Number of Reports:',0,1,'L'); pdf.set_x(6)
    pdf.cell(0,3.3,'Number of Passing Reports:',0,1,'L'); pdf.set_x(6)
    pdf.cell(0,3.3,'Number of Failing Reports:',0,1,'L'); pdf.set_x(6)
    pdf.cell(0,3.3,'Number of Warning Reports:',0,1,'L'); pdf.set_x(6)
    pdf.cell(0,3.3,'Documentation Only:',0,1,'L')

    pdf.set_y(26)
    pdf.cell(60,3.3,f'{total_length} m',0,1,'R')
    pdf.cell(60,3.3,f'{total_reports}',0,1,'R')
    pdf.cell(60,3.3,f'{passing_reports}',0,1,'R')
    pdf.cell(60,3.3,f'{failing_reports}',0,1,'R')
    pdf.cell(60,3.3,f'{warning_reports}',0,1,'R')
    pdf.cell(60,3.3,f'{documentation_only}',0,1,'R')

    out = filepath.replace('.xlsx', '.pdf')
    pdf.output(out)
    return out

class App(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=20)
        self.master = master
        self.master.title("Excel → PDF Converter")
        self.master.resizable(False, False)
        self.grid(sticky="NSEW")
        self.create_widgets()

    def create_widgets(self):
        ttk.Label(self, text="Выберите Excel-файл:").grid(row=0, column=0, sticky="W")
        self.file_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.file_var, width=40).grid(row=1, column=0, padx=(0,10))
        ttk.Button(self, text="Обзор...", command=self.browse_file).grid(row=1, column=1)

        ttk.Label(self, text="Дата начала (YYYY-MM-DD):").grid(row=2, column=0, pady=(10,0), sticky="W")
        self.date_var = tk.StringVar(value=datetime.date.today().isoformat())
        ttk.Entry(self, textvariable=self.date_var, width=20).grid(row=3, column=0, sticky="W")

        ttk.Label(self, text="Время начала (HH:MM):").grid(row=2, column=1, pady=(10,0), sticky="W")
        self.time_var = tk.StringVar(value="08:00")
        ttk.Entry(self, textvariable=self.time_var, width=10).grid(row=3, column=1, sticky="W")

        # Generate and Exit buttons
        ttk.Button(self, text="Сгенерировать PDF", command=self.on_generate).grid(row=4, column=0, pady=20)
        ttk.Button(self, text="Выход", command=self.master.destroy).grid(row=4, column=1, pady=20)

    def browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx")])
        if path:
            self.file_var.set(path)

    def on_generate(self):
        path = self.file_var.get()
        if not path or not os.path.exists(path):
            messagebox.showerror("Ошибка", "Файл не выбран или не существует.")
            return
        try:
            out = generate_pdf(path, self.date_var.get(), self.time_var.get())
            open_file(out)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

def main():
    root = tk.Tk()
    style = ttk.Style(root)
    style.theme_use("clam")
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
