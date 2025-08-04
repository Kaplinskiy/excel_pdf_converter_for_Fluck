import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from fpdf import FPDF
import datetime
import random
import os
import subprocess

# PDF-класс (оставлен без изменений)
class PDF(FPDF):
    def __init__(self, filename, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.filename = filename.replace('.xlsx', '.flw')
        self.show_headers = True

    def header(self):
        # логотип и шрифты
        self.image('fluck.png', 0, 4, 206)
        self.add_font('Arial', '', 'Arial.ttf', uni=True)
        self.add_font('Arial', 'B', 'Arial-Bold.ttf', uni=True)
        self.add_font('Arial', 'I', 'Arial-Italic.ttf', uni=True)
        self.add_font('Arial', 'BI', 'Arial-BoldItalic.ttf', uni=True)
        self.set_font('Arial', 'B', 6.8)
        if self.show_headers:
            self.ln(16)
            self.cell(36, 4, 'Cable ID', 0, 0, 'L')
            self.cell(33, 4, 'Summary', 0, 0, 'L')
            self.cell(28, 4, 'Test Limit', 0, 0, 'L')
            self.cell(30, 4, 'Length', 0, 0, 'L')
            self.cell(32, 4, 'Headroom', 0, 0, 'L')
            self.cell(30, 4, 'Date / Time', 0, 0, 'L')
            self.ln(3.8)

    def footer(self):
        self.set_y(-28.5)
        self.image('blue_line.png', 6, 264, 195)
        self.image('fl.png', 145, 270, 50)
        self.set_font('Arial', 'B', 9)
        self.set_x(6)
        current_date = datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S %p')
        self.cell(90, 8, current_date)
        self.cell(0, 8, f'Page {self.page_no()}', 0, 0, 'L')
        self.ln(4)
        self.set_x(6)
        self.cell(30, 8, self.filename)

# Функция выбора файла
def select_file():
    path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")])
    if path:
        file_var.set(path)

# Основная функция конверсии
def convert():
    filepath = file_var.get()
    if not filepath:
        messagebox.showerror("Ошибка", "Выберите файл для конвертации")
        return

    try:
        df = pd.read_excel(filepath, engine='openpyxl')
        pdf = PDF(os.path.basename(filepath))
        pdf.add_page()
        pdf.set_font('Arial', '', 6.8)
        h = 2.8

        # стартовая дата/время
        current_time = datetime.datetime.strptime(
            f"{date_var.get()} {time_var.get()}", "%Y-%m-%d %H:%M")

        for _, row in df.iterrows():
            if pdf.get_y() > 260:
                pdf.add_page()

            pdf.cell(36, h, str(row[0]))
            pdf.cell(33, h, str(row[1]))
            pdf.cell(28, h, str(row[2]))
            pdf.cell(30, h, str(row[3]) + ' m')
            pdf.cell(32, h, f'{random.randrange(100,250)/10:.2f} dB (NEXT)')
            pdf.cell(30, h, current_time.strftime("%d/%m/%Y %H:%M"), ln=True)
            current_time += datetime.timedelta(seconds=random.randint(45, 245))

        # Итоговая страница
        total_length      = round(df['Length'].sum(), 2) if 'Length' in df.columns else 0
        total_reports     = len(df)
        passing_reports   = df[df['Summary']=='PASS'].shape[0] if 'Summary' in df.columns else 0
        failing_reports   = df[df['Summary']=='FAIL'].shape[0] if 'Summary' in df.columns else 0
        warning_reports   = df[df['Summary']=='WARNING'].shape[0] if 'Summary' in df.columns else 0
        documentation_only= df[df['Summary']=='DOCUMENTATION ONLY'].shape[0] if 'Summary' in df.columns else 0

        pdf.show_headers = False
        pdf.add_page()
        pdf.set_y(26)
        pdf.set_x(6)
        pdf.set_font('Arial', '', 9)
        pdf.cell(0, 3.3, 'Total Length:', 0, 1, 'L')
        pdf.set_x(6); pdf.cell(0, 3.3, 'Number of Reports:', 0, 1, 'L')
        pdf.set_x(6); pdf.cell(0, 3.3, 'Number of Passing Reports:', 0, 1, 'L')
        pdf.set_x(6); pdf.cell(0, 3.3, 'Number of Failing Reports:', 0, 1, 'L')
        pdf.set_x(6); pdf.cell(0, 3.3, 'Number of Warning Reports:', 0, 1, 'L')
        pdf.set_x(6); pdf.cell(0, 3.3, 'Documentation Only:', 0, 1, 'L')

        pdf.set_y(26)
        pdf.cell(60, 3.3, f'{total_length} m', 0, 1, 'R')
        pdf.cell(60, 3.3, f'{total_reports}', 0, 1, 'R')
        pdf.cell(60, 3.3, f'{passing_reports}', 0, 1, 'R')
        pdf.cell(60, 3.3, f'{failing_reports}', 0, 1, 'R')
        pdf.cell(60, 3.3, f'{warning_reports}', 0, 1, 'R')
        pdf.cell(60, 3.3, f'{documentation_only}', 0, 1, 'R')

        output = filepath.replace('.xlsx', '.pdf')
        pdf.output(output)

        # Открыть PDF сразу после генерации
        if os.name == 'nt':
            os.startfile(output)
        else:
            subprocess.Popen(['open', output])

    except Exception as e:
        messagebox.showerror("Ошибка", str(e))

# Построение GUI
def create_ui():
    global file_var, date_var, time_var

    root = tk.Tk()
    root.title("Excel → PDF Converter")
    # removed: root.geometry("480x200")
    root.resizable(False, False)

    # переменные
    file_var  = tk.StringVar()
    date_var  = tk.StringVar(value=datetime.date.today().strftime("%Y-%m-%d"))
    time_var  = tk.StringVar(value=datetime.datetime.now().strftime("%H:%M"))

    # виджеты
    tk.Label(root, text="Excel-файл:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=file_var, width=50).grid(row=0, column=1, padx=5, pady=5)
    tk.Button(root, text="Обзор…", command=select_file).grid(row=0, column=2, padx=5, pady=5)

    tk.Label(root, text="Start Date (YYYY-MM-DD):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=date_var).grid(row=1, column=1, padx=5, pady=5, sticky="w")

    tk.Label(root, text="Start Time (HH:MM):").grid(row=2, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=time_var).grid(row=2, column=1, padx=5, pady=5, sticky="w")

    tk.Button(root, text="Преобразовать", command=convert).grid(row=3, column=1, pady=10)
    tk.Button(root, text="Выход", command=root.quit).grid(row=3, column=2, pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_ui()
