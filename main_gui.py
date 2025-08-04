#!/usr/bin/env python3
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from fpdf import FPDF
import os
import datetime
import random

# ====================== КЛАСС PDF (не менять форматирование!) ======================
class PDF(FPDF):
    def __init__(self, filename, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # имя файла с расширением .flw для показа в футере
        self.filename = filename.replace('.xlsx', '.flw')
        self.show_headers = True

    def header(self):
        # Логотип вверху
        self.image('fluck.png', 0, 4, 206)
        # Добавление шрифтов Arial (оставлено без изменений)
        self.add_font('Arial', '', 'Arial.ttf')
        self.add_font('Arial', 'B', 'Arial-Bold.ttf')
        self.add_font('Arial', 'I', 'Arial-Italic.ttf')
        self.add_font('Arial', 'BI', 'Arial-BoldItalic.ttf')
        self.set_font('Arial', 'B', 6.8)
        # Показываем заголовки, только если нужно
        if self.show_headers:
            self.ln(16)
            self.cell(36, 4, 'Cable ID',     0, 0, 'L')
            self.cell(33, 4, 'Summary',      0, 0, 'L')
            self.cell(28, 4, 'Test Limit',   0, 0, 'L')
            self.cell(30, 4, 'Length',       0, 0, 'L')
            self.cell(32, 4, 'Headroom',     0, 0, 'L')
            self.cell(30, 4, 'Date / Time',  0, 0, 'L')
            self.ln(3.8)

    def footer(self):
        # Нижний колонтитул
        self.set_y(-28.5)
        self.image('blue_line.png', 6, 264, 195)
        self.image('fl.png',       145, 270, 50)
        self.set_font('Arial', 'B', 9)
        self.set_x(6)
        current_date = datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S %p')
        self.cell(90, 8, current_date)
        self.cell(0,   8, f'Page {self.page_no()}', 0, 0, 'L')
        self.ln(4)
        self.set_x(6)
        self.cell(30,  8, self.filename)

# ================================= UI-функции ====================================
def select_file():
    path = filedialog.askopenfilename(
        title="Выберите Excel-файл",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if path:
        file_var.set(path)

def convert():
    path = file_var.get()
    if not path:
        messagebox.showerror("Ошибка", "Файл не выбран")
        return

    # Парсим дату и время старта
    start_date = date_var.get()
    start_time = time_var.get()
    try:
        current_time = datetime.datetime.strptime(
            f"{start_date} {start_time}", "%Y-%m-%d %H:%M"
        )
    except ValueError:
        messagebox.showerror("Ошибка", "Неверный формат даты или времени")
        return

    # Читаем Excel
    df = pd.read_excel(path, engine='openpyxl')

    # Генерируем PDF
    pdf = PDF(filename=os.path.basename(path))
    pdf.add_page()
    pdf.set_font('Arial', '', 6.8)
    h = 2.8

    for _, row in df.iterrows():
        if pdf.get_y() > 260:
            pdf.add_page()
        pdf.cell(36, h, txt=str(row[0]))
        pdf.cell(33, h, txt=str(row[1]))
        pdf.cell(28, h, txt=str(row[2]))
        pdf.cell(30, h, txt=str(row[3]) + ' m')
        pdf.cell(32, h, txt=f'{random.randrange(100,250)/10:.2f} dB (NEXT)')
        pdf.cell(30, h, txt=current_time.strftime("%d/%m/%Y %H:%M"), ln=True)
        # добавляем случайные 15–35 секунд
        current_time += datetime.timedelta(seconds=random.randint(15,35))

    # Считаем итоги
    total_length       = round(df.get('Length', pd.Series()).sum(), 2)
    total_reports      = len(df)
    passing_reports    = df[df.get('Summary')=='PASS'].shape[0]
    failing_reports    = df[df.get('Summary')=='FAIL'].shape[0]
    warning_reports    = df[df.get('Summary')=='WARNING'].shape[0]
    documentation_only = df[df.get('Summary')=='DOCUMENTATION ONLY'].shape[0]

    pdf.show_headers = False
    pdf.add_page()
    pdf.set_y(26)
    pdf.set_x(6)
    pdf.set_font('Arial','',9)
    pdf.cell(0,10,'Total Length:',                 0,1,'L')
    pdf.cell(0,10,'Number of Reports:',            0,1,'L')
    pdf.cell(0,10,'Number of Passing Reports:',    0,1,'L')
    pdf.cell(0,10,'Number of Failing Reports:',    0,1,'L')
    pdf.cell(0,10,'Number of Warning Reports:',    0,1,'L')
    pdf.cell(0,10,'Documentation Only:',           0,1,'L')

    pdf.set_y(26)
    pdf.cell(60,10,f'{total_length} m',            0,1,'R')
    pdf.cell(60,10,f'{total_reports}',             0,1,'R')
    pdf.cell(60,10,f'{passing_reports}',           0,1,'R')
    pdf.cell(60,10,f'{failing_reports}',           0,1,'R')
    pdf.cell(60,10,f'{warning_reports}',           0,1,'R')
    pdf.cell(60,10,f'{documentation_only}',        0,1,'R')

    # Сохраняем и открываем
    output = path.replace('.xlsx', '.pdf')
    pdf.output(output)
    messagebox.showinfo("Готово", f"PDF сохранён:\n{output}")
    os.startfile(output)

def create_ui():
    global file_var, date_var, time_var
    root = tk.Tk()

    # Инициализируем переменные **после** создания корня
    file_var = tk.StringVar(master=root)
    date_var = tk.StringVar(
        master=root,
        value=datetime.date.today().isoformat()
    )
    time_var = tk.StringVar(
        master=root,
        value=datetime.datetime.now().strftime("%H:%M")
    )

    root.title("Excel → PDF Converter")
    root.geometry("480x200")
    root.resizable(False, False)

    # Разметка
    tk.Label(root, text="Excel-файл:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    tk.Entry(root,  textvariable=file_var, width=50).grid(row=0, column=1, padx=5, pady=5)
    tk.Button(root, text="Обзор…", command=select_file).grid(row=0, column=2, padx=5)

    tk.Label(root, text="Start Date (YYYY-MM-DD):").grid(row=1, column=0, sticky="e", padx=5)
    tk.Entry(root,  textvariable=date_var).grid(row=1, column=1, padx=5, pady=5, sticky="w")

    tk.Label(root, text="Start Time (HH:MM):").grid(row=2, column=0, sticky="e", padx=5)
    tk.Entry(root,  textvariable=time_var).grid(row=2, column=1, padx=5, pady=5, sticky="w")

    tk.Button(root, text="Преобразовать", command=convert, width=15).grid(row=3, column=1, pady=15)
    tk.Button(root, text="Выход",       command=root.destroy, width=10).grid(row=3, column=2)

    root.mainloop()

# ==================== Точка входа ====================
if __name__ == "__main__":
    create_ui()
