import sys
import os
import datetime
import random
import pandas as pd
from fpdf import FPDF
import tkinter as tk
from tkinter import filedialog, messagebox

# Определяем папку с ресурсами (для работы как .py, так и внутри PyInstaller EXE)
if getattr(sys, 'frozen', False):
    BASE_PATH = sys._MEIPASS
else:
    BASE_PATH = os.path.abspath(os.path.dirname(__file__))

class PDF(FPDF):
    def __init__(self, filename, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Сохраняем имя файла для footer
        self.filename = filename.replace('.xlsx', '.flw')
        self.show_headers = True

    def header(self):
        # Логотип
        self.image(os.path.join(BASE_PATH, 'fluck.png'), 0, 4, 206)
        # Регистрируем шрифты из ресурсов
        arial = os.path.join(BASE_PATH, 'Arial.ttf')
        arial_b = os.path.join(BASE_PATH, 'Arial-Bold.ttf')
        arial_i = os.path.join(BASE_PATH, 'Arial-Italic.ttf')
        arial_bi= os.path.join(BASE_PATH, 'Arial-BoldItalic.ttf')
        self.add_font('Arial', '', arial,    uni=True)
        self.add_font('Arial', 'B', arial_b, uni=True)
        self.add_font('Arial', 'I', arial_i, uni=True)
        self.add_font('Arial', 'BI',arial_bi,uni=True)
        self.set_font('Arial', 'B', 6.8)
        # Заголовки столбцов на всех, кроме итоговой страницы
        if self.show_headers:
            self.ln(16)
            self.cell(36, 4, 'Cable ID',       0, 0, 'L')
            self.cell(33, 4, 'Summary',        0, 0, 'L')
            self.cell(28, 4, 'Test Limit',     0, 0, 'L')
            self.cell(30, 4, 'Length',         0, 0, 'L')
            self.cell(32, 4, 'Headroom',       0, 0, 'L')
            self.cell(30, 4, 'Date / Time',    0, 0, 'L')
            self.ln(3.8)

    def footer(self):
        # Нижний колонтитул
        self.set_y(-28.5)
        self.image(os.path.join(BASE_PATH, 'blue_line.png'), 6, 264, 195)
        self.image(os.path.join(BASE_PATH, 'fl.png'),          145, 270, 50)
        self.set_font('Arial', 'B', 9)
        self.set_x(6)
        now = datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S %p')
        self.cell(90, 8, now)
        self.cell(0, 8, f'Page {self.page_no()}', 0, 0, 'L')
        self.ln(4)
        self.set_x(6)
        self.cell(30, 8, self.filename)

def select_file():
    """Открыть диалог выбора Excel-файла."""
    path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx;*.xls")],
        title="Выберите Excel-файл"
    )
    if path:
        file_var.set(path)

def convert():
    """Собрать PDF из выбранного Excel и открыть его для просмотра."""
    excel_path = file_var.get()
    date_str     = date_var.get()
    time_str     = time_var.get()

    if not excel_path:
        messagebox.showerror("Ошибка", "Файл не выбран.")
        return
    try:
        start_dt = datetime.datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M")
    except Exception:
        messagebox.showerror("Ошибка", "Неверный формат даты или времени.")
        return

    try:
        df = pd.read_excel(excel_path, engine='openpyxl')
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось прочитать Excel: {e}")
        return

    # Генерируем PDF
    pdf = PDF(filename=os.path.basename(excel_path))
    pdf.add_page()
    pdf.set_font('Arial', '', 6.8)
    h = 2.8
    current_time = start_dt

    for _, row in df.iterrows():
        if pdf.get_y() > 260:
            pdf.add_page()
        pdf.cell(36, h, txt=str(row[0]))
        pdf.cell(33, h, txt=str(row[1]))
        pdf.cell(28, h, txt=str(row[2]))
        pdf.cell(30, h, txt=f"{row[3]} m")
        pdf.cell(32, h, txt=f"{random.randrange(100,250)/10:.2f} dB (NEXT)")
        pdf.cell(30, h, txt=current_time.strftime("%d/%m/%Y %H:%M"), ln=True)
        # Случайный шаг 15–35 сек
        current_time += datetime.timedelta(seconds=random.randint(15,35))

    # Итоговая страница
    total_length       = round(df['Length'].sum(),2) if 'Length' in df.columns else 0
    total_reports      = len(df)
    passing            = df[df['Summary']=='PASS'].shape[0] if 'Summary' in df.columns else 0
    failing            = df[df['Summary']=='FAIL'].shape[0] if 'Summary' in df.columns else 0
    warning_reports    = df[df['Summary']=='WARNING'].shape[0] if 'Summary' in df.columns else 0
    doc_only           = df[df['Summary']=='DOCUMENTATION ONLY'].shape[0] if 'Summary' in df.columns else 0

    pdf.show_headers = False
    pdf.add_page()
    pdf.set_y(26)
    pdf.set_x(6)
    pdf.set_font('Arial', '', 9)
    pdf.cell(0, 3.3, 'Total Length:',               0, 1, 'L')
    pdf.cell(0, 3.3, 'Number of Reports:',          0, 1, 'L')
    pdf.cell(0, 3.3, 'Number of Passing Reports:',  0, 1, 'L')
    pdf.cell(0, 3.3, 'Number of Failing Reports:',  0, 1, 'L')
    pdf.cell(0, 3.3, 'Number of Warning Reports:',  0, 1, 'L')
    pdf.cell(0, 3.3, 'Documentation Only:',         0, 1, 'L')

    pdf.set_y(26)
    pdf.cell(60, 3.3, f"{total_length} m",      0, 1, 'R')
    pdf.cell(60, 3.3, f"{total_reports}",       0, 1, 'R')
    pdf.cell(60, 3.3, f"{passing}",             0, 1, 'R')
    pdf.cell(60, 3.3, f"{failing}",             0, 1, 'R')
    pdf.cell(60, 3.3, f"{warning_reports}",     0, 1, 'R')
    pdf.cell(60, 3.3, f"{doc_only}",            0, 1, 'R')

    # Сохраняем рядом с исходным файлом
    pdf_path = excel_path.replace('.xlsx', '.pdf')
    try:
        pdf.output(pdf_path)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось сохранить PDF: {e}")
        return

    # Открываем PDF в стандартном просмотрщике Windows
    try:
        os.startfile(pdf_path)
    except Exception:
        messagebox.showinfo("Готово", f"PDF сохранён: {pdf_path}")

def create_ui():
    root = tk.Tk()
    root.title("Excel → PDF Converter")
    root.geometry("480x200")
    root.resizable(False, False)

    tk.Label(root, text="Excel-файл:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    tk.Entry(root, textvariable=file_var, width=50).grid(row=0, column=1, padx=5, pady=5)
    tk.Button(root, text="Обзор…", command=select_file).grid(row=0, column=2, padx=5)

    tk.Label(root, text="Start Date (YYYY-MM-DD):").grid(row=1, column=0, sticky="e", padx=5)
    tk.Entry(root, textvariable=date_var).grid(row=1, column=1, padx=5, pady=5, sticky="w")

    tk.Label(root, text="Start Time (HH:MM):").grid(row=2, column=0, sticky="e", padx=5)
    tk.Entry(root, textvariable=time_var).grid(row=2, column=1, padx=5, pady=5, sticky="w")

    tk.Button(root, text="Преобразовать", command=convert, width=15).grid(row=3, column=1, pady=15)
    tk.Button(root, text="Выход",       command=root.destroy, width=10).grid(row=3, column=2)

    root.mainloop()

if __name__ == "__main__":
    # Глобальные переменные для UI
    file_var = tk.StringVar()
    date_var = tk.StringVar(value=datetime.date.today().isoformat())
    time_var = tk.StringVar(value=datetime.datetime.now().strftime("%H:%M"))
    create_ui()
