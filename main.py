#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import random
import datetime
import warnings

import pandas as pd
from fpdf import FPDF

# 🔇 Глушим предупреждения fpdf2, чтобы консоль не засорялась
warnings.filterwarnings("ignore", category=DeprecationWarning, module="fpdf")
warnings.filterwarnings("ignore", category=UserWarning, module="fpdf")

# Папка с ресурсами (картинки/шрифты). Если всё лежит рядом с main.py — оставьте '.'
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

IMG_FLUCK      = os.path.join(BASE_DIR, "fluck.png")
IMG_BLUE_LINE  = os.path.join(BASE_DIR, "blue_line.png")
IMG_FL         = os.path.join(BASE_DIR, "fl.png")

FONT_ARIAL      = os.path.join(BASE_DIR, "Arial.ttf")
FONT_ARIAL_B    = os.path.join(BASE_DIR, "Arial-Bold.ttf")
FONT_ARIAL_I    = os.path.join(BASE_DIR, "Arial-Italic.ttf")
FONT_ARIAL_BI   = os.path.join(BASE_DIR, "Arial-BoldItalic.ttf")


class PDF(FPDF):
    def __init__(self, filename, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.filename = filename.replace('.xlsx', '.flw')
        self.show_headers = True

    def header(self):
        # Логотип
        if os.path.exists(IMG_FLUCK):
            self.image(IMG_FLUCK, 0, 4, 206)

        # Шрифты
        # Параметр uni больше не нужен в новых версиях fpdf2
        self.add_font('Arial',  '',  FONT_ARIAL)
        self.add_font('Arial',  'B', FONT_ARIAL_B)
        self.add_font('Arial',  'I', FONT_ARIAL_I)
        self.add_font('Arial',  'BI', FONT_ARIAL_BI)

        self.set_font('Arial', 'B', 6.8)

        # Если есть метод set_stretching (вдруг в версии PyFPDF) — используем
        try:
            self.set_stretching(95.0)
        except AttributeError:
            pass

        if self.show_headers:
            self.ln(16)
            # Оставляем старые параметры ln/txt ради идентичной верстки
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
        self.cell(0,  8, f'Page {self.page_no()}', 0, 0, 'L')
        self.ln(4)
        self.set_x(6)
        self.cell(30, 8, self.filename)


def generate_pdf(filepath, start_date, start_time):
    """
    Генерация PDF из Excel без изменения верстки.
    """
    df = pd.read_excel(filepath, engine='openpyxl')

    pdf = PDF(filename=os.path.basename(filepath))
    pdf.add_page()
    pdf.set_font('Arial', '', 6.8)

    # Если метод set_stretching есть — применим, как было в веб-версии
    try:
        pdf.set_stretching(95.0)
    except AttributeError:
        pass

    h = 2.8

    current_time = datetime.datetime.strptime(f"{start_date} {start_time}", "%Y-%m-%d %H:%M")

    for _, row in df.iterrows():
        # Проверка переполнения страницы
        if pdf.get_y() > 260:
            pdf.add_page()

        pdf.cell(36, h, txt=str(row[0]))
        pdf.cell(33, h, txt=str(row[1]))
        pdf.cell(28, h, txt=str(row[2]))
        pdf.cell(30, h, txt=str(row[3]) + ' m')
        pdf.cell(32, h, txt=f'{random.randrange(100, 250, 1)/10:.2f} dB (NEXT)')
        pdf.cell(30, h, txt=current_time.strftime("%d/%m/%Y %H:%M"), ln=True)

        current_time += datetime.timedelta(seconds=random.randint(45, 245))

    # Итоговая страница
    total_length       = round(df['Length'].sum(), 2) if 'Length'  in df.columns else 0
    total_reports      = len(df)
    passing_reports    = df[df['Summary'] == 'PASS'].shape[0]          if 'Summary' in df.columns else 0
    failing_reports    = df[df['Summary'] == 'FAIL'].shape[0]          if 'Summary' in df.columns else 0
    warning_reports    = df[df['Summary'] == 'WARNING'].shape[0]       if 'Summary' in df.columns else 0
    documentation_only = df[df['Summary'] == 'DOCUMENTATION ONLY'].shape[0] if 'Summary' in df.columns else 0

    pdf.show_headers = False
    pdf.add_page()
    pdf.set_y(26)
    pdf.set_x(6)
    pdf.set_font('Arial', '', 9)

    # Левый столбец
    pdf.cell(0, 3.3, 'Total Length:',              0, 1, 'L'); pdf.set_x(6)
    pdf.cell(0, 3.3, 'Number of Reports:',         0, 1, 'L'); pdf.set_x(6)
    pdf.cell(0, 3.3, 'Number of Passing Reports:', 0, 1, 'L'); pdf.set_x(6)
    pdf.cell(0, 3.3, 'Number of Failing Reports:', 0, 1, 'L'); pdf.set_x(6)
    pdf.cell(0, 3.3, 'Number of Warning Reports:', 0, 1, 'L'); pdf.set_x(6)
    pdf.cell(0, 3.3, 'Documentation Only:',        0, 1, 'L')

    # Правый столбец (значения)
    pdf.set_y(26)
    pdf.cell(60, 3.3, f'{total_length} m',      0, 1, 'R')
    pdf.cell(60, 3.3, f'{total_reports}',       0, 1, 'R')
    pdf.cell(60, 3.3, f'{passing_reports}',     0, 1, 'R')
    pdf.cell(60, 3.3, f'{failing_reports}',     0, 1, 'R')
    pdf.cell(60, 3.3, f'{warning_reports}',     0, 1, 'R')
    pdf.cell(60, 3.3, f'{documentation_only}',  0, 1, 'R')

    output_path = filepath.replace('.xlsx', '.pdf')
    pdf.output(output_path)
    return output_path


def main():
    # Простой CLI: спросим у пользователя параметры
    excel_path = input("Путь к Excel-файлу (*.xlsx): ").strip().strip('"').strip("'")
    if not excel_path:
        print("Файл не указан.")
        return
    if not os.path.exists(excel_path):
        print("Файл не найден.")
        return

    start_date = input("Дата начала (YYYY-MM-DD): ").strip() or datetime.date.today().isoformat()
    start_time = input("Время начала (HH:MM): ").strip() or "08:00"

    try:
        out_pdf = generate_pdf(excel_path, start_date, start_time)
        print(f"Готово! PDF: {out_pdf}")
    except Exception as e:
        print("Ошибка:", e)


if __name__ == "__main__":
    main()
