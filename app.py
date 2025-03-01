from flask import Flask, request, send_file, render_template, send_from_directory
import pandas as pd
from fpdf import FPDF
import os
import datetime
import random

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['EXAMPLE_FOLDER'] = 'static/'

# ✅ Перемещаем класс PDF ВЫШЕ, перед его первым использованием
class PDF(FPDF):
    def __init__(self, filename, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.filename = filename.replace('.xlsx', '.flw')
        self.show_headers = True

    def header(self):
        self.image('fluck.png', 0, 4, 206)
        self.add_font('Arial', '', 'Arial.ttf', uni=True)
        self.add_font('Arial', 'B', 'Arial-Bold.ttf', uni=True)
        self.add_font('Arial', 'I', 'Arial-Italic.ttf', uni=True)
        self.add_font('Arial', 'BI', 'Arial-BoldItalic.ttf', uni=True)
        self.set_font('Arial', 'B', 6.8)
        self.set_stretching(95.0)
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

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part"
    file = request.files['file']
    start_date = request.form.get('start_date')
    start_time = request.form.get('start_time')
    if file.filename == '':
        return "No selected file"
    if file:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)

        pdf_filename = filepath.replace('.xlsx', '.flw')

        df = pd.read_excel(filepath, engine='openpyxl')
        pdf = PDF(filename=file.filename)
        pdf.add_page()
        pdf.set_font('Arial', '', 6.8)
        pdf.set_stretching(95.0)
        h = 2.8

        current_time = datetime.datetime.strptime(f"{start_date} {start_time}", "%Y-%m-%d %H:%M")

        for index, row in df.iterrows():
            if pdf.get_y() > 260:
                pdf.add_page()
            pdf.cell(36, h, txt=str(row[0]))
            pdf.cell(33, h, txt=str(row[1]))
            pdf.cell(28, h, txt=str(row[2]))
            pdf.cell(30, h, txt=str(row[3]) + ' m')
            pdf.cell(32, h, txt=str(row[4]))#f'{random.randrange(69, 96, 1)/10:.2f} dB (NEXT)')
            pdf.cell(30, h, txt=str(row[5]))#current_time.strftime("%d/%m/%Y %H:%M"), ln=True)
            current_time += datetime.timedelta(seconds=random.randint(45, 245))

        total_length = round(df['Length'].sum(), 2) if 'Length' in df.columns else 0
        total_reports = len(df)
        passing_reports = df[df['Summary'] == 'PASS'].shape[0] if 'Summary' in df.columns else 0
        failing_reports = df[df['Summary'] == 'FAIL'].shape[0] if 'Summary' in df.columns else 0
        warning_reports = df[df['Summary'] == 'WARNING'].shape[0] if 'Summary' in df.columns else 0
        documentation_only = df[df['Summary'] == 'DOCUMENTATION ONLY'].shape[0] if 'Summary' in df.columns else 0

        pdf.show_headers = False
        pdf.add_page()
        pdf.set_y(26)
        pdf.set_x(6)
        pdf.set_font('Arial', '', 9)
        pdf.cell(0, 3.3, 'Total Length:', 0, 1, 'L')
        pdf.set_x(6)
        pdf.cell(0, 3.3, 'Number of Reports:', 0, 1, 'L')
        pdf.set_x(6)
        pdf.cell(0, 3.3, 'Number of Passing Reports:', 0, 1, 'L')
        pdf.set_x(6)
        pdf.cell(0, 3.3, 'Number of Failing Reports:', 0, 1, 'L')
        pdf.set_x(6)
        pdf.cell(0, 3.3, 'Number of Warning Reports:', 0, 1, 'L')
        pdf.set_x(6)
        pdf.cell(0, 3.3, 'Documentation Only:', 0, 1, 'L')

        pdf.set_y(26)
        pdf.cell(60, 3.3, f'{total_length} m', 0, 1, 'R')
        pdf.cell(60, 3.3, f'{total_reports}', 0, 1, 'R')
        pdf.cell(60, 3.3, f'{passing_reports}', 0, 1, 'R')
        pdf.cell(60, 3.3, f'{failing_reports}', 0, 1, 'R')
        pdf.cell(60, 3.3, f'{warning_reports}', 0, 1, 'R')
        pdf.cell(60, 3.3, f'{documentation_only}', 0, 1, 'R')

        pdf_filename = filepath.replace('.xlsx', '.pdf')
        pdf.output(pdf_filename)

        return send_file(pdf_filename, as_attachment=True)

@app.route('/download-example')
def download_example():
    return send_from_directory(app.config['EXAMPLE_FOLDER'], 'fluck.xlsx', as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=True)