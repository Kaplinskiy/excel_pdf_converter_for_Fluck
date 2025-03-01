import argparse
import base64
import pandas as pd
import reportlab

base64.encodestring = base64.encodebytes
base64.decodestring = base64.decodebytes

from argparse import Namespace
from datetime import datetime, timedelta
# Standard library imports
from random import randint

# ReportLab imports
from reportlab.lib.colors import *
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.charts.axes import *
from reportlab.lib.units import inch, cm
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.graphics import renderPDF


try:
    # import _rl_accel

    ACCEL = 1
except ImportError:
    ACCEL = 0

# precalculate some basics parameters
top_margin = A4[0] - inch / 2
bottom_margin = inch / 2
left_margin = inch / 2
right_margin = A4[1] - inch / 2
frame_width = right_margin - left_margin

def getDataExel(file_name):
    return pd.read_excel(file_name)
    pass

def createParser():
    parser = argparse.ArgumentParser()
    parser.add_argument('-n', '--name', default='Default Name')
    parser.add_argument('-j', '--job', default='Default Job')
    parser.add_argument('-c', '--cable', default='Default Cable')
    parser.add_argument('--trace', default=0)
    parser.add_argument('-ft', '--fibertype', default="Custom")
    parser.add_argument('-o', '--operator', default="RAK.NET")
    parser.add_argument('-l', '--length', default=0.0)
    parser.add_argument('-d', '--date', default=datetime.today().strftime('%d-%m-%Y'))
    parser.add_argument('-t', '--time', default=datetime.now().strftime('%H:%M:%S'))
    parser.add_argument('-x', '--exel', default='')
    return parser


def drawPageFrame(canv,f_name,flag):
    # canv.line(left_margin, top_margin, right_margin, bottom_margin)
    canv.setFont("Helvetica", 9)
    canv.setLineWidth(0.5)
    canv.drawImage("veex.jpg", right_margin - 7 * inch, top_margin + 6, width=400, height=25)
    if flag == 1:
        num = str(randint(1, 3))
        canv.drawImage("graf_new"+num+".png", left_margin + 1 * cm, top_margin - 8 * cm, width=350, height=200)
    else:
        canv.drawImage("graf_new.png", left_margin + 1 * cm, top_margin - 8 * cm, width=350, height=200)
    canv.drawString(bottom_margin, left_margin+7.1*inch, f_name + '.sor')
    canv.line(left_margin-10, top_margin, right_margin, top_margin)
    canv.line(left_margin-10, top_margin - 4.5 * inch, right_margin, top_margin - 4.5 * inch)
    canv.line(left_margin-10, top_margin - 4.7 * inch, right_margin, top_margin - 4.7 * inch)
    canv.drawString(left_margin, top_margin-4.66 * inch, " #")
    canv.drawString(left_margin + 55, top_margin - 4.66 * inch, "Location, km")
    canv.drawString(left_margin + 175, top_margin - 4.66 * inch, "Loss, dB")
    canv.drawString(left_margin + 265, top_margin - 4.66 * inch, "Reflectance, dB")
    canv.drawString(left_margin + 380, top_margin - 4.66 * inch, "Attenuation, dB/km")
    canv.drawString(left_margin + 510, top_margin - 4.66 * inch, "Total Loss, dB")
    canv.drawString(left_margin, top_margin - 4.66 * inch - 15, "[0       R")
    canv.drawString(left_margin, top_margin - 4.66 * inch - 30, " 1]      R")
    canv.line(left_margin-10, bottom_margin, right_margin, bottom_margin)
    canv.drawCentredString(0.5 * A4[1], 0.5 * inch / 2,
                           "Page %d" % canv.getPageNumber())


def printScale(canv, length_cable, loss):

    # from reportlab.graphics.charts.axes import XValueAxis

    drawing = Drawing(200, 200)
    data = [(-length_cable * 0.65, length_cable * 2.55)]
    from reportlab.graphics.charts.axes import XValueAxis
    xAxis = XValueAxis()
    xAxis.setPosition(70, 50, 350)
    x_start = float(format(0 + randint(0, 99) / 100000, '.5f'))
    xAxis.valueSteps = [x_start, length_cable+x_start, length_cable*1.855]
    xAxis.strokeColor = pink
    # xAxis.strokeDashArray = [1, 10]
    xAxis.strokeWidth = 1
    xAxis.labelTextFormat = '%0.5f'
    xAxis.tickUp = 5
    xAxis.configure(data)
    drawing.add(xAxis)
    data = [(-55, 0)]
    yAxis = YValueAxis()
    yAxis.setPosition(70, 50, 220)
    yAxis.valueSteps = [-43, -32.267, -21.338, -10.693]
    yAxis.strokeColor = pink
    # yAxis.strokeDashArray = [1, 10]
    yAxis.strokeWidth = 1
    yAxis.labelTextFormat = '%0.3f'
    yAxis.configure(data)
    drawing.add(yAxis)
    canv.setFont("Helvetica", 9)
    strX='X: [ , ' + format(length_cable * 8.8, '.5f') + '] km;  ' + format(length_cable * 2.5, '.5f') + ' km/div'
    canv.drawString(bottom_margin-10, left_margin+3.4*inch, strX)
    strY='Y: [-55.000, 0.000] dB;  ' + format(21.338-10.693, '.3f') + ' dB/div'
    canv.drawString(bottom_margin-10, left_margin + 3.2 * inch, strY)
    canv.drawString(left_margin + 55, top_margin - 4.66 * inch - 15, format(x_start, '.5f'))
    canv.drawString(left_margin + 175, top_margin - 4.66 * inch - 30, format(length_cable + x_start, '.5f'))
    canv.drawString(left_margin + 175, top_margin - 4.66 * inch - 15, '0.000')
    canv.drawString(left_margin + 265, top_margin - 4.66 * inch - 15, format(-43.000 - randint(0, 999)/1000, '.3f'))
    canv.drawString(left_margin + 265, top_margin - 4.66 * inch - 30, format(-16.000 - randint(0, 999)/1000, '.3f'))
    canv.drawString(left_margin + 380, top_margin - 4.66 * inch - 30, format(loss / length_cable, '.3f'))
    canv.drawString(left_margin + 510, top_margin - 4.66 * inch - 30, format(loss, '.3f'))

    renderPDF.draw(drawing, canv, -10, 270)


def run(names: Namespace):
    if sys.platform[0:4] == 'java':
        impl = 'Jython'
    else:
        impl = 'Python'
    verStr = '%d.%d' % (sys.version_info[0:2])
    if ACCEL:
        accelStr = 'with _rl_accel'
    else:
        accelStr = 'without _rl_accel'
    print('Benchmark of %s %s %s' % (impl, verStr, accelStr))
    file_name = 'results/' + names.job + '_' + names.fibertype
    canv = canvas.Canvas(file_name + '.pdf', pagesize=landscape(A4), invariant=1)
    canv.setPageCompression(1)
    # do some title page stuff
    canv.setFont("Helvetica", 9)
    y = 11
    dx = 3.5 * cm
    ddx = 2.5
    t = names.time
    d = names.date
    print(names.trace)

    for i in range(1, int(names.trace) + 1):
        names.length = format(float(names.length) + randint(-99, 99) / 100000, '.5f')
        loss = float(names.length) * 2.4 + 0.037 + randint(0, 3) / 100
        printScale(canv, float(names.length), loss)
        textX = 0.50 * A4[1]
        textY = 7.5 * inch
        drawPageFrame(canv, file_name, names.flag)
        canv.setFont("Helvetica", 9)
        canv.drawString(textX, textY, "Report date:", charSpace=0)
        canv.drawString(textX + dx, textY, datetime.today().strftime('%d-%m-%Y'), charSpace=0)
        textY = textY - y
        canv.drawString(textX, textY, "Timestamp:", charSpace=0)
        canv.drawString(textX + dx, textY, d + "   " + t, charSpace=0)
        t = datetime.strftime(datetime.strptime(t, '%H:%M:%S') + timedelta(seconds=randint(25, 70)), '%H:%M:%S')
        textY = textY - y * 2
        canv.drawString(textX, textY, "MTT-PLUS 410+/PRO-QUAD1676221", charSpace=0)
        textY = textY - y
        canv.drawString(textX, textY, "/SM1310/SM1550/MM850/MM1300", charSpace=0)
        textY = textY - y
        canv.drawString(textX, textY, "1300 nm", charSpace=0)
        textY = textY - y * 2
        canv.drawString(textX, textY, "Job:", charSpace=0)
        canv.drawString(textX + dx, textY, names.job, charSpace=0)
        textY = textY - y
        canv.drawString(textX, textY, "Cable:", charSpace=0)
        canv.drawString(textX + dx, textY, names.cable, charSpace=0)
        textY = textY - y
        canv.drawString(textX, textY, "Fiber:", charSpace=0)
        canv.drawString(textX + dx, textY, str(i), charSpace=0)
        textY = textY - y
        canv.drawString(textX, textY, "Trace:", charSpace=0)
        textY = textY - y
        canv.drawString(textX, textY, "Fiber type:", charSpace=0)
        canv.drawString(textX + dx, textY, names.fibertype, charSpace=0)
        textY = textY - y
        canv.drawString(textX, textY, "Operator:", charSpace=0)
        canv.drawString(textX + dx, textY, names.operator, charSpace=0)
        textY = textY - y
        last = format(float(names.length) + randint(1, 9) / 10000, '.5f')
        canv.drawString(textX, textY, "First data point:", charSpace=0)
        canv.drawString(textX + dx, textY, format(float(last) - float(names.length), ".5f"), charSpace=0)
        canv.drawString(textX + ddx * inch, textY, "Km", charSpace=0)
        textY = textY - y
        canv.drawString(textX, textY, "Last data point:", charSpace=0)
        canv.drawString(textX + dx, textY, last, charSpace=0)
        canv.drawString(textX + ddx * inch, textY, "Km", charSpace=0)
        textY = textY - y
        canv.drawString(textX, textY, "Resolution:", charSpace=0)
        canv.drawString(textX + dx, textY, "0.032", charSpace=0)
        canv.drawString(textX + ddx * inch, textY, "m", charSpace=0)
        textY = textY - y
        canv.drawString(textX, textY, "Pulse Width:", charSpace=0)
        canv.drawString(textX + dx, textY, "10", charSpace=0)
        canv.drawString(textX + ddx * inch, textY, "ns", charSpace=0)
        textY = textY - y
        canv.drawString(textX, textY, "Averaging Time:", charSpace=0)
        canv.drawString(textX + dx, textY, "12", charSpace=0)
        canv.drawString(textX + ddx * inch, textY, "s", charSpace=0)
        textY = textY - y * 2
        canv.setFont("Helvetica-Bold", 9)
        canv.drawString(textX, textY, "Length:", charSpace=0)
        canv.drawString(textX + dx, textY, names.length, charSpace=0)
        canv.drawString(textX + ddx * inch, textY, "Km", charSpace=0)
        textY = textY - y
        canv.drawString(textX, textY, "Total loss:", charSpace=0)
        canv.drawString(textX + dx, textY, format(loss, ".2f"), charSpace=0)
        canv.drawString(textX + ddx * inch, textY, "Db", charSpace=0)
        textY = textY - y
        canv.drawString(textX, textY, "Total ORL:", charSpace=0)
        canv.drawString(textX + ddx * inch, textY, "Db", charSpace=0)

        canv.showPage()

    canv.save()

    fileSize = os.stat('odyssey.pdf')[6] / 1024
    print('%s pages , file size %d kb' % (names.trace, fileSize))
    import hashlib
    print('file digest: %s' % hashlib.md5(open('odyssey.pdf', 'rb').read()).hexdigest())


if __name__ == '__main__':
    parser = createParser()
    namespace: Namespace = parser.parse_args()
    if namespace.exel != '':
        print(namespace.exel)
        data_xls = getDataExel(namespace.exel)
        print(data_xls)
        for row in data_xls.itertuples():
            namespace.name = row.Project
            namespace.job = row.Job
            namespace.length = row.Length
            namespace.date = str(row.Date)[:-9]
            namespace.time = str(row.Time)
            namespace.fibertype = row.Cable
            namespace.trace = row.Fibers
            namespace.flag = row.Flag
            print(namespace)
            run(namespace)
            print("All Done!")
    else:
        run(namespace)
    quiet = ('-q' in sys.argv)

# -d 11.10.2021 -t 07:34:47 --trace 6 -ft Multimode -l 0.1452 -j Isracard -c Floor_(-3)_Floor_5_RED