import pandas as pd
from math import sin, cos, sqrt, atan2, radians
import xlsxwriter
import os
import sys

if getattr(sys, 'frozen', False):
    # frozen
    d = os.path.dirname(sys.executable)
else:
    # unfrozen
    d = os.path.dirname(os.path.realpath(__file__))

p = r'{}/data'.format(d)
data_path = os.path.join(p, "musteri.xlsx")
data = pd.read_excel(data_path)
data.index += 1
simp = data[['MÜŞTERİ KODU', 'ENLEM', 'BOYLAM', 'Ortalama Ciro']]


def calculate_distance(lat1, long1, lat2, long2):
    R = 6371
    lat1 = radians(lat1)
    long1 = radians(long1)
    lat2 = radians(lat2)
    long2 = radians(long2)

    dlong = long2 - long1
    dlat = lat2 - lat1

    a = sin(dlat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(dlong / 2) ** 2
    c = 2 * atan2(sqrt(a), sqrt(1 - a))

    distance = R * c

    return distance


def prepare_output_file():
    wb = xlsxwriter.Workbook("Distances.xlsx")
    ws = wb.add_worksheet("Sheet1")

    row = 1
    col = 0

    for i in simp['MÜŞTERİ KODU']:
        ws.write(row, col, i)
        row = row + 1

    row = 0
    col = 1

    for i in simp['MÜŞTERİ KODU']:
        ws.write(row, col, i)
        col = col + 1

    row, col = 1, 1

    for i in simp['MÜŞTERİ KODU']:
        lat1 = simp[simp['MÜŞTERİ KODU'] == i]['ENLEM']
        long1 = simp[simp['MÜŞTERİ KODU'] == i]['BOYLAM']

        for j in simp['MÜŞTERİ KODU']:
            lat2 = simp[simp['MÜŞTERİ KODU'] == j]['ENLEM']
            long2 = simp[simp['MÜŞTERİ KODU'] == j]['BOYLAM']
            dist = calculate_distance(lat1, long1, lat2, long2)
            ws.write(row, col, dist)
            col += 1

        col = 1
        row += 1

    wb.close()
    return

# prepare_output_file()
