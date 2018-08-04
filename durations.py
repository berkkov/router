import pandas as pd
import xlsxwriter
import requests
import os
import distances
import sys
import json


if getattr(sys, 'frozen', False):
    # frozen
    d = os.path.dirname(sys.executable)
else:
    # unfrozen
    d = os.path.dirname(os.path.realpath(__file__))

p = r'{}/data'.format(d)
API_KEY = '[YOUR_TRAFI_API_KEY]'  #Trafi API key - visit trafi.com

datapath = os.path.join(p, "clusters.xlsx")


def construct_url(start_lat, start_lng, end_lat, end_lng, api_key = API_KEY):
    url = "http://api-ext.trafi.com/routes?start_lat=" + str(start_lat) + "&start_lng=" + str(start_lng) + "&end_lat=" + \
    str(end_lat) + "&end_lng=" + str(end_lng) + "&time=2018-04-0412%3A00" "&is_arrival=false&api_key=" + API_KEY
    return url


def get_duration(start_lat, start_lng, end_lat, end_lng):
    r = requests.get(construct_url(start_lat, start_lng, end_lat, end_lng))

    try:
        if r.json()['Routes']:
            duration = r.json()['Routes'][0]['DurationMinutes']
            if duration > 120:
                duration = round(round(distances.calculate_distance(start_lat, start_lng, end_lat, end_lng)) * (1.7))
        else:
            duration = round(distances.calculate_distance(start_lat, start_lng, end_lat, end_lng) * 1.5)

    except json.decoder.JSONDecodeError:
        print("JSONDecodeError caught between: " + str(start_lng, start_lat) + " and " + str(end_lat, end_lng) + "\n")
        duration = round(distances.calculate_distance(start_lat, start_lng, end_lat, end_lng) * 1.5)
    return duration


def prepare_duration_matrices():

    data = pd.read_excel(datapath)
    cluster_ids = list(data['Clusters'].unique())

    wb = xlsxwriter.Workbook(os.path.join(p, "durations.xlsx"))
    for i in cluster_ids:

        tempdata = data[data['Clusters'] == i]
        zipped = zip(tempdata['Şube Kodu'], tempdata['ENLEM'], tempdata['BOYLAM'])
        l = list(zipped)
        #wb = xlsxwriter.Workbook('C:\\Users\\berkm\\OneDrive\\Masaüstü\\cluster durations\\cluster' + str(i) + '.xlsx')
        ws = wb.add_worksheet('Cluster' + str(i))
        row = 0
        col = 0
        ws.write(row, col, "Yolda geçen süre (dk)")
        ws.write(row+1, col, 'dummy')
        for a in l:
            ws.write(row + 2, col, 'm' + str(a[0]))
            row += 1

        ws.write(row+2, col, 'dummy')

        row = 0

        ws.write(row, col+1, 'dummy')

        for a in l:
            ws.write(row, col + 2, 'm' + str(a[0]))
            col += 1

        ws.write(row, col+2, 'dummy')

        for r in range(0, len(l)):
            ws.write(r + 2, r + 2, 0)



        for j in range(0, len(l)):
            for k in range(j + 1, len(l)):


                dur = get_duration(l[j][1], l[j][2], l[k][1], l[k][2])
                print(l[j][0], l[k][0], "Duration :" + str(dur))
                ws.write(j + 2, k + 2, dur)
                ws.write(k + 2, j + 2, dur)
        #wb.close()

        row = 1
        col = 1

        for j in range(0,len(l)+2):
            ws.write(row, j+1, 0)
            ws.write(row + len(l) + 1, j+1 ,0)
            ws.write(j+1, row, 0)
            ws.write(j+1, row+len(l)+1, 0)

    wb.close()

#prepare_duration_matrices()
