import pandas as pd
import os
import xlsxwriter


#d = os.path.dirname(__file__)
import sys

if getattr(sys, 'frozen', False):
    # frozen
    d = os.path.dirname(sys.executable)
else:
    # unfrozen
    d = os.path.dirname(os.path.realpath(__file__))
p = r'{}/data'.format(d)

path = os.path.join(p, "model.xlsx")
out_path = os.path.join(p, "schedule.xlsx")
duration_path = os.path.join(p, "durations.xlsx")
musteri_path = os.path.join(p, "musteri.xlsx")
cache_path = os.path.join(p, "cache.xlsx")


def get_customer_ids(cluster_id):
    sheet_name = "Cluster" + str(cluster_id)
    dur = pd.read_excel(duration_path, header=None, sheet_name=sheet_name)
    id_list = list(dur[0])
    id_list = id_list[2:len(id_list)-1]
    for i in range(len(id_list)):
        id_list[i] = int(id_list[i][1:])

    return id_list


def get_adresses(ids):
    musteri = pd.read_excel(musteri_path)
    adresses = []
    names = []
    for i in ids:
        adresses.append(list(musteri[musteri['MÜŞTERİ KODU'] == i]['ADRES']))
        names.append(list(musteri[musteri['MÜŞTERİ KODU'] == i]['Müşteri Grubu']))
    return adresses, names


def write_schedule(cluster_id, ws, cache_wb):

    sheet_name = "Cluster" + str(cluster_id)
    data = pd.read_excel(path, sheet_name=sheet_name, header=None)

    cache_ws = cache_wb.add_worksheet(sheet_name)
    cache_row = 1
    cache_col = 0

    musteri = pd.read_excel(musteri_path)
    data = data.fillna('XXX')

    # wb = xlsxwriter.Workbook(out_path)
    # ws = wb.add_worksheet()

    id_list = get_customer_ids(cluster_id=cluster_id)
    cluster_len = len(id_list)
    addresses, names = get_adresses(id_list)
    daily_ordered = order_daily(data=data, cluster_len=cluster_len)


    ws.write(0, 0, 'Customer ID')
    ws.write(0, 1, "Monday")
    ws.write(0, 2, "Tuesday")
    ws.write(0, 3, "Wednesday")
    ws.write(0, 4, "Thursday")
    ws.write(0, 5, "Friday")
    ws.write(0, 6, "Saturday")

    ws.write(0, 8, "Customer ID")
    ws.write(0, 9, "Customer Name")
    ws.write(0, 10, "Adress")

    row = 1
    col = 0

    for i in id_list:
        ws.write(row, col, i)
        ws.write(row, 8, i)
        row += 1

    row = 1

    for i in addresses:
        ws.write(row, 10, i[0])
        row += 1

    row = 1
    for i in names:
        ws.write(row, 9, i[0])
        row += 1


    for day in range(1, 7):
        for mus in range(1, len(id_list) + 1):
            search_string_g = "g(" + str(mus) + ", " + str(day) + ")"
            search_string_s = "S(" + str(mus) + ", " + str(day) + ")"
            search_string_extra = "extra(" + str(day) + ")"

            search_string_n_visit = 'z' + str(day)
            n_visit = list(data[data[8] == search_string_n_visit][9])[0]

            search_string_def = 'def' + str(day)
            if data[data[10] == search_string_def].empty:
                deficit = 0
            if not data[data[10] == search_string_def].empty:
                deficit = list(data[data[10] == search_string_def][11])[0]

            if (data[data[6] == search_string_extra]).empty:
                extra = 0
            if not (data[data[6] == search_string_extra]).empty:
                extra = list(data[data[6] == search_string_extra][7])[0]

            if not (data[data[4] == search_string_g]).empty:
                coef = daily_ordered[day].index(mus)
                start_min = list(data[data[2] == search_string_s][3])[0] + (coef)* (round(deficit/n_visit)) - coef*(round(extra/n_visit))

                servis_suresi = list(musteri[musteri['MÜŞTERİ KODU'] == id_list[mus - 1]]['Servis Süresi'])[0]
                servis_suresi = servis_suresi + round(deficit / n_visit) - round(extra/n_visit)

                cache_ws.write(cache_row, cache_col, id_list[mus-1])
                cache_ws.write(cache_row, cache_col+1, servis_suresi)
                cache_row +=1

                basla = format_min_to_time(start_min)
                bitir = format_min_to_time(start_min + servis_suresi)
                ws.write(mus, day, basla + " - " + bitir)
    # wb.close()
    return

def write_abstract(cluster_id, ws):
    sheet_name = "Cluster" + str(cluster_id)
    data = pd.read_excel(path, sheet_name=sheet_name, header=None)
    musteri = pd.read_excel(musteri_path)

    id_list = get_customer_ids(cluster_id=cluster_id)
    addresses, names = get_adresses(id_list)

    ws.write(0, 0, 'Customer ID')
    ws.write(0, 1, "Monday")
    ws.write(0, 2, "Tuesday")
    ws.write(0, 3, "Wednesday")
    ws.write(0, 4, "Thursday")
    ws.write(0, 5, "Friday")
    ws.write(0, 6, "Saturday")

    ws.write(0, 8, "Customer ID")
    ws.write(0, 9, "Customer Name")
    ws.write(0, 10, "Adress")

    row = 1
    col = 0

    for i in id_list:
        ws.write(row, col, i)
        ws.write(row, 8, i)
        row += 1

    row = 1

    for i in addresses:
        ws.write(row, 10, i[0])
        row += 1

    row = 1
    for i in names:
        ws.write(row, 9, i[0])
        row += 1

    return


def format_min_to_time(start_min, mesai_baslangici=8):
    start_min += 60 * mesai_baslangici
    minute = start_min % 60
    if minute < 10:
        min_str = "0" + str(minute)
    else:
        min_str = str(minute)

    hour = int(start_min/60)
    if hour < 10:
        hour_str = "0" + str(hour)
    else:
        hour_str = str(hour)

    time_string = hour_str + ":" + min_str

    return time_string

def order_daily(data, cluster_len):
    git = {1: [], 2: [], 3: [], 4: [], 5: [], 6: []}
    for p in range(1, 7):
        for i in range(1, 9):
            search_string_g = "g(" + str(i) + ", " + str(p) + ")"
            p_str = str(p)
            if not data[data[4] == search_string_g].empty:
                git[p].append(i)

    ordered_git = {}


    for p in range(1, 7):
        ordered_git[p] = [-1] * len(git[p])

    for p in range(1, 7):
        if len(ordered_git[p])>0:
            for i in git[p]:

                bas = x(0, i, p)
                bit = x(i, cluster_len+1, p) #  TODO: 9 değişecek
                if not (data[data[0] == bas]).empty:
                    ordered_git[p][0] = i
                elif not (data[data[0] == bit]).empty:
                    leng = len(git[p])
                    ordered_git[p][leng - 1] = i

    for p in range(1, 7):
        if len(ordered_git[p])>1:
            for i in git[p]:
                bas = ordered_git[p][0]
                bit = ordered_git[p][len(git[p]) - 1]

                if not (data[data[0] == x(bas, i, p)]).empty:
                    ordered_git[p][1] = i
                elif not (data[data[0] == x(i, bit, p)]).empty:
                    leng = len(git[p])
                    ordered_git[p][leng - 2] = i


    for p in range(1, 7):
        if len(ordered_git[p]) > 2:
            for i in git[p]:
                bas = ordered_git[p][1]
                bit = ordered_git[p][len(git[p]) - 2]

                if not (data[data[0] == x(bas, i, p)]).empty:
                            ordered_git[p][2] = i
                elif not (data[data[0] == x(i, bit, p)]).empty:
                    leng = len(git[p])
                    ordered_git[p][leng - 3] = i


    for p in range(1, 7):
        if len(ordered_git[p]) > 3:
            for i in git[p]:
                print(p)
                bas = ordered_git[p][2]
                bit = ordered_git[p][len(git[p]) - 3]

                if not (data[data[0] == x(bas, i, p)]).empty:
                    ordered_git[p][3] = i
                elif not (data[data[0] == x(i, bit, p)]).empty:
                    leng = len(git[p])
                    ordered_git[p][leng - 4] = i

    return ordered_git

def x(i, j, p):
    return 'x(' + str(i) + ", " + str(j) + ", " + str(p) + ")"


#write_schedule(19)
