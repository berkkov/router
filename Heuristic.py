import pandas as pd
import xlsxwriter
import os
import sys
import random
import operator

if getattr(sys, 'frozen', False):
    # frozen
    d = os.path.dirname(sys.executable)
else:
    # unfrozen
    d = os.path.dirname(os.path.realpath(__file__))

pa = r'{}/data'.format(d)

musteri_path = os.path.join(pa, "clusters.xlsx")
duration_path = os.path.join(pa, "durations.xlsx")

K_ARRAY = [[1, 2, 3, 4, 5, 6], [[1, 3], [1, 4], [1, 5], [1, 6], [2, 4], [2, 5], [2, 6], [3, 5], [3, 6], [4, 6]],
           [[1, 2, 4], [1, 2, 5], [1, 2, 6], [1, 3, 4], [1, 3, 5], [1, 4, 5], [1, 3, 6], [1, 4, 6], [1, 5, 6],
            [2, 3, 5],
            [2, 3, 6], [2, 4, 5], [2, 4, 6], [2, 5, 6], [3, 4, 6], [3, 5, 6]],
           [[1, 2, 3, 5], [1, 2, 3, 6], [1, 2, 4, 6], [1, 2, 5, 6], [1, 3, 4, 5], [1, 3, 5, 6], [1, 4, 5, 6],
            [2, 3, 4, 6], [2, 3, 5, 6], [2, 4, 5, 6], [1, 3, 4, 6]],
           [[1, 2, 3, 4, 5], [1, 2, 3, 4, 6], [1, 2, 3, 5, 6], [1, 2, 4, 5, 6], [1, 3, 4, 5, 6], [2, 3, 4, 5, 6]],
           [[1, 2, 3, 4, 5, 6]]]


def get_cluster_detail(cluster_no):
    m_data = pd.read_excel(musteri_path)
    members = m_data[m_data['Clusters'] == cluster_no]

    cluster_str = "Cluster" + str(cluster_no)
    durs = pd.read_excel(duration_path, cluster_str)
    durs = durs.drop(['dummy', 'dummy.1'], axis=1)
    durs = durs.drop(0)
    durs = durs.drop(len(durs))
    return members, durs


def frame_writer(cluster_id, ws):
    mems, durs = get_cluster_detail(cluster_id)
    id_list = list(mems['Şube Kodu'])

    ws.write(0, 0, 'Müşteri Kodu')
    ws.write(0, 1, "Pazartesi")
    ws.write(0, 2, "Salı")
    ws.write(0, 3, "Çarşamba")
    ws.write(0, 4, "Perşembe")
    ws.write(0, 5, "Cuma")
    ws.write(0, 6, "Cumartesi")

    ws.write(0, 8, "Müşteri Kodu")
    ws.write(0, 9, "Müşteri")
    ws.write(0, 10, "Adres")

    row = 1
    col = 0

    for i in id_list:
        ws.write(row, col, i)
        ws.write(row, 8, i)
        row += 1

    return


def initializer(cluster_no):
    visits_left = {}
    pattern_track = {}
    days_track = {}
    mems, durs = get_cluster_detail(cluster_no)
    freqs = list(mems['Sıklık'])
    j = 0

    for day in range(1, 7):
        days_track[day] = []

    for i in freqs:
        visits_left[j] = i
        pattern_index = random.randint(0, len(K_ARRAY[i - 1]) - 1)
        pattern_track[j] = K_ARRAY[i - 1][pattern_index]
        j += 1

    for day in range(1, 7):
        for i in range(0, len(freqs)):
            if day in pattern_track[i]:
                days_track[day].append(i)

    stimes = list(mems['Servis Süreleri'])

    return freqs, stimes, visits_left, pattern_track, days_track


def sequencer(durs, days_track):
    visit_sequence = {}
    durs_list = durs.values.tolist()
    for day in range(1, 7):
        visit_sequence[day] = []

    print(durs_list)
    for day in range(1, 7):
        totals = {}
        latest = None
        for j in days_track[day]:
            visits = days_track[day]
            totals[j] = 0
            print(j)
            for i in visits:
                print(i)
                totals[j] += durs_list[j][i + 1]

        print(days_track)
        print(totals)
        first_visit = max(totals.items(), key=operator.itemgetter(1))[0]
        latest = first_visit
        visit_sequence[day].append(first_visit)
        days_track[day].remove(first_visit)

        while days_track[day]:
            next_visit = None
            min_dist = 9999
            min_dist_cust = None

            for i in days_track[day]:
                dist = durs_list[latest][i + 1]

                if dist < min_dist:
                    min_dist = dist
                    min_dist_cust = i

            next_visit = min_dist_cust
            visit_sequence[day].append(next_visit)
            days_track[day].remove(next_visit)
            latest = next_visit

    return visit_sequence


def sequence_writer(visit_sequence, durs, ws, cluster_id):
    freqs, stimes, visits_left, pattern_track, days_track = initializer(cluster_id)
    durs_list = durs.values.tolist()
    for day in range(1, 7):
        col = day

        total = 0
        transport_count = 0
        for i in visit_sequence[day]:
            total += stimes[i]
            if i != visit_sequence[day][0]:
                total += durs_list[visit_sequence[day][transport_count - 1]][i + 1]
            transport_count += 1

        time_deficit = 450 - total
        deficit_share = int(time_deficit / len(visit_sequence[day]))

        start = 0
        count = 0

        for vis in visit_sequence[day]:
            if vis == visit_sequence[day][0]:
                start = 0
                visit_start = start
            else:
                visit_start = visit_end + durs_list[visit_sequence[day][count - 1]][vis + 1]

            row = vis + 1

            time_spent = stimes[vis] + deficit_share
            visit_end = visit_start + time_spent
            ws.write(row, col, format_min_to_time(visit_start) + " - " + format_min_to_time(visit_end))
            count += 1
    return


def format_min_to_time(start_min, mesai_baslangici=8):
    start_min += 60 * mesai_baslangici
    minute = start_min % 60
    if minute < 10:
        min_str = "0" + str(minute)
    else:
        min_str = str(minute)

    hour = int(start_min / 60)
    if hour < 10:
        hour_str = "0" + str(hour)
    else:
        hour_str = str(hour)

    time_string = hour_str + ":" + min_str

    return time_string


def sanity_check(days_track):
    switch = True
    for day in range(1, 7):
        if len(days_track[day]) < 2:
            switch = False
    return switch


def fix_sanity_failure(days_track, freqs):
    switch = sanity_check(days_track)

    while not switch:
        for day in range(1, 7):
            if len(days_track[day]) < 2:
                or_visits = days_track[day]
                for cust in range(0, len(freqs)):
                    if cust not in or_visits:
                        days_track[day].append(cust)
                        freqs[cust] += 1
                        break
        switch = sanity_check(days_track)

    return switch


def heuristic(cluster_id, ws):
    frame_writer(cluster_id, ws)
    freqs, stimes, visits_left, pattern_track, days_track = initializer(cluster_id)

    SANITY_CHECK = sanity_check(days_track)
    print(days_track)
    print(SANITY_CHECK)

    if not SANITY_CHECK:
        fix_sanity_failure(days_track, freqs)

    print(sanity_check(days_track))
    print(days_track)
    mems, durs = get_cluster_detail(cluster_id)
    sq = sequencer(durs, days_track)
    sequence_writer(sq, durs, ws, cluster_id)
    return
