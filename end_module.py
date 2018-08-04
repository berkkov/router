import model2
import xlsxwriter
import os
import ScheduleWriter
import pandas as pd
import ModelStrict
import sys
import Heuristic
from pulp.constants import PulpError

if getattr(sys, 'frozen', False):
    # frozen
    d = os.path.dirname(sys.executable)
else:
    # unfrozen
    d = os.path.dirname(os.path.realpath(__file__))


pa = r'{}/data'.format(d)
datapath = os.path.join(pa, "clusters.xlsx")
STRICT = False


def solve_all():
    data = pd.read_excel(datapath)
    cluster_ids = list(data.Clusters.unique())
    wb = xlsxwriter.Workbook(os.path.join(model2.pa, "model.xlsx"))
    schedule_wb = xlsxwriter.Workbook(os.path.join(model2.pa, "Ziyaret Planları.xlsx"))
    cache_wb = xlsxwriter.Workbook(os.path.join(model2.pa, "cache.xlsx"))
    solved = []

    for i in cluster_ids:

            if i in []:
                continue

            print("Program şu anda " + str(i) + ". problemi çözüyor. Toplam problem sayısı: " + str(len(cluster_ids)))
            ws = wb.add_worksheet("Cluster" + str(i))
            try:
                if not STRICT:
                    problem = model2.problem(data=data, cluster_no=i)
                if STRICT:
                    problem = ModelStrict.problem(data=data, cluster_no=i)
            except PulpError:
                print("PulpError caught.")

            # print(problem.frequencies)
            # print(problem.durations)
            # print(problem.service_times)
            try:
                solution = problem.solve()
                model2.prepare_output_file(solution=solution, ws=ws)
                solved.append(i)
            except IndexError:
                print("error1")
            except UnboundLocalError:
                print("error2")
            except KeyError:
                print("error3")

    wb.close()

    for i in cluster_ids:
        if i in []:
            continue

        try:
            print("Problem " + str(i))
            ws_s = schedule_wb.add_worksheet("TTE" + str(i))
            ScheduleWriter.write_schedule(i, ws=ws_s, cache_wb=cache_wb)
            ws_check_name = "Cluster" + str(i)
            check_empty = pd.read_excel(os.path.join(model2.pa, "model.xlsx"), ws_check_name)

            if check_empty.empty:
                Heuristic.heuristic(cluster_id=i, ws=ws_s)

        except TypeError:
            print(i)
            print("Type Error caught")
            ScheduleWriter.write_abstract(i, ws=ws_s)

            try:
                Heuristic.heuristic(cluster_id=i, ws=ws_s)
            except ValueError:
                print("Value Error for Group: " + str(i))
            except KeyError:
                print("Key Error for Group: " + str(i))
            except IndexError:
                print("Index Error for Group: " + str(i))

        except KeyError:
            print(i)
            print("Key Error caught")
            ScheduleWriter.write_abstract(i, ws=ws_s)

            try:
                Heuristic.heuristic(cluster_id=i, ws=ws_s)
            except ValueError:
                print("Value Error for Group: " + str(i))
            except KeyError:
                print("Key Error for Group: " + str(i))
            except IndexError:
                print("Index Error for Group: " + str(i))

        except IndexError:
            print(i)
            print("Index Error caught")
            ScheduleWriter.write_abstract(i, ws=ws_s)

            try:
                Heuristic.heuristic(cluster_id=i, ws=ws_s)
            except ValueError:
                print("Value Error for Group: " + str(i))
            except KeyError:
                print("Key Error for Group: " + str(i))
            except IndexError:
                print("Index Error for Group: " + str(i))

    cache_wb.close()
    schedule_wb.close()
    return

#solve_all()
'''
wb = xlsxwriter.Workbook("new.xlsx")
for i in range(0,26):
    ws = wb.add_worksheet("New" + str(i))
    Heuristic.heuristic(18, ws)

wb.close()
'''
