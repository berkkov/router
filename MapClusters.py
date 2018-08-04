import folium
import pandas as pd
import os

from pathlib import Path

#d = os.path.dirname(__file__)
import sys

if getattr(sys, 'frozen', False):
    # frozen
    d = os.path.dirname(sys.executable)
else:
    # unfrozen
    d = os.path.dirname(os.path.realpath(__file__))

pa = r'{}/data'.format(d)
output_path = os.path.join(pa,"/Gruplar_yeni.html")
#daa = Path(__file__).resolve().parents[1]
#print(daa)
def map_clusters():

    data_path = os.path.join(pa,"clusters_yeni.xlsx")
    data = pd.read_excel(data_path)
    lats = list(data['ENLEM'])
    lngs = list(data['BOYLAM'])
    clusters = list(data['Clusters'])
    sube_kodu = list(data['Şube Kodu'])
    color_domain = ['red', 'blue', 'green', 'purple', 'orange', 'darkred',
                 'lightred', 'beige', 'darkblue', 'darkgreen', 'cadetblue',
                 'darkpurple', 'white', 'pink', 'lightblue', 'lightgreen',
                 'gray', 'black', 'lightgray']

    m = folium.Map(location=[lats[1], lngs[1]], zoom_start=8)
    for i in range(0,len(lats)):
        folium.Marker([lats[i], lngs[i]], popup = str("Şube kodu: " + str(sube_kodu[i]) + "\nKısa kod: " + str(i) + "\nTTE: " + str(clusters[i])), icon = folium.Icon(color=color_domain[clusters[i]%(len(color_domain)-1)])).add_to(m)

    m.save("Gruplar_eski.html")
    return

map_clusters()
