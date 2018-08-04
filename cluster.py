import pulp
import random
from math import radians, sin, cos, atan2, sqrt
import pandas as pd
import xlsxwriter
from time_calculation import prepare_file
import os
import sys

NUM_ITER = 1
MAX_WEIGHT_CIRO = 70000
MIN_WEIGHT_CIRO = 35000
MAX_WEIGHT_TIME = 2000
MIN_WEIGHT_TIME = 1500
TTE_SAYISI = 25
CIRO_LIMITED = False
MAX_ITERS = 9
MAX_CUST = 8
#toplam süre: 39575  BURAYI OTOMATİKLEŞTİR

#d = os.path.dirname(__file__)
if getattr(sys, 'frozen', False):
    # frozen
    d = os.path.dirname(sys.executable)
else:
    # unfrozen
    d = os.path.dirname(os.path.realpath(__file__))

pa = r'{}/data'.format(d)
data_path = os.path.join(pa, "musteri.xlsx")
output_path = os.path.join(pa, "Clusters.xlsx")

total_time_col = "Toplam Süre"
latitude_col = "ENLEM"
longitude_col = "BOYLAM"
revenue_col = "Ortalama Ciro"

TIME_LIMIT = 900

# Cluster limitlerini belirleme kriteri

if CIRO_LIMITED:
    MAX_WEIGHT = MAX_WEIGHT_CIRO
    MIN_WEIGHT = MIN_WEIGHT_CIRO
else:
    MAX_WEIGHT = MAX_WEIGHT_TIME
    MIN_WEIGHT = MIN_WEIGHT_TIME


def prepare_data(ciro_limited=CIRO_LIMITED):

    data = pd.read_excel(data_path)
    if not ciro_limited:

        data_copy = data
        temp = data_copy[data_copy['Toplam Süre'] > MAX_WEIGHT]
        temp_ind = data_copy[data_copy['Toplam Süre'] > MAX_WEIGHT].index
        temp_ind_list = list(temp_ind)
        extra_len = temp_ind_list
        data_copy.at[temp_ind, 'Toplam Süre'] = MAX_WEIGHT - 120
        weights = list(data_copy['Toplam Süre'])
        coordinates = list(zip(data_copy['ENLEM'], data_copy['BOYLAM']))

    if ciro_limited:
        data_copy = data
        temp = data_copy[data_copy[revenue_col] > MAX_WEIGHT]
        temp_ind = data_copy[data_copy[revenue_col] > MAX_WEIGHT].index
        temp_ind_list = list(temp_ind)
        extra_len = len(temp_ind_list)
        data_copy.at[temp_ind,revenue_col] = MAX_WEIGHT-10000
        weights = list(data_copy[revenue_col])
        coordinates = list(zip(data_copy['ENLEM'], data_copy['BOYLAM']))


    return coordinates, weights


def calculate_distance(lat1, long1, lat2, long2):
    '''
    İki nokta arasındaki uzaklığı hesaplar.
    :param lat1: 1.Noktanın Enlemi
    :param long1: 1.Noktanın Boylamı
    :param lat2: 2.Noktanın Enlemi
    :param long2: 2.Noktanın Boylamı
    :return:
    '''
    R = 6371
    lat1 = radians(lat1)
    long1 = radians(long1)
    lat2 = radians(lat2)
    long2 = radians(long2)

    dlong = long2 - long1
    dlat = lat2 - lat1

    a = sin(dlat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(dlong / 2) ** 2
    c = 2 * atan2(sqrt(a), sqrt(1 - a))

    distance = round(R * c, 2)

    return distance**2


class subproblem(object):
    def __init__(self, centroids, data, weights, min_weight, max_weight, tmlim=TIME_LIMIT):

        self.centroids = centroids
        self.data = data
        self.weights = weights
        self.min_weight = min_weight
        self.max_weight= max_weight
        self.n = len(data)
        self.k = len(centroids)
        self.create_model()
        self.time_limit = tmlim

    def create_model(self):
        def distances(assignment):
            return calculate_distance(self.data[assignment[0]][0], self.data[assignment[0]][1], self.centroids[assignment[1]][0], self.centroids[assignment[1]][1])

        assignments = [(i, j) for i in range(self.n) for j in range(self.k)]

        # assignment variables
        self.y = pulp.LpVariable.dicts('data-to-cluster assignments',
                                       assignments,
                                       lowBound=0,
                                       upBound=1,
                                       cat=pulp.LpInteger)


        self.model = pulp.LpProblem("Model for assignment subproblem", pulp.LpMinimize)

        # objective function
        self.model += pulp.lpSum([distances(assignment) * self.y[assignment] for assignment in assignments]), 'Objective Function - sum weighted squared distances to assigned centroid'

        for j in range(self.k):
            self.model += pulp.lpSum([self.weights[i] * self.y[(i, j)] for i in
                                      range(self.n)]) >= self.min_weight, "minimum weight for cluster {}".format(j)
            self.model += pulp.lpSum([self.weights[i] * self.y[(i, j)] for i in
                                      range(self.n)]) <= self.max_weight, "maximum weight for cluster {}".format(j)

        # make sure each point is assigned at least once, and only once
        for i in range(self.n):
            self.model += pulp.lpSum([self.y[(i, j)] for j in range(self.k)]) == 1, "must assign point {}".format(i)

        # Optional

        for j in range(self.k):
            self.model += pulp.lpSum([self.y[(i,j)] for i in range(self.n)]) <= MAX_CUST

    def solve(self, tmlim):
        self.status = self.model.solve(pulp.GLPK_CMD(options=['--tmlim', str(tmlim)]))

        clusters = None

        if self.status == 1:
            clusters = [-1 for i in range(self.n)]
            for i in range(self.n):
                for j in range(self.k):
                    if self.y[(i, j)].value() > 0:
                        clusters[i] = j
        return clusters


def initialize_centers(dataset, k):
    """
    Farthest point initialization
    :param dataset: Tüm data
    :param k: TTE elemanı sayısı
    :return: (lat,long) şeklinde gruplandırılmış centroid koordinatları
    """

    #coordinates = list(zip(dataset['ENLEM'], dataset['BOYLAM'])) #TODO: buradan hata gelebilir
    coordinates = dataset
    centroids = []
    indices = []
    ind = random.randint(0, len(dataset))
    centroids.append(coordinates[ind])
    indices.append(ind)

    for i in range(k-1): #-1 maybe
        print(i)
        max_sum = 0
        best_cand = 99999
        for cand in range(len(coordinates)):
            if (cand in indices):
                continue
            sum = 0
            for cent in centroids:
                dist = calculate_distance(coordinates[cand][0], coordinates[cand][1], cent[0], cent[1])
                if (dist < 3):
                    sum -= 999999
                sum += dist

            if (sum > max_sum):
                max_sum = sum
                best_cand = cand

        centroids.append(coordinates[best_cand])
        indices.append(best_cand)

    return centroids

def compute_centers(clusters, dataset, weights=None):
    """
    Ciroya göre ağırlıklı orta nokta bulur
    :param clusters:
    :param dataset: coordinates
    :param weights: cirolar
    :return:
    """


    if weights is None:
        weights = [1]*len(dataset)

    # canonical labeling of clusters - {cluster numaraları : cluster_id } dicti oluştur
    ids = list(set(clusters))
    c_to_id = dict()
    for j, c in enumerate(ids):
        c_to_id[c] = j

    for j, c in enumerate(clusters):
        clusters[j] = c_to_id[c]

    k = len(ids)                                        # TTE sayısını cluster sayısından çek
    dim = len(dataset[0])                               # coordinate kısmından boyunu al (2 : (lat,long))
    cluster_centers = [[0.0] * dim for i in range(k)]   # cluster centerları için storage'ı düzgün formatla başlat
    cluster_weights = [0] * k                           # cluster weight(ciro toplamı) için düzgün düzgün format

    # her cluster için elemanlarının lat*weight ve long*weight toplamlarını hesapla
    for j, c in enumerate(clusters):
        for i in range(dim):
            cluster_centers[c][i] += dataset[j][i] * weights[j]
        cluster_weights[c] += weights[j]

    # Toplamları total cluster weight'ine bölüp koordinatları bul.
    for j in range(k):
        for i in range(dim):
            cluster_centers[j][i] = cluster_centers[j][i]/float(cluster_weights[j])
    return clusters, cluster_centers


# Bu function'a girecek data list'e dönüştürülmeli ve sadece koordinatları içermeli
def minsize_kmeans_weighted(dataset, k, weights=None, min_weight=MIN_WEIGHT, max_weight=MAX_WEIGHT, max_iters=MAX_ITERS, uiter=None, tmlim=TIME_LIMIT):
    """
    @dataset - numpy matrix (or list of lists) - of point coordinates
    @k - number of clusters
    @weights - list of point weights, length equal to len(@dataset)
    @min_weight - minimum total weight per cluster
    @max_weight - maximum total weight per cluster
    @max_iters - if no convergence after this number of iterations, stop anyway
    @uiter - iterator like tqdm to print a progress bar.
    """

    n = len(dataset)
    if weights is None:
        weights = [-1]*n
    if max_weight == None:
        max_weight = sum(weights)

    uiter = uiter or list

    centers = initialize_centers(dataset, k)
    clusters = [-1] * n                         # (1xn) lik liste. Her müşterinin ait olduğu clusterlar yazıyor
    centers_ = [-1] * TTE_SAYISI
    quality_bench = 99999

    # Iterations
    for ind in uiter(range(max_iters)):
        print(ind)
        m = subproblem(centers, dataset, weights, min_weight, max_weight, tmlim=tmlim)   # LP Model oluşturma
        clusters_ = m.solve(tmlim=tmlim)                                               # Müşteri assignmentları bulma (Model çözümü)
        if not clusters_:
            return None, None

        quality = compute_quality(dataset, clusters_)
        print(quality)
        if quality < quality_bench:
            clusters_, centers = compute_centers(clusters_, dataset)            # Cluster merkezi hesaplamaları

            converged = all([clusters[i] == clusters_[i] for i in range(n)])
            clusters = clusters_

            alt_converged = all([centers[j] == centers_[j] for j in range(TTE_SAYISI)])
            centers_ = centers

            quality = quality_bench

            if converged:
                break
            if alt_converged:
                break


    return clusters, centers


def cluster_quality(cluster):
    '''
    Tek bir cluster için cluster kalitesini hesaplar. Kalite = her elemanın cluster içindeki diğer elemanlarla arasındaki
    mesafelerin kareleri toplamı.
    :param cluster: Kalitesi hesaplanacak cluster numarası
    :return: Cluster kalitesi
    '''
    if len(cluster) == 0:
        return 0.0

    quality = 0.0
    for i in range(len(cluster)):
        for j in range(i, len(cluster)):
            quality += calculate_distance(cluster[i][0],cluster[i][1], cluster[j][0], cluster[j][1])
    return quality / len(cluster)


def compute_quality(data, cluster_indices):
    '''
    Tüm modül için kalite hesaplar.
    :param data:
    :param cluster_indices:
    :return:
    '''

    # Cluster içindeki noktaların koordinatlarını gösteren dict oluştur
    clusters = dict()
    for i, c in enumerate(cluster_indices):
        if c in clusters:
            clusters[c].append(data[i])
        else:
            clusters[c] = [data[i]]
    return sum(cluster_quality(c) for c in clusters.values())

#clusters, centers = minsize_kmeans_weighted(coordinates, 25, weights, min_weight=30000, max_weight=90000, max_iters=30, uiter=None)

def run_and_write(ciro_limited=CIRO_LIMITED, MIN_WEIGHT= MIN_WEIGHT, MAX_WEIGHT=MAX_WEIGHT, TTE_SAYISI=TTE_SAYISI, max_iters=MAX_ITERS, tmlim=TIME_LIMIT):

    TTE_SAYISI = TTE_SAYISI + 1
    coordinates, weights = prepare_data(ciro_limited=ciro_limited)
    best = None
    best_clusters = None
    for i in range(NUM_ITER):
        print("TOTAL ITERATION no:" + str(i + 1) + "/" + str(NUM_ITER))
        clusters, centers = minsize_kmeans_weighted(coordinates, k=TTE_SAYISI , weights=weights, min_weight=MIN_WEIGHT, max_weight=MAX_WEIGHT, max_iters=max_iters , uiter=None, tmlim=tmlim)

        if clusters:
            quality = compute_quality(coordinates, clusters)
            if not best or (quality < best):
                best = quality
                best_clusters = clusters

    if best:
        print('cluster assignments:')
        for i in range(len(best_clusters)):
            print('%d: %d'%(i, best_clusters[i]))
        print('sum of squared distances: %.4f'%(best))
    else:
        print('no clustering found')

    try:
        row = 1
        col = 0
        wb = xlsxwriter.Workbook(output_path)
        ws = wb.add_worksheet('Sheet1')
        ws.write(0, 0, "Şube Kodu")
        ws.write(0, 1, "Clusters")
        ws.write(0, 2, "ENLEM")
        ws.write(0, 3, "BOYLAM")
        ws.write(0, 4, "Ortalama Ciro")
        ws.write(0, 5, "Servis Süreleri")
        ws.write(0, 6, "Sıklık")

        data = pd.read_excel(data_path)
        customer_ids = list(data['MÜŞTERİ KODU'])
        revenues = list(data[revenue_col])
        stimes = list(data['Servis Süresi'])
        freqs = list(data["Sıklık"])

        # Geçirilecek Süre hesabı time_calculation module kullanılıyor


        for i in range(0,len(customer_ids)):

            lat_temp = coordinates[i][0]
            long_temp = coordinates[i][1]
            ws.write(row, col, customer_ids[i])
            ws.write(row, col+1, best_clusters[i])
            ws.write(row, col+2, lat_temp)
            ws.write(row, col+3, long_temp)
            ws.write(row, col+4, revenues[i])
            ws.write(row, col+5, stimes[i])
            ws.write(row, col+6, freqs[i])
            row += 1

        wb.close()

    except TypeError:
        print("\n\n******HATA******\n"
              "Şu anki parametreler, müşterileri bölmeye uygun değil. Üst limiti yükseltmeyi, alt limiti azaltmayı veya TTE "
              "sayısını arttırmayı deneyebilirsiniz."
              "\n****************")
    return

#run_and_write()