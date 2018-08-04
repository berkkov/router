import pulp
import pandas as pd
import os
import sys

if getattr(sys, 'frozen', False):
    # frozen
    d = os.path.dirname(sys.executable)
else:
    # unfrozen
    d = os.path.dirname(os.path.realpath(__file__))


pa = r'{}/data'.format(d)
datapath = os.path.join(pa, "clusters.xlsx")
NUMBER_OF_DAYS = 6
SWITCH = 550
costpath = os.path.join(pa, "durations.xlsx")

'''
K_ARRAY = [[1,2,3,4,5,6],[[1,2],[1,3],[1,4],[1,5],[1,6],[2,3],[2,4],[2,5],[2,6],[3,4],[3,5],[4,5],[3,6],[4,6],[5,6]],[[1,2,3],[1,2,4],[1,2,5],[1,2,6],[1,3,4],[1,3,5],[1,4,5],[1,3,6],[1,4,6],[1,5,6],[2,3,4],[2,3,5],
           [2,3,6],[2,4,5],[2,4,6],[2,5,6],[3,4,5],[3,4,6],[3,5,6]],[[1,2,3,5],[1,2,3,6],[1,2,4,6],[1,2,5,6],[1,3,4,5],[1,3,5,6],[1,4,5,6],
             [2,3,4,6],[2,3,5,6],[2,4,5,6],[1,3,4,6]],[[1,2,3,4,5],[1,2,3,4,6],[1,2,3,5,6],[1,2,4,5,6],[1,3,4,5,6],[2,3,4,5,6]],
           [[1,2,3,4,5,6]]]
'''

K_ARRAY = [[1,2,3,4,5,6],[[1,3],[1,4],[1,5],[1,6],[2,4],[2,5],[2,6],[3,5],[3,6],[4,6]],[[1,2,4],[1,2,5],[1,2,6],[1,3,4],[1,3,5],[1,4,5],[1,3,6],[1,4,6],[1,5,6],[2,3,5],
          [2,3,6],[2,4,5],[2,4,6],[2,5,6],[3,4,6],[3,5,6]],[[1,2,3,5],[1,2,3,6],[1,2,4,6],[1,2,5,6],[1,3,4,5],[1,3,5,6],[1,4,5,6],
            [2,3,4,6],[2,3,5,6],[2,4,5,6],[1,3,4,6]],[[1,2,3,4,5],[1,2,3,4,6],[1,2,3,5,6],[1,2,4,5,6],[1,3,4,5,6],[2,3,4,5,6]],
           [[1,2,3,4,5,6]]]

NUM_OF_POS = 0

for j in K_ARRAY:
    NUM_OF_POS += len(j)


def prepare_duration_matrix(cluster_no):
    """
    :param cluster_no: Cluster numarası
    :return: n+2*n+2 lik duration matrix'i (+2 dummy'den geliyor, n clusterdaki eleman sayısı)
    """

    durations = pd.read_excel(costpath, "Cluster" +str(cluster_no))

    durations.drop(['Yolda geçen süre (dk)'], axis=1)
    l = durations.as_matrix().tolist()
    for i in l:
        i.pop(0)

    return l


class problem(object):
    def __init__(self, data, cluster_no):
        data = pd.read_excel(datapath)
        self.cluster_no = cluster_no
        self.cluster = data[data['Clusters'] == cluster_no]
        self.durations = prepare_duration_matrix(cluster_no)

        # Ziyaret sıklıkları (ilk ve son dummy'lerin. Her gün gidiliyor)
        freq = self.cluster['Sıklık'].tolist()
        freq.insert(0, 6)
        freq.insert(len(freq), 6)
        self.frequencies = freq

        # Servis Süreleri (ilk ve son dummy'lerin de var. Dummy servis süreleri 0)
        serv = self.cluster['Servis Süreleri'].tolist()
        serv.insert(0, 0)
        serv.insert(len(freq), 0)
        self.service_times = serv

        self.periods = [1, 2, 3, 4, 5, 6]
        self.n = len(self.cluster)
        self.n_with_dummies = self.n + 2

        self.nodes = [-1] * self.n_with_dummies
        for j in range(self.n_with_dummies):  # TODO: Son dummy sıkıntı yapabilir
            self.nodes[j] = j

        self.create_model()

    def create_model(self):
        '''
        MIP modeli
        :return:
        '''
        Routes = [(i,j,p) for i in range(self.n_with_dummies) for j in range(self.n_with_dummies) for p in
                  self.periods]
        Y = [(i,k) for i in range(1,self.n + 1) for k in range(1,NUM_OF_POS+1)]
        Time = [(i,p) for i in range(self.n_with_dummies) for p in self.periods]
        U = [(i,p) for i in range(self.n_with_dummies) for p in self.periods]

        self.x = pulp.LpVariable.dicts(name='X', indexs=Routes, lowBound=0, upBound=1, cat=pulp.LpInteger)

        self.S = pulp.LpVariable.dicts(name='S', indexs=Time, lowBound=0, upBound=SWITCH)

        self.u = pulp.LpVariable.dicts(name='u', indexs=U)  # , upBound=1, lowBound=0, cat= pulp.LpInteger

        self.y = pulp.LpVariable.dicts(name='y', indexs=Y, lowBound=0, upBound=1, cat=pulp.LpInteger)

        self.penalty = pulp.LpVariable.dicts(name='penalty', indexs=self.periods, lowBound=0) # TODO:işe yarmazsa çıkar

        self.g = pulp.LpVariable.dicts(name='g', indexs=Time)

        self.z = pulp.LpVariable.dicts(name='z', indexs=self.periods) # O gün gidilen toplam market sayısı

        self.sigma = pulp.LpVariable.dicts(name='sigma', indexs=self.periods, lowBound=0, upBound=1, cat=pulp.LpInteger) #

        self.extra = pulp.LpVariable.dicts(name='extra', indexs=self.periods)  # Mesai saatinin üzerine çıkış

        self.deficit = pulp.LpVariable.dicts(name='def', indexs=self.periods)

        bigM = 1000

        self.model = pulp.LpProblem("RoutingProblem" + str(self.cluster_no), pulp.LpMinimize)

        # Objective Function
        print(self.durations)
        self.model += pulp.lpSum(
            [self.x[(i, j, p)] * self.durations[i][j] + bigM * self.extra[(p)] for i in range(self.n_with_dummies)  for j in
             range(self.n_with_dummies) for p in self.periods])

        # Constraints
        for i in range(1,self.n+1):
            self.model += pulp.lpSum([self.y[(i,k)] for k in range(1,NUM_OF_POS+1)]) == 1

        for i in range(1,self.n + 1):
            size_until = 0

            for a in range(0,self.frequencies[i]-1):
                size_until += len(K_ARRAY[a])
            for p in self.periods:
                k_pos = []
                for li in K_ARRAY[self.frequencies[i] - 1]:
                    if p not in li:
                        continue
                    k_pos += [size_until + K_ARRAY[self.frequencies[i]-1].index(li) + 1]

                
                self.model += (pulp.lpSum([[self.x[(i,j,p)]] for j in (s for s in range(1,self.n+1) if s!=i)])) == (pulp.lpSum([self.y[(i,k)] for k in k_pos]))

        for p in self.periods:
            #for i in range(self.n + 1):  # YANLIŞ OLABİLİR

                    #self.model += pulp.lpSum(
                    #[self.x[(i, j, p)] for j in (s for s in range(1, self.n_with_dummies) if s != i)]) <= 1

            for i in range(self.n_with_dummies):  # EKLENTI 1
                for j in range(self.n_with_dummies):
                    A=5
                    #self.model += self.x[(i, j, p)] + self.x[(j, i, p)] <= 1


            for j in range(1, self.n + 1):
                # Constraint 2  == Xpress Constraint 1
                self.model += pulp.lpSum(
                    [self.x[(i, j, p)] for i in (s for s in range(self.n + 1) if s != j)]) == pulp.lpSum(
                    [self.x[(j, i, p)] for i in (s for s in range(1, self.n_with_dummies) if s != j)])


            # C3 Sıfırdan çıkılacak = XPress constraint 2
            self.model += pulp.lpSum(
                [self.x[(0, j, p)] for j in (s for s in range(0, self.n_with_dummies) if s != 0)]) == 1

            # C4 Sıfıra Giriş yok = Xpress Constraint 3
            self.model += pulp.lpSum([self.x[(j, 0, p)] for j in range(0, self.n_with_dummies)]) == 0


            #### C5  Bitiş dummy node'una dönülecek = Xpress Constraint 4
            self.model += pulp.lpSum([self.x[(j, self.n + 1, p)] for j in range(1, self.n + 1)]) == 1

            # C6 Bİtiş dummy node'undan çıkış yok = Xpress Constraint 5
            self.model += pulp.lpSum([self.x[(self.n + 1, j, p)] for j in range(0, self.n_with_dummies)]) == 0

            # C7 Bir node'dan kendisine gidilemez = Xpress Constraint 7
            for i in range(self.n_with_dummies):
                self.model += self.x[(i, i, p)] == 0

            # C7  Mesai başlangıcı = Xpress Constraint 9
            self.model += self.S[(0, p)] == 0

            # C8 Mesai toplam süresi = Xpress Constraint 11
            self.model += self.S[(self.n + 1, p)] <= SWITCH

            # C9 Subtour ve süre arttırımı = Xpress Constraint 10
            for i in range(self.n_with_dummies):
                for j in range(self.n_with_dummies):
                    self.model += (self.S[(i, p)] + self.durations[i][j] + self.service_times[i] - SWITCH * (
                                1 - self.x[(i, j, p)])) <= self.S[(j, p)]

            # C10 Mesai bitiş süresi matematiksel mantığı = Xpress Constraint 11
            self.model += self.S[(self.n + 1, p)] == pulp.lpSum(
                [((self.durations[i][j] + self.service_times[i]) * self.x[(i, j, p)]) for i in
                 range(self.n_with_dummies) for j in range(self.n_with_dummies)])


        # C11 Haftalık ziyaret toplamı frekansa eşit = Xpress Constraint 8
        for i in range(1, self.n + 1):
            self.model += pulp.lpSum([self.x[(i, j, p)] for p in self.periods for j in
                                      (s for s in range(1, self.n_with_dummies) if s != i)]) == self.frequencies[i]


        for p in self.periods:
            for i in range(1,self.n + 1):
                self.model += self.g[(i,p)] == pulp.lpSum(self.x[(i,j,p)] for j in range(1, self.n_with_dummies))

        for p in self.periods:
            
            self.model += self.S[(self.n+1,p)] >= 450 - bigM*(1-self.sigma[(p)])
            self.model += self.S[(self.n+1,p)] <= 450 + bigM*(self.sigma[(p)])
            self.model += self.S[(self.n+1,p)] - 450 - bigM*(1 - self.sigma[(p)]) <= self.extra[(p)]
            self.model += self.extra[(p)] <= self.S[(self.n+1,p)] - 450 + bigM * (1-self.sigma[(p)])
            self.model += -bigM* self.sigma[(p)] <= self.extra[(p)]
            self.model += self.extra[(p)] <= bigM * self.sigma[(p)]
            
            self.model += self.z[(p)] == pulp.lpSum([self.g[(i, p)] for i in range(1, self.n+1)])
            self.model += self.deficit[(p)] == 450 - self.S[(self.n+1, p)]


    def solve(self):
        '''
        Model çözüm fonksiyonu, GLPK open-source solver'ı kullanır.
        :return:
        '''
        self.status = self.model.solve(pulp.GLPK_CMD(options=['--tmlim', '480'],keepFiles=False))

        solution_dict = {}

        if self.status == 1:
            for i in range(self.n_with_dummies):
                for j in range(self.n_with_dummies):
                    for p in self.periods:
                        if self.x[(i, j, p)].value():
                            solution_dict['x' + str((i, j, p))] = 1

            for i in range(self.n_with_dummies):
                for p in self.periods:

                    solution_dict['S' + str((i, p))] = self.S[(i,p)].value()

            for i in range(1,self.n + 1):
                for p in self.periods:
                    if self.g[(i, p)].value():
                        solution_dict['g' + str((i,p))] = self.g[(i,p)].value()

            for p in self.periods:
                if self.extra[(p)].value():
                    solution_dict['extra' + str((p))] = self.extra[(p)].value()

            for p in self.periods:
                solution_dict['z' + str((p))] = self.z[(p)].value()

            for p in self.periods:
                if self.deficit[(p)].value():
                    solution_dict['def' + str((p))] = self.deficit[(p)].value()
        return solution_dict


def prepare_output_file(solution,ws):
    '''
    Bir problem için çıktı dosyası(excel) oluşturur.
    :param solution: problemin çözümünden gelen rota değerleri.
    :param ws: yazılacak olan excel sheet'i
    :return:
    '''

    try:
        row = 0
        col = 0
        col_s = 2
        row_s = 0
        col_g = 4
        row_g = 0
        row_e = 0
        col_e = 6
        row_z = 0
        col_z = 8
        row_rem = 0
        col_rem = 10
        for i in solution.keys():
            if i[0] == 'x' or i[0] == 'X' :
                ws.write(row,col, i)
                ws.write(row,col+1, solution[i])
                row += 1
            if i[0] == 'S' or i[0] == 's' :
                ws.write(row_s, col_s, i)
                ws.write(row_s, col_s + 1, solution[i])
                row_s += 1

            if i[0] == 'g' or i[0] == 'G':

                ws.write(row_g, col_g, i)
                ws.write(row_g, col_g + 1, solution[i])
                row_g +=1

            if i[0] == 'e' or i[0] == 'E':
                ws.write(row_e, col_e, i)
                ws.write(row_e, col_e + 1, solution[i])
                row_e += 1

            if i[0] == 'z' or i[0] == 'Z':
                ws.write(row_z, col_z, i)
                ws.write(row_z, col_z+1, solution[i])
                row_z += 1

            if i[0] == 'd':
                ws.write(row_rem, col_rem, i)
                ws.write(row_rem, col_rem + 1, solution[i])
                row_rem += 1

    except TypeError and IndexError:
        print("Problem çözülemedi (infeasible).")

    return
