from __future__ import division
import xlrd
import xlwt
from math import exp
from pprint import pprint as pp

import collections

from pyomo.environ import *

model = AbstractModel()

opt = SolverFactory('ipopt')

model.te = Set(initialize=['2015.0', '2020.0', '2025.0', '2030.0', '2035.0', '2040.0', '2045.0', '2050.0'],
               ordered=True)
model.tb = Set(initialize=['2015.0'], ordered=True)
model.t = Set(initialize=['2015.0', '2020.0', '2025.0', '2030.0', '2035.0', '2040.0', '2045.0', '2050.0'], ordered=True)
model.lc = Set(initialize=['peak', 'high', 'medium', 'low'], ordered=True)
model.sf = Set(initialize=['NuclearF', 'SolarF', 'SolidsF', 'Gas', 'HydroLakesF', 'HydroRorF', 'WindF', 'BiomassF'],
               ordered=True)
model.pl = Set(
    initialize=['NuclearP', 'SolarPVP', 'SolidsP', 'GTCCP', 'GasPeakP', 'OtherThermalP', 'CCS_coalP', 'HydroLakesP',
                'HydroRorP', 'WindP', 'BiomassP'], ordered=True)
model.im = Set(initialize=['imp1', 'imp2', 'imp3', 'imp4'], ordered=True)
model.fuel = Set(initialize=['NuclearF', 'SolarF', 'SolidsF', 'Gas', 'HydroLakesF', 'HydroRorF', 'WindF', 'BiomassF'],
                 ordered=True)
model.mapnf = Set(
    initialize=[('NuclearP', 'NuclearF'), ('SolarPVP', 'SolarF'), ('SolidsP', 'SolidsF'), ('GTCCP', 'Gas'),
                ('GasPeakP', 'Gas'), ('OtherThermalP', 'Gas'), ('CCS_coalP', 'SolidsF'), ('HydroLakesP', 'HydroLakesF'),
                ('HydroRorP', 'HydroRorF'), ('WindP', 'WindF'), ('BiomassP', 'BiomassF')], ordered=True)


def lag_lead(model, model_onoma_set, onoma_set, ord_):
    l = [i for i in model_onoma_set]
    indx = l.index(onoma_set)
    m = indx + ord_
    if m > -1 and m < len(l):
        return l[m]
    else:
        return None


file_location = "D:\irini (HD)\New folder\SupplyModel\SupplyModel\Data_elec.xlsx"

book = xlrd.open_workbook(file_location)

sheet=book.sheet_by_name('Sets')


number_of_sheets = book.nsheets

sheets = {}

for i in range(number_of_sheets):
    sheets[i] = book.sheet_by_index(i)


def find_set_or_param_values(indx_sheet_num,
                             crow):  # sinartisi pou epistrefei to sinolo twn timwn
    C_colum = []                     # sto excel pou kanoun initialize set / param
    split_C = []
    letter1 = ''
    letter2 = ''
    letter3 = ''
    letter4 = ''
    num2 = ''
    num1 = ''
    colums = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    times = []
    firstrow = ''
    lastrow = ''
    firstcolum = ''
    lastcolum = ''

    for k in range(sheets[indx_sheet_num].nrows):
        C_colum.append(sheets[indx_sheet_num].cell_value(rowx=k, colx=2))
    for l in range(sheets[indx_sheet_num].nrows):
        split_C.append(C_colum[l].split('!'))
    if split_C[crow][0] == book.sheet_by_name('Sets').name:
        bi1_split = split_C[crow][1].split(':')  # 2o split gia na exw mono ta onomata twn cell
        for m in bi1_split[0]:
            if m.isalpha():
                letter3 += m
            if m.isdigit():
                num1 += str(m)
                firstrow = str(int(num1) - 1)
        if len(letter3) > 1:
            for l in letter3:
                for j in colums:
                    if l == j:
                        firstcolum += str(colums.index(j) + 1)

        else:
            for j in colums:
                if letter3 == j:
                    firstcolum = str(colums.index(j))

        for m in bi1_split[1]:
            if m.isalpha():
                letter4 += m
            if m.isdigit():
                num2 += str(m)
                lastrow = str(int(num2) - 1)
        if len(letter4) > 1:
            for l in letter4:
                for j in colums:
                    if l == j:
                        lastcolum += str(colums.index(j) + 1)
            lastcolum = str(int(lastcolum) + 15)

        else:
            for j in colums:
                if letter4 == j:
                    lastcolum = str(colums.index(j))

        for c in range(int(firstcolum), int(lastcolum) + 1):
            for i in range(int(firstrow), int(lastrow) + 2):
                times.append(str(book.sheet_by_name('Sets').cell_value(i, c)))
    if split_C[crow][0] == book.sheet_by_name('Params').name:
        bi1_split = split_C[crow][1].split(':')  # 2o split gia na exw mono ta onomata twn cell
        for m in bi1_split[0]:
            if m.isalpha():
                letter3 += m
            if m.isdigit():
                num1 += str(m)
                firstrow = str(int(num1) - 1)
        if len(letter3) > 1:
            for l in letter3:
                for j in colums:
                    if l == j:
                        firstcolum += str(colums.index(j) + 1)
        else:
            for j in colums:
                if letter3 == j:
                    firstcolum = str(colums.index(j))

        for m in bi1_split[1]:
            if m.isalpha():
                letter4 += m
            if m.isdigit():
                num2 += str(m)
                lastrow = str(int(num2) - 1)
        if len(letter4) > 1:
            for l in letter4:
                for j in colums:
                    if l == j:
                        lastcolum += str(colums.index(j) + 1)
            lastcolum = str(int(lastcolum) + 15)

        else:
            for j in colums:
                if letter4 == j:
                    lastcolum = str(colums.index(j))

        for c in range(int(firstcolum), int(lastcolum) + 1):
            for i in range(int(firstrow), int(lastrow) + 2):
                times.append(str(book.sheet_by_name('Params').cell_value(i, c)))
    return times


def split_colums(indx_sheet_num, crow):  # sinartisi pou diaxwrizei
    exp = find_set_or_param_values(indx_sheet_num, crow)  # tis stiles apo to sinolo twn timwn aftwn
    kena = []
    for i in exp:
        if i == '':
            kena.append(i)
    num_of_sublists = len(kena)
    sublists = [[] for i in xrange(num_of_sublists)]

    for j in range(num_of_sublists):
        for i in range(len(exp)):
            while exp[i] != '':
                sublists[j].append(exp[i])
                exp.remove(exp[i])
            exp.remove(exp[i])
            break
    return sublists


def initialize_set(indx_sheet_num, crow):
    lista = split_colums(indx_sheet_num, crow)
    if len(lista) > 1:
        init = zip(*lista)
    else:
        init = lista[0]
    return init


def initialize_param(indxsheet_name, param_name):
    values = []
    indx_sheet_num = book.sheet_by_name(indxsheet_name).number
    for i in range(int(book.sheet_by_name(indxsheet_name).nrows)):
        if str(book.sheet_by_name(indxsheet_name).cell_value(i, 1)) == param_name:
            crow = i
    lista = split_colums(indx_sheet_num, crow)
    if len(lista) > 2:
        for i in lista[-1]:
            values.append(float(i))
        lista.remove(lista[-1])
        keys = zip(*lista)
        return dict(zip(keys, values))
    elif len(lista) == 2:
        for i in lista[1]:
            values.append(float(i))
        keys = lista[0]
        return dict(zip(keys, values))
    else:
        for i in lista[0]:
            return float(i)


def create_set_from_excel(indx_sheet_num, crow):
    return 'model.' + str(sheets[indx_sheet_num].cell_value(rowx=crow, colx=1)) + " = " + str(
        sheets[indx_sheet_num].cell_value(rowx=crow, colx=0)) + '(initialize=' + str(
        initialize_set(indx_sheet_num, crow)) + ',ordered=True' + ')'


def load_all_data(indxsheet_name):
    indx_sheet_num = book.sheet_by_name(indxsheet_name).number
    crow = 0
    a = 'from pyomo.environ import *\n\nmodel=AbstractModel()\n\n'
    if book.sheet_by_name(indxsheet_name).cell_value(crow, 0) == '':
        crow += 1
    for i in range(1, int(book.sheet_by_name(indxsheet_name).nrows)):
        a = a + create_set_from_excel(indx_sheet_num, crow) + '\n'
        crow += 1
    return a


sets = open('Create_sets_from_excel.py', 'w')
sets.write(load_all_data('indexSet'))
sets.close()

model.labels = Set(
    initialize=['duration', 'demand', 'growth', 'Capacity', 'max_fuel', 'Pgen', 'heatrate', 'eff', 'invcapcost',
                'maxinv', 'lifetime'], ordered=True)

model.tt = Set(initialize=model.te)

model.pf = Param(initialize=3.500, mutable=True)
model.r = Param(initialize=0.09, mutable=True)
model.disc = Param(initialize=0.04, mutable=True)

model.dd = Param(model.labels, model.lc, initialize=initialize_param('indexPar', 'dd'), mutable=True, default=0.0)
model.avail = Param(model.pl, model.lc, initialize=initialize_param('indexPar', 'avail'), mutable=True, default=0.0)
model.d = Param(model.lc, model.te, mutable=True, default=0.0)
model.h = Param(model.lc, mutable=True, default=0.0)
model.data = Param(model.labels, model.pl, initialize=initialize_param('indexPar', 'data'), mutable=True, default=0.0)
model.fuelcostdata = Param(model.sf, initialize=initialize_param('indexPar', 'fuelcostdata'), mutable=True, default=0.0)
model.maxfueldata = Param(model.sf, initialize=initialize_param('indexPar', 'maxfueldata'), mutable=True, default=0.0)
model.cap = Param(model.pl, model.te, mutable=True, default=0.0)
model.maxfuel = Param(model.sf, model.te, mutable=True, default=0.0)
model.heatrate = Param(model.pl, model.te, mutable=True, default=0.0)
model.availability = Param(model.pl, model.lc, model.te, mutable=True, default=0.0)
model.fuelcost = Param(model.sf, model.te, mutable=True, default=0.0)
model.invcost = Param(model.pl, model.te, mutable=True, default=0.0)
model.lifetime = Param(model.pl, model.te, mutable=True, default=0.0)
model.maxinv = Param(model.pl, model.te, mutable=True, default=0.0)
model.emisfactordata = Param(model.sf, initialize=initialize_param('indexPar', 'emisfactordata'), mutable=True,
                             default=0.0)
model.emisfactor = Param(model.sf, mutable=True, default=0.0)
model.etsprice = Param(model.te, initialize=initialize_param('indexPar', 'etsprice'), mutable=True, default=0.0)
model.etsprice_proj = Param(model.te, mutable=True, default=0.0)
model.surv = Param(model.pl, model.te, model.te, mutable=True, default=0.0)
model.capexog_data = Param(model.pl, model.te, initialize=initialize_param('indexPar', 'capexog_data'), mutable=True,
                           default=0.0)
model.capexog = Param(model.pl, model.te, model.te, mutable=True, default=0.0)
model.ccsco2price = Param(model.te, initialize=initialize_param('indexPar', 'ccsco2price'), mutable=True, default=0.0)
model.carboncapt = Param(model.pl, initialize=initialize_param('indexPar', 'carboncapt'), mutable=True, default=0.0)

model.capelc = Param(model.pl, model.te, mutable=True, default=0.0)
model.gen = Param(model.pl, model.te, mutable=True, default=0.0)
model.fuelcons = Param(model.pl, model.te, mutable=True, default=0.0)
model.report = Param(mutable=True, default=0.0)
model.dishpatch = Param(mutable=True, default=0.0)
model.total_cost = Param(mutable=True, default=0.0)

instance = model.create_instance()

for lc in instance.lc:
    for tb in instance.tb:
        instance.d[lc, tb] = value(instance.dd['demand', lc] / 1000)

for lc in instance.lc:
    instance.h[lc] = value(instance.dd['duration', lc])

for te in instance.te:
    if float(te) >= 2020:
        for lc in instance.lc:
            instance.d[lc, te] = value(dict(instance.d).get((lc, lag_lead(instance, instance.te, te, -1)), 0)) * (
                                                                                                                 1 + value(
                                                                                                                     instance.dd[
                                                                                                                         'growth', lc])) ** 5

for pl in instance.pl:
    for tb in instance.tb:
        for te in instance.te:
            instance.capexog[pl, tb, te] = value(instance.capexog_data[pl, te])

for sf in instance.sf:
    for te in instance.te:
        instance.maxfuel[sf, te] = value(instance.maxfueldata[sf])

for pl in instance.pl:
    for te in instance.te:
        instance.heatrate[pl, te] = value(instance.data['heatrate', pl])

for pl in instance.pl:
    for lc in instance.lc:
        for te in instance.te:
            instance.availability[pl, lc, te] = value(instance.avail[pl, lc])

for sf in instance.sf:
    for te in instance.te:
        instance.fuelcost[sf, te] = value(instance.fuelcostdata[sf])

for pl in instance.pl:
    for te in instance.te:
        instance.invcost[pl, te] = value(instance.data['invcapcost', pl])

for pl in instance.pl:
    for te in instance.te:
        instance.lifetime[pl, te] = value(instance.data['lifetime', pl])

for pl in instance.pl:
    for te in instance.te:
        instance.maxinv[pl, te] = value(instance.data['maxinv', pl]) / 1000

for sf in instance.sf:
    instance.emisfactor[sf] = value(instance.emisfactordata[sf]) * 0.086 / 1000

for pl in instance.pl:
    for tt in instance.tt:
        for te in instance.te:
            if (float(te) - float(tt)) <= value(instance.lifetime[pl, tt]) and (float(te) - float(tt)) >= 0:
                instance.surv[pl, tt, te] = 1


instance.g = Var(instance.pl, instance.tt, instance.lc, instance.te, within=NonNegativeReals)
instance.k = Var(instance.pl, instance.tt, within=NonNegativeReals)
instance.cut = Var(instance.lc, instance.te, within=NonNegativeReals)
instance.emis = Var(instance.te, within=NonNegativeReals)
instance.emis_capt = Var(instance.te, within=NonNegativeReals)
instance.cost = Var()


def eqdemand_rule(instance, lc, te):
    if te in instance.t:
        return value(instance.h[lc]) * sum(
            sum(instance.g[pl, tt, lc, te] for pl in instance.pl) for tt in instance.tt if
            float(tt) <= float(te)) >= value(instance.h[lc]) * (value(instance.d[lc, te]) - instance.cut[lc, te])
    else:
        return Constraint.Skip

instance.eqdemand = Constraint(instance.lc, instance.te, rule=eqdemand_rule)


def eqcapacity_rule(instance, pl, tt, lc, te):
    if te in instance.t and float(tt) <= float(te):
        return value(instance.h[lc]) * instance.g[pl, tt, lc, te] <= value(instance.h[lc]) * (
        value(instance.capexog[pl, tt, te]) + instance.k[pl, tt] * value(instance.surv[pl, tt, te])) * value(
            instance.availability[pl, lc, tt])
    else:
        return Constraint.Skip

instance.eqcapacity = Constraint(instance.pl, instance.tt, instance.lc, instance.te, rule=eqcapacity_rule)


def eqfuel_rule(instance, sf, te):
    if te in instance.t and value(instance.maxfuel[sf, te]) < float('inf'):
        return sum(sum(sum(
            value(instance.h[lc]) * instance.g[pl, tt, lc, te] * value(instance.heatrate[pl, tt]) for lc in instance.lc)
                       for tt in instance.tt if float(tt) <= float(te)) for pl in instance.pl if
                   (pl, sf) in instance.mapnf) <= value(instance.maxfuel[sf, te])
    else:
        return Constraint.Skip


instance.eqfuel = Constraint(instance.sf, instance.te, rule=eqfuel_rule)


def eqmaxinv_rule(instance, pl, te):
    if te in instance.t and value(instance.maxinv[pl, te]) < float('inf'):
        return sum(instance.k[pl, tt] * value(instance.surv[pl, tt, te]) for tt in instance.tt if
                   float(tt) <= float(te)) <= value(instance.maxinv[pl, te])
    else:
        return Constraint.Skip


instance.eqmaxinv = Constraint(instance.pl, instance.te, rule=eqmaxinv_rule)


def eqemis_rule(instance, te):
    return instance.emis[te] == sum(value(instance.emisfactor[sf]) * sum((1 - value(instance.carboncapt[pl])) * sum(
        sum(value(instance.h[lc]) * instance.g[pl, tt, lc, te] * value(instance.heatrate[pl, tt]) for lc in instance.lc)
        for tt in instance.tt if float(tt) <= float(te)) for pl in instance.pl if (pl, sf) in instance.mapnf) for sf in
                                    instance.sf)


instance.eqemis = Constraint(instance.te, rule=eqemis_rule)


def eqemiscapt_rule(instance, te):
    return instance.emis_capt[te] == sum(value(instance.emisfactor[sf]) * sum(value(instance.carboncapt[pl]) * sum(
        sum(value(instance.h[lc]) * instance.g[pl, tt, lc, te] * value(instance.heatrate[pl, tt]) for lc in instance.lc)
        for tt in instance.tt if float(tt) <= float(te)) for pl in instance.pl if (pl, sf) in instance.mapnf) for sf in
                                         instance.sf)


instance.eqemiscapt = Constraint(instance.te, rule=eqemiscapt_rule)


def eqobjective_rule(instance):
    return sum((
                   sum(sum(
                       value(instance.h[lc]) * instance.g[pl, tt, lc, te] * value(instance.heatrate[pl, tt]) * value(
                           instance.fuelcost[sf, te]) for lc in instance.lc) for pl in instance.pl for sf in instance.sf
                       for tt in instance.tt if float(tt) <= float(te) and (pl, sf) in instance.mapnf)
                   + sum(
                       value(instance.h[lc]) * instance.cut[lc, te] * value(instance.pf) for lc in instance.lc) + value(
                       instance.etsprice[te]) * instance.emis[te]
                   + value(instance.ccsco2price[te]) * instance.emis_capt[te]
                   + sum(sum(
                       instance.k[pl, tt] * value(instance.surv[pl, tt, te]) * value(instance.invcost[pl, tt]) * value(
                           instance.r) * exp(value(instance.r) * value(instance.lifetime[pl, tt])) / (
                       exp(value(instance.r) * value(instance.lifetime[pl, tt])) - 1) for pl in instance.pl) for tt in
                         instance.tt if tt in instance.t and float(tt) <= float(te))
                   + sum(sum(
                       value(instance.capexog[pl, tt, te]) * value(instance.invcost[pl, tt]) * value(instance.r) * exp(
                           value(instance.r) * value(instance.lifetime[pl, tt])) / (
                       exp(value(instance.r) * value(instance.lifetime[pl, tt])) - 1) for pl in instance.pl) for tt in
                         instance.tt if tt in instance.t and float(tt) <= float(te))
               ) for te in instance.t)


instance.eqobjective = Objective(rule=eqobjective_rule)

for pl in instance.pl:
    for tb in instance.tb:
        instance.k[pl, tb].setlb(0)
        instance.k[pl, tb].setub(0)

for lc in instance.lc:
    for te in instance.te:
        instance.cut[lc,te].setub(instance.d[lc,te])
        instance.cut[lc,te].setub(0)
        instance.cut[lc,te].setlb(0)


def include_NoNuc_dec(model,instance):
    for te in instance.te:
        instance.k['NuclearP', te].setlb(0)
        instance.k['NuclearP', te].setub(0)
        if float(te) >= 2035:
            instance.etsprice[te] = 2.5 * instance.etsprice[te]


include_NoNuc_dec(model,instance)

results = opt.solve(instance)

instance.solutions.load_from(results)

display(instance)