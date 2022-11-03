# -*- coding: utf-8 -*-

import openpyxl as openpyxl
from ortools.linear_solver import pywraplp
from IOfunctionsExcel import *

def merge_dic_intersection(dic1, dic2):
    """
    Intersecta dos diccionarios y los une en una lista.

    :param dic1: diccionario cuyos valores seran los primeros de las listas
    :param dic2: diccionario cuyos valores seran los segundos de las listas
    :return:
    """
    key1 = list(dic1.keys())
    key2 = list(dic2.keys())
    c_keys = list(set(key1) & set(key2))

    aux = {}
    for key in c_keys:
        aux[key] = [dic1[key], dic2[key]]
    return aux
def main():
    """
    x:
        x0 = xa
        x1 = xb
        ...

    v:
        v0 =  infraestructuras
        v1 = educacion
        v2 = sanidad
        v3 = administración
        v4 = sector primario
        v5 = i+d
        v6 = desempleo

    :return:
    """
    name='caso1_excel.xlsx'
    excel_doc=openpyxl.load_workbook(name, data_only=True)
    sheet=excel_doc['Hoja1']

    s = Read_Excel_to_List(sheet, 'b2', 'b6')
    p = Read_Excel_to_List(sheet, 'c2', 'c6')
    ni = Read_Excel_to_List(sheet, 'd2', 'd6')

    g = list(Read_Excel_to_NesteDic(sheet, 'a10', 'i11').values())[0]

    solver = pywraplp.Solver.CreateSolver('GLOP')

    x = []
    for i in range(len(s)):
        x.append(solver.NumVar(0, solver.infinity(), f"x{i}"))

    v = []
    for i in range(len(g)-1):
        v.append(solver.NumVar(0, solver.infinity(), f"v{i}"))

    # R1
    solver.Add(solver.Sum(v[i] for i in range(len(v))) >= 0.75*g['TOTAL'], f"R1")
    # R2
    solver.Add((solver.Sum(s[i]*p[i]*(ni[i]+x[i]) for i in range(len(s))) -
               solver.Sum(v[i] for i in range(len(v)))) == 27000, f"R2")
    # R3
    solver.Add(v[4] >= 2500, f"R3")
    # R4
    solver.Add(v[5] >= 900, f"R4")
    # R5
    solver.Add(v[6] >= 6000, f"R5")
    # R6
    solver.Add(x[0]/s[0] == x[1]/s[1], f"R6")
    # R7
    solver.Add(x[1]/s[1] == x[2]/s[2], f"R7")
    # R8
    solver.Add(x[2]/s[2] == x[3]/s[3], f"R8")
    # R9
    solver.Add(x[3]/s[3] == x[4]/s[4], f"R9")
    # R10
    solver.Add(v[0]/g['INFRA'] == v[3]/g['ADMIN'], f"R10")
    # R11
    solver.Add(v[1] == g['EDUCACION'], f"R11")
    # R12
    solver.Add(v[2] == g['SANIDAD'], f"R12")

    solver.Minimize(solver.Sum(s[i]*p[i]*(ni[i]+x[i]) for i in range(len(s))))

    status=solver.Solve()

    if status==pywraplp.Solver.OPTIMAL:
        print('El problema tiene solucion.')

        t = []
        for sol in x:
            t.append(sol.solution_value())
        n_ni = [ni[i] + t[i] for i in range(len(ni))]
        ci = []
        for cii in v:
            ci.append(cii.solution_value())
        ing = sum([s[i]*p[i]*(ni[i]+t[i]) for i in range(len(s))])
        gas = sum(ci)
        print(f"El incremento de las tasas impositivas son: {t}")
        print(f"Dejando las tasas impositivas como: {n_ni}")
        print(f"Los nuevos costes de inversión son: {ci}")
        print(f"Los ingresos son: {round(ing, 2)}")
        print(f"Los gastos son: {gas}")
        print(f"Balance del año: {round(ing-gas, 2)}")

        # dic = {}
        # dic['NUEVO_NIVEL_IMPOSITIVO'] = [ni[i]+t[i] for i in range(len(ni))]
        # dic['NUEVO_ING_ANU'] = [s[i]*p[i]*dic['NUEVO_NIVEL_IMPOSITIVO'][i] for i in range(len(s))].append(ing)
        #
        # # Write_NesteDic_to_Excel(excel_doc, name, sheet, dic, 'f1', 'g7')
        #
        # dic2 = {}
        # # conseguimos las claves y quitamos el total
        # dic2_keys = list(g.keys())[:-1]
        # for i in range(len(dic2_keys)):
        #     dic2[dic2_keys[i]] = ci[i]
        # dic2['TOTAL'] = sum(ci)
        # merged = merge_dic_intersection(g, dic2)
        # # Write_NesteDic_to_Excel(excel_doc, name, sheet, merged, 'b10', 'i12')




    else:
        print('No hay solución óptima. Error.')

    return

if __name__=='__main__':
    main()
