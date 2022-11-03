# -*- coding: utf-8 -*-

import openpyxl as openpyxl
from ortools.linear_solver import pywraplp
from IOfunctionsExcel import *

def parse_list(lst):
    aux = []
    for item in lst:
        aux.append(int(item))
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

    s = parse_list(Read_Excel_to_List(sheet, 'b2', 'b6'))
    p = parse_list(Read_Excel_to_List(sheet, 'c2', 'c6'))
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

        # sol = {}
        # crel = {}
        # for i in fabricas:
        #     sol[i] = {j: 0.0 for j in almacenes}
        #     crel[i] = {j: 0.0 for j in almacenes}
        # for i in fabricas:
        #     for j in almacenes:
        #         sol[i][j]=x[i][j].solution_value()
        #         crel[i][j]=x[i][j].reduced_cost()
        #
        # Write_NesteDic_to_Excel(excel_doc, name, sheet, sol,'f10', 'k16')
        # Write_NesteDic_to_Excel(excel_doc, name, sheet, crel,'f18', 'k24')

        t = []
        for sol in x:
            t.append(sol.solution_value())
        ci = []
        for cii in v:
            ci.append(cii.solution_value())

        print(f"El incremento de las tasas impositivas son: {t}\nLos nuevos costes de inversión son: {ci}")

    else:
        print('No hay solución óptima. Error.')


if __name__=='__main__':
    main()
