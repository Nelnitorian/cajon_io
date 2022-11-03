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
    name='NOMBRE_EXCEL.xlsx'
    excel_doc=openpyxl.load_workbook(name,data_only=True)
    sheet=excel_doc['Hoja 1']

    arcos = parse_list(Read_Excel_to_List(sheet, 'a2', 'a11'))
    nodos = parse_list(Read_Excel_to_List(sheet, 'f2', 'f7'))
    prod_dem = parse_list(Read_Excel_to_List(sheet, 'g2', 'g7'))
    info_arcos = parse_list(Read_Excel_to_List(sheet, 'a1', 'd11'))

    A = {}
    for n in nodos:
        A[n] = {}
        for a in arcos:
            A[n][a] = 0.0
    for n in nodos:
        for a in arcos:
            if info_arcos[a]['i'] == n:
                A[n][a] = 1
            elif info_arcos[a]['j'] == n:
                A[n][a] = -1
            else:
                A[n][a] = 0

    a = Read_Excel_to_List(sheet, 'b2', 'b7')
    b = Read_Excel_to_List(sheet, 'd2', 'd6')

    c = Read_Excel_to_NesteDic(sheet, 'f1', 'k7')

    solver=pywraplp.Solver.CreateSolver('GLOP')

    x = {}
    for i in fabricas:
        x[i] = {}
        for j in almacenes:
            x[i][j] = solver.NumVar(0, solver.infinity(), f"x{i}{j}")

    for i in fabricas:
        solver.Add(sum(x[i][j] for j in almacenes)==a[i-1], f"RF{i}")
    for j in almacenes:
        solver.Add(sum(x[i][j] for i in fabricas)==b[j-1], f"RA{j}")

    solver.Minimize(solver.Sum(c[i][j]*x[i][j] for i in fabricas for j in almacenes))

    status=solver.Solve()
    if status==pywraplp.Solver.OPTIMAL:
        print('El problema tiene solucion')

        sol = {}
        crel = {}
        for i in fabricas:
            sol[i] = {j: 0.0 for j in almacenes}
            crel[i] = {j: 0.0 for j in almacenes}
        for i in fabricas:
            for j in almacenes:
                sol[i][j]=x[i][j].solution_value()
                crel[i][j]=x[i][j].reduced_cost()

        Write_NesteDic_to_Excel(excel_doc, name, sheet, sol,'f10', 'k16')
        Write_NesteDic_to_Excel(excel_doc, name, sheet, crel,'f18', 'k24')


if __name__=='__main__':
    main()
