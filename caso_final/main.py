# -*- coding: utf-8 -*-

import openpyxl as openpyxl
from ortools.linear_solver import pywraplp
from utils import *
import pandas as pd

# Constants
EXCEL_FILE_NAME = 'casofinal_excel.xlsx'
DATA_SHEET_NAME = 'Datos'
ANSWER_SHEET_BASE_NAME = 'Resultados'
PATIENTS_ANSWER_SHEET = ANSWER_SHEET_BASE_NAME + " Pacientes"
ORS_ANSWER_SHEET = ANSWER_SHEET_BASE_NAME + " Quirofanos"
SURGEONS_ANSWER_SHEET = ANSWER_SHEET_BASE_NAME + " Cirujanos"


PATIENTS_DATA_SEET_NAME = 'Patients'
SURGEONS_DATA_SEET_NAME = 'Surgeons'
ORS_DATA_SEET_NAME = 'ORs'
TIME_SLOTS_DATA_SEET_NAME = 'Time_Slots'
DAYS_DATA_SEET_NAME = 'Days'

OPERATION_DURATION = 2 # hours

def main():
    """


    :return:
    """


    solver = pywraplp.Solver.CreateSolver('CBC')

    # Read from excel
    excel_doc = openpyxl.load_workbook(EXCEL_FILE_NAME, data_only=True)

    # Se lee el diccionario donde está la información de las tablas
    table_information = {PATIENTS_DATA_SEET_NAME: ['A1', 'D101'], SURGEONS_DATA_SEET_NAME: ['B4', 'N10'], ORS_DATA_SEET_NAME: ['A1', 'M5'],
                         TIME_SLOTS_DATA_SEET_NAME: ['A1', 'A4'], DAYS_DATA_SEET_NAME: ['A1', 'A6']}
    table_contents = {}
    for key, value in table_information.items():
        table_contents[key] = read_from_excel_to_dataframe(EXCEL_FILE_NAME, key, value)

    # Definición de indices

    quirofanos_chr = list(table_contents[ORS_DATA_SEET_NAME]["ORs"])
    cirujanos_chr = list(table_contents[SURGEONS_DATA_SEET_NAME]["Sid"])
    pacientes_chr = list(table_contents[PATIENTS_DATA_SEET_NAME]["patient_id"])
    dias_chr = list(table_contents[DAYS_DATA_SEET_NAME]["Days"])
    turnos_chr = list(table_contents[TIME_SLOTS_DATA_SEET_NAME]["Time_Slots"])

    quirofanos = list(range(len(quirofanos_chr)))
    cirujanos = list(range(len(cirujanos_chr)))
    pacientes = list(range(len(pacientes_chr)))
    dias = list(range(len(dias_chr)))
    turnos = list(range(len(turnos_chr)))

    # Definición de constantes

    skill_quirofanos = extract_indexes_with_value(table_contents[ORS_DATA_SEET_NAME])
    skill_cirujanos = extract_indexes_with_value(table_contents[SURGEONS_DATA_SEET_NAME])
    skill_pacientes = list(table_contents[PATIENTS_DATA_SEET_NAME]["sType"])

    g = list(table_contents[PATIENTS_DATA_SEET_NAME]["imp"])
    z = calculate_list_date_difference(list(table_contents[PATIENTS_DATA_SEET_NAME]["admision_date"]))
    Kid = 'OR{i}'
    Lid = 'S{i}'
    Pid = '#p#{i}'

    # Definición de variables de compatibilidad (paso necesario para la optimización)
    indexes = create_list_empty_nested_dics(len(quirofanos))
    for i in quirofanos:
        for j in cirujanos:
            # Patologías operables en el quirófano i por el cirujano j
            skills_operables = list(set(skill_quirofanos[i])-(set(skill_quirofanos[i])-set(skill_cirujanos[j])))
            if skills_operables:
                # El cirujano puede operar
                # Se calcula a qué pacientes puede operar:
                skills_a_operar = list(set(skills_operables)-(set(skills_operables)-set(skill_pacientes)))
                pacientes_a_operar = [i for i, x in enumerate(skill_pacientes) if x in skills_a_operar]
                if pacientes_a_operar:
                    indexes[i][j] = pacientes_a_operar

    # Definición de variables
    x = []
    for i in dias:
        x.append([])
        for j in turnos:
            x[i].append([])
            for k in quirofanos:
                x[i][j].append([])

    for i in dias:
        for j in turnos:
            for k in quirofanos:
                cirujanos_disp = indexes[k]
                if cirujanos_disp:
                    x[i][j][k] = {}
                    for l, pacientes_disp in cirujanos_disp.items():
                        x[i][j][k][l] = {}
                        for m in pacientes_disp:
                            x[i][j][k][l][m] = solver.IntVar(0, 1, f"x_{i}_{j}_{k}_{l}_{m}")

    # R1
    for i in dias:
        for j in turnos:
            for k in quirofanos:
                solver.Add(solver.Sum([var for l, dic in x[i][j][k].items() for m, var in dic.items()]) <= 1)

    # R2 y R3
    # TODO compactar
    for l in cirujanos:
        accumulated_hours = []
        for i in dias:
            for j in turnos:
                for k in quirofanos:
                    if l in x[i][j][k].keys():
                        dic = x[i][j][k][l]
                        for m, var in dic.items():
                            accumulated_hours.append(OPERATION_DURATION*var)
        solver.Add(solver.Sum(accumulated_hours) <= 18)
        solver.Add(solver.Sum(accumulated_hours) >= 10)

    # R4
    # TODO compactar
    for m in pacientes:
        accumulated_hours = []
        for i in dias:
            for j in turnos:
                for k in quirofanos:
                    for l, dic in x[i][j][k].items():
                        if m in dic.keys():
                            accumulated_hours.append(dic[m])
        solver.Add(solver.Sum(accumulated_hours) <= 1)

    # R5
    # TODO compactar
    days_reduced = []
    for m in pacientes:
        patient_operation = []
        for i in dias:
            for j in turnos:
                for k in quirofanos:
                    for l, dic in x[i][j][k].items():
                        if m in dic.keys():
                            patient_operation.append(dic[m])
        days_reduced.append(z[m] * solver.Sum(patient_operation))

    solver.Add(solver.Sum(days_reduced) >= 0.45*sum(z[m] for m in pacientes))

    # R6
    # al resolver el problema equivalente con menos coste computacional,
    # no hace falta esta restricción

    # FO
    # TODO compactar
    punctuation = []
    for m in pacientes:
        patient_operation = []
        for i in dias:
            for j in turnos:
                for k in quirofanos:
                    for l, dic in x[i][j][k].items():
                        if m in dic.keys():
                            patient_operation.append(dic[m])
        punctuation.append(z[m] * g[m] * (1-solver.Sum(patient_operation)))
    FO = solver.Sum(punctuation)
    solver.Minimize(FO)
    status = solver.Solve()

    if status == pywraplp.Solver.OPTIMAL:
        print('El problema tiene solucion.')

        """
        Representar:
            · Horario de cada cirujano | done 
            · Horario de cada quirófano | done
            · Lista de pacientes operados (con día, turno, cirujano, dolencia) | done
            · Lista de pacientes en cola (con toda la info que viene en los datos) | falta
        """
        x_sol = []
        for i in dias:
            x_sol.append([])
            for j in turnos:
                x_sol[i].append([])
                for k in quirofanos:
                    x_sol[i][j].append([])
        for i in dias:
            for j in turnos:
                for k in quirofanos:
                    cirujanos_disp = indexes[k]
                    if cirujanos_disp:
                        x_sol[i][j][k] = {}
                        for l, pacientes_disp in cirujanos_disp.items():
                            x_sol[i][j][k][l] = {}
                            for m in pacientes_disp:
                                x_sol[i][j][k][l][m] = x[i][j][k][l][m].solution_value()

        # El calendario tendrá índices lji (ciru turno dia)
        # Con valores (quirofano, paciente, dolencia)
        surgeon_calendar = {}
        for l in cirujanos_chr:
            surgeon_calendar[l] = create_empty_nested_dics(turnos_chr)

        for l in cirujanos:
            for i in dias:
                for j in turnos:
                    for k in quirofanos:
                        if l in x_sol[i][j][k].keys():
                            dic = x_sol[i][j][k][l]
                            for m, var in dic.items():
                                if var == 1:
                                    surgeon_calendar[cirujanos_chr[l]][turnos_chr[j]][dias_chr[i]] = f"({quirofanos_chr[k]}, {pacientes_chr[m]}, {skill_pacientes[m]})"
                                else:
                                    surgeon_calendar[cirujanos_chr[l]][turnos_chr[j]][dias_chr[i]] = ''

        # El calendario tendrá índices kji (quiro turno dia)
        # Con valores (cirujano, turno, dia)
        ors_calendar = {}
        for l in quirofanos_chr:
            ors_calendar[l] = create_empty_nested_dics(turnos_chr)

        for i in dias:
            for j in turnos:
                for k in quirofanos:
                    for l, dic in x_sol[i][j][k].items():
                        for m, var in dic.items():
                            if var == 1:
                                ors_calendar[quirofanos_chr[k]][turnos_chr[j]][dias_chr[i]] = f"({cirujanos_chr[l]}, {pacientes_chr[m]}, {skill_pacientes[m]})"
                            else:
                                ors_calendar[quirofanos_chr[k]][turnos_chr[j]][dias_chr[i]] = ''

        # El calendario tendrá índices kji (quiro turno dia)
        # Con valores (con día, turno, cirujano, dolencia)

        patients_calendar = {}
        for i in dias:
            for j in turnos:
                for k in quirofanos:
                    for l, dic in x_sol[i][j][k].items():
                        for m, var in dic.items():
                            if var == 1:
                                patients_calendar[pacientes_chr[m]] = {}
                                patients_calendar[pacientes_chr[m]]['Dia'] = dias_chr[i]
                                patients_calendar[pacientes_chr[m]]['Turno'] = turnos_chr[j]
                                patients_calendar[pacientes_chr[m]]['Cirujano que opera'] = cirujanos_chr[l]
                                patients_calendar[pacientes_chr[m]]['Quirofano'] = quirofanos_chr[k]
                                patients_calendar[pacientes_chr[m]]['Dolencia a operar'] = skill_pacientes[m]

        pacientes_sin_operar_keys = list(set(pacientes_chr) - set(patients_calendar.keys()))
        pacientes_sin_operar_keys.sort()
        pacientes_sin_operar = create_empty_nested_dics(pacientes_sin_operar_keys)
        df = table_contents[PATIENTS_DATA_SEET_NAME]
        claves_secundarias = list(df.keys())
        claves_secundarias.remove('patient_id')
        for m_chr in pacientes_sin_operar.keys():
            df_aux = df[df["patient_id"] == m_chr]
            for clave in claves_secundarias:
                pacientes_sin_operar[m_chr][clave] = list(df_aux[clave])[0]

        # Escribimos los datos de los pacientes

        patient_answer_excel_sheet = excel_doc[PATIENTS_ANSWER_SHEET]
        patient_dics_to_save_array = [pacientes_sin_operar, patients_calendar]

        patient_ranges = calculate_write_ranges_from_dic_array(patient_dics_to_save_array, start='C3')
        for j in range(len(patient_dics_to_save_array)):
            write_nested_dicts_to_excel(excel_doc, EXCEL_FILE_NAME, patient_answer_excel_sheet, patient_dics_to_save_array[j], patient_ranges[j], f'S{j + 1}')

        # Escribimos los datos de los cirujanos
        # surgeon_dics_to_save_array = [surgeon_calendar]
        surgeon_answer_excel_sheet = excel_doc[SURGEONS_ANSWER_SHEET]
        surgeon_dics_array = list(surgeon_calendar.values())
        surgeon_dics_keys = list(surgeon_calendar.keys())
        surgeon_ranges = calculate_write_ranges_from_dic_array(surgeon_dics_array, start='C3')
        for i in range(len(surgeon_dics_array)):
            write_nested_dicts_to_excel(excel_doc, EXCEL_FILE_NAME, surgeon_answer_excel_sheet, surgeon_dics_array[i], surgeon_ranges[i], f'S_{surgeon_dics_keys[i]}')

        # Escribimos los datos de las salas
        # ors_dics_to_save_array = [ors_calendar]
        ors_answer_excel_sheet = excel_doc[ORS_ANSWER_SHEET]
        ors_dics_array = list(ors_calendar.values())
        ors_dics_keys = list(ors_calendar.keys())
        ors_ranges = calculate_write_ranges_from_dic_array(ors_dics_array, start='C3')
        for i in range(len(ors_dics_array)):
            write_nested_dicts_to_excel(excel_doc, EXCEL_FILE_NAME, ors_answer_excel_sheet, ors_dics_array[i], ors_ranges[i], f'S_{ors_dics_keys[i]}')





        print(f"Valor de la función objetivo total: {FO.solution_value()}")
    else:
        print('No hay solución óptima. Error.')

    return

if __name__=='__main__':
    main()
