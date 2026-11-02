# -*- coding: utf-8 -*-
"""
PRACTIPLAN GIOI-DOE — Entrega final
Asignación de alumnos a grupos de prácticas y asignación de sesiones DOE a laboratorios.
Mejoras incorporadas:
1) Fase 2 greedy con criterio de equilibrio de grupos.
2) Fase 3 con Búsqueda Tabú para minimizar FO_FASE3.
3) Tratamiento correcto de sesiones externas al 3º GIOI:
   - Las sesiones de asignaturas externas usan el nº de alumnos de la hoja DOE.
   - Solo se penaliza "sin grupo" cuando la sesión pertenece a una asignatura GIOI y no se puede mapear.
4) Evaluación rápida con índices de solape y continuidad.
5) Exportación Excel visual: Dashboard, Resultados, Resultados DOE, Resumen Laboratorios e Historial Tabú.
Autor original: MGILMAR1
Versión final asistida: búsqueda tabú + presentación visual
"""

import os
import sys
import glob
import re
import time
import random
import datetime as dt
from collections import defaultdict, Counter
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.comments import Comment

VERSION_ENTREGA = "Gil_Miguel_5ªENTREGA - versión corregida: Excel de entrada = Excel de salida"

# =============================================================================
# 0. PARÁMETROS GENERALES
# =============================================================================
# Añadimos "PRACTIPLAN DATOS.xlsx" a los candidatos para que coincida exactamente con el archivo subido.
INPUT_CANDIDATES = [
    "PRACTIPLAN DATOS.xlsx",
    "Gil_Miguel_5ªENTREGA.xlsx",
    "Gil_Miguel_5ªENTREGA_PRACTIPLAN.xlsx",
]

OUTPUT_FILE = None
RANDOM_SEED = None

MAX_TIME_TABU = 55
MAX_ITER_TABU = 3000
TABU_TENURE = 11
SIN_MEJORA_MAX = 450
CANDIDATE_SESSIONS = 90

PESOS_FO3 = {
    "solape_laboratorio": 1,
    "exceso_aforo": 1,
    "software": 1,
    "continuidad_docente": 1,
    "sin_grupo": 1,
}

DIAS = {0: "Lunes", 1: "Martes", 2: "Miércoles", 3: "Jueves", 4: "Viernes"}
LABS = ["0.01", "0.02", "0.03", "S.01", "S.02", "S.03"]
CAPACIDAD = {"0.01": 30, "0.02": 30, "S.02": 30, "0.03": 20, "S.01": 20, "S.03": 20}
LABS_SOFTWARE = {"0.02", "S.02", "S.03"}
PLANTA = {lab: ("PB" if lab.startswith("0.") else "Sótano") for lab in LABS}

if RANDOM_SEED is not None:
    random.seed(RANDOM_SEED)

# =============================================================================
# 1. UTILIDADES GENERALES
# =============================================================================
def localizar_excel_entrada():
    """
    Localiza el Excel de entrada y devuelve su ruta absoluta.
    Busca entre los nombres candidatos y comprueba las hojas necesarias.
    """
    hojas_necesarias = {"Datos alumnado", "Datos practicas 3 GIOI", "Datos prácticas DOE"}
    script_dir = os.path.dirname(os.path.abspath(__file__))

    def ruta_absoluta(nombre_o_ruta):
        return nombre_o_ruta if os.path.isabs(nombre_o_ruta) else os.path.join(script_dir, nombre_o_ruta)

    if len(sys.argv) > 1 and sys.argv[1].lower().endswith(".xlsx"):
        ruta = ruta_absoluta(sys.argv[1])
        if os.path.exists(ruta):
            return ruta
        raise FileNotFoundError(f"No se ha encontrado el archivo Excel argumentado: {ruta}")

    for candidate in INPUT_CANDIDATES:
        ruta = ruta_absoluta(candidate)
        if os.path.exists(ruta):
            return ruta

    for candidate in glob.glob(os.path.join(script_dir, "*.xlsx")):
        if os.path.basename(candidate).startswith("~$"):
            continue
        try:
            wb_tmp = openpyxl.load_workbook(candidate, read_only=True, data_only=True)
            if hojas_necesarias.issubset(set(wb_tmp.sheetnames)):
                wb_tmp.close()
                return candidate
            wb_tmp.close()
        except Exception:
            pass

    raise FileNotFoundError(
        "No se ha encontrado ningún Excel válido con las hojas: "
        "Datos alumnado, Datos practicas 3 GIOI y Datos prácticas DOE."
    )

def solapan_minutos(a_ini, a_fin, b_ini, b_fin):
    return max(a_ini, b_ini) < min(a_fin, b_fin)

def solapan_dt(a_ini, a_fin, b_ini, b_fin):
    return max(a_ini, b_ini) < min(a_fin, b_fin)

def normalizar_si_no(valor):
    return str(valor).strip().lower() in ("sí", "si", "s", "yes", "y", "true", "1")

def parsear_incompatibilidades(texto):
    if texto is None:
        return []
    texto = str(texto).strip()
    if texto == "" or texto.lower() == "ninguna":
        return []
    partes = re.split(r"\s+y\s+", texto)
    resultado = []
    for p in partes:
        p = p.strip()
        if "toda la tarde" in p.lower():
            dia = p.split()[0]
            resultado.append((dia, 15 * 60, 21 * 60))
            continue
        m = re.search(
            r"(Lunes|Martes|Miércoles|Jueves|Viernes)\s+de\s+(\d+):(\d+)\s+a\s+(\d+):(\d+)", p
        )
        if m:
            dia = m.group(1)
            h1, m1, h2, m2 = map(int, m.groups()[1:])
            resultado.append((dia, h1 * 60 + m1, h2 * 60 + m2))
    return resultado

def limpiar_hojas(wb, nombres):
    for nombre in nombres:
        if nombre in wb.sheetnames:
            del wb[nombre]

def corregir_formulas_incompatibles(wb):
    if "Datos prácticas DOE" in wb.sheetnames:
        ws = wb["Datos prácticas DOE"]
        docentes = {str(ws.cell(r, 7).value).strip() for r in range(2, ws.max_row + 1) if ws.cell(r, 7).value}
        if ws.max_row >= 1394:
            ws.cell(1394, 6).value = "Número de profesores"
            ws.cell(1394, 7).value = len(docentes)

# =============================================================================
# 2. LECTURA DE DATOS
# =============================================================================
def leer_datos_base(wb):
    ws_alumnos = wb["Datos alumnado"]
    ws_practicas = wb["Datos practicas 3 GIOI"]
    nombres_asig = [ws_alumnos.cell(1, c).value for c in range(2, 8)]
    codigo_asig = {}
    for r in range(2, ws_practicas.max_row + 1):
        cod, nom = ws_practicas.cell(r, 1).value, ws_practicas.cell(r, 2).value
        if cod and nom:
            codigo_asig[nom] = int(cod)
    codigos_gioi = set(codigo_asig.values())
    grupos = defaultdict(set)
    sesiones = defaultdict(list)
    for r in range(2, ws_practicas.max_row + 1):
        cod, grp = ws_practicas.cell(r, 1).value, ws_practicas.cell(r, 4).value
        if not cod or not grp:
            continue
        cod = int(cod)
        grupos[cod].add(grp)
        f = ws_practicas.cell(r, 5).value
        h = ws_practicas.cell(r, 6).value
        m = ws_practicas.cell(r, 7).value
        d = ws_practicas.cell(r, 8).value
        if isinstance(f, dt.datetime):
            inicio = dt.datetime(f.year, f.month, f.day, int(h), int(m))
        elif isinstance(f, dt.date):
            inicio = dt.datetime(f.year, f.month, f.day, int(h), int(m))
        else:
            continue
        sesiones[(cod, grp)].append((inicio, inicio + dt.timedelta(minutes=int(d))))
    alumnos = []
    matricula = defaultdict(dict)
    incompat = {}
    for r in range(2, ws_alumnos.max_row + 1):
        nombre = ws_alumnos.cell(r, 1).value
        if not nombre or str(nombre).strip().lower().startswith("total"):
            continue
        nombre = str(nombre).strip()
        alumnos.append(nombre)
        for i, asig in enumerate(nombres_asig):
            valor = ws_alumnos.cell(r, 2 + i).value
            matricula[nombre][codigo_asig[asig]] = normalizar_si_no(valor)
        incompat[nombre] = parsear_incompatibilidades(ws_alumnos.cell(r, 9).value)
    return nombres_asig, codigo_asig, codigos_gioi, grupos, sesiones, alumnos, matricula, incompat

def aleatorizar_escenario(alumnos, matricula, incompat, codigos_asig):
    for a in alumnos:
        for cod in codigos_asig:
            matricula[a][cod] = random.random() < 0.6
        if random.random() < 0.3:
            incompat[a] = [(random.choice(list(DIAS.values())), 15 * 60, 17 * 60)]
        else:
            incompat[a] = []

def leer_sesiones_doe(wb, codigos_gioi, sheet_name="Datos prácticas DOE"):
    ws = wb[sheet_name]
    out = []
    for r in range(2, ws.max_row + 1):
        fecha = ws.cell(r, 1).value
        cod = ws.cell(r, 2).value
        escuela = ws.cell(r, 3).value
        h = ws.cell(r, 4).value
        m = ws.cell(r, 5).value
        dur = ws.cell(r, 6).value
        docente = ws.cell(r, 7).value
        num_alumnos_excel = ws.cell(r, 8).value
        sw = ws.cell(r, 9).value
        if not fecha or cod is None or h is None or m is None or dur is None:
            continue
        try:
            cod_int = int(cod)
            alumnos_excel = int(num_alumnos_excel) if num_alumnos_excel not in (None, "") else 0
        except Exception:
            continue
        if isinstance(fecha, dt.datetime):
            inicio = dt.datetime(fecha.year, fecha.month, fecha.day, int(h), int(m))
        elif isinstance(fecha, dt.date):
            inicio = dt.datetime(fecha.year, fecha.month, fecha.day, int(h), int(m))
        else:
            continue
        fin = inicio + dt.timedelta(minutes=int(dur))
        out.append({
            "idx": r,
            "cod_asig": cod_int,
            "escuela": str(escuela).strip() if escuela else "",
            "inicio": inicio,
            "fin": fin,
            "docente": str(docente).strip() if docente else "",
            "n_alumnos_excel": alumnos_excel,
            "requiere_sw": normalizar_si_no(sw),
            "es_gioi": cod_int in codigos_gioi,
        })
    return out

# =============================================================================
# 3. FASE 2 — GREEDY CONSTRUCTIVO PARA ALUMNOS-GRUPOS
# =============================================================================
def calcular_conflictos_alumno(alumno, cod_asig, grupo_id, sesiones, incompat, asignacion):
    fallos = 0
    for ini, fin in sesiones[(cod_asig, grupo_id)]:
        dia = DIAS[ini.weekday()]
        i1, i2 = ini.hour * 60 + ini.minute, fin.hour * 60 + fin.minute
        for d, j1, j2 in incompat[alumno]:
            if d == dia and solapan_minutos(i1, i2, j1, j2):
                fallos += 1
        for asig_ya, grp_ya in asignacion[alumno].items():
            for ini2, fin2 in sesiones[(asig_ya, grp_ya)]:
                if ini.date() == ini2.date() and solapan_dt(ini, fin, ini2, fin2):
                    fallos += 1
    return fallos

def asignar_alumnos_greedy_equilibrado(alumnos, matricula, codigos_asig, grupos, sesiones, incompat):
    asignacion = defaultdict(dict)
    carga_grupo = defaultdict(lambda: defaultdict(int))
    alumnos_ordenados = sorted(
        alumnos,
        key=lambda a: (sum(matricula[a].values()), len(incompat[a])),
        reverse=True,
    )
    for a in alumnos_ordenados:
        for cod in codigos_asig:
            if not matricula[a].get(cod):
                continue
            mejor_grp = None
            mejor_score = float("inf")
            for grp in sorted(grupos[cod], key=str):
                conflictos = calcular_conflictos_alumno(a, cod, grp, sesiones, incompat, asignacion)
                carga = carga_grupo[cod][grp]
                score = conflictos * 100 + carga
                if score < mejor_score:
                    mejor_score = score
                    mejor_grp = grp
            asignacion[a][cod] = mejor_grp
            carga_grupo[cod][mejor_grp] += 1
    fo2_conflictos = sum(
        calcular_conflictos_alumno(a, c, g, sesiones, incompat, asignacion)
        for a in alumnos for c, g in asignacion[a].items()
    )
    return asignacion, carga_grupo, fo2_conflictos

# =============================================================================
# 4. FASE 3 — MAPEO, EVALUACIÓN Y BÚSQUEDA TABÚ
# =============================================================================
def mapear_sesion_doe_a_grupo(sesiones_doe, sesiones_fase2):
    idx = defaultdict(list)
    for (cod, grp), lista in sesiones_fase2.items():
        idx[int(cod)].append((grp, lista))
    grupo_por_sesion = {}
    for s in sesiones_doe:
        key = (s["cod_asig"], s["inicio"])
        if not s["es_gioi"]:
            grupo_por_sesion[key] = "Externa"
            continue
        best_grp = None
        best_overlap = -1
        for grp, lista in idx.get(int(s["cod_asig"]), []):
            for ini2, fin2 in lista:
                if s["inicio"] == ini2 and s["fin"] == fin2:
                    best_grp = grp
                    best_overlap = float("inf")
                    break
                if solapan_dt(s["inicio"], s["fin"], ini2, fin2):
                    overlap = (min(s["fin"], fin2) - max(s["inicio"], ini2)).total_seconds()
                    if overlap > best_overlap:
                        best_overlap = overlap
                        best_grp = grp
            if best_overlap == float("inf"):
                break
        grupo_por_sesion[key] = best_grp
    return grupo_por_sesion

def contar_alumnos_por_sesion_doe(sesiones_doe, asignacion, alumnos, sesiones_fase2):
    grupo_mapeado = mapear_sesion_doe_a_grupo(sesiones_doe, sesiones_fase2)
    n_alumnos = {}
    for s in sesiones_doe:
        key = (s["cod_asig"], s["inicio"])
        if not s["es_gioi"]:
            n_alumnos[key] = s["n_alumnos_excel"]
            continue
        grp = grupo_mapeado.get(key)
        if grp is None:
            n_alumnos[key] = 0
            continue
        n_alumnos[key] = sum(1 for a in alumnos if asignacion[a].get(s["cod_asig"]) == grp)
    return n_alumnos, grupo_mapeado

def construir_indices_fase3(sesiones_doe, n_alumnos, grupo_mapeado):
    ses_by_key = {(s["cod_asig"], s["inicio"]): s for s in sesiones_doe}
    keys = list(ses_by_key.keys())
    by_date = defaultdict(list)
    for k, s in ses_by_key.items():
        by_date[s["inicio"].date()].append(k)
    overlap_neighbors = {k: [] for k in keys}
    for klist in by_date.values():
        klist.sort(key=lambda kk: ses_by_key[kk]["inicio"])
        for i, k1 in enumerate(klist):
            s1 = ses_by_key[k1]
            for j in range(i + 1, len(klist)):
                k2 = klist[j]
                s2 = ses_by_key[k2]
                if s2["inicio"] >= s1["fin"]:
                    break
                if solapan_dt(s1["inicio"], s1["fin"], s2["inicio"], s2["fin"]):
                    overlap_neighbors[k1].append(k2)
                    overlap_neighbors[k2].append(k1)
    continuity_neighbors = {k: [] for k in keys}
    by_doc = defaultdict(list)
    for k, s in ses_by_key.items():
        if s["docente"]:
            by_doc[s["docente"]].append(k)
    for klist in by_doc.values():
        klist.sort(key=lambda kk: ses_by_key[kk]["inicio"])
        for i in range(len(klist) - 1):
            kA, kB = klist[i], klist[i + 1]
            if ses_by_key[kA]["fin"] == ses_by_key[kB]["inicio"]:
                continuity_neighbors[kA].append(kB)
                continuity_neighbors[kB].append(kA)
    return {
        "ses_by_key": ses_by_key,
        "keys": keys,
        "overlap_neighbors": overlap_neighbors,
        "continuity_neighbors": continuity_neighbors,
        "n_alumnos": n_alumnos,
        "grupo_mapeado": grupo_mapeado,
    }

def penalizacion_individual(k, lab, indices):
    s = indices["ses_by_key"][k]
    pen = 0
    if s["es_gioi"] and indices["grupo_mapeado"].get(k) is None:
        pen += PESOS_FO3["sin_grupo"]
    if indices["n_alumnos"].get(k, 0) > CAPACIDAD[lab]:
        pen += PESOS_FO3["exceso_aforo"]
    if s["requiere_sw"] and lab not in LABS_SOFTWARE:
        pen += PESOS_FO3["software"]
    return pen

def evaluar_solucion_fase3_rapida(lab_asignado, indices):
    detalle = {"solape_laboratorio": 0, "exceso_aforo": 0, "software": 0, "continuidad_docente": 0, "sin_grupo": 0}
    for k in indices["keys"]:
        lab = lab_asignado[k]
        s = indices["ses_by_key"][k]
        if s["es_gioi"] and indices["grupo_mapeado"].get(k) is None:
            detalle["sin_grupo"] += 1
        if indices["n_alumnos"].get(k, 0) > CAPACIDAD[lab]:
            detalle["exceso_aforo"] += 1
        if s["requiere_sw"] and lab not in LABS_SOFTWARE:
            detalle["software"] += 1
    vistos = set()
    for k in indices["keys"]:
        for v in indices["overlap_neighbors"][k]:
            par = tuple(sorted((k, v), key=str))
            if par in vistos:
                continue
            vistos.add(par)
            if lab_asignado[k] == lab_asignado[v]:
                detalle["solape_laboratorio"] += 1
    vistos = set()
    for k in indices["keys"]:
        for v in indices["continuity_neighbors"][k]:
            par = tuple(sorted((k, v), key=str))
            if par in vistos:
                continue
            vistos.add(par)
            if lab_asignado[k] != lab_asignado[v]:
                detalle["continuidad_docente"] += 1
    fo = sum(detalle[k] * PESOS_FO3[k] for k in detalle)
    return fo, detalle

def delta_movimiento(k, lab_nuevo, lab_asignado, indices):
    lab_antiguo = lab_asignado[k]
    if lab_antiguo == lab_nuevo:
        return 0
    delta = -penalizacion_individual(k, lab_antiguo, indices)
    delta += penalizacion_individual(k, lab_nuevo, indices)
    for v in indices["overlap_neighbors"][k]:
        lab_v = lab_asignado[v]
        if lab_v == lab_antiguo:
            delta -= PESOS_FO3["solape_laboratorio"]
        if lab_v == lab_nuevo:
            delta += PESOS_FO3["solape_laboratorio"]
    for v in indices["continuity_neighbors"][k]:
        lab_v = lab_asignado[v]
        antes = 1 if lab_antiguo != lab_v else 0
        despues = 1 if lab_nuevo != lab_v else 0
        delta += (despues - antes) * PESOS_FO3["continuidad_docente"]
    return delta

def asignar_labs_aleatorio(sesiones_doe):
    return {(s["cod_asig"], s["inicio"]): random.choice(LABS) for s in sesiones_doe}

def solucion_inicial_greedy_fase3(indices):
    lab_asignado = {}
    keys_ordenadas = sorted(indices["keys"], key=lambda k: (k[1], k[0]))
    for k in keys_ordenadas:
        mejor_lab = None
        mejor_coste = float("inf")
        labs = LABS[:]
        random.shuffle(labs)
        for lab in labs:
            coste = penalizacion_individual(k, lab, indices)
            for v in indices["overlap_neighbors"][k]:
                if v in lab_asignado and lab_asignado[v] == lab:
                    coste += PESOS_FO3["solape_laboratorio"]
            for v in indices["continuity_neighbors"][k]:
                if v in lab_asignado and lab_asignado[v] != lab:
                    coste += PESOS_FO3["continuidad_docente"]
            if coste < mejor_coste:
                mejor_coste = coste
                mejor_lab = lab
        lab_asignado[k] = mejor_lab if mejor_lab is not None else random.choice(LABS)
    return lab_asignado

def claves_conflictivas(lab_asignado, indices):
    conflictivas = set()
    for k in indices["keys"]:
        s = indices["ses_by_key"][k]
        lab = lab_asignado[k]
        if s["es_gioi"] and indices["grupo_mapeado"].get(k) is None:
            conflictivas.add(k)
        if indices["n_alumnos"].get(k, 0) > CAPACIDAD[lab]:
            conflictivas.add(k)
        if s["requiere_sw"] and lab not in LABS_SOFTWARE:
            conflictivas.add(k)
        for v in indices["overlap_neighbors"][k]:
            if lab_asignado[v] == lab:
                conflictivas.add(k)
                conflictivas.add(v)
        for v in indices["continuity_neighbors"][k]:
            if lab_asignado[v] != lab:
                conflictivas.add(k)
                conflictivas.add(v)
    return list(conflictivas)

def busqueda_tabu_laboratorios(sesiones_doe, n_alumnos, grupo_mapeado, max_iter=MAX_ITER_TABU, tenure=TABU_TENURE, max_time=MAX_TIME_TABU, candidate_sessions=CANDIDATE_SESSIONS, sin_mejora_max=SIN_MEJORA_MAX):
    t0 = time.time()
    indices = construir_indices_fase3(sesiones_doe, n_alumnos, grupo_mapeado)
    keys = indices["keys"]
    sol_aleatoria = asignar_labs_aleatorio(sesiones_doe)
    fo_aleatoria, detalle_aleatoria = evaluar_solucion_fase3_rapida(sol_aleatoria, indices)
    actual = solucion_inicial_greedy_fase3(indices)
    fo_actual, detalle_actual = evaluar_solucion_fase3_rapida(actual, indices)
    mejor = actual.copy()
    fo_mejor = fo_actual
    detalle_mejor = detalle_actual.copy()
    tabu_until = {}
    historial = []
    iter_sin_mejora = 0
    iteracion = 0
    while iteracion < max_iter and (time.time() - t0) < max_time and iter_sin_mejora < sin_mejora_max:
        iteracion += 1
        for attr, vence in list(tabu_until.items()):
            if vence <= iteracion:
                del tabu_until[attr]
        conflictivas = claves_conflictivas(actual, indices)
        if conflictivas:
            if candidate_sessions is None or candidate_sessions >= len(conflictivas):
                keys_iter = conflictivas[:]
            else:
                keys_iter = random.sample(conflictivas, candidate_sessions)
        else:
            keys_iter = random.sample(keys, min(candidate_sessions or len(keys), len(keys)))
        mejor_movimiento = None
        mejor_delta = float("inf")
        mejor_fo_vecino = float("inf")
        for k in keys_iter:
            lab_actual = actual[k]
            labs = [lab for lab in LABS if lab != lab_actual]
            random.shuffle(labs)
            for lab_nuevo in labs:
                d = delta_movimiento(k, lab_nuevo, actual, indices)
                fo_vecino = fo_actual + d
                atributo_prohibido = (k, lab_nuevo)
                es_tabu = atributo_prohibido in tabu_until
                aspiracion = fo_vecino < fo_mejor
                if es_tabu and not aspiracion:
                    continue
                if fo_vecino < mejor_fo_vecino or (fo_vecino == mejor_fo_vecino and d < mejor_delta):
                    mejor_fo_vecino = fo_vecino
                    mejor_delta = d
                    mejor_movimiento = (k, lab_actual, lab_nuevo)
        if mejor_movimiento is None:
            break
        k_mov, lab_origen, lab_destino = mejor_movimiento
        actual[k_mov] = lab_destino
        fo_actual = mejor_fo_vecino
        tabu_until[(k_mov, lab_origen)] = iteracion + tenure
        if fo_actual < fo_mejor:
            mejor = actual.copy()
            fo_mejor = fo_actual
            detalle_mejor = evaluar_solucion_fase3_rapida(mejor, indices)[1]
            iter_sin_mejora = 0
        else:
            iter_sin_mejora += 1
        historial.append({
            "iteracion": iteracion,
            "fo_actual": fo_actual,
            "fo_mejor": fo_mejor,
            "delta": mejor_delta,
            "movimiento": f"{k_mov[0]} {k_mov[1].strftime('%Y-%m-%d %H:%M')} : {lab_origen} -> {lab_destino}",
        })
        if fo_mejor == 0:
            break
    info = {
        "metodo": "Búsqueda Tabú",
        "fo_aleatoria": fo_aleatoria,
        "detalle_aleatoria": detalle_aleatoria,
        "fo_inicial_greedy": evaluar_solucion_fase3_rapida(solucion_inicial_greedy_fase3(indices), indices)[0],
        "fo_final_tabu": fo_mejor,
        "detalle_final_tabu": detalle_mejor,
        "iteraciones": iteracion,
        "tenencia_tabu": tenure,
        "tiempo_seg": round(time.time() - t0, 2),
        "historial": historial,
        "pesos": PESOS_FO3.copy(),
        "semilla": RANDOM_SEED if RANDOM_SEED is not None else "Aleatoria",
    }
    return mejor, fo_mejor, detalle_mejor, info

# =============================================================================
# 5. EXPORTACIÓN VISUAL A EXCEL
# =============================================================================
COLORS = {
    "navy": "1F4E78", "blue": "5B9BD5", "teal": "00A6A6", "green": "70AD47",
    "red": "C00000", "orange": "ED7D31", "yellow": "FFC000", "light_blue": "D9EAF7",
    "light_green": "E2F0D9", "light_red": "FCE4D6", "light_yellow": "FFF2CC",
    "gray": "F2F2F2", "dark_gray": "595959", "white": "FFFFFF",
}
THIN_GRAY = Side(style="thin", color="D9E2F3")
BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)

def apply_title(ws, cell_range, title, subtitle=None):
    start = cell_range.split(":")[0]
    ws.merge_cells(cell_range)
    c = ws[start]
    c.value = title if subtitle is None else f"{title}\n{subtitle}"
    c.font = Font(bold=True, size=16, color=COLORS["white"])
    c.fill = PatternFill("solid", fgColor=COLORS["navy"])
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

def style_header_range(rng):
    for item in rng:
        cells = item if isinstance(item, (list, tuple)) else (item,)
        for cell in cells:
            cell.font = Font(bold=True, color=COLORS["white"])
            cell.fill = PatternFill("solid", fgColor=COLORS["navy"])
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = BORDER

def style_kpi(cell, label_cell=None, good=True):
    cell.font = Font(bold=True, size=18, color=COLORS["green"] if good else COLORS["red"])
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.fill = PatternFill("solid", fgColor=COLORS["white"])
    cell.border = BORDER
    if label_cell:
        label_cell.font = Font(bold=True, color=COLORS["dark_gray"])
        label_cell.fill = PatternFill("solid", fgColor=COLORS["gray"])
        label_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        label_cell.border = BORDER

def safe_table(ws, ref, name):
    return

def auto_width(ws, min_width=8, max_width=38):
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_len = 0
        for cell in col_cells:
            val = cell.value
            if val is None:
                continue
            max_len = max(max_len, len(str(val)))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, min_width), max_width)

def exportar_resultados_fase2(wb, nombres_asig, codigo_asig, alumnos, matricula, asignacion, total_matriculados, fo2):
    if "Resultados" in wb.sheetnames:
        del wb["Resultados"]
    ws = wb.create_sheet("Resultados")
    apply_title(ws, "A1:H2", "PRACTIPLAN — Resultados Fase 2", "Asignación de alumnos a grupos de prácticas")
    ws["A4"] = "Función objetivo Fase 2"
    ws["B4"] = fo2
    style_kpi(ws["B4"], ws["A4"], good=fo2 == 0)
    ws["A6"] = "Total matriculados"
    for j, asig in enumerate(nombres_asig, start=2):
        ws.cell(6, j).value = total_matriculados[codigo_asig[asig]]
    style_header_range(ws["A6:G6"])
    start_row = 9
    ws.cell(start_row, 1).value = "Alumno"
    for j, asig in enumerate(nombres_asig, start=2):
        ws.cell(start_row, j).value = asig
    style_header_range(ws[start_row][0:7])
    r = start_row + 1
    for a in alumnos:
        ws.cell(r, 1).value = a
        for j, asig in enumerate(nombres_asig, start=2):
            cod = codigo_asig[asig]
            ws.cell(r, j).value = asignacion[a].get(cod, "No Matr.") if matricula[a].get(cod) else "No Matr."
        r += 1
    for row in ws.iter_rows(min_row=start_row, max_row=r-1, min_col=1, max_col=7):
        for cell in row:
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center" if cell.column > 1 else "left", vertical="center")
    ws.freeze_panes = "A10"
    ws.auto_filter.ref = f"A{start_row}:G{r-1}"
    safe_table(ws, f"A{start_row}:G{r-1}", "TablaResultadosFase2")
    auto_width(ws)

def exportar_resultados_fase3(wb, sesiones_doe, lab_asignado, n_alumnos, grupo_mapeado, fo3, detalle3, info3):
    if "Resultados DOE" in wb.sheetnames:
        del wb["Resultados DOE"]
    ws = wb.create_sheet("Resultados DOE")
    apply_title(ws, "A1:Q2", "PRACTIPLAN — Resultados Fase 3", "Asignación de sesiones DOE a laboratorios con Búsqueda Tabú")
    kpis = [
        ("FO Tabú", fo3, fo3 <= info3["fo_aleatoria"]),
        ("FO aleatoria", info3["fo_aleatoria"], False),
        ("Mejora %", f"{(1 - fo3 / info3['fo_aleatoria']) * 100:.1f}%" if info3["fo_aleatoria"] else "—", True),
        ("Iteraciones", info3["iteraciones"], True),
        ("Tiempo (s)", info3["tiempo_seg"], True),
    ]
    col = 1
    for label, value, good in kpis:
        ws.cell(4, col).value = label
        ws.cell(5, col).value = value
        style_kpi(ws.cell(5, col), ws.cell(4, col), good=good)
        col += 2

    ws["A8"] = "Desglose de infracciones"
    ws["A8"].font = Font(bold=True, color=COLORS["navy"], size=12)
    resumen = [
        ("Solape laboratorio", detalle3["solape_laboratorio"]),
        ("Exceso aforo", detalle3["exceso_aforo"]),
        ("Software incorrecto", detalle3["software"]),
        ("Continuidad docente", detalle3["continuidad_docente"]),
        ("Sesión GIOI sin grupo", detalle3["sin_grupo"]),
    ]
    ws["A9"], ws["B9"] = "Tipo", "Nº infracciones"
    style_header_range(ws["A9:B9"])
    for i, (label, val) in enumerate(resumen, start=10):
        ws.cell(i, 1).value = label
        ws.cell(i, 2).value = val
        for c in range(1, 3):
            ws.cell(i, c).border = BORDER
            ws.cell(i, c).alignment = Alignment(horizontal="left" if c == 1 else "center")

    start_row = 17
    headers = [
        "Sesión", "Tipo", "Escuela", "Asignatura", "Grupo", "Fecha", "Inicio", "Fin",
        "Docente", "Laboratorio", "Planta", "Alumnos", "Capacidad", "Requiere SW",
        "Software OK", "Aforo OK", "Incidencias"
    ]
    for j, h in enumerate(headers, 1):
        ws.cell(start_row, j).value = h
    style_header_range(ws[start_row][0:len(headers)])

    sesiones_sorted = sorted(sesiones_doe, key=lambda s: (s["inicio"], s["cod_asig"], s["idx"]))
    r = start_row + 1
    for i, s in enumerate(sesiones_sorted, start=1):
        key = (s["cod_asig"], s["inicio"])
        lab = lab_asignado.get(key, "")
        cap = CAPACIDAD.get(lab, None)
        grp = grupo_mapeado.get(key)
        alumnos_s = n_alumnos.get(key, 0)
        software_ok = (not s["requiere_sw"]) or lab in LABS_SOFTWARE
        aforo_ok = cap is not None and alumnos_s <= cap
        incidencias = []
        if s["es_gioi"] and grp is None:
            incidencias.append("Sin grupo")
        if not software_ok:
            incidencias.append("Software")
        if not aforo_ok:
            incidencias.append("Aforo")

        values = [
            i, "GIOI" if s["es_gioi"] else "Externa", s["escuela"], s["cod_asig"],
            grp if grp is not None else "", s["inicio"].date(), s["inicio"].strftime("%H:%M"),
            s["fin"].strftime("%H:%M"), s["docente"], lab, PLANTA.get(lab, ""),
            alumnos_s, cap, "Sí" if s["requiere_sw"] else "No",
            "OK" if software_ok else "NO", "OK" if aforo_ok else "NO",
            ", ".join(incidencias) if incidencias else "—",
        ]
        for j, val in enumerate(values, 1):
            c = ws.cell(r, j)
            c.value = val
            c.border = BORDER
            c.alignment = Alignment(horizontal="center" if j not in (3, 9, 17) else "left", vertical="center", wrap_text=True)
        ws.cell(r, 6).number_format = "yyyy-mm-dd"
        if s["es_gioi"]:
            ws.cell(r, 2).fill = PatternFill("solid", fgColor=COLORS["light_blue"])
        if not software_ok:
            ws.cell(r, 15).fill = PatternFill("solid", fgColor=COLORS["light_red"])
        else:
            ws.cell(r, 15).fill = PatternFill("solid", fgColor=COLORS["light_green"])
        if not aforo_ok:
            ws.cell(r, 16).fill = PatternFill("solid", fgColor=COLORS["light_red"])
        else:
            ws.cell(r, 16).fill = PatternFill("solid", fgColor=COLORS["light_green"])
        if incidencias:
            ws.cell(r, 17).fill = PatternFill("solid", fgColor=COLORS["light_yellow"])
        r += 1

    ws.freeze_panes = f"A{start_row+1}"
    ws.auto_filter.ref = f"A{start_row}:Q{r-1}"
    safe_table(ws, f"A{start_row}:Q{r-1}", "TablaResultadosDOE")
    auto_width(ws, max_width=34)
    ws.column_dimensions["I"].width = 26
    ws.column_dimensions["Q"].width = 24

def exportar_historial_tabu(wb, info3):
    if "Historial Tabu" in wb.sheetnames:
        del wb["Historial Tabu"]
    ws = wb.create_sheet("Historial Tabu")
    apply_title(ws, "A1:E2", "Historial de Búsqueda Tabú", "Evolución de la FO durante el proceso")
    headers = ["Iteración", "FO actual", "FO mejor", "Delta", "Movimiento"]
    for j, h in enumerate(headers, 1):
        ws.cell(4, j).value = h
    style_header_range(ws[4][0:5])
    for r, h in enumerate(info3.get("historial", []), start=5):
        ws.cell(r, 1).value = h["iteracion"]
        ws.cell(r, 2).value = h["fo_actual"]
        ws.cell(r, 3).value = h["fo_mejor"]
        ws.cell(r, 4).value = h["delta"]
        ws.cell(r, 5).value = h["movimiento"]
        for c in range(1, 6):
            ws.cell(r, c).border = BORDER
    auto_width(ws, max_width=60)
    ws.freeze_panes = "A5"
    if ws.max_row >= 6:
        chart = LineChart()
        chart.title = "Evolución de la función objetivo"
        chart.y_axis.title = "FO"
        chart.x_axis.title = "Iteración"
        data = Reference(ws, min_col=2, max_col=3, min_row=4, max_row=ws.max_row)
        cats = Reference(ws, min_col=1, min_row=5, max_row=ws.max_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.height = 9
        chart.width = 18
        ws.add_chart(chart, "G4")

def exportar_resumen_laboratorios(wb, sesiones_doe, lab_asignado, n_alumnos):
    if "Resumen Laboratorios" in wb.sheetnames:
        del wb["Resumen Laboratorios"]
    ws = wb.create_sheet("Resumen Laboratorios")
    apply_title(ws, "A1:H2", "Resumen por laboratorio", "Carga, capacidad y uso de software")
    headers = ["Laboratorio", "Planta", "Capacidad", "Nº sesiones", "Alumnos acumulados", "Sesiones con SW", "Ocupación media", "% sesiones"]
    for j, h in enumerate(headers, 1):
        ws.cell(4, j).value = h
    style_header_range(ws[4][0:8])
    total_ses = len(sesiones_doe)
    stats = {lab: {"sesiones": 0, "alumnos": 0, "sw": 0} for lab in LABS}
    for s in sesiones_doe:
        k = (s["cod_asig"], s["inicio"])
        lab = lab_asignado[k]
        stats[lab]["sesiones"] += 1
        stats[lab]["alumnos"] += n_alumnos.get(k, 0)
        if s["requiere_sw"]:
            stats[lab]["sw"] += 1
    for i, lab in enumerate(LABS, start=5):
        ses = stats[lab]["sesiones"]
        alumnos_acum = stats[lab]["alumnos"]
        ocup_media = alumnos_acum / (ses * CAPACIDAD[lab]) if ses else 0
        values = [
            lab, PLANTA[lab], CAPACIDAD[lab], ses, alumnos_acum,
            stats[lab]["sw"], ocup_media, ses / total_ses if total_ses else 0
        ]
        for j, val in enumerate(values, 1):
            ws.cell(i, j).value = val
            ws.cell(i, j).border = BORDER
            ws.cell(i, j).alignment = Alignment(horizontal="center")
        ws.cell(i, 7).number_format = "0.0%"
        ws.cell(i, 8).number_format = "0.0%"
    auto_width(ws)
    chart = BarChart()
    chart.title = "Sesiones asignadas por laboratorio"
    chart.y_axis.title = "Nº sesiones"
    chart.x_axis.title = "Laboratorio"
    data = Reference(ws, min_col=4, min_row=4, max_row=10)
    cats = Reference(ws, min_col=1, min_row=5, max_row=10)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 8
    chart.width = 15
    ws.add_chart(chart, "J4")

def exportar_dashboard(wb, fo2, fo3, detalle3, info3):
    if "Dashboard" in wb.sheetnames:
        del wb["Dashboard"]
    ws = wb.create_sheet("Dashboard", 0)
    apply_title(ws, "A1:M2", "PRACTIPLAN GIOI-DOE", "Panel ejecutivo de resultados finales")
    kpi_data = [
        ("FO Fase 2", fo2),
        ("FO aleatoria F3", info3["fo_aleatoria"]),
        ("FO Tabú F3", fo3),
        ("Mejora Tabú", f"{(1 - fo3 / info3['fo_aleatoria']) * 100:.1f}%" if info3["fo_aleatoria"] else "—"),
        ("Tiempo Tabú", f"{info3['tiempo_seg']} s"),
        ("Iteraciones", info3["iteraciones"]),
    ]
    positions = [("A4", "B4"), ("C4", "D4"), ("E4", "F4"), ("G4", "H4"), ("I4", "J4"), ("K4", "L4")]
    for (label, value), (lab_cell, val_cell) in zip(kpi_data, positions):
        ws[lab_cell] = label
        ws[val_cell] = value
        style_kpi(ws[val_cell], ws[lab_cell], good=True)

    ws["A8"] = "Comparación FO"
    ws["A8"].font = Font(bold=True, size=12, color=COLORS["navy"])
    ws["A9"], ws["B9"] = "Método", "FO"
    rows = [("Aleatorio", info3["fo_aleatoria"]), ("Greedy inicial", info3["fo_inicial_greedy"]), ("Búsqueda Tabú", fo3)]
    style_header_range(ws["A9:B9"])
    for i, row in enumerate(rows, 10):
        ws.cell(i, 1).value, ws.cell(i, 2).value = row
        ws.cell(i, 1).border = ws.cell(i, 2).border = BORDER

    ws["D8"] = "Infracciones finales"
    ws["D8"].font = Font(bold=True, size=12, color=COLORS["navy"])
    ws["D9"], ws["E9"] = "Tipo", "Cantidad"
    style_header_range(ws["D9:E9"])
    breakdown = [
        ("Solapes", detalle3["solape_laboratorio"]),
        ("Aforo", detalle3["exceso_aforo"]),
        ("Software", detalle3["software"]),
        ("Docente", detalle3["continuidad_docente"]),
        ("Sin grupo GIOI", detalle3["sin_grupo"]),
    ]
    for i, row in enumerate(breakdown, 10):
        ws.cell(i, 4).value, ws.cell(i, 5).value = row
        ws.cell(i, 4).border = ws.cell(i, 5).border = BORDER

    ws["A37"] = "Lectura rápida"
    ws["A37"].font = Font(bold=True, size=12, color=COLORS["navy"])
    ws["A38"] = (
        "La Búsqueda Tabú parte de una solución inicial y explora cambios de laboratorio. "
        "La lista tabú evita ciclos y el criterio de aspiración permite aceptar un movimiento tabú "
        "si mejora la mejor solución global. Las sesiones externas usan el número de alumnos del Excel; "
        "solo se penaliza ausencia de grupo cuando la sesión pertenece a una asignatura de 3º GIOI."
    )
    ws["A38"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A38:M40")
    ws["A38"].fill = PatternFill("solid", fgColor=COLORS["light_blue"])
    ws["A38"].border = BORDER

    bar = BarChart()
    bar.title = "Comparación de función objetivo"
    bar.y_axis.title = "FO"
    data = Reference(ws, min_col=2, min_row=9, max_row=12)
    cats = Reference(ws, min_col=1, min_row=10, max_row=12)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.height = 8
    bar.width = 13
    ws.add_chart(bar, "G8")

    pie = PieChart()
    pie.title = "Desglose de infracciones"
    labels = Reference(ws, min_col=4, min_row=10, max_row=14)
    data = Reference(ws, min_col=5, min_row=9, max_row=14)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.height = 8
    pie.width = 12
    ws.add_chart(pie, "G23")

    for col in range(1, 14):
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 30
    ws.sheet_view.showGridLines = False

def exportar_readme(wb):
    if "README" in wb.sheetnames:
        del wb["README"]
    ws = wb.create_sheet("README")
    apply_title(ws, "A1:H2", "README — PRACTIPLAN", "Guía rápida de interpretación del archivo")
    contenido = [
        ["Hoja", "Contenido"],
        ["Dashboard", "Panel visual con KPIs, comparación aleatorio/greedy/tabú y gráficos principales."],
        ["Resultados", "Asignación de alumnos a grupos de prácticas y función objetivo de Fase 2."],
        ["Resultados DOE", "Detalle de cada sesión DOE con laboratorio asignado, capacidad, software e incidencias."],
        ["Resumen Laboratorios", "Resumen agregado por laboratorio: sesiones, alumnos, ocupación y software."],
        ["Historial Tabu", "Evolución de la función objetivo durante la búsqueda tabú."],
    ]
    for r, row in enumerate(contenido, 4):
        for c, val in enumerate(row, 1):
            ws.cell(r, c).value = val
            ws.cell(r, c).border = BORDER
            ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")
    style_header_range(ws[4][0:2])
    ws["A12"] = "Puntos de mejora incorporados"
    ws["A12"].font = Font(bold=True, color=COLORS["navy"], size=12)
    mejoras = [
        "Separación entre sesiones GIOI y sesiones externas para evitar penalizar grupos inexistentes.",
        "Búsqueda Tabú con memoria por atributos y criterio de aspiración.",
        "Solución inicial greedy antes de iniciar la búsqueda local.",
        "Dashboard y tablas con formato condicional para facilitar la lectura.",
        "Historial de iteraciones para justificar el comportamiento experimental.",
    ]
    for i, mejora in enumerate(mejoras, 13):
        ws.cell(i, 1).value = f"• {mejora}"
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=8)
        ws.cell(i, 1).alignment = Alignment(wrap_text=True)
    auto_width(ws, max_width=70)
    ws.column_dimensions["B"].width = 80
    ws.sheet_view.showGridLines = False

# =============================================================================
# 6. PROGRAMA PRINCIPAL
# =============================================================================
def main():
    print(VERSION_ENTREGA)
    input_file = localizar_excel_entrada()
    wb = openpyxl.load_workbook(input_file)
    nombres_asig, codigo_asig, codigos_gioi, grupos, sesiones, alumnos, matricula, incompat = leer_datos_base(wb)
    codigos_ordenados = list(codigo_asig.values())
    aleatorizar_escenario(alumnos, matricula, incompat, codigos_ordenados)
    total_matriculados = {
        cod: sum(1 for a in alumnos if matricula[a].get(cod, False)) for cod in codigos_ordenados
    }
    asignacion, carga_grupo, fo2 = asignar_alumnos_greedy_equilibrado(
        alumnos, matricula, codigos_ordenados, grupos, sesiones, incompat
    )
    sesiones_doe = leer_sesiones_doe(wb, codigos_gioi)
    n_alumnos_doe, grupo_mapeado = contar_alumnos_por_sesion_doe(
        sesiones_doe, asignacion, alumnos, sesiones
    )
    lab_asignado, fo3, detalle3, info3 = busqueda_tabu_laboratorios(
        sesiones_doe, n_alumnos_doe, grupo_mapeado
    )
    limpiar_hojas(wb, ["Dashboard", "Resultados", "Resultados DOE", "Resumen Laboratorios", "Historial Tabu", "README"])
    exportar_dashboard(wb, fo2, fo3, detalle3, info3)
    exportar_resultados_fase2(wb, nombres_asig, codigo_asig, alumnos, matricula, asignacion, total_matriculados, fo2)
    exportar_resultados_fase3(wb, sesiones_doe, lab_asignado, n_alumnos_doe, grupo_mapeado, fo3, detalle3, info3)
    exportar_resumen_laboratorios(wb, sesiones_doe, lab_asignado, n_alumnos_doe)
    exportar_historial_tabu(wb, info3)
    corregir_formulas_incompatibles(wb)
    exportar_readme(wb)
    
    wb["Dashboard"]["A42"] = "Parámetros"
    wb["Dashboard"]["A42"].font = Font(bold=True, color=COLORS["navy"])
    wb["Dashboard"]["A43"] = f"Tenencia tabú={TABU_TENURE}; max_iter={MAX_ITER_TABU}; max_time={MAX_TIME_TABU}s; seed={info3['semilla']}"
    wb["Dashboard"]["A43"].comment = Comment("Parámetros principales de la ejecución final.", "PRACTIPLAN")
    wb["Dashboard"].merge_cells("A43:M43")
    wb["Dashboard"]["A43"].alignment = Alignment(wrap_text=True)
    
    output_file = input_file
    wb.save(output_file)
    print("=== PRACTIPLAN FINAL ===")
    print(f"Excel entrada: {input_file}")
    print(f"Excel salida sobrescrita: {output_file}")
    print(f"FO_F2: {fo2}")
    print(f"FO_F3 aleatoria: {info3['fo_aleatoria']}")
    print(f"FO_F3 Tabú: {fo3}")
    print(f"Detalle F3: {detalle3}")
    print(f"Iteraciones: {info3['iteraciones']} | Tiempo: {info3['tiempo_seg']} s")

if __name__ == "__main__":
    main()
