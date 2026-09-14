# -*- coding: utf-8 -*-
"""Agrega Docentes, Personal Administrativo y Estudiantes 2025 (SNIES-MEN)
en un JSON compacto agrupado por IES (institucion) y semestre."""
import openpyxl, json
from collections import defaultdict

BASE = "C:/Users/ASUS vivobook/Downloads/educacion/"

def load(fname):
    wb = openpyxl.load_workbook(BASE + fname, read_only=True, data_only=True)
    ws = wb["1."]
    rows = ws.iter_rows(min_row=7, values_only=True)
    # excluir filas de notas al pie (sector None) y vacias
    data = [r for r in rows if r[0] is not None and r[5] is not None]
    wb.close()
    return data

def num(v):
    if v is None:
        return 0
    try:
        return int(str(v).strip())
    except Exception:
        return 0

def norm_acred(v):
    if v is None:
        return "Sin información"
    s = str(v).strip().upper()
    return "Si" if s in ("S", "SI") else ("No" if s in ("N", "NO") else s.title())

# modalidad: consolidar variantes hibridas/mixtas/duales
def norm_modalidad(m):
    if m is None:
        return "Sin información"
    s = str(m).strip()
    if "Dual" in s:
        return "Dual"
    if "Híbrida" in s or "Hibrida" in s:
        return "Híbrida"
    if s in ("Presencial", "Virtual", "A distancia"):
        return s
    if s == "Presencial-Virtual":
        return "Híbrida"
    # combinaciones tipo "Virtual-A distancia", "Presencial-A distancia", etc.
    return "Mixta"

insts = defaultdict(lambda: {
    "codigo": None, "nombre": None, "sector": None, "caracter": None,
    "tipo": None, "departamento": None, "municipio": None, "acreditada": None,
    "sedes": set(), "tieneSeccionales": False,
    "sem1": {"docentes": {"total": 0, "sexo": {}, "formacion": {}, "dedicacion": {}, "contrato": {}},
             "admin": {"total": 0, "categoria": {}},
             "estudiantes": {"total": 0, "nivelAcademico": {}, "formacion": {}, "modalidad": {},
                             "area": {}, "sexo": {}, "programaAcreditado": {}, "deptoOferta": {}},
             "graduados": {"total": 0, "nivelAcademico": {}, "formacion": {}, "modalidad": {},
                           "area": {}, "sexo": {}, "programaAcreditado": {}, "deptoOferta": {}}},
    "sem2": {"docentes": {"total": 0, "sexo": {}, "formacion": {}, "dedicacion": {}, "contrato": {}},
             "admin": {"total": 0, "categoria": {}},
             "estudiantes": {"total": 0, "nivelAcademico": {}, "formacion": {}, "modalidad": {},
                             "area": {}, "sexo": {}, "programaAcreditado": {}, "deptoOferta": {}},
             "graduados": {"total": 0, "nivelAcademico": {}, "formacion": {}, "modalidad": {},
                           "area": {}, "sexo": {}, "programaAcreditado": {}, "deptoOferta": {}}},
})

def ensure_meta(inst, codigo_sede, codigo_padre, nombre, tipo, sector, caracter, depto, muni, acred):
    if inst["codigo"] is None:
        inst["codigo"] = str(codigo_padre)
        inst["nombre"] = nombre
        inst["sector"] = sector
        inst["caracter"] = caracter
        inst["tipo"] = tipo
        inst["acreditada"] = norm_acred(acred)
        inst["departamento"] = depto
        inst["municipio"] = muni
    else:
        # la sede principal define depto/municipio de domicilio
        if tipo == "Principal":
            inst["tipo"] = "Principal"
            inst["departamento"] = depto
            inst["municipio"] = muni
    if inst["nombre"] is None and nombre is not None:
        inst["nombre"] = nombre
    inst["sedes"].add(str(codigo_sede))
    if tipo == "Seccional":
        inst["tieneSeccionales"] = True

def bump(d, k, v):
    d[k] = d.get(k, 0) + v

# ---------------- DOCENTES ----------------
print("Leyendo Docentes...")
for r in load("Docentes.xlsx"):
    padre, nombre, tipo = r[1], r[2], r[3]
    sector, caracter = r[5], r[7]
    depto, muni, acred = r[9], r[11], r[12]
    sexo, formacion = r[14], r[16]
    dedicacion, contrato = r[18], r[20]
    sem, cnt = r[22], num(r[23])
    sem = int(sem) if sem is not None else 1
    key = str(padre)
    inst = insts[key]
    ensure_meta(inst, r[0], padre, nombre, tipo, sector, caracter, depto, muni, acred)
    blk = inst["sem1"] if sem == 1 else inst["sem2"]
    d = blk["docentes"]
    d["total"] += cnt
    if sexo: bump(d["sexo"], str(sexo).strip(), cnt)
    if formacion: bump(d["formacion"], str(formacion).strip(), cnt)
    if dedicacion: bump(d["dedicacion"], str(dedicacion).strip(), cnt)
    if contrato: bump(d["contrato"], str(contrato).strip(), cnt)

# ---------------- PERSONAL ADMINISTRATIVO ----------------
print("Leyendo Personal Administrativo...")
for r in load("personal administrativo.xlsx"):
    padre, nombre, tipo = r[1], r[2], r[3]
    sector, caracter = r[5], r[7]
    depto, muni, acred = r[9], r[11], r[12]
    sem = str(r[14]).strip()
    sem = int(sem) if sem in ("1", "2") else 1
    key = str(padre)
    inst = insts[key]
    ensure_meta(inst, r[0], padre, nombre, tipo, sector, caracter, depto, muni, acred)
    blk = inst["sem1"] if sem == 1 else inst["sem2"]
    a = blk["admin"]
    a["total"] += num(r[19])
    for k, idx in (("Auxiliar", 15), ("Técnico", 16), ("Profesional", 17), ("Directivo", 18)):
        bump(a["categoria"], k, num(r[idx]))

# ---------------- ESTUDIANTES ----------------
print("Leyendo Estudiantes...")
for r in load("Estudiantes 2025.xlsx"):
    padre, nombre, tipo = r[1], r[2], r[3]
    sector, caracter = r[5], r[7]
    depto, muni, acred = r[9], r[11], r[12]
    nivel_acad, formacion = r[17], r[19]
    modalidad, area = norm_modalidad(r[21]), r[23]
    sexo, prog_acred = r[37], r[15]
    depto_oferta = r[33]
    sem, cnt = r[39], num(r[40])
    sem = int(sem) if sem is not None else 1
    key = str(padre)
    inst = insts[key]
    ensure_meta(inst, r[0], padre, nombre, tipo, sector, caracter, depto, muni, acred)
    blk = inst["sem1"] if sem == 1 else inst["sem2"]
    e = blk["estudiantes"]
    e["total"] += cnt
    if nivel_acad: bump(e["nivelAcademico"], str(nivel_acad).strip(), cnt)
    if formacion: bump(e["formacion"], str(formacion).strip(), cnt)
    bump(e["modalidad"], modalidad, cnt)
    if area: bump(e["area"], str(area).strip(), cnt)
    if sexo: bump(e["sexo"], str(sexo).strip(), cnt)
    if prog_acred: bump(e["programaAcreditado"], str(prog_acred).strip(), cnt)
    if depto_oferta: bump(e["deptoOferta"], str(depto_oferta).strip(), cnt)

# ---------------- GRADUADOS ----------------
print("Leyendo Graduados...")
for r in load("estudianrtes graduados.xlsx"):
    padre, nombre, tipo = r[1], r[2], r[3]
    sector, caracter = r[5], r[7]
    depto, muni, acred = r[9], r[11], r[12]
    nivel_acad, formacion = r[17], r[19]
    modalidad, area = norm_modalidad(r[21]), r[23]
    sexo, prog_acred = r[37], r[15]
    depto_oferta = r[33]
    sem, cnt = r[39], num(r[40])
    sem = int(sem) if sem is not None else 1
    key = str(padre)
    inst = insts[key]
    ensure_meta(inst, r[0], padre, nombre, tipo, sector, caracter, depto, muni, acred)
    blk = inst["sem1"] if sem == 1 else inst["sem2"]
    g = blk["graduados"]
    g["total"] += cnt
    if nivel_acad: bump(g["nivelAcademico"], str(nivel_acad).strip(), cnt)
    if formacion: bump(g["formacion"], str(formacion).strip(), cnt)
    bump(g["modalidad"], modalidad, cnt)
    if area: bump(g["area"], str(area).strip(), cnt)
    if sexo: bump(g["sexo"], str(sexo).strip(), cnt)
    if prog_acred: bump(g["programaAcreditado"], str(prog_acred).strip(), cnt)
    if depto_oferta: bump(g["deptoOferta"], str(depto_oferta).strip(), cnt)

print("Total instituciones:", len(insts))

# convertir set de sedes a conteo
for i in insts.values():
    i["sedes"] = len(i["sedes"])

# ordenar instituciones por nombre
lista = sorted(insts.values(), key=lambda x: (x["nombre"] or "").lower())

# metadata de dimensiones (para filtros)
def dims(field):
    s = set()
    for i in lista:
        v = i[field]
        if v: s.add(v)
    return sorted(s)

out = {
    "meta": {
        "titulo": "Educación Superior - Colombia 2025",
        "fuente": "SNIES - Ministerio de Educación Nacional",
        "corte": "Mayo 31 de 2026",
        "anio": 2025,
        "desarrolladoPor": "John A. Echeverry",
        "para": "Mauricio",
    },
    "dimensiones": {
        "sector": dims("sector"),
        "caracter": dims("caracter"),
        "departamento": dims("departamento"),
        "municipio": dims("municipio"),
        "acreditada": dims("acreditada"),
    },
    "instituciones": lista,
}

with open(BASE + "data.json", "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, separators=(",", ":"))

print("data.json escrito OK")
# resumen rapido
d_total = sum(i["sem1"]["docentes"]["total"] + i["sem2"]["docentes"]["total"] for i in lista)
a_total = sum(i["sem1"]["admin"]["total"] + i["sem2"]["admin"]["total"] for i in lista)
e_total = sum(i["sem1"]["estudiantes"]["total"] + i["sem2"]["estudiantes"]["total"] for i in lista)
g_total = sum(i["sem1"]["graduados"]["total"] + i["sem2"]["graduados"]["total"] for i in lista)
print("Docentes:", d_total, "| Admin:", a_total, "| Estudiantes:", e_total, "| Graduados:", g_total)
