"""
Genera los JSON del portal Pinulito para un mes:
  - <mes>.json                (colaboradores)
  - supervisores-<mes>.json   (supervisores: bono + equipo)

El personal se asigna a sucursal/supervisor EXACTAMENTE como sale en la hoja
"Resumen Final" del Reporte (una fila por colaborador-sucursal).

Uso:
  python3 build_portal.py <Reporte.xlsx> <Bono_Sup.xlsx> <AAAA-MM> "<Mes Año>" <supervisores-mes-anterior.json> <salida_dir>
"""
import sys, json, calendar
from datetime import datetime, timedelta
from collections import defaultdict
import openpyxl

GC1_TIERS = [(7501, 100), (6501, 80), (6001, 60), (5501, 40), (5001, 20), (4500, 10)]
BONO_JEFE_UMBRAL, BONO_JEFE_MONTO = 6000, 30
GC2_VENTA_MIN, GC2_DIAS_MIN = 5225, 25


def tier_actual(v):
    for u, m in GC1_TIERS:
        if v >= u:
            return m
    return 0


def siguiente_tier(v):
    for u, m in sorted(GC1_TIERS):
        if v < u:
            return {"monto": m, "venta_faltante": round(u - v, 2), "umbral": u}
    return None


def hhmm(h):
    if h is None:
        return "0:00"
    mins = int(h.total_seconds() // 60) if isinstance(h, timedelta) else round(float(h) * 60)
    return f"{mins // 60}:{mins % 60:02d}"


def es_dato(emp_id):
    return emp_id is not None and str(emp_id).strip().upper() != "TOTAL"


def build_empleados(wb, mes_texto):
    # --- Detalle (por día) ---
    det = [r for r in wb["Detalle"].iter_rows(min_row=2, values_only=True) if es_dato(r[2])]
    fechas = sorted({r[1] for r in det if r[1]})
    dias_trans = len(fechas)
    dias_mes = calendar.monthrange(fechas[-1].year, fechas[-1].month)[1]
    det_emp = defaultdict(lambda: defaultdict(list))
    for r in det:
        suc, fecha, eid, horas, calif, solo1, venta_asig = r[0], r[1], str(r[2]), r[4], r[5], r[6], r[10]
        det_emp[eid][suc].append({
            "fecha": fecha.strftime("%d/%m") if fecha else "",
            "horas": hhmm(horas), "califica": calif == "Sí",
            "solo_una_marcacion": solo1 == "Sí", "venta_asignada": round(venta_asig or 0, 2)})

    # --- Resumen por Empleado (venta consolidada, aplica GC2) ---
    res = {}
    for r in wb["Resumen por Empleado"].iter_rows(min_row=2, values_only=True):
        if not es_dato(r[0]):
            continue
        eid = str(r[0])
        x = res.setdefault(eid, {"venta": 0, "aplica": False, "cargo": r[8]})
        x["venta"] = round(r[10] or 0, 2) or x["venta"]
        x["aplica"] = x["aplica"] or bool(r[12])

    # --- Resumen Final (fuente de sucursal / supervisor / montos) ---
    filas = []
    for r in wb["Resumen Final"].iter_rows(min_row=2, values_only=True):
        if not es_dato(r[0]) or not r[2]:
            continue
        filas.append({"id": str(r[0]), "nombre": r[1], "sucursal": r[2], "supervisor": r[3],
                      "plaza": r[4], "dias_trab": r[5] or 0, "dias_calif": r[6] or 0,
                      "gc1": r[9] or 0, "bono": r[10] or 0, "gc2": r[11] or 0, "total": r[12] or 0})
    por_emp = defaultdict(list)
    for f in filas:
        por_emp[f["id"]].append(f)

    def detalle_suc(eid, suc):
        dias = sorted(det_emp.get(eid, {}).get(suc, []), key=lambda d: d["fecha"])
        return {"sucursal": suc, "dias_trabajados": len(dias),
                "dias_califican": sum(1 for d in dias if d["califica"]),
                "venta_asignada": round(sum(d["venta_asignada"] for d in dias), 2), "dias": dias}

    def proyeccion(venta, dias_calif, es_jefe, aplica):
        ritmo = round(venta / dias_trans, 2) if dias_trans else 0
        proy = round(ritmo * dias_mes, 2)
        return {"dias_transcurridos": dias_trans, "dias_mes": dias_mes, "ritmo_diario": ritmo,
                "venta_proyectada": proy, "tier_actual": tier_actual(venta),
                "tier_proyectado": tier_actual(proy), "siguiente_tier": siguiente_tier(proy),
                "es_jefe": es_jefe, "bono_jefe_umbral": BONO_JEFE_UMBRAL,
                "bono_jefe_monto": BONO_JEFE_MONTO, "aplica_gc2": aplica,
                "gc2_dias_min": GC2_DIAS_MIN, "gc2_venta_min": GC2_VENTA_MIN,
                "gc2_dias_faltan": max(0, GC2_DIAS_MIN - dias_calif),
                "gc2_venta_falta": max(0, round(GC2_VENTA_MIN - venta, 2))}

    empleados, filas_equipo = {}, []
    for eid, fs in por_emp.items():
        # fila principal: la que lleva el premio; si ninguna, la de más días trabajados
        principal = max(fs, key=lambda f: (f["total"], f["dias_trab"]))
        info = res.get(eid, {"venta": 0, "aplica": False, "cargo": ""})
        sucs = sorted((detalle_suc(eid, f["sucursal"]) for f in fs), key=lambda s: -s["venta_asignada"])
        dias_trab = len({d["fecha"] for s in det_emp.get(eid, {}).values() for d in s})
        dias_cal = len({d["fecha"] for s in det_emp.get(eid, {}).values() for d in s if d["califica"]})
        es_jefe = info["cargo"] == "JEFE DE SUCURSAL"
        gc1 = sum(f["gc1"] for f in fs); bono = sum(f["bono"] for f in fs); gc2 = sum(f["gc2"] for f in fs)
        base = {"nombre": principal["nombre"], "cargo": info["cargo"], "plaza": principal["plaza"],
                "sucursal_principal": principal["sucursal"], "supervisor": principal["supervisor"],
                "dias_trabajados": dias_trab, "dias_califican": dias_cal,
                "venta_consolidada": info["venta"], "gc1": gc1, "bono_jefe": bono, "gc2": gc2,
                "total": round(gc1 + bono + gc2, 2), "sucursales": sucs,
                "proyeccion": proyeccion(info["venta"], dias_cal, es_jefe, info["aplica"])}
        empleados[eid] = base
        # una entrada de equipo por fila del Resumen Final (sucursal tal cual la hoja)
        for f in fs:
            e = dict(base)
            e.update({"id": eid, "sucursal_principal": f["sucursal"], "supervisor": f["supervisor"],
                      "plaza": f["plaza"], "dias_trabajados": f["dias_trab"],
                      "dias_califican": f["dias_calif"], "gc1": f["gc1"], "bono_jefe": f["bono"],
                      "gc2": f["gc2"], "total": round(f["gc1"] + f["bono"] + f["gc2"], 2),
                      "sucursales": [detalle_suc(eid, f["sucursal"])]})
            filas_equipo.append(e)

    data = {"meta": {"mes": mes_texto, "dias_transcurridos": dias_trans, "dias_mes": dias_mes,
                     "actualizado": datetime.now().strftime("%d/%m/%Y")}, "empleados": empleados}
    return data, filas_equipo


def build_supervisores(wb_bono, mes_texto, filas_equipo, prev_sup, nuevos=None):
    nuevos = nuevos or {}
    cod, comp, activo = {}, {}, {}
    roll = {}
    for c, s in prev_sup["supervisores"].items():
        if s.get("supervisores_subordinados") is not None:
            roll[c] = s
        else:
            cod[s["nombre"]], comp[s["nombre"]], activo[s["nombre"]] = c, s["nombre_completo"], s.get("activo")
    for corto, (c, completo) in nuevos.items():
        cod[corto], comp[corto], activo[corto] = c, completo, True

    grupos, actual = {}, None
    for row in wb_bono["Comisiones"].iter_rows(min_row=3, values_only=True):
        (_i, _g, sup, suc, vta, meta, pct, _d, _q, _f, _vt, _pr, _c, _de, ing, cal_pct, cc, crit, inv, tot, *_r) = row
        if suc is None:
            continue
        if str(suc).strip().upper().startswith("TOTAL GENERAL"):
            break
        if isinstance(sup, str) and sup.strip():
            actual = sup.strip()
        if actual is None:
            continue
        grupos.setdefault(actual, []).append({
            "sucursal": suc, "venta": round(vta or 0, 2), "meta": round(meta or 0, 2),
            "pct_venta": round((pct or 0) * 100, 1), "ingresos": round(ing or 0, 2),
            "calidad_pct": round((cal_pct or 0) * 100, 1), "calidad_monto": round(cc or 0, 2),
            "inventario_criterio": (crit or "").strip() or "No Aplica",
            "inventarios": round(inv or 0, 2), "total": round(tot or 0, 2)})

    desglose = {}
    for r in wb_bono["Desglose comisiones"].iter_rows(min_row=2, values_only=True):
        if len(r) >= 12 and r[2] == "SUPERVISORES" and r[1]:
            desglose[r[1].strip()] = {"bono_ingresos": round(r[8] or 0, 2), "bono_calidad": round(r[9] or 0, 2),
                                      "bono_inventarios": round(r[10] or 0, 2), "bono_total": round(r[11] or 0, 2)}

    # Si el supervisor de la fila no es un supervisor conocido (p. ej. "Guatemala"),
    # se usa el supervisor que el archivo de bono asigna a esa sucursal.
    conocidos = set(comp.values())
    suc_a_sup = {s["sucursal"]: comp[corto] for corto, ss in grupos.items() if corto in comp for s in ss}
    equipo_por = defaultdict(list)
    for e in filas_equipo:
        sup_fila = e["supervisor"]
        if sup_fila not in conocidos and e["sucursal_principal"] in suc_a_sup:
            sup_fila = suc_a_sup[e["sucursal_principal"]]
            e["supervisor"] = sup_fila
        equipo_por[sup_fila].append(e)

    sups, ignorados = {}, []
    for corto, sucs in sorted(grupos.items()):
        if corto not in cod:
            ignorados.append(corto)
            continue
        m = desglose.get(comp[corto]) or {
            "bono_ingresos": round(sum(s["ingresos"] for s in sucs), 2),
            "bono_calidad": round(sum(s["calidad_monto"] for s in sucs), 2),
            "bono_inventarios": round(sum(s["inventarios"] for s in sucs), 2),
            "bono_total": round(sum(s["total"] for s in sucs), 2)}
        sups[cod[corto]] = {"nombre": corto, "nombre_completo": comp[corto], "equipo": equipo_por.get(comp[corto], []),
                            **m, "sucursales": sucs, "activo": activo.get(corto, True)}
    # supervisores sin fila de bono este mes pero con equipo (p. ej. inactivos)
    for corto, c in cod.items():
        if c not in sups and equipo_por.get(comp[corto]):
            sups[c] = {"nombre": corto, "nombre_completo": comp[corto], "equipo": equipo_por[comp[corto]],
                       "bono_ingresos": 0, "bono_calidad": 0, "bono_inventarios": 0, "bono_total": 0,
                       "sucursales": [], "activo": activo.get(corto, False)}

    base = dict(sups)
    for c, r in roll.items():
        nombres = r.get("_subs") or [x["nombre"] for x in r["supervisores_subordinados"]]
        subs = [s for s in sups.values() if s["nombre_completo"] in nombres]
        if len(nombres) > 1:  # Aquiles / Administrativo: todos los supervisores de sucursal
            subs = [s for s in base.values()]
        sups[c] = {"nombre": r["nombre"], "nombre_completo": r["nombre_completo"],
                   "equipo": [e for s in subs for e in s["equipo"]], "bono_ingresos": None, "bono_calidad": None,
                   "bono_inventarios": None, "bono_total": None, "sucursales": [],
                   "supervisores_subordinados": subs, "activo": r.get("activo", True)}
    return {"meta": {"mes": mes_texto}, "supervisores": sups}, ignorados


if __name__ == "__main__":
    reporte, bono, key, mes, prev, out = sys.argv[1:7]
    wb = openpyxl.load_workbook(reporte, data_only=True)
    emp, filas = build_empleados(wb, mes)
    json.dump(emp, open(f"{out}/{key}.json", "w", encoding="utf-8"), ensure_ascii=False)
    sup, ign = build_supervisores(openpyxl.load_workbook(bono, data_only=True), mes, filas,
                                  json.load(open(prev, encoding="utf-8")))
    json.dump(sup, open(f"{out}/supervisores-{key}.json", "w", encoding="utf-8"), ensure_ascii=False)
    print(len(emp["empleados"]), "empleados | ignorados:", ign)
