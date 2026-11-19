#!/usr/bin/env python3
"""Modelo con matrices V/A, poda de candidatos y MILP en dos fases."""
from collections import defaultdict
from pathlib import Path

import asignacion_milp as base
from salidas_asignacion import export_tobe

ROOT = Path(__file__).resolve().parent
INPUTS = ROOT / "RESULTADOS_LIMPIEZA"
RESULTS = ROOT / "RESULTADOS MODELOS"
EXCEL_FILE = RESULTS / "Asignacion_To-Be_Tensores_Lexicografico.xlsx"
V_FILE = RESULTS / "Viabilidad_V.csv"
A_FILE = RESULTS / "Afinidad_A.csv"


def make_slots(projects):
    slots = []
    for index, project in enumerate(projects):
        skill_count = len(base.required_skill_names(
            project, base.TECH_FIELDS + base.FUNCTIONAL_FIELDS + base.SUGGESTED_FIELDS))
        if not (project.get("Rol") or "").strip() or not (project.get("Seniority") or "").strip() or skill_count < base.MIN_SKILLS_PER_ROLE:
            continue
        for field, category in (("Internal", "interno"), ("External", "externo")):
            raw = (project.get(field) or "").strip()
            try:
                count = int(float(raw)) if raw else 0
            except ValueError:
                count = 0
            for number in range(1, max(0, count) + 1):
                demand = dict(project, _required_type=category, _slot_number=str(number))
                try:
                    source_row = int(project.get("Fila origen") or index + 2)
                except ValueError:
                    source_row = index + 2
                slots.append((source_row, demand))
    return slots


def affinity(project, person):
    required = base.required_skill_names(project, base.TECH_FIELDS + base.FUNCTIONAL_FIELDS)
    values = [person["skills"].get(skill, 0) for skill in required]
    skill_fit = sum(values) / (4 * len(values)) if values else 0.5
    return 0.8 * skill_fit + 0.2  # El seniority es filtro exacto en Rglas.md.


def solve_lexicographic(slots, people, role_map):
    import numpy as np
    from scipy.optimize import Bounds, LinearConstraint, milp
    from scipy.sparse import coo_matrix

    ns, nperson = len(slots), len(people)
    V = np.zeros((ns, nperson), dtype=np.uint8)
    A = np.zeros((ns, nperson), dtype=np.float32)
    for s, (_, project) in enumerate(slots):
        for p, person in enumerate(people):
            ok, _, _ = base.eligibility(project, person, set(), role_map)
            if ok:
                V[s, p] = 1
                A[s, p] = affinity(project, person)
    pairs = [(s, p) for s in range(ns) for p in range(nperson) if V[s, p]]
    if not pairs:
        return V, A, [], {"total": ns * nperson, "valid": 0, "pruned_pct": 100.0,
                          "coverage": 0, "closed": 0}

    # Una variable binaria por candidato viable y una por proyecto cerrado.
    project_slots = defaultdict(list)
    for s, (_, project) in enumerate(slots):
        key = (base.key(project.get("SDATOOL")), base.norm(project.get("Nombre Proyecto")),
               base.norm(project.get("Mesa")))
        project_slots[key].append(s)
    project_keys = list(project_slots)
    z0 = len(pairs)
    repeat_people = sorted({person for _, person in pairs})
    repeat0 = z0 + len(project_keys)
    repeat_index = {person: repeat0 + i for i, person in enumerate(repeat_people)}
    rows, lower, upper = [], [], []
    def constraint(coeff, lo=-np.inf, hi=np.inf):
        rows.append(coeff); lower.append(lo); upper.append(hi)
    by_slot, by_person, by_mesa = defaultdict(list), defaultdict(list), defaultdict(list)
    for j, (s, p) in enumerate(pairs):
        by_slot[s].append(j); by_person[p].append(j)
        mesa = base.norm(slots[s][1].get("Mesa"))
        if mesa:
            by_mesa[mesa].append(j)
    for items in by_slot.values(): constraint({j: 1 for j in items}, hi=1)
    for person, items in by_person.items():
        coeff = {j: 1 for j in items}
        coeff[repeat_index[person]] = -1
        constraint(coeff, hi=base.MAX_ASSIGNMENTS_PER_PERSON)
    for items in by_mesa.values(): constraint({j: 1 for j in items}, hi=base.MAX_PEOPLE_PER_MESA)
    for k, key in enumerate(project_keys):
        z = z0 + k
        # z=1 solo si se cubrieron todos los slots del proyecto.
        for s in project_slots[key]:
            coeff = {j: -1 for j in by_slot.get(s, [])}
            coeff[z] = 1
            constraint(coeff, hi=0)
        # Evita z=1 cuando no existen plazas o no se cubre proyecto completo.
        constraint({z: 1}, hi=1)
    nvars = repeat0 + len(repeat_people)
    rr, cc, vv = [], [], []
    for ri, row in enumerate(rows):
        for col, val in row.items(): rr.append(ri); cc.append(col); vv.append(val)
    matrix = coo_matrix((vv, (rr, cc)), shape=(len(rows), nvars)).tocsr()
    bounds = Bounds(np.zeros(nvars), np.ones(nvars))
    integ = np.ones(nvars)
    def run(obj, extra=()):
        mats = [matrix]; lows = [np.array(lower)]; highs = [np.array(upper)]
        for coeff, lo, hi in extra:
            ar, ac, av = [], [], []
            for col, value in coeff.items(): ar.append(0); ac.append(col); av.append(value)
            from scipy.sparse import csr_matrix
            mats.append(csr_matrix((av, (ar, ac)), shape=(1, nvars)))
            lows.append(np.array([lo])); highs.append(np.array([hi]))
        from scipy.sparse import vstack
        result = milp(obj, integrality=integ, bounds=bounds,
                      constraints=LinearConstraint(vstack(mats), np.concatenate(lows), np.concatenate(highs)),
                      options={"time_limit": 300})
        if result.x is None:
            raise RuntimeError(f"MILP sin solución: {result.message}")
        return result
    # Fase 1: maximizar cobertura total y, fijada esta, maximizar iniciativas cerradas.
    coverage_obj = np.zeros(nvars); coverage_obj[:len(pairs)] = -1_000_000
    coverage_obj[repeat0:] = 1
    phase1a = run(coverage_obj)
    coverage = int(round(sum(phase1a.x[:len(pairs)])))
    repeat_count = int(round(sum(phase1a.x[repeat0:])))
    # Preserva el estado para reportar cualquier parada no óptima del solver.
    closed_obj = np.zeros(nvars); closed_obj[z0:repeat0] = -1
    phase1b = run(closed_obj, [({j: 1 for j in range(len(pairs))}, coverage, coverage),
                               ({repeat0 + i: 1 for i in range(len(repeat_people))}, repeat_count, repeat_count)])
    closed = int(round(sum(phase1b.x[z0:repeat0])))
    # Fase 2: maximizar afinidad, con cobertura y cierres congelados; externalidad es proxy de costo.
    talent_obj = np.zeros(nvars)
    for j, (s, p) in enumerate(pairs):
        talent_obj[j] = -(float(A[s, p]) * 10000 - (1 if people[p]["resource_type"] == "externo" else 0))
    final = run(talent_obj, [({j: 1 for j in range(len(pairs))}, coverage, coverage),
                             ({z0 + k: 1 for k in range(len(project_keys))}, closed, closed),
                             ({repeat0 + i: 1 for i in range(len(repeat_people))}, repeat_count, repeat_count)])
    selected = [(s, p, float(A[s, p])) for j, (s, p) in enumerate(pairs) if final.x[j] > 0.5]
    total = ns * nperson
    stats = {"total": total, "valid": len(pairs),
             "pruned_pct": round(100 * (total - len(pairs)) / total, 2) if total else 100.0,
             "coverage": coverage, "closed": closed, "projects": len(project_keys),
             "capacity_overflow_people": repeat_count,
             "solver_status": str(phase1a.message)}
    return V, A, selected, stats


def main():
    RESULTS.mkdir(exist_ok=True)
    header, projects = base.load_projects()
    people = base.load_people()
    role_map, _ = base.build_role_map(projects, people)
    slots = make_slots(projects)
    V, A, selected, stats = solve_lexicographic(slots, people, role_map)
    import csv
    codes = [p["code"] for p in people]
    for path, matrix in ((V_FILE, V), (A_FILE, A)):
        with path.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f); writer.writerow(["Slot"] + codes)
            for i, row in enumerate(matrix):
                project = slots[i][1]
                label = " | ".join((str(i + 1), str(slots[i][0]), project.get("Mesa", ""),
                                    project.get("Nombre Proyecto", ""), project.get("Rol", ""),
                                    project.get("_required_type", "")))
                writer.writerow([label] + row.tolist())
    export_tobe(EXCEL_FILE, header, projects, slots, selected, people,
                affinity, base.TECH_FIELDS + base.FUNCTIONAL_FIELDS + base.SUGGESTED_FIELDS)
    print(" | ".join(f"{k}={v}" for k, v in stats.items()))
    print(f"Poda: {stats['pruned_pct']}% | V: {V_FILE} | A: {A_FILE} | Excel: {EXCEL_FILE}")
    if stats["pruned_pct"] < 80:
        print("AVISO: los datos no alcanzaron el objetivo de poda >80%; se reporta la poda observada sin alterar reglas.")


if __name__ == "__main__":
    main()
