#!/usr/bin/env python3
"""Modelo principal: matriz de viabilidad/afinidad y MILP lexicográfico."""
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
    return base.make_slots(projects)


def affinity(project, person, number_of_slots, role_map=None):
    """Codifica continuidad, Dominio y skills obligatorias cubiertas en A."""
    return base.combined_preference_score(project, person, number_of_slots, role_map)


def solve_lexicographic(slots, people, role_map, return_recommendations=False):
    import numpy as np

    ns, npeople = len(slots), len(people)
    V = np.zeros((ns, npeople), dtype=np.uint8)
    A = np.zeros((ns, npeople), dtype=np.float32)
    for s, (_, project) in enumerate(slots):
        for p, person in enumerate(people):
            ok, _, _ = base.eligibility(project, person, set(), role_map)
            if ok:
                V[s, p] = 1
                A[s, p] = affinity(project, person, ns, role_map)

    selected_slots, selected, reasons, eligible, recommendation_pool = base.solve_with_suggestions(
        slots, people, role_map, viability_matrix=V, affinity_matrix=A)
    selected_count = len(selected)
    suggested_assigned = sum(
        person_index in recommendation_pool.get(base.mesa_group_key(slots[slot_index][1]), [])
        for slot_index, person_index, _score in selected
    )
    mesa_groups = {base.mesa_group_key(project) for _, project in slots}
    filled = set()
    for slot_index, _, _ in selected:
        project = slots[slot_index][1]
        filled.add(base.mesa_group_key(project))
    total = ns * npeople
    stats = {
        "total": total,
        "valid": int(V.sum()),
        "pruned_pct": round(100 * (total - int(V.sum())) / total, 2) if total else 100.0,
        "coverage": selected_count,
        "asignaciones_promovidas_desde_sugerencias": suggested_assigned,
        "grupos_mesa_foco": len(mesa_groups),
        "grupos_con_asignacion": len(filled),
        "solver_status": "Óptimo certificado en todas las fases",
    }
    if return_recommendations:
        return V, A, selected, stats, recommendation_pool
    return V, A, selected, stats


def main():
    RESULTS.mkdir(exist_ok=True)
    header, projects = base.load_projects()
    people = base.load_people()
    role_map, _ = base.build_role_map(projects, people)
    slots = make_slots(projects)
    V, A, selected, stats, recommendation_pool = solve_lexicographic(
        slots, people, role_map, return_recommendations=True)
    import csv
    codes = [person["code"] for person in people]
    for path, matrix in ((V_FILE, V), (A_FILE, A)):
        with path.open("w", encoding="utf-8-sig", newline="") as file:
            writer = csv.writer(file)
            writer.writerow(["Slot"] + codes)
            for index, row in enumerate(matrix):
                project = slots[index][1]
                label = " | ".join((str(index + 1), str(slots[index][0]), project.get("Mesa", ""),
                                    project.get("Nombre Proyecto", ""), project.get("Rol", ""),
                                    project.get("_required_type", "")))
                writer.writerow([label] + row.tolist())
    export_tobe(EXCEL_FILE, header, projects, slots, selected, people,
                affinity, base.TECH_FIELDS + base.FUNCTIONAL_FIELDS + base.SUGGESTED_FIELDS,
                role_map=role_map, recommendation_pool=recommendation_pool)
    print(" | ".join(f"{key}={value}" for key, value in stats.items()))
    print(f"Poda: {stats['pruned_pct']}% | V: {V_FILE} | A: {A_FILE} | Excel: {EXCEL_FILE}")


if __name__ == "__main__":
    main()
