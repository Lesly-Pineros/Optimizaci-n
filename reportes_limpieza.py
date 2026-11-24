#!/usr/bin/env python3
"""Genera reportes auxiliares a partir de las fuentes y salidas de limpieza.

Ejecutar después de `limpieza.py` cuando se necesiten los reportes de cruce y
de skills de iniciativas sin mapeo.
"""

from limpieza import (
    RESULTS,
    SKILLS_FILE,
    SKILLS_HEADER_ROW,
    STAFFING_FILE,
    code_key,
    find_column,
    normalize,
    read_csv,
    skill_match_key,
    skills_catalog,
    write_review_workbook,
)

UNMATCHED_OUTPUT = RESULTS / "Codigos sin cruce SkillsCOE-Staffing.xlsx"
UNMAPPED_SKILLS_OUTPUT = RESULTS / "Skills de iniciativas sin mapeo.xlsx"


def prepare_unmatched_report():
    """Genera listas de códigos presentes solo en una de las dos fuentes."""
    skills_header, people = read_csv(SKILLS_FILE, SKILLS_HEADER_ROW - 1)
    staffing_header, staffing_rows = read_csv(STAFFING_FILE)
    skills_code_column = find_column(skills_header, "Código")
    skills_name_column = find_column(skills_header, "Nombre")
    skills_role_column = find_column(skills_header, "Rol Principal")
    skills_type_column = find_column(skills_header, "Tipo de Recurso")
    staffing_code_column = find_column(staffing_header, "codigo")

    skills_by_code = {}
    for person in people:
        code = code_key(person.get(skills_code_column))
        if code:
            skills_by_code.setdefault(code, person)
    staffing_by_person = {}
    for person in staffing_rows:
        code = code_key(person.get(staffing_code_column))
        if code:
            staffing_by_person.setdefault(code, person)

    skills_only = [
        [row.get(skills_code_column, ""), row.get(skills_name_column, ""),
         row.get(skills_role_column, ""), row.get(skills_type_column, ""), "BACKLOG"]
        for code, row in sorted(skills_by_code.items()) if code not in staffing_by_person
    ]
    staffing_only = [
        [row.get("codigo", ""), row.get("interno", ""), row.get("% dedicacion", ""),
         row.get("rol", ""), row.get("dominio", ""), row.get("direccion", ""),
         row.get("SDATOOL - Proyecto", ""), row.get("Fecha Inicio", ""),
         row.get("Fecha Fin Actual", ""), row.get("Mesa actual", "")]
        for code, row in sorted(staffing_by_person.items()) if code not in skills_by_code
    ]
    write_review_workbook(UNMATCHED_OUTPUT, [
        ("Skills sin Staffing", ["Código", "Nombre", "Rol Principal", "Tipo de Recurso",
                                 "Estado Staffing"], skills_only),
        ("Staffing sin Skills", ["codigo", "interno", "% dedicacion", "rol", "dominio", "direccion",
                                  "SDATOOL - Proyecto", "Fecha Inicio", "Fecha Fin Actual", "Mesa actual"],
         staffing_only),
    ])
    print(f"Sin cruce: {len(skills_only)} códigos en SkillsCOE sin Staffing; "
          f"{len(staffing_only)} en Staffing sin SkillsCOE.")
    print(f"Excel de revisión: {UNMATCHED_OUTPUT}")


def prepare_unmapped_skills_report():
    """Lista skills requeridas sin columna equivalente en la matriz SkillsCOE."""
    catalog_keys = {skill_match_key(name) for name in skills_catalog()}
    rows = []
    for path in sorted(RESULTS.glob("*-iniciativas con skills.csv")):
        header, initiatives = read_csv(path)
        index = {normalize(name): name for name in header}
        skill_fields = [name for name in header if normalize(name).startswith((
            "habilidades tecnicas obligatorias", "habilidades funcionales obligatorias"))]
        for initiative in initiatives:
            for field in skill_fields:
                requested = (initiative.get(field) or "").strip()
                if requested and skill_match_key(requested) not in catalog_keys:
                    rows.append([
                        initiative.get("Dominio", ""),
                        initiative.get(index.get(normalize("SDATOOL"), ""), ""),
                        initiative.get(index.get(normalize("Nombre Proyecto"), ""), ""),
                        initiative.get(index.get(normalize("Rol"), ""), ""),
                        "Técnica" if "tecnicas obligatorias" in normalize(field) else "Funcional",
                        requested,
                        field,
                    ])

    write_review_workbook(UNMAPPED_SKILLS_OUTPUT, [
        ("Skills sin mapeo",
         ["Dominio", "SDATOOL", "Nombre Proyecto", "Rol", "Tipo de skill",
          "Skill solicitada", "Campo origen"], rows),
    ])
    print(f"Skills solicitadas sin columna en SkillsCOE: {len(rows)}; "
          f"Excel de revisión: {UNMAPPED_SKILLS_OUTPUT}")


def main():
    RESULTS.mkdir(exist_ok=True)
    prepare_unmapped_skills_report()
    prepare_unmatched_report()


if __name__ == "__main__":
    main()
