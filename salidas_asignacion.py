"""Exporta un To-Be XLSX sin dependencias adicionales de escritura."""
from pathlib import Path
import re
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile


def _col(n):
    result = ""
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def write_xlsx(path, headers, rows, sheet_name="Asignaciones"):
    """Escribe un libro de una hoja, compatible con Excel/LibreOffice."""
    write_xlsx_sheets(path, [(sheet_name, headers, rows)])


def write_xlsx_sheets(path, sheets):
    """Escribe un libro con varias hojas y celdas de texto inline."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(path, "w", ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?>'
                    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                    '<Default Extension="xml" ContentType="application/xml"/>'
                    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                    + ''.join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                              for i in range(1, len(sheets) + 1)) + '</Types>')
        zf.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                    '</Relationships>')
        workbook_sheets = ''.join(
            f'<sheet name="{escape(name[:31])}" sheetId="{i}" r:id="rId{i}"/>'
            for i, (name, _, _) in enumerate(sheets, 1))
        zf.writestr("xl/workbook.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                    f'<sheets>{workbook_sheets}</sheets></workbook>')
        rels = ''.join(
            f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
            for i in range(1, len(sheets) + 1))
        zf.writestr("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    + rels + '</Relationships>')
        for sheet_index, (sheet_name, headers, rows) in enumerate(sheets, 1):
            all_rows = [headers] + [[row.get(h, "") for h in headers] for row in rows]
            xml_rows = []
            for ri, row in enumerate(all_rows, 1):
                cells = []
                for ci, value in enumerate(row, 1):
                    ref = f"{_col(ci)}{ri}"
                    value = "" if value is None else str(value)
                    cells.append(f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{escape(value)}</t></is></c>')
                xml_rows.append(f'<row r="{ri}">{"".join(cells)}</row>')
            sheet_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                         '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                         f'<sheetData>{"".join(xml_rows)}</sheetData></worksheet>')
            zf.writestr(f"xl/worksheets/sheet{sheet_index}.xml", sheet_xml)


def export_tobe(path, source_header, projects, slots, selected, people, skill_fit, skill_fields,
                role_map=None, recommendation_pool=None):
    """Exporta asignaciones y candidatos opcionales en pestañas separadas."""
    from collections import defaultdict
    import asignacion_milp as base
    role_map = role_map or {}
    recommendation_pool = recommendation_pool or {}
    assignments = {slot: (person, score) for slot, person, score in selected}
    source_fields = list(source_header[:18])
    extra_fields = [
        "Dominio requerido", "Fila origen", "N.º plaza", "Tipo requerido", "Estado asignación",
        "Motivo sin asignar", "Persona asignada", "Código persona", "Origen de asignación", "Rol persona",
        "Seniority persona", "Ajuste seniority", "Tipo persona", "Ajuste tipo de recurso", "Dominio persona",
        "Puntaje preferencias", "Penalización por skills",
        "Porcentaje coincidencia skills obligatorias",
        "Puntaje skills obligatorias",
        "Comentario skills obligatorias",
        "Skills obligatorias coincidentes", "Skills con match", "Skills faltantes",
        "Comentario asignación Staffing",
        "SDATOOL - Proyecto Staffing actual",
        "Fecha Inicio Staffing actual", "Fecha Fin Actual Staffing", 
        "% Dedicación Staffing actual",
    ]
    assignment_headers = source_fields + extra_fields

    def skill_report(project, person):
        required = []
        seen = set()
        for field in base.TECH_FIELDS + base.FUNCTIONAL_FIELDS:
            raw_name = (project.get(field) or "").strip()
            normalized = base.norm(raw_name)
            if raw_name and normalized not in seen:
                required.append((raw_name, normalized))
                seen.add(normalized)
        matched, missing = [], []
        for raw_name, normalized in required:
            value = person["skills"].get(normalized, 0)
            if value > 0:
                matched.append(f"{raw_name} (nivel {value})")
            else:
                missing.append(f"{raw_name} (sin skill registrada)")
        percent = f"{100 * len(matched) / len(required):.1f}%" if required else "N/A"
        ratio = f"{len(matched)}/{len(required)}" if required else "0/0"
        return (percent, ratio, "; ".join(matched), "; ".join(missing),
                len(matched), base.skill_observation(project, person, role_map))

    assigned_by_mesa = defaultdict(set)
    assigned_types = defaultdict(set)
    eligible_types_by_mesa = defaultdict(set)
    for _slot, (_source_row, project) in enumerate(slots):
        mesa = base.mesa_group_key(project)
        for candidate in people:
            if base.eligibility(project, candidate, set(), role_map)[0]:
                eligible_types_by_mesa[mesa].add(candidate["resource_type"])
    for slot, person_index, _ in selected:
        project = slots[slot][1]
        mesa = base.mesa_group_key(project)
        assigned_by_mesa[mesa].add(person_index)
        assigned_types[mesa].add(people[person_index]["resource_type"])

    assignment_rows = []
    for slot_index, (source_row, project) in enumerate(slots):
        person_index, score = assignments.get(slot_index, (None, ""))
        if person_index is None:
            mesa_key = base.mesa_group_key(project)
            candidates = [i for i, candidate in enumerate(people)
                          if base.eligibility(project, candidate, set(), role_map)[0]]
            if not candidates:
                reason = base.diagnose_no_candidate(project, people, set(), role_map)
            elif (eligible_types_by_mesa[mesa_key] == {"interno", "externo"}
                  and assigned_types[mesa_key] != {"interno", "externo"}
                  and assigned_by_mesa[mesa_key]):
                reason = "La Mesa no cumple la mezcla interna/externa requerida"
            elif (eligible_types_by_mesa[mesa_key] == {"interno", "externo"}
                  and not assigned_by_mesa[mesa_key]):
                reason = "Mesa sin asignar: no fue posible asegurar presencia interna y externa"
            else:
                reason = "Capacidad o solapamiento de proyectos"
            person = None
            percent, ratio, matched, missing, skill_score, skill_note = "", "", "", "", "", ""
            status = "Vacante"
        else:
            person = people[person_index]
            reason = ""
            percent, ratio, matched, missing, skill_score, skill_note = skill_report(project, person)
            status = "Asignada"
        origin = ""
        if person_index is not None:
            origin = ("Candidato sugerido adicional"
                      if person_index in recommendation_pool.get(base.mesa_group_key(project), [])
                      else "Asignación base")

        row = {field: project.get(field, "") for field in source_fields}
        row.update({
            "Dominio requerido": project.get("Dominio", ""),
            "Fila origen": source_row,
            "N.º plaza": project.get("_slot_number", ""),
            "Tipo requerido": project.get("_required_type", ""),
            "Estado asignación": status,
            "Motivo sin asignar": reason,
            "Persona asignada": person["name"] if person else "",
            "Código persona": person["code"] if person else "",
            "Origen de asignación": origin,
            "Rol persona": person["role_raw"] if person else "",
            "Seniority persona": person["seniority_raw"] if person else "",
            "Ajuste seniority": ({2: "Exacto", 1: "Un nivel por encima", 0: "Un nivel por debajo"}[
                base.seniority_fit(project.get("Seniority"), person["seniority"])] if person else ""),
            "Tipo persona": person["resource_type"] if person else "",
            "Ajuste tipo de recurso": (
                "Exacto" if person and person["resource_type"] == project.get("_required_type")
                else (f"Sustitución: {person['resource_type']} cubre plaza {project.get('_required_type')}"
                      if person else "")),
            "Dominio persona": " | ".join(sorted(person.get("domains", set()))) if person else "",
            "Puntaje preferencias": score if person else "",
            "Penalización por skills": (
                -base.skill_match_penalty(project, person, role_map) if person else ""),
            "Porcentaje coincidencia skills obligatorias": percent,
            "Puntaje skills obligatorias": skill_score,
            "Comentario skills obligatorias": skill_note,
            "Skills obligatorias coincidentes": ratio,
            "Skills con match": matched,
            "Skills faltantes": missing,
            **({key: value for key, value in base.current_staffing_fields(person, project).items()
                if key != "Asignaciones Staffing Q4 (detalle)"} if person else {
                    "Comentario asignación Staffing": "",
                    "SDATOOL - Proyecto Staffing actual": "",
                    "Fecha Inicio Staffing actual": "",
                    "Fecha Fin Actual Staffing": "",
                    "% Dedicación Staffing actual": "",
                }),
        })
        assignment_rows.append(row)

    def mesa_sort(value):
        match = re.search(r"\d+", str(value or ""))
        return (int(match.group()) if match else 10**9, base.norm(value))

    assignment_rows.sort(key=lambda row: (
        mesa_sort(row.get("Mesa")), base.norm(row.get("Foco estratégico (Fábrica)")),
        base.norm(row.get("Nombre Proyecto")),
        base.norm(row.get("Rol")), base.norm(row.get("Tipo requerido")),
        int(row.get("N.º plaza") or 0), int(row.get("Fila origen") or 0)))

    recommendations = base.recommend_candidates(slots, selected, people, role_map)
    commitments_by_person = defaultdict(list)
    for slot, person_index, _score in selected:
        project = slots[slot][1]
        commitments_by_person[person_index].append(
            f"{project.get('Mesa', '')} / {project.get('Nombre Proyecto', '')}")
    option_headers = [
        "equipo", "codigo", "nombre", "direccion", "seniority", "tipo_recurso",
        "familia_sugerida", "dominio", "skills_requeridas", "skills", "skills_faltantes",
        "coincidencia", "seniority_solicitado", "compromiso_to_be",
    ]
    option_rows = []
    for mesa in sorted(recommendations):
        for person_index in recommendations[mesa]:
            person = people[person_index]
            matching_projects = []
            for _, project in slots:
                family = base.demand_family(project.get("Rol"), project.get("Dominio"))
                if base.mesa_group_key(project) != mesa or family not in {"DEV", "QA"}:
                    continue
                if base.eligibility(project, person, set(), role_map)[0]:
                    matching_projects.append(project)
            if not matching_projects:
                continue
            project = max(
                matching_projects,
                key=lambda item: base.recommendation_priority(item, person, role_map),
            )
            _percent, _ratio, matched, missing, matched_count, _note = skill_report(project, person)
            required_names = []
            seen_skills = set()
            for field in base.TECH_FIELDS + base.FUNCTIONAL_FIELDS:
                name = (project.get(field) or "").strip()
                skill_key = base.norm(name)
                if name and skill_key not in seen_skills:
                    required_names.append(name)
                    seen_skills.add(skill_key)
            # Un requerimiento QA sin skills no tiene denominador para calcular
            # coincidencia; se indica N/A en vez de presentar cero falsamente.
            skill_count = len(required_names)
            coincidence = (round(matched_count / skill_count, 1) if skill_count else "N/A")
            team = " | ".join(str(project.get(field) or "").strip() for field in
                              ("Mesa", "Dirección", "Foco estratégico (Fábrica)")
                              if str(project.get(field) or "").strip())
            commitments = commitments_by_person.get(person_index, [])
            option_rows.append({
                "equipo": team,
                "codigo": person["code"],
                "nombre": person["name"],
                "direccion": person.get("staffing_direction", ""),
                "seniority": person["seniority_raw"],
                "tipo_recurso": person["resource_type"],
                "familia_sugerida": base.demand_family(project.get("Rol"), project.get("Dominio")),
                "dominio": project.get("Dominio", ""),
                "skills_requeridas": "; ".join(required_names),
                "skills": matched,
                "skills_faltantes": missing,
                "coincidencia": coincidence,
                "seniority_solicitado": project.get("Seniority", ""),
                "compromiso_to_be": " | ".join(commitments) if commitments else "Sin asignación To-Be",
            })

    option_rows.sort(key=lambda row: (
        base.norm(row["equipo"]), base.norm(row["familia_sugerida"]),
        -float(row["coincidencia"]) if isinstance(row["coincidencia"], (int, float)) else 0,
        base.norm(row["nombre"]), row["codigo"],
    ))
    write_xlsx_sheets(path, [
        ("Asignaciones", assignment_headers, assignment_rows),
        ("Opciones", option_headers, option_rows),
    ])
