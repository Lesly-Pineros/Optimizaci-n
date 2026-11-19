"""Exporta un To-Be XLSX sin dependencias adicionales de escritura."""
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile


def _col(n):
    result = ""
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def write_xlsx(path, headers, rows, sheet_name="Asignaciones"):
    """Escribe un libro sencillo con texto inline, compatible con Excel/LibreOffice."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    all_rows = [headers] + [[row.get(h, "") for h in headers] for row in rows]
    xml_rows = []
    for ri, row in enumerate(all_rows, 1):
        cells = []
        for ci, value in enumerate(row, 1):
            ref = f"{_col(ci)}{ri}"
            value = "" if value is None else str(value)
            cells.append(f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{escape(value)}</t></is></c>')
        xml_rows.append(f'<row r="{ri}">{"".join(cells)}</row>')
    sheet = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
             f'<sheetData>{"".join(xml_rows)}</sheetData></worksheet>')
    workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                f'<sheets><sheet name="{escape(sheet_name[:31])}" sheetId="1" r:id="rId1"/></sheets></workbook>')
    with ZipFile(path, "w", ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?>'
                    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                    '<Default Extension="xml" ContentType="application/xml"/>'
                    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                    '</Types>')
        zf.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                    '</Relationships>')
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                    '</Relationships>')
        zf.writestr("xl/worksheets/sheet1.xml", sheet)


def export_tobe(path, source_header, projects, slots, selected, people, skill_fit, skill_fields):
    """Una fila por fila fuente; detalle de varios slots en celdas multilínea."""
    from collections import defaultdict
    import asignacion_milp as base
    assignment_counts = defaultdict(int)
    for _slot, person_index, _score in selected:
        assignment_counts[person_index] += 1
    by_source = defaultdict(list)
    for slot, person_index, _score in selected:
        source_row, project = slots[slot]
        person = people[person_index]
        covered = []
        for field in skill_fields:
            name = project.get(field, "")
            normalized = __import__("asignacion_milp").norm(name)
            level = person["skills"].get(normalized, 0)
            if name.strip() and level > 0:
                covered.append(f"{name.strip()} (nivel {level})")
        expected = base.SENIORITY_SKILL_LEVEL.get(base.norm(project.get("Seniority")), 1)
        missing = []
        for field in base.TECH_FIELDS + base.FUNCTIONAL_FIELDS:
            name = project.get(field, "")
            if not name.strip():
                continue
            level = person["skills"].get(base.norm(name), 0)
            if level < expected:
                missing.append(f"{name.strip()} (nivel {level}; requiere {expected})")
        by_source[int(source_row)].append({
            "Nombre": person["name"], "Código": person["code"],
            "Rol": person["role_raw"], "Seniority": person["seniority_raw"],
            "Skills": "; ".join(covered) or "Sin skill coincidente",
            "Missing": "; ".join(missing) or "Ninguna",
            "Capacity": (f"Asignada a {assignment_counts[person_index]} plazas; capacidad base 1"
                         if assignment_counts[person_index] > base.MAX_ASSIGNMENTS_PER_PERSON else ""),
            "Tipo": project.get("_required_type", ""),
        })
    headers = list(source_header) + ["Personas asignadas", "Códigos asignados", "Roles persona",
                                      "Seniority personas", "Tipo de asignación", "Skills con match (nivel)",
                                      "Skills faltantes", "Alerta de capacidad"]
    output = []
    for idx, project in enumerate(projects, 2):
        try:
            source_id = int(project.get("Fila origen") or idx)
        except ValueError:
            source_id = idx
        assigned = by_source.get(source_id, [])
        row = {h: project.get(h, "") for h in source_header}
        row.update({
            "Personas asignadas": "\n".join(p["Nombre"] for p in assigned),
            "Códigos asignados": "\n".join(p["Código"] for p in assigned),
            "Roles persona": "\n".join(p["Rol"] for p in assigned),
            "Seniority personas": "\n".join(p["Seniority"] for p in assigned),
            "Tipo de asignación": "\n".join(p["Tipo"] for p in assigned),
            "Skills con match (nivel)": "\n".join(p["Skills"] for p in assigned),
            "Skills faltantes": "\n".join(f"{p['Nombre']}: {p['Missing']}" for p in assigned),
            "Alerta de capacidad": "\n".join(f"{p['Nombre']}: {p['Capacity']}" for p in assigned if p["Capacity"]),
        })
        output.append(row)
    write_xlsx(path, headers, output)
