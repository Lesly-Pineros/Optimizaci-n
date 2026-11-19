#!/usr/bin/env python3
"""Prepara las matrices de personas e iniciativas para el flujo MILP."""

import csv
import json
import re
import unicodedata
import zipfile
from collections import defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET

from limpiar_staffing import OUTPUT as STAFFING_FILE, extract_q4
from mapeo_roles import demand_role_family, person_role_families


ROOT = Path(__file__).resolve().parent
DATA = ROOT / "DATA"
RESULTS = ROOT / "RESULTADOS_LIMPIEZA"
SKILLS_FILE = DATA / "Matriz SkillsCOE.csv"
INITIATIVES_FILE = DATA / "ENG-Matriz Iniciativas.csv"
PEOPLE_OUTPUT = RESULTS / "Matriz SkillsCOE enriquecida.csv"
INITIATIVES_OUTPUT = RESULTS / "ENG-Matriz Iniciativas con skills.csv"
UNMATCHED_OUTPUT = RESULTS / "Codigos sin cruce SkillsCOE-Staffing.xlsx"
PEOPLE_REVIEW_OUTPUT = RESULTS / "Personas que requieren tratamiento.xlsx"

SKILLS_HEADER_ROW = 2
INITIATIVES_HEADER_ROW = 2
PLANNING_QUARTER = "1Q 2027"
PLANNING_QUARTER_INDEX = 2027 * 4 + 1
# La matriz fuente corresponde a iniciativas ENG; la Dirección es una dimensión
# organizacional distinta y no cambia el dominio requerido por estas plazas.
DOMAIN_BY_DIRECTION = {
    "fcp": "ENG",
    "systems": "ENG",
    "operaciones": "ENG",
    "ai tech": "ENG",
    "infraestructura": "ENG",
}

FILL_DOWN_FIELDS = (
    "Mesa",
    "Foco estratégico (Fábrica)",
    "SDATOOL",
    "Nombre Proyecto",
    "Fecha Inicio",
    "Fecha Fin",
    "Dirección",
    "Estado",
    "Workspace",
)
REQUIRED_FILLED_FIELDS = ("Mesa", "Dirección", "Estado")
STAFFING_FIELDS = (
    ("interno", "Tipo de Recurso Staffing"),
    ("% dedicacion", "% Dedicación Staffing"),
    ("rol", "Rol Staffing"),
    ("dominio", "Dominio"),
    ("direccion", "Dirección Staffing"),
    ("SDATOOL - Proyecto", "SDATOOL - Proyecto"),
    ("Fecha Inicio", "Fecha Inicio"),
    ("Fecha Fin Actual", "Fecha Fin Actual"),
    ("Proyecto", "Proyecto Staffing"),
    ("Mesa actual", "Mesa Staffing actual"),
)
STAFFING_ASSIGNMENTS_COLUMN = "Asignaciones Staffing Q4"
STAFFING_ASSIGNMENT_SOURCE_COLUMN = "Asignación Staffing"
STAFFING_STATUS_COLUMN = "Estado Staffing"
INVALID_VALUES = {"na", "n a", "n/a", "na no aplica", "null", "none", "nan",
                  "cod eliminado", "codigo eliminado", "eliminado", "sin rol", ""}
VALID_SENIORITY = {"junior", "mid level", "senior", "lead"}
SENIORITY_LABELS = {
    "junior": "Junior",
    "mid level": "Mid-Level",
    "senior": "Senior",
    "lead": "Lead",
}
XLSX_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XLSX_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", XLSX_NS)
ET.register_namespace("r", XLSX_REL_NS)


def normalize(value):
    value = unicodedata.normalize("NFKD", str(value or ""))
    value = "".join(char for char in value if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", value).strip().casefold()


def code_key(value):
    return re.sub(r"[^A-Z0-9]", "", unicodedata.normalize("NFKC", str(value or "")).upper())


def read_csv(path, header_row=0):
    last_error = None
    for encoding in ("utf-8-sig", "cp1252", "latin1"):
        try:
            with path.open(encoding=encoding, newline="") as file:
                rows = list(csv.reader(file))
            header = rows[header_row]
            records = []
            for values in rows[header_row + 1:]:
                values = values + [""] * max(0, len(header) - len(values))
                records.append(dict(zip(header, values)))
            return header, records
        except UnicodeDecodeError as exc:
            last_error = exc
    raise last_error


def find_column(header, name):
    matches = [column for column in header if normalize(column) == normalize(name)]
    if len(matches) != 1:
        raise ValueError(f"Se esperaba una columna '{name}' única; encontradas: {len(matches)}")
    return matches[0]


def normalize_resource_type(value):
    value = normalize(value)
    if value == "subsidiaria":
        return "externo"
    return value


def staffing_by_code():
    _, staffing = read_csv(STAFFING_FILE)
    code_column = find_column(list(staffing[0]), "codigo") if staffing else "codigo"
    grouped = defaultdict(lambda: {source: [] for source, _ in STAFFING_FIELDS})
    assignments = defaultdict(list)
    for row in staffing:
        code = code_key(row.get(code_column))
        if not code:
            continue
        for source, _ in STAFFING_FIELDS:
            value = (row.get(source) or "").strip()
            if value and value not in grouped[code][source]:
                grouped[code][source].append(value)
        raw_assignment = row.get(STAFFING_ASSIGNMENT_SOURCE_COLUMN, "")
        if raw_assignment:
            try:
                assignment = json.loads(raw_assignment)
            except json.JSONDecodeError:
                assignment = {}
        else:
            assignment = {}
        if not assignment:
            assignment = {
                "sda": row.get("SDATOOL - Proyecto", ""),
                "proyecto": row.get("Proyecto", ""),
                "mesa": row.get("Mesa actual", ""),
                "fecha_inicio": row.get("Fecha Inicio", ""),
                "fecha_fin_actual": row.get("Fecha Fin Actual", ""),
                "dedicacion": row.get("% dedicacion", ""),
                "fuente": row.get("fuente staffing", ""),
            }
        if any(str(value or "").strip() for value in assignment.values()):
            assignments[code].append(assignment)

    result = {}
    for code, fields in grouped.items():
        result[code] = {source: " | ".join(values) for source, values in fields.items()}
        result[code][STAFFING_ASSIGNMENTS_COLUMN] = json.dumps(
            assignments[code], ensure_ascii=False)
    return result


def prepare_people():
    skills_header, people = read_csv(SKILLS_FILE, SKILLS_HEADER_ROW - 1)
    staffing = staffing_by_code()
    skills_code_column = find_column(skills_header, "Código")
    name_column = find_column(skills_header, "Nombre")
    role_column = find_column(skills_header, "Rol Principal")
    seniority_column = find_column(skills_header, "Nivel del Rol Actual")
    type_column = find_column(skills_header, "Tipo de Recurso")
    output_header = skills_header + [target for _, target in STAFFING_FIELDS]
    for extra in ("Seniority", STAFFING_STATUS_COLUMN, STAFFING_ASSIGNMENTS_COLUMN):
        if extra not in output_header:
            output_header.append(extra)
    code_counts = defaultdict(int)
    for person in people:
        key = code_key(person.get(skills_code_column))
        if key:
            code_counts[key] += 1

    ready = []
    treatment = []
    matched_codes = set()
    for index, person in enumerate(people):
        key = code_key(person.get(skills_code_column))
        linked = staffing.get(key, {})
        output = dict(person)
        for source, target in STAFFING_FIELDS:
            output[target] = linked.get(source, "") or output.get(target, "")
        output[STAFFING_STATUS_COLUMN] = "EN_STAFFING_Q4" if key and key in staffing else "BACKLOG"
        output[STAFFING_ASSIGNMENTS_COLUMN] = linked.get(STAFFING_ASSIGNMENTS_COLUMN, "[]")

        reasons = []
        if not key:
            reasons.append("Código vacío en SkillsCOE")
        elif code_counts[key] > 1:
            reasons.append("Código duplicado en SkillsCOE")
        if key and key in staffing:
            matched_codes.add(key)

        resource_type = normalize_resource_type(person.get(type_column))
        if resource_type not in {"interno", "externo"}:
            reasons.append(f"Tipo de recurso no soportado o vacío: {person.get(type_column, '').strip() or 'vacío'}")
        staffing_types = {normalize_resource_type(value) for value in linked.get("interno", "").split("|")
                          if normalize_resource_type(value)}
        if len(staffing_types) > 1:
            reasons.append("Tipos de recurso contradictorios en Staffing Q4")
        if staffing_types and resource_type in {"interno", "externo"} and resource_type not in staffing_types:
            reasons.append("Tipo de recurso no coincide con Staffing Q4")
        role_value = normalize(person.get(role_column))
        if not person_role_families(role_value):
            reasons.append("Rol principal nulo o basura")
        seniority_value = normalize(person.get(seniority_column))
        if seniority_value not in VALID_SENIORITY:
            reasons.append("Nivel del rol actual vacío o inválido")
        output["Seniority"] = SENIORITY_LABELS.get(seniority_value, "")
        domain_value = normalize(output.get("Dominio"))
        if (not domain_value or domain_value in INVALID_VALUES
                or domain_value.startswith(("na ", "n a ", "sin dominio"))):
            reasons.append("Dominio vacío o basura; excluido del pool elegible")

        if reasons:
            treatment.append([
                str(index + SKILLS_HEADER_ROW + 1),
                person.get(skills_code_column, ""), person.get(name_column, ""),
                person.get(role_column, ""), person.get(seniority_column, ""),
                person.get(type_column, ""), output.get("Tipo de Recurso Staffing", ""),
                output.get("Dominio", ""), output.get("Dirección Staffing", ""),
                output.get("% Dedicación Staffing", ""), output[STAFFING_STATUS_COLUMN],
                output.get("SDATOOL - Proyecto", ""), output.get("Fecha Inicio", ""),
                output.get("Fecha Fin Actual", ""), " | ".join(dict.fromkeys(reasons)),
            ])
        else:
            ready.append(output)

    with PEOPLE_OUTPUT.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=output_header)
        writer.writeheader()
        for output in ready:
            writer.writerow(output)

    write_review_workbook(PEOPLE_REVIEW_OUTPUT, [
        ("Requieren tratamiento",
         ["Fila SkillsCOE", "Código", "Nombre", "Rol Principal", "Nivel del Rol Actual",
          "Tipo de Recurso", "Tipo de Recurso Staffing", "Dominio Staffing",
          "Dirección Staffing", "% Dedicación Staffing", "Estado Staffing",
          "SDATOOL - Proyecto", "Fecha Inicio", "Fecha Fin Actual", "Motivo"], treatment),
    ])
    print(f"Personas listas para MILP: {len(ready)}; con tratamiento: {len(treatment)}; "
          f"códigos con cruce en Staffing Q4: {len(matched_codes)}.")
    print(f"Excel de personas para revisión: {PEOPLE_REVIEW_OUTPUT}")


def quarter_index(value):
    value = str(value or "").strip().upper()
    match = re.search(r"([1-4])\s*Q\s*(\d{4})", value)
    if not match:
        match = re.search(r"Q\s*([1-4])\s*(\d{4})", value)
    if not match:
        return None
    return int(match.group(2)) * 4 + int(match.group(1))


def prepare_initiatives():
    header, all_rows = read_csv(INITIATIVES_FILE, INITIATIVES_HEADER_ROW - 1)
    indexes = {normalize(name): name for name in header}
    missing = [name for name in (*FILL_DOWN_FIELDS, "Rol", "Seniority")
               if normalize(name) not in indexes]
    if missing:
        raise ValueError(f"Faltan columnas en la matriz de iniciativas: {', '.join(missing)}")

    filled = {name: "" for name in FILL_DOWN_FIELDS}
    selected = []
    validation_errors = []
    excluded = defaultdict(int)
    unmapped_roles = set()
    mesa_focus_groups_source = set()
    source_rows = 0
    for source_index, source in enumerate(all_rows, start=INITIATIVES_HEADER_ROW + 1):
        excel_row = source_index
        row = dict(source)
        if not any((value or "").strip() for value in row.values()):
            continue
        source_rows += 1

        for field in FILL_DOWN_FIELDS:
            actual = indexes[normalize(field)]
            value = (row.get(actual) or "").strip()
            if value:
                filled[field] = value
            elif filled[field]:
                row[actual] = filled[field]

        mesa_value = normalize(row.get(indexes[normalize("Mesa")], ""))
        focus_value = normalize(row.get(indexes[normalize("Foco estratégico (Fábrica)")], ""))
        if re.fullmatch(r"mesa\s+\d+", mesa_value) and focus_value:
            mesa_focus_groups_source.add((mesa_value, focus_value))

        role = row.get(indexes[normalize("Rol")], "")
        seniority = normalize(row.get(indexes[normalize("Seniority")], ""))
        seniority = re.sub(r"[-–—]+", " ", seniority)
        seniority = re.sub(r"\s+", " ", seniority).strip()
        skill_fields = [name for name in header if normalize(name).startswith(
            ("habilidades tecnicas obligatorias", "habilidades funcionales obligatorias"))]
        skill_count = len({normalize(row.get(name, "")) for name in skill_fields
                           if (row.get(name) or "").strip()})
        family = demand_role_family(role)
        if family is None:
            excluded["rol sin mapeo"] += 1
            if role.strip():
                unmapped_roles.add(role.strip())
            continue
        if seniority not in VALID_SENIORITY:
            excluded["seniority vacío o no homologado"] += 1
            continue
        if skill_count < 1:
            excluded["sin skills obligatorias solicitadas"] += 1
            continue

        for field in REQUIRED_FILLED_FIELDS:
            if not (row.get(indexes[normalize(field)]) or "").strip():
                validation_errors.append(f"fila Excel {excel_row}: falta {field}")
        direction = row.get(indexes[normalize("Dirección")], "")
        mapped_domain = DOMAIN_BY_DIRECTION.get(normalize(direction))
        if not mapped_domain:
            validation_errors.append(
                f"fila Excel {excel_row}: no hay regla de dominio para Dirección '{direction}'"
            )
            continue
        row["Dominio"] = mapped_domain
        row["Fila origen"] = str(excel_row)
        selected.append(row)

    if validation_errors:
        raise ValueError("Hay filas sin contexto requerido:\n" + "\n".join(validation_errors))
    if not selected:
        raise ValueError("No hay iniciativas que cumplan rol, seniority y skills obligatorias.")

    output_header = list(header)
    if normalize("Dominio") not in indexes:
        output_header.append("Dominio")
    output_header.append("Fila origen")

    with INITIATIVES_OUTPUT.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=output_header)
        writer.writeheader()
        writer.writerows(selected)

    print(f"Iniciativas revisadas: {source_rows}; conservadas: {len(selected)}; "
          f"exclusiones: {dict(excluded)}.")
    eligible_groups = {
        (normalize(row.get(indexes[normalize("Mesa")], "")),
         normalize(row.get(indexes[normalize("Foco estratégico (Fábrica)")], "")))
        for row in selected
    }
    print(f"Grupos Mesa + Foco estratégico: {len(mesa_focus_groups_source)} en la matriz; "
          f"{len(eligible_groups)} con demanda elegible.")
    if unmapped_roles:
        print(f"Roles que requieren mapeo en mapeo_roles.py: {sorted(unmapped_roles)}")


def excel_column(number):
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def write_review_workbook(path, sheets):
    content_types = ET.Element("{http://schemas.openxmlformats.org/package/2006/content-types}Types")
    ET.SubElement(content_types, "{http://schemas.openxmlformats.org/package/2006/content-types}Default",
                  {"Extension": "rels", "ContentType": "application/vnd.openxmlformats-package.relationships+xml"})
    ET.SubElement(content_types, "{http://schemas.openxmlformats.org/package/2006/content-types}Default",
                  {"Extension": "xml", "ContentType": "application/xml"})
    ET.SubElement(content_types, "{http://schemas.openxmlformats.org/package/2006/content-types}Override",
                  {"PartName": "/xl/workbook.xml", "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"})
    for index in range(1, len(sheets) + 1):
        ET.SubElement(content_types, "{http://schemas.openxmlformats.org/package/2006/content-types}Override",
                      {"PartName": f"/xl/worksheets/sheet{index}.xml", "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"})

    package_rels = ET.Element("{http://schemas.openxmlformats.org/package/2006/relationships}Relationships")
    ET.SubElement(package_rels, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
                  {"Id": "rId1", "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "Target": "xl/workbook.xml"})

    workbook = ET.Element(f"{{{XLSX_NS}}}workbook")
    workbook_sheets = ET.SubElement(workbook, f"{{{XLSX_NS}}}sheets")
    workbook_rels = ET.Element("{http://schemas.openxmlformats.org/package/2006/relationships}Relationships")
    worksheet_xml = {}
    for index, (title, headers, rows) in enumerate(sheets, 1):
        ET.SubElement(workbook_sheets, f"{{{XLSX_NS}}}sheet",
                      {"name": title, "sheetId": str(index), f"{{{XLSX_REL_NS}}}id": f"rId{index}"})
        ET.SubElement(workbook_rels, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
                      {"Id": f"rId{index}", "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "Target": f"worksheets/sheet{index}.xml"})
        sheet_root = ET.Element(f"{{{XLSX_NS}}}worksheet")
        sheet_data = ET.SubElement(sheet_root, f"{{{XLSX_NS}}}sheetData")
        for row_number, values in enumerate([headers, *rows], 1):
            row_element = ET.SubElement(sheet_data, f"{{{XLSX_NS}}}row", {"r": str(row_number)})
            for column_number, value in enumerate(values, 1):
                cell = ET.SubElement(row_element, f"{{{XLSX_NS}}}c",
                                     {"r": f"{excel_column(column_number)}{row_number}", "t": "inlineStr"})
                inline = ET.SubElement(cell, f"{{{XLSX_NS}}}is")
                text_node = ET.SubElement(inline, f"{{{XLSX_NS}}}t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"})
                text_node.text = str(value or "")
        auto_filter = ET.SubElement(sheet_root, f"{{{XLSX_NS}}}autoFilter")
        auto_filter.set("ref", f"A1:{excel_column(len(headers))}{len(rows) + 1}")
        worksheet_xml[f"xl/worksheets/sheet{index}.xml"] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as workbook_file:
        workbook_file.writestr("[Content_Types].xml", ET.tostring(content_types, encoding="utf-8", xml_declaration=True))
        workbook_file.writestr("_rels/.rels", ET.tostring(package_rels, encoding="utf-8", xml_declaration=True))
        workbook_file.writestr("xl/workbook.xml", ET.tostring(workbook, encoding="utf-8", xml_declaration=True))
        workbook_file.writestr("xl/_rels/workbook.xml.rels", ET.tostring(workbook_rels, encoding="utf-8", xml_declaration=True))
        for part, xml in worksheet_xml.items():
            workbook_file.writestr(part, xml)


def prepare_unmatched_report():
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
    print(f"Sin cruce: {len(skills_only)} códigos en SkillsCOE sin Staffing; {len(staffing_only)} en Staffing sin SkillsCOE.")
    print(f"Excel de revisión: {UNMATCHED_OUTPUT}")


def main():
    RESULTS.mkdir(exist_ok=True)
    # Mantiene Staffing Q4 al día desde el Excel antes de hacer el cruce.
    extract_q4()
    prepare_people()
    prepare_initiatives()
    prepare_unmatched_report()
    print(f"Salidas MILP: {PEOPLE_OUTPUT} ; {INITIATIVES_OUTPUT}")


if __name__ == "__main__":
    main()
