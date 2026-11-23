#!/usr/bin/env python3
"""Prepara las matrices de personas e iniciativas para el flujo MILP.

El flujo cruza la matriz de skills con Staffing Q4, conserva los registros que
cumplen las reglas de calidad y genera reportes Excel para los casos pendientes.
"""

import csv
import json
import re
import unicodedata
import zipfile
from collections import defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET

from limpiar_staffing import (
    NS as STAFFING_XLSX_NS,
    OUTPUT as STAFFING_FILE,
    extract_q4,
    read_shared_strings,
    worksheet_rows,
)
from mapeo_roles import (
    canonical_demand_role,
    canonical_person_role,
    demand_role_family,
    person_role_families,
)


# Rutas del proyecto: los insumos originales viven en DATA y las salidas en
# RESULTADOS_LIMPIEZA, ambas relativas a la ubicación de este script.
ROOT = Path(__file__).resolve().parent
DATA = ROOT / "DATA"
RESULTS = ROOT / "RESULTADOS_LIMPIEZA"
# ENG llega en CSV; los demás dominios se detectan en pestañas de Iniciativas.xlsx
# cuyo nombre siga el patrón "<DOMINIO> X Mesa" o "<DOMINIO> X Mesas".
SKILLS_FILE = DATA / "Matriz SkillsCOE.csv"
ENG_INITIATIVES_FILE = DATA / "ENG-iniciativas.csv"
INITIATIVES_WORKBOOK = DATA / "Iniciativas.xlsx"
# Las hojas que cumplan este patrón aportan su prefijo como dominio de salida.
INITIATIVE_DOMAIN_SHEET_PATTERN = r"^(.+?)\s+x\s+mesas?(?:\s+\d+)?$"
# Archivos de salida para el modelo y para la revisión manual de excepciones.
PEOPLE_OUTPUT = RESULTS / "Matriz SkillsCOE enriquecida.csv"
INITIATIVES_OUTPUT = RESULTS / "GRM-iniciativas con skills.csv"
UNMATCHED_OUTPUT = RESULTS / "Codigos sin cruce SkillsCOE-Staffing.xlsx"
PEOPLE_REVIEW_OUTPUT = RESULTS / "Personas que requieren tratamiento.xlsx"
UNMAPPED_SKILLS_OUTPUT = RESULTS / "Skills de iniciativas sin mapeo.xlsx"
UNMAPPED_ROLES_OUTPUT = RESULTS / "Roles sin mapeo.xlsx"

# Las fuentes tienen una fila de título antes de sus encabezados.
SKILLS_HEADER_ROW = 2
INITIATIVES_HEADER_ROW = 2
# Periodo usado por las reglas de planificación del proyecto.
PLANNING_QUARTER = "1Q 2027"
PLANNING_QUARTER_INDEX = 2027 * 4 + 1
# Campos de iniciativa que suelen aparecer en una sola fila por celdas combinadas.
FILL_DOWN_FIELDS = (
    "Mesa",
    "Foco estratégico (Fábrica)",
    "Nombre Proyecto",
    "Fecha Inicio",
    "Fecha Fin",
    "Dirección",
    "Estado",
    "Workspace",
)
# Si estos campos siguen vacíos después del relleno, la fila se reporta como error.
REQUIRED_FILLED_FIELDS = ("Mesa", "Dirección", "Estado")
# Pares de columna origen en Staffing y nombre que tendrá en la matriz enriquecida.
STAFFING_FIELDS = (
    ("interno", "Tipo de Recurso Staffing"),
    ("% dedicacion", "% Dedicación Staffing"),
    ("dominio", "Dominio"),
    ("direccion", "Dirección Staffing"),
    ("SDATOOL - Proyecto", "SDATOOL - Proyecto"),
    ("Fecha Inicio", "Fecha Inicio"),
    ("Fecha Fin Actual", "Fecha Fin Actual"),
    ("Proyecto", "Proyecto Staffing"),
    ("Mesa actual", "Mesa Staffing actual"),
)
# Columnas añadidas para el estado del cruce y el detalle de todas las asignaciones.
STAFFING_ASSIGNMENTS_COLUMN = "Asignaciones Staffing Q4"
STAFFING_ASSIGNMENT_SOURCE_COLUMN = "Asignación Staffing"
STAFFING_STATUS_COLUMN = "Estado Staffing"
# Indicadores de valores inválidos usados para filtrar campos incompletos.
INVALID_VALUES = {"na", "n a", "n/a", "na no aplica", "null", "none", "nan",
                  "cod eliminado", "codigo eliminado", "eliminado", "sin rol", ""}
VALID_SENIORITY = {"junior", "mid level", "senior", "lead"}
# Etiquetas de seniority que se guardan normalizadas en la salida enriquecida.
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
    """Normaliza texto para comparar encabezados y categorías sin tildes/caso."""
    value = unicodedata.normalize("NFKD", str(value or ""))
    value = "".join(char for char in value if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", value).strip().casefold()


def code_key(value):
    """Deja un código comparable quitando espacios y signos de puntuación."""
    return re.sub(r"[^A-Z0-9]", "", unicodedata.normalize("NFKC", str(value or "")).upper())


def read_csv(path, header_row=0):
    """Lee un CSV con encabezado configurable y prueba codificaciones habituales.

    Las filas se devuelven como diccionarios. Si una fila tiene menos celdas que
    el encabezado, se completa con cadenas vacías para mantener las columnas.
    """
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
    """Busca una columna por nombre normalizado y exige una coincidencia única."""
    matches = [column for column in header if normalize(column) == normalize(name)]
    if len(matches) != 1:
        raise ValueError(f"Se esperaba una columna '{name}' única; encontradas: {len(matches)}")
    return matches[0]


def normalize_resource_type(value):
    """Homologa subsidiaria como externo; devuelve las demás categorías tal cual."""
    value = normalize(value)
    if value == "subsidiaria":
        return "externo"
    return value


def seniority_key(value):
    """Homologa variantes de seniority a junior/mid level/senior/lead."""
    value = normalize(value)
    value = re.sub(r"[-–—]+", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    aliases = {
        "midlevel": "mid level",
        "mid lvl": "mid level",
        "middle level": "mid level",
    }
    return aliases.get(value, value)


def skill_match_key(value):
    """Crea una clave de skill insensible a tildes, caso y puntuación."""
    value = unicodedata.normalize("NFKD", str(value or ""))
    value = "".join(char for char in value if not unicodedata.combining(char))
    value = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", " ", value)
    # Conserva símbolos que cambian el significado de lenguajes como C# o C++.
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9+#]+", " ", value.casefold())).strip()


def normalized_unknown_skill(value):
    """Da un formato uniforme a skills sin columna equivalente en SkillsCOE."""
    value = unicodedata.normalize("NFKD", str(value or ""))
    value = "".join(char for char in value if not unicodedata.combining(char))
    value = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", " ", value)
    return re.sub(r"[^A-Za-z0-9]+", " ", value).strip().upper()


def skills_catalog():
    """Devuelve las etiquetas oficiales de skills de la matriz SkillsCOE."""
    header, _ = read_csv(SKILLS_FILE, SKILLS_HEADER_ROW - 1)
    role_level_index = next(i for i, name in enumerate(header)
                            if normalize(name) == normalize("Nivel del Rol Actual"))
    end_index = next((i for i, name in enumerate(header)
                      if normalize(name) == normalize("Skill a desarrollar en los próximos meses")),
                     len(header))
    return [name.strip() for name in header[role_level_index + 1:end_index]
            if name and name.strip()]


def staffing_by_code():
    """Agrupa los registros de Staffing Q4 por código de persona.

    Conserva valores distintos de cada campo sin duplicarlos y reúne las
    asignaciones de proyecto como JSON para no perder el detalle por proyecto.
    """
    _, staffing = read_csv(STAFFING_FILE)
    code_column = find_column(list(staffing[0]), "codigo") if staffing else "codigo"
    grouped = defaultdict(lambda: {source: [] for source, _ in STAFFING_FIELDS})
    assignments = defaultdict(list)
    original_codes = {}
    for row in staffing:
        code = code_key(row.get(code_column))
        if not code:
            continue
        original_codes.setdefault(code, (row.get(code_column) or "").strip())
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
        result[code]["codigo"] = original_codes[code]
        result[code][STAFFING_ASSIGNMENTS_COLUMN] = json.dumps(
            assignments[code], ensure_ascii=False)
    return result


def prepare_people():
    """Enriquece la población de SkillsCOE con Staffing Q4.

    SkillsCOE define personas, skills y rol. Los registros que no estén en
    Staffing se conservan como BACKLOG; si no tienen Dominio no serán elegibles.
    Roles NA o no mapeados se excluyen y se documentan en el Excel de revisión.
    """
    skills_header, people = read_csv(SKILLS_FILE, SKILLS_HEADER_ROW - 1)
    staffing = staffing_by_code()
    skills_code_column = find_column(skills_header, "Código")
    name_column = find_column(skills_header, "Nombre")
    role_column = find_column(skills_header, "Rol Principal")
    seniority_column = find_column(skills_header, "Nivel del Rol Actual")
    skills_by_code = defaultdict(list)
    for index, person in enumerate(people, start=SKILLS_HEADER_ROW + 1):
        key = code_key(person.get(skills_code_column))
        if key:
            skills_by_code[key].append((index, person))

    # La salida conserva el esquema esperado por el modelo. El campo de rol es
    # Rol Principal de SkillsCOE; no se agrega ni se conserva el rol de Staffing.
    output_header = [
        column for column in skills_header
        if normalize(column) not in {"rol homologado", "rol staffing"}
    ] + [target for _, target in STAFFING_FIELDS]
    for extra in ("Seniority", STAFFING_STATUS_COLUMN, STAFFING_ASSIGNMENTS_COLUMN):
        if extra not in output_header:
            output_header.append(extra)

    ready = []
    treatment = []
    excluded_roles = defaultdict(int)
    for key, matches in skills_by_code.items():
        if len(matches) > 1:
            index, person = matches[0]
            treatment.append([
                str(index), person.get(skills_code_column, ""), person.get(name_column, ""),
                person.get(role_column, ""), person.get(seniority_column, ""), "", "", "", "",
                "", "", "", "", "", "Código duplicado en SkillsCOE",
            ])
            continue

        index, person = matches[0]
        original_role = (person.get(role_column) or "").strip()
        canonical_role = canonical_person_role(original_role)
        if not canonical_role:
            reason = "Rol NA/vacío" if normalize(original_role).startswith("na") or not original_role else "Rol SkillsCOE sin mapeo"
            excluded_roles[original_role or "(vacío)"] += 1
            treatment.append([
                str(index), person.get(skills_code_column, ""), person.get(name_column, ""),
                original_role, person.get(seniority_column, ""), "", "", "", "", "",
                "", "", "", "", reason,
            ])
            continue

        linked = staffing.get(key, {})
        output = dict(person)
        for source, target in STAFFING_FIELDS:
            output[target] = linked.get(source, "")
        output[skills_code_column] = linked.get("codigo", person.get(skills_code_column, ""))
        if linked:
            output[find_column(output_header, "Tipo de Recurso")] = linked.get("interno", "")
        output[role_column] = canonical_role
        output[STAFFING_STATUS_COLUMN] = "EN_STAFFING_Q4" if linked else "BACKLOG"
        output[STAFFING_ASSIGNMENTS_COLUMN] = linked.get(STAFFING_ASSIGNMENTS_COLUMN, "[]")

        seniority_value = seniority_key(person.get(seniority_column))
        output["Seniority"] = SENIORITY_LABELS.get(seniority_value, "")
        output[seniority_column] = SENIORITY_LABELS.get(seniority_value, "")
        if seniority_value not in VALID_SENIORITY:
            treatment.append([
                str(index), output.get(skills_code_column, ""), output.get(name_column, ""),
                canonical_role, output.get(seniority_column, ""), output.get("Tipo de Recurso", ""),
                output.get("Tipo de Recurso Staffing", ""), output.get("Dominio", ""),
                output.get("Dirección Staffing", ""), output.get("% Dedicación Staffing", ""),
                output[STAFFING_STATUS_COLUMN], output.get("SDATOOL - Proyecto", ""),
                output.get("Fecha Inicio", ""), output.get("Fecha Fin Actual", ""),
                "Seniority vacío o inválido",
            ])
            continue
        if linked:
            resource_type = normalize_resource_type(linked.get("interno", ""))
            staffing_types = {normalize_resource_type(value) for value in linked.get("interno", "").split("|")
                              if normalize_resource_type(value)}
            if resource_type not in {"interno", "externo"} or len(staffing_types) > 1:
                treatment.append([
                    str(index), person.get(skills_code_column, ""), person.get(name_column, ""),
                    canonical_role, output[seniority_column], linked.get("interno", ""),
                    linked.get("interno", ""), linked.get("dominio", ""), linked.get("direccion", ""),
                    linked.get("% dedicacion", ""), output[STAFFING_STATUS_COLUMN],
                    linked.get("SDATOOL - Proyecto", ""), linked.get("Fecha Inicio", ""),
                    linked.get("Fecha Fin Actual", ""), "Tipo de recurso vacío o contradictorio",
                ])
                continue
        ready.append(output)
        domain_value = normalize(output.get("Dominio"))
        if (not domain_value or domain_value in INVALID_VALUES
                or domain_value.startswith(("na ", "n a ", "sin dominio"))):
            treatment.append([
                str(index), output.get(skills_code_column, ""), output.get(name_column, ""),
                canonical_role, output.get(seniority_column, ""), output.get("Tipo de Recurso", ""),
                output.get("Tipo de Recurso Staffing", ""), output.get("Dominio", ""),
                output.get("Dirección Staffing", ""), output.get("% Dedicación Staffing", ""),
                output[STAFFING_STATUS_COLUMN], output.get("SDATOOL - Proyecto", ""),
                output.get("Fecha Inicio", ""), output.get("Fecha Fin Actual", ""),
                "Dominio vacío o basura; excluido del pool elegible",
            ])

    with PEOPLE_OUTPUT.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=output_header, extrasaction="ignore")
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
    print(f"Personas SkillsCOE enriquecidas listas para el modelo: {len(ready)}; "
          f"BACKLOG (sin Staffing Q4): {sum(row[STAFFING_STATUS_COLUMN] == 'BACKLOG' for row in ready)}; "
          f"con tratamiento/excluidas: {len(treatment)}.")
    if excluded_roles:
        print(f"Roles de SkillsCOE excluidos y reportados: {dict(excluded_roles)}")
    print(f"Excel de personas para revisión: {PEOPLE_REVIEW_OUTPUT}")


def quarter_index(value):
    """Convierte etiquetas como '1Q 2027' o 'Q1 2027' a un índice ordenable."""
    value = str(value or "").strip().upper()
    match = re.search(r"([1-4])\s*Q\s*(\d{4})", value)
    if not match:
        match = re.search(r"Q\s*([1-4])\s*(\d{4})", value)
    if not match:
        return None
    return int(match.group(2)) * 4 + int(match.group(1))


def matrix_records(matrix, header_row_index):
    """Convierte filas leídas desde una pestaña Excel a diccionarios por fila."""
    if len(matrix) <= header_row_index:
        raise ValueError("La pestaña de iniciativas no contiene la fila de encabezados")
    header = matrix[header_row_index]
    records = []
    for values in matrix[header_row_index + 1:]:
        values = values + [""] * max(0, len(header) - len(values))
        records.append(dict(zip(header, values)))
    return header, records


def initiative_output_path(domain):
    """Devuelve el CSV de salida del dominio, manteniendo estable la ruta GRM."""
    if domain == "GRM":
        return INITIATIVES_OUTPUT
    safe_domain = re.sub(r"[^A-Z0-9_-]+", "_", domain.upper())
    return RESULTS / f"{safe_domain}-iniciativas con skills.csv"


def prepare_initiative_file(domain, header, all_rows, first_excel_row=3,
                            unmapped_role_rows=None):
    """Depura una matriz de iniciativas para el dominio indicado.

    Completa el contexto de proyectos en filas continuadas. Conserva demandas
    con rol mapeado, seniority válido y skills obligatorias. QA puede no declarar
    skills, porque las reglas permiten cubrir esa vacante con menor afinidad.
    Los campos de nivel general técnico y nivel IA se mantienen sin cambios.
    """
    indexes = {normalize(name): name for name in header}
    missing = [name for name in (*FILL_DOWN_FIELDS, "Rol", "Seniority")
               if normalize(name) not in indexes]
    if missing:
        raise ValueError(f"Faltan columnas en la matriz de iniciativas {domain}: {', '.join(missing)}")

    filled = {name: "" for name in FILL_DOWN_FIELDS}
    selected = []
    sdas_by_group = defaultdict(dict)
    excluded = defaultdict(int)
    unmapped_roles = set()
    mesa_focus_groups_source = set()
    source_rows = 0
    canonical_skills = {}
    for name in skills_catalog():
        key = skill_match_key(name)
        if key in canonical_skills and canonical_skills[key] != name:
            raise ValueError(
                f"Hay nombres de skills ambiguos en SkillsCOE: "
                f"{canonical_skills[key]!r} y {name!r}")
        canonical_skills[key] = name
    mandatory_skill_fields = [name for name in header if normalize(name).startswith(
        ("habilidades tecnicas obligatorias", "habilidades funcionales obligatorias"))]
    all_initiative_skill_fields = [name for name in header if normalize(name).startswith((
        "habilidades tecnicas obligatorias", "habilidades funcionales obligatorias",
        "habilidades sugerida"))]
    unmapped_skills = set()
    for source_index, source in enumerate(all_rows, start=first_excel_row):
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

        source_domain_column = indexes.get(normalize("Dominio"))
        source_domain = (row.get(source_domain_column) or "").strip() if source_domain_column else ""
        mesa_value = normalize(row.get(indexes[normalize("Mesa")], ""))
        focus_value = normalize(row.get(indexes[normalize("Foco estratégico (Fábrica)")], ""))
        group_key = (mesa_value, focus_value)
        sda_column = indexes[normalize("SDATOOL")]
        for sda in re.split(r"[|;,\n]+", (row.get(sda_column) or "").strip()):
            sda = sda.strip()
            sda_key = code_key(sda)
            if sda and sda_key:
                sdas_by_group[group_key].setdefault(sda_key, sda)
        if re.fullmatch(r"mesa\s+\d+", mesa_value) and focus_value:
            mesa_focus_groups_source.add((mesa_value, focus_value))

        role = row.get(indexes[normalize("Rol")], "")
        seniority = seniority_key(row.get(indexes[normalize("Seniority")], ""))
        # Cuenta skills obligatorias técnicas/funcionales distintas. No se usan
        # habilidades sugeridas, Nivel general skills técnicos ni Nivel IA.
        skill_count = len({skill_match_key(row.get(name, "")) for name in mandatory_skill_fields
                           if (row.get(name) or "").strip()})
        family = demand_role_family(role, domain)
        if family is None:
            excluded["rol sin mapeo"] += 1
            if role.strip():
                unmapped_roles.add(role.strip())
                if unmapped_role_rows is not None:
                    unmapped_role_rows.append([domain, role.strip(), source_index])
            continue
        if seniority not in VALID_SENIORITY:
            excluded["seniority vacío o no homologado"] += 1
            continue
        # La fila necesita una skill obligatoria, salvo QA. El orden de
        # preferencia por dos, una o cero coincidencias se resuelve en el MILP.
        if skill_count == 0 and family != "QA":
            excluded["sin skills obligatorias (solo se admite para QA)"] += 1
            continue

        missing_context = [field for field in REQUIRED_FILLED_FIELDS
                           if not (row.get(indexes[normalize(field)]) or "").strip()]
        if missing_context:
            excluded["campos de contexto vacíos"] += 1
            continue
        # Convierte las etiquetas usadas para cruces a la misma forma canónica:
        # rol por familia mapeada, seniority por catálogo y skills con el nombre
        # exacto de la columna de SkillsCOE cuando existe una coincidencia.
        row[indexes[normalize("Rol")]] = canonical_demand_role(role, domain)
        row[indexes[normalize("Seniority")]] = SENIORITY_LABELS[seniority]
        for skill_field in all_initiative_skill_fields:
            skill_value = (row.get(skill_field) or "").strip()
            if not skill_value:
                continue
            canonical_skill = canonical_skills.get(skill_match_key(skill_value))
            if canonical_skill:
                row[skill_field] = canonical_skill
            else:
                row[skill_field] = normalized_unknown_skill(skill_value)
                unmapped_skills.add(row[skill_field])
        row["Dominio fuente"] = source_domain
        if source_domain_column:
            row[source_domain_column] = domain
        else:
            row["Dominio"] = domain
        row["Fila origen"] = str(source_index)
        selected.append(row)

    # Cada demanda conserva su fila, pero SDATOOL representa todas las SDA
    # relacionadas con su grupo Mesa + Foco estratégico.
    sda_column = indexes[normalize("SDATOOL")]
    for row in selected:
        group_key = (
            normalize(row.get(indexes[normalize("Mesa")], "")),
            normalize(row.get(indexes[normalize("Foco estratégico (Fábrica)")], "")),
        )
        row[sda_column] = " | ".join(sdas_by_group[group_key].values())

    output_header = list(header)
    if normalize("Dominio") not in indexes:
        output_header.append("Dominio")
    if normalize("Dominio fuente") not in indexes:
        output_header.append("Dominio fuente")
    output_header.append("Fila origen")

    output_path = initiative_output_path(domain)
    with output_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=output_header)
        writer.writeheader()
        writer.writerows(selected)

    print(f"{domain}: filas revisadas {source_rows}; demandas conservadas: {len(selected)}; "
          f"exclusiones: {dict(excluded)}; salida: {output_path}.")
    eligible_groups = {
        (normalize(row.get(indexes[normalize("Mesa")], "")),
         normalize(row.get(indexes[normalize("Foco estratégico (Fábrica)")], "")))
        for row in selected
    }
    print(f"{domain}: grupos Mesa + Foco estratégico: {len(mesa_focus_groups_source)} "
          f"en la matriz; {len(eligible_groups)} con demanda elegible.")
    print(f"{domain}: SDA consolidadas por Mesa + Foco estratégico en la columna SDATOOL.")
    if unmapped_roles:
        print(f"{domain}: roles que requieren mapeo en mapeo_roles.py: {sorted(unmapped_roles)}")
    if unmapped_skills:
        print(f"{domain}: skills de iniciativas no encontradas como columna en SkillsCOE: "
              f"{sorted(unmapped_skills)}")


def prepare_initiatives():
    """Limpia ENG desde CSV y los demás dominios desde hojas de Iniciativas.xlsx.

    Para agregar un dominio basta con incorporar una pestaña llamada, por
    ejemplo, "FIN X Mesa" o "FIN X Mesas". El texto anterior a "X Mesa(s)" se
    escribe como Dominio en todas las filas de salida de esa pestaña.
    """
    unmapped_role_rows = []
    eng_header, eng_rows = read_csv(ENG_INITIATIVES_FILE, INITIATIVES_HEADER_ROW - 1)
    prepare_initiative_file("ENG", eng_header, eng_rows,
                            unmapped_role_rows=unmapped_role_rows)

    xlsx_ns = {
        "main": STAFFING_XLSX_NS["main"],
        "pkg": STAFFING_XLSX_NS["pkg"],
    }
    sheet_pattern = re.compile(INITIATIVE_DOMAIN_SHEET_PATTERN, re.IGNORECASE)
    processed_domains = set()
    with zipfile.ZipFile(INITIATIVES_WORKBOOK) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {item.attrib["Id"]: item.attrib["Target"]
                   for item in relationships.findall("pkg:Relationship", xlsx_ns)}
        shared_strings = read_shared_strings(archive)

        for sheet in workbook.findall("main:sheets/main:sheet", xlsx_ns):
            sheet_name = sheet.attrib["name"].strip()
            match = sheet_pattern.fullmatch(sheet_name)
            if not match:
                # ENG se procesa desde su CSV dedicado; las pestañas auxiliares
                # del libro que no sigan el patrón no son dominios de iniciativas.
                continue
            domain = normalize(match.group(1)).upper()
            if domain == "ENG":
                continue

            relationship_id = sheet.attrib[f"{{{STAFFING_XLSX_NS['rel']}}}id"]
            matrix = list(worksheet_rows(archive, targets[relationship_id], shared_strings))
            header, rows = matrix_records(matrix, INITIATIVES_HEADER_ROW - 1)
            prepare_initiative_file(domain, header, rows,
                                    unmapped_role_rows=unmapped_role_rows)
            processed_domains.add(domain)

    if not processed_domains:
        raise ValueError(
            f"No se encontraron pestañas de dominio en {INITIATIVES_WORKBOOK.name}; "
            "usa nombres como 'GRM X Mesa' o 'GRM X Mesas'.")
    write_review_workbook(UNMAPPED_ROLES_OUTPUT, [
        ("Roles iniciativas sin mapeo", ["Dominio", "Rol fuente", "Fila origen"],
         unmapped_role_rows),
    ])
    print(f"Roles de iniciativas sin mapeo: {len(unmapped_role_rows)}; "
          f"Excel de revisión: {UNMAPPED_ROLES_OUTPUT}")


def excel_column(number):
    """Convierte un índice de columna Excel (1-based) a letras: 1 -> A."""
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def write_review_workbook(path, sheets):
    """Escribe un XLSX mínimo con hojas filtrables usando solo la biblioteca estándar.

    Cada elemento de sheets contiene nombre de hoja, encabezados y filas. El
    formato se genera como un paquete XML XLSX para evitar dependencias externas.
    """
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
    """Genera dos listas de códigos: solo en SkillsCOE y solo en Staffing."""
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
    """Ejecuta en orden la extracción de Staffing y los tres pasos de limpieza."""
    RESULTS.mkdir(exist_ok=True)
    # Mantiene Staffing Q4 al día desde el Excel antes de hacer el cruce.
    extract_q4()
    prepare_people()
    prepare_initiatives()
    prepare_unmapped_skills_report()
    prepare_unmatched_report()
    print(f"Salidas MILP: matriz de personas {PEOPLE_OUTPUT}; "
          f"matrices de iniciativas por dominio en {RESULTS}.")


if __name__ == "__main__":
    main()
