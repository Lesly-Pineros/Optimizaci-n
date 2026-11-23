#!/usr/bin/env python3
"""Asigna personas a plazas de iniciativas con un MILP (SciPy/HiGHS).

Las filas con una o más skills obligatorias generan plazas individuales según
Internal/External. Rol, seniority, tipo y dominio son requisitos; más skills
obligatorias cubiertas aumentan la puntuación del candidato.
"""
import csv
import json
import re
import unicodedata
from datetime import date, datetime, timedelta
from collections import defaultdict, Counter
from pathlib import Path
from salidas_asignacion import export_tobe
from mapeo_roles import demand_role_family, person_role_families

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "DATA"
INPUTS = ROOT / "RESULTADOS_LIMPIEZA"
RESULTS = ROOT / "RESULTADOS MODELOS"
PEOPLE_FILE = INPUTS / "Matriz SkillsCOE enriquecida.csv"
PROJECT_FILE = INPUTS / "GRM-iniciativas con skills.csv"
STAFFING_FILE = DATA / "STAFFING.csv"
ASSIGNMENTS_FILE = RESULTS / "Asignaciones GRM MILP.csv"
UNFILLED_FILE = RESULTS / "Plazas GRM sin asignar MILP.csv"
ROLE_MATCH_FILE = RESULTS / "Similitud de roles GRM.csv"
EXCEL_FILE = RESULTS / "Asignacion GRM To-Be_Baseline.xlsx"

MAX_BASE_ASSIGNMENTS_PER_GROUP = 7
MAX_SUGGESTED_PER_MESA = 5
MAX_SIMULTANEOUS_PROJECTS = 2
MIN_SKILLS_PER_ROLE = 1
NO_SKILL_MATCH_PENALTY = 10
SENIORITY_ORDER = {"junior": 1, "mid level": 2, "senior": 3, "lead": 4}


def seniority_fit(project_seniority, person_seniority):
    """Admite exacto, un nivel superior y finalmente un nivel inferior."""
    requested = SENIORITY_ORDER.get(norm(project_seniority))
    actual = SENIORITY_ORDER.get(norm(person_seniority))
    if requested is None or actual is None or abs(actual - requested) > 1:
        return None
    # Un valor mayor expresa preferencia lexicográfica: exacto > superior > inferior.
    return 2 if actual == requested else (1 if actual > requested else 0)
# Mapeo de Dirección a Dominio cuando la fuente no proporciona el dominio.
LEADER_ROLE_MARKERS = ("lider", "lead", "tech manager")
TESTING_ROLE_MARKERS = ("qa", "test", "testing", "tester", "quality")
INVALID_ROLE_MARKERS = ("na ", "n a ", "cod eliminado", "codigo eliminado", "retirado")
HORIZON_START = date(2027, 1, 1)
HORIZON_END = date(2027, 6, 30)

FILL_DOWN_PROJECT_FIELDS = [
    "Mesa", "Foco estratégico (Fábrica)", "SDATOOL", "Nombre Proyecto",
    "Fecha Inicio", "Fecha Fin", "Dominio", "Estado", "Workspace",
]
TECH_FIELDS = [f"Habilidades técnicas obligatorias {i}" for i in range(1, 6)]
FUNCTIONAL_FIELDS = [f"Habilidades funcionales obligatorias {i}" for i in range(1, 6)]
SUGGESTED_FIELDS = [f"Habilidades sugerida {i}" for i in range(1, 4)]


def norm(value):
    value = unicodedata.normalize("NFKD", str(value or ""))
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", " ", value.casefold()).strip()


def key(value):
    return re.sub(r"[^A-Z0-9]", "", unicodedata.normalize("NFKC", str(value or "")).upper())


def read_csv(path, header_row=0):
    last_error = None
    for encoding in ("utf-8-sig", "cp1252", "latin1"):
        try:
            with path.open(encoding=encoding, newline="") as f:
                rows = list(csv.reader(f))
            header = rows[header_row]
            return header, [dict(zip(header, row + [""] * max(0, len(header) - len(row))))
                            for row in rows[header_row + 1:]]
        except UnicodeDecodeError as exc:
            last_error = exc
    raise last_error


def level(value):
    match = re.match(r"\s*([0-4])", str(value or ""))
    return int(match.group(1)) if match else 0


def parse_date(value, *, quarter_end=False):
    value = str(value or "").strip()
    if not value:
        return None
    quarter = re.search(r"(?i)([1-4])\s*Q\s*(\d{4})|Q\s*([1-4])\s*(\d{4})", value)
    if quarter:
        q = int(quarter.group(1) or quarter.group(3))
        year = int(quarter.group(2) or quarter.group(4))
        month = q * 3 if quarter_end else (q - 1) * 3 + 1
        if quarter_end:
            next_month = month + 1
            next_year = year
            if next_month == 13:
                next_month, next_year = 1, year + 1
            return date(next_year, next_month, 1) - timedelta(days=1)
        return date(year, month, 1)
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(value[:10], fmt).date()
        except ValueError:
            pass
    try:
        serial = float(value)
        if 20000 <= serial <= 80000:
            return (datetime(1899, 12, 30) + timedelta(days=serial)).date()
    except ValueError:
        pass
    return None


def quarter_label(value, *, quarter_end=False):
    parsed = parse_date(value, quarter_end=quarter_end)
    if not parsed:
        return "trimestre no informado"
    return f"{(parsed.month - 1) // 3 + 1}Q {parsed.year}"


def staffing_project_periods(person):
    assignments = person.get("staffing_assignments") or []
    if not assignments:
        current = person.get("current_assignment") or {}
        if any(str(current.get(field) or "").strip() for field in
               ("sda", "proyecto", "mesa", "fecha_inicio", "fecha_fin_actual")):
            assignments = [current]

    periods = []
    for index, item in enumerate(assignments):
        if not isinstance(item, dict):
            continue
        start_raw = item.get("fecha_inicio", "")
        end_raw = item.get("fecha_fin_actual", "")
        start = parse_date(start_raw) or HORIZON_START
        end = parse_date(end_raw, quarter_end=True) or HORIZON_END
        if end < HORIZON_START or start > HORIZON_END:
            continue
        start, end = max(start, HORIZON_START), min(end, HORIZON_END)
        if end < start:
            continue
        sda = key(item.get("sda"))
        project = norm(item.get("proyecto"))
        mesa = norm(item.get("mesa"))
        identity = (sda, project, mesa) if (sda or project or mesa) else ("fila", index)
        periods.append({
            "identity": identity, "start": start, "end": end, "item": item,
            "start_raw": start_raw, "end_raw": end_raw,
        })
    return periods


def staffing_matches_project(assignment, project):
    staffing_sda = key(assignment.get("sda"))
    project_sda = key(project.get("SDATOOL"))
    if staffing_sda and project_sda and staffing_sda == project_sda:
        return True
    staffing_name = norm(assignment.get("proyecto"))
    project_name = norm(project.get("Nombre Proyecto"))
    staffing_mesa = norm(assignment.get("mesa"))
    project_mesa = norm(project.get("Mesa"))
    return bool(
        staffing_name and project_name and staffing_name == project_name
        and (not staffing_mesa or not project_mesa or staffing_mesa == project_mesa)
    )


def staffing_capacity_allows_project(person, project):
    start = max(parse_date(project.get("Fecha Inicio")) or HORIZON_START, HORIZON_START)
    end = min(parse_date(project.get("Fecha Fin"), quarter_end=True) or HORIZON_END,
              HORIZON_END)
    periods = [period for period in staffing_project_periods(person)
               if period["start"] <= end and period["end"] >= start]
    boundaries = {start}
    for period in periods:
        boundaries.add(max(start, period["start"]))
        if period["end"] < end:
            boundaries.add(period["end"] + timedelta(days=1))
    for point in boundaries:
        active = [period for period in periods if period["start"] <= point <= period["end"]]
        active_ids = {period["identity"] for period in active}
        already_on_project = any(staffing_matches_project(period["item"], project)
                                 for period in active)
        if len(active_ids) > MAX_SIMULTANEOUS_PROJECTS:
            return False
        if not already_on_project and len(active_ids) >= MAX_SIMULTANEOUS_PROJECTS:
            return False
    return True


def staffing_overlap_comment(person, project):
    project_start = max(parse_date(project.get("Fecha Inicio")) or HORIZON_START,
                        HORIZON_START)
    project_end = min(parse_date(project.get("Fecha Fin"), quarter_end=True) or HORIZON_END,
                      HORIZON_END)
    comments = []
    seen = set()
    for period in staffing_project_periods(person):
        if max(project_start, period["start"]) > min(project_end, period["end"]):
            continue
        item = period["item"]
        label = (item.get("proyecto") or item.get("sda") or "Proyecto sin nombre").strip()
        mesa = (item.get("mesa") or "").strip()
        identity = (period["identity"], period["start"], period["end"])
        if identity in seen:
            continue
        seen.add(identity)
        if mesa:
            label += f" / {mesa}"
        start_q = quarter_label(period["start_raw"])
        end_q = quarter_label(period["end_raw"], quarter_end=True)
        comments.append(f"Asignación vigente: {label} ({start_q}–{end_q})")
    return "; ".join(comments)


def dedication_fraction(value):
    value = str(value or "").strip().replace("%", "").replace(",", ".")
    try:
        parsed = float(value)
    except ValueError:
        return 0.0
    return parsed / 100.0 if parsed > 1 else parsed


def role_families(value, seniority=None):
    return person_role_families(value, seniority)


def demand_family(value, domain=None):
    return demand_role_family(value, domain)


def qa_without_skills(project, role_map=None):
    role = (role_map or {}).get(norm(project.get("Rol")),
            demand_family(project.get("Rol"), project.get("Dominio")))
    no_skills = not any(
        (project.get(field) or "").strip()
        for field in TECH_FIELDS + FUNCTIONAL_FIELDS + SUGGESTED_FIELDS
    )
    return role == "QA" and no_skills


def skill_match_penalty(project, person, role_map=None):
    if qa_without_skills(project, role_map):
        return NO_SKILL_MATCH_PENALTY
    if (project.get("_allow_no_skill_fallback")
            and skill_fit(project, person, role_map)[0] == 0):
        return NO_SKILL_MATCH_PENALTY
    return 0


def role_compatible(role, person):
    """Aplica las familias GRM y mantiene a Dev Int como rol interno estricto."""
    families = person.get("families") or role_families(
        person.get("role_raw", person.get("role", "")), person.get("seniority_raw"))
    if role == "DEV_INT":
        return "DEV" in families and person.get("resource_type") == "interno"
    if role == "DEV_EXT":
        return "DEV" in families and person.get("resource_type") in {"interno", "externo"}
    return bool(role and role in families)


def current_staffing_assignment(row):
    raw = row.get("Asignaciones Staffing Q4", "")
    try:
        assignments = json.loads(raw) if raw else []
    except json.JSONDecodeError:
        assignments = []
    if not isinstance(assignments, list):
        assignments = []
    if not assignments:
        assignments = [{
            "sda": row.get("SDATOOL - Proyecto", ""),
            "proyecto": row.get("Proyecto Staffing", ""),
            "mesa": row.get("Mesa Staffing actual", ""),
            "fecha_inicio": row.get("Fecha Inicio", ""),
            "fecha_fin_actual": row.get("Fecha Fin Actual", ""),
            "dedicacion": row.get("% Dedicación Staffing", ""),
        }]
    return max(assignments, key=lambda item: dedication_fraction(item.get("dedicacion")), default={})


def load_projects():
    header, rows = read_csv(PROJECT_FILE)
    filled = []
    previous = {name: "" for name in FILL_DOWN_PROJECT_FIELDS}
    for source in rows:
        row = dict(source)
        for field in FILL_DOWN_PROJECT_FIELDS:
            value = (row.get(field) or "").strip()
            if value:
                previous[field] = value
            elif previous[field]:
                row[field] = previous[field]
        filled.append(row)
    retained = []
    for row in filled:
        row["Dominio"] = "GRM"
        skill_count = len(set(required_skill_names(row, TECH_FIELDS + FUNCTIONAL_FIELDS)))
        role_defined = demand_family(row.get("Rol"), row.get("Dominio")) is not None
        family = demand_family(row.get("Rol"), row.get("Dominio"))
        no_skills_diligenced = qa_without_skills(row)
        seniority_defined = norm(row.get("Seniority")) in SENIORITY_ORDER
        skills_rule_ok = (skill_count >= MIN_SKILLS_PER_ROLE
                          or (family == "QA" and no_skills_diligenced))
        if role_defined and seniority_defined and skills_rule_ok:
            retained.append(row)
    print(f"Iniciativas GRM filtradas: {len(retained)} conservadas de {len(filled)} "
          "(GRM, rol y seniority válidos; QA sin skills también se admite).")
    return header, retained


def load_people():
    _, people = read_csv(PEOPLE_FILE)
    result = []
    seen = set()
    for row in people:
        person_key = key(row.get("Código"))
        if not person_key or person_key in seen:
            continue
        seen.add(person_key)
        direction = (row.get("Dominio") or "").strip()
        resource_type = norm(row.get("Tipo de Recurso"))
        if resource_type == "subsidiaria":
            resource_type = "externo"
        current = current_staffing_assignment(row)
        assignments_raw = row.get("Asignaciones Staffing Q4", "")
        try:
            assignments = json.loads(assignments_raw) if assignments_raw else []
        except json.JSONDecodeError:
            assignments = []
        skills = {norm(name): level(value) for name, value in row.items() if name.strip()}
        seniority_raw = (row.get("Seniority") or row.get("Nivel del Rol Actual") or "").strip()
        result.append({
            "key": person_key,
            "code": (row.get("Código") or "").strip(),
            "name": (row.get("Nombre") or "").strip(),
            "role": norm(row.get("Rol Principal")),
            "role_raw": (row.get("Rol Principal") or "").strip(),
            "seniority": norm(seniority_raw),
            "seniority_raw": seniority_raw,
            "company": norm(row.get("Empresa")),
            "resource_type": {"interno": "interno", "externo": "externo"}.get(resource_type, ""),
            "direction": direction,
            "domains": {norm(value) for value in direction.split("|") if norm(value)},
            # El SDA seleccionado proviene de SDATOOL - Proyecto, emparejado con
            # el registro de mayor dedicación en el detalle de Staffing enriquecido.
            "staffing_sda": current.get("sda") or row.get("SDATOOL - Proyecto", ""),
            "families": role_families(row.get("Rol Principal"), seniority_raw),
            "staffing_status": row.get("Estado Staffing", ""),
            "staffing_assignments": assignments,
            "current_assignment": current,
            "skills": skills,
        })
    return result


def build_role_map(projects, people):
    mapping, report = {}, []
    role_domains = sorted({((p.get("Rol") or "").strip(),
                            (p.get("Dominio") or "").strip())
                           for p in projects if (p.get("Rol") or "").strip()})
    for role, domain in role_domains:
        source = norm(role)
        family = demand_family(role, domain)
        mapping[source] = family
        report.append({
            "Rol iniciativa": role,
            "Dominio": domain,
            "Familia de rol": family,
            "Método": "mapeo a familia",
            "Usar en asignación": "Sí" if family else "No",
        })
    return mapping, report


def required_skill_names(row, fields):
    return [norm(row.get(field)) for field in fields if (row.get(field) or "").strip()]


def numeric_tokens(value):
    """Identificadores numéricos completos de Mesa/SDA (no compara dígitos parciales)."""
    return set(re.findall(r"\d+", norm(value)))


def labels_match_by_text_or_number(left, right):
    left_norm, right_norm = norm(left), norm(right)
    if not left_norm or not right_norm:
        return False
    left_parts = {norm(part) for part in str(left or "").split("|") if norm(part)}
    right_parts = {norm(part) for part in str(right or "").split("|") if norm(part)}
    if left_parts & right_parts:
        return True
    return bool(numeric_tokens(left_norm) & numeric_tokens(right_norm))


def skill_fit(project, person, role_map=None):
    """Cuenta skills obligatorias registradas, sin filtrar por nivel."""
    required = required_skill_names(project, TECH_FIELDS + FUNCTIONAL_FIELDS)
    levels = [person["skills"].get(skill, 0) for skill in required]
    covered = sum(value > 0 for value in levels)
    known = sum(value > 0 for value in levels)
    return covered, len(required), known, None, covered


def skill_observation(project, person, role_map=None):
    """Explica el respaldo sin match o el caso de skill LT en nivel Básico."""
    if skill_match_penalty(project, person, role_map):
        return "Sin match de skills; se aplicó respaldo con penalización de -10 puntos."
    role = (role_map or {}).get(norm(project.get("Rol")),
            demand_family(project.get("Rol"), project.get("Dominio")))
    if role != "LT":
        return ""
    required = required_skill_names(project, TECH_FIELDS + FUNCTIONAL_FIELDS)
    if any(person["skills"].get(skill, 0) == 1 for skill in required):
        return "LT: cumple al menos una skill obligatoria, pero en nivel Básico (1)."
    return ""


def affinity_components(project, person):
    """Return the continuity and Domain preference components."""
    current = person.get("current_assignment", {})
    staffing_sda = person.get("staffing_sda", current.get("sda", ""))
    current_project = norm(current.get("proyecto"))
    current_mesa = norm(current.get("mesa"))
    initiative_sda = project.get("SDATOOL", "")
    project_name = norm(project.get("Nombre Proyecto"))
    project_mesa = norm(project.get("Mesa"))
    continuity = int(bool(
        labels_match_by_text_or_number(current_mesa, project_mesa)
        or labels_match_by_text_or_number(staffing_sda, initiative_sda)
        or (current_project and project_name and current_project == project_name)
    ))
    domain = norm(project.get("Dominio", ""))
    domains = person.get("domains") or {
        norm(value) for value in str(person.get("direction", "")).split("|") if norm(value)
    }
    domain_match = int(bool(domain and domain in domains))
    return continuity, domain_match


def affinity_score(project, person, role_map=None):
    continuity, domain_match = affinity_components(project, person)
    matched_skills = skill_fit(project, person, role_map)[0]
    return continuity * 1000 + domain_match * 100 + matched_skills


def combined_preference_score(project, person, number_of_slots, role_map=None):
    """Ordena por continuidad, Dominio y finalmente skills obligatorias cubiertas."""
    continuity, domain_match = affinity_components(project, person)
    matched_skills = skill_fit(project, person, role_map)[0]
    max_skills = len(TECH_FIELDS) + len(FUNCTIONAL_FIELDS)
    skill_total_ceiling = number_of_slots * max_skills
    domain_weight = skill_total_ceiling + 1
    continuity_weight = number_of_slots * (domain_weight + max_skills) + 1
    return continuity * continuity_weight + domain_match * domain_weight + matched_skills


def eligibility(project, person, skill_columns, role_map, allow_skillless=False):
    mapped_role = role_map.get(norm(project.get("Rol")))
    role = mapped_role if mapped_role in {"LT", "DEV", "QA", "DEV_INT", "DEV_EXT"} else demand_family(
        project.get("Rol"), project.get("Dominio"))
    seniority = norm(project.get("Seniority"))
    if not role or not role_compatible(role, person):
        return False, 0, "rol"
    if seniority_fit(seniority, person["seniority"]) is None:
        return False, 0, "seniority"

    required_type = project.get("_required_type", "")
    if required_type not in {"interno", "externo"}:
        return False, 0, "tipo interno/externo"

    person_domains = person.get("domains") or {
        norm(value) for value in str(person.get("direction", "")).split("|") if norm(value)
    }
    required_domain = norm(project.get("Dominio"))
    if not required_domain or required_domain not in person_domains:
        return False, 0, "dominio distinto de " + (project.get("Dominio") or "la iniciativa")

    if role == "LT" and (required_type != "interno" or person.get("resource_type") != "interno"):
        return False, 0, "LT solo BBVA y sin sustitución de tipo"

    required = required_skill_names(project, TECH_FIELDS + FUNCTIONAL_FIELDS)
    skillless_fallback = (allow_skillless or project.get("_allow_no_skill_fallback")
                          or qa_without_skills(project, role_map))
    if len(required) < MIN_SKILLS_PER_ROLE and not skillless_fallback:
        return False, 0, "sin skills obligatorias en la demanda"
    covered, _, _, _, _ = skill_fit(project, person, role_map)
    if covered < MIN_SKILLS_PER_ROLE and not skillless_fallback:
        return False, 0, "ninguna skill obligatoria registrada"

    display_score = affinity_score(project, dict(person, domains=person_domains), role_map)
    display_score -= skill_match_penalty(project, person, role_map)
    return True, display_score, ""


def prepare_skill_fallbacks(slots, people, role_map):
    """Enciende respaldo de skills solo si ningún candidato elegible tiene match."""
    for _, project in slots:
        if qa_without_skills(project, role_map):
            project["_allow_no_skill_fallback"] = True
            continue
        required = required_skill_names(project, TECH_FIELDS + FUNCTIONAL_FIELDS)
        if not required:
            project["_allow_no_skill_fallback"] = False
            continue
        has_skill_candidate = any(
            eligibility(project, person, set(), role_map, allow_skillless=True)[0]
            and staffing_capacity_allows_project(person, project)
            and skill_fit(project, person, role_map)[0] >= MIN_SKILLS_PER_ROLE
            for person in people
        )
        project["_allow_no_skill_fallback"] = not has_skill_candidate


def diagnose_no_candidate(project, people, skill_columns, role_map):
    """Devuelve el primer criterio que deja la plaza sin candidatos."""
    pool = people
    mapped = role_map.get(norm(project.get("Rol")))
    role = mapped if mapped in {"LT", "DEV", "QA", "DEV_INT", "DEV_EXT"} else demand_family(
        project.get("Rol"), project.get("Dominio"))
    pool = [p for p in pool if role_compatible(role, p)]
    if not pool:
        return "No hay persona con rol compatible"
    seniority = norm(project.get("Seniority"))
    pool = [p for p in pool if seniority_fit(seniority, p["seniority"]) is not None]
    if not pool:
        return "No hay persona con seniority compatible"
    required_type = project.get("_required_type", "")
    if role == "LT":
        pool = [p for p in pool if p["resource_type"] == "interno"
                and required_type == "interno"]
        if not pool:
            return "No hay LT interno elegible; LT no admite sustitución de tipo"
    else:
        pool = [p for p in pool if p["resource_type"] in {"interno", "externo"}]
        if not pool:
            return "No hay persona interna ni externa elegible"
    required_domain = norm(project.get("Dominio"))
    pool = [p for p in pool if required_domain and required_domain in (p.get("domains") or {
        norm(value) for value in str(p.get("direction", "")).split("|") if norm(value)
    })]
    if not pool:
        return "No hay persona con Dominio " + (project.get("Dominio") or "de la iniciativa")
    if role == "LT":
        pool = [p for p in pool if p["resource_type"] == "interno"]
        if not pool:
            return "No hay LT elegible que sea del banco"
    skill_failures = Counter()
    for person in pool:
        ok, _, reason = eligibility(project, person, skill_columns, role_map)
        if not ok:
            skill_failures[reason] += 1
    if skill_failures:
        reason = skill_failures.most_common(1)[0][0]
        return {
            "sin skills obligatorias en la demanda": "La vacante no declara skills obligatorias",
            "ninguna skill obligatoria registrada": "Ninguna persona tiene una skill obligatoria registrada",
            "no liberado para el inicio del horizonte de asignación": (
                "No hay persona liberada para el inicio del horizonte de asignación"),
            "LT solo BBVA y sin sustitución de tipo": (
                "No hay LT interno de BBVA elegible"),
        }.get(reason, "No hay persona con skills coincidentes")
    return "No hay persona elegible"


def make_slots(projects):
    slots = []
    for index, project in enumerate(projects):
        if not (project.get("Rol") or "").strip() or not (project.get("Seniority") or "").strip():
            continue
        start = parse_date(project.get("Fecha Inicio"))
        end = parse_date(project.get("Fecha Fin"), quarter_end=True)
        if start and start > HORIZON_END or end and end < HORIZON_START:
            continue
        categories = []
        for field, category in (("Internal", "interno"), ("External", "externo")):
            raw = (project.get(field) or "").strip()
            try:
                count = int(float(raw)) if raw else 0
            except ValueError:
                count = 0
            categories.extend([category] * max(0, count))
        for number, category in enumerate(categories, start=1):
            demand = dict(project)
            demand["_required_type"] = category
            demand["_slot_number"] = str(number)
            source_row = (project.get("Fila origen") or "").strip()
            try:
                source_row = int(source_row) if source_row else index + 2
            except ValueError:
                source_row = index + 2
            slots.append((source_row, demand))
    return slots


def project_key(project):
    return (key(project.get("SDATOOL")), norm(project.get("Nombre Proyecto")),
            norm(project.get("Mesa")), project.get("Fecha Inicio", ""), project.get("Fecha Fin", ""))


def mesa_group_key(project):
    """Clave interna de conteo; no se agrega como columna de salida."""
    return (norm(project.get("Mesa")), norm(project.get("Foco estratégico (Fábrica)")))


def intervals_overlap(project_a, project_b):
    start_a = parse_date(project_a.get("Fecha Inicio")) or HORIZON_START
    end_a = parse_date(project_a.get("Fecha Fin"), quarter_end=True) or HORIZON_END
    start_b = parse_date(project_b.get("Fecha Inicio")) or HORIZON_START
    end_b = parse_date(project_b.get("Fecha Fin"), quarter_end=True) or HORIZON_END
    start_a, start_b = max(start_a, HORIZON_START), max(start_b, HORIZON_START)
    end_a, end_b = min(end_a, HORIZON_END), min(end_b, HORIZON_END)
    return max(start_a, start_b) <= min(end_a, end_b)


def solve_slots(slots, people, role_map, viability_matrix=None, affinity_matrix=None,
                recommendation_pool=None, fixed_assignments=None):
    import numpy as np
    from scipy.optimize import Bounds, LinearConstraint, milp
    from scipy.sparse import coo_matrix, csr_matrix, vstack

    if viability_matrix is not None and viability_matrix.shape != (len(slots), len(people)):
        raise ValueError("La matriz V no coincide con las dimensiones de plazas y personas")
    if affinity_matrix is not None and affinity_matrix.shape != (len(slots), len(people)):
        raise ValueError("La matriz A no coincide con las dimensiones de plazas y personas")

    prepare_skill_fallbacks(slots, people, role_map)

    pairs = []
    projects_by_key = {project_key(project): project for _, project in slots}
    restrict_to_fixed_and_suggested = fixed_assignments is not None
    fixed_assignments = set(fixed_assignments or [])
    reasons = defaultdict(int)
    eligible_by_slot = defaultdict(list)
    for slot_index, (source_row, project) in enumerate(slots):
        project_key_value = project_key(project)
        mesa_key = mesa_group_key(project)
        if not any(mesa_key):
            reasons[(slot_index, "mesa")]
            continue
        for person_index, person in enumerate(people):
            if viability_matrix is not None and not viability_matrix[slot_index, person_index]:
                continue
            group_key = mesa_group_key(project)
            if restrict_to_fixed_and_suggested and ((slot_index, person_index) not in fixed_assignments
                                                    and person_index not in (recommendation_pool or {}).get(group_key, [])):
                continue
            ok, score, reason = eligibility(project, person, set(), role_map)
            if viability_matrix is not None and not ok:
                raise ValueError(f"La matriz V marca viable un par que no pasa los filtros: slot {slot_index}, persona {person_index}")
            if ok:
                pairs.append((slot_index, person_index, score, project_key_value, mesa_key))
                eligible_by_slot[slot_index].append(person_index)
            else:
                reasons[(slot_index, reason)] += 1

    if not pairs:
        return slots, [], reasons, eligible_by_slot

    n_x = len(pairs)
    mesa_person_keys = sorted({(person, mesa) for _, person, _, _, mesa in pairs})
    mesa_person_index = {item: n_x + i for i, item in enumerate(mesa_person_keys)}
    project_person_keys = sorted({(person, project_id) for _, person, _, project_id, _ in pairs})
    project_person_start = n_x + len(mesa_person_keys)
    project_person_index = {item: project_person_start + i
                            for i, item in enumerate(project_person_keys)}
    mesa_active_start = project_person_start + len(project_person_keys)
    mesa_active_keys = sorted({mesa_group_key(project) for _, project in slots
                               if any(mesa_group_key(project))})
    mesa_active_index = {mesa: mesa_active_start + i for i, mesa in enumerate(mesa_active_keys)}
    mesa_type_active_start = mesa_active_start + len(mesa_active_keys)
    mesa_type_keys = sorted({(mesa, people[person]["resource_type"])
                             for _, person, _, _, mesa in pairs})
    mesa_type_active_index = {item: mesa_type_active_start + i
                              for i, item in enumerate(mesa_type_keys)}
    mesa_mixed_start = mesa_type_active_start + len(mesa_type_keys)
    mesa_mixed_keys = sorted({mesa for mesa in mesa_active_keys
                              if (mesa, "interno") in mesa_type_active_index
                              and (mesa, "externo") in mesa_type_active_index})
    mesa_mixed_index = {mesa: mesa_mixed_start + i
                        for i, mesa in enumerate(mesa_mixed_keys)}
    n_vars = mesa_mixed_start + len(mesa_mixed_keys)
    by_slot, by_person_project, by_mesa_person = defaultdict(list), defaultdict(list), defaultdict(list)
    by_mesa_type = defaultdict(lambda: defaultdict(list))
    rows, lower, upper = [], [], []
    def add_constraint(coefficients, lo=-np.inf, hi=np.inf):
        rows.append(coefficients); lower.append(float(lo)); upper.append(float(hi))

    for j, (slot, person, score, project_id, mesa) in enumerate(pairs):
        by_slot[slot].append(j)
        by_person_project[(person, project_id)].append(j)
        by_mesa_person[mesa].append(mesa_person_index[(person, mesa)])
        z = project_person_index[(person, project_id)]
        y = mesa_person_index[(person, mesa)]
        add_constraint({j: 1, z: -1}, hi=0)
        add_constraint({j: 1, y: -1}, hi=0)
        add_constraint({j: 1, mesa_active_index[mesa]: -1}, hi=0)
        resource_type = people[person]["resource_type"]
        add_constraint({j: 1, mesa_type_active_index[(mesa, resource_type)]: -1}, hi=0)
        by_mesa_type[mesa][resource_type].append(j)

    for variables in by_slot.values():
        add_constraint({v: 1 for v in variables}, hi=1)
    pair_index = {(slot, person): j for j, (slot, person, *_rest) in enumerate(pairs)}
    for fixed_pair in fixed_assignments:
        if fixed_pair not in pair_index:
            raise ValueError(f"Asignación base fija no es viable en la expansión: {fixed_pair}")
        add_constraint({pair_index[fixed_pair]: 1}, lo=1, hi=1)
    for variables in by_person_project.values():
        add_constraint({v: 1 for v in variables}, hi=1)
    recommendation_pool = recommendation_pool or {}
    for mesa in by_mesa_person:
        suggested_people = set(recommendation_pool.get(mesa, []))
        fixed_variables = [mesa_person_index[(person, group)]
                           for person, group in mesa_person_keys
                           if group == mesa and person not in suggested_people]
        suggested_variables = [mesa_person_index[(person, group)]
                               for person, group in mesa_person_keys
                               if group == mesa and person in suggested_people]
        add_constraint({v: 1 for v in set(fixed_variables)}, hi=MAX_BASE_ASSIGNMENTS_PER_GROUP)
        add_constraint({v: 1 for v in set(suggested_variables)}, hi=MAX_SUGGESTED_PER_MESA)

    # Staffing actual consume una plaza de simultaneidad; la capacidad teórica
    # sigue siendo 100% y se permiten hasta dos proyectos activos por persona.
    for person in range(len(people)):
        project_keys = [project_id for p, project_id in project_person_keys if p == person]
        staffing_periods = staffing_project_periods(people[person])
        boundaries = set()
        for project_id in project_keys:
            project = projects_by_key[project_id]
            start = max(parse_date(project.get("Fecha Inicio")) or HORIZON_START, HORIZON_START)
            end = min(parse_date(project.get("Fecha Fin"), quarter_end=True) or HORIZON_END, HORIZON_END)
            boundaries.add(start)
            if end < HORIZON_END:
                boundaries.add(end + timedelta(days=1))
        for period in staffing_periods:
            boundaries.add(period["start"])
            if period["end"] < HORIZON_END:
                boundaries.add(period["end"] + timedelta(days=1))
        for point in boundaries:
            active_project_ids = []
            for project_id in project_keys:
                project = projects_by_key[project_id]
                start = max(parse_date(project.get("Fecha Inicio")) or HORIZON_START, HORIZON_START)
                end = min(parse_date(project.get("Fecha Fin"), quarter_end=True) or HORIZON_END, HORIZON_END)
                if start <= point <= end:
                    active_project_ids.append(project_id)
            active_staffing = [period for period in staffing_periods
                               if period["start"] <= point <= period["end"]]
            active_staffing_ids = {period["identity"] for period in active_staffing}
            already_counted = {
                project_id for project_id in active_project_ids
                if any(staffing_matches_project(period["item"], projects_by_key[project_id])
                       for period in active_staffing)
            }
            new_project_variables = [
                project_person_index[(person, project_id)]
                for project_id in active_project_ids if project_id not in already_counted
            ]
            if new_project_variables:
                available_capacity = max(
                    0, MAX_SIMULTANEOUS_PROJECTS - len(active_staffing_ids))
                add_constraint({z: 1 for z in new_project_variables}, hi=available_capacity)

    # Vincula la composición real de recursos a las asignaciones. La mezcla
    # interna/externa se optimiza como preferencia tras maximizar cobertura, así
    # una falta de un tipo no bloquea plazas que sí puede cubrir el tipo opuesto.
    for mesa, resource_type in mesa_type_keys:
        type_active = mesa_type_active_index[(mesa, resource_type)]
        variables = by_mesa_type[mesa][resource_type]
        add_constraint({type_active: 1, **{v: -1 for v in variables}}, hi=0)
    for mesa in mesa_mixed_keys:
        internal = mesa_type_active_index[(mesa, "interno")]
        external = mesa_type_active_index[(mesa, "externo")]
        mixed = mesa_mixed_index[mesa]
        add_constraint({mixed: 1, internal: -1}, hi=0)
        add_constraint({mixed: 1, external: -1}, hi=0)
        add_constraint({mixed: 1, internal: -1, external: -1}, lo=-1)

    rr, cc, vv = [], [], []
    for row_index, coefficients in enumerate(rows):
        for variable, coefficient in coefficients.items():
            rr.append(row_index); cc.append(variable); vv.append(coefficient)
    matrix = coo_matrix((vv, (rr, cc)), shape=(len(rows), n_vars)).tocsr()
    integrality = np.ones(n_vars)
    bounds = Bounds(np.zeros(n_vars), np.ones(n_vars))
    mesa_demand = Counter(mesa_group_key(project) for _, project in slots if any(mesa_group_key(project)))
    coverage_weights = np.array([1 / max(1, mesa_demand[pair[4]]) for pair in pairs], dtype=float)

    def run(objective, extra_constraints=()):
        matrices = [matrix]
        lows, highs = [np.array(lower)], [np.array(upper)]
        if extra_constraints:
            extra_rows, extra_cols, extra_values = [], [], []
            extra_lows, extra_highs = [], []
            for row_id, (coeff, lo, hi) in enumerate(extra_constraints):
                for col, value in coeff.items():
                    extra_rows.append(row_id); extra_cols.append(col); extra_values.append(value)
                extra_lows.append(lo); extra_highs.append(hi)
            matrices.append(csr_matrix((extra_values, (extra_rows, extra_cols)),
                                       shape=(len(extra_constraints), n_vars)))
            lows.append(np.array(extra_lows)); highs.append(np.array(extra_highs))
            active_matrix = vstack(matrices)
        else:
            active_matrix = matrix
        result = milp(objective, integrality=integrality, bounds=bounds,
                      constraints=LinearConstraint(active_matrix, np.concatenate(lows), np.concatenate(highs)),
                      options={"time_limit": 300})
        if result.x is None or result.status != 0:
            raise RuntimeError(f"Fase MILP sin óptimo certificado (status {result.status}): {result.message}")
        return result

    weighted_objective = np.zeros(n_vars)
    weighted_objective[:n_x] = -coverage_weights
    phase1 = run(weighted_objective)
    best_weighted = round(float(np.dot(coverage_weights, phase1.x[:n_x])), 8)
    weighted_constraint = ({j: float(coverage_weights[j]) for j in range(n_x)}, best_weighted, best_weighted)

    total_objective = np.zeros(n_vars); total_objective[:n_x] = -1
    phase2 = run(total_objective, [weighted_constraint])
    best_total = int(round(float(np.sum(phase2.x[:n_x]))))
    total_constraint = ({j: 1.0 for j in range(n_x)}, best_total, best_total)
    no_skill_fallback_scores = np.array([
        int(bool(slots[pair[0]][1].get("_allow_no_skill_fallback"))) for pair in pairs
    ], dtype=float)
    no_skill_fallback_constraint = []
    if no_skill_fallback_scores.any():
        fallback_objective = np.zeros(n_vars)
        fallback_objective[:n_x] = -no_skill_fallback_scores
        phase_fallback = run(fallback_objective, [weighted_constraint, total_constraint])
        best_fallback = int(round(float(np.dot(no_skill_fallback_scores,
                                                phase_fallback.x[:n_x]))))
        no_skill_fallback_constraint = [
            ({j: float(no_skill_fallback_scores[j]) for j in range(n_x)},
             best_fallback, best_fallback)
        ]
    mix_constraints = []
    if mesa_mixed_keys:
        mix_objective = np.zeros(n_vars)
        mix_objective[[mesa_mixed_index[mesa] for mesa in mesa_mixed_keys]] = -1
        phase_mix = run(mix_objective, [weighted_constraint, total_constraint,
                                        *no_skill_fallback_constraint])
        best_mixed_groups = int(round(sum(phase_mix.x[mesa_mixed_index[mesa]]
                                           for mesa in mesa_mixed_keys)))
        mix_constraints = [({mesa_mixed_index[mesa]: 1.0 for mesa in mesa_mixed_keys},
                            best_mixed_groups, best_mixed_groups)]
    requested_type_scores = np.array([
        int(people[pair[1]]["resource_type"] == slots[pair[0]][1].get("_required_type"))
        for pair in pairs
    ], dtype=float)
    requested_type_objective = np.zeros(n_vars)
    requested_type_objective[:n_x] = -requested_type_scores
    phase_type = run(requested_type_objective,
                     [weighted_constraint, total_constraint,
                      *no_skill_fallback_constraint, *mix_constraints])
    best_requested_type = int(round(float(np.dot(requested_type_scores, phase_type.x[:n_x]))))
    requested_type_constraint = ({j: float(requested_type_scores[j]) for j in range(n_x)},
                                 best_requested_type, best_requested_type)
    seniority_tiers = [seniority_fit(slots[pair[0]][1].get("Seniority"),
                                     people[pair[1]]["seniority"]) for pair in pairs]
    exact_scores = np.array([int(tier == 2) for tier in seniority_tiers], dtype=float)
    above_scores = np.array([int(tier == 1) for tier in seniority_tiers], dtype=float)
    exact_objective = np.zeros(n_vars); exact_objective[:n_x] = -exact_scores
    phase4 = run(exact_objective, [weighted_constraint, total_constraint,
                                  *no_skill_fallback_constraint, *mix_constraints,
                                  requested_type_constraint])
    best_exact = int(round(float(np.dot(exact_scores, phase4.x[:n_x]))))
    exact_constraint = ({j: float(exact_scores[j]) for j in range(n_x)}, best_exact, best_exact)
    above_objective = np.zeros(n_vars); above_objective[:n_x] = -above_scores
    phase5 = run(above_objective, [weighted_constraint, total_constraint,
                                   *no_skill_fallback_constraint, *mix_constraints,
                                   requested_type_constraint, exact_constraint])
    best_above = int(round(float(np.dot(above_scores, phase5.x[:n_x]))))
    above_constraint = ({j: float(above_scores[j]) for j in range(n_x)}, best_above, best_above)
    if affinity_matrix is not None:
        final_affinity_scores = np.array([
            affinity_matrix[pair[0], pair[1]] for pair in pairs
        ], dtype=float)
    else:
        final_affinity_scores = np.array([
            combined_preference_score(slots[pair[0]][1], people[pair[1]], len(slots), role_map)
            for pair in pairs
        ], dtype=float)
    affinity_objective = np.zeros(n_vars)
    affinity_objective[:n_x] = -final_affinity_scores
    final = run(affinity_objective, [weighted_constraint, total_constraint,
                                     *no_skill_fallback_constraint, *mix_constraints,
                                     requested_type_constraint,
                                     exact_constraint, above_constraint])

    selected = [(pairs[j][0], pairs[j][1], pairs[j][2])
                for j, value in enumerate(final.x[:n_x]) if value > 0.5]
    return slots, selected, reasons, eligible_by_slot


def solve(projects, people, skill_columns, role_map):
    slots = make_slots(projects)
    return solve_slots(slots, people, role_map)


def recommend_candidates(slots, selected, people, role_map, limit=5, only_unfilled=False):
    assigned_by_mesa = defaultdict(set)
    assigned_slots = set()
    for slot_index, person_index, _score in selected:
        mesa = mesa_group_key(slots[slot_index][1])
        assigned_by_mesa[mesa].add(person_index)
        assigned_slots.add(slot_index)
    candidates_by_mesa = defaultdict(dict)
    for slot_index, (_, project) in enumerate(slots):
        if only_unfilled and slot_index in assigned_slots:
            continue
        mesa = mesa_group_key(project)
        family = role_map.get(norm(project.get("Rol")), demand_family(
            project.get("Rol"), project.get("Dominio")))
        if family not in {"DEV", "DEV_INT", "DEV_EXT", "QA"}:
            continue
        for index, person in enumerate(people):
            ok, _, _ = eligibility(project, person, set(), role_map)
            # Una sugerencia es opcional: puede estar comprometida en otra
            # iniciativa. La capacidad solo se exige si se promueve a asignación.
            if ok and index not in assigned_by_mesa[mesa]:
                candidates_by_mesa[mesa][index] = max(
                    candidates_by_mesa[mesa].get(index, 0), affinity_score(project, person, role_map))
    result = {}
    for mesa, ranked in candidates_by_mesa.items():
        result[mesa] = [person for person, _ in sorted(
            ranked.items(), key=lambda item: (-item[1], people[item[0]]["name"], people[item[0]]["code"]))[:limit]]
    return result


def solve_with_suggestions(slots, people, role_map, viability_matrix=None,
                           affinity_matrix=None):
    """Resuelve primero la base de 7, luego permite asignar hasta 5 candidatos sugeridos."""
    base_result = solve_slots(slots, people, role_map,
                              viability_matrix=viability_matrix,
                              affinity_matrix=affinity_matrix)
    initial_suggestions = recommend_candidates(slots, base_result[1], people, role_map,
                                               limit=MAX_SUGGESTED_PER_MESA,
                                               only_unfilled=True)
    if not any(initial_suggestions.values()):
        return (*base_result, initial_suggestions)
    expanded_result = solve_slots(
        slots, people, role_map, viability_matrix=viability_matrix,
        affinity_matrix=affinity_matrix, recommendation_pool=initial_suggestions,
        fixed_assignments={(slot, person) for slot, person, _score in base_result[1]})
    return (*expanded_result, initial_suggestions)


def current_staffing_fields(person, project=None):
    assignment = person.get("current_assignment", {})
    return {
        "SDATOOL - Proyecto Staffing actual": assignment.get("sda", ""),
        "Fecha Inicio Staffing actual": assignment.get("fecha_inicio", ""),
        "Fecha Fin Actual Staffing": assignment.get("fecha_fin_actual", ""),
        "Mesa Staffing actual": assignment.get("mesa", ""),
        "% Dedicación Staffing actual": assignment.get("dedicacion", ""),
        "Comentario asignación Staffing": staffing_overlap_comment(person, project or {}),
        "Asignaciones Staffing Q4 (detalle)": json.dumps(
            person.get("staffing_assignments", []), ensure_ascii=False),
    }


def main():
    RESULTS.mkdir(exist_ok=True)
    initiative_header, projects = load_projects()
    people = load_people()
    role_map, role_report = build_role_map(projects, people)
    skill_columns = {norm(name) for name in read_csv(PEOPLE_FILE)[0]}
    slots = make_slots(projects)
    slots, selected, reasons, eligible_by_slot, recommendation_pool = solve_with_suggestions(
        slots, people, role_map)

    assignments = []
    assigned_slots = set()
    people_by_mesa = defaultdict(set)
    for slot, person_index, score in selected:
        source_row, project = slots[slot]
        person = people[person_index]
        group = mesa_group_key(project)
        assignment_source = ("Candidato sugerido adicional"
                             if person_index in recommendation_pool.get(group, [])
                             else "Asignación base")
        assigned_slots.add(slot)
        people_by_mesa[mesa_group_key(project)].add(person_index)
        continuity, domain = affinity_components(project, person)
        assignments.append({
            "Mesa": project.get("Mesa", ""), "SDATOOL": project.get("SDATOOL", ""),
            "Proyecto": project.get("Nombre Proyecto", ""), "Dominio asignado": project.get("Dominio", ""),
            "Rol requerido": project.get("Rol", ""), "Seniority requerido": project.get("Seniority", ""),
            "Familia de rol": role_map.get(norm(project.get("Rol")), demand_family(
                project.get("Rol"), project.get("Dominio"))),
            "Tipo requerido": project.get("_required_type", ""), "Tipo persona": person["resource_type"],
            "Origen de asignación": assignment_source,
            "Dominio persona": " | ".join(sorted(person.get("domains", set()))),
            "Código": person["code"], "Nombre": person["name"], "Rol persona": person["role_raw"],
            "Seniority persona": person["seniority_raw"], "Estado Staffing": person.get("staffing_status", ""),
            "Ajuste seniority": {2: "Exacto", 1: "Un nivel por encima", 0: "Un nivel por debajo"}[
                seniority_fit(project.get("Seniority"), person["seniority"])],
            "Ajuste tipo de recurso": (
                "Exacto" if person["resource_type"] == project.get("_required_type")
                else f"Sustitución: {person['resource_type']} cubre plaza {project.get('_required_type')}"),
            "FTE asignado": 1.0, "Afinidad SDA/Mesa actual": continuity,
            "Afinidad Dominio": domain, "Puntaje preferencias": score,
            "Ajuste coincidencia sin skills": -skill_match_penalty(project, person, role_map),
            "Puntaje skills obligatorias": skill_fit(project, person, role_map)[0],
            "Comentario skills obligatorias": skill_observation(project, person, role_map),
            **current_staffing_fields(person, project),
            "Skills obligatorias registradas": (
                f"{skill_fit(project, person, role_map)[0]}/{skill_fit(project, person, role_map)[1]}"),
            "Regla de nivel de skill": "No se filtra por nivel",
            "Fila origen": source_row,
        })

    unmet = []
    assigned_types_by_mesa = defaultdict(set)
    eligible_types_by_mesa = defaultdict(set)
    for slot, candidates in eligible_by_slot.items():
        group = mesa_group_key(slots[slot][1])
        eligible_types_by_mesa[group].update(people[index]["resource_type"] for index in candidates)
    for slot, person_index, _ in selected:
        assigned_types_by_mesa[mesa_group_key(slots[slot][1])].add(
            people[person_index]["resource_type"])
    for slot, (source_row, project) in enumerate(slots):
        if slot in assigned_slots:
            continue
        mesa = project.get("Mesa", "")
        candidates = eligible_by_slot.get(slot, [])
        if not (project.get("Rol") or "").strip() or not (project.get("Seniority") or "").strip():
            reason = "Falta rol o seniority en la demanda"
        elif not candidates:
            reason = diagnose_no_candidate(project, people, skill_columns, role_map)
        elif (eligible_types_by_mesa[mesa_group_key(project)] == {"interno", "externo"}
              and assigned_types_by_mesa[mesa_group_key(project)] != {"interno", "externo"}):
            reason = "Mesa sin asignar: no fue posible asegurar presencia interna y externa"
        elif len(people_by_mesa[mesa_group_key(project)]) >= MAX_BASE_ASSIGNMENTS_PER_GROUP + MAX_SUGGESTED_PER_MESA:
            reason = "El grupo alcanzó 7 asignaciones base más 5 sugeridas"
        elif (len(people_by_mesa[mesa_group_key(project)]) >= MAX_BASE_ASSIGNMENTS_PER_GROUP
              and not (set(recommendation_pool.get(mesa_group_key(project), []))
                       - people_by_mesa[mesa_group_key(project)])):
            reason = "El grupo ya usó sus candidatos sugeridos disponibles"
        else:
            reason = "Capacidad disponible o solapamiento de proyectos"
        unmet.append({
            "Mesa": mesa, "SDATOOL": project.get("SDATOOL", ""),
            "Proyecto": project.get("Nombre Proyecto", ""), "Dominio": project.get("Dominio", ""),
            "Tipo requerido": project.get("_required_type", ""),
            "Rol requerido": project.get("Rol", ""), "Seniority requerido": project.get("Seniority", ""),
            "Skills técnicas": " | ".join(required_skill_names(project, TECH_FIELDS)),
            "Skills funcionales": " | ".join(required_skill_names(project, FUNCTIONAL_FIELDS)),
            "Motivo": reason, "Fila origen": source_row,
        })

    def write(path, rows):
        headers = list(rows[0]) if rows else ["Mesa", "SDATOOL", "Proyecto", "Motivo"]
        with path.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=headers)
            writer.writeheader(); writer.writerows(rows)
    write(ASSIGNMENTS_FILE, assignments)
    write(UNFILLED_FILE, unmet)
    write(ROLE_MATCH_FILE, role_report)
    export_tobe(EXCEL_FILE, initiative_header, projects, slots, selected, people,
                skill_fit, TECH_FIELDS + FUNCTIONAL_FIELDS + SUGGESTED_FIELDS,
                role_map=role_map, recommendation_pool=recommendation_pool)
    print(f"Plazas modeladas: {len(slots)} | asignadas: {len(assignments)} | pendientes: {len(unmet)}")
    suggested_assigned = sum(
        person_index in recommendation_pool.get(mesa_group_key(slots[slot][1]), [])
        for slot, person_index, _score in selected
    )
    print(f"Asignaciones promovidas desde sugerencias: {suggested_assigned}")
    print(f"Dominios de iniciativa: {sorted({(p.get('Dominio') or '').strip() for p in projects if (p.get('Dominio') or '').strip()})}")
    print(f"Archivos: {ASSIGNMENTS_FILE} ; {UNFILLED_FILE} ; {ROLE_MATCH_FILE} ; {EXCEL_FILE}")


if __name__ == "__main__":
    main()
