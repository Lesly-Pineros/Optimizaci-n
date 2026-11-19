#!/usr/bin/env python3
"""Asigna personas a plazas de iniciativas con un MILP (SciPy/HiGHS).

Las filas con dos o más skills generan plazas individuales según Internal/
External. Rol, seniority, tipo y dominio son requisitos; las skills se optimizan
por cobertura y el nivel esperado depende del seniority.
"""
import csv
import re
import unicodedata
from collections import defaultdict, Counter
from difflib import SequenceMatcher
from pathlib import Path
from salidas_asignacion import export_tobe

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "DATA"
INPUTS = ROOT / "RESULTADOS_LIMPIEZA"
RESULTS = ROOT / "RESULTADOS MODELOS"
PEOPLE_FILE = INPUTS / "Matriz SkillsCOE enriquecida.csv"
PROJECT_FILE = INPUTS / "ENG-Matriz Iniciativas con skills.csv"
STAFFING_FILE = DATA / "STAFFING.csv"
ASSIGNMENTS_FILE = RESULTS / "Asignaciones MILP.csv"
UNFILLED_FILE = RESULTS / "Plazas sin asignar MILP.csv"
ROLE_MATCH_FILE = RESULTS / "Similitud de roles.csv"
EXCEL_FILE = RESULTS / "Asignacion_To-Be_Baseline.xlsx"

MAX_PEOPLE_PER_MESA = 7
# Capacidad normal por persona. Si impide cubrir la demanda, el MILP permite
# como máximo una segunda plaza y minimiza el número de personas duplicadas.
MAX_ASSIGNMENTS_PER_PERSON = 1
# Equivalencias claras entre abreviaturas de demanda y cargos de SkillsCOE.
ROLE_ALIASES = {
    "lt": "tech manager/líder",
    "dev ext": "desarrollador",
    "dev int": "desarrollador",
}
ROLE_SIMILARITY_THRESHOLD = 0.80
MIN_SKILLS_PER_ROLE = 2
# Nivel mínimo para considerar una skill cubierta: 1 básico, 2 intermedio,
# 3 avanzado. Se aplica a skills técnicas y funcionales obligatorias.
SENIORITY_SKILL_LEVEL = {
    "junior": 1,
    "mid level": 2,
    "senior": 3,
    "lead": 3,
}
# Regla de negocio confirmada para esta corrida: Operaciones requiere dominio ENG.
DIRECTION_MAP = {"Operaciones": "ENG"}

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
    return header, filled


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
        skills = {norm(name): level(value) for name, value in row.items() if name.strip()}
        result.append({
            "key": person_key,
            "code": (row.get("Código") or "").strip(),
            "name": (row.get("Nombre") or "").strip(),
            "role": norm(row.get("Rol Principal")),
            "role_raw": (row.get("Rol Principal") or "").strip(),
            "seniority": norm(row.get("Nivel del Rol Actual")),
            "seniority_raw": (row.get("Nivel del Rol Actual") or "").strip(),
            "company": norm(row.get("Empresa")),
            "resource_type": {"interno": "interno", "externo": "externo"}.get(norm(row.get("Tipo de Recurso")), ""),
            "direction": direction,
            "skills": skills,
        })
    return result


def role_target(value):
    normalized = norm(value)
    return norm(ROLE_ALIASES.get(normalized, value))


def build_role_map(projects, people):
    candidate_roles = {}
    for person in people:
        if person["role"]:
            candidate_roles.setdefault(person["role"], person["role_raw"])
    mapping, report = {}, []
    for role in sorted({(p.get("Rol") or "").strip() for p in projects if (p.get("Rol") or "").strip()}):
        source = norm(role)
        suggested_target = ""
        if source in ROLE_ALIASES:
            target = norm(ROLE_ALIASES[source])
            suggested_target = target
            score, method = (1.0, "equivalencia de abreviatura") if target in candidate_roles else (0.0, "equivalencia sin rol en SkillsCOE")
        elif source in candidate_roles:
            target, score, method = source, 1.0, "coincidencia exacta"
            suggested_target = target
        else:
            ranked = []
            source_tokens = set(source.split())
            for candidate in candidate_roles:
                tokens = set(candidate.split())
                token_score = len(source_tokens & tokens) / max(1, len(source_tokens | tokens))
                char_score = SequenceMatcher(None, source, candidate, autojunk=False).ratio()
                ranked.append((max(char_score, token_score), candidate))
            score, target = max(ranked, default=(0.0, ""))
            suggested_target = target if score >= ROLE_SIMILARITY_THRESHOLD else ""
            method = "sugerencia para revisión" if score >= ROLE_SIMILARITY_THRESHOLD else "sin similitud suficiente"
            # Las similitudes no exactas se reportan, pero requieren revisión humana.
            target = ""
        mapping[source] = target
        report.append({
            "Rol iniciativa": role,
            "Rol SkillsCOE sugerido": candidate_roles.get(suggested_target, ""),
            "Similitud": round(score, 3),
            "Método": method,
            "Usar en asignación": "Sí" if target else "No",
        })
    return mapping, report


def required_skill_names(row, fields):
    return [norm(row.get(field)) for field in fields if (row.get(field) or "").strip()]


def skill_fit(project, person):
    """Resume la cobertura de skills sin exigir que se cumplan todas."""
    required = required_skill_names(project, TECH_FIELDS + FUNCTIONAL_FIELDS)
    seniority = norm(project.get("Seniority"))
    minimum = SENIORITY_SKILL_LEVEL.get(seniority, 1)
    levels = [person["skills"].get(skill, 0) for skill in required]
    covered = sum(value >= minimum for value in levels)
    known = sum(value > 0 for value in levels)
    # La cobertura al nivel esperado manda; luego se favorecen skills con
    # conocimiento parcial y mejores niveles dentro del umbral.
    score = covered * 10000 + known * 100 + sum(min(value, minimum) for value in levels)
    suggested = required_skill_names(project, SUGGESTED_FIELDS)
    score += sum(person["skills"].get(skill, 0) > 0 for skill in suggested) * 10
    return covered, len(required), known, minimum, score


def eligibility(project, person, skill_columns, role_map):
    role = role_map.get(norm(project.get("Rol")), "")
    seniority = norm(project.get("Seniority"))
    if not role or person["role"] != role:
        return False, 0, "rol"
    if not seniority or person["seniority"] != seniority:
        return False, 0, "seniority"

    required_type = project.get("_required_type", "")
    if not required_type or person["resource_type"] != required_type:
        return False, 0, "tipo interno/externo"

    project_direction = norm(DIRECTION_MAP.get(
        (project.get("Dominio") or "").strip(), project.get("Dominio")))
    person_domains = {norm(v) for v in person["direction"].split("|") if norm(v)}
    if not project_direction or project_direction not in person_domains:
        return False, 0, "dominio"

    is_technical_lead = norm(project.get("Rol")) in {"lt", "lider tecnico", "lider técnico"}
    if is_technical_lead and person["company"] != norm("BBVA"):
        return False, 0, "LT solo BBVA"

    required_tech = required_skill_names(project, TECH_FIELDS)
    required_functional = required_skill_names(project, FUNCTIONAL_FIELDS)
    skill_count = len(required_skill_names(project, TECH_FIELDS + FUNCTIONAL_FIELDS + SUGGESTED_FIELDS))
    if skill_count < MIN_SKILLS_PER_ROLE:
        return False, 0, "menos de dos skills"
    covered, _, known, _, fit_score = skill_fit(project, person)
    if known == 0:
        return False, 0, "sin skills coincidentes"
    return True, fit_score, ""


def diagnose_no_candidate(project, people, skill_columns, role_map):
    """Devuelve el primer criterio que deja la plaza sin candidatos."""
    pool = people
    role = role_map.get(norm(project.get("Rol")), "")
    pool = [p for p in pool if role and p["role"] == role]
    if not pool:
        return "No hay persona con rol compatible"
    seniority = norm(project.get("Seniority"))
    pool = [p for p in pool if seniority and p["seniority"] == seniority]
    if not pool:
        return "No hay persona con seniority compatible"
    required_type = project.get("_required_type", "")
    pool = [p for p in pool if p["resource_type"] == required_type]
    if not pool:
        return "No hay persona del tipo interno/externo requerido"
    domain = norm(DIRECTION_MAP.get(project.get("Dominio", ""), project.get("Dominio", "")))
    pool = [p for p in pool if domain and domain in {norm(v) for v in p["direction"].split("|") if norm(v)}]
    if not pool:
        return "No hay persona con el dominio requerido por la iniciativa"
    if norm(project.get("Rol")) in {"lt", "lider tecnico"}:
        pool = [p for p in pool if p["company"] == norm("BBVA")]
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
            "sin skills coincidentes": "Ninguna persona elegible tiene alguna de las skills obligatorias",
            "menos de dos skills": "La plaza declara menos de dos skills",
        }.get(reason, "No hay persona con skills coincidentes")
    return "No hay persona elegible"


def solve(projects, people, skill_columns, role_map):
    slots = []
    for index, project in enumerate(projects):
        skill_count = len(required_skill_names(project, TECH_FIELDS + FUNCTIONAL_FIELDS + SUGGESTED_FIELDS))
        if not (project.get("Rol") or "").strip() or not (project.get("Seniority") or "").strip() or skill_count < MIN_SKILLS_PER_ROLE:
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

    pairs = []
    reasons = defaultdict(int)
    eligible_by_slot = defaultdict(list)
    for slot_index, (source_row, project) in enumerate(slots):
        for person_index, person in enumerate(people):
            ok, score, reason = eligibility(project, person, skill_columns, role_map)
            if ok:
                project_key = (key(project.get("SDATOOL")), norm(project.get("Nombre Proyecto")))
                mesa_key = norm(project.get("Mesa"))
                domain_key = norm(DIRECTION_MAP.get(project.get("Dominio", ""), project.get("Dominio", "")))
                pairs.append((slot_index, person_index, score, project_key, mesa_key, domain_key))
                eligible_by_slot[slot_index].append(person_index)
            else:
                reasons[(slot_index, reason)] += 1

    if not pairs:
        return slots, [], reasons, eligible_by_slot

    try:
        import numpy as np
        from scipy.optimize import Bounds, LinearConstraint, milp
        from scipy.sparse import coo_matrix
    except ImportError as exc:
        raise RuntimeError("Falta SciPy. Instala las dependencias con: python3 -m pip install -r requirements-milp.txt") from exc

    # x: asignaciones; y: persona presente en una mesa; r: desborde de una plaza.
    # El desborde solo se habilita si es necesario para cubrir toda la demanda.
    n_x = len(pairs)
    y_keys = sorted({(person, mesa) for _, person, _, _, mesa, _ in pairs if mesa})
    y_index = {value: n_x + i for i, value in enumerate(y_keys)}
    repeat_start = n_x + len(y_keys)
    repeat_keys = sorted({person for _, person, *_ in pairs})
    repeat_index = {person: repeat_start + i for i, person in enumerate(repeat_keys)}
    n_vars = repeat_start + len(repeat_keys)

    by_slot, by_person, by_person_project = defaultdict(list), defaultdict(list), defaultdict(list)
    by_mesa = defaultdict(list)
    rows, upper = [], []
    def add_constraint(coefficients, ub):
        rows.append(coefficients); upper.append(float(ub))

    for j, (slot, person, score, project_key, mesa, domain) in enumerate(pairs):
        by_slot[slot].append(j)
        by_person[person].append(j)
        by_person_project[(person, project_key)].append(j)
        if mesa:
            by_mesa[mesa].append(y_index[(person, mesa)])
            add_constraint({j: 1.0, y_index[(person, mesa)]: -1.0}, 0)

    for variables in by_slot.values():
        add_constraint({v: 1.0 for v in variables}, 1)
    for person, variables in by_person.items():
        coefficients = {v: 1.0 for v in variables}
        coefficients[repeat_index[person]] = -1.0
        add_constraint(coefficients, MAX_ASSIGNMENTS_PER_PERSON)
    for variables in by_person_project.values():
        add_constraint({v: 1.0 for v in variables}, 1)
    for variables in by_mesa.values():
        add_constraint({v: 1.0 for v in set(variables)}, MAX_PEOPLE_PER_MESA)
    rr, cc, vv = [], [], []
    for row_index, coefficients in enumerate(rows):
        for variable, coefficient in coefficients.items():
            rr.append(row_index); cc.append(variable); vv.append(coefficient)
    matrix = coo_matrix((vv, (rr, cc)), shape=(len(rows), n_vars)).tocsr()
    # Prioriza cubrir plazas; el ajuste por skills desempata.
    objective = np.zeros(n_vars)
    objective[:n_x] = [-1_000_000_000.0 - score for _, _, score, _, _, _ in pairs]
    objective[repeat_start:] = 1_000_000.0
    result = milp(objective, integrality=np.ones(n_vars), bounds=Bounds(0, 1),
                  constraints=LinearConstraint(matrix, np.full(len(rows), -np.inf), np.array(upper)),
                  options={"time_limit": 300})
    if result.x is None:
        raise RuntimeError(f"El solver no encontró solución: {result.message}")
    selected = [(pairs[j][0], pairs[j][1], pairs[j][2]) for j, value in enumerate(result.x[:n_x]) if value > 0.5]
    return slots, selected, reasons, eligible_by_slot


def main():
    RESULTS.mkdir(exist_ok=True)
    initiative_header, projects = load_projects()
    people = load_people()
    role_map, role_report = build_role_map(projects, people)
    skill_columns = {norm(name) for name in read_csv(PEOPLE_FILE)[0]}
    slots, selected, reasons, eligible_by_slot = solve(projects, people, skill_columns, role_map)

    assignments = []
    assignment_counts = Counter(person for _, person, _ in selected)
    assigned_slots = set()
    projects_by_person = defaultdict(set)
    people_by_mesa = defaultdict(set)
    project_key_by_slot = {}
    for slot_index, (_, demand) in enumerate(slots):
        project_key_by_slot[slot_index] = (key(demand.get("SDATOOL")), norm(demand.get("Nombre Proyecto")))
    for slot, person_index, score in selected:
        source_row, project = slots[slot]
        person = people[person_index]
        assigned_slots.add(slot)
        projects_by_person[person_index].add(project_key_by_slot[slot])
        people_by_mesa[norm(project.get("Mesa"))].add(person_index)
        assignments.append({
            "Mesa": project.get("Mesa", ""), "SDATOOL": project.get("SDATOOL", ""),
            "Proyecto": project.get("Nombre Proyecto", ""), "Dominio asignado": project.get("Dominio", ""),
            "Rol requerido": project.get("Rol", ""), "Seniority requerido": project.get("Seniority", ""),
            "Rol SkillsCOE equivalente": next((r["Rol SkillsCOE sugerido"] for r in role_report if norm(r["Rol iniciativa"]) == norm(project.get("Rol"))), ""),
            "Tipo requerido": project.get("_required_type", ""), "Tipo persona": person["resource_type"],
            "Código": person["code"], "Nombre": person["name"], "Rol persona": person["role_raw"],
            "Seniority persona": person["seniority_raw"], "Puntaje ajuste": score,
            "Excepción capacidad": (f"Asignado a {assignment_counts[person_index]} plazas; capacidad base 1"
                                     if assignment_counts[person_index] > MAX_ASSIGNMENTS_PER_PERSON else ""),
            "Skills cubiertas al nivel esperado": f"{skill_fit(project, person)[0]}/{skill_fit(project, person)[1]}",
            "Nivel mínimo de skill": skill_fit(project, person)[3],
            "Fila origen": source_row,
        })

    unmet = []
    for slot, (source_row, project) in enumerate(slots):
        if slot in assigned_slots:
            continue
        mesa = project.get("Mesa", "")
        candidates = eligible_by_slot.get(slot, [])
        if not (project.get("Rol") or "").strip() or not (project.get("Seniority") or "").strip():
            reason = "Falta rol o seniority en la demanda"
        elif not (project.get("Dominio") or "").strip():
            reason = "Falta Dominio en iniciativa"
        elif not candidates:
            reason = diagnose_no_candidate(project, people, skill_columns, role_map)
        elif len(people_by_mesa[norm(mesa)]) >= MAX_PEOPLE_PER_MESA:
            reason = "La mesa ya tiene 7 personas únicas"
        elif all(len(projects_by_person[p]) >= MAX_ASSIGNMENTS_PER_PERSON for p in candidates):
            reason = "Las personas elegibles alcanzaron la capacidad máxima del modelo"
        elif all(project_key_by_slot[slot] in projects_by_person[p] for p in candidates):
            reason = "Las personas elegibles ya están asignadas a este proyecto"
        else:
            reason = "Conflicto de capacidad entre plazas elegibles"
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
                skill_fit, TECH_FIELDS + FUNCTIONAL_FIELDS + SUGGESTED_FIELDS)
    overflow_people = sum(count > MAX_ASSIGNMENTS_PER_PERSON for count in assignment_counts.values())
    print(f"Plazas modeladas: {len(slots)} | asignadas: {len(assignments)} | pendientes: {len(unmet)} | "
          f"personas con sobrecapacidad mínima: {overflow_people}")
    print(f"Dominios de iniciativa: {sorted({(p.get('Dominio') or '').strip() for p in projects if (p.get('Dominio') or '').strip()})}")
    print(f"Archivos: {ASSIGNMENTS_FILE} ; {UNFILLED_FILE} ; {ROLE_MATCH_FILE} ; {EXCEL_FILE}")


if __name__ == "__main__":
    main()
