"""Normalización de roles para las personas y las iniciativas.

Las personas usan exclusivamente el rol de la columna E de SkillsCOE. Los dos
catálogos terminan en una de las familias de asignación DEV, QA o LT.
"""
import re
import unicodedata


def normalize_role(value):
    """Quita tildes y puntuación para comparar alias de roles."""
    value = unicodedata.normalize("NFKD", str(value or ""))
    value = "".join(char for char in value if not unicodedata.combining(char))
    return re.sub(r"[^a-z0-9]+", " ", value.casefold()).strip()


# El rol canónico de ambos lados del cruce queda limitado a estas tres familias.
ROLE_ALIASES = {
    "dev": "DEV",
    "dev int": "DEV",
    "dev interno": "DEV",
    "dev ext": "DEV",
    "dev externo": "DEV",
    "desarrollador": "DEV",
    "desarrollador interno": "DEV",
    "desarrollador externo": "DEV",
    "developer": "DEV",
    "arquitecto": "DEV",
    "arquitectura": "DEV",
    "despliegues": "DEV",
    "data ing": "DEV",
    "qa": "QA",
    "qa 0 5 fte": "QA",
    "testing": "QA",
    "test": "QA",
    "tester": "QA",
    "quality assurance": "QA",
    "quality context": "QA",
    "lt": "LT",
    "lt 0 5 fte": "LT",
    "lider tecnico": "LT",
    "lider primera linea": "LT",
    "technical lead": "LT",
    "tech lead": "LT",
    "tech manager lider": "LT",
    "service reliability lead": "LT",
    "ptl": "LT",
    "product engineer lt ia": "LT",
}

# La matriz SkillsCOE es la fuente de verdad para la familia de las personas.
PERSON_ROLE_ALIASES = {
    "desarrollador": "DEV",
    "arquitecto": "DEV",
    "despliegues": "DEV",
    "qa": "QA",
    "tech manager lider": "LT",
    "service reliability lead": "LT",
    "lider primera linea": "LT",
    "ptl": "LT",
}


def _is_na_role(value):
    """Detecta NA únicamente al inicio, como establece la regla de limpieza."""
    return normalize_role(value).startswith("na")


def demand_role_family(value, domain=None):
    """Homologa un rol de iniciativa; None indica que debe excluirse/reportarse."""
    return ROLE_ALIASES.get(normalize_role(value))


def canonical_demand_role(value, domain=None):
    """Devuelve DEV, QA o LT; conserva el original desconocido para reportarlo."""
    return demand_role_family(value, domain) or str(value or "").strip()


def canonical_person_role(value, seniority=None, resource_type=None, domain=None):
    """Homologa el rol de SkillsCOE columna E sin inferirlo por seniority."""
    role = normalize_role(value)
    if _is_na_role(value):
        return ""
    return PERSON_ROLE_ALIASES.get(role, "")


def person_role_families(value, seniority=None):
    """Devuelve la única familia asignable para un rol de SkillsCOE reconocido."""
    canonical = canonical_person_role(value)
    return {canonical} if canonical else set()
