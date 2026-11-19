"""Catálogo explícito de roles de demanda para la asignación."""
import re
import unicodedata


def normalize_role(value):
    value = unicodedata.normalize("NFKD", str(value or ""))
    value = "".join(char for char in value if not unicodedata.combining(char))
    return re.sub(r"[^a-z0-9]+", " ", value.casefold()).strip()


# Alias observados en la matriz y equivalencias aprobadas para el MILP.
DEMAND_ROLE_MAP = {
    "lt": "LT",
    "lider tecnico": "LT",
    "technical lead": "LT",
    "tech lead": "LT",
    "dev int": "DEV",
    "dev interno": "DEV",
    "desarrollador interno": "DEV",
    "dev ext": "DEV",
    "dev externo": "DEV",
    "desarrollador externo": "DEV",
    "desarrollador": "DEV",
    "data ing": "DEV",
    "ux": "DEV",
    "po": "DEV",  # Mapeo provisional; estas filas aún deben pasar el filtro técnico.
    "product engineer lt ia": "LT",
    "qa": "QA",
    "testing": "QA",
    "test": "QA",
    "tester": "QA",
    "quality assurance": "QA",
    "quality context": "QA",
}

LEADERSHIP_ROLE_MARKERS = (
    "service reliability lead", "tech manager", "lider", "lead",
)
# Roles adicionales que aparecen con seniority Lead en la matriz SkillsCOE.
# Solo los perfiles Lead con estos roles suman LT; el resto conserva su
# familia DEV/QA. Los perfiles de Testing se excluyen antes de aplicar esto.
LEAD_SENIORITY_LT_ROLES = {
    "desarrollador", "arquitecto", "despliegues", "otros",
    "software expert", "tech",
}
TESTING_ROLE_MARKERS = ("qa", "testing", "test", "tester", "quality")
INVALID_PERSON_ROLE_MARKERS = (
    "na", "n a", "n a no aplica", "null", "none", "nan",
    "cod eliminado", "codigo eliminado", "eliminado", "sin rol",
)


def demand_role_family(value):
    return DEMAND_ROLE_MAP.get(normalize_role(value))


def person_role_families(value, seniority=None):
    """Familias que puede cubrir cada perfil según el mapeo provisional BD2."""
    role = normalize_role(value)
    if not role or role in INVALID_PERSON_ROLE_MARKERS or role.startswith(
        ("na ", "n a ", "cod eliminado", "codigo eliminado")
    ):
        return set()
    if any(marker in role for marker in TESTING_ROLE_MARKERS):
        return {"QA"}
    families = {"DEV", "QA"}  # Desarrollo, arquitectura, datos, infraestructura y otros.
    if role == "product engineer lt ia" or any(marker in role for marker in LEADERSHIP_ROLE_MARKERS):
        families.add("LT")
    if (normalize_role(seniority) == "lead"
            and role in LEAD_SENIORITY_LT_ROLES):
        families.add("LT")
    return families
