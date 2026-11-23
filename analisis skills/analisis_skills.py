#!/usr/bin/env python3
"""Analiza skills de personas ENG/GRM y contrasta la cobertura con iniciativas.

Edición rápida: cambia DOMINIOS, INPUT_PERSONAS e INPUT_INICIATIVAS si
quieres analizar otros dominios o usar archivos fuente diferentes.
Ejecutar desde cualquier carpeta: python3 "analisis skills/analisis_skills.py"
"""

import csv
import re
import sys
import unicodedata
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from xml.sax.saxutils import escape

# Rutas relativas a la carpeta del proyecto para que el script sea portable.
ROOT = Path(__file__).resolve().parent.parent
OUTPUT_DIR = Path(__file__).resolve().parent
INPUT_PERSONAS = ROOT / "RESULTADOS_LIMPIEZA" / "Matriz SkillsCOE enriquecida.csv"
# El análisis usa las matrices de iniciativas depuradas para ambos dominios.
INPUT_INICIATIVAS = {
    "ENG": (ROOT / "RESULTADOS_LIMPIEZA" / "ENG-iniciativas con skills.csv", 0),
    "GRM": (ROOT / "RESULTADOS_LIMPIEZA" / "GRM-iniciativas con skills.csv", 0),
}

# Dominios a analizar. Se admiten dominios combinados como "RCS | ENG".
DOMINIOS = ("ENG", "GRM")

# Valores de skills: 0 - Nulo y vacío se tratan de forma diferente. El primero
# indica que la persona declaró no tener la skill; el vacío indica sin dato.
NIVELES_POSITIVOS = {"1 - básico", "2 - intermedio", "3 - avanzado", "4 - experto"}
NIVEL_NULO = "0 - nulo"

# Los nombres de demanda de las iniciativas se traducen a familias comunes
# (DEV, QA, LT, etc.) con el catálogo del proyecto.
sys.path.insert(0, str(ROOT))
from mapeo_roles import demand_role_family, person_role_families  # noqa: E402


def normalizar(valor):
    """Quita tildes, espacios repetidos y diferencias entre mayúsculas."""
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", texto).strip().casefold()


def leer_csv(ruta, fila_encabezados=0):
    """Lee CSV UTF-8 con BOM y permite indicar la fila donde están los encabezados."""
    with ruta.open("r", encoding="utf-8-sig", newline="") as archivo:
        registros = list(csv.reader(archivo))
    if not registros or fila_encabezados >= len(registros):
        return [], []
    encabezados = registros[fila_encabezados]
    filas = []
    for valores in registros[fila_encabezados + 1:]:
        valores += [""] * max(0, len(encabezados) - len(valores))
        filas.append(dict(zip(encabezados, valores)))
    return encabezados, filas


def referencia_columna(indice):
    """Convierte un índice base cero a la referencia Excel (A, B, ..., AA)."""
    referencia = ""
    numero = indice + 1
    while numero:
        numero, resto = divmod(numero - 1, 26)
        referencia = chr(65 + resto) + referencia
    return referencia


def escribir_xlsx(ruta, hojas):
    """Crea un XLSX sencillo con XML estándar, sin requerir paquetes externos."""
    tipos = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
             '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
             '<Default Extension="xml" ContentType="application/xml"/>'
             '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
    relaciones_libro = []
    definiciones_hoja = []
    for numero, (nombre, columnas, filas) in enumerate(hojas, start=1):
        tipos += (f'<Override PartName="/xl/worksheets/sheet{numero}.xml" '
                  'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
        relaciones_libro.append(
            f'<Relationship Id="rId{numero}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{numero}.xml"/>')
        definiciones_hoja.append(
            f'<sheet name="{escape(nombre)}" sheetId="{numero}" r:id="rId{numero}"/>')

    with zipfile.ZipFile(ruta, "w", compression=zipfile.ZIP_DEFLATED) as libro:
        tipos += '</Types>'
        libro.writestr("[Content_Types].xml", tipos)
        libro.writestr("_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>')
        libro.writestr("xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets>' + "".join(definiciones_hoja) + '</sheets></workbook>')
        libro.writestr("xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(relaciones_libro) + '</Relationships>')

        for numero, (_, columnas, filas) in enumerate(hojas, start=1):
            lineas = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>']
            registros = [columnas] + [[fila.get(columna, "") for columna in columnas] for fila in filas]
            for num_fila, registro in enumerate(registros, start=1):
                celdas = []
                for indice, valor in enumerate(registro):
                    if valor is None or valor == "":
                        continue
                    ref = f"{referencia_columna(indice)}{num_fila}"
                    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
                        celdas.append(f'<c r="{ref}"><v>{valor}</v></c>')
                    else:
                        celdas.append(f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">'
                                      f'{escape(str(valor))}</t></is></c>')
                lineas.append(f'<row r="{num_fila}">' + "".join(celdas) + '</row>')
            lineas.append('</sheetData></worksheet>')
            libro.writestr(f"xl/worksheets/sheet{numero}.xml", "".join(lineas))


def dividir_dominios(valor):
    """Convierte un campo de dominio individual o combinado en códigos."""
    return {normalizar(x) for x in re.split(r"\s*[|,;/]\s*", str(valor or "")) if x.strip()}


def nivel_skill(valor):
    """Clasifica una respuesta de skill como nivel, ausencia declarada o sin dato."""
    valor_norm = normalizar(valor)
    if not valor_norm:
        return "Sin dato"
    if valor_norm == NIVEL_NULO:
        return "No la tiene (Nulo)"
    if valor_norm in NIVELES_POSITIVOS:
        return valor.strip()
    return "Otro valor"


def tiene_skill(valor):
    """True si hay un nivel positivo conocido (Básico a Experto)."""
    return normalizar(valor) in NIVELES_POSITIVOS


def encontrar_skills(encabezados):
    """Toma las columnas de skills entre Nivel del Rol Actual y el campo libre."""
    inicio = next(i for i, h in enumerate(encabezados)
                  if normalizar(h) == normalizar("Nivel del Rol Actual")) + 1
    fin = next((i for i, h in enumerate(encabezados)
                if normalizar(h) == normalizar("Skill a desarrollar en los próximos meses")),
               len(encabezados))
    # Omite encabezados vacíos en caso de que el Excel original tuviera columnas sin nombre.
    return [h for h in encabezados[inicio:fin] if h and h.strip()]


def es_compatible(persona, familia, rol_demanda):
    """Comprueba familia de rol y, para Dev Int/Ext, tipo de recurso."""
    familias = person_role_families(persona.get("Rol Staffing") or persona.get("Rol Principal"),
                                    persona.get("Seniority") or persona.get("Nivel del Rol Actual"))
    if not familia:
        return False
    tipo = normalizar(persona.get("Tipo de Recurso Staffing") or persona.get("Tipo de Recurso"))
    if familia in {"DEV_INT", "DEV_EXT"}:
        if "DEV" not in familias:
            return False
        if familia == "DEV_INT" and tipo != "interno":
            return False
        if familia == "DEV_EXT" and tipo not in {"externo", "subsidiaria"}:
            return False
    elif familia not in familias:
        return False
    return True


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    headers, personas = leer_csv(INPUT_PERSONAS)
    iniciativas_headers_por_dominio = {}
    iniciativas_por_dominio = {}
    for dominio, (ruta, fila_encabezados) in INPUT_INICIATIVAS.items():
        headers_iniciativas, filas_iniciativas = leer_csv(ruta, fila_encabezados)
        # El archivo fuente de ENG no incluye una columna de dominio: se la asignamos
        # según el archivo del que provienen sus iniciativas.
        if "Dominio" not in headers_iniciativas:
            headers_iniciativas = list(headers_iniciativas) + ["Dominio"]
        # En el Excel de ENG varios datos del proyecto aparecen una sola vez y
        # las filas siguientes contienen otras demandas/roles con celdas vacías.
        # Rellenamos esos campos de contexto para conservar cada demanda asociada.
        campos_rellenables = (
            "Mesa", "Foco estratégico (Fábrica)", "SDATOOL", "Nombre Proyecto",
            "Fecha Inicio", "Fecha Fin", "Dirección", "Estado", "Workspace",
        )
        contexto = {campo: "" for campo in campos_rellenables}
        for fila in filas_iniciativas:
            fila.setdefault("Dominio", dominio)
            if not any(str(valor or "").strip() for valor in fila.values()):
                continue
            for campo in campos_rellenables:
                valor = (fila.get(campo) or "").strip()
                if valor:
                    contexto[campo] = valor
                elif contexto[campo]:
                    fila[campo] = contexto[campo]
        # Cada fila con rol representa una demanda. Ignora filas separadoras y
        # roles que no tienen equivalencia en el catálogo del modelo.
        filas_iniciativas = [fila for fila in filas_iniciativas
                             if (fila.get("Rol") or "").strip()
                             and demand_role_family(fila.get("Rol", ""), dominio)]
        iniciativas_headers_por_dominio[dominio] = headers_iniciativas
        iniciativas_por_dominio[dominio] = filas_iniciativas
    skills = encontrar_skills(headers)
    # Se admiten valores exactos o combinados, por ejemplo "RCS | ENG".
    personas_por_dominio = {
        dominio: [p for p in personas if normalizar(dominio) in dividir_dominios(p.get("Dominio", ""))]
        for dominio in DOMINIOS
    }
    for dominio in DOMINIOS:
        personas_dom = personas_por_dominio[dominio]
        iniciativas_dom = [i for i in iniciativas_por_dominio[dominio]
                           if normalizar(dominio) in dividir_dominios(i.get("Dominio", ""))]
        iniciativas_headers = iniciativas_headers_por_dominio[dominio]

        # Detalle: una fila por persona y skill, incluidos niveles nulos y sin dato.
        detalle = []
        resumen = defaultdict(Counter)
        for persona in personas_dom:
            seniority = (persona.get("Seniority") or persona.get("Nivel del Rol Actual") or "Sin seniority").strip()
            for skill in skills:
                valor = (persona.get(skill) or "").strip()
                clasificacion = nivel_skill(valor)
                detalle.append({
                    "Dominio": dominio,
                    "Código": persona.get("Código", ""),
                    "Nombre": persona.get("Nombre", ""),
                    "Rol": persona.get("Rol Staffing") or persona.get("Rol Principal", ""),
                    "Nivel del Rol Actual": persona.get("Nivel del Rol Actual", ""),
                    "Seniority": seniority,
                    "Skill": skill,
                    "Nivel reportado": valor,
                    "Estado skill": clasificacion,
                })
                resumen[(seniority, skill)][clasificacion] += 1
        detalle_columnas = ["Dominio", "Código", "Nombre", "Rol", "Nivel del Rol Actual", "Seniority",
                            "Skill", "Nivel reportado", "Estado skill"]

        # Resumen por seniority: cada columna muestra cuántas personas reportan ese nivel.
        categorias = ["4 - Experto", "3 - Avanzado", "2 - Intermedio", "1 - Básico",
                      "No la tiene (Nulo)", "Sin dato", "Otro valor"]
        resumen_filas = []
        for (seniority, skill), conteos in sorted(resumen.items()):
            fila = {"Dominio": dominio, "Seniority": seniority, "Skill": skill,
                    "Personas": sum(conteos.values()),
                    "Con skill": sum(conteos[c] for c in categorias[:4]),
                    "Sin skill declarada como Nulo": conteos["No la tiene (Nulo)"],
                    "Sin dato": conteos["Sin dato"]}
            fila.update({nivel: conteos[nivel] for nivel in categorias})
            resumen_filas.append(fila)
        resumen_columnas = ["Dominio", "Seniority", "Skill", "Personas", "Con skill",
                            "Sin skill declarada como Nulo", "Sin dato"] + categorias

        # Lista de personas Lead con todas sus skills para revisar sus perfiles.
        leads = [p for p in personas_dom
                 if normalizar(p.get("Nivel del Rol Actual")) == "lead"]
        lead_columnas = ["Dominio", "Código", "Nombre", "Rol Principal", "Nivel del Rol Actual",
                         "Tipo de Recurso Staffing"] + skills
        lead_filas = [{"Dominio": dominio, **{c: p.get(c, "") for c in lead_columnas if c != "Dominio"}}
                      for p in leads]
        # Para cada iniciativa y skill obligatoria, cuenta perfiles compatibles que la reportan.
        # Si el nombre de skill no está en la matriz, queda señalado como no encontrado.
        mapa_skills = {normalizar(s): s for s in skills}
        requeridas_por_iniciativa = []
        for iniciativa in iniciativas_dom:
            familia = demand_role_family(iniciativa.get("Rol", ""), dominio)
            candidatos = [p for p in personas_dom if es_compatible(p, familia, iniciativa.get("Rol", ""))]
            columnas_requeridas = [(h, "Técnica") for h in iniciativas_headers
                                   if "habilidades tecnicas obligatorias" in normalizar(h)]
            columnas_requeridas += [(h, "Funcional") for h in iniciativas_headers
                                    if "habilidades funcionales obligatorias" in normalizar(h)]
            for columna, tipo_skill in columnas_requeridas:
                nombre_requerido = (iniciativa.get(columna) or "").strip()
                if not nombre_requerido:
                    continue
                skill_columna = mapa_skills.get(normalizar(nombre_requerido))
                con_skill = sum(1 for p in candidatos if skill_columna and tiene_skill(p.get(skill_columna, "")))
                requeridas_por_iniciativa.append({
                    "Dominio": dominio,
                    "SDATOOL": iniciativa.get("SDATOOL", ""),
                    "Nombre Proyecto": iniciativa.get("Nombre Proyecto", ""),
                    "Rol requerido": iniciativa.get("Rol", ""),
                    "Familia de rol": familia or "Sin mapeo",
                    "Tipo de skill": tipo_skill,
                    "Skill requerida": nombre_requerido,
                    "Skill en matriz": skill_columna or "NO ENCONTRADA",
                    "Personas compatibles": len(candidatos),
                    "Personas compatibles con skill": con_skill,
                    "Brecha": "Sí" if not skill_columna or con_skill == 0 else "No",
                })
        requeridas_columnas = ["Dominio", "SDATOOL", "Nombre Proyecto", "Rol requerido", "Familia de rol",
                               "Tipo de skill", "Skill requerida", "Skill en matriz", "Personas compatibles",
                               "Personas compatibles con skill", "Brecha"]

        # Agrega las brechas por rol y skill para ver qué capacidades faltan con más frecuencia.
        agregadas = defaultdict(lambda: {"iniciativas": set(), "personas_compatibles": 0,
                                         "personas_con_skill": 0, "no_encontrada": False})
        for fila in requeridas_por_iniciativa:
            clave = (fila["Rol requerido"], fila["Skill requerida"])
            item = agregadas[clave]
            item["iniciativas"].add(fila["Nombre Proyecto"])
            item["personas_compatibles"] = max(item["personas_compatibles"], fila["Personas compatibles"])
            item["personas_con_skill"] = max(item["personas_con_skill"], fila["Personas compatibles con skill"])
            item["no_encontrada"] |= fila["Skill en matriz"] == "NO ENCONTRADA"
        brechas = []
        for (rol, skill), item in sorted(agregadas.items()):
            falta = item["no_encontrada"] or item["personas_con_skill"] == 0
            brechas.append({"Dominio": dominio, "Rol requerido": rol, "Skill requerida": skill,
                            "Iniciativas que la piden": len(item["iniciativas"]),
                            "Personas compatibles": item["personas_compatibles"],
                            "Personas compatibles con skill": item["personas_con_skill"],
                            "Brecha": "Sí" if falta else "No",
                            "Motivo": "Skill no existe en la matriz" if item["no_encontrada"] else
                                      ("Ninguna persona compatible la reporta" if item["personas_con_skill"] == 0 else "")})
        brechas_columnas = ["Dominio", "Rol requerido", "Skill requerida", "Iniciativas que la piden",
                            "Personas compatibles", "Personas compatibles con skill", "Brecha", "Motivo"]

        # Documenta los supuestos dentro del propio archivo para que el análisis
        # pueda revisarse sin consultar el código fuente.
        criterio = [
            {"Tema": "Filtro de personas", "Criterio": "Se incluye a quien tenga ENG o GRM como token del campo Dominio; los dominios combinados también cuentan."},
            {"Tema": "Filtro de iniciativas", "Criterio": "Se incluyen filas de demanda con rol reconocido para el dominio. En ENG se rellenan los datos de contexto del proyecto que vienen en blanco en filas continuadas."},
            {"Tema": "Skill presente", "Criterio": "Se considera presente con nivel 1 - Básico, 2 - Intermedio, 3 - Avanzado o 4 - Experto."},
            {"Tema": "Skill ausente", "Criterio": "0 - Nulo se reporta como ausencia declarada; celda vacía se reporta aparte como Sin dato."},
            {"Tema": "Persona Lead", "Criterio": "Se usa Nivel del Rol Actual igual a Lead. Las skills aparecen con el nivel original de la matriz."},
            {"Tema": "Compatibilidad del rol", "Criterio": "Se mapea el rol de la iniciativa y de Staffing con mapeo_roles.py. Para Dev Int/Dev Ext también se filtra por tipo de recurso interno/externo."},
            {"Tema": "Brecha de iniciativa", "Criterio": "Se revisan skills obligatorias. Brecha = Sí cuando la skill requerida no existe como columna en la matriz o ninguna persona compatible la reporta con nivel positivo. No evalúa el nivel mínimo requerido."},
            {"Tema": "Brechas por rol", "Criterio": "Agrupa por etiqueta de rol y skill. Personas compatibles y con skill muestran el máximo observado entre las iniciativas agrupadas."},
            {"Tema": f"Iniciativas {dominio}", "Criterio": f"El análisis recibió {len(iniciativas_dom)} filas de demanda con rol reconocido para {dominio}; si es cero, el reporte queda solo con encabezados."},
        ]
        escribir_xlsx(OUTPUT_DIR / f"Analisis skills {dominio}.xlsx", [
            ("Personas", headers, personas_dom),
            ("Skills detalle", detalle_columnas, detalle),
            ("Skills por seniority", resumen_columnas, resumen_filas),
            ("Personas Lead", lead_columnas, lead_filas),
            ("Brechas iniciativas", requeridas_columnas, requeridas_por_iniciativa),
            ("Brechas por rol", brechas_columnas, brechas),
            ("Criterios", ["Tema", "Criterio"], criterio),
        ])

        print(f"{dominio}: {len(personas_dom)} personas, {len(leads)} Lead, "
              f"{len(iniciativas_dom)} filas de demanda con rol reconocido; resultados en {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
