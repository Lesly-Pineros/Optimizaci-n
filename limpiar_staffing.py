#!/usr/bin/env python3
"""Extrae las asignaciones Q4 de Staffing desde el libro Excel original.

Lee directamente el formato XLSX con la biblioteca estándar, selecciona las
hojas Q4 de Systems, Data y Arquitectura, y consolida sus filas en un CSV.
Cada fila conserva tanto los campos de Staffing como el detalle de asignación.
"""

import csv
import json
import re
import unicodedata
import zipfile
from datetime import datetime, timedelta
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET


# Ubicación de los archivos de entrada y de la salida consolidada.
ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "DATA" / "STAFFING.xlsx"
RESULTS = ROOT / "RESULTADOS_LIMPIEZA"
OUTPUT = RESULTS / "Staffing Q4 limpio.csv"

# Espacios de nombres XML usados por el formato OpenXML de Excel.
NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkg": "http://schemas.openxmlformats.org/package/2006/relationships",
}
# Esquema estable del CSV: limpieza.py consume estos encabezados para enriquecer
# personas y conservar la asignación original en formato JSON.
HEADERS = [
    "codigo", "interno", "% dedicacion", "rol", "dominio", "direccion",
    "SDATOOL - Proyecto", "Fecha Inicio", "Fecha Fin Actual", "Proyecto",
    "Mesa actual", "Asignación Staffing", "fuente staffing",
]


def normalize(value):
    """Normaliza encabezados para encontrar alias aunque cambien tildes o caso."""
    value = unicodedata.normalize("NFKD", str(value or ""))
    value = "".join(char for char in value if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", value).strip().casefold()


def column_index(header, *names):
    """Devuelve el índice de una columna obligatoria entre varios nombres posibles."""
    normalized = [normalize(cell) for cell in header]
    for name in names:
        matches = [i for i, cell in enumerate(normalized) if cell == normalize(name)]
        if len(matches) == 1:
            return matches[0]
    raise ValueError(f"No se encontró una columna única entre: {', '.join(names)}")


def optional_column(header, *names):
    """Busca una columna opcional y devuelve None si no existe."""
    normalized = [normalize(cell) for cell in header]
    for name in names:
        matches = [i for i, cell in enumerate(normalized) if cell == normalize(name)]
        if matches:
            return matches[0]
    return None


def staffing_date(value):
    """Convierte números de fecha Excel a ISO; deja intactos otros formatos."""
    value = str(value or "").strip()
    if not value:
        return ""
    try:
        serial = float(value)
    except ValueError:
        return value
    if not (20000 <= serial <= 80000):
        return value
    return (datetime(1899, 12, 30) + timedelta(days=serial)).date().isoformat()


def cell_value(cell, shared_strings):
    """Lee el contenido de una celda, tanto texto en línea como compartido."""
    value = cell.find("main:v", NS)
    if cell.attrib.get("t") == "inlineStr":
        return "".join(node.text or "" for node in cell.findall(".//main:t", NS))
    if value is None or value.text is None:
        return ""
    if cell.attrib.get("t") == "s":
        return shared_strings[int(value.text)]
    return value.text


def worksheet_rows(archive, target, shared_strings):
    """Recorre las filas de una hoja y expande celdas vacías hasta su columna."""
    path = str(PurePosixPath("xl") / target)
    # Las relaciones del libro suelen dar la ruta desde xl/ o desde la raíz.
    if target.startswith("/"):
        path = target.lstrip("/")
    elif target.startswith("xl/"):
        path = target
    sheet = ET.fromstring(archive.read(path))
    for row in sheet.findall(".//main:sheetData/main:row", NS):
        cells = row.findall("main:c", NS)
        values = []
        last_column = 0
        for cell in cells:
            ref = cell.attrib.get("r", "")
            letters = re.match(r"([A-Z]+)", ref)
            if letters:
                column = 0
                for char in letters.group(1):
                    column = column * 26 + ord(char) - ord("A") + 1
                values.extend([""] * max(0, column - last_column - 1))
                last_column = column
            values.append(cell_value(cell, shared_strings))
        yield values


def read_shared_strings(archive):
    """Carga la tabla opcional de cadenas compartidas del archivo XLSX."""
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return ["".join(node.text or "" for node in item.findall(".//main:t", NS))
            for item in root.findall("main:si", NS)]


def extract_csv_source():
    """Normaliza el Staffing tabular cuando el archivo fuente es CSV.

    En algunas entregas DATA/STAFFING.xlsx contiene realmente texto CSV con
    encabezados en español. Se mapea ese esquema al formato común de salida y se
    conserva una asignación JSON por fila fuente.
    """
    with SOURCE.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        header = reader.fieldnames or []
        source_rows = list(reader)

    columns = {
        "codigo": column_index(header, "Código", "Codigo", "codigo"),
        "interno": column_index(header, "Interno", "Interno/ Externo", "Tipo de Recurso"),
        "% dedicacion": column_index(header, "% Dedicación", "% dedicacion", "FTE Proy"),
        "rol": column_index(header, "Rol", "Rol homologado"),
        "dominio": column_index(header, "Dominio"),
        "direccion": optional_column(header, "Dirección", "Direccion", "Dirección SDA"),
        "sda": optional_column(header, "SDATOOL - Proyecto", "SDATOOL", "SDA"),
        "fecha_inicio": optional_column(header, "Fecha Inicio"),
        "fecha_fin": optional_column(header, "Fecha Fin Actual"),
        "proyecto": optional_column(header, "Proyecto"),
        "mesa": optional_column(header, "Mesa fábrica 2026", "Mesa Fábrica", "Mesa Asignada"),
    }

    combined = []
    for row in source_rows:
        get = lambda field: (row.get(header[columns[field]]) or "").strip() \
            if columns[field] is not None else ""
        code = get("codigo")
        if not code:
            continue
        sda = get("sda")
        project = get("proyecto")
        assignment = {
            "sda": sda,
            "proyecto": project,
            "mesa": get("mesa"),
            "fecha_inicio": get("fecha_inicio"),
            "fecha_fin_actual": get("fecha_fin"),
            "dedicacion": get("% dedicacion"),
            "fuente": SOURCE.name,
        }
        combined.append([
            code, get("interno"), get("% dedicacion"), get("rol"), get("dominio"),
            get("direccion"), sda, get("fecha_inicio"), get("fecha_fin"), project,
            get("mesa"), json.dumps(assignment, ensure_ascii=False), SOURCE.name,
        ])

    RESULTS.mkdir(exist_ok=True)
    with OUTPUT.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(HEADERS)
        writer.writerows(combined)
    print(f"{len(combined)} filas de Staffing desde CSV; salida: {OUTPUT}")


def extract_q4():
    """Extrae y consolida Staffing Q4 de las tres direcciones requeridas.

    Identifica columnas mediante alias por hoja, convierte fechas serializadas,
    construye un objeto JSON por asignación y valida que existan las tres hojas.
    """
    # Algunas entregas guardan un CSV con extensión .xlsx. Detectamos el formato
    # real del archivo para normalizarlo por la ruta apropiada.
    if not zipfile.is_zipfile(SOURCE):
        extract_csv_source()
        return

    RESULTS.mkdir(exist_ok=True)
    with zipfile.ZipFile(SOURCE) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {item.attrib["Id"]: item.attrib["Target"]
                   for item in relationships.findall("pkg:Relationship", NS)}
        shared_strings = read_shared_strings(archive)
        combined = []
        found = set()

        # Recorre todas las pestañas y descarta las que no sean del trimestre Q4.
        for sheet in workbook.findall("main:sheets/main:sheet", NS):
            sheet_name = sheet.attrib["name"]
            if not re.search(r"\bq4\b", sheet_name, flags=re.IGNORECASE):
                continue
            name = normalize(sheet_name)
            # Los encabezados varían entre áreas; los alias se adaptan a cada una.
            if "system" in name:
                direction = "system"
                aliases = {
                    "codigo": ("codigo", "código"),
                    "interno": ("interno", "interno/ externo"),
                    "% dedicacion": ("% dedicacion", "dedicacion q", "fte proy", "% total x usuario"),
                    "rol": ("rol",), "dominio": ("dominio",),
                    "direccion": ("dirección",),
                }
            elif "data" in name:
                direction = "data"
                aliases = {
                    "codigo": ("codigo", "código"),
                    "interno": ("interno/ externo", "interno"),
                    "% dedicacion": ("fte proy", "dedicacion q", "% total x usuario", "% dedicacion"),
                    "rol": ("rol",), "dominio": ("dominio",),
                    "direccion": ("dirección",),
                }
            elif "arquitectura" in name:
                direction = "arquitectura"
                aliases = {
                    "codigo": ("codigo", "código"),
                    "interno": ("interno", "interno/ externo"),
                    "% dedicacion": ("% dedicacion", "dedicacion q", "fte proy", "% total x usuario"),
                    "rol": ("rol",), "dominio": ("dominio",),
                    "direccion": ("dirección",),
                }
            else:
                continue

            rows = iter(worksheet_rows(archive, targets[sheet.attrib[f"{{{NS['rel']}}}id"]], shared_strings))
            try:
                header = next(rows)
            except StopIteration:
                raise ValueError(f"La hoja '{sheet_name}' está vacía")
            indices = {field: column_index(header, *names)
                       for field, names in aliases.items() if field != "direccion"}
            indices["direccion"] = optional_column(header, *aliases["direccion"])
            sda_indices = [optional_column(header, alias) for alias in
                           ("SDATOOL - Proyecto", "SDATOOL", "SDA", "Programa SDA")]
            sda_indices = [index for index in sda_indices if index is not None]
            start_index = optional_column(header, "Fecha Inicio")
            end_index = optional_column(header, "Fecha Fin Actual")
            project_index = optional_column(header, "Proyecto")
            mesa_index = optional_column(header, "Mesa fábrica 2026", "Mesa Fábrica", "Mesa Asignada")
            # Conserva filas que tengan al menos un dato identificador y agrega
            # la asignación completa a la salida junto a las columnas comunes.
            for row in rows:
                row += [""] * max(0, len(header) - len(row))
                selected = [row[index] if index is not None and index < len(row) else ""
                            for index in indices.values()]
                if not any(value.strip() for value in selected):
                    continue
                get = lambda index: row[index] if index is not None and index < len(row) else ""
                sda = next((get(index).strip() for index in sda_indices if get(index).strip()), "")
                start_date = staffing_date(get(start_index))
                end_date = staffing_date(get(end_index))
                project = get(project_index).strip()
                mesa = get(mesa_index).strip()
                assignment = {
                    "sda": sda,
                    "proyecto": project,
                    "mesa": mesa,
                    "fecha_inicio": start_date,
                    "fecha_fin_actual": end_date,
                    "dedicacion": selected[list(indices).index("% dedicacion")],
                    "fuente": sheet_name,
                }
                combined.append(selected + [sda, start_date, end_date, project, mesa,
                                             json.dumps(assignment, ensure_ascii=False), sheet_name])
            found.add(direction)

        if found != {"system", "data", "arquitectura"}:
            missing = sorted({"system", "data", "arquitectura"} - found)
            raise ValueError(f"No se encontraron las tres hojas Q4 requeridas; faltan: {', '.join(missing)}")

    with OUTPUT.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(HEADERS)
        writer.writerows(combined)
    print(f"{len(combined)} filas de Staffing Q4; salida: {OUTPUT}")


if __name__ == "__main__":
    # Permite usar el módulo desde limpieza.py o ejecutarlo de manera directa.
    extract_q4()
