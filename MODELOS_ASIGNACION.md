# Modelos de asignación Q1 2027

## Ejecución

```bash
python3 -m pip install -r requirements-milp.txt
python3 -m unittest discover -s tests -v
python3 asignacion_milp.py
python3 asignacion_tensores_milp.py
```

Ambos programas leen `RESULTADOS_LIMPIEZA/Matriz SkillsCOE enriquecida.csv` y
`RESULTADOS_LIMPIEZA/ENG-Matriz Iniciativas con skills.csv`; los resultados se
guardan en `RESULTADOS MODELOS/`. Las fechas no filtran iniciativas: se incluyen
todas las filas del bloque procesado con rol, seniority y al menos dos skills
declaradas. El Excel de cada modelo mantiene una fila por fila de iniciativa y
añade columnas multilínea para personas, códigos, rol, seniority, tipo y skills
coincidentes con su nivel (0–4).
La columna `Skills faltantes` lista, para cada persona asignada, las skills
técnicas y funcionales obligatorias cuyo nivel está por debajo del esperado
para el seniority de la plaza; muestra también el nivel actual y el requerido.

## Archivos generados

- `RESULTADOS MODELOS/Asignacion_To-Be_Baseline.xlsx`
- `RESULTADOS MODELOS/Asignacion_To-Be_Tensores_Lexicografico.xlsx`
- `RESULTADOS MODELOS/Viabilidad_V.csv` (1 = par elegible)
- `RESULTADOS MODELOS/Afinidad_A.csv` (puntaje continuo 0–1)

El baseline reutiliza el MILP de una fase de `asignacion_milp.py`: maximiza
cobertura y usa el ajuste de skills para desempatar. El modelo de tensores crea
V y A antes de llamar al solver y optimiza en dos fases: (1) cobertura máxima,
(2) máximo número de proyectos cerrados con esa cobertura fijada; después
optimiza afinidad manteniendo ambas metas. El límite de siete personas por mesa
y el uso de una plaza por persona se aplica normalmente. Si ese límite deja
demanda pendiente, se permite como máximo una segunda plaza para las personas
estrictamente necesarias; el número de personas duplicadas se minimiza y la
excepción queda marcada en `Alerta de capacidad`.

## Interpretación de reglas y presupuesto

`Rglas.md` establece seniority exacto como requisito duro y es la referencia
operativa implementada por ambos modelos. `SUPUESTOS.md` describe otra versión
del borrador: seniority blando, familias de rol y elegibilidad más amplia. Esas
reglas contradicen `Rglas.md`; por ello no se mezclan silenciosamente. En la
afinidad A se conserva la ponderación 80% skills / 20% seniority; como el
seniority ya debe coincidir exactamente, su componente vale 1 para cada par
viable y no cambia el orden. Si Planeación confirma seniority blando, se debe
relajar ese filtro en el modelo de tensores.

Los archivos disponibles no contienen tarifas, costos por persona ni topes
CapEx/OpEx. El segundo modelo usa una penalización pequeña por recurso externo
como proxy de costo, solo para desempatar soluciones con la misma afinidad.
No representa ni valida un presupuesto monetario. Los roles no exactos ni
explícitamente equivalentes se reportan para revisión y no se habilitan
automáticamente.

El porcentaje de poda se calcula como `1 - pares viables / pares posibles`.
En la corrida actual fue 98,93%; el código reporta el valor observado y no
modifica reglas para alcanzar un porcentaje prefijado.
