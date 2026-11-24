# Asignación conjunta ENG y GRM

## Ejecución

```bash
python3 limpieza.py
python3 asignacion_milp.py
```

Los reportes auxiliares de códigos sin cruce y skills de iniciativas sin mapeo
se generan solo cuando se soliciten:

```bash
python3 reportes_limpieza.py
```

La limpieza toma Staffing Q4 como población y le agrega rol, seniority y skills
de SkillsCOE mediante el código de persona. Solo los códigos presentes en ambas
fuentes entran al pool; el Excel de diferencias se genera con el script auxiliar.
ENG se lee desde
`DATA/ENG-iniciativas.csv`; los demás dominios se leen de hojas con nombre
`<DOMINIO> X Mesa(s)` en `DATA/Iniciativas.xlsx`. Cada salida de iniciativas
incluye la columna `Dominio`. Se mantienen las filas de vacantes y `SDATOOL` se
reemplaza por la lista deduplicada de SDA de cada grupo Mesa + Foco estratégico.

`mapeo_roles.py` normaliza ambos lados del cruce a DEV, QA o LT. Para personas,
el rol fuente es SkillsCOE, columna E (`Rol Principal`): Desarrollador,
Arquitecto y Despliegues se normalizan a DEV; QA a QA; Tech Manager/Líder,
Service Reliability Lead, Líder Primera Línea y PTL a LT. No se infiere LT por
seniority. Roles desconocidos o con NA al inicio se excluyen y se muestran en
`RESULTADOS_LIMPIEZA/Personas que requieren tratamiento.xlsx`; roles de demanda
no mapeados se registran en `RESULTADOS_LIMPIEZA/Roles sin mapeo.xlsx`.

## Reglas del modelo

- Se asignan ENG y GRM en una sola corrida, con capacidad compartida. Para
  incluir otro dominio, agrégalo a `ASSIGNMENT_DOMAINS` en
  `asignacion_milp.py`.
- Se excluyen del pool personas sin rol válido, sin Dominio o sin tipo interno/
  externo válido.
- Cada plaza es 1 FTE indivisible. Se consideran hasta dos proyectos/Mesas
  simultáneos por persona, contando asignaciones de Staffing que se traslapan
  con el horizonte.
- Por mesa y Dominio se asignan hasta siete profesionales base. Se sugieren
  hasta tres personas adicionales DEV/QA. Los grupos tienen la misma prioridad
  proporcional de cobertura.
- LT se cubre exclusivamente con personal interno del banco. Cuando hay
  candidatos de ambos tipos, el modelo maximiza la presencia interna y externa
  de cada mesa sin reducir la cobertura alcanzada.
- Tras maximizar cobertura y presencia interna/externa, el orden de prioridad es:
  1. SDA actual incluida en la lista `SDATOOL` del grupo Mesa + Foco;
  2. coincidencia de Dominio;
  3. rol como filtro estricto y seniority exacto, superior y luego inferior;
  4. skills: dos o más coincidencias, después una, después ninguna.
  Si una persona tiene varias asignaciones Staffing, para la continuidad se toma
  la SDA de mayor dedicación; todas las asignaciones se conservan para capacidad.
- Las skills se cuentan con cualquier nivel positivo. Una sola coincidencia en
  una vacante que solicita dos o más skills recibe una penalización de 5 puntos;
  ninguna coincidencia, 10 puntos. El respaldo sin skills exige coincidencia de
  Dominio. QA puede no declarar skills obligatorias y entra en ese respaldo.
  Seniority permite solo niveles adyacentes cuando no hay exacto que prevalezca.
- Skills sugeridas, `Nivel general skills técnicos` y `Nivel IA` no se modifican
  ni se usan como requisitos.

## Relación con los cuatro pasos propuestos

1. **Cruce de proyectos y skills:** sí. `limpieza.py` normaliza roles,
   seniority y skills, agrega las SDA por Mesa+Foco y prepara el pool de
   personas enriquecido con Staffing Q4.
2. **Viabilidad binaria:** las reglas sí se aplican, pero como pares elegibles
   que genera `eligibility()`; actualmente no se materializa ni exporta una
   matriz V completa con ceros y unos.
3. **Matriz A ponderada por nivel y tarifa:** no. El modelo actual no usa
   tarifa ni pondera skills por nivel 0–4. Cualquier nivel positivo cuenta y
   las preferencias se resuelven por fases lexicográficas: SDA, Dominio,
   seniority y skills.
4. **Solver MILP:** sí. SciPy/HiGHS elige pares persona-vacante con variables
   binarias y las restricciones de cobertura, simultaneidad, límite por Mesa,
   mezcla de recursos y LT interno.

## Salidas

La corrida principal escribe en `RESULTADOS MODELOS/` un XLSX To-Be, el detalle
CSV de asignaciones, el CSV de plazas sin cubrir y el reporte de mapeo de roles.
El XLSX contiene `Asignaciones` y una pestaña `Opciones` con hasta tres
candidatos DEV/QA por Mesa+Foco, sus skills coincidentes/faltantes y la
coincidencia como fracción entre 0 y 1.
