# Modelos de asignación Q1–Q2 2027

## Ejecución

```bash
python3 limpieza.py
python3 asignacion_milp.py
python3 asignacion_tensores_milp.py
```

`limpieza.py` primero ejecuta la extracción de Staffing Q4 desde `DATA/STAFFING.xlsx`,
cruza códigos con SkillsCOE y prepara la demanda. Ambos modelos leen los CSV de
`RESULTADOS_LIMPIEZA/` y escriben sus resultados en `RESULTADOS MODELOS/`.
La limpieza conserva solo iniciativas con rol reconocido en `mapeo_roles.py`,
seniority homologado y al menos una skill técnica/funcional obligatoria.
Recorre todas las filas de datos de la matriz, no un rango fijo; las direcciones
organizacionales observadas se normalizan al dominio `ENG` porque este archivo es
la matriz de iniciativas ENG.

## Reglas que comparten ambos modelos

`RESULTADOS_LIMPIEZA/Reglas.md` es la referencia operativa. Los dos modelos usan
la misma matriz de viabilidad y las mismas restricciones duras:

- Solo se asignan roles válidos, tipo requerido y personas con Dominio `ENG`.
  Las iniciativas también deben tener rol mapeado, seniority válido y al menos
  una skill obligatoria para entrar en la demanda.
  El seniority puede ser exacto, un nivel superior o, como último recurso, un
  nivel inferior (máximo un nivel de diferencia, en el orden Junior, Mid-Level,
  Senior, Lead). Entre soluciones con igual cobertura se prioriza exacto,
  después superior y al final inferior.
- Los perfiles de Testing solo cubren QA; los demás perfiles válidos cubren DEV
  y QA. Los roles líderes ya reconocidos pueden cubrir LT; además, los perfiles
  con seniority Lead de `Desarrollador`, `Arquitecto`, `Despliegues`, `Otros`,
  `SOFTWARE EXPERT` y `tech` también pueden cubrir LT. Esta ampliación depende
  del seniority individual y no aplica a Testing. LT sigue siendo interno.
  Para todos los roles basta una skill obligatoria registrada; el nivel no
  filtra. Si un match de LT está en Básico, la salida lo señala en el comentario
  de skills.
- Se prioriza el tipo solicitado (`Internal`/`External`). En roles distintos de
  LT, el tipo opuesto puede cubrir una plaza cuando no haya candidatos elegibles
  del tipo requerido; la salida marca la sustitución. LT sigue siendo interno y
  no admite este respaldo.
- El candidato debe tener al menos una skill técnica/funcional obligatoria
  registrada, sin importar el nivel. Entre candidatos con la misma continuidad
  y afinidad de Dominio, registrar más skills obligatorias aumenta el puntaje;
  el nivel no cambia el puntaje.
- La fecha de liberación se calcula con Staffing Q4. Se permiten hasta dos
  proyectos simultáneos por persona; una asignación solapada no se bloquea
  automáticamente. Cada plaza sigue siendo indivisible de 1.0 FTE.
- Cada grupo identificado por `Mesa` + `Foco estratégico (Fábrica)` tiene hasta
  siete asignaciones base más cinco personas sugeridas adicionales. Si quedan
  roles por cubrir, el modelo puede promover sugerencias elegibles si están
  disponibles; el límite de siete no deja vacantes por sí solo.
  La concatenación se usa internamente para contar y no aparece como columna.
  Todos los grupos tienen la misma prioridad.

El objetivo prioriza cobertura proporcional por grupo Mesa+Foco, ajuste de seniority,
continuidad en la SDA/Mesa actual (por texto o por número completo, tomando el
registro de mayor dedicación), afinidad de Dominio y cantidad de skills obligatorias
registradas. El nivel no cambia elegibilidad ni puntaje; las skills sugeridas no cuentan.

## Diferencia entre modelos

- `asignacion_milp.py` resuelve cobertura proporcional, volumen total, seniority,
  continuidad y Dominio en fases sucesivas.
- `asignacion_tensores_milp.py` construye matrices `V` (viabilidad) y `A`
  (afinidad) y las usa en la solución lexicográfica: `V` define los pares
  elegibles y `A` ordena continuidad SDA/Mesa antes de afinidad de Dominio.
  Ambos modelos suman un punto por cada skill obligatoria registrada, sin importar el nivel.

La demanda se limita a proyectos que se solapan con Q1 o Q2 de 2027. Cada plaza
de `Internal`/`External` representa 1.0 FTE indivisible; el modelo no reparte una
plaza entre varias personas. Tras maximizar cobertura, el modelo favorece grupos
con ambos tipos. Si no hay una combinación viable, puede asignar el tipo
disponible e informa la mezcla incompleta en las vacantes restantes.

## Staffing en la limpieza y en las salidas

Cada Excel tiene una sola hoja `Asignaciones`, una fila por plaza, ordenada por
Mesa. Conserva las columnas A:R de la matriz de iniciativas y después incluye los
campos de asignación, `Origen de asignación` (base o sugerido adicional), el
`Ajuste tipo de recurso` cuando se usa el tipo opuesto,
porcentaje de skills obligatorias cumplidas, las skills con match y las faltantes
en columnas separadas. Incluye `Ajuste seniority` para
identificar si quedó exacto, un nivel por encima o un nivel por debajo. Las hasta
cinco recomendaciones DEV/QA por grupo Mesa+Foco van en la última columna y se
muestran una vez por grupo, aunque las personas estén comprometidas en otra
iniciativa. La lista indica su asignación To-Be actual. Cuando una recomendación
cubre una vacante se incluye como asignación y se identifica como `Candidato
sugerido adicional`; su disponibilidad se vuelve a validar en el MILP.

El CSV enriquecido conserva `SDATOOL - Proyecto`, `Fecha Inicio`, `Fecha Fin
Actual`, Mesa y dedicación. También guarda las asignaciones Q4 estructuradas para
mantener SDA, fechas y dedicación emparejadas cuando una persona tiene más de un
registro. La entrega muestra los campos del registro de mayor dedicación para cada
persona asignada. Las recomendaciones que no se promueven no forman parte de la
asignación ni consumen capacidad. Las promovidas ocupan la plaza y deben respetar
los filtros de elegibilidad y capacidad.

## Archivos generados

- `RESULTADOS MODELOS/Asignacion_To-Be_Baseline.xlsx`
- `RESULTADOS MODELOS/Asignacion_To-Be_Tensores_Lexicografico.xlsx`
- `RESULTADOS MODELOS/Asignaciones MILP.csv`
- `RESULTADOS MODELOS/Plazas sin asignar MILP.csv`
- `RESULTADOS MODELOS/Similitud de roles.csv` (ahora contiene el mapeo a familia)
- `RESULTADOS MODELOS/Viabilidad_V.csv` (1 = par elegible)
- `RESULTADOS MODELOS/Afinidad_A.csv` (afinidad por continuidad y Dominio)
