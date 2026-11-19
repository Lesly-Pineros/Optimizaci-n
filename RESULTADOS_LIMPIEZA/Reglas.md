# Reglas del modelo MILP de asignación

## 1. Población y disponibilidad

1. **Pool base:** se parte de los profesionales de la Matriz de SkillsCOE. Un registro con rol nulo o basura, o sin Dominio asignado, queda excluido del pool elegible y se reporta con su motivo.
2. **Recursos en backlog:** una persona de SkillsCOE sin registro en Staffing Q4 se clasifica como `BACKLOG`. Puede ser candidata desde el inicio del horizonte si tiene rol válido, Dominio y tipo de recurso disponibles en la fuente; el modelo no inventa datos faltantes. Si no tiene Dominio asignado, queda excluida.
3. **Capacidad teórica:** cada persona puede ocupar una plaza de **1.0 FTE**. La disponibilidad empieza el día siguiente a su última `Fecha Fin Actual` registrada en Staffing Q4; antes de ese día no se asigna. Si la persona no tiene fecha fin utilizable, se considera disponible desde el inicio del horizonte y la fecha faltante queda visible en la salida.
4. **Proyectos simultáneos:** una persona puede participar en como máximo dos proyectos/mesas que se solapen en el tiempo. No se bloquea la segunda asignación por el solo hecho de que se cruce con la primera. Las plazas se modelan indivisibles de 1.0 FTE; esta versión no reparte el FTE entre proyectos. Dos proyectos no se consideran simultáneos si sus fechas no se traslapan.
5. **Horizonte:** el modelo debe usar las fechas de demanda y de liberación para Q1/Q2 que se estén planificando. La disponibilidad se calcula por período/intervalo de fechas; no se debe asumir que la capacidad de Q1 aplica sin cambios a Q2.

## 2. Plazas, familias de rol y filtros técnicos

6. **Demanda FTE:** cada unidad indicada en `Internal` o `External` representa una necesidad indivisible de 1.0 FTE. Cada plaza conserva Mesa, fechas, rol, seniority, tipo y skills obligatorias. La versión actual no divide una plaza entre varios profesionales.
7. **Filtro de iniciativas:** se recorre toda la matriz fuente después del encabezado; no se limita a un rango fijo. Solo entran a la demanda las filas con rol reconocido en `mapeo_roles.py`, seniority homologado y al menos una skill técnica/funcional obligatoria. Esta matriz pertenece al dominio ENG; las direcciones observadas FCP, Systems, Operaciones, AI TECH e Infraestructura se asignan a ENG. Los roles nuevos deben agregarse al catálogo antes de incluir sus plazas.
   - Mapeo de roles de demanda: `LT` y `Product Engineer (LT - IA)` → `LT`; `QA` y `QUALITY & CONTEXT` → `QA`; `Dev Int`, `Dev Ext`, `Desarrollador`, `Data ing`, `UX` y provisionalmente `PO` → `DEV`. Es independiente del mapeo de roles de personas.
8. **Familias de rol de personas (mapeo provisional BD2):** se aplica el catálogo `person_role_families` de `mapeo_roles.py`:
   - Desarrollador, Analista Técnico, Arquitecto y Funcional: `DEV`, `QA`.
   - QA/Testing: `QA`.
   - Service Reliability Lead, Tech Manager/Líder, Líder y Líder de Servicio: `LT`, `DEV`, `QA`.
   - Data Service, Infraestructura, Despliegues, Operativa y Otros: `DEV`, `QA`.
   - Rol nulo, `NA...` o `cod eliminado`: ninguna familia; se excluye del pool.
   - Los roles válidos que no aparecen en estas categorías se tratan provisionalmente como Otros (`DEV`, `QA`); cualquier rol nuevo debe revisarse en el catálogo.
   - La matriz SkillsCOE registra perfiles con seniority `Lead` en `Desarrollador`, `Arquitecto`, `Despliegues`, `Otros`, `SOFTWARE EXPERT` y `tech`. Esos perfiles suman `LT` a sus familias habituales solo cuando el seniority individual es `Lead`. Testing/QA permanece exclusivamente en `QA`; todas las plazas LT siguen requiriendo personal interno.
9. **Rol y seniority:** después de normalizar al catálogo acordado, el candidato debe ser elegible para la familia solicitada. Se admite seniority exacto, un nivel por encima o, como último recurso, un nivel por debajo. La prioridad del modelo es exacto > superior > inferior; solo usa una categoría adyacente en la escala `Junior < Mid-Level < Senior < Lead`. Un seniority vacío/no homologado o con diferencia mayor a un nivel no pasa el filtro. La prioridad de seniority se aplica después de maximizar la cobertura y antes de continuidad y afinidad.
10. **Filtro y puntaje técnico:** para cada vacante se cuentan únicamente las skills técnicas y funcionales marcadas como obligatorias. El candidato debe tener al menos **una** de esas skills registrada, sin filtrar por su nivel. Cada skill obligatoria registrada suma un punto; tener niveles más altos no da puntos adicionales. Las skills sugeridas no cuentan. Si la demanda no declara skills obligatorias, la vacante se reporta para corregir.
11. **Tipo de recurso y sustitución:** primero se busca una persona del tipo solicitado en `Internal` o `External`. Si no hay una candidata elegible de ese tipo, se permite usar una persona del tipo opuesto y se marca la sustitución en la salida. Esta flexibilidad aplica a todos los roles salvo `LT`; LT solo puede cubrirse con personal interno y no admite sustitución de tipo. Una marca de tipo desconocido no se convierte automáticamente en interna o externa.
12. **Dominio ENG:** toda persona asignada a estas iniciativas debe tener `ENG` entre sus dominios. Es un filtro duro, no una preferencia. La salida muestra el dominio de la persona asignada.
13. **Líder Técnico:** toda vacante `LT` solo puede asignarse a personal interno del banco y debe cumplir el seniority admitido. Aplica la misma regla de skills: al menos una skill obligatoria registrada, sin filtrar por nivel. Si el match está en Básico, la salida lo indica en `Comentario skills obligatorias`.

## 3. Cobertura y límites por Mesa

14. **Mezcla interna/externa:** se prioriza que cada grupo `Mesa` + `Foco estratégico (Fábrica)` incluya personas internas y externas, siempre que sea posible con la cobertura máxima. Si no hay una combinación viable de ambos tipos, se asigna el tipo disponible y se informa la mezcla incompleta. La recomendación de candidatos no cuenta como asignación.
15. **Clave de conteo de mesa:** para contar equipos se usa internamente la pareja `Mesa` + `Foco estratégico (Fábrica)`. Esta unión solo identifica el grupo; no se agrega como columna ni modifica los valores de origen en la entrega.
16. **Límite base y ampliación:** cada grupo tiene hasta **7 asignaciones base** y puede incorporar hasta **5 personas adicionales** de la lista de sugerencias cuando hay plazas restantes. Si la mesa solicita más de siete roles, el modelo debe probar y asignar candidatos sugeridos que cubran esas vacantes; no debe dejar plazas vacantes solo por alcanzar siete asignaciones base. El máximo es 12 personas por grupo.
17. **Candidatos sugeridos opcionales:** el modelo muestra hasta **5 candidatos** por grupo, enfocados en `DEV` y `QA`, aunque estén comprometidos con otra iniciativa. La salida los etiqueta como opcionales e indica su asignación To-Be actual cuando la tengan. Una recomendación no asignada no cuenta como cobertura ni para la mezcla interna/externa. Solo al promoverla a una plaza se validan disponibilidad y máximo de dos proyectos simultáneos.
18. **Prioridad de grupos:** todos los grupos `Mesa` + `Foco estratégico (Fábrica)` tienen igual prioridad. La función objetivo no favorece un grupo por tener más plazas; compara cobertura proporcional por grupo antes de afinar preferencias individuales.

## 4. Objetivo del MILP

19. **Orden lexicográfico recomendado:** resolver en etapas para evitar que una preferencia blanda compense una restricción dura:
   1. Maximizar la cobertura factible respetando capacidad, número máximo de proyectos simultáneos, límite por grupo Mesa+Foco y presencia interna/externa.
   2. Con la cobertura fijada, favorecer una cobertura equilibrada entre grupos Mesa+Foco de igual prioridad.
   3. Con la cobertura máxima fijada, priorizar la presencia de internos y externos en cada grupo Mesa+Foco cuando sea viable.
   4. Maximizar asignaciones del tipo solicitado; usar el tipo opuesto como respaldo cuando no haya una combinación factible con el tipo requerido. LT no admite esta sustitución.
   5. Entre soluciones restantes, priorizar seniority exacto, después superior y finalmente inferior.
   6. Favorecer continuidad: para SDA se compara el valor de `SDATOOL - Proyecto` de `Matriz SkillsCOE enriquecida.csv` (tomado del registro de mayor dedicación cuando hay varios) contra `SDATOOL` de `ENG-Matriz Iniciativas con skills.csv`. Coinciden por texto o por identificador numérico completo aunque las letras sean distintas; no se comparan dígitos parciales. La Mesa actual se compara por separado con la columna `Mesa`.
   7. Después priorizar afinidad de Dominio y cantidad de skills obligatorias registradas.
20. **Puntaje de skills:** cada skill obligatoria registrada suma un punto de preferencia, sin importar su nivel. Las skills sugeridas no suman puntos.
21. **No asignados:** una plaza que no pueda cubrirse con candidatos elegibles queda vacante y se reporta con la causa principal (rol/familia, seniority, skills, tipo, Dominio, capacidad, simultaneidad o composición de Mesa+Foco).

## 5. Traducción a restricciones del MILP

- `x[p,s] ∈ {0,1}` = asignar a la persona `p` a la plaza completa `s`.
- Cada plaza recibe como máximo una persona; cada persona cubre como máximo una plaza por proyecto.
- Una persona puede cubrir plazas en hasta dos proyectos simultáneos. Las variables de actividad por persona/proyecto aplican este máximo; la asignación a dos proyectos solapados está permitida.
- Por grupo `Mesa` + `Foco estratégico (Fábrica)`, máximo siete personas únicas en la base. Tras maximizar cobertura, se favorece que estén presentes ambos tipos; si no es viable, se permite asignar el tipo disponible y se informa la mezcla incompleta.
- El tipo solicitado se prioriza. Los roles distintos de LT pueden usar el tipo opuesto como respaldo; LT debe ser interno y no permite sustitución.
- El objetivo prioriza cobertura proporcional por grupo `Mesa` + `Foco estratégico (Fábrica)` (prioridad igual), ajuste de seniority (exacto > superior > inferior), continuidad en la SDA/Mesa actual, afinidad de Dominio y cantidad de skills obligatorias cubiertas.
- Los filtros de rol, seniority adyacente como máximo, al menos una skill obligatoria registrada (sin filtro de nivel), tipo, fecha de liberación y Dominio asignado se aplican antes del objetivo.

## 6. Datos y decisiones que deben quedar explícitos para implementar

- Si se requiere asignación parcial de FTE para que una persona cubra simultáneamente dos proyectos; la versión actual trata cada plaza como indivisible de 1 FTE.
- Identificador y fechas de inicio/fin de cada proyecto, para determinar simultaneidad y disponibilidad en cada intervalo.
- Catálogo controlado que clasifica cada Rol Principal como LT, Testing, DEV/QA u otra familia; reglas para detectar valores basura.
- Correspondencia entre Dominio del profesional y Dominio requerido por cada Mesa, incluida la dirección de afinidad cuando no coinciden.
- Historial de SDA/Mesa actual y porcentaje de dedicación por persona; si hay más de una asignación actual, escoger la de mayor dedicación para continuidad.
- Pesos o prioridad lexicográfica entre continuidad, afinidad de Dominio y afinidad de Mesa actual. Las reglas indican preferencia, pero no fijan pesos numéricos.
- Texto recibido incompleto: “Los profesionales en la Matriz de Skills que no fig…”. Se interpreta únicamente la regla completa de recursos sin registro en Staffing Q4 indicada en el punto 2; cualquier condición adicional de la frase cortada requiere el texto completo.
