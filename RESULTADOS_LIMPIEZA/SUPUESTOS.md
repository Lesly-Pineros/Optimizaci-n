# Supuestos del Borrador — PoC Asignación de Talento 2027

> Estos supuestos resuelven provisionalmente las discrepancias detectadas entre el
> contexto (steering + handoff) y la data real (`data/DATA1_personas.xlsx`,
> `data/DATA2_demanda.xlsx`). Son **decisiones del borrador**, no reglas de negocio
> confirmadas. Cada uno debe revisarse cuando Planeación entregue las reglas reales
> (Bloque 2). Todos se aplican en código y se referencian por ID.

Fecha: generado para el primer borrador funcional (datos reales como fuente, no dummy).

---

## A. Escala de skills (discrepancia #1 y #2)

- **BA1 — Rango 0–4, no 0–3.** La data trae un quinto nivel `4 - Experto`. Se adopta la
  escala observada **0–4** como oficial del borrador (en vez de forzar 0–3). El contexto
  (S12) asumía 0–3; se deja constancia de que la data lo contradice.
  *Impacto:* el `fit` de skills se gradúa 0–4. Si Planeación confirma 0–3, basta con
  fusionar Avanzado+Experto (mapa `4→3`).
- **BA2 — Parseo de prefijo.** Las celdas de skill vienen como texto `"<n> - <etiqueta>"`.
  Se extrae el entero inicial. Mapa: `0 - Nulo→0`, `1 - Básico→1`, `2 - Intermedio→2`,
  `3 - Avanzado→3`, `4 - Experto→4`.
- **BA3 — Celdas de skill no numéricas → 0.** Algunas celdas contienen texto libre
  ("ASO - Control M - DATAX", "IA, LRBA", etc.) o están vacías. Toda celda de skill que
  no parsee a un entero 0–4 se trata como **0 (sin skill)**. Se cuenta cuántas se
  descartan y se reporta.

## B. Staff ausente (discrepancia #3 y #4)

- **BB1 — No hay base Staff (Q3 2026).** Solo existen Skills y Demanda. En el borrador:
  - **Capacidad = 1.0 FTE libre para todas las personas** (no hay dato de compromisos
    vigentes; D1 sin resolver).
  - **Continuidad equipo-fábrica desactivada** (peso 0): requiere histórico persona↔mesa
    que no existe (S5). El término queda en la función objetivo con coeficiente 0, listo
    para activarse cuando llegue el histórico.

## C. Seniority de persona sucio (discrepancia #5)

- **BC1 — Vocabulario controlado de seniority.** Valores válidos:
  `Lead`, `Senior`, `Mid-Level`, `Junior`. Se normaliza por coincidencia exacta
  (case-insensitive, trim).
- **BC2 — Seniority inválido → persona NO elegible como LT, pero sí para Dev/QA.**
  Las filas con seniority basura ("NA ...", "cod eliminado", nulo, etc.) conservan su
  skill pero se les asigna seniority `DESCONOCIDO`. Una persona con seniority
  `DESCONOCIDO` **no puede ser líder técnico** (la regla S4 exige Lead explícito) pero
  **sí puede cubrir slots Dev/QA** si su rol lo permite. Se reporta cuántas quedan en
  cada categoría.

## D. Rol de persona sucio y mapeo a roles del modelo (discrepancia #6)

- **BD1 — Vocabulario de rol de slot (DATA2):** `LT`, `Dev Int`, `Dev Ext`, `QA`.
  Se normaliza a tres **familias de rol de slot**: `LT`, `DEV`, `QA`.
  (`Dev Int`/`Dev Ext` → `DEV`; la marca int/ext se maneja aparte, ver BE1.)
- **BD2 — Mapeo provisional rol-persona → familias que puede cubrir** (S2/D3, provisional):
  | `Rol Principal` (persona, normalizado) | Puede cubrir |
  |---|---|
  | Desarrollador / Analista Tecnico / Arquitecto / Funcional | DEV, QA |
  | QA / Testing | QA |
  | Service Reliability Lead / Tech Manager/Líder / Lider / Lider servicio | LT, DEV, QA |
  | Data Service / Infraestructura / Despliegues / Operativa / Otros | DEV, QA |
  | (rol nulo o basura "NA...", "cod eliminado") | NINGUNA (excluida del pool) |
  - **BD3 — Jerarquía por seniority para LT:** cubrir un slot `LT` exige además
    `Seniority == Lead` **e** `Interno` (regla dura S4), con independencia del rol
    principal. Es decir, BD2 habilita la *familia* LT, pero el filtro duro de líder
    técnico manda.
  - **BD4 — Service Reliability Lead (D3):** provisionalmente se trata como perfil capaz
    de LT/DEV/QA (ver BD2). Mapeo real pendiente de Planeación.

## E. Interno / Externo y mezcla por slot (discrepancia #10)

- **BE1 — Interno vs externo de la persona:** `Tipo de Recurso == "Interno"` → interno;
  `Externo` y `Subsidiaria` → **externo** (S15: interno = BBVA; externo = otra empresa).
  Se declara que *Subsidiaria* se cuenta como externo en el borrador.
- **BE2 — Marca del slot:** `Internal == 1` → el slot debe cubrirse con persona interna;
  `External == 1` → con persona externa. Es exacta por slot (S8). Se asume que cada slot
  tiene exactamente una de las dos marcas (se valida; si ambas o ninguna, se reporta).

## F. Grano y agrupación de la Demanda (discrepancia #8 y #9)

- **BF1 — 1 fila = 1 slot** (S11, confirmado por la data: 28 filas = 28 slots).
- **BF2 — Reconstrucción de proyecto por forward-fill.** Las columnas de proyecto
  (`Mesa`, `Nombre Proyecto`, `Dirección`, `Estado`) solo aparecen en la primera fila de
  cada bloque (celdas combinadas en Excel). Se hace *forward-fill* para propagar la
  identidad del proyecto a cada slot de su bloque.
- **BF3 — Clave de proyecto:** como `SDATOOL` está casi vacío y solo trae flags
  `FVAL 1/2` (= flags F-Val, S17 ignorados), el proyecto se identifica por **`Mesa`**
  (Mesa 1–4). Cada mesa = un equipo/proyecto a componer.
- **BF4 — Composición por proyecto (S18), corregido.** S18 asumía 1 LT + 3 DEV + 3 QA.
  La demanda REAL por mesa es **1 LT + 5 DEV (3 Ext + 2 Int) + 1 QA = 7 FTEs**. Ambas
  suman 7, pero la mezcla difiere.
  - **Decisión [BF4']:** la **demanda define los slots a cubrir**; la composición no
    recorta slots válidos de la demanda. Topes duros que sí se imponen: **1 LT por
    proyecto** (S4) y **≤ 7 FTEs por equipo** (S18).
  - Los topes por familia DEV/QA (3/3) quedan como **referencia a confirmar con
    Planeación**, no como recorte (si se impusieran, dejarían 8 slots DEV
    forzosamente vacíos aunque haya candidatos).

## G. Duplicados de llave (discrepancia #7)

- **BG1 — `Código` duplicado.** Hay 4 códigos repetidos en DATA1. Se conserva la
  **primera** aparición de cada `Código` y se descartan las demás; se reporta cuántas se
  eliminaron. (Criterio provisional; Planeación debe indicar cuál registro es el válido.)

## H. Objetivo (pesos del borrador)

- **BH1 — Fit de skills:** suma de niveles (0–4) de la persona sobre las skills
  **obligatorias** (técnicas + funcionales) del slot, normalizada. Es el término
  dominante del objetivo.
- **BH2 — Prioridad de dominio uniforme** (S7 provisional): peso igual para todos los
  proyectos/mesas.
- **BH3 — Penalización por slot no cubierto:** constante por slot, de modo que el modelo
  prefiera cubrir a dejar vacío, pero pueda dejar vacío si no hay candidato elegible
  (cobertura parcial, S6).
- **BH4 — Continuidad:** coeficiente 0 (ver BB1).

## I. Elegibilidad de skills (S3 provisional)

- **BI1 — Skills obligatorias NO son umbral duro en el borrador.** S3 (umbrales por rol)
  es provisional. Las habilidades obligatorias de la demanda entran como **puntaje blando**
  (fit), no como filtro duro de elegibilidad. Lo único duro es: familia de rol compatible
  (BD2), interno/externo (BE2), y para LT la regla S4 (BD3). Así se evita declarar
  infactibles proyectos por falta de un skill exacto antes de tener los umbrales reales.

## J. Seniority como preferencia blanda en el fit (decisión del usuario)

- **BJ1 — El fit combina skills y seniority.** La demanda (DATA2) especifica un seniority
  por slot (`Junior`/`Mid-Level`/`Senior`). Antes el modelo lo ignoraba; ahora entra como
  **preferencia blanda** (no filtra, no es regla dura). Fórmula:

  ```
  fit = W_FIT_SKILLS · fit_skills + W_FIT_SENIORITY · fit_seniority
  W_FIT_SKILLS = 0.8   W_FIT_SENIORITY = 0.2   (suman 1 → fit en [0,1])
  ```

  - **`fit_skills`** ∈ [0,1]: nivel medio de la persona sobre las skills obligatorias del
    slot / nivel máximo (4). Sin skills obligatorias → 0.5. (= el fit anterior, BH1.)
  - **`fit_seniority`** ∈ [0,1]: cercanía entre el seniority de la persona y el pedido:
    | Relación | Puntaje | Razón |
    |---|---|---|
    | coincide | 1.00 | encaje exacto |
    | por encima (sobrecalificado) | 0.75 | sirve, pero desaprovecha talento |
    | por debajo | 1.00 − 0.25·(niveles por debajo), piso 0 | riesgo creciente |
    | sin dato | 0.50 | neutro |

- **BJ2 — Pesos provisionales.** `0.8 / 0.2` es un supuesto del borrador: las skills
  dominan (núcleo del encaje técnico), el seniority desempata. Pendiente de que Planeación
  defina la ponderación real y si "sobrecalificado" debe premiarse o penalizarse (hoy:
  penalización suave). Ajustable en `config.py`.

---

## Filtros duros efectivos del borrador (resumen)

Una persona `p` es **elegible** para un slot `s` si y solo si:
1. `p` tiene rol principal válido (no basura) → BD2 la habilita para la familia de `s`.
2. La familia de rol del slot `s` ∈ familias que `p` puede cubrir (BD2).
3. int/ext: `p` coincide con la marca del slot (BE1/BE2).
4. Si `s` es `LT`: `p` es `Lead` **e** `Interno` (S4 / BD3).
5. Capacidad de `p` = 1.0 (BB1), consumida como ≤ 1 FTE en total (restricción de capacidad).

Términos blandos: fit = skills (BH1) + seniority (BJ1); prioridad uniforme (BH2),
−penalización no cubierto (BH3), continuidad=0 (BH4).
