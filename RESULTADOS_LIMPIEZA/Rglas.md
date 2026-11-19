# Reglas del modelo MILP de asignación

## Alcance

1. La corrida corresponde a Q1 de 2027. Se conservan las iniciativas cuya fecha de inicio y fin incluye ese trimestre.
2. Para las iniciativas cuya `Dirección` es `Operaciones`, el dominio requerido es `ENG`.
3. Se supone que cada persona está disponible al 100% para Q1 de 2027, independientemente de su dedicación en Staffing Q4. Como cada plaza representa una persona a tiempo completo, cada persona puede cubrir como máximo una plaza en el trimestre.

## Requisitos de las plazas

4. Se modelan filas con rol, seniority y al menos dos skills declaradas entre skills técnicas obligatorias, funcionales obligatorias y sugeridas.
5. Los valores de `Internal` y `External` indican el número de personas requeridas de cada tipo; cada unidad crea una plaza.
6. El rol de iniciativa se compara con `Rol Principal` de SkillsCOE. Se aplican las equivalencias de rol explícitas en el código y las demás similitudes deben revisarse en el reporte.
7. La persona debe coincidir exactamente en seniority. Se favorece la cobertura de skills obligatorias; se admite cobertura parcial y se valoran las skills sugeridas.
8. El nivel esperado para contar una skill como cubierta depende del seniority: Junior = básico (1), Mid-Level = intermedio (2), Senior y Lead = avanzado (3).

## Restricciones de asignación

9. Para una iniciativa de Operaciones, la persona debe tener `ENG` entre sus dominios de Staffing. Un dominio vacío no se considera disponibilidad universal; los registros sin dominio se separan para tratamiento y no entran al modelo.
10. El tipo debe coincidir: `Interno` con `Internal` y `Externo` con `External`. `Subsidiaria`, códigos duplicados, personas sin código y registros sin rol, seniority, cruce o dominio se reportan en el Excel de tratamiento.
11. Cada mesa puede tener como máximo siete personas distintas.
12. Los líderes técnicos (`LT`) deben ser personas internas de BBVA.

## Objetivo

13. El modelo maximiza la cantidad de plazas cubiertas sin incumplir las restricciones anteriores. Entre soluciones con igual cobertura, favorece más skills obligatorias cubiertas al nivel esperado, luego coincidencias parciales y finalmente skills sugeridas.
