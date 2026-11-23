# Reglas aplicadas en limpieza y asignación

La especificación completa y el orden de ejecución están en
`MODELOS_ASIGNACION.md`.

1. **Población:** SkillsCOE es la base. Se enriquece con Staffing Q4. Personas
   que no aparecen en Staffing se marcan `BACKLOG`; sin Dominio o sin tipo
   interno/externo válido no entran al pool. Roles vacíos, con `NA` al inicio o
   fuera del catálogo se excluyen y se reportan.
2. **Fuente de rol:** solo se usa SkillsCOE, columna E, `Rol Principal`.
   Desarrollador, Arquitecto y Despliegues → DEV; QA → QA; Tech Manager/Líder,
   Service Reliability Lead, Líder Primera Línea y PTL → LT. Personas e
   iniciativas se normalizan a los mismos tres valores. No se infiere LT por
   seniority.
3. **Dominios:** ENG y GRM se asignan juntos, con capacidad compartida. Otros
   dominios se pueden incorporar en `ASSIGNMENT_DOMAINS`.
   La limpieza conserva cada fila de vacante, pero escribe en `SDATOOL` la lista
   deduplicada de todas las SDA del grupo Mesa + Foco estratégico.
4. **Capacidad:** cada plaza equivale a 1 FTE y cada persona puede estar en un
   máximo de dos proyectos simultáneos, contando Staffing Q4. Se usa una
   capacidad teórica del 100%; la dedicación actual solo resuelve continuidad.
5. **Prioridad, después de maximizar cobertura:** (1) continuidad SDA incluida
   en la lista del grupo y coincidencia de Dominio; el modelo maximiza primero
   SDA y luego Dominio; (2) rol como filtro estricto y seniority exacto,
   superior y luego inferior; (3) skills: dos o más coincidencias, luego una y
   por último ninguna. Para varias asignaciones Staffing, continuidad usa la
   SDA de mayor dedicación.
6. **Skills:** cualquier nivel positivo cuenta. Una coincidencia en una demanda
   de dos o más skills recibe penalización de 5; ninguna, 10. El respaldo sin
   skills exige coincidencia de Dominio. QA puede no declarar skills y entra
   por su rol y seniority, con penalización de 10. No se aplica un umbral de
   nivel de skill según seniority.
7. **Seniority:** se permite una categoría adyacente como máximo. No se
   transforma el seniority en umbral de skill.
8. **LT y mezcla:** LT solo se asigna a personal interno del banco. Se maximiza
   la presencia interna y externa por mesa cuando ambos tipos están disponibles.
9. **Límite por mesa:** máximo siete asignaciones base y tres recomendaciones
   DEV/QA. Todos los grupos Mesa+Dominio tienen la misma prioridad proporcional
   de cobertura.
10. **Campos no usados como requisito:** no se modifican `Nivel general skills
    técnicos` ni `Nivel IA`; las skills sugeridas tampoco cuentan como match.
