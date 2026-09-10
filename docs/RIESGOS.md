# Riesgos y mitigaciones

Prioridad = probabilidad × impacto sobre el cierre mensual.

> Los nombres de módulo que se mencionan abajo (`transform/reps.py`, `validate/excepciones.py`…)
> son piezas del **motor a futuro**; hoy la validación vive en `notebooks/validacion.py` y el
> escritor del `.xlsm` está diseñado en `ARQUITECTURA.md` (fase F8). El plan por fases está en
> `REEVALUACION_2026-09.md` §8.

---

## R1 · La matriz tiene fórmulas vivas y macros VBA  ·  ALTO

**Qué pasa:** un `df.to_excel()` o un `openpyxl` mal configurado destruye las fórmulas de
rollup `NL:NO` de `General` y **borra el proyecto VBA** (las 4 macros `Exportar_*`), dejando
al analista sin la parte que hoy sí funciona.

**Evidencia:** verificado sobre julio — `NL#=NM#+NN#`, `NM#=COUNTIF(A#:NK#,"Cumple")`,
`NN#=COUNTIF(A#:NK#,"No cumple")`, `NO#=NM#/NL#`. `wb.vba_archive is not None`. Las hojas
`Informe_B_C/E/M/A` son `VLOOKUP(B7, General!$A:$AIC, <ordinal>, 0)` — dependen del **orden**
de columnas de `General`.

**Mitigación:**
- `load_workbook(plantilla, keep_vba=True, data_only=False)`.
- Escritura **celda a celda** en la ventana `A3:NK` únicamente.
- `limpiar_datos()` borra solo `A3:NK`; nunca `NL:NO`.
- Snapshot (hash) de todas las hojas `Informe_*`, de `NL:NO` y del `vbaProject.bin` antes de
  escribir; se re-verifica antes de `save()`. Si cambió algo no previsto → aborta sin guardar.
- Validación de estructura al inicio: si la fila 2 de `General` no coincide en nombre y orden
  con `config/matriz.yaml`, aborta (alguien movió una columna → hay que regenerar el catálogo
  y avisar).

---

## R2 · "Cumple / No cumple" no es una sola regla  ·  ALTO

**Qué pasa:** si se codifica "resultado ≤ meta → Cumple" a secas, fallan al menos 5 casos
reales y el analista pierde la confianza en la primera corrida.

**Variantes confirmadas en los datos de julio:**
| # | Situación | valor (RESULTADO) | cumplimiento |
|---|---|---|---|
| 1 | Regla general: dentro del umbral | resultado | `Cumple` |
| 2 | Meta = `SD` (días estancia) | resultado | `Cumple` (automático) |
| 3 | IPS en el informe, midió `NA` por falta de casos | `NA` | **`Cumple`** |
| 4 | `NA` en numerador **y** denominador (oxígeno) | `NA` | `NA` |
| 5 | Reporta ≥1 indicador de la 1552 | `Si` (meta `Reporta`) | `Satisfactorio` |
| 6 | Nota técnica cápita (informativo, sin deducciones) | resultado | `NA` |
| 7 | Componente no aplica al contrato | `NA` | `NA` |

**Mitigación:** `config/reglas_cumplimiento.yaml` — reglas ordenadas, primer match gana.
Cada contrato nuevo o ajuste de política es un cambio de YAML, **no un despliegue**.
Cobertura de las 7 variantes es criterio de salida de la Fase 2.

---

## R3 · Cruces por texto frágiles  ·  MEDIO-ALTO

**Qué pasa:** `Municipio` y `Nombre IPS` son llaves en financieros-cápita, nota técnica,
autorizaciones, extramuralidad, domiciliaria, salud oral, transporte, etc. Diferencias de
tildes/espacios/mayúsculas rompen el cruce silenciosamente (fila queda en `NA` sin avisar).

**Evidencia:** el FO-GC-47 lo advierte 4 veces ("sintaxis exacta"). En los datos, régimen
llega como `SUBSIDIADO Y CONTRIBUTIVO` combinado y hay que ponderar RS/RC.

**Mitigación:**
- `transform/normalizar.py` + `config/equivalencias.yaml` (normalización + diccionarios de alias).
- Todo texto que no cruza va al reporte de excepciones (`municipios_sin_equivalencia`,
  `ips_sin_equivalencia`) — el cruce nunca falla en silencio.
- El diccionario de alias se completa iterativamente con los primeros reportes.

---

## R4 · El REPS  ·  MEDIO

**Qué pasa:** la Resolución 1552 (banda más grande de la matriz, 15+ indicadores) se cruza
por **código REPS de sede**, no por NIT. Hoy Jonatan lo digita a mano.

**Matiz (verificado):** el `CONSOLIDADO...SIRECI` **ya trae** un `REPS` principal por contrato
(sirve para datos generales). Lo que falta automatizar es el **desglose por sede**: un contrato
→ N sedes, cada una con su código de habilitación, que reportan por separado.

**Mitigación:** `transform/reps.py` — resolvedor `NIT ↔ REPS ↔ sede ↔ municipio` alimentado
por: (a) `config/reps_maestra.csv` (overrides manuales), (b) cache local, (c) **REPS público**
(datos.gov.co / prestadores.minsalud.gov.co) filtrado por NIT. Sin cobertura suficiente, la
banda 1552 se marca para diligenciamiento manual y el resto del proceso continúa.

---

## R5 · Contratos desactualizados en los insumos  ·  MEDIO (no resoluble por código)

**Qué pasa:** las áreas reportan números de contrato vencidos. Caso real (transcripción):
Hospital General de Medellín reportado como `0039` cuando el vigente era `0451`. El cruce
falla y el indicador queda sin diligenciar; hoy se descubre el día 19.

**Mitigación:** no se corrige automáticamente. `validate/excepciones.py` emite
`contratos_vencidos` y `contratos_sin_cruce` **el día 10**, para que el analista los devuelva
al área con tiempo. `validar_vigencia: true` en `fuentes.yaml` para medicamentos y
extramuralidad (los que el instructivo marca).

---

## Riesgos secundarios

| Riesgo | Mitigación |
|---|---|
| `CATEGORÍA DEL CONTRATO` sin lista confiable (col mantenida a mano) | Bloqueo B1, ver REEVALUACION_2026-09.md §7; `equivalencias.yaml` marcado `pendiente_confirmar` |
| El instructivo FO-GC-47 tiene errores de copy-paste | `docs/ERRORES_INSTRUCTIVO_FO-GC-47.md`; las correcciones ya van en `fuentes.yaml` |
| Insumos que cambian de formato mes a mes | Cada loader valida encabezados esperados y falla explícito |
| Google Sheets (PGP, PQRSD) sin acceso vía API | Fase 0; alternativa: export manual a `.xlsx` en una carpeta acordada |
| Archivos individuales por IPS (PF, ayudas dx) | `consolidar_por_ips: true` — paso previo que arma un auxiliar |
| Metas que vienen en el archivo fuente y cambian por contratación | `meta_origen: fuente`; RIAS/PGP exige avisar si cambió un indicador antes del cargue |
