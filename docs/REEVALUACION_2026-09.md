# Re-evaluación del proyecto — septiembre 2026

> Disparador: se recibió la **data real** en `docs/Insumos/` (todos los insumos crudos de
> julio 2026) y el **objetivo** en `docs/Objetivo/` (la matriz de julio ya diligenciada),
> más `docs/Insumos/Indicaciones.txt` (Jonatan indica qué hoja usar de cada fuente).
>
> Este documento re-evalúa el alcance con esos datos en la mano: qué se confirmó, qué
> cambió, qué falta, y el plan para llegar a una corrida mensual repetible.
>
> **Objetivo final:** que cada mes, con los insumos de ese mes, el proceso **diligencie la
> matriz nueva** (`…_<MES>.xlsm`) y entregue un **reporte de control**, de modo que el
> analista pase de *armar* la matriz a **validar y ajustar** el resultado. La comparación
> contra julio de este documento es el paso previo para asegurar que lo que el proceso va a
> llenar coincide con el diligenciamiento manual.

---

## 0. TL;DR

| | |
|---|---|
| **Lo más importante** | Ya tenemos el **par de referencia de julio completo**: `docs/Insumos/…xlsm` viene con la hoja `General` **vacía** (plantilla de entrada) y `docs/Objetivo/…xlsm` viene **diligenciada** (657 contratos, 245.884 celdas con dato, 228.145 en bandas de indicadores). Esto **desbloquea la validación por regresión** (era el bloqueo §7.3 del README). |
| **Fuentes** | De 20 componentes: **16 con archivo presente y validado**, **1 hoja heterogénea sin automatizar** (Ayudas Diagnósticas), **1 no aplica** (Salud Mental), **2 en Google Sheets sin acceso** (PGP, PQRSD), **1 archivo faltante** (Tablero de Cápitas para nota técnica), **1 dudoso** (INFORME CÁPITA Y MOVILIDAD, no lo menciona `Indicaciones.txt`). |
| **Reconstrucción** | El notebook reconstruye hoy **10 componentes** con **~92 %** de concordancia (validación de fuentes) y, con el **escritor** (`notebooks/escritor.py`), **escribe la matriz de julio** —celda a celda, preservando `NL:NO`, `Informe_*` y VBA— con **~96 %** de concordancia contra la diligenciada a mano. Las diferencias están **clasificadas** (ver §6). |
| **Bloqueos que siguen** | (a) lista cerrada de `CATEGORÍA DEL CONTRATO` y su mapa a componentes que aplican; (b) resolvedor `NIT ↔ REPS ↔ sede` para la Resolución 1552; (c) fórmula exacta de ponderación RS/RC en financieros. |
| **Errores de `Indicaciones.txt` / `fuentes.yaml`** | 8 discrepancias con los archivos reales, ya corregidas en `config/fuentes.yaml` (ver §4). |

**Entregables de esta re-evaluación**

1. `config/fuentes.yaml` — mapa de fuentes **verificado archivo por archivo** (qué archivo, qué hoja, qué fila de encabezado, qué llave, qué columnas se esperan, a qué columna de `General` va). Es el que consume el notebook.
2. `notebooks/validacion.py` — motor: validación de esquema + extractores + comparador + escritor del reporte.
3. `notebooks/01_validacion_fuentes_y_concordancia.ipynb` — el notebook pedido: se le pasa
   la **ruta de la unidad compartida** (`RUTA_UNIDAD`), busca los insumos **recursivamente**,
   y entrega el **objeto reporte** (`salidas/validacion/<fecha>/reporte_validacion.xlsx`) con
   la revisión del análisis y de lo que concuerda. Hojas clave: **`que_revisar`** (acciones)
   y **`faltantes`** (qué fuente no se encontró o cambió de formato). No se rompe si la
   unidad no está conectada o falta un archivo: lo reporta y sigue.

---

## 1. Qué cambió respecto del plan original

| Antes (plan original) | Ahora |
|---|---|
| "Sin el mes de referencia completo no se puede validar nada." | **Disponible** para julio (salvo los 3-4 insumos que no llegaron). |
| "El instructivo no lista los valores de `CATEGORÍA DEL CONTRATO`." | Se ven en los datos: `BAJA COMPLEJIDAD`, `MEDIANA COMPLEJIDAD`, `MOS Y AYUDAS ORTOPEDICAS` en la columna I de `General` y `M` del SIRECI. **Sigue faltando** la lista cerrada y el mapa categoría→indicadores que aplican. |
| "Confirmar hoja exacta de cada fuente." | **Confirmado** con `Indicaciones.txt` + apertura de cada archivo (§3). |
| "REPS de sede: ¿tabla propia o REPS público?" | El `CONSOLIDADO…SIRECI` sí trae 1 REPS principal por contrato (col D). El desglose por sede para la 1552 **sigue sin resolver**. |
| Fuentes en `Z:` | Los insumos llegaron por carpeta local; el mapeo a `Z:` se hará al tener el acceso, pero **la estructura interna de cada archivo ya está caracterizada**. |

---

## 2. Inventario de lo recibido

`docs/Insumos/` — 29 archivos + subcarpeta `Reporte Ayudas DX/` (24 archivos).
`docs/Objetivo/` — 1 archivo: `7.SEGUIMIENTO CONTRACTUAL SAVIA PPAL_JULIO.xlsm` **diligenciado**.

Verificación de las dos copias del `.xlsm`:

| Copia | Hoja `General` | Interpretación |
|---|---|---|
| `docs/Insumos/7.SEGUIMIENTO…JULIO.xlsm` | 0 filas de datos, 0 celdas | **plantilla de entrada** (equivale al `.xlsm` del mes anterior con `A3:NK` limpio) |
| `docs/Objetivo/7.SEGUIMIENTO…JULIO.xlsm` | 657 contratos · 245.884 celdas · 228.145 en indicadores | **salida esperada** (la "verdad" de julio) |

Ambas conservan las 379 columnas, las hojas `Informe_*`, `Exportar`, auxiliares y el VBA (`keep_vba` sigue siendo obligatorio al escribir).

---

## 3. Mapa de fuentes confirmado (por componente)

Fuente de verdad para "qué hoja": `docs/Insumos/Indicaciones.txt`. Verificado abriendo cada archivo.
Detalle completo (columnas esperadas, fila de encabezado, llaves, destino en `General`): **`config/fuentes.yaml`**.

| # | Componente (banda General) | Archivo real | Hoja real | Llave de cruce | Estado |
|--:|---|---|---|---|---|
| 1 | Datos generales `A:AE` | `7_CONSOLIDADO FINANCIERA_CONTRATACION_SIRECI_JULIO_7_RED.xlsx` | `SIRECI` (fila 1; 658 filas reales) | N.º contrato | migración directa |
| 1b | ↳ complemento | `Contratos.xlsm` | `BD CONTRATOS 2019-2026` (fila 3) | N.º contrato | vigencia / prórrogas / complejidad |
| 2 | Auditoría calidad `AF:AG` | `AUDITORIA DE CALIDAD_MATRIZ_Act_Julio_2026.xlsx` | `Indicador` (fila 1: NIT · CONTRATISTA · FECHA · RESULTADO) | NIT | **reconstruible** (99,7 %) |
| 3 | Financieros `AH:AP` | `9.Consolidado indicadores por IPS JULIO 2026.xlsx` | `2026` (fila 1; filtrar `Periodo`) | NIT + Régimen | **reconstruible** (90,4 %) |
| 4 | Nota técnica `AQ:AS` | `Tablero de Seguimiento Ejecucion Capitas ENE-JUL 2026` | `resumen` | Modalidad + Municipio | **falta insumo** |
| 5 | Autorizaciones – Trámite `AT:AV` | `Trámite interno Julio 2026.xlsx` | `Julio` (encabezado doble, filas 1-2) | NIT + Municipio (+ régimen para elegir columna) | **reconstruible** |
| 6 | Autorizaciones – Marcación `AW:AY` | `Marcación prestación efectiva Julio 2026.xlsx` | `Julio` (encabezado doble) | NIT + Municipio (+ régimen) | **reconstruible** |
| 7 | Concurrencia `AZ:BE` | `Indicadores_Hospitalarios_Julio_2026.xlsx` | `IPS_Sin_Concurrencia` **+** `IPS_Concurrencia_CONEXIONBES` (fila 2) | NIT | **reconstruible** (100 %) |
| 8 | MIPRES `BF:BK` | `8. Entregas Efectivas…xlsx` (`Dinamica`, fila 5) **+** `8.Cumplimiento Junta de profesionales…xlsx` (`Dinamica`, fila 3) | NIT | **reconstruible parcial** (63 %) |
| 9 | Medicina domiciliaria `BL:CX` | `INDICADORES ATENCION DOMICILIARIA JULIO 2026.xlsx` (1 hoja por contrato) **+** `Indicadores Curativ Julio 2026.xlsx` (`CONTRATO 0227 2024`, `CONTRATO 0429 2025`) | N.º contrato | **pendiente** (layout por bloques) |
| 10 | Salud oral `CY:DP` | `INDICADORES ODONTOLOGICOS IPS SAN JOSE JULIO 2026.xlsx` | `JULIO` (fila 1; 1 sola fila de datos) | N.º contrato | **pendiente** |
| 11 | Transporte `DQ:DV` | `7. Indicador Transporte Asistencia Julio_2026.xlsx` | `Indicadores_Contratos_Evento` (fila 1) | N.º contrato | **reconstruible** (75 %) |
| 12 | Hogares de paso `DW:EE` | `INDICADORES HOGARES DE PASO JULIO 2026.xlsx` | `SUBSIDIADO` + `CONTRIBUTIVO` (**no tabular**, bloques verticales) | Nombre proveedor | **pendiente** |
| 13 | Oxígeno `EF:FC` | `INDICADOR PROGRAMA GESTION OXIGENO -SEGUIMIENTO A LA RED JULIO 2026.xlsx` | `Hoja1` (**2 bloques**: EVENTO fila 2, PGP fila 9) | N.º contrato | **pendiente** |
| 14 | MOS `FD:FL` | `REPORTE DE INDICADORES MOS Y AOM JULIO 2026 (supervision).xlsx` | `Hoja1` (fila 1; histórico 2022+, filtrar corte) | N.º contrato | **pendiente** (mapeo 6→3 indicadores) |
| 15 | Medicamentos Fénix `FM:GA` | `FO-RS-96 Formato reporte de indicadores para supervisión JULIO 2026.xlsx` | `FO-RS-96 Supervisión` (fila 6) | N.º contrato | **reconstruible** (95 %) |
| 16 | Resolución 1552 `GB:JA` | `07.CONSOLIDADO_R1552 JULIO 2026..xlsx` | `CONSOLIDADO` (**encabezado en fila 17**) + `METAS` (CUPS→meta) | **Código habilitación (REPS de sede) + CUPS** | **pendiente** (necesita resolvedor de sedes) |
| 17 | Tutelas `JB:JD` | `Base de Servicios de tutelas Julio 2026.xlsx` | `Detallado` (fila 1) | N.º contrato (con prefijo RS-/RC-) | **reconstruible** (100 %) |
| 18 | PQRSD `JE:JV` | Google Sheets "RESULTADOS CONTRACTUALES INDICADORES PQRSD 2026" | 6 hojas | N.º contrato | **falta acceso** |
| 19 | Salud mental `JW:KB` | — | — | — | **NA por definición** |
| 20 | Planificación familiar `KC:LC` | `INDICADORES IPS EXPERTASALUD…xlsx` (`RESULTADO INDICADORES 2025`) **+** `INDICADORES IPS SALUD REPRODUCTIVA…xlsx` (`Indicadores contrato 2025`) | N.º contrato (filtrar `MES DE REPORTE`) | **pendiente** (layout ancho por bloques de 4 col) |
| 21 | Ayudas diagnósticas `LD:NH` | carpeta `Reporte Ayudas DX/` (24 archivos, **estructura distinta cada uno**) | varias | N.º contrato / nombre IPS | **manual / heterogéneo** |
| 22 | Campañas extramural `NI:NK` | `Formato de indicadores extramuralidad julio 2026.xlsx` | `JULIO` (fila 1) | Municipio | **reconstruible** (97 %) |
| — | Financieros Cápita `AH:AP` | `INFORME CAPITA Y MOVILIDAD` | ? | Municipio | **dudoso** (no lo menciona `Indicaciones.txt`) |

---

## 4. Discrepancias detectadas (Indicaciones.txt ↔ fuentes.yaml ↔ archivos reales)

Ya corregidas en `config/fuentes.yaml`. Confirmar las marcadas ⚠️ con Jonatan.

| # | Componente | Decía | Es en realidad | Acción |
|--:|---|---|---|---|
| 1 | Concurrencia | hoja `IPS_Concurrencia_CONEXIONES` | hoja **`IPS_Concurrencia_CONEXIONBES`** (typo en el archivo) | usar el nombre real; además en esa hoja **las columnas L y M están invertidas** (L=denominador, M=numerador) |
| 2 | Transporte | hoja `Indicaddores_Contratos_Evento` | hoja **`Indicadores_Contratos_Evento`** (typo en `Indicaciones.txt`) | usar el nombre real |
| 3 | Domiciliaria/Curativa | hoja `CONTRATO 0426 2025` | hojas **`CONTRATO 0227 2024`** y **`CONTRATO 0429 2025`** | ⚠️ confirmar cuál(es) aplican a julio |
| 4 | Financieros | `fuentes.yaml`: `9. Consolidado indicadores Financieros por IPS`, hoja `null` | archivo `9.Consolidado indicadores por IPS <MES> <AÑO>`, hoja **`2026`** | corregido |
| 5 | Auditoría calidad | `fuentes.yaml` espera `fecha de auditoría`, `% de calificación` | hoja `Indicador`: **`NIT` · `CONTRATISTA` · `FECHA` · `RESULTADO`** | corregido |
| 6 | Nota técnica | `Indicaciones.txt` lista `Tablero de Seguimiento Ejecucion Capitas` | **archivo ausente** en `docs/Insumos/` | ⚠️ pedir el archivo |
| 7 | Financieros Cápita / PGP / PQRSD | `INFORME CAPITA Y MOVILIDAD` + 2 Google Sheets | **ausentes** | ⚠️ acceso a Drive / carpeta compartida |
| 8 | Salud mental | banda `JW:KB` | **todo `NA`** (`Indicaciones.txt`: "ya no mandan indicadores") | marcar la banda como NA fija |

Errores del instructivo FO-GC-47 confirmados a la luz de los datos (ver `docs/ERRORES_INSTRUCTIVO_FO-GC-47.md`):

- Hogares de Paso: filtrar por `"Hogares de Paso"`, no `"TRANSPORTE ASISTENCIAL"`. ✅
- Planificación Familiar / Ayudas Diagnósticas: **no** filtran por `"MOS y ayudas ortopédicas"`. En julio, PF fueron 2 contratos concretos (0217-2022, 0443-2025). ✅

---

## 5. Hallazgos de estructura (los que rompen un lector ingenuo)

| Componente | Trampa | Cómo se maneja |
|---|---|---|
| **SIRECI** | `ws.max_row` reporta 1.048.548 (formato hasta el final). Datos reales: **658 filas**. Además hay hoja `SIRECI_JUNIO` (mes anterior). | contar filas con dato en col A; leer solo hoja `SIRECI` |
| **Resolución 1552** | Las primeras **16 filas** son el membrete FO-RS-90 + glosario de siglas. **Encabezado real = fila 17.** | `fila_encabezado: 17` |
| **Concurrencia** | 2 hojas a unificar; nombre real con typo (`CONEXIONBES`); columnas L/M invertidas en la hoja con auditoría; col D trae **varios contratos concatenados** | unir hojas, resolver por NIT, `contratos_en_celda()` |
| **MIPRES (`Dinamica`)** | Tabla dinámica: encabezado en 3 filas; columnas `Numerador/Denominador` repetidas por régimen (Contributivo / Subsidiado / en blanco) | leer con `header=fila 5`, agregar por régimen |
| **Trámite / Marcación** | Encabezado en 2 filas; `% TRÁMITE` / `% FECHA PRESTACIÓN` aparecen **3 veces** (subsidiado / contributivo / total); 1 hoja por mes | elegir la columna por el régimen del contrato |
| **Financieros (`9.Consolidado`)** | Una hoja por año (`2022`…`2026`); trae el **año completo**, hay que filtrar `Periodo == mes`; el archivo **ya trae Meta y Cumplimiento** por indicador | usar solo la hoja del año; filtrar mes; los Cumplimiento del archivo sirven para validar el motor de reglas |
| **Medicina domiciliaria** | 1 hoja por contrato; cada hoja = 12 meses en filas + separador de año en col E; patrón META/INDICADOR **intercalado y distinto** entre "atención" y "curativa"; el nº de contrato cambia de sufijo de año (`0351-2025`→`0351-2026`) | parser por hoja; normalizar nº de contrato base |
| **Hogares de paso** | **No es tabular**: ficha de indicadores arriba, luego 3 mini-tablas verticales `PROVEEDOR/MES/INDICADOR/NUMERADOR/DENOMINADOR/RESULTADO` | parser por bloques; cruce por nombre de proveedor (necesita alias) |
| **Oxígeno** | 2 bloques en una hoja (EVENTO fila 2, PGP fila 9); metas embebidas en celdas de datos (`≤ 8 Horas`, `>95 %`); mucho `N/A` por zona | parser de 2 bloques; regla `oxigeno_na_num_y_den` |
| **MOS** | Histórico desde 2022 (687 filas); valores `SR`/`CR`/`N/A`; NIT con sufijo `-1`; la fuente trae **6 indicadores**, `General` usa **3** | filtrar corte del mes; `norm_nit()`; ⚠️ confirmar mapeo 6→3 |
| **Medicamentos FO-RS-96** | Encabezado en fila 6; metas como texto (`2 DÍAS`, `48 HORAS`); `Mes Reportado` a veces trae una frase ("NO REPORTA INDICADOR PORQUE…") | `fila_encabezado: 6`; tratar frase como no-reporte |
| **Planificación familiar** | Layout ancho: bloques de 4 columnas por indicador (`INDICADOR n / META / NUMERADOR / DENOMINADOR`); filas futuras con `#DIV/0!`; encabezado en fila 1 (Expertasalud) o 2 (Salud Reproductiva) | filtrar `MES DE REPORTE`; parser de bloques de 4 |
| **Ayudas diagnósticas** | 24 archivos con estructuras **totalmente distintas**: unos con `DETALLE_1552_Art_2` (41-65 col), otros con hojas ad-hoc (`IMAGENES`, `Mamografias`, `INDICADORES UVDXM`…); 1 archivo `.xls` legacy | parser **por familia de plantilla**, no por archivo; fase posterior |
| **Tutelas** | `Indicador` ya viene calculado (razón ~1e-6); contrato con prefijo `RS-`/`RC-` | `norm_contrato()` conserva el prefijo para cruzar con `General` |

---

## 6. Resultado de la primera reconstrucción

Corrida del notebook sobre julio (`notebooks/01_validacion_fuentes_y_concordancia.ipynb`):

```
contratos en la matriz .............. 657
fuentes con esquema OK ............... 22 / 28 objetivos
fuentes FALTA_INSUMO ................. 4   (nota técnica, INFORME CÁPITA, PGP, PQRSD)
celdas reconstruidas ................ 9.147
celdas que coinciden ................ 8.420   (92,1 %)
```

Concordancia por banda:

| Banda | Celdas | Coinciden | % |
|---|--:|--:|--:|
| Auditoría calidad | 1.314 | 1.310 | 99,7 |
| Concurrencia | 544 | 544 | 100 |
| Tutelas | 178 | 178 | 100 |
| Campañas extramural | 382 | 370 | 96,9 |
| Medicamentos Fénix | 1.245 | 1.186 | 95,3 |
| Financieros | 3.816 | 3.450 | 90,4 |
| Indicadores de acceso | 1.271 | 1.130 | 88,9 |
| Transporte | 12 | 9 | 75,0 |
| MIPRES | 385 | 243 | 63,1 |

Clasificación de las **727 diferencias**:

| Clasificación | Celdas | Qué significa | Cómo se cierra |
|---|--:|---|---|
| `objetivo_NA_proceso_valor` | 254 | El analista dejó `NA` (el componente **no aplica** a ese contrato) pero el proceso produjo un valor porque el NIT/municipio cruzó | **mapa `CATEGORÍA DEL CONTRATO → componentes que aplican`** (bloqueo abierto) |
| `diferencia_numerica` | 234 | Coincide el concepto pero no el número: redondeo, **ponderación RS/RC**, o agregación distinta (sedes) | fórmula exacta de ponderación + regla de agregación por NIT |
| `llave_no_cruzada_o_dato_faltante` | 162 | El proceso dejó `NA` donde el analista puso valor | **alias** de municipio / nombre de IPS / NIT secundario |
| `diferencia_texto` | 77 | Formato (`100,0%` vs `1.0`) o **regla de cumplimiento** | normalización de salida + motor de reglas |

> Lectura: el 92 % "en bruto" ya es alto para 10 componentes sin el mapa de aplicabilidad
> ni los diccionarios de alias. Las 3 causas de fondo (aplicabilidad, ponderación, alias)
> explican la mayoría de las diferencias y son trabajo acotado, no rediseño.

---

## 7. Bloqueos que siguen abiertos

| # | Bloqueo | Impacto | Quién |
|--:|---|---|---|
| B1 | **Lista cerrada de `CATEGORÍA DEL CONTRATO`** (col I de `General`, la que el analista mantiene a mano) **y el mapa categoría → qué componentes le aplican** | Sin esto, ~254 celdas quedan como falso positivo (proceso pone valor donde debe ir `NA`). Es la llave de segmentación de todo. | Jonatan |
| B2 | **Resolvedor `NIT ↔ REPS ↔ sede ↔ municipio`** para la Resolución 1552 (banda `GB:JA`, la más grande) | La 1552 se cruza por código de habilitación de sede, no por NIT. Hoy Jonatan lo digita. | REPS público (datos.gov.co) + `config/reps_maestra.csv` |
| B3 | **Fórmula de ponderación RS/RC** en financieros-evento cuando el contrato es "SUBSIDIADO Y CONTRIBUTIVO" y la fuente llega separada | ~parte de las 234 diferencias numéricas | Jonatan / Financiera (¿promedio simple? ¿ponderado por facturación?) |
| B4 | **Insumos faltantes**: Tablero de Cápitas (nota técnica), INFORME CÁPITA Y MOVILIDAD, Google Sheets PGP y PQRSD | 4 bandas no reconstruibles hasta tenerlos | Jonatan / Yuli (acceso Drive) |
| B5 | **Ayudas Diagnósticas**: 24 plantillas distintas | Banda `LD:NH`. Requiere agrupar por familia de plantilla. | Trabajo de análisis (fase 5) |
| B6 | **Acceso a `Z:`** y a las rutas reales | Operación mensual (hoy se trabaja sobre copia local) | Yuli |

---

## 8. Plan de trabajo revisado (para corrida mensual repetible)

Cada fase deja algo **ejecutable** y **medible contra julio**.

| Fase | Entregable | Criterio de salida | Depende de |
|---|---|---|---|
| **F0 — hecho** | `config/fuentes.yaml` + `validacion.py` + notebook + este documento | El notebook corre sobre `docs/Insumos` y emite el reporte | — |
| **F8 — hecho** | `notebooks/escritor.py`: escribe `A3:NK` celda a celda, `keep_vba`, `verificar_integridad()` compara hojas preservadas + `NL:NO` + VBA contra la plantilla tras guardar | La matriz generada preserva fórmulas/plantillas/VBA (integridad OK) y llega a ~96 % vs julio en las 10 bandas con extractor | F0 |
| **F1 — Aplicabilidad** | `config/categorias.yaml`: lista cerrada de `CATEGORÍA` + mapa a componentes; función `aplica(componente, contrato)` | Los 254 `objetivo_NA_proceso_valor` bajan a < 20 | B1 |
| **F2 — Normalización + alias** | `transform/normalizar.py` real + `equivalencias.yaml` sembrado con los alias que salgan del reporte (`municipios_sin_equivalencia`, `ips_sin_equivalencia`) | Los `llave_no_cruzada` bajan a < 30 | reporte F0 |
| **F3 — Motor de reglas** | `transform/motor_reglas.py` implementado (las 7 variantes) + validado contra los `Cumplimiento` que ya traen `9.Consolidado` y `FO-RS-96` | 100 % de coincidencia en celdas `CUMPLIMIENTO` de los 10 componentes actuales | F2 |
| **F4 — Extractores restantes** | Domiciliaria, Salud Oral, Hogares de Paso, Oxígeno, MOS, Planificación Familiar | Cada banda ≥ 95 % vs julio | F1-F3 |
| **F5 — Resolución 1552** | `transform/reps.py` (resolvedor de sedes) + extractor 1552 | Banda `GB:JA` ≥ 95 % vs julio | B2 |
| **F6 — Ayudas Diagnósticas** | 1 parser por familia de plantilla + consolidador | Banda `LD:NH` ≥ 90 % de las IPS que reportaron | B5 |
| **F7 — Insumos faltantes** | Conectores Google Sheets (PGP, PQRSD) + lectores Tablero Cápitas / INFORME CÁPITA | Bandas de acceso/financieros/PQRSD completas | B4 |
| **F9 — Regresión formal** | umbral: `CUMPLIMIENTO` 100 %, `RESULTADO`/datos generales ≥ 99 %, cero `dato` sin explicar | Umbral alcanzado sobre julio (y un 2.º mes cuando llegue agosto) | F1–F4 |
| **F10 — Operación asistida** | corrida en paralelo al proceso manual 1-2 meses + acta de cambios entre versiones del entregable | El analista valida 2 cierres seguidos sin corrección relevante | F9 |

Orden recomendado: **F1 → F2 → F3** primero (suben la concordancia de lo que ya hay a ~99 % sin escribir un solo extractor nuevo), y en paralelo pedir B1/B4. El escritor (F8) ya está;
al agregar cada extractor (F4–F7) su banda se llena automáticamente.

---

## 9. Cómo se repite cada mes (runbook operativo)

1. Conectar la **unidad compartida** (disco `Z:` o el que sea).
2. Abrir `notebooks/01_validacion_fuentes_y_concordancia.ipynb`, ajustar en el bloque de
   parámetros:
   - `MES` = `AAAA-MM` del cierre;
   - `RUTA_UNIDAD` = raíz del disco compartido (los archivos se buscan **recursivamente**
     en todas las subcarpetas — no hay que listar cada ruta);
   - `SUBCARPETA` = subcarpeta del mes si aplica;
   - `MATRIZ_OBJETIVO` = el `.xlsm` de referencia (opcional: si no está, igual valida fuentes).
3. Ejecutar el notebook. **No se rompe** si la unidad no está o falta un insumo: lo reporta.
4. Revisar primero la hoja **`que_revisar`** y **`faltantes`** del reporte:
   - `FALTA_ARCHIVO` / `FALTA_HOJA` → el insumo llegó incompleto o con otro nombre → pedirlo.
   - `DRIFT` → un área **cambió una columna** de su reporte → avisar al área antes de seguir.
   - `FALTA_INSUMO` → fuente que no se entrega por este canal (Google Sheets, Tablero de Cápitas).
5. Revisar **`concordancia_por_banda`** y **`concordancia_detalle`**:
   - diferencias nuevas respecto del mes anterior → clasificar (dato / regla / llave).
6. El notebook genera `salidas/<mes>/7.SEGUIMIENTO…_<MES>.xlsm` + `reporte_llenado.xlsx`.
   Revisar `reporte_llenado.xlsx` → `resumen` (que la integridad diga `OK`) y `a_revisar`.
7. Abrir la matriz generada, completar/ajustar las bandas que falten, correr las macros como siempre.

El notebook **no reemplaza** el criterio del analista todavía: es el tablero que le dice
**qué mirar** y **dónde no coincide**, para que su revisión sea dirigida y no celda por celda.

---

## 10. Preguntas concretas para Jonatan

1. **`CATEGORÍA DEL CONTRATO`**: ¿lista cerrada de valores? ¿qué componentes de indicadores le aplican a cada uno? (bloqueo B1)
2. **Curativa**: en julio, ¿aplican las hojas `CONTRATO 0227 2024` **y** `CONTRATO 0429 2025`, o solo una?
3. **Financieros "SUBSIDIADO Y CONTRIBUTIVO"**: ¿cómo se pondera exactamente RS/RC? (promedio simple / ponderado por facturación / por población)
4. **MOS**: la fuente trae 6 indicadores (PQRD, cancelación CX, oportunidad entrega, eventos adversos, envío cotizaciones, entrega productos); `General` banda MOS (`FD:FL`) usa 3. ¿Cuál mapea a cuál?
5. **Nota técnica**: ¿de dónde sale el `% ejecución` si no está el "Tablero de Seguimiento Ejecución Cápitas"? ¿Ese archivo lo puedes compartir?
6. **INFORME CÁPITA Y MOVILIDAD**: ¿se usa o quedó reemplazado por el Tablero de Cápitas / el `9.Consolidado`?
7. **Google Sheets** (PGP, PQRSD): ¿nos compartes acceso, o exportas `.xlsx` a una carpeta acordada cada mes?
8. **1552**: ¿tienes hoy alguna tabla propia `NIT → código de habilitación → sede`? ¿Cuántas sedes maneja un contrato grande típico?

### Borrador de correo

> **Asunto:** Automatización matriz de seguimiento — datos para avanzar
>
> Jonatan, ya revisamos toda la data de julio contra `Indicaciones.txt`. Para seguir necesito:
>
> 1. **Permiso a `Z:\10.INDICADORES SEGUIMIENTO CONTRACTUAL`** (lo autoriza Yuli) y acceso a
>    las **2 Google Sheets** (Matriz Seguimiento PGP y Resultados Contractuales PQRSD).
> 2. El archivo **"Tablero de Seguimiento Ejecución Cápitas ENE-JUL 2026"** (no vino en el paquete).
> 3. La **lista cerrada de valores de "CATEGORÍA DEL CONTRATO"** y, si puedes, qué indicadores
>    le aplican a cada una. Es la llave con la que se segmenta todo.
>
> Y 3 dudas puntuales: (a) en Curativa, ¿aplican las hojas `0227 2024` y `0429 2025` o solo una?
> (b) en financieros "SUBSIDIADO Y CONTRIBUTIVO", ¿cómo se pondera RS/RC? (c) en MOS, ¿cuál de
> los 6 indicadores de la fuente mapea a cada uno de los 3 de la matriz?
