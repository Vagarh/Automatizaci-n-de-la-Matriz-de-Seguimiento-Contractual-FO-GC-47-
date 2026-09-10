# Automatización de la Matriz de Seguimiento Contractual (FO-GC-47)

Reemplazar los ~10–12 días de cruces manuales con `BUSCARV` que hace el Analista de
Información entre el día 10 y el 21 de cada mes, por un proceso **Python reproducible y
auditable**, **sin cambiar su flujo de trabajo**: la salida es el mismo `.xlsm` que él
abre, revisa y sobre el que corre sus macros.

> **No es un RPA de interfaz.** Es un **ETL de archivos Excel**: ~20 áreas dejan sus
> reportes en un disco compartido y en 2 Google Sheets; el proceso los cruza y diligencia
> la hoja `General` de `7.SEGUIMIENTO CONTRACTUAL SAVIA PPAL_<MES>.xlsm` (379 columnas,
> ~657 contratos). La exportación de plantillas (macros VBA `Exportar_*`) **ya funciona y no se toca.**

---

## Estado (septiembre 2026)

Ya se recibió la **data real de julio** (insumos crudos + matriz diligenciada). Con eso:

- **Mapa de fuentes verificado archivo por archivo** → [`config/fuentes.yaml`](config/fuentes.yaml)
- **Motor + notebook operativos** → validan el esquema de las 20+ fuentes y reconstruyen
  10 componentes con **~92 % de concordancia** celda a celda contra la matriz de julio.
- Re-evaluación completa, discrepancias, bloqueos y plan por fases → [`docs/REEVALUACION_2026-09.md`](docs/REEVALUACION_2026-09.md)

**Lo que falta** para reproducir el 100 % de la matriz: extractores de 8 bandas
(1552, ayudas dx, domiciliaria, salud oral, hogares, oxígeno, MOS, planificación familiar),
3 decisiones de negocio (categorías, REPS de sede, ponderación RS/RC), 4 insumos que no
llegan por este canal, y el escritor del `.xlsm` (`keep_vba`, celda a celda).

---

## Cómo se usa (corrida mensual)

**Requisitos:** Python 3.11+, `pip install -r requirements.txt`, y acceso al disco compartido.

1. Conectar la unidad compartida (`Z:` o la que sea).
2. Abrir **`notebooks/01_validacion_fuentes_y_concordancia.ipynb`** y ajustar el bloque de
   parámetros:

   ```python
   MES         = "2026-08"                                  # AAAA-MM del cierre
   RUTA_UNIDAD = r"Z:\10.INDICADORES SEGUIMIENTO CONTRACTUAL"  # raíz del disco compartido
   SUBCARPETA  = ""                                         # subcarpeta del mes, si aplica
   MATRIZ_OBJETIVO = ...                                    # .xlsm de referencia (opcional)
   ```

   Los insumos se buscan **recursivamente** en todas las subcarpetas — no hay que listar
   cada ruta. Si la unidad no está conectada, cae a `docs/Insumos/` y avisa.
3. Ejecutar el notebook. Genera **`salidas/validacion/<fecha>/reporte_validacion.xlsx`**.
4. Revisar ese Excel — guía paso a paso en **[`docs/GUIA_DE_USO.md`](docs/GUIA_DE_USO.md)**.

El notebook **solo lee**: no escribe en la matriz ni toca los archivos fuente. **No se
rompe** si falta la unidad o un insumo — lo reporta y sigue.

### Sin abrir Jupyter

```bash
python notebooks/validacion.py          # corre con los valores por defecto y escribe el reporte
```

o desde cualquier script / REPL:

```python
import sys; sys.path.insert(0, "notebooks")
import validacion as V
res = V.run(mes="2026-08", ruta_unidad=r"Z:\10.INDICADORES SEGUIMIENTO CONTRACTUAL")
print(res["reporte"])       # ruta del .xlsx
res["que_revisar"]          # DataFrame con las acciones pendientes
```

---

## El objeto reporte (`reporte_validacion.xlsx`)

| Hoja | Para qué |
|---|---|
| `resumen` | métricas de la corrida (fuentes OK / con problema, % de concordancia) |
| **`que_revisar`** | lista de acciones priorizada — **empezar por aquí** |
| **`faltantes`** | qué fuente esperada **no se encontró** o **cambió de formato**, y dónde se buscó |
| `fuentes_esquema` | estado de cada fuente: `OK` / `DRIFT` / `FALTA_ARCHIVO` / `FALTA_HOJA` / `FALTA_INSUMO` |
| `discrepancias_indicaciones` | bitácora: dónde `Indicaciones.txt` no coincide con los archivos reales |
| `cobertura_matriz` | tamaño real del trabajo por banda y si el proceso ya la reconstruye |
| `concordancia_por_banda` | % de celdas que coinciden con la matriz de referencia, por banda |
| `concordancia_clasificacion` | por qué difieren las que difieren (dato / regla / llave / redondeo) |
| `concordancia_detalle` | una fila por celda comparada (contrato, columna, valor proceso vs objetivo) |
| `log_extraccion` | qué componente se extrajo y cuántas celdas produjo |

---

## Cómo se implementa / se extiende

Todo el motor está en **`notebooks/validacion.py`** (un solo archivo, ~900 líneas, sin
dependencias del proyecto). Estructura interna:

```
norm_texto / norm_nit / norm_contrato / es_na / parse_umbral   normalización de llaves y metas
evaluar_cumplimiento                                           las 7 variantes de Cumple/No cumple/NA
validar_fuentes(cfg, carpeta)                                  esquema de cada fuente -> OK / DRIFT / FALTA_*
leer_general(xlsm)                                             matriz objetivo -> DataFrame wide por contrato
indice_llaves(general)                                         mapas NIT->contratos, municipio->contratos, ...
ex_<componente>(ruta, cfg, idx)                                un extractor por componente -> [{contrato, col, valor}]
comparar(extraido, general, enc)                               celda a celda -> DataFrame clasificado
run(mes, ruta_unidad, ...)                                     orquesta todo y escribe el reporte
```

**Para agregar un componente** (p. ej. `salud_oral`):

1. En [`config/fuentes.yaml`](config/fuentes.yaml), su entrada ya tiene `archivo_glob`,
   `hoja`, `fila_encabezado`, `columnas_esperadas`, `destino`. Ajustar si hace falta.
2. En `validacion.py`, escribir `def ex_salud_oral(ruta, cfg, idx) -> list[dict]:` que
   devuelva filas `{"numero_contrato", "col", "valor"}` (mirar `ex_transporte` como molde).
3. Registrarlo en `run()`:
   ```python
   if (r := ruta("salud_oral")):
       intentar("salud_oral", lambda: ex_salud_oral(r, C["salud_oral"], idx))
   ```
4. Correr el notebook; su banda aparece en `concordancia_por_banda`. Iterar contra el %.

Orden recomendado (ver `docs/REEVALUACION_2026-09.md` §8): primero **aplicabilidad**
(mapa `CATEGORÍA DEL CONTRATO → componentes`), **normalización/alias** y **motor de reglas**
— eso sube lo ya hecho de 92 % a ~99 % sin escribir un extractor nuevo.

### Utilidad: regenerar el catálogo de columnas de `General`

```bash
python notebooks/catalogo_indicadores.py --matriz "ruta\a\matriz.xlsm"          # imprime col/banda/rol
python notebooks/catalogo_indicadores.py --matriz ... --yaml > config/indicadores.generado.yaml
```

### El notebook se genera desde un script

`notebooks/01_...ipynb` se mantiene desde `notebooks/_build_notebook.py` (editar celdas
en JSON es incómodo). Tras cambiarlo: `python notebooks/_build_notebook.py` y volver a ejecutarlo.

---

## Estructura del repo

```
config/
  fuentes.yaml               ← ACTIVO. Mapa de fuentes verificado (lo consume el notebook)
  matriz.yaml                estructura de la hoja General: 379 col, bandas, NL:NO protegido, VBA
  indicadores.yaml           catálogo de tripletas META/RESULTADO/CUMPLIMIENTO por columna
  reglas_cumplimiento.yaml   spec de las 7 variantes de cumplimiento
  equivalencias.yaml         normalización + diccionarios de alias (se completan con los reportes)
  reps_maestra.csv.example   plantilla para el resolvedor NIT↔REPS↔sede (Resolución 1552, futuro)
notebooks/
  01_validacion_fuentes_y_concordancia.ipynb   ← punto de entrada
  validacion.py                                 el motor
  catalogo_indicadores.py                       utilidad standalone
  _build_notebook.py                            genera el .ipynb
docs/
  GUIA_DE_USO.md             ← guía para el analista (qué mirar en el reporte, mes a mes)
  REEVALUACION_2026-09.md     re-evaluación con la data real: mapa, discrepancias, bloqueos, plan
  ARQUITECTURA.md             flujo del motor + diseño del escritor .xlsm (futuro)
  RIESGOS.md                  los 5 riesgos duros y su mitigación
  VALIDACION_Y_CONTROL_CAMBIOS.md   regresión vs mes de referencia + acta de cambios
  ERRORES_INSTRUCTIVO_FO-GC-47.md   erratas de copy-paste del instructivo
  Insumos/  Objetivo/         datos de referencia de julio (no van al repo)
salidas/                      reportes generados (no van al repo)
```

---

## Anatomía de la matriz (`7.SEGUIMIENTO CONTRACTUAL SAVIA PPAL_<MES>.xlsm`)

Hoja **`General`** ("la sábana"):

| Elemento | Valor |
|---|---|
| Fila 1 | banda de grupo (`DATOS GENERALES`, `FINANCIEROS`, `CONCURRENCIA`, `MIPRES`, …, `AYUDAS DIAGNÓSTICAS`, `Campañas extramural`) |
| Fila 2 | 379 encabezados. La meta suele ir embebida en el título (`Meta-≤ 180 días`, `≥95% META…`). |
| Fila 3 → 659 | una fila por contrato (~657) |
| Llaves `A:K` | `A` N.º CONTRATO · `B` NIT · `C` REPS · `D` CONTRATISTA · `E` OBJETO · `F` RÉGIMEN · `G` TIPO DE RED · `H` MODALIDAD · `I` CATEGORÍA DEL CONTRATO · `J` SUBREGIÓN · `K` MUNICIPIO |
| Indicadores `AF:NK` | tripletas repetidas `META` / `RESULTADO` (o `INDICADOR`/`PFAFS`) / `CUMPLIMIENTO` |
| **Fórmulas vivas** | **solo** `NL:NO` → `NL=NM+NN` · `NM=COUNTIF(A#:NK#,"Cumple")` · `NN=COUNTIF(…,"No cumple")` · `NO=NM/NL` |

**Consecuencia dura (para el futuro escritor):** escribir **valores literales** solo en
`A3:NK`, **nunca** `NL:NO`; celda a celda con `openpyxl(keep_vba=True)` sobre una **copia**;
un `df.to_excel()` destruiría las fórmulas y borraría el VBA. El **orden de las columnas de
`General` es estructural** — las plantillas `Informe_*` son `VLOOKUP` por posición y las
macros rompen si una columna se desplaza. Detalle en `docs/ARQUITECTURA.md` y `docs/RIESGOS.md`.

---

## Los 5 riesgos que hay que resolver (resumen — detalle en `docs/RIESGOS.md`)

1. **Fórmulas vivas + VBA en la matriz** → escribir celda a celda `A3:NK`, `keep_vba`, hash de hojas preservadas.
2. **"Cumple/No cumple" no es una sola regla** → 7 variantes en `config/reglas_cumplimiento.yaml`, primer match gana.
3. **Cruces por texto frágiles** (municipio, nombre IPS) → `config/equivalencias.yaml` + reporte de los que no cruzan.
4. **El REPS de sede** (Resolución 1552) → resolvedor `NIT ↔ REPS ↔ sede` (REPS público + `reps_maestra.csv`).
5. **Contratos vencidos en los insumos** → no se corrige por código; se emite excepción el día 10.
