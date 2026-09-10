"""Genera notebooks/01_validacion_fuentes_y_concordancia.ipynb.

El notebook se mantiene desde este script (editar celdas en JSON es incómodo).
    python notebooks/_build_notebook.py
Después ejecutarlo con Jupyter, o:
    python -c "import nbformat;from nbclient import NotebookClient;nb=nbformat.read('notebooks/01_validacion_fuentes_y_concordancia.ipynb',as_version=4);NotebookClient(nb,resources={'metadata':{'path':'notebooks/'}}).execute();nbformat.write(nb,'notebooks/01_validacion_fuentes_y_concordancia.ipynb')"
"""
from pathlib import Path

import nbformat as nbf

nb = nbf.v4.new_notebook()
c = []
def md(s): c.append(nbf.v4.new_markdown_cell(s.strip("\n")))
def code(s): c.append(nbf.v4.new_code_cell(s.strip("\n")))

md(r"""
# Validación de fuentes y concordancia contra la matriz objetivo

**Qué hace este notebook**

1. **Valida el esquema de cada fuente** de `docs/Insumos/` contra lo que se espera
   (archivo presente · hoja correcta · columnas esperadas). Sirve para **detectar a futuro
   si un área cambia el formato de su reporte** — si un mes cambian una hoja o una columna,
   el notebook lo marca como `DRIFT` / `FALTA_HOJA` antes de que rompa nada.
2. **Reconstruye** los componentes de la matriz que hoy son reproducibles a partir de los
   insumos, y los **compara celda a celda contra la matriz ya diligenciada**
   (`docs/Objetivo/…xlsm`) — la "verdad" de julio.
3. Emite el **objeto reporte** (`salidas/validacion/<fecha>/reporte_validacion.xlsx`) con:
   `resumen`, **`que_revisar`**, **`faltantes`**, `fuentes_esquema`, `discrepancias_indicaciones`,
   `cobertura_matriz`, `concordancia_por_banda`, `concordancia_clasificacion`,
   `concordancia_detalle`, `log_extraccion`.

**No escribe nada en la matriz. No toca los archivos fuente. Solo lee.**
**No se rompe** si la unidad no está conectada o falta un insumo: lo reporta y sigue.

El mapa de fuentes vive en [`config/fuentes.yaml`](../config/fuentes.yaml)
y la lógica en [`notebooks/validacion.py`](validacion.py). Para agregar un componente:
añadir su extractor en `validacion.py` y registrarlo en `run()`.
""")

code(r"""
import sys, warnings
from pathlib import Path
warnings.filterwarnings("ignore")            # openpyxl es ruidoso con celdas mal tipadas

RAIZ = Path.cwd().parents[0] if Path.cwd().name == "notebooks" else Path.cwd()
sys.path.insert(0, str(RAIZ / "notebooks"))

import pandas as pd
pd.set_option("display.max_rows", 120)
pd.set_option("display.max_colwidth", 70)
pd.set_option("display.width", 200)

import validacion as V
print("raíz del proyecto:", RAIZ)
""")

code(r'''
# ======================  PARÁMETROS DE LA CORRIDA  ============================
MES = "2026-07"                       # AAAA-MM del mes a validar

# --- Dónde están los insumos --------------------------------------------------
# Poné aquí la letra/ruta de la unidad compartida. Se buscan los archivos
# RECURSIVAMENTE (en todas las subcarpetas: FINANCIERA, MIPRES, AUTORIZACION, ...).
RUTA_UNIDAD = r"Z:\10.INDICADORES SEGUIMIENTO CONTRACTUAL"
SUBCARPETA  = ""                      # opcional: subcarpeta del mes bajo RUTA_UNIDAD

# Fallback local mientras no haya acceso al disco compartido:
if not Path(RUTA_UNIDAD).exists():
    print(f"[aviso] no se ve la unidad {RUTA_UNIDAD!r} -> uso la copia local docs/Insumos")
    RUTA_UNIDAD, SUBCARPETA = str(RAIZ / "docs" / "Insumos"), ""

# --- Matriz objetivo (la ya diligenciada, para comparar) --------------------
# Si no existe, el notebook igual valida las fuentes y avisa qué falta.
MATRIZ_OBJETIVO = RAIZ / "docs" / "Objetivo" / "7.SEGUIMIENTO CONTRACTUAL SAVIA PPAL_JULIO.xlsm"
# ==========================================================================

res = V.run(mes=MES, ruta_unidad=RUTA_UNIDAD, subcarpeta_insumos=SUBCARPETA,
            matriz_objetivo=MATRIZ_OBJETIVO)
res["resumen"]
''')

md(r"""
## 0. ¿Qué tengo que revisar?  (lo primero)

`que_revisar` = lista de acciones. `faltantes` = qué fuente esperada **no se encontró**
o **cambió de formato** en la carpeta indicada. Si esto está vacío: todas las fuentes
llegaron completas y con el formato esperado.
""")

code(r"""
display(res["que_revisar"])
print("\nFuentes no encontradas / con cambio de formato:")
display(res["faltantes"])
""")

md(r"""
## 1. Validación de esquema de las fuentes

`OK` = archivo, hoja y columnas esperadas presentes.
`DRIFT` = el archivo está pero **cambió una columna** (revisar antes de confiar en su extracción).
`FALTA_HOJA` / `FALTA_ARCHIVO` = el insumo llegó incompleto o con otro nombre.
`FALTA_INSUMO` = no se entregó este mes (Google Sheets, Tablero de Cápitas, INFORME CÁPITA Y MOVILIDAD).
`NO_APLICA` = por definición del negocio va todo `NA` (Salud Mental).
`MANUAL` = estructura heterogénea, sin automatizar aún (Ayudas Diagnósticas).
""")

code(r"""
f = res["fuentes"]
f[["componente", "estado", "archivo", "filas", "columnas_faltantes", "detalle"]]
""")

code(r"""
# Solo lo que requiere acción
f[f.estado.isin(["DRIFT", "FALTA_ARCHIVO", "FALTA_HOJA", "FALTA_INSUMO"])][
    ["componente", "estado", "archivo", "columnas_faltantes", "detalle"]
]
""")

md(r"""
## 2. Discrepancias entre `Indicaciones.txt`, `config/fuentes.yaml` y los archivos reales

Halladas en la re-evaluación (ver `docs/REEVALUACION_2026-09.md`). Ya están corregidas en
`config/fuentes.yaml`; esta tabla queda como bitácora.
""")

code(r"""
pd.read_excel(res["reporte"], sheet_name="discrepancias_indicaciones")
""")

md(r"""
## 3. Concordancia contra la matriz objetivo

Para cada componente reconstruible, se compara el valor calculado contra el de la matriz
de julio ya diligenciada. `pct` = % de celdas que coinciden (numérico con tolerancia,
texto normalizado).
""")

code(r"""
if res["concordancia_por_banda"].empty:
    print("Sin matriz objetivo -> no hay comparación. (Se validó solo el esquema de las fuentes.)")
res["concordancia_por_banda"]
""")

code(r"""
d = res["concordancia_detalle"]
if d.empty:
    print("Sin comparación disponible.")
else:
  print("Clasificación de las diferencias:")
  display(d[~d.igual].groupby("clasificacion").size().sort_values(ascending=False).rename("celdas").to_frame())

# objetivo_NA_proceso_valor -> el contrato NO estaba en alcance de ese componente
#                              (falta el mapa CATEGORÍA -> componentes que aplican)
# llave_no_cruzada...        -> falta alias de municipio / NIT / nombre de IPS
# diferencia_numerica        -> redondeo, ponderación RS/RC, o agregación distinta
# diferencia_texto           -> formato ('100,0%' vs 1.0), o regla de cumplimiento
""")

code(r"""
# Muestra de discrepancias por banda (para revisar con el analista)
if not d.empty:
  for banda in d.loc[~d.igual, "banda"].unique():
    sub = d[(d.banda == banda) & (~d.igual)].head(8)
    print(f"\n---  {banda}  ---")
    display(sub[["numero_contrato", "col", "encabezado", "valor_proceso", "valor_objetivo", "clasificacion"]])
""")

md(r"""
## 4. Cobertura: cuánto falta para reproducir el 100 % de la matriz

`celdas_objetivo_con_valor_no_NA` = tamaño real del trabajo manual por banda.
`reconstruible_ahora` = si este notebook ya produce algo para esa banda.
""")

code(r"""
res["cobertura"]
""")

code(r"""
cob = res["cobertura"]
if cob.empty:
    print("Sin matriz objetivo -> no se calcula cobertura.")
else:
    total = cob["celdas_objetivo_con_valor_no_NA"].sum()
    hecho = cob.loc[cob.reconstruible_ahora, "celdas_objetivo_con_valor_no_NA"].sum()
    print(f"Celdas de indicadores con valor en la matriz de julio : {total:,}")
    print(f"En bandas que el notebook ya reconstruye              : {hecho:,}  ({100*hecho/total:.0f}%)")
    print("Bandas pendientes: "
          + ", ".join(cob.loc[~cob.reconstruible_ahora & (cob.celdas_objetivo_con_valor_no_NA>0), 'banda']))
""")

md(r"""
## 5. El objeto reporte

Todo lo anterior queda consolidado en un único `.xlsx` — es el entregable que revisa el analista.
""")

code(r"""
print("Objeto reporte:", res["reporte"])
print("\nHojas:")
import openpyxl
wb = openpyxl.load_workbook(res["reporte"], read_only=True)
for h in wb.sheetnames:
    ws = wb[h]
    print(f"  - {h:32s} {ws.max_row-1:>6} filas")
wb.close()
""")

md(r"""
## 6. Generar la matriz del mes (escritura del `.xlsm`)

Esto es el **objetivo**: a partir de los insumos, escribir la matriz diligenciada.
Toma la **plantilla** (el `.xlsm` del mes anterior), limpia `A3:NK`, escribe celda a
celda las bandas que hoy tienen extractor, y **verifica** que no se tocaron las fórmulas
`NL:NO`, las hojas `Informe_*` ni el VBA.

Salida: `salidas/<mes>/7.SEGUIMIENTO CONTRACTUAL SAVIA PPAL_<MES>.xlsm` + `reporte_llenado.xlsx`.
""")

code(r'''
import escritor as E

PLANTILLA = RAIZ / "docs" / "Insumos" / "7.SEGUIMIENTO CONTRACTUAL SAVIA PPAL_JULIO.xlsm"
# En operación: el .xlsm del mes ANTERIOR (agosto para cerrar septiembre, etc.).

gen = E.llenar(mes=MES, ruta_unidad=RUTA_UNIDAD, subcarpeta_insumos=SUBCARPETA,
               plantilla_xlsm=PLANTILLA, matriz_objetivo=MATRIZ_OBJETIVO)

print("\\nMatriz generada :", gen["matriz"])
print("Integridad VBA/fórmulas/plantillas:", "OK" if not gen["problemas_integridad"] else gen["problemas_integridad"])
gen["resumen"]
''')

code(r"""
# Concordancia de la matriz GENERADA contra la de referencia, por banda
if not gen["concordancia_por_banda"].empty:
    display(gen["concordancia_por_banda"])
    print(f"\nGlobal: {gen['concordancia_global']} %  (sobre las bandas con extractor)")
""")

md(r"""
> La matriz generada tiene diligenciadas las **10 bandas con extractor** (~96 % de
> concordancia con el trabajo manual). El resto queda como en la plantilla. Las
> diferencias que faltan son los bloqueos conocidos (mapa de categorías, ponderación
> RS/RC, alias) — ver `reporte_llenado.xlsx` hoja `a_revisar` y `docs/REEVALUACION_2026-09.md`.
""")

md(r"""
---

### Estado y próximos pasos

| Bloque | Estado |
|---|---|
| Validación de esquema de fuentes (20+ componentes) | ✅ operativo |
| Extracción + concordancia | ✅ 10 componentes |
| **Escritura del `.xlsm`** (celda a celda, `keep_vba`, verifica fórmulas/VBA) | ✅ operativo — 10 bandas, ~96 % vs julio |
| Bandas pendientes de extractor | Medicina Domiciliaria · Salud Oral · Hogares de Paso · Oxígeno · MOS · **Resolución 1552** · Planificación Familiar · **Ayudas Diagnósticas** |
| Bloqueos de negocio | mapa `CATEGORÍA DEL CONTRATO → componentes que aplican` · resolvedor `NIT ↔ REPS ↔ sede` (1552) · ponderación RS/RC en financieros |
| Insumos faltantes | Tablero de Cápitas · INFORME CÁPITA Y MOVILIDAD · 2 Google Sheets (PGP, PQRSD) |

Detalle en **`docs/REEVALUACION_2026-09.md`**.
""")

nb["cells"] = c
nb["metadata"] = {
    "kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"},
    "language_info": {"name": "python", "version": "3.11"},
}
out = Path(__file__).with_name("01_validacion_fuentes_y_concordancia.ipynb")
nbf.write(nb, out)
print("escrito", out)
