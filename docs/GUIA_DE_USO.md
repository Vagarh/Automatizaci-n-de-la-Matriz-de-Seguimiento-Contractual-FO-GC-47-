# Guía de uso — corrida mensual

Para el Analista de Información. No hay que programar nada: se ajustan 3 valores y se
ejecuta un notebook. El resultado es **un Excel** que dice **qué revisar** y **dónde el
proceso no coincide** con el diligenciamiento manual.

> El proceso **solo lee**. No modifica la matriz ni los archivos de las áreas.

---

## Una vez: instalación

1. Instalar **Python 3.11 o superior** (marcar "Add Python to PATH" en el instalador).
2. Abrir una terminal en la carpeta del proyecto y ejecutar:

   ```
   python -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   pip install jupyterlab
   ```

3. Listo. Para las siguientes corridas basta con `.venv\Scripts\activate`.

---

## Cada mes

### 1. Conectar el disco compartido

Que la unidad (`Z:` o la que corresponda) esté visible en el explorador de archivos.

### 2. Abrir el notebook

```
.venv\Scripts\activate
jupyter lab
```

En el navegador, abrir `notebooks/01_validacion_fuentes_y_concordancia.ipynb`.

### 3. Ajustar los parámetros (segunda celda de código)

```python
MES         = "2026-08"                                      # el mes que se está cerrando, AAAA-MM
RUTA_UNIDAD = r"Z:\10.INDICADORES SEGUIMIENTO CONTRACTUAL"   # raíz del disco compartido
SUBCARPETA  = ""                                             # dejar vacío salvo que haya una carpeta por mes
```

- Los archivos se buscan **en todas las subcarpetas** de `RUTA_UNIDAD` (FINANCIERA, MIPRES,
  AUTORIZACION, …). No hay que indicar cada ruta.
- Si un mes hay **varios archivos** que sirven (p. ej. de meses distintos), toma el **más
  reciente**.
- `MATRIZ_OBJETIVO` solo se usa para comparar contra un mes ya diligenciado (validación).
  Si no aplica, dejar la ruta que viene: si no existe, el notebook igual valida las fuentes.

### 4. Ejecutar todo

Menú **Run → Run All Cells**. Tarda ~1–2 minutos.

Al final produce dos cosas:

```
salidas\validacion\<fecha>\reporte_validacion.xlsx          ← control de fuentes + concordancia
salidas\<mes>\7.SEGUIMIENTO CONTRACTUAL SAVIA PPAL_<MES>.xlsm  ← la matriz diligenciada
salidas\<mes>\reporte_llenado.xlsx                          ← control de lo que se llenó
```

### 5. Revisar

**Primero** `reporte_validacion.xlsx` (hojas `que_revisar` / `faltantes`) para saber si
llegó todo. **Después** abrir la matriz generada y revisarla con `reporte_llenado.xlsx`:

- Hoja `resumen` → que `integridad_vba_formulas` diga **OK** (si no, no usar el archivo:
  avisar a quien mantiene el proyecto).
- Hoja `a_revisar` → celdas donde el cálculo no coincidió con la referencia.
- Hoja `cobertura` → qué bandas quedaron llenas y cuáles hay que completar a mano.

Completar/ajustar lo que falte en la matriz generada y correr las macros como siempre.

> La matriz generada hoy llena **10 bandas** (Datos Generales, Auditoría, Financieros,
> Acceso, Concurrencia, MIPRES, Transporte, Medicamentos, Tutelas, Extramuralidad). El
> resto sigue siendo manual hasta que se agreguen sus extractores.

---

## Cómo leer `reporte_validacion.xlsx`

### Hoja `que_revisar` — **empezar aquí**

Lista de pendientes con prioridad. Cada fila es una acción concreta:

| prioridad | tema | qué hacer |
|---|---|---|
| ALTA | `fuente FALTA_ARCHIVO` | el insumo de esa área no llegó o tiene otro nombre → **pedirlo al área** |
| ALTA | `fuente FALTA_HOJA` | el archivo llegó pero le falta la hoja esperada → devolver al área |
| MEDIA | `fuente DRIFT` | el área **cambió una columna** de su reporte → confirmar con el área antes de usarlo |
| MEDIA | `fuente FALTA_INSUMO` | fuente que no llega por el disco (Google Sheets, Tablero de Cápitas) → gestionarla aparte |
| MEDIA | `concordancia NN%` | esa banda no coincide del todo con lo manual → mirar `concordancia_detalle` |

Si esta hoja dice **"sin observaciones"**: todas las fuentes llegaron completas y con el
formato de siempre.

### Hoja `faltantes`

El detalle de lo anterior: por cada fuente con problema, **qué archivo/hoja se esperaba**,
**qué se encontró**, y **en qué carpeta se buscó**. Útil para el correo al área.

### Hoja `fuentes_esquema`

Estado de las 20+ fuentes. Estados:

- `OK` — archivo, hoja y columnas como se esperaba.
- `DRIFT` — el archivo está, pero una columna esperada **cambió de nombre o desapareció**.
  Columna `columnas_faltantes` dice cuál. **Es la señal temprana de que un área cambió su plantilla.**
- `FALTA_ARCHIVO` / `FALTA_HOJA` — no se encontró el archivo, o le falta la hoja.
- `FALTA_INSUMO` — fuente que por definición no llega por este canal (2 Google Sheets,
  el Tablero de Cápitas, el INFORME CÁPITA Y MOVILIDAD).
- `NO_APLICA` — Salud Mental: va todo `NA` (ya no mandan indicadores).
- `MANUAL` — Ayudas Diagnósticas: 24 plantillas distintas, aún sin automatizar.

### Hoja `concordancia_por_banda`

Solo si se comparó contra una matriz de referencia. `pct` = % de celdas de esa banda que
el proceso calculó **igual** que el diligenciamiento manual.

### Hoja `concordancia_detalle`

Una fila por celda comparada: `numero_contrato`, `col` (columna de `General`), `encabezado`,
`valor_proceso`, `valor_objetivo`, y `clasificacion` de la diferencia:

| clasificación | significa |
|---|---|
| `match` | coinciden |
| `objetivo_NA_proceso_valor` | el manual dejó `NA` (ese indicador **no aplica** a ese contrato) y el proceso puso un número |
| `llave_no_cruzada_o_dato_faltante` | el proceso dejó `NA` donde el manual tiene valor (falta un alias de municipio/IPS, o el dato) |
| `diferencia_numerica` | mismo concepto, número distinto (redondeo, ponderación, forma de sumar) |
| `diferencia_texto` | formato distinto (`100,0%` vs `1.0`) o la regla de Cumple/No cumple |

Para revisar una banda concreta: filtrar `col` por el rango de esa banda (ver `config/matriz.yaml`)
o filtrar `banda`.

### Hojas de contexto

- `discrepancias_indicaciones` — dónde `Indicaciones.txt` no coincide con los archivos reales
  (typos de nombre de hoja, columnas invertidas…). Ya están contempladas; es bitácora.
- `cobertura_matriz` — cuánto trabajo manual hay por banda y si el proceso ya la reproduce.
- `log_extraccion` — traza técnica de la corrida.

---

## Preguntas frecuentes

**El notebook dice `[aviso] no se ve la unidad ... -> uso la copia local docs/Insumos`.**
El disco compartido no está conectado. Conectarlo y volver a ejecutar. (Si solo se quería
probar con los datos de ejemplo, ignorar el aviso.)

**Sale `[ERROR] NO se encontró la carpeta de insumos`.**
La ruta de `RUTA_UNIDAD` está mal o la unidad no está montada. Corregir y reejecutar. El
notebook no se cae: escribe un reporte con una sola hoja `ERROR`.

**Una fuente sale `DRIFT` pero el área dice que no cambió nada.**
Revisar la columna `columnas_faltantes` del reporte: a veces es un cambio menor (una tilde,
un espacio doble, "2024" → "2025" en el nombre de una hoja). Si el cambio es real y
correcto, se ajusta `config/fuentes.yaml` (lo hace quien mantiene el proyecto).

**¿Esto ya llena la matriz?**
Todavía no. Hoy el proceso **valida las fuentes** y **compara** lo que puede reconstruir
contra el trabajo manual, para dirigir la revisión. El llenado del `.xlsm` es una fase
siguiente (ver `docs/REEVALUACION_2026-09.md` §8).

**¿Puedo correrlo sin Jupyter?**
Sí: `python notebooks/validacion.py` (control de fuentes) y `python notebooks/escritor.py`
(escribe la matriz) corren con los valores por defecto.
Para cambiar el mes o la ruta, editar la última sección de ese archivo o llamar a
`validacion.run(...)` desde Python.
