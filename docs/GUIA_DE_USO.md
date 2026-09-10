# Guía de uso — corrida mensual

Para el Analista de Información. No hay que programar nada: se ajustan unos valores en una
celda y se ejecuta el notebook. Produce **la matriz del mes diligenciada** (las bandas que
el proceso ya sabe llenar) más dos Excel de control que dicen **qué revisar** y **dónde no
coincide** con el trabajo manual.

> No modifica los archivos de las áreas ni la plantilla: escribe una **copia** en la
> carpeta de salida, y verifica que las fórmulas y las macros de esa copia quedaron intactas.

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

### 3. Ajustar los parámetros — **solo la celda `①  Parámetros`**

```python
MES          = "2026-09"                                    # el mes que se está cerrando, AAAA-MM
RUTA_FUENTES = r"Z:\10.INDICADORES SEGUIMIENTO CONTRACTUAL" # carpeta de insumos del mes
SUBCARPETA   = ""                                           # subcarpeta del mes, si aplica
PLANTILLA    = r"...\7.SEGUIMIENTO...AGOSTO.xlsm"           # la matriz del MES ANTERIOR
MATRIZ_REFERENCIA = ""                                      # .xlsm ya diligenciado, solo para comparar
DIR_SALIDA   = ""                                           # dónde dejar todo; vacío = salidas\<MES>
```

- `RUTA_FUENTES`: los archivos se buscan **en todas las subcarpetas** (FINANCIERA, MIPRES,
  AUTORIZACION, …). No hay que indicar cada ruta. Si hay varios archivos que sirven, toma
  el más reciente.
- `PLANTILLA`: la matriz del mes anterior. Se **copia** y se diligencia; el original no se toca.
- `MATRIZ_REFERENCIA`: opcional. Un `.xlsm` ya diligenciado para medir la concordancia. Vacío = no comparar.
- `DIR_SALIDA`: los 3 resultados quedan juntos en esta carpeta.
- Si una ruta no existe, el notebook usa la copia local de ejemplo y avisa (no se rompe).

### 4. Ejecutar todo

Menú **Run → Run All Cells**. Tarda ~1–2 minutos.

Al final deja 3 archivos en `DIR_SALIDA` (por defecto `salidas\<MES>\`):

```
reporte_validacion.xlsx                       ← control de fuentes + concordancia
7.SEGUIMIENTO CONTRACTUAL SAVIA PPAL_<MES>.xlsm  ← la matriz diligenciada
reporte_llenado.xlsx                          ← control de lo que se llenó
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
La ruta de `RUTA_FUENTES` está mal o la unidad no está montada. Corregir y reejecutar. El
notebook no se cae: escribe un reporte con una sola hoja `ERROR`.

**Una fuente sale `DRIFT` pero el área dice que no cambió nada.**
Revisar la columna `columnas_faltantes` del reporte: a veces es un cambio menor (una tilde,
un espacio doble, "2024" → "2025" en el nombre de una hoja). Si el cambio es real y
correcto, se ajusta `config/fuentes.yaml` (lo hace quien mantiene el proyecto).

**¿Esto ya llena la matriz?**
Sí, las **10 bandas con extractor** (Datos Generales, Auditoría, Financieros, Acceso,
Concurrencia, MIPRES, Transporte, Medicamentos, Tutelas, Extramuralidad) — con ~96 % de
concordancia contra el trabajo manual de julio. Las otras 8 bandas siguen siendo manuales
hasta que se agreguen sus extractores (ver `docs/REEVALUACION_2026-09.md` §8).

**¿Puedo correrlo sin Jupyter?**
Sí: `python notebooks/validacion.py` (control de fuentes) y `python notebooks/escritor.py`
(escribe la matriz) corren con los valores por defecto.
Para cambiar el mes o la ruta, editar la última sección de ese archivo o llamar a
`validacion.run(...)` desde Python.
