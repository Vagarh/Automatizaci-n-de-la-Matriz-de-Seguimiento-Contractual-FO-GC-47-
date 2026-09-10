# Validación y control de cambios

Dos mecanismos distintos, ambos requisito del negocio:

1. **Validación por regresión** — *¿lo que hizo el proceso es igual a lo que hace el analista?*
   — **operativo** (hojas `concordancia_*` del reporte del notebook).
2. **Control de cambios del entregable** — *cuando se libera un paquete, ¿qué cambió, respecto
   de qué, y por qué?* — **diseño; se implementa junto con el escritor del `.xlsm` (fase F8).**

> Las rutas `salidas/<mes>/matriz/vNNN/` y el `control_cambios.py` que se mencionan abajo
> describen ese mecanismo #2 futuro. Hoy el notebook solo produce
> `salidas/validacion/<fecha>/reporte_validacion.xlsx`.

---

## 1. Validación por regresión

### Insumo
Un **mes de referencia completo**:
- La matriz `.xlsm` de ese mes **ya diligenciada a mano** por el analista (la "verdad").
- **Todos** los insumos crudos de ese mismo mes (los archivos que estaban en `Z:` y Drive
  cuando el analista trabajó).

Recomendado: **julio 2026** — ya disponible: `docs/Insumos/` (insumos crudos) y
`docs/Objetivo/7.SEGUIMIENTO CONTRACTUAL SAVIA PPAL_JULIO.xlsm` (matriz diligenciada).

### Ejecución
Hoy lo hace el notebook `notebooks/01_validacion_fuentes_y_concordancia.ipynb`
(`MES = "2026-07"`, `MATRIZ_OBJETIVO` apuntando al `.xlsm` de `docs/Objetivo/`), o:

```python
import sys; sys.path.insert(0, "notebooks"); import validacion as V
V.run(mes="2026-07",
      carpeta_insumos="docs/Insumos",
      matriz_objetivo="docs/Objetivo/7.SEGUIMIENTO CONTRACTUAL SAVIA PPAL_JULIO.xlsm")
```

El proceso corre sobre los insumos de julio y compara lo reconstruido contra la matriz
manual (hojas `concordancia_*` del reporte).

### Comparación
Celda a celda en la ventana `A3:NK`, alineando por `NÚMERO DE CONTRATO` (col A):

| Tipo de celda | Criterio de igualdad |
|---|---|
| `CUMPLIMIENTO` (texto) | igualdad exacta tras normalizar (`Cumple`/`No cumple`/`NA`/`Satisfactorio`/`Si`) |
| `RESULTADO` numérico | \|a−b\| ≤ tolerancia (`1e-6` relativa por defecto, configurable) |
| `META` | igualdad exacta tras normalizar (`<=8%` == `≤8%`) |
| Datos generales | igualdad exacta; fechas por valor, no por formato |

### Salida — `salidas/2026-07/reportes/regresion.xlsx`
- **Resumen**: tasa de acierto global, por banda y por tipo de celda.
- **Detalle**: una fila por discrepancia →
  `contrato · columna · banda · valor_proceso · valor_manual · Δ · clasificación`.
- **Clasificación** de cada discrepancia:
  - `dato` — el proceso leyó otro número de la fuente (revisar loader).
  - `regla_cumplimiento` — coinciden los valores pero difiere el `Cumple/No cumple` (revisar YAML).
  - `llave_no_cruzada` — el proceso dejó `NA` donde el analista puso valor (revisar equivalencias).
  - `redondeo` — diferencia solo de decimales.
  - `fuente_distinta` — el analista usó otro archivo/mes de corte.
  - `posible_error_manual` — el proceso parece correcto y el manual no (se revisa con el analista).

### Criterio de aceptación
- `CUMPLIMIENTO`: **100 %** de coincidencia.
- `RESULTADO` / datos generales: **≥ 99 %**.
- Cero discrepancias de tipo `dato` sin explicar.

Hasta cumplirlo, el proceso corre **en paralelo** al trabajo manual (1–2 meses) sin reemplazarlo.

---

## 2. Control de cambios del entregable

### Versionado
Cada corrida escribe en:
```
salidas/<mes>/matriz/v001/  7.SEGUIMIENTO CONTRACTUAL SAVIA PPAL_<MES>.xlsm
                            trazas.parquet          # (contrato, columna, valor, regla, fuente, fila)
                            manifiesto.json         # insumos usados: nombre, mtime, hash, filas
salidas/<mes>/matriz/v002/  ...
```
`manifiesto.json` deja registro de **qué versión de cada insumo** entró (nombre, hash SHA-256,
fecha de modificación, n.º de filas). Reproducible.

### Acta de cambios
Al generar `vNNN` (N>1), o al re-liberar una tanda tras una corrección,
`control_cambios.py` compara `vNNN` contra `vNNN-1` (o contra la versión marcada como
"liberada") y produce `salidas/<mes>/reportes/acta_de_cambios_vNNN.xlsx`:

| Columna | Contenido |
|---|---|
| `numero_contrato` | contrato afectado |
| `columna` / `banda` / `indicador` | celda que cambió |
| `valor_anterior` → `valor_nuevo` | el cambio |
| `regla_id` | regla de cumplimiento que produjo el valor nuevo |
| `origen_del_cambio` | archivo fuente + fila que lo originó (desde `trazas.parquet`) |
| `motivo` | clasificación automática (ver abajo) |
| `corrida` / `usuario` / `timestamp` | quién generó `vNNN` y cuándo |

**Clasificación automática del `motivo`:**
- `insumo_actualizado` — el `manifiesto.json` muestra que ese archivo fuente cambió de hash.
- `contrato_corregido` — cambió la llave / vigencia del contrato entre versiones.
- `regla_ajustada` — cambió `reglas_cumplimiento.yaml` o `indicadores.yaml` entre corridas.
- `equivalencia_agregada` — se agregó un alias que ahora hace cruzar una fila antes en `NA`.
- `sin_causa_identificada` — se marca para revisión (no debería ocurrir).

### Notificación
`control_cambios.py --notificar` genera además un resumen en texto plano
(`acta_resumen_vNNN.txt`) apto para pegar en un correo:

```
Corrida v003 · 2026-08 · Juan Felipe · 2026-08-19 14:32
Cambios respecto de v002 (liberada el 2026-08-15):
  · 12 contratos, 34 celdas.
  · 28 celdas: insumo_actualizado (FINANCIERA reemplazó "9. Consolidado ... agosto" el 2026-08-18).
  · 4 celdas: contrato_corregido (Hospital General de Medellín 0039 -> 0451).
  · 2 celdas: regla_ajustada (oxígeno: NA num/den).
  · 0 sin causa identificada.
Detalle: salidas/2026-08/reportes/acta_de_cambios_v003.xlsx
```

Así, cuando el analista entrega el paquete —o lo vuelve a entregar el día 26 tras una
corrección—, **el sistema dice exactamente qué cambió y por qué**, y esa nota viaja con el
entregable a los supervisores.

---

## 3. Reporte de cobertura y excepciones (siempre)

`salidas/<mes>/reportes/excepciones.xlsx`, generado en **toda** corrida:

| Hoja | Contenido | Acción esperada |
|---|---|---|
| `estructura_general` | columnas de `General` fuera de orden vs `matriz.yaml` | regenerar catálogo / avisar |
| `contratos_sin_cruce` | contratos de CIESI sin datos en un componente que les aplica | devolver al área (día 10) |
| `contratos_vencidos` | número de contrato reportado no vigente | devolver al área (día 10) |
| `municipios_sin_equivalencia` | textos de municipio que no cruzaron | agregar alias a `equivalencias.yaml` |
| `ips_sin_equivalencia` | nombres de IPS que no cruzaron | agregar alias |
| `nit_sin_reps` | NIT sin REPS resoluble (afecta 1552) | completar `reps_maestra.csv` |
| `metas_ausentes` | indicador `meta_origen: fuente` sin meta en el insumo | pedir al área |
| `reglas_sin_match` | celdas que cayeron en `sin_causa_identificada` | revisar YAML |
