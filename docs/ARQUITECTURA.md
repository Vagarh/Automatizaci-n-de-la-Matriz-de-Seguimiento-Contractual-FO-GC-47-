# Arquitectura del proceso

## Principio rector

**No se cambia el flujo de trabajo del analista.** La salida (cuando exista el escritor)
es el mismo `.xlsm` diligenciado que él abriría a mano. Todo lo demás —validación de
fuentes, comparación, actas de cambio— es andamiaje alrededor de ese único entregable.

---

## Hoy: el motor de validación (`notebooks/validacion.py`)

Un solo módulo, sin dependencias del proyecto. Lo maneja el notebook
`notebooks/01_validacion_fuentes_y_concordancia.ipynb`.

```
config/fuentes.yaml ──► validar_fuentes() ──► fuentes_esquema / faltantes / que_revisar
                                                 (OK · DRIFT · FALTA_ARCHIVO · FALTA_HOJA · FALTA_INSUMO)

docs/Insumos (o Z:) ──► ex_<componente>() ──► [{numero_contrato, col, valor}]
   por componente          extractor              (long, sin lógica de cumplimiento)
                              │
                              ▼
                        normalización        norm_texto / norm_nit / norm_contrato
                              │              parse_umbral / es_na
                              ▼
                    evaluar_cumplimiento()   las 7 variantes -> (valor, cumplimiento, regla_id)
                              │
docs/Objetivo/…xlsm ─► leer_general() ─► DataFrame wide por contrato (col A..NO)
                              │
                              ▼
                         comparar()          celda a celda -> clasificación de cada diferencia
                              │
                              ▼
                     salidas/validacion/<fecha>/reporte_validacion.xlsx   (10 hojas)
```

### Piezas

| Función | Rol |
|---|---|
| `norm_texto` / `norm_nit` / `norm_contrato` / `es_na` / `parse_umbral` | normalización de llaves y metas |
| `evaluar_cumplimiento(resultado, meta, …)` | intérprete de las 7 reglas de `config/reglas_cumplimiento.yaml` |
| `validar_fuentes(cfg, carpeta, mes, año)` | por cada componente: ¿está el archivo? ¿la hoja? ¿las `columnas_esperadas`? |
| `resolver_archivo(carpeta, glob)` | busca **recursivamente** (los insumos están repartidos en subcarpetas de `Z:`); el más reciente si hay varios |
| `leer_general(xlsm)` | hoja `General` → DataFrame indexado por número de contrato, columnas `A`…`NO` |
| `indice_llaves(general)` | mapas de cruce derivados de la matriz: `NIT → contratos`, `municipio → contratos`, régimen, modalidad, categoría |
| `ex_<componente>(ruta, cfg, idx)` | un extractor por componente → filas `{numero_contrato, col, valor}` |
| `comparar(extraido, general, enc)` | alinea por contrato+columna, compara con tolerancia, clasifica la diferencia |
| `run(mes, ruta_unidad, …)` | orquesta todo, arma `que_revisar` / `faltantes`, escribe el `.xlsx` |

### Contrato de un extractor

```python
def ex_<componente>(ruta: Path, cfg: dict, idx: dict) -> list[dict]:
    """
    Lee la(s) hoja(s) del insumo, resuelve el/los contrato(s) por su llave
    (NIT / municipio / número de contrato) usando `idx`, y devuelve filas:
        {"numero_contrato": <str>, "col": <letra de General>, "valor": <valor crudo>}
    Sin lógica de cumplimiento: solo leer, filtrar por periodo, tipar.
    """
```

`cfg` es la entrada del componente en `config/fuentes.yaml` (trae `hoja`,
`fila_encabezado`, `columnas_esperadas`, `destino`, filtros…).

### Reglas de cumplimiento (`config/reglas_cumplimiento.yaml`)

Se evalúan **en orden**; gana el primer `cuando` verdadero:

1. `no_aplica_componente` — el contrato no pertenece a la categoría/modalidad → `NA` / `NA`
2. `reporta_1552_satisfactorio` — reporta ≥1 indicador de la 1552 → `Si` / `Satisfactorio`
3. `meta_SD_cumple_automatico` — meta `SD` (días estancia) → `Cumple`
4. `nota_tecnica_informativa` — % ejecución nota técnica → cumplimiento `NA`
5. `oxigeno_na_num_y_den` — `NA` en numerador **y** denominador → `NA` / `NA`
6. `ips_midio_NA_sin_casos` — la IPS está en el informe pero midió `NA` → `NA` / `Cumple`
7. `ips_ausente_del_informe` — aplica pero no figura → `NA` / `NA`
8. `regla_general` — comparar `resultado` vs `umbral(meta)` según el operador

---

## El escritor del `.xlsm` (`notebooks/escritor.py`) — fase F8, operativo

Escribe la matriz del mes a partir de los insumos. Reusa los extractores de `validacion.py`.

```
SIRECI ──► ex_datos_generales() ──► orden de filas + A:AE (migración directa)
insumos ─► ex_<componente>()   ──► celdas {contrato, col, valor}  (10 bandas hoy)
            calcular_cumplimientos()   añade META fija + CUMPLIMIENTO donde la fuente no lo trae
                    │
                    ▼
   EscritorMatriz(plantilla)          load_workbook(keep_vba=True, data_only=False)
     .mapear_filas(orden)             contrato -> nº de fila (usa la col A si ya venía)
     .limpiar_datos()                 borra A3:NK ; nunca NL:NO
     .volcar(celdas)                  escribe literales; rechaza col > NK
     .guardar(destino)                guarda una COPIA (destino != plantilla) y llama a:
                    │
                    ▼
   verificar_integridad(plantilla, generado) -> list[str]
     · VBA presente en el generado
     · mismo conjunto de hojas
     · cada hoja distinta de General idéntica celda a celda
       (números con tolerancia: openpyxl re-serializa el último dígito de los
        valores cacheados de las fórmulas — eso es ruido, no un cambio)
     · fórmulas NL:NO de General intactas fila por fila
```

Salida: `salidas/<mes>/7.SEGUIMIENTO CONTRACTUAL SAVIA PPAL_<MES>.xlsm` +
`reporte_llenado.xlsx` (resumen con la verificación de integridad, cobertura,
concordancia por banda vs la matriz de referencia, y `a_revisar`).

Sobre julio: integridad **OK**, ~96 % de concordancia en las 10 bandas con extractor.
Cada extractor nuevo (F4–F7) llena su banda automáticamente.

Invariantes (de `docs/RIESGOS.md` R1):

- **Nunca** `pandas.to_excel` sobre este libro (borra fórmulas y VBA).
- `data_only=False` para conservar las fórmulas de `NL:NO`.
- Escritura **celda a celda** solo en la ventana `A3:NK`.
- El **orden de las columnas de `General`** es estructural: las hojas `Informe_*` son
  `VLOOKUP($B$7, General!$A:$AIC, <ordinal>, 0)` y las macros `Exportar_*` rompen si una
  columna se desplaza. Validar la fila 2 contra `config/matriz.yaml` al inicio y abortar
  si no coincide.
- Cada celda escrita produce una **traza** `(contrato, columna, valor, regla_id, componente,
  archivo_fuente, fila_fuente, timestamp)` → insumo del acta de cambios y del reporte de
  excepciones.

Referencia de estructura: `config/matriz.yaml` (bandas, columnas protegidas, hojas
preservadas, macros esperadas).

---

## Qué NO hace el proceso

- No ejecuta las macros VBA (lo hace el analista).
- No genera las plantillas FO-RS-99 ni los informes RIAS/PGP.
- No sube nada a `Z:` ni a Drive.
- No modifica los archivos fuente.
