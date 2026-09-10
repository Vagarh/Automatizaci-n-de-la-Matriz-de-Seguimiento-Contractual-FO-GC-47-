# Errores detectados en el instructivo FO-GC-47

Encontrados leyendo `DOCUMENTOS/FO-GC-47 - Informe Mensual De Seguimiento A Contratos De Salud.doc`.
Son errores de copy-paste. Si se codifican tal cual, el proceso rompe o diligencia mal.
Las correcciones ya están aplicadas en `config/fuentes.yaml` (marcadas con `# CORRECCIÓN`).

| # | Dónde | Dice | Debería decir | Impacto |
|---|---|---|---|---|
| 1 | Actividad 11 · Hogares de Paso | filtrar `CATEGORÍA DEL CONTRATO` = `"TRANSPORTE ASISTENCIAL"` | `"Hogares de Paso"` | Diligenciaría hogares de paso sobre los contratos de transporte |
| 2 | Actividad 18 · Planificación Familiar | "no pertenezcan a la categoría filtrada (`MOS y ayudas ortopédicas`)" | filtro real de planificación familiar (¿por categoría? ¿sin filtro?) | Excluye las IPS equivocadas |
| 3 | Actividad 18 · Ayudas Diagnósticas | idem — `"MOS y ayudas ortopédicas"` | filtro real de ayudas diagnósticas | idem |
| 4 | Numeración de actividades | **tres** actividades distintas numeradas `18` (Planificación Familiar, Ayudas Diagnósticas, Campañas Extramural) | 18, 19, 20… | Ambigüedad al referenciar pasos |
| 5 | Sección 3 · Definiciones | "Indicadores de Acceso" **definido dos veces** (texto idéntico) | una sola vez | Cosmético |
| 6 | Actividades 13 (MOS) y 14 (Medicamentos) | ambas traen el mismo párrafo de "validar que los números de contrato estén vigentes… informar al área" | confirmar si ambos validan vigencia o es copy-paste | Falsos positivos/negativos en el reporte de contratos vencidos |
| 7 | Actividad 5 · Nota técnica | cumplimiento "se debe registrar siempre como NA" pero la meta es "100%" | ok, es informativo — solo confirmar que RESULTADO sí se diligencia | Menor |
| 8 | Rutas `Z:` | el propio Jonatan dice en la reunión que "están desactualizadas… se han movido" | validar cada ruta al obtener acceso | Todos los loaders |
| 9 | Actividad 20 | nombra 4 hojas `Informe_B_C / B_E / B_M / B_A` | en el archivo real son `Informe_B_C`, `Informe_B_E`, `Informe_M`, `Informe_A` | Menor (nombres de hoja) |
| 10 | Actividad 20 | macros `Exportar_Baja_C/E/M/A` | en el VBA real: `Exportar_Baja_C`, `Exportar_Baja_E`, `Exportar_Informe_M` (`_Mediana_Alta`), `Exportar_Informe_A` (`_Apoyo_DX`) | Solo referencia — no las ejecuta el proceso |

## Además (no son errores, son cosas que el instructivo no aterriza)

- La columna `CATEGORÍA DEL CONTRATO` la mantiene el analista a mano y mezcla complejidad
  (`BAJA/MEDIANA COMPLEJIDAD`) con programa (`ATENCION DOMICILIARIA`, `OXIGENO`…). El instructivo
  la usa como llave de segmentación pero nunca lista sus valores. → Bloqueo B1, ver `REEVALUACION_2026-09.md` §7 y §10.
- La ponderación `FINANCIEROS (5%) / PQ-TUTELAS (20%) / TÉCNICOS (75%)` que el instructivo dice
  "no pisar" **no está como columnas en `General`** (termina en `NO` = `CUMPLIMIENTO`). Vive en
  las hojas `Informe_*`. El proceso igual las preserva porque no toca esas hojas.
