"""
Escritura de la matriz del mes  ·  fase F8.

Toma la plantilla `.xlsm` (el archivo del mes anterior) y, a partir de los insumos
del mes, **diligencia la hoja `General`** celda a celda en la ventana `A3:NK`, sin
tocar las fórmulas de `NL:NO`, las hojas `Informe_*` ni el proyecto VBA.

Salida:
    salidas/<mes>/7.SEGUIMIENTO CONTRACTUAL SAVIA PPAL_<MES>.xlsm   (matriz diligenciada)
    salidas/<mes>/reporte_llenado.xlsx                              (control: cobertura + concordancia)

Reutiliza los extractores y la normalización de `validacion.py`.

Estado: reconstruye las 10 bandas que hoy tienen extractor. El resto de las celdas
queda como en la plantilla. La cobertura y la comparación contra la matriz de
referencia van en el reporte de control.
"""
from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

import validacion as V

RAIZ = V.RAIZ
PLANTILLA_DEFECTO = RAIZ / "docs" / "Insumos" / "7.SEGUIMIENTO CONTRACTUAL SAVIA PPAL_JULIO.xlsm"
MATRIZ_OBJETIVO_DEFECTO = V.MATRIZ_OBJETIVO_DEFECTO
DIR_SALIDAS = RAIZ / "salidas"

COL_FIN_ESCRITURA = "NK"          # nunca se escribe más allá (NL:NO son fórmulas)
FILA_DATOS = 3

# --- Tripletas cuyo CUMPLIMIENTO hay que calcular (la fuente no lo trae) ------
#   (meta_col, resultado_col, cumplimiento_col, operador)
TRIPLETAS_CALCULAR = [
    ("AZ", "BA", "BB", "<="), ("BC", "BD", "BE", "<="),        # concurrencia
    ("BF", "BG", "BH", ">="), ("BI", "BJ", "BK", ">="),        # mipres
    ("AT", "AU", "AV", ">="), ("AW", "AX", "AY", ">="),        # autorizaciones
    ("DQ", "DR", "DS", ">="), ("DT", "DU", "DV", "<="),        # transporte
    ("FM", "FN", "FO", "<="), ("FP", "FQ", "FR", ">="), ("FS", "FT", "FU", ">="),  # medicamentos
    ("JB", "JC", "JD", "<="),                                  # tutelas
    ("NI", "NJ", "NK", ">="),                                  # extramuralidad
]
META_FIJA = {"AT": 0.8, "AW": 0.8, "BF": 0.7, "BI": 1.0, "JB": 0.0}


# ===========================================================================
#  Escritor seguro
# ===========================================================================
class EscritorMatriz:
    HOJAS_PRESERVAR_PREFIJO = ("Informe_", "EPOC", "Cuello_", "Prostata", "OXIGENO")

    def __init__(self, plantilla_xlsm: Path):
        self.plantilla = Path(plantilla_xlsm)
        self.wb = load_workbook(self.plantilla, keep_vba=True, data_only=False)
        self.ws = self.wb["General"]
        self._col_fin = column_index_from_string(COL_FIN_ESCRITURA)
        self._fila_por_contrato: dict[str, int] = {}
        self.escritas = 0
        self.sin_fila = 0

    # -- estructura ---------------------------------------------------------
    def mapear_filas(self, orden_contratos: list[str]) -> None:
        """Asigna contrato -> fila. Usa el orden dado (SIRECI); si la plantilla ya
        traía la columna A poblada, respeta esa posición."""
        ya = {}
        for r in range(FILA_DATOS, self.ws.max_row + 1):
            v = self.ws.cell(r, 1).value
            if v not in (None, ""):
                ya[V.norm_contrato(v)] = r
        fila = FILA_DATOS
        for c in orden_contratos:
            self._fila_por_contrato[c] = ya.get(c, fila)
            fila = max(fila, self._fila_por_contrato[c]) + 1

    def limpiar_datos(self) -> None:
        """Borra A{3}:NK{última}. Nunca NL:NO."""
        ultima = max([FILA_DATOS - 1, *self._fila_por_contrato.values()])
        for r in range(FILA_DATOS, ultima + 1):
            for c in range(1, self._col_fin + 1):
                self.ws.cell(r, c).value = None

    # -- escritura --------------------------------------------------------
    def escribir(self, numero_contrato: str, col_letra: str, valor) -> bool:
        ci = column_index_from_string(col_letra)
        if ci > self._col_fin:
            raise ValueError(f"columna protegida: {col_letra} (más allá de {COL_FIN_ESCRITURA})")
        fila = self._fila_por_contrato.get(V.norm_contrato(numero_contrato))
        if fila is None:
            self.sin_fila += 1
            return False
        if valor is None or (isinstance(valor, float) and pd.isna(valor)):
            return False
        self.ws.cell(fila, ci).value = valor
        self.escritas += 1
        return True

    def volcar(self, filas: list[dict]) -> None:
        for f in filas:
            self.escribir(f["numero_contrato"], f["col"], f["valor"])

    # -- guardado con verificación --------------------------------------
    def guardar(self, destino: Path) -> list[str]:
        destino = Path(destino)
        if destino.resolve() == self.plantilla.resolve():
            raise ValueError("el destino no puede ser la propia plantilla")
        destino.parent.mkdir(parents=True, exist_ok=True)
        self.wb.save(destino)
        return verificar_integridad(self.plantilla, destino)


def _celda_igual(a, b) -> bool:
    """Igualdad tolerante: los números se comparan con tolerancia (openpyxl re-serializa
    el último dígito de los VALORES CACHEADOS de las fórmulas). Fórmulas y texto: exacto."""
    if a == b:
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= max(1e-9, abs(b) * 1e-9)
    return False


def _hojas_preservadas(wb) -> list[str]:
    keep = []
    for h in wb.sheetnames:
        if h == "General":
            continue
        keep.append(h)
    return keep


def verificar_integridad(plantilla: Path, generado: Path) -> list[str]:
    """Compara el archivo generado contra la plantilla: hojas distintas de General
    idénticas celda a celda, VBA presente, y fórmulas NL:NO de General intactas.
    Devuelve la lista de problemas (vacía = OK)."""
    problemas: list[str] = []
    wa = load_workbook(plantilla, keep_vba=True, data_only=False)
    wb = load_workbook(generado, keep_vba=True, data_only=False)

    if (wa.vba_archive is not None) and (wb.vba_archive is None):
        problemas.append("el archivo generado PERDIÓ el proyecto VBA")

    fa, fb = set(wa.sheetnames), set(wb.sheetnames)
    if fa != fb:
        problemas.append(f"cambió el conjunto de hojas: falta {fa - fb}, sobra {fb - fa}")

    for h in _hojas_preservadas(wa):
        if h not in wb.sheetnames:
            continue
        sa, sb = wa[h], wb[h]
        if (sa.max_row, sa.max_column) != (sb.max_row, sb.max_column):
            problemas.append(f"hoja '{h}': cambió el tamaño "
                             f"{(sa.max_row, sa.max_column)} -> {(sb.max_row, sb.max_column)}")
            continue
        difs = 0
        for ra, rb in zip(sa.iter_rows(), sb.iter_rows()):
            for ca, cb in zip(ra, rb):
                if _celda_igual(ca.value, cb.value):
                    continue
                difs += 1
        if difs:
            problemas.append(f"hoja preservada '{h}': {difs} celda(s) distinta(s) "
                             f"(fórmula o texto — no ruido decimal)")

    # NL:NO de General
    ga, gb = wa["General"], wb["General"]
    for r in range(FILA_DATOS, ga.max_row + 1):
        for col in ("NL", "NM", "NN", "NO"):
            va = ga[f"{col}{r}"].value
            vb = gb[f"{col}{r}"].value
            if va != vb:
                problemas.append(f"General!{col}{r}: fórmula de rollup cambió {va!r} -> {vb!r}")
                break
    wa.close()
    wb.close()
    return problemas


# ===========================================================================
#  Cálculo de las tripletas META/RESULTADO/CUMPLIMIENTO faltantes
# ===========================================================================
def calcular_cumplimientos(celdas: dict[tuple[str, str], object]) -> list[dict]:
    """`celdas` = {(contrato, col): valor} ya extraído. Devuelve filas nuevas con
    la META fija (si falta) y el CUMPLIMIENTO calculado, para las tripletas de
    TRIPLETAS_CALCULAR."""
    contratos = {c for (c, _) in celdas}
    nuevas = []
    for meta_col, res_col, cmpl_col, op in TRIPLETAS_CALCULAR:
        for contrato in contratos:
            res = celdas.get((contrato, res_col))
            if res is None:
                continue
            meta = celdas.get((contrato, meta_col))
            if meta is None and meta_col in META_FIJA:
                meta = META_FIJA[meta_col]
                nuevas.append({"numero_contrato": contrato, "col": meta_col, "valor": meta})
            valor, cumpl, _regla = V.evaluar_cumplimiento(res, meta, op)
            nuevas.append({"numero_contrato": contrato, "col": cmpl_col, "valor": cumpl})
    return nuevas


# ===========================================================================
#  Orquestación:  insumos del mes  ->  matriz diligenciada + reporte
# ===========================================================================
def llenar(mes: str = "2026-07",
           ruta_unidad: Path | str | None = None,
           carpeta_insumos: Path | str | None = None,
           subcarpeta_insumos: str = "",
           plantilla_xlsm: Path | str = PLANTILLA_DEFECTO,
           matriz_objetivo: Path | str | None = MATRIZ_OBJETIVO_DEFECTO,
           salida: Path | str | None = None) -> dict:
    anio, m = mes.split("-")
    mes_nombre = V.MESES_ES[int(m) - 1]

    if carpeta_insumos:
        carpeta = Path(carpeta_insumos)
    elif ruta_unidad:
        carpeta = Path(ruta_unidad) / subcarpeta_insumos if subcarpeta_insumos else Path(ruta_unidad)
    else:
        carpeta = V.DIR_INSUMOS_DEFECTO
    plantilla = Path(plantilla_xlsm)
    obj = Path(matriz_objetivo) if matriz_objetivo else None
    out_dir = Path(salida) if salida else (DIR_SALIDAS / mes)
    out_dir.mkdir(parents=True, exist_ok=True)

    if not carpeta.exists():
        raise RuntimeError(f"no existe la carpeta de insumos: {carpeta}")
    if not plantilla.exists():
        raise RuntimeError(f"no existe la plantilla: {plantilla}")

    cfg = V.cargar_cfg()
    C = cfg["componentes"]

    def ruta(comp):
        return V.resolver_archivo(carpeta, C.get(comp, {}).get("archivo_glob", ""))

    # --- 1. orden de filas + índice de llaves (desde SIRECI) ---------------
    r_sireci = ruta("datos_generales")
    if r_sireci is None:
        raise RuntimeError("no se encontró el CONSOLIDADO...SIRECI (define las filas de General)")
    orden = V.orden_contratos_sireci(r_sireci, C["datos_generales"])
    print(f"[1/5] {len(orden)} contratos en SIRECI")

    # idx de llaves: si hay objetivo, se usa el suyo; si no, se arma desde SIRECI
    if obj and obj.exists():
        general_obj = V.leer_general(obj)
        idx = V.indice_llaves(general_obj)
        enc = V.encabezados_general(obj)
    else:
        general_obj, enc = None, {}
        idx = _indice_desde_sireci(r_sireci, C["datos_generales"])

    # --- 2. extracción (misma que validacion.run) -----------------------
    print("[2/5] extrayendo componentes ...")
    filas: list[dict] = list(V.ex_datos_generales(r_sireci, C["datos_generales"]))
    log = [("datos_generales", len(filas))]

    def add(nombre, fn):
        try:
            f = fn()
            filas.extend(f)
            log.append((nombre, len(f)))
            print(f"      {nombre:24s} -> {len(f)}")
        except Exception as e:  # noqa: BLE001
            log.append((nombre, f"ERROR {e}"))
            print(f"      {nombre:24s} -> ERROR {e}")

    if (r := ruta("auditoria_calidad")):
        add("auditoria_calidad", lambda: V.ex_auditoria_calidad(r, C["auditoria_calidad"], idx))
    if (r := ruta("financieros_evento")):
        add("financieros_evento", lambda: V.ex_financieros_evento(r, C["financieros_evento"], idx, mes_nombre))
    if (r := ruta("concurrencia")):
        add("concurrencia", lambda: V.ex_concurrencia(r, C["concurrencia"], idx))
    re_, rj_ = ruta("mipres_entregas"), ruta("mipres_juntas")
    if re_ or rj_:
        add("mipres", lambda: V.ex_mipres(re_, rj_, C["mipres_entregas"], C["mipres_juntas"], idx))
    if (r := ruta("autorizaciones_tramite")):
        add("autorizaciones_tramite",
            lambda: V._autoriz(r, C["autorizaciones_tramite"], idx, "% TRÁMITE", "AU", mes_nombre))
    if (r := ruta("autorizaciones_marcacion")):
        add("autorizaciones_marcacion",
            lambda: V._autoriz(r, C["autorizaciones_marcacion"], idx, "% FECHA PRESTACIÓN", "AX", mes_nombre))
    if (r := ruta("transporte_asistencial")):
        add("transporte_asistencial", lambda: V.ex_transporte(r, C["transporte_asistencial"], idx))
    if (r := ruta("tutelas")):
        add("tutelas", lambda: V.ex_tutelas(r, C["tutelas"], idx))
    if (r := ruta("extramuralidad")):
        add("extramuralidad", lambda: V.ex_extramuralidad(r, C["extramuralidad"], idx))
    if (r := ruta("medicamentos")):
        add("medicamentos", lambda: V.ex_medicamentos(r, C["medicamentos"], idx))

    # --- 3. calcular CUMPLIMIENTO faltante -----------------------------
    celdas = {(f["numero_contrato"], f["col"]): f["valor"] for f in filas}
    calc = calcular_cumplimientos(celdas)
    filas.extend(calc)
    print(f"[3/5] {len(calc)} celdas de CUMPLIMIENTO calculadas")

    # --- 4. escribir la matriz ---------------------------------------
    print("[4/5] escribiendo la matriz ...")
    esc = EscritorMatriz(plantilla)
    esc.mapear_filas(orden)
    esc.limpiar_datos()
    esc.volcar(filas)
    nombre = f"7.SEGUIMIENTO CONTRACTUAL SAVIA PPAL_{mes_nombre.upper()}.xlsm"
    destino = out_dir / nombre
    problemas = esc.guardar(destino)
    print(f"      {esc.escritas} celdas escritas · {esc.sin_fila} sin fila · "
          f"integridad: {'OK' if not problemas else problemas}")

    # --- 5. reporte de control -------------------------------------
    print("[5/5] reporte de control ...")
    generado = V.leer_general(destino)
    cob = _cobertura(generado)
    if general_obj is not None:
        extraido_para_cmp = [{"numero_contrato": c, "col": col,
                              "valor": generado.at[c, col] if (c in generado.index and col in generado.columns) else None}
                             for (c, col) in {(f["numero_contrato"], f["col"]) for f in filas}]
        cmp = V.comparar(extraido_para_cmp, general_obj, enc)
        cmp["banda"] = cmp["col"].map(V.banda_de_col)
        conc = (cmp.groupby("banda").agg(celdas=("igual", "size"), coinciden=("igual", "sum"))
                .assign(pct=lambda d: (100 * d.coinciden / d.celdas).round(1)).reset_index())
        pct_global = round(100 * cmp["igual"].mean(), 1)
    else:
        cmp, conc, pct_global = pd.DataFrame(), pd.DataFrame(), None

    rep = out_dir / "reporte_llenado.xlsx"
    resumen = pd.DataFrame([
        ["mes", mes], ["plantilla", plantilla.name], ["matriz_generada", str(destino)],
        ["contratos", len(orden)], ["celdas_escritas", esc.escritas],
        ["integridad_vba_formulas", "OK" if not problemas else "; ".join(problemas)],
        ["concordancia_global_%", pct_global],
    ], columns=["metrica", "valor"])
    with pd.ExcelWriter(rep, engine="xlsxwriter") as xw:
        resumen.to_excel(xw, sheet_name="resumen", index=False)
        pd.DataFrame(log, columns=["componente", "celdas"]).to_excel(xw, sheet_name="log_extraccion", index=False)
        cob.to_excel(xw, sheet_name="cobertura", index=False)
        if not conc.empty:
            conc.to_excel(xw, sheet_name="concordancia_por_banda", index=False)
        if not cmp.empty:
            cmp[~cmp.igual].sort_values(["banda", "numero_contrato", "col"]).to_excel(
                xw, sheet_name="a_revisar", index=False)
    print(f"      matriz  -> {destino}")
    print(f"      reporte -> {rep}")

    return {"matriz": destino, "reporte": rep, "problemas_integridad": problemas,
            "concordancia_por_banda": conc, "concordancia_global": pct_global,
            "cobertura": cob, "resumen": resumen}


def _indice_desde_sireci(ruta, cfg) -> dict:
    df = V.leer_hoja(ruta, cfg["hoja"], cfg["fila_encabezado"])
    cols = {V.norm_texto(c): c for c in df.columns}
    g = lambda n: cols.get(V.norm_texto(n))
    nit_a, muni_a, reg, moda, cat = {}, {}, {}, {}, {}
    for _, r in df.iterrows():
        c = V.norm_contrato(r.get(g("NUMERO DE CONTRATO")))
        if not c:
            continue
        nit = V.norm_nit(r.get(g("CONTRATISTA : NUMERO DEL NIT")))
        if nit:
            nit_a.setdefault(nit, []).append(c)
        muni = V.norm_texto(r.get(g("MUNICIPIO")))
        if muni:
            muni_a.setdefault(muni, []).append(c)
        reg[c] = V.norm_texto(r.get(g("REGIMEN")))
        moda[c] = V.norm_texto(r.get(g("MODALIDAD")))
        cat[c] = V.norm_texto(r.get(g("CATEGORIA DEL CONTRATO SEGUIMIENTO A LA RED")))
    return {"nit_a_contratos": nit_a, "muni_a_contratos": muni_a, "regimen": reg,
            "modalidad": moda, "categoria": cat, "contratos": set(reg)}


def _cobertura(generado: pd.DataFrame) -> pd.DataFrame:
    out = []
    for banda, (a, b) in V.BANDAS.items():
        ca, cb = column_index_from_string(a), column_index_from_string(b)
        letras = [get_column_letter(i) for i in range(ca, cb + 1) if get_column_letter(i) in generado.columns]
        if not letras:
            continue
        sub = generado[letras]
        con_dato = int(sub.notna().sum().sum())
        out.append({"banda": banda, "columnas": len(letras), "celdas_con_valor": con_dato})
    return pd.DataFrame(out)


if __name__ == "__main__":
    res = llenar()
    print("\n=== RESUMEN ===")
    print(res["resumen"].to_string(index=False))
    if not res["concordancia_por_banda"].empty:
        print("\n=== CONCORDANCIA POR BANDA (matriz generada vs objetivo) ===")
        print(res["concordancia_por_banda"].to_string(index=False))
