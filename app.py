"""Generador de fichas de identificación de individuos forestales (podas / talas).

Uso independiente:  streamlit run app.py
Uso desde el hub:   from app import render; render()
"""
from __future__ import annotations

import zipfile
from datetime import date
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from core.especies import cargar_catalogo, resolver_especies
from core.ficha import generar_documento
from core.fotos import clave_id, desde_carpeta, desde_zip, indexar, preparar
from core.inventario import leer_inventario, norm, orden_natural

RAIZ = Path(__file__).resolve().parent
PLANTILLA = RAIZ / "assets" / "plantilla_ficha.docx"

# Profesionales que firman las fichas: nombre -> matrícula profesional
PROFESIONALES = {
    "MIGUEL ANDRÉS BELTRÁN": "011226-0804858 ANT",
    "MATEO VARGAS MEDINA": "011226-0806919 ANT",
}

MEDIDAS = ["Tala", "Poda"]
MODOS = ["Tala (cercas vivas)", "Poda", "Mixto: según el motivo"]
TIPOS_PODA = ["Poda de ramas", "Poda de aclareo", "Poda de levante", "Poda de despunte",
              "Poda sanitaria", "Poda de formación"]
TEXTO_OBS = {
    "Tala": "Árbol de la especie {nombre} ({especie}), que {justificacion}. Se recomienda su tala.",
    "Poda": "Árbol de la especie {nombre} ({especie}), que {justificacion}. Se recomienda {tipo_poda} "
            "sobre el {pct} % de la copa, conservando el individuo.",
}


def regla_por_defecto(motivo: str) -> tuple[str, str]:
    m = norm(motivo)
    if "cerramiento" in m or "cerca" in m:
        return "Tala", "hace parte de la cerca viva (cerramiento) del predio, dentro del área de implantación del proyecto"
    if "sombra" in m:
        return "Poda", "proyecta sombra sobre el área de los paneles solares"
    if "obstacul" in m:
        return "Tala", "obstaculiza el arreglo de los paneles solares"
    return "Poda", (motivo or "interfiere con las obras del proyecto").lower()


def txt(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def num(v, dec=2, coma=True) -> str:
    try:
        s = f"{float(str(v).replace(',', '.')):.{dec}f}"
    except (TypeError, ValueError):
        return "—"
    if s == "nan":
        return "—"
    return s.replace(".", ",") if coma else s


def coordenadas(x, y, formato: str, coma: bool) -> str:
    if x is None or y is None or pd.isna(x) or pd.isna(y):
        return "—"
    if formato == "Lat / Long":
        return f"Lat {num(y, 6, coma)} / Long {num(x, 6, coma)}"
    ns = "N" if y >= 0 else "S"
    eo = "E" if x >= 0 else "W"
    return f"{num(abs(y), 6, coma)}° {ns} – {num(abs(x), 6, coma)}° {eo}"


@st.cache_data(show_spinner=False)
def _leer(datos: bytes):
    return leer_inventario(datos)


def render():
    st.title("🌳 Fichas de individuos forestales")
    st.caption("Inventario + fotos → una ficha por árbol en la plantilla de Unergy (podas y talas de cercas vivas).")

    # ---------------- Barra lateral ----------------
    with st.sidebar:
        st.header("Datos del registro")
        opciones = list(PROFESIONALES) + ["Otro"]
        quien = st.selectbox("Profesional responsable", opciones)
        if quien == "Otro":
            elaboro = st.text_input("Nombre")
            matricula = st.text_input("Matrícula profesional")
        else:
            elaboro = quien
            matricula = PROFESIONALES[quien]
            st.caption(f"Matrícula: {matricula}")
        st.header("Formato")
        coma = st.toggle("Decimales con coma", True)
        fmt_coord = st.radio("Coordenadas", ["Lat / Long", "Grados N/W"], horizontal=True)
        foto_alto = st.slider("Alto máximo de fotos (cm)", 3.5, 6.5, 4.5, 0.1,
                              help="Si alguna ficha se pasa a una segunda página, bájalo.")
        orden = st.radio("Orden de las fichas", ["Por ID (C1, C2, …)", "Como viene en el inventario"])
        numeracion = st.radio("Numeración de fichas", ["Independiente por tala / poda", "Continua"])
        individuales = st.checkbox("Incluir también un .docx por árbol", False)
        st.header("Plantilla")
        plantilla_up = st.file_uploader("Otra plantilla (opcional)", type="docx")
        with st.expander("Textos de observación"):
            obs_tala = st.text_area("Tala", TEXTO_OBS["Tala"])
            obs_poda = st.text_area("Poda", TEXTO_OBS["Poda"])
            st.caption("Variables: {nombre} {especie} {familia} {justificacion} {pct} {tipo_poda} {dap} {ht} {id}")

    # ---------------- 1. Inventario ----------------
    st.subheader("1. Inventario")
    inv = st.file_uploader("Excel del inventario forestal", type=["xlsx", "xlsm"])
    if not inv:
        st.info("Sube el inventario para empezar. Debe tener una fila de encabezados que empiece por **ID**.")
        return
    try:
        df, meta = _leer(inv.getvalue())
    except Exception as e:  # noqa: BLE001
        st.error(f"No pude leer el inventario: {e}")
        return

    if "catalogo" not in st.session_state:
        st.session_state.catalogo = cargar_catalogo()
    df = resolver_especies(df, st.session_state.catalogo)

    c1, c2 = st.columns(2)
    proyecto = c1.text_input("Proyecto", meta.get("proyecto", ""), placeholder="Ej. Minigranja Solar …")
    fecha = c2.text_input("Fecha", meta.get("fecha") or date.today().strftime("%d/%m/%Y"))
    st.caption(f"{len(df)} individuos · {df['Nombre común'].nunique()} especies · "
               f"{(df['n_fustes'] > 1).sum()} multifustales (DAP equivalente)")

    # Especies sin catálogo
    faltan = sorted(df.loc[~df["especie_en_catalogo"], "nombre_inventario"].unique())
    with st.expander(f"Catálogo de especies {'⚠️ ' + str(len(faltan)) + ' sin identificar' if faltan else '✅'}",
                     expanded=bool(faltan)):
        if faltan:
            st.write("Completa estas especies (se aplican a todas las fichas):")
            nuevas = st.data_editor(
                pd.DataFrame({"clave": faltan, "nombre_comun": faltan, "nombre_cientifico": "", "familia": ""}),
                hide_index=True, disabled=["clave"], key="nuevas_especies")
            if st.button("Agregar al catálogo"):
                nuevas["clave"] = nuevas["clave"].map(norm)
                st.session_state.catalogo = pd.concat([st.session_state.catalogo, nuevas], ignore_index=True)
                st.rerun()
        resumen = df.groupby(["nombre_inventario", "Nombre común", "Nombre científico", "Familia"]).size()
        st.dataframe(resumen.rename("n").reset_index(), hide_index=True, width="stretch")
        st.download_button("Descargar catálogo actualizado (CSV)",
                           st.session_state.catalogo.to_csv(index=False).encode("utf-8"), "especies.csv")

    # ---------------- 2. Fotos ----------------
    st.subheader("2. Fotos")
    st.caption("Nombres tipo **C90_1.jpg**, **C90_2.jpg** (también C90-1, C90 (1)…). "
               "Silueta 1 = foto _1, Silueta 2 = foto _2.")
    t_zip, t_arch, t_dir = st.tabs(["ZIP de la carpeta", "Seleccionar archivos", "Ruta local"])
    archivos: dict[str, bytes] = {}
    with t_zip:
        z = st.file_uploader("Carpeta comprimida en .zip", type="zip")
        if z:
            archivos.update(desde_zip(z.getvalue()))
    with t_arch:
        for f in st.file_uploader("Fotos", type=["jpg", "jpeg", "png"], accept_multiple_files=True) or []:
            archivos[f.name] = f.getvalue()
    with t_dir:
        ruta = st.text_input("Ruta de la carpeta (solo si corres la app en tu PC)")
        if ruta:
            try:
                archivos.update(desde_carpeta(ruta))
            except FileNotFoundError as e:
                st.error(str(e))
    fotos = indexar(archivos)

    ids = {clave_id(i): i for i in df["ID"]}
    df["n_fotos"] = [len(fotos.get(clave_id(i), [])) for i in df["ID"]]
    huerfanas = sorted(k for k in fotos if k not in ids)
    if archivos:
        a, b, c = st.columns(3)
        a.metric("Con 2+ fotos", int((df["n_fotos"] >= 2).sum()))
        b.metric("Con 1 foto", int((df["n_fotos"] == 1).sum()))
        c.metric("Sin fotos", int((df["n_fotos"] == 0).sum()))
        if huerfanas:
            st.warning(f"Fotos sin árbol en el inventario: {', '.join(huerfanas[:30])}")
        extra = [i for i in df["ID"] if len(fotos.get(clave_id(i), [])) > 2]
        if extra:
            st.caption(f"Con más de 2 fotos (se usan las dos primeras): {', '.join(extra)}")

    # ---------------- 3. Reglas por motivo ----------------
    st.subheader("3. Tipo de intervención")
    med_inv_df = df["Medida inventario"].map(lambda m: next((x for x in MEDIDAS if x.lower() in norm(m)), ""))
    detectado = set(med_inv_df) - {""}
    idx_modo = 0 if detectado == {"Tala"} else 1 if detectado == {"Poda"} else 2
    modo = st.radio("Este inventario va de…", MODOS, index=idx_modo, horizontal=True,
                    help="Si subes un Excel por tipo, escoge Tala o Poda y todos los árboles quedan con esa medida. "
                         "'Mixto' asigna según el motivo de intervención.")
    forzada = {"Tala (cercas vivas)": "Tala", "Poda": "Poda"}.get(modo)

    motivos = sorted(m for m in df["Motivo"].unique())
    reglas_base = pd.DataFrame(
        [{"Motivo": m or "(vacío)", "n": int((df["Motivo"] == m).sum()),
          "Medida": forzada or regla_por_defecto(m)[0], "Justificación": regla_por_defecto(m)[1]} for m in motivos])
    if forzada:
        st.caption("La justificación por motivo se usa en la observación de cada ficha. Puedes editarla.")
    else:
        st.caption("La medida se asigna por motivo; si la columna **Medida de manejo** del Excel trae Tala o Poda, "
                   "manda esa.")
    reglas = st.data_editor(
        reglas_base, hide_index=True, width="stretch", key=f"reglas_{modo}",
        disabled=["Motivo", "n"] + (["Medida"] if forzada else []),
        column_config={"Medida": st.column_config.SelectboxColumn(options=MEDIDAS, required=True)})
    regla = {("" if r["Motivo"] == "(vacío)" else r["Motivo"]): r for _, r in reglas.iterrows()}

    d1, d2, d3, d4, d5 = st.columns(5)
    tipo_poda = d1.selectbox("Tipo de poda", TIPOS_PODA, disabled=forzada == "Tala")
    pct_poda = d2.number_input("% de copa a podar", 5, 100, 30, 5, disabled=forzada == "Tala")
    estado_def = d3.selectbox("Estado fitosanitario", ["Bueno", "Regular", "Malo"])
    fauna_def = d4.selectbox("Fauna/epífitas", ["NO", "SÍ"])
    tipo_fauna_def = d5.text_input("Tipo fauna/epífitas", "No se registra")

    # ---------------- 4. Revisión por árbol ----------------
    st.subheader("4. Revisión árbol por árbol")
    base = df.copy()
    base["_med_inv"] = med_inv_df
    if orden.startswith("Por ID"):
        base = base.iloc[sorted(range(len(base)), key=lambda i: orden_natural(base["ID"].iloc[i]))]
    base["Medida"] = [forzada or mi or regla[m]["Medida"] for mi, m in zip(base["_med_inv"], base["Motivo"])]
    base["Tratamiento"] = ["TALA" if m == "Tala" else tipo_poda.upper() for m in base["Medida"]]
    base["Justificación"] = [regla[m]["Justificación"] for m in base["Motivo"]]
    base["Estado"] = base["Estado fitosanitario"].replace("", estado_def)
    base["% copa"] = [100 if m == "Tala" else pct_poda for m in base["Medida"]]
    base["Fauna"] = fauna_def
    base["Tipo fauna"] = tipo_fauna_def
    cols = ["ID", "Nombre común", "Nombre científico", "Familia", "Medida", "Tratamiento", "Justificación",
            "DAP (m)", "HT (m)", "DC N-S (m)", "DC E-O (m)", "Estado", "% copa", "Fauna", "Tipo fauna",
            "Observaciones inventario", "n_fotos"]
    st.caption("Puedes corregir cualquier celda antes de generar (por ejemplo, pasar un árbol puntual de poda a tala). "
               "Si cambias las opciones de arriba, la tabla se recalcula.")
    clave_tabla = hash((reglas.to_json(), modo, tipo_poda, pct_poda, estado_def, fauna_def, tipo_fauna_def, orden))
    tabla = st.data_editor(
        base[cols].reset_index(drop=True), hide_index=True, width="stretch", height=420,
        disabled=["ID", "n_fotos"], key=f"tabla_{clave_tabla}",
        column_config={
            "Medida": st.column_config.SelectboxColumn(options=MEDIDAS, required=True),
            "Estado": st.column_config.SelectboxColumn(options=["Bueno", "Regular", "Malo"]),
            "Fauna": st.column_config.SelectboxColumn(options=["NO", "SÍ"]),
            "% copa": st.column_config.NumberColumn(min_value=0, max_value=100, step=5),
            "n_fotos": st.column_config.NumberColumn("Fotos"),
        })
    n_tala = int((tabla["Medida"] == "Tala").sum())
    st.info(f"**{n_tala}** fichas de tala · **{len(tabla) - n_tala}** de poda")
    coords = df.set_index("ID")[["X", "Y", "Municipio", "Departamento"]]

    sin_dc = tabla["DC N-S (m)"].isna().sum()
    if sin_dc:
        st.caption(f"ℹ️ {sin_dc} árboles sin diámetro de copa: saldrán con “—”.")

    # ---------------- 5. Generar ----------------
    st.subheader("5. Generar fichas")
    if not proyecto:
        st.warning("Escribe el nombre del proyecto.")
    if st.button("Generar", type="primary", disabled=not proyecto):
        plantilla = plantilla_up.getvalue() if plantilla_up else PLANTILLA.read_bytes()
        textos = {"Tala": obs_tala, "Poda": obs_poda}
        salida = BytesIO()
        barra = st.progress(0.0, "Preparando…")
        contador = {"Tala": 0, "Poda": 0, "total": 0}
        grupos = {"Tala": [], "Poda": []}

        for i, fila in tabla.iterrows():
            r = {k: txt(v) for k, v in fila.items()}  # celdas borradas en la tabla llegan como None
            med = r["Medida"] or "Tala"
            contador[med] += 1
            contador["total"] += 1
            fp = fotos.get(clave_id(r["ID"]), [])
            c = coords.loc[r["ID"]]
            if isinstance(c, pd.DataFrame):
                c = c.iloc[0]
            tratamiento = r["Tratamiento"] or ("TALA" if med == "Tala" else tipo_poda.upper())
            pct = fila["% copa"]
            pct_txt = "—" if pct is None or pd.isna(pct) else str(int(pct))
            variables = dict(nombre=r["Nombre común"], especie=r["Nombre científico"], familia=r["Familia"],
                             justificacion=r["Justificación"], pct=pct_txt, id=r["ID"],
                             tipo_poda=tratamiento.lower() if med == "Poda" else "",
                             dap=num(fila["DAP (m)"], 2, coma), ht=num(fila["HT (m)"], 2, coma))
            try:
                obs = textos[med].format(**variables)
            except (KeyError, ValueError, IndexError) as e:
                st.error(f"Revisa el texto de observación de {med}: variable inválida {e}")
                return
            if r["Observaciones inventario"]:
                obs += " " + r["Observaciones inventario"]
            grupos[med].append(dict(
                titulo="TALA" if med == "Tala" else "PODAS",
                ficha_no=contador[med] if numeracion.startswith("Independiente") else contador["total"],
                id=r["ID"], proyecto=proyecto, ubicacion=f"{txt(c['Municipio'])}, {txt(c['Departamento'])}",
                coordenadas=coordenadas(c["X"], c["Y"], fmt_coord, coma),
                nombre=r["Nombre común"], especie=r["Nombre científico"], familia=r["Familia"],
                elaboro=elaboro, matricula=matricula, fecha=fecha,
                foto1=preparar(fp[0][2]) if len(fp) > 0 else None,
                foto2=preparar(fp[1][2]) if len(fp) > 1 else None,
                ht=num(fila["HT (m)"], 2, coma), dap=num(fila["DAP (m)"], 2, coma),
                dc_ns=num(fila["DC N-S (m)"], 2, coma), dc_eo=num(fila["DC E-O (m)"], 2, coma),
                estado=r["Estado"] or "—", pct_copa=pct_txt, fauna=r["Fauna"] or "—",
                tipo_fauna=r["Tipo fauna"] or "—", medida=tratamiento, observacion=obs))
            barra.progress((i + 1) / len(tabla) * 0.5, f"Fotos {r['ID']}")

        nombre_base = norm(proyecto).replace(" ", "_")[:40] or "proyecto"
        with zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED) as zf:
            for med, fichas in grupos.items():
                if not fichas:
                    continue
                barra.progress(0.6 if med == "Tala" else 0.8, f"Armando fichas de {med.lower()}…")
                zf.writestr(f"Fichas_{med.upper()}_{nombre_base}.docx",
                            generar_documento(plantilla, fichas, foto_alto_cm=foto_alto))
                if individuales:
                    for f in fichas:
                        for k in ("foto1", "foto2"):
                            if f[k] is not None:
                                f[k].seek(0)
                        zf.writestr(f"individuales/{med.upper()}_{f['id']}.docx",
                                    generar_documento(plantilla, [f], foto_alto_cm=foto_alto))
            resumen = tabla.copy()
            buf = BytesIO()
            resumen.to_excel(buf, index=False)
            zf.writestr(f"Resumen_fichas_{nombre_base}.xlsx", buf.getvalue())
        barra.progress(1.0, "Listo")

        st.success(f"{len(grupos['Tala'])} fichas de tala y {len(grupos['Poda'])} de poda.")
        st.download_button("⬇️ Descargar ZIP", salida.getvalue(), f"Fichas_{nombre_base}.zip",
                           mime="application/zip", type="primary")


if __name__ == "__main__":
    st.set_page_config(page_title="Fichas forestales", page_icon="🌳", layout="wide")
    render()
