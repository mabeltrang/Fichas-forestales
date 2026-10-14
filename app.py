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
# Profesionales que firman: nombre corto (botón) -> (nombre en la ficha, matrícula)
PROFESIONALES = {
    "Miguel Beltrán": ("MIGUEL ANDRÉS BELTRÁN", "011226-0804858 ANT"),
    "Mateo Vargas": ("MATEO VARGAS MEDINA", "011226-0806919 ANT"),
}

MEDIDAS = ["Tala", "Poda"]
# Texto que va después de "IDENTIFICACIÓN DE ÁRBOLES A INTERVENIR - "
TITULOS = {"Tala": "CERCA VIVA PARA TALA", "Poda": "ÁRBOLES PARA PODA"}

# Botón de la barra lateral -> medida
MODOS = {"🪓 Tala de cerca viva": "Tala", "✂️ Poda": "Poda"}
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
    # ---------------- Barra lateral: solo lo esencial ----------------
    with st.sidebar:
        st.markdown("### 👷 ¿Quién firma?")
        quien = st.radio("Profesional", list(PROFESIONALES), label_visibility="collapsed", key="ff_quien")
        elaboro, matricula = PROFESIONALES[quien]
        st.caption(f"Matrícula {matricula}")
        st.markdown("### 🌳 ¿Qué se va a hacer?")
        modo = st.radio("Medida", list(MODOS), label_visibility="collapsed", key="ff_modo")
        medida = MODOS[modo]

    st.title("🌳 Fichas de individuos forestales")
    st.caption("Sube el inventario y las fotos, revisa y descarga las fichas en Word.")

    # ---------------- 1. Inventario ----------------
    st.subheader("1. Inventario")
    inv = st.file_uploader("Excel del inventario forestal", type=["xlsx", "xlsm"])
    if not inv:
        st.info("Sube el Excel del inventario para empezar.")
        return
    try:
        df, meta = _leer(inv.getvalue())
    except Exception as e:  # noqa: BLE001
        st.error(f"No pude leer el inventario: {e}")
        return

    if "ff_catalogo" not in st.session_state:
        st.session_state.ff_catalogo = cargar_catalogo()
    df = resolver_especies(df, st.session_state.ff_catalogo)

    c1, c2 = st.columns([3, 1])
    proyecto = c1.text_input("Nombre del proyecto", meta.get("proyecto", ""),
                             placeholder="Ej. Minigranja Solar Firavitoba Occidente T123")
    fecha = c2.text_input("Fecha", meta.get("fecha") or date.today().strftime("%d/%m/%Y"))
    m1, m2, m3 = st.columns(3)
    m1.metric("Árboles", len(df))
    m2.metric("Especies", df["Nombre común"].nunique())
    m3.metric("Multifustales", int((df["n_fustes"] > 1).sum()), help="Se usa el DAP equivalente")

    faltan = sorted(df.loc[~df["especie_en_catalogo"], "nombre_inventario"].unique())
    if faltan:
        st.warning(f"No reconozco {len(faltan)} especie(s). Complétalas para que salgan bien en las fichas:")
        nuevas = st.data_editor(
            pd.DataFrame({"clave": faltan, "nombre_comun": faltan, "nombre_cientifico": "", "familia": ""}),
            hide_index=True, disabled=["clave"], key="ff_nuevas_especies",
            column_config={"clave": "En el inventario", "nombre_comun": "Nombre común",
                           "nombre_cientifico": "Nombre científico", "familia": "Familia"})
        if st.button("Guardar especies"):
            nuevas["clave"] = nuevas["clave"].map(norm)
            st.session_state.ff_catalogo = pd.concat([st.session_state.ff_catalogo, nuevas], ignore_index=True)
            st.rerun()
    with st.expander("Ver especies del inventario"):
        resumen = df.groupby(["Nombre común", "Nombre científico", "Familia"]).size()
        st.dataframe(resumen.rename("Árboles").reset_index(), hide_index=True, width="stretch")
        st.download_button("Descargar catálogo de especies (CSV)",
                           st.session_state.ff_catalogo.to_csv(index=False).encode("utf-8"), "especies.csv")

    # ---------------- 2. Fotos ----------------
    st.subheader("2. Fotos")
    st.caption("Nombradas con el ID del árbol: **C90_1.jpg** (Silueta 1) y **C90_2.jpg** (Silueta 2).")
    t_zip, t_arch, t_dir = st.tabs(["📦 Carpeta en ZIP", "🖼️ Seleccionar fotos", "📁 Ruta en mi PC"])
    archivos: dict[str, bytes] = {}
    with t_zip:
        z = st.file_uploader("Comprime la carpeta de fotos y súbela aquí", type="zip")
        if z:
            archivos.update(desde_zip(z.getvalue()))
    with t_arch:
        for f in st.file_uploader("Fotos", type=["jpg", "jpeg", "png"], accept_multiple_files=True) or []:
            archivos[f.name] = f.getvalue()
    with t_dir:
        ruta = st.text_input("Ruta de la carpeta", help="Solo funciona si corres la app en tu computador.")
        if ruta:
            try:
                archivos.update(desde_carpeta(ruta))
            except FileNotFoundError as e:
                st.error(str(e))
    fotos = indexar(archivos)
    ids = {clave_id(i) for i in df["ID"]}
    df["n_fotos"] = [len(fotos.get(clave_id(i), [])) for i in df["ID"]]
    if archivos:
        a, b, c = st.columns(3)
        a.metric("✅ Con 2 fotos", int((df["n_fotos"] >= 2).sum()))
        b.metric("⚠️ Con 1 foto", int((df["n_fotos"] == 1).sum()))
        c.metric("❌ Sin fotos", int((df["n_fotos"] == 0).sum()))
        huerfanas = sorted(k for k in fotos if k not in ids)
        if huerfanas:
            st.warning(f"Fotos que no corresponden a ningún ID del inventario: {', '.join(huerfanas[:30])}")
    else:
        st.caption("Sin fotos, las fichas salen con “Sin registro fotográfico”.")

    # ---------------- 3. Medida ----------------
    st.subheader(f"3. {'Tala de cerca viva' if medida == 'Tala' else 'Poda'}")
    tipo_poda, pct_poda = TIPOS_PODA[0], 30
    if medida == "Poda":
        p1, p2 = st.columns(2)
        tipo_poda = p1.selectbox("Tipo de poda", TIPOS_PODA, key="ff_tipo_poda")
        pct_poda = p2.number_input("% de copa a podar", 5, 100, 30, 5, key="ff_pct")

    with st.expander("⚙️ Opciones avanzadas"):
        a1, a2, a3 = st.columns(3)
        estado_def = a1.selectbox("Estado fitosanitario", ["Bueno", "Regular", "Malo"])
        fauna_def = a2.selectbox("Fauna/epífitas", ["NO", "SÍ"])
        tipo_fauna_def = a3.text_input("Tipo fauna/epífitas", "No se registra")
        b1, b2, b3 = st.columns(3)
        fmt_coord = b1.radio("Coordenadas", ["Lat / Long", "Grados N/W"])
        orden = b2.radio("Orden", ["Por ID (C1, C2, …)", "Como viene en el inventario"])
        coma = b3.toggle("Decimales con coma", True)
        individuales = b3.checkbox("Un .docx por árbol además del consolidado")
        foto_alto = st.slider("Alto máximo de fotos (cm)", 3.5, 6.5, 4.5, 0.1,
                              help="Si alguna ficha se pasa a una segunda página en Word, bájalo.")
        plantilla_up = st.file_uploader("Usar otra plantilla de ficha (.docx)", type="docx")
        obs_tala = st.text_area("Texto de observación · Tala", TEXTO_OBS["Tala"])
        obs_poda = st.text_area("Texto de observación · Poda", TEXTO_OBS["Poda"])
        st.caption("Variables: {nombre} {especie} {familia} {justificacion} {pct} {tipo_poda} {dap} {ht} {id}")

    # Justificación por motivo
    motivos = sorted(df["Motivo"].unique())
    reglas_base = pd.DataFrame([{"Motivo": m or "(vacío)", "Árboles": int((df["Motivo"] == m).sum()),
                                 "Justificación": regla_por_defecto(m)[1]} for m in motivos])

    # ---------------- 4. Revisión ----------------
    st.subheader("4. Revisión")
    with st.expander("✏️ Revisar o corregir árbol por árbol"):
        st.caption("Justificación que va en la observación, según el motivo del inventario:")
        reglas = st.data_editor(reglas_base, hide_index=True, width="stretch", disabled=["Motivo", "Árboles"],
                                key="ff_reglas")
        regla = {("" if r["Motivo"] == "(vacío)" else r["Motivo"]): r["Justificación"] for _, r in reglas.iterrows()}

        base = df.copy()
        if orden.startswith("Por ID"):
            base = base.iloc[sorted(range(len(base)), key=lambda i: orden_natural(base["ID"].iloc[i]))]
        base["Medida"] = medida
        base["Tratamiento"] = "TALA" if medida == "Tala" else tipo_poda.upper()
        base["Justificación"] = [regla[m] for m in base["Motivo"]]
        base["Estado"] = base["Estado fitosanitario"].replace("", estado_def)
        base["% copa"] = 100 if medida == "Tala" else pct_poda
        base["Fauna"] = fauna_def
        base["Tipo fauna"] = tipo_fauna_def
        cols = ["ID", "Nombre común", "Nombre científico", "Familia", "Medida", "Tratamiento", "Justificación",
                "DAP (m)", "HT (m)", "DC N-S (m)", "DC E-O (m)", "Estado", "% copa", "Fauna", "Tipo fauna",
                "Observaciones inventario", "n_fotos"]
        st.caption("Si un árbol puntual lleva otra medida, cámbiala aquí. "
                   "Al cambiar las opciones de arriba, la tabla se recalcula.")
        clave = hash((reglas.to_json(), medida, tipo_poda, pct_poda, estado_def, fauna_def, tipo_fauna_def, orden))
        tabla = st.data_editor(
            base[cols].reset_index(drop=True), hide_index=True, width="stretch", height=380,
            disabled=["ID", "n_fotos"], key=f"ff_tabla_{clave}",
            column_config={
                "Medida": st.column_config.SelectboxColumn(options=MEDIDAS, required=True),
                "Estado": st.column_config.SelectboxColumn(options=["Bueno", "Regular", "Malo"]),
                "Fauna": st.column_config.SelectboxColumn(options=["NO", "SÍ"]),
                "% copa": st.column_config.NumberColumn(min_value=0, max_value=100, step=5),
                "Observaciones inventario": "Observaciones",
                "n_fotos": st.column_config.NumberColumn("Fotos"),
            })

    n_tala = int((tabla["Medida"] == "Tala").sum())
    n_poda = len(tabla) - n_tala
    partes = [f"**{n}** de {t}" for n, t in ((n_tala, "tala"), (n_poda, "poda")) if n]
    st.info(f"Se generarán {' y '.join(partes)} · firma **{elaboro}**")
    sin_dc = int(tabla["DC N-S (m)"].isna().sum())
    if sin_dc:
        st.caption(f"ℹ️ {sin_dc} árboles sin diámetro de copa en el inventario: saldrán con “—”.")
    coords = df.set_index("ID")[["X", "Y", "Municipio", "Departamento"]]

    # ---------------- 5. Generar ----------------
    if not proyecto:
        st.warning("Escribe el nombre del proyecto (paso 1).")
    if st.button("📄 Generar fichas", type="primary", disabled=not proyecto):
        plantilla = plantilla_up.getvalue() if plantilla_up else PLANTILLA.read_bytes()
        textos = {"Tala": obs_tala, "Poda": obs_poda}
        salida = BytesIO()
        barra = st.progress(0.0, "Preparando fotos…")
        contador = {"Tala": 0, "Poda": 0}
        grupos = {"Tala": [], "Poda": []}

        for i, fila in tabla.iterrows():
            r = {k: txt(v) for k, v in fila.items()}  # celdas borradas en la tabla llegan como None
            med = r["Medida"] or "Tala"
            contador[med] += 1
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
                titulo=TITULOS[med],
                ficha_no=contador[med],
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
            barra.progress((i + 1) / len(tabla) * 0.5, f"Preparando fotos… {r['ID']}")

        nombres_catalogo = st.session_state.ff_catalogo["nombre_cientifico"].tolist()
        nombre_base = norm(proyecto).replace(" ", "_")[:40] or "proyecto"
        with zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED) as zf:
            for med, fichas in grupos.items():
                if not fichas:
                    continue
                barra.progress(0.6 if med == "Tala" else 0.8, f"Armando fichas de {med.lower()}…")
                zf.writestr(f"Fichas_{med.upper()}_{nombre_base}.docx",
                            generar_documento(plantilla, fichas, foto_alto_cm=foto_alto,
                                              otros_nombres=nombres_catalogo))
                if individuales:
                    for f in fichas:
                        for k in ("foto1", "foto2"):
                            if f[k] is not None:
                                f[k].seek(0)
                        zf.writestr(f"individuales/{med.upper()}_{f['id']}.docx",
                                    generar_documento(plantilla, [f], foto_alto_cm=foto_alto,
                                                              otros_nombres=nombres_catalogo))
            resumen = tabla.copy()
            buf = BytesIO()
            resumen.to_excel(buf, index=False)
            zf.writestr(f"Resumen_fichas_{nombre_base}.xlsx", buf.getvalue())
        barra.progress(1.0, "Listo")

        hechas = [f"{len(v)} de {k.lower()}" for k, v in grupos.items() if v]
        st.success(f"Listo: {' y '.join(hechas)}.")
        st.download_button("⬇️ Descargar ZIP", salida.getvalue(), f"Fichas_{nombre_base}.zip",
                           mime="application/zip", type="primary")


if __name__ == "__main__":
    try:  # dentro del hub la página ya está configurada
        st.set_page_config(page_title="Fichas forestales", page_icon="🌳", layout="wide")
    except st.errors.StreamlitAPIException:
        pass
    render()
