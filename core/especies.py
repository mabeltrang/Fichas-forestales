"""Catálogo de especies: nombre del inventario -> nombre común, científico y familia."""
from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

from core.inventario import norm

RUTA_CATALOGO = Path(__file__).resolve().parent.parent / "data" / "especies.csv"


def cargar_catalogo(ruta: Path | None = None) -> pd.DataFrame:
    df = pd.read_csv(ruta or RUTA_CATALOGO, dtype=str).fillna("")
    df["clave"] = df["clave"].map(norm)
    return df


def _limpiar(nombre: str) -> str:
    n = norm(nombre)
    n = re.sub(r"\b(sp|spp|sp\.|spp\.)\.?$", "", n).strip()  # "Cupressus sp" -> "cupressus"
    return n.rstrip(".").strip()


def buscar(nombre: str, catalogo: pd.DataFrame) -> dict | None:
    """Busca por coincidencia exacta, luego sin 'sp', luego por género."""
    if not nombre:
        return None
    tabla = dict(zip(catalogo["clave"], catalogo.to_dict("records")))
    for clave in (norm(nombre), _limpiar(nombre)):
        if clave in tabla:
            return tabla[clave]
    genero = _limpiar(nombre).split(" ")[0]
    if genero in tabla:
        return tabla[genero]
    return None


def resolver_especies(df: pd.DataFrame, catalogo: pd.DataFrame) -> pd.DataFrame:
    """Agrega columnas Nombre común / Nombre científico / Familia."""
    comun, cient, fam, ok = [], [], [], []
    for _, r in df.iterrows():
        hit = buscar(r["cientifico_inventario"], catalogo) or buscar(r["nombre_inventario"], catalogo)
        if hit:
            comun.append(hit["nombre_comun"])
            cient.append(r["cientifico_inventario"] or hit["nombre_cientifico"])
            fam.append(r["familia_inventario"] or hit["familia"])
            ok.append(True)
        else:
            comun.append(r["nombre_inventario"])
            cient.append(r["cientifico_inventario"])
            fam.append(r["familia_inventario"])
            ok.append(False)
    out = df.copy()
    out["Nombre común"] = comun
    out["Nombre científico"] = cient
    out["Familia"] = fam
    out["especie_en_catalogo"] = ok
    return out
