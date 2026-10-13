"""Emparejamiento de fotos con los ID del inventario.

Nombres aceptados: C90_1.jpg, C90-1.jpg, C90 1.jpg, C90 (1).jpg, c90_2.JPEG ...
Silueta 1 = foto con el número más bajo, Silueta 2 = la siguiente.
"""
from __future__ import annotations

import re
import zipfile
from collections import defaultdict
from io import BytesIO
from pathlib import Path

from PIL import Image, ImageOps

EXTENSIONES = {".jpg", ".jpeg", ".png", ".webp"}
# ID + separador obligatorio (_ - espacio o paréntesis) + número de foto
PATRON = re.compile(r"^(?P<id>.+?)(?:[\s_\-]+\(?|\()(?P<n>\d+)\)?$")


def clave_id(id_: str) -> str:
    return re.sub(r"[\s_\-]", "", str(id_)).upper()


def _parsear(nombre_archivo: str):
    stem = Path(nombre_archivo).stem.strip()
    m = PATRON.match(stem)
    if not m:  # sin número de foto (C90.jpg): se toma como foto 1 de ese ID
        return clave_id(stem), 1
    return clave_id(m.group("id")), int(m.group("n"))


def indexar(archivos: dict[str, bytes]) -> dict[str, list[tuple[int, str, bytes]]]:
    """archivos: {nombre: bytes}. Devuelve {ID: [(n, nombre, bytes), ...]} ordenado por n."""
    idx = defaultdict(list)
    for nombre, contenido in archivos.items():
        if Path(nombre).suffix.lower() not in EXTENSIONES:
            continue
        id_, n = _parsear(Path(nombre).name)
        if id_:
            idx[id_].append((n, Path(nombre).name, contenido))
    for k in idx:
        idx[k].sort(key=lambda t: t[0])
    return dict(idx)


def desde_zip(datos: bytes) -> dict[str, bytes]:
    out = {}
    with zipfile.ZipFile(BytesIO(datos)) as z:
        for info in z.infolist():
            if not info.is_dir() and not Path(info.filename).name.startswith("."):
                out[info.filename] = z.read(info)
    return out


def desde_carpeta(ruta: str) -> dict[str, bytes]:
    p = Path(ruta).expanduser()
    if not p.is_dir():
        raise FileNotFoundError(f"No existe la carpeta: {p}")
    return {f.name: f.read_bytes() for f in p.rglob("*") if f.suffix.lower() in EXTENSIONES}


def preparar(contenido: bytes, lado_max: int = 1400, calidad: int = 82) -> BytesIO:
    """Corrige la rotación EXIF del celular y reduce peso para que el Word no quede enorme."""
    img = Image.open(BytesIO(contenido))
    img = ImageOps.exif_transpose(img).convert("RGB")
    img.thumbnail((lado_max, lado_max))
    buf = BytesIO()
    img.save(buf, "JPEG", quality=calidad, optimize=True)
    buf.seek(0)
    return buf
