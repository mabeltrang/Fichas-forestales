# Fichas de individuos forestales

App Streamlit que genera las fichas de identificación de árboles a intervenir (podas y talas de cercas vivas) a partir del inventario forestal y la carpeta de fotos, usando la plantilla de Unergy (`assets/plantilla_ficha.docx`).

> Herramienta interna del equipo forestal (no guarda datos ni maneja usuarios; todo se procesa en memoria). Si pasa a ser una herramienta oficial de Unergy, debe migrarse al stack estándar (Nuxt + Supabase).

## Cómo correrla

```bash
pip install -r requirements.txt
streamlit run app.py
```

Desde el hub: `from app import render; render()`.

## Qué necesita

**Inventario (.xlsx)** con una fila de encabezados que empiece por `ID`. Columnas que usa: `Nombre común` (o `Nombre científico`), `Familia`, `CAP A/B/C (cm)` o `DAP A/B/C (m)`, `HT (m)`, `DC (N-S) (m)`, `DC (E-O) (m)`, `Coordenada X`, `Coordenada Y`, `Municipio`, `Departamento`, `Estado fitosanitario`, `Motivo de intervención`, `Observaciones`, `Medida de manejo`. El bloque superior (`PROYECTO`, `FECHA`) se lee si está diligenciado.

**Fotos** nombradas `ID_n.jpg` (C90_1, C90_2…). Se aceptan también `C90-1`, `C90 (1)`. Silueta 1 = `_1`, Silueta 2 = `_2`. Se puede subir la carpeta en .zip, seleccionar los archivos o, corriendo local, indicar la ruta.

## Qué calcula o completa

| Campo de la ficha | Origen |
|---|---|
| Nombre común, nombre técnico, familia | Catálogo `data/especies.csv` (editable desde la app) |
| DAP | Desde CAP si las fórmulas del Excel no tienen valor guardado; multifustales con DAP equivalente √(ΣDAP²) |
| Tratamiento y título de la ficha | Selector del inventario: **Tala (cercas vivas)**, **Poda** (con tipo de poda) o **Mixto** (según la columna `Medida de manejo` o una regla por motivo). Se puede cambiar árbol por árbol |
| Profesional y matrícula | Selector (Miguel, Mateo u otro) |
| Observación / justificación | Texto plantilla por medida + observaciones del inventario |
| Estado, % copa, fauna | Del inventario o valores por defecto, editables por árbol |

La foto se reduce sola cuando la observación es larga para que cada ficha quede en una página.

Salida: ZIP con `Fichas_TALA_*.docx`, `Fichas_PODA_*.docx` (una ficha por página), un resumen en Excel y, opcionalmente, un .docx por árbol.

## Estructura

```
app.py              interfaz
core/inventario.py  lectura del Excel y cálculo de DAP
core/especies.py    catálogo de especies
core/fotos.py       emparejamiento y compresión de fotos (corrige rotación del celular)
core/ficha.py       llenado de la plantilla Word
```
