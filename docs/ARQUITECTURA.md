# Arquitectura

## Una sola fuente: `src/template.html`

Todo el tablero (estilos, estructura y JavaScript) vive en `src/template.html`. `scripts/build.py` reemplaza marcadores y produce dos salidas:

| Marcador | Página independiente (`dist/Cemex_Orios.html`) | Artifact (`dist/orios_gdl_artifact.html`) |
|---|---|---|
| `__BRAND__` | Cemex Orios | Orios GDL |
| `__LOGO__` | `assets/logo_cemex.png` (base64) | ícono genérico |
| `__SEED__` | historial inicial (JSON en línea) | `null` (usa la base de datos del Artifact) |
| `__SEEDFILES__` | textos de los archivos originales | `{}` |
| `__MODLOGOS__` | logos de cada técnica (`assets/logos/<id>.webp` como data URI) | `{}` (sin marca; se usa un ícono genérico) |
| `__REPAIR__` | diccionario de acentos | igual |
| `__FILEBRAND__` | prefijo del Excel exportado | igual |

Librerías: Chart.js 4.4.1 y SheetJS 0.18.5 desde cdnjs (la página necesita internet para gráficas y exportar a Excel; los datos y tablas funcionan sin conexión).

## Actualización

El botón **Actualizar** ofrece:

1. **Cargar desde carpeta** (`loadFromFolder`): usa la API de carpetas de Chrome/Edge (`showDirectoryPicker`). La persona elige la carpeta una vez (se recuerda por técnica en IndexedDB, clave `dir_<id>`); se toma el archivo más reciente cuyo nombre coincide con `FOLDER_RX` y se avisa si tiene 2 días o más.
2. **Descargar de Orios**: guía de 4 pasos (Equipos → embudo → Guadalajara → Búsqueda Avanzada → Descargar a CSV).
3. **Arrastrar un archivo**.

`readFile` detecta si el archivo es Excel real (`.xlsx/.xlsm`, firma ZIP) y lo convierte con SheetJS (`workbookRows`, prefiere la hoja "Ultima Toma"); si no, lo trata como texto de Orios. La página no entra a Orios ni automatiza nada en él.

## Inicio y técnicas

`MODS` (en `template.html`) define las 8 técnicas: `vib` Vibraciones, `tri` Tribología, `pnd` Pruebas no destructivas (grupo Mecánico) y `mot` Motores, `int` Interruptores, `tra` Transformadores, `bat` Bancos de baterías, `rel` Relevadores (grupo Eléctrico). Para agregar o renombrar una técnica basta con editar esa lista.

- `S.mod` es la técnica abierta (`"home"` = pantalla de inicio, `renderHome`).
- `S.store[id] = {eq, cargas}` guarda los datos de cada técnica; al abrir una (`openMod`), `S.eq` y `S.cargas` apuntan a los suyos, así que todas las vistas del tablero funcionan igual para cualquier técnica.
- Una técnica sin datos muestra una pantalla de bienvenida con el botón Actualizar.
- Cada técnica puede tener en `MODS`: `rx` (qué archivo buscar en la carpeta), `parser` (`"pruebas"` = Excel de pruebas eléctricas), `labels` (nombres de estatus, p. ej. Alarma/Precaución/Normal en Motores) y `stn` (textos largos). Internamente siempre son P1/P2/DL.
- Lectores: `parseFile` (exportación de Orios, por omisión), `parsePruebas` (`parser:"pruebas"`, Motores) y `parseBloques` (`parser:"bloques"`, Interruptores y Transformadores; detecta los bloques de prueba de cada hoja solo, así que otro libro "GDL - Pruebas …" con la misma forma —por ejemplo bancos de baterías o relevadores— debería funcionar con solo agregar `parser:"bloques"`, `rx`, `archivo` y `labels` en `MODS`). Ver `docs/DATOS.md`.
- `vence` (días) cambia el aviso "sin monitoreo +90 días" por "prueba vencida" (interruptores 1095, transformadores 365).
- Semilla: `SEED.eq/cargas` es vibraciones y `SEED.mods[id]` las demás técnicas (hoy `mot`, `int`, `tra`). `scripts/build.py` la arma desde `data/processed`.

## Modos de almacenamiento

`boot()` al final del script decide el modo:

- **`SEED` presente** (página independiente): datos en `localStorage`, una clave por técnica (`orios_data` para vibraciones, `orios_data_<id>` para las demás), y archivos originales en IndexedDB (`orios_files`). Al primer arranque copia la semilla de cada técnica que tenga una.
- **Artifact con capacidad `db`**: por técnica, colecciones `equipos`/`cargas` (vibraciones) y `equipos_<id>`/`cargas_<id>` (las demás); un documento por equipo y uno por actualización. Los archivos originales se suben como *assets* y la carga guarda su `url`.
- **Sin nada de lo anterior**: modo de solo vista; se avisa con un banner.

## Modelo de datos

Documento de equipo (`S.eq`, colección `equipos`):

```js
{
  id: "e1a2b3c4",            // djb2("EQUIPO|SECCION"), igual en Python (build_seed.py) y JS (djb)
  equipo, seccion, area, linea, planta,
  activo: true,              // EstaActivo de la última toma
  estatus: "P1"|"P2"|"DL",   // de la última toma
  fecha: "2026-10-14", hora: "00:00",
  comp: [ {c, s, r, k} ],    // componentes de la última toma: componente, síntoma, recomendación, comentario
  hist: [                    // TODAS las tomas conocidas, ordenadas por fecha
    { f, h, e, a, comp:[...], src }   // src = id de la carga de donde salió ("c_seed_txt", "c1792...")
  ]
}
```

Carga (`S.cargas`, colección `cargas`): `{id, ts, archivo, fname, corte, equipos, filas, duplicados, nuevas, nuevosEq, P1, P2, DL, empeoraron:[{equipo,de,a}], asset?:{id,url}, local?:true}`.

## Procesamiento de un archivo de Orios (`parseFile` → `merge`)

1. `decode()` detecta UTF‑16 (la exportación ".xls"), UTF‑8 o Windows‑1252.
2. `fixText()` repara palabras con `�` usando `REPAIR` (y reglas para "-ción"/"-sión").
3. `parseTSV()` acepta tabulador, `;` o `,`; respeta comillas.
4. Encabezados por nombre (sin acentos): Planta, Area, Linea, Seccion, Fecha monitoreo, Equipo, Estatus, EstaActivo, Componente/Union, Sintoma, Recomendacion, Comentarios del Analista.
5. Se quitan **filas idénticas** y se descartan plantas distintas de Guadalajara.
6. Se agrupa por equipo + fecha (una "toma"); el estatus del equipo es el más grave de sus filas.
7. `merge()` agrega la toma al historial (si la fecha ya existía, la actualiza) y, si es más reciente, reemplaza los datos actuales.

## Cálculos

- **Falla desde**: recorriendo `hist` hacia atrás desde la última toma, la fecha más antigua de la racha sin DL. Si la racha llega a la primera toma registrada se muestra "al menos desde" / "≤".
- **Días en falla**: hoy − falla desde.
- **Tendencia**: compara el estatus de la última toma con la anterior (▲ empeoró, ▼ mejoró).
- **Sin monitoreo +90 días**: días desde la última toma > 90.
- **Rutas pendientes**: equipos sin ninguna toma en el mes elegido.

## Vistas

Inicio (8 técnicas con su resumen) · por técnica: Resumen (indicadores, dona, barras por área, empeoraron, modo de falla, monitoreos por mes, antigüedad, mapa sección × modo) · Ver por áreas · Diagnóstico por estatus · Equipos · Rutas pendientes · Actualizaciones. Ficha de equipo (`openDrawer`) con línea de tiempo, todas las tomas y "formato original" (`openOrig`).
