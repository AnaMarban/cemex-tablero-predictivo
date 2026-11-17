# Mantenimiento predictivo · Planta Guadalajara

La pantalla de inicio, **Mantenimiento predictivo**, agrupa las técnicas en dos bloques:

- **Mecánico:** Vibraciones · Tribología · Pruebas no destructivas
- **Eléctrico:** Motores · Interruptores · Transformadores · Bancos de baterías · Relevadores

Cada cuadro muestra su resumen P1/P2/DL y abre su propio tablero. Todos los tableros son el mismo (el de Vibraciones, que es el proyecto base) con sus propios datos e historial. Hoy tienen datos:

- **Vibraciones**: exportación de Orios (P1 / P2 / DL).
- **Motores**: Excel "GDL - Pruebas Motores" (Alarma / Precaución / Normal, por resistencia de aislamiento).
- **Interruptores**: Excel "GDL - Pruebas Interruptores" (Fuera de parámetros / Monitorear / Buen estado; bloques Condición física, Operación, Pruebas eléctricas y Relevador).
- **Transformadores**: Excel "GDL - Pruebas Transformadores" (Fuera de tolerancia / Monitorear / Buen estado; bloques Fisicoquímico del aceite, Cromatografía y Pruebas físicas/eléctricas).

Tribología, Pruebas no destructivas, Bancos de baterías y Relevadores se llenan al cargar su reporte.

Tablero para revisar de un vistazo el estado de los equipos monitoreados por vibración en planta Guadalajara. Toma el archivo que exporta **MPC Orios** (`https://cemex.oriosmpc.mx/`), quita duplicados y guarda el historial de cada equipo. Con eso muestra lo que en Orios cuesta trabajo ver:

- cuántos equipos hay en **P1** (situación de emergencia), **P2** (fuera de límites permisibles) y **DL** (dentro de límites), por área y por sección, con enlace directo desde cada gráfica;
- **desde cuándo** está cada equipo en falla y cuántos días lleva;
- el **diagnóstico y la recomendación** de cada componente, con el comentario del analista;
- qué equipos **empeoraron** respecto a su toma anterior;
- qué equipos **no se han monitoreado** en el mes (rutas pendientes);
- una **lista de diagnóstico** con la sección, el diagnóstico y la recomendación de cada equipo;
- el **historial completo** de cada equipo (número de muestras) y cada dato en su **formato original**;
- **Exportar a Excel**: un solo archivo con una sola hoja, un renglón por equipo con estatus, sección, área, falla desde, número de muestras, diagnóstico, recomendación, comentario e historial.

Los filtros de área y sección son de **opción múltiple** (se pueden combinar varias áreas o secciones).

> El repositorio contiene datos internos de planta (equipos, fallas, recomendaciones). Mantenlo **privado**.

## Uso diario

1. Abre `dist/Cemex_Orios.html` con doble clic (Chrome o Edge). No necesita instalar nada.
2. En el inicio, elige la técnica (por ejemplo **Vibraciones**) y presiona **Actualizar**.
3. **Cargar desde carpeta:** la primera vez elige la carpeta sincronizada de SharePoint/OneDrive (**Tablero predictivo**) o `C:\7SAP\Export`. El tablero toma el archivo más reciente (`EquiposOrios.txt`, `equipos.xls` o `EquiposOriosRed.xlsm`) y avisa si tiene más de 2 días sin actualizarse.
   En **Motores**, **Interruptores** y **Transformadores** se carga su Excel `GDL - Pruebas …xlsm` de la misma carpeta (toma todas las áreas del `Indice` y el historial de la hoja de cada equipo).
4. Si ese archivo no está al día, despliega **Descargar de Orios** y sigue los pasos (Equipos → embudo → Guadalajara → Búsqueda Avanzada → Descargar a CSV); luego arrastra el archivo.

Se quitan duplicados, se agregan las tomas nuevas al historial y se guarda el original. El historial se guarda en el navegador de esa computadora. La versión `dist/orios_gdl_artifact.html` es la que se publica como Artifact de Claude con historial compartido.

## Estructura del proyecto

```
cemex-orios/
├── README.md                    este archivo
├── requirements.txt             dependencias de Python
├── src/
│   └── template.html            CÓDIGO FUENTE: el tablero completo (HTML + CSS + JavaScript)
├── scripts/
│   ├── build_seed.py            data/raw -> data/processed (historial inicial y reparación de acentos)
│   └── build.py                 src + data/processed -> dist/
├── data/
│   ├── raw/                     archivos originales (Orios, Excel, Power BI). El .xlsm va sin macros; su código está en docs/
│   └── processed/               lo que generan los scripts (semilla del historial y copias de originales)
├── assets/                      logo CEMEX y logos de cada técnica (assets/logos/<id>.webp, 240 px)
├── tests/
│   └── test_pagina.py           prueba automática de la página (carga, vistas, importar archivo)
├── dist/                        versiones compiladas, listas para usar
└── docs/                        documentación técnica
```

## Para editar el proyecto

```bash
git clone <repo> && cd cemex-orios
python -m venv .venv && source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python -m playwright install chromium                     # solo para la prueba

python scripts/build_seed.py      # solo si cambian los archivos de data/raw
python scripts/build.py           # genera dist/
python tests/test_pagina.py       # prueba la página
```

Se edita `src/template.html`, se compila con `scripts/build.py` y se prueba con `tests/test_pagina.py`. No edites `dist/` a mano: se sobrescribe en cada build.

## Documentación

- [`docs/ARQUITECTURA.md`](docs/ARQUITECTURA.md): cómo está armado el tablero, modelo de datos y cálculos ("falla desde", tendencia, rutas).
- [`docs/DATOS.md`](docs/DATOS.md): qué hay en cada archivo de `data/raw`, columnas de la exportación de Orios y significado de estatus.
- [`docs/PENDIENTES.md`](docs/PENDIENTES.md): limitaciones conocidas y siguientes pasos.

## Seguridad

- El proyecto **no guarda usuarios ni contraseñas** y no entra a Orios por su cuenta: la persona descarga el archivo y lo carga.
- No agregues credenciales al código ni a `data/`.
