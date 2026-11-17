# Datos

## `data/raw` (originales, sin modificar)

| Archivo | Qué es |
|---|---|
| `EquiposOrios.txt` | Exportación de Orios de planta GDL (corte 15/10/2026). Texto separado por tabuladores, UTF‑8, **con acentos dañados** (`Transformaci�n`). 1,244 filas, 76 duplicadas exactas, 341 equipos. Solo la **última toma** de cada equipo. |
| `equipos.xls` | La misma exportación tal como la baja el ícono de Excel de Orios: **no es Excel**, es texto UTF‑16 con tabuladores y acentos correctos. |
| `2024_Monitor_Status.xlsx` | Reporte 2024. Hoja `equipos vibraciones`: tomas de 09/2021 a 10/2024 con componentes y síntomas. Hojas `Tabla` y `Resumen`: tablas dinámicas. |
| `EquiposOriosRed_sin_macros.xlsx` | Libro usado para procesar la exportación (originalmente `.xlsm`). Se guardó **sin macros** para que los antivirus no bloqueen el proyecto; el código de sus dos macros (refrescar la consulta, quitar duplicados y actualizar la tabla dinámica) está en `docs/macros_EquiposOriosRed.bas`. Hojas: `Ultima Toma` (exportación + columnas Año, Mes, Toma = OK/Falta, Id. Equipo = Equipo-Sección), `Todas Tomas`, `P1`, `P2`, `Rutas` (pendientes del mes), `Tiempo` (estatus por equipo hasta 03/2025), `Todos`, `x`. |
| `GDL - Pruebas Motores - Prueba_v2_sin_macros.xlsx` | Libro de **pruebas de aislamiento de motores** (megger/PdMA), guardado sin macros; su código VBA está en `docs/macros_PruebasMotores.bas`. Hoja `Indice`: un renglón por motor (Hoja, Motor, Área, Tipo, datos de placa, última prueba de Rotor/Armadura y de Estator/Campo con R, IP y Estado, Comentarios, OT, Estatus Activo/No activo, Planta y una columna sin título con el estado general Normal/Precaución/Alarma). Una hoja por motor (`1`…`190`): bloque izquierdo Estator (CA)/Armadura (CD), bloque derecho Rotor (CA)/Campo (CD), límites en filas 13–14 y pruebas desde la fila 19. |
| `GDL - Pruebas Interruptores_v2_sin_macros.xlsx` | Libro de pruebas de **interruptores** (57), sin macros; código VBA en `docs/macros_PruebasInterruptores.bas`. `Indice` + una hoja por interruptor con 4 bloques: Revisión de condición física, Operación del equipo, Pruebas eléctricas, Relevador de protección. |
| `GDL - Pruebas Transformadores_v1_sin_macros.xlsx` | Libro de pruebas de **transformadores** (34), sin macros; código VBA en `docs/macros_PruebasTransformadores.bas`. `Indice` + una hoja por transformador con bloques Fisicoquímico del aceite, Cromatografía de gases y Pruebas físicas (eléctricas); BPC y Furanos no tienen columna Estado y no cuentan para el estatus (igual que en el `Indice`). |
| `Tablero_Predictivo_GDL.pbix.zip` | Tablero de Power BI anterior. Páginas: Inicio, Vibraciones, Tribología. Medidas: Equipos P1/P2/DL, Equipos por Estatus; segmentadores por Área, Línea y Sección. |

## De dónde sale el archivo diario

- Orios exporta `EquiposOrios.txt` (o `equipos.xls` desde la pantalla Equipos con "Descargar a CSV").
- El libro `EquiposOriosRed` (que está en SharePoint/OneDrive, carpeta **Tablero predictivo**) **no se actualiza solo**: sus consultas de Power Query leen `C:\7SAP\Export\EquiposOrios.txt` de la computadora donde se abre (código en `docs/powerquery_EquiposOriosRed.m`) y se ponen al día cuando alguien lo abre ahí y ejecuta la macro Actualizar.
- El tablero puede leer cualquiera de los tres: el `.txt`, el `.xls` de Orios o el Excel `EquiposOriosRed` (hoja "Ultima Toma").

## Columnas de la exportación de Orios

`Planta, Area, Linea, Seccion, Fecha monitoreo (dd/mm/aaaa hh:mm), Equipo, Estatus, EstaActivo (Si/No), Componente/Union, Sintoma, Recomendacion, Comentarios del Analista`

Una fila por componente. El estatus es del equipo (se repite en todas sus filas).

## Estatus (nombres de Orios)

| Código | Orios | Uso en el tablero |
|---|---|---|
| P1 | Situación de emergencia | rojo |
| P2 | Fuera de límites permisibles | ámbar |
| DL | Dentro de límites | verde |

## Motores: criterio y lectura

El tablero de **Motores** lee el Excel de pruebas (el mismo que la persona abre en SharePoint/OneDrive; equivale a pararse en `Indice` con el filtro en **todas las áreas**). La función `parsePruebas` (página) y la sección MOTORES de `scripts/build_seed.py` hacen lo mismo:

- **Estatus actual** = la columna de estado general del `Indice`; si está vacía, el peor estado de las dos últimas pruebas. Alarma = P1, Precaución = P2, Normal = DL (igual que la hoja `BDatos` del libro).
- **Historial**: cada prueba de cada hoja de motor es una toma. Su estado se calcula con la resistencia de aislamiento corregida a 40 °C contra los límites de esa hoja (por omisión precaución < 100 MΩ, alarma < 75 MΩ). Se guardan también IA, IP, equipo de medición, observaciones y OT.
- Se ignoran fechas en cero (30/12/1899) y fechas a futuro (posibles errores de captura; hoy hay 2, por ejemplo una del 10/05/2027).
- **Área** = columna Área del índice; **Línea** = "Tipo" del motor. Los motores "No activo" se ocultan con el filtro *Solo activos*.
- El Excel no tiene una columna de recomendaciones; el diagnóstico es el valor fuera de límite y el comentario es el de la prueba.

## Interruptores y transformadores: criterio y lectura

`parseBloques` (página) y la sección PRUEBAS POR BLOQUES de `scripts/build_seed.py` hacen lo mismo:

- En cada hoja de equipo los bloques se detectan solos: fila 14 = "Fecha de la prueba" / "Fecha de muestra", fila 15 = Estado (o "Recomendación"), Observaciones, Acción, OT; pruebas en filas 16–32. Se usan tantos bloques como pares Fecha/Estado tenga el `Indice`.
- **Estado por prueba**: Buen Estado = DL, Monitorear = P2, cualquier otro resultado (Fuera de tolerancia / parámetros, Falla eléctrica, Regenerar, Secado vacío, Reacondicionar…) = P1. Es el mismo mapeo que la hoja `BDatos` del libro.
- **Estatus del equipo en cada fecha** = el peor de la última prueba de cada bloque hasta esa fecha (como la columna "Estatus general" del `Indice`, que exige que todos los bloques estén en Buen Estado).
- **Diagnóstico** = el estado del bloque; **recomendación** = la columna "Acción a realizar o realizada"; el comentario lleva compañía, mediciones (los puntos "OK" se resumen), observaciones y OT.
- **Prueba vencida**: interruptores sin prueba en más de 3 años, transformadores en más de 1 año (mismos plazos que `PBIAnual`).
- Diferencia con Power BI: `PBIAnual` / `BDatos` de transformadores solo consideran las pruebas de aceite, por eso ahí casi todos salen en Buen Estado; el tablero sigue el `Indice`, que incluye las pruebas eléctricas.
- Error detectado en el libro de transformadores: en el `Indice`, el renglón de la hoja 34 (Alumbrado CCM TRIT-1) lee la hoja `33`. El tablero lee la hoja 34.

## Identificador de equipo

`Equipo + "|" + Sección` (hay nombres repetidos en dos secciones: BOMBA FULLER, COMPRESOR FULLER, SOPLADOR NORTE). Equivale a `Id. Equipo` del `.xlsm` y de Power BI.

## `data/processed` (generado por `scripts/build_seed.py`)

- `seed_equipos.json`: 341 equipos, 911 tomas (exportación actual + reporte 2024 + hoja Tiempo).
- `cargas_seed.json`: los tres orígenes del historial inicial.
- `seed_motores.json` / `cargas_motores.json`: 188 motores y 1,656 tomas leídas del Excel de pruebas (corte 14/10/2026).
- `seed_int.json` / `cargas_int.json`: 57 interruptores, 141 tomas (desde 2022).
- `seed_tra.json` / `cargas_tra.json`: 34 transformadores, 370 tomas (desde 2007).
- `repair.json`: 55 palabras dañadas → correctas.
- `originales/`: la exportación y las hojas usadas, en texto, para "Ver formato original".
