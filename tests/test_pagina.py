"""Prueba de la página independiente dist/Cemex_Orios.html.

1. Abre la página: debe mostrar el inicio "Mantenimiento predictivo" con 8 técnicas.
2. Entra a Vibraciones: historial inicial (341 equipos), recorre todas las vistas e importa data/raw/equipos.xls
   (ya está en la semilla, así que no debe traer equipos nuevos).
3. Entra a Motores (188), Interruptores (57) y Transformadores (34): historial inicial y carga de su Excel de pruebas
   (ya está en la semilla, 0 nuevos). Un Excel cargado en la técnica equivocada debe dar aviso.
4. Regresa al inicio, entra a Tribología (vacía) y le carga un archivo de prueba: los datos deben quedar solo en
   Tribología y conservarse al recargar.
Uso:  python scripts/build.py && python tests/test_pagina.py
"""
import asyncio, os, pathlib
from playwright.async_api import async_playwright

ROOT = pathlib.Path(__file__).resolve().parents[1]
PAGE = (ROOT / "dist/Cemex_Orios.html").as_uri()
XLS = str(ROOT / "data/raw/equipos.xls")


async def main():
    async with async_playwright() as p:
        b = await p.chromium.launch()
        pg = await b.new_page(viewport={"width": 1300, "height": 900})
        vend = os.environ.get("VENDOR_DIR")  # opcional: carpeta con chart.umd.js y xlsx.full.min.js si no hay internet

        async def route(r):
            u = r.request.url
            if vend and "Chart.js" in u: return await r.fulfill(path=os.path.join(vend, "chart.umd.js"), content_type="text/javascript")
            if vend and "xlsx" in u: return await r.fulfill(path=os.path.join(vend, "xlsx.full.min.js"), content_type="text/javascript")
            await r.continue_()
        await pg.route("**/*", route)
        errs = []
        pg.on("pageerror", lambda e: errs.append(str(e)))
        await pg.goto(PAGE)
        await pg.wait_for_timeout(1200)
        tiles = await pg.query_selector_all("#homeBox .tile")
        print("técnicas en inicio:", len(tiles))
        assert len(tiles) == 8

        await pg.click("#homeBox [data-m=vib]"); await pg.wait_for_timeout(600)
        corte = await pg.text_content("#corte"); print("vibraciones:", corte)
        assert "341" in corte
        for tab in ["areas", "diag", "equipos", "lista", "rutas", "cargas", "resumen"]:
            await pg.click(f"#tabs [data-t={tab}]"); await pg.wait_for_timeout(200)
        await pg.click("#btnUpload")
        await pg.set_input_files("#file", XLS); await pg.wait_for_timeout(600)
        print("importar en vibraciones:", await pg.text_content("#mSave"))
        assert "(0 equipos)" in await pg.text_content("#mSave")
        # el Excel EquiposOriosRed (hoja "Ultima Toma") también se puede cargar
        await pg.click("#mClose"); await pg.click("#btnUpload")
        await pg.set_input_files("#file", str(ROOT / "data/raw/EquiposOriosRed_sin_macros.xlsx")); await pg.wait_for_timeout(1500)
        body = " ".join((await pg.text_content("#mBody")).split())
        print("Excel:", body[:160])
        assert "Ultima Toma" in body and "EquiposGDL341" in body.replace(" ", "")
        await pg.click("#mClose")

        await pg.click("#btnHome"); await pg.wait_for_timeout(300)
        await pg.click("#homeBox [data-m=mot]"); await pg.wait_for_timeout(600)
        corte = " ".join((await pg.text_content("#corte")).split()); print("motores:", corte)
        assert "188" in corte
        assert "Alarma" in await pg.text_content("#kpis") if await pg.query_selector("#kpis") else True
        for tab in ["areas", "diag", "equipos", "lista", "resumen"]:
            await pg.click(f"#tabs [data-t={tab}]"); await pg.wait_for_timeout(200)
        await pg.click("#btnUpload")
        await pg.set_input_files("#file", str(ROOT / "data/raw/GDL - Pruebas Motores - Prueba_v2_sin_macros.xlsx"))
        await pg.wait_for_selector("#mSave", timeout=60000)
        print("importar en motores:", await pg.text_content("#mSave"))
        assert "(0 equipos)" in await pg.text_content("#mSave")
        await pg.click("#mClose")

        for mod, n, xl in [("int", "57", "GDL - Pruebas Interruptores_v2_sin_macros.xlsx"), ("tra", "34", "GDL - Pruebas Transformadores_v1_sin_macros.xlsx")]:
            await pg.click("#btnHome"); await pg.wait_for_timeout(300)
            await pg.click(f"#homeBox [data-m={mod}]"); await pg.wait_for_timeout(600)
            corte = " ".join((await pg.text_content("#corte")).split()); print(mod + ":", corte)
            assert n in corte
            for tab in ["areas", "diag", "equipos", "lista", "rutas", "resumen"]:
                await pg.click(f"#tabs [data-t={tab}]"); await pg.wait_for_timeout(150)
            await pg.click("#tabs [data-t=equipos]"); await pg.wait_for_timeout(200)
            await pg.click("#eqBody tr, tbody tr"); await pg.wait_for_timeout(300)
            assert await pg.is_visible("#drawer")
            await pg.click("#dClose")
            await pg.click("#btnUpload")
            await pg.set_input_files("#file", str(ROOT / "data/raw" / xl))
            await pg.wait_for_selector("#mSave", timeout=60000)
            print("importar en", mod + ":", await pg.text_content("#mSave"))
            assert "(0 equipos)" in await pg.text_content("#mSave")
            await pg.click("#mClose")
        # Excel en la técnica equivocada: el de transformadores en interruptores
        await pg.click("#btnHome"); await pg.click("#homeBox [data-m=int]"); await pg.wait_for_timeout(300)
        await pg.click("#btnUpload")
        await pg.set_input_files("#file", str(ROOT / "data/raw/GDL - Pruebas Transformadores_v1_sin_macros.xlsx")); await pg.wait_for_timeout(4000)
        err = await pg.text_content("#mErr"); print("aviso:", err)
        assert "transformadores" in err.lower()
        await pg.click("#mClose")

        await pg.click("#btnHome"); await pg.wait_for_timeout(300)
        await pg.click("#homeBox [data-m=tri]"); await pg.wait_for_timeout(400)
        assert "Sin datos" in await pg.text_content("#corte")
        await pg.click("#btnUpload")
        await pg.set_input_files("#file", XLS); await pg.wait_for_timeout(600)
        await pg.click("#mSave"); await pg.wait_for_timeout(600)
        print("tribología:", await pg.text_content("#corte"))
        assert "341" in await pg.text_content("#corte")

        await pg.reload(); await pg.wait_for_timeout(1000)
        txt = await pg.text_content("#homeBox [data-m=tri]")
        print("inicio tras recargar:", " ".join(txt.split())[:80])
        assert "P1" in txt
        assert "5 de 8" in " ".join((await pg.text_content("#corte")).split())
        assert not errs, errs
        await b.close()
    print("OK: inicio, tablero por técnica e importación funcionan")

asyncio.run(main())
