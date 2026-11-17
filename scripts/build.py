"""Compila el tablero a partir de src/template.html y data/processed/.

  dist/Cemex_Orios.html            página independiente (doble clic en Chrome o Edge). Historial en el navegador.
  dist/orios_gdl_artifact.html     versión para publicar como Artifact de Claude (historial compartido; sin marca CEMEX)

Uso:  python scripts/build.py
"""
import base64, json, pathlib

ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC, DATA, DIST = ROOT / "src", ROOT / "data/processed", ROOT / "dist"

tpl = (SRC / "template.html").read_text(encoding="utf-8")
REPAIR = json.dumps(json.load(open(DATA / "repair.json", encoding="utf-8")), ensure_ascii=True)
SEED_EQ = json.load(open(DATA / "seed_equipos.json", encoding="utf-8"))
CARGAS = json.load(open(DATA / "cargas_seed.json", encoding="utf-8"))
# semillas de las demás técnicas: id del módulo -> (equipos, cargas). Los Excel originales no se incrustan (pesan mucho).
ld = lambda n: json.load(open(DATA / n, encoding="utf-8"))
MODS = {"mot": (ld("seed_motores.json"), ld("cargas_motores.json")),
        "int": (ld("seed_int.json"), ld("cargas_int.json")),
        "tra": (ld("seed_tra.json"), ld("cargas_tra.json"))}
ORIG = DATA / "originales"
SEED_FILES = {
    "c_seed_txt": dict(name="EquiposOrios_2026-10-15.txt", text=(ORIG / "EquiposOrios_2026-10-15.txt").read_text(encoding="utf-8")),
    "c_seed_2024": dict(name="2024_Monitor_Status_equipos_vibraciones.csv", text=(ORIG / "2024_Monitor_Status_equipos_vibraciones.csv").read_text(encoding="utf-8-sig")),
    "c_seed_tiempo": dict(name="EquiposOriosRed_hoja_Tiempo.csv", text=(ORIG / "EquiposOriosRed_hoja_Tiempo.csv").read_text(encoding="utf-8-sig")),
}
# logos de cada técnica (assets/logos/<id>.webp) -> data URI; solo en la versión archivo, el Artifact no lleva marca
MODLOGOS = {f.stem: "data:image/webp;base64," + base64.b64encode(f.read_bytes()).decode() for f in sorted((ROOT / "assets/logos").glob("*.webp"))}
LOGO = '<img src="data:image/png;base64,' + base64.b64encode((ROOT / "assets/logo_cemex.png").read_bytes()).decode() + '" alt="CEMEX">'
ICON = ('<svg width="34" height="34" viewBox="0 0 34 34" aria-hidden="true"><rect width="34" height="34" rx="8" fill="var(--accent)"/>'
        '<path d="M4 18h5l3-8 4 15 4-11 3 6h7" fill="none" stroke="#fff" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/></svg>')
js = lambda o: json.dumps(o, ensure_ascii=True, separators=(",", ":"))


def fill(brand, logo, filebrand, seed, seedfiles, modlogos="{}"):
    return (tpl.replace("__MODLOGOS__", modlogos).replace("__BRAND__", brand).replace("__LOGO__", logo).replace("__FILEBRAND__", filebrand)
               .replace("__REPAIR__", REPAIR).replace("__SEED__", seed).replace("__SEEDFILES__", seedfiles))


def as_document(body):
    i = body.index("</style>") + 8
    head = body[:i].replace('<meta name="claude:viewport-fit-sides" content="cover">\n', "")
    return ('<!doctype html>\n<html lang="es">\n<head>\n<meta charset="utf-8">\n'
            '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">\n'
            + head + "\n</head>\n<body>\n" + body[i:] + "\n</body>\n</html>\n")


DIST.mkdir(exist_ok=True)
seed = js({"eq": SEED_EQ, "cargas": [dict(c, local=True) for c in CARGAS],
           "mods": {k: {"eq": eq, "cargas": [dict(c, local=False) for c in cg]} for k, (eq, cg) in MODS.items()}})
(DIST / "Cemex_Orios.html").write_text(as_document(fill("Mantenimiento predictivo GDL", LOGO, "Mantenimiento_predictivo_GDL", seed, js(SEED_FILES), js(MODLOGOS))), encoding="utf-8")
# Artifact: fragmento sin <html>/<head> (el Artifact lo envuelve al publicar); los datos viven en su base compartida.
(DIST / "orios_gdl_artifact.html").write_text(fill("Mantenimiento predictivo GDL", ICON, "Mantenimiento_predictivo_GDL", "null", "{}"), encoding="utf-8")
print("dist/Cemex_Orios.html\ndist/orios_gdl_artifact.html")
