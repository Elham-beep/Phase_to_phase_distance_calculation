"""
IFC ➜ OBJ converter
• uses world coords
• optional pythonocc-core fallback (auto-detected, never imported directly)
• skips and reports elements that still fail
"""

import logging
import pathlib
import importlib.util

import ifcopenshell
import ifcopenshell.geom

# ──────────────────────────────── paths ────────────────────────────────
IN_IFC  = r"C:\Users\bimax\DC\ACCDocs\Axpo Grid AG\QC_test_Elham\I-17736_RP_MOD_KOO_NEU3.ifc"
OUT_OBJ = r"C:\Users\bimax\DC\ACCDocs\Axpo Grid AG\QC_test_Elham\building.obj"

# ─────────────────────────────── logging ───────────────────────────────
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

# ─────────────────────── detect pythonocc-core safely ──────────────────
PY_OCC_AVAILABLE = importlib.util.find_spec("OCC.Core") is not None
logging.info("OCC fallback %s", "enabled" if PY_OCC_AVAILABLE else "disabled")

# ────────────────────── helper to set settings safely ──────────────────
settings = ifcopenshell.geom.settings()

def safe_set(key: str, value):
    if key in settings.setting_names():
        settings.set(key, value)

safe_set("use-world-coords", True)
safe_set("sew-shells", False)
safe_set("disable-opening-subtractions", False)

# ───────────────────────── open IFC file ───────────────────────────────
if not pathlib.Path(IN_IFC).is_file():
    raise FileNotFoundError(IN_IFC)
model = ifcopenshell.open(IN_IFC)

verts_all, faces_all, failed = [], [], []

def append_shape(shape):
    verts, faces = shape.geometry.verts, shape.geometry.faces
    offset = len(verts_all) // 3
    verts_all.extend(verts)
    faces_all.extend([i + offset for i in faces])

# ───────────────────── iterate over products ───────────────────────────
logging.info("Serialising geometry …")
for prod in model.by_type("IfcProduct"):
    if not prod.Representation:
        continue
    try:
        append_shape(ifcopenshell.geom.create_shape(settings, prod))

    except RuntimeError:
        # optional OCC fallback
        if (PY_OCC_AVAILABLE and
            "use-python-opencascade" in settings.setting_names()):
            try:
                settings.set("use-python-opencascade", True)
                append_shape(ifcopenshell.geom.create_shape(settings, prod))
                settings.set("use-python-opencascade", False)
                continue
            except Exception:
                settings.set("use-python-opencascade", False)

        failed.append((prod.GlobalId, prod.is_a()))
        logging.debug("Skipped %s (%s)", prod.GlobalId, prod.is_a())

# ───────────────────────── write OBJ file ──────────────────────────────
logging.info("Writing OBJ → %s", OUT_OBJ)
with open(OUT_OBJ, "w", encoding="utf-8") as f:
    for x, y, z in zip(*[iter(verts_all)] * 3):
        f.write(f"v {x} {y} {z}\n")
    for a, b, c in zip(*[iter(faces_all)] * 3):
        f.write(f"f {a+1} {b+1} {c+1}\n")

# ───────────────────────────── summary ─────────────────────────────────
logging.info("Done.  Vertices: %d  Faces: %d  Skipped elements: %d",
             len(verts_all)//3, len(faces_all)//3, len(failed))
if failed:
    logging.warning("First skipped elements:\n%s",
                    "\n".join(f"{g} ({c})" for g, c in failed[:10]))
