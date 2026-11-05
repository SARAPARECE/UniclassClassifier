"""Extraction of the classifier inputs from an IFC file (any authoring tool), so that the interoperable classifier
(feature version v14) can be applied without Revit.

Mapping of the inputs (Revit export -> IFC):
  Bounding_Box_Width/Depth/Height  axis-aligned box of the tessellated body in world coordinates (x, y, z)
  Volume                           Qto_*BaseQuantities NetVolume/GrossVolume, else volume of the closed mesh
  Centroid_Z                       centre of the box in z (height above the lowest element is computed later)
  Length, Height, Thickness/Width  Qto_*BaseQuantities Length, Height (or Depth), Width (or Thickness)
  Area                             Qto_*BaseQuantities NetSideArea/GrossSideArea (walls), NetArea/GrossArea, else 0
  load_bearing_status              Pset_*Common.LoadBearing
  Family and Type                  type name and element type (IfcTypeObject.Name, ObjectType) or element Name
  Category                         IFC class mapped to the Revit category, used only by the category check
Values that are not available are written as 0, which the classifier treats as missing, as in the Revit exports.
Usage: python ifc_extract.py model.ifc out.csv
"""
import sys

import numpy as np
import pandas as pd
import ifcopenshell
import ifcopenshell.geom
import ifcopenshell.util.element as E
import ifcopenshell.util.unit as UU

IFC_TO_CATEGORY = {"IfcWall": "Walls", "IfcWallStandardCase": "Walls", "IfcCurtainWall": "Walls", "IfcPlate": "Curtain Panels",
                   "IfcSlab": "Floors", "IfcRoof": "Roofs", "IfcDoor": "Doors", "IfcWindow": "Windows", "IfcStair": "Stairs",
                   "IfcStairFlight": "Stairs", "IfcColumn": "Structural Columns", "IfcBeam": "Structural Framing",
                   "IfcMember": "Structural Framing", "IfcFooting": "Structural Foundations"}


def _q(qto, *names):
    for n in names:
        v = qto.get(n)
        if isinstance(v, (int, float)) and v > 0:
            return float(v)
    return 0.0


def _mesh_volume(v, f):
    a, b, c = v[f[:, 0]], v[f[:, 1]], v[f[:, 2]]
    return abs(np.einsum("ij,ij->i", a, np.cross(b, c)).sum()) / 6.0


def extract(path):
    m = ifcopenshell.open(path)
    s = UU.calculate_unit_scale(m)          # file length unit -> metre (geometry is already returned in metres)
    sa, sv = UU.calculate_unit_scale(m, "AREAUNIT"), UU.calculate_unit_scale(m, "VOLUMEUNIT")
    st = ifcopenshell.geom.settings()
    st.set("use-world-coords", True)
    rows = []
    for cls in IFC_TO_CATEGORY:
        for el in m.by_type(cls, include_subtypes=False):
            if el.IsDecomposedBy and not el.Representation:
                continue                         # aggregate (for example an IfcRoof made of slabs): its parts are classified
            if el.is_a("IfcSlab") and E.get_predefined_type(el) == "ROOF":
                cat = "Roofs"
            else:
                cat = IFC_TO_CATEGORY[cls]
            try:
                sh = ifcopenshell.geom.create_shape(st, el)
                v = np.array(sh.geometry.verts).reshape(-1, 3)
                f = np.array(sh.geometry.faces).reshape(-1, 3)
            except Exception:
                v, f = np.zeros((0, 3)), np.zeros((0, 3), int)
            bb = (v.max(0) - v.min(0)) if len(v) else np.zeros(3)
            cz = float((v[:, 2].max() + v[:, 2].min()) / 2) if len(v) else 0.0
            ps = E.get_psets(el)
            qto = {k: val for n, d in ps.items() if n.startswith("Qto_") for k, val in d.items()}
            common = {k: val for n, d in ps.items() if n.startswith("Pset_") and n.endswith("Common") for k, val in d.items()}
            vol = sv * _q(qto, "NetVolume", "GrossVolume") or (_mesh_volume(v, f) if len(f) else 0.0)
            typ = E.get_type(el)
            name = " - ".join(x for x in ((typ.Name if typ else None), el.ObjectType, None if typ else el.Name) if x) or (el.Name or "")
            rows.append({"ElementID": el.GlobalId, "IfcClass": el.is_a(), "Category": cat, "Family and Type": name,
                         "Bounding_Box_Width": bb[0], "Bounding_Box_Depth": bb[1], "Bounding_Box_Height": bb[2],
                         "Volume": vol, "Centroid_Z": cz, "Number_of_Faces": 0,
                         "Length": s * _q(qto, "Length"), "Height": s * _q(qto, "Height", "Depth"), "Thickness/Width": s * _q(qto, "Width", "Thickness"),
                         "Area": sa * _q(qto, "NetSideArea", "GrossSideArea", "NetArea", "GrossArea"),
                         "load_bearing_status": "Load-Bearing" if common.get("LoadBearing") is True else "Non Load-Bearing"})
    return pd.DataFrame(rows), s


if __name__ == "__main__":
    df, s = extract(sys.argv[1])
    df.to_csv(sys.argv[2], index=False)
    print(len(df), "elements; length unit scale", s)
    print(df.groupby("Category").size().to_string())
