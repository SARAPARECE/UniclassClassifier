# -*- coding: utf-8 -*-
"""Uniclass Ss classification of an IFC file (any authoring tool), without Revit.

Usage: python predict_ifc.py <model.ifc> <out.csv> <model_dir> [tau_L2] [tau_L1]
Extracts the inputs with ifc_extract.py (IfcOpenShell), applies the final model, the confidence gate and the category
check (IFC class mapped to the Revit category), and writes one row per element with its GlobalId.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import uniclass_v1 as U
from ifc_extract import extract

ifc, out, model_dir = sys.argv[1], sys.argv[2], sys.argv[3]
tau2 = float(sys.argv[4]) if len(sys.argv) > 4 else 0.8
tau1 = float(sys.argv[5]) if len(sys.argv) > 5 else tau2
df, _ = extract(ifc)
df["project"] = os.path.basename(ifc)
b = U.load(os.path.join(model_dir, os.environ.get("UNICLASS_MODEL", "uniclass_v1")))
res = U.decide(b, U.proba(b, df), tau2, tau1, categories=df["Category"].values)
res.insert(0, "GlobalId", df["ElementID"].values)
res.insert(1, "IfcClass", df["IfcClass"].values)
res.insert(2, "Name", df["Family and Type"].values)
res.to_csv(out, index=False, encoding="utf-8-sig")
print(f"Classified {len(res)} elements: " + ", ".join(f"{k} {v}" for k, v in res.level.value_counts().items()))
