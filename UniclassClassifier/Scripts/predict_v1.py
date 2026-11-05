# -*- coding: utf-8 -*-
"""Uniclass Ss classification v1, called by the Revit add-in (ClassifierCommand).
The model does not use the Revit category; the category is only compared with the prediction (category check).

Usage: python predict_v1.py <csv_path> <json_path> <model_dir> [tau_L2] [tau_L1]
  json_path  elements exported by the add-in (export.json)
  model_dir  folder with uniclass_v1.joblib and uniclass_v1_xgb.json
  tau_L2     confidence threshold for Sub-group (Level 2) codes, default 0.8
  tau_L1     confidence threshold for Group (Level 1) codes, default tau_L2
Environment: UNICLASS_MODEL selects another model file stem in model_dir (e.g. uniclass_v1_noP6).

Writes next to json_path:
  classified.json        {ElementID: code}: Level 2 code if p2 >= tau_L2, else Level 1 code if p1 >= tau_L1;
                         elements below both thresholds are not included, so existing codes are kept
  classified_report.csv  per element: code, level (L2 / L1 / L1 review), Level 2 and Level 1 proposals, confidences,
                         the usual Level 1 group of the element's category and whether the prediction disagrees with it
  summary.json           counts per level and the ElementIDs left for review
"""
import json
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import uniclass_v1 as U

json_path, model_dir = sys.argv[2], sys.argv[3]
tau2 = float(sys.argv[4]) if len(sys.argv) > 4 else float(os.environ.get("UNICLASS_TAU", "0.8"))
tau1 = float(sys.argv[5]) if len(sys.argv) > 5 else tau2
model = os.environ.get("UNICLASS_MODEL", "uniclass_v1")
base = os.path.dirname(os.path.abspath(json_path))

df = pd.read_json(json_path, dtype=False)
if df.empty:
    raise SystemExit("No elements were exported.")
for c in ["Bounding_Box_Width", "Bounding_Box_Height", "Bounding_Box_Depth", "Volume", "Centroid_Z", "Number_of_Faces",
          "Length", "Height", "Thickness/Width", "Area"]:
    if c not in df.columns:
        df[c] = 0
for c in ["Category", "Family and Type", "load_bearing_status"]:
    if c not in df.columns:
        df[c] = ""

bundle = U.load(os.path.join(model_dir, model))
res = U.decide(bundle, U.proba(bundle, df), tau2, tau1, categories=df["Category"].astype(str).values)
res.insert(0, "ElementID", df["ElementID"].astype(str).values)
res.insert(1, "Family and Type", df["Family and Type"].values)
res.insert(2, "Category", df["Category"].values)
res.to_csv(os.path.join(base, "classified_report.csv"), index=False, encoding="utf-8-sig")

keep = res.level != "L1 review"
with open(os.path.join(base, "classified.json"), "w", encoding="utf-8") as f:
    json.dump(dict(zip(res.ElementID[keep], res.code[keep])), f, indent=2)
counts = res.level.value_counts()
summary = {"model": model, "tau_L2": tau2, "tau_L1": tau1, "n": int(len(res)),
           "L2": int(counts.get("L2", 0)), "L1": int(counts.get("L1", 0)), "review": int(counts.get("L1 review", 0)),
           "category_check": int(res.category_mismatch.sum()), "category_check_ids": res.ElementID[res.category_mismatch].tolist(),
           "review_ids": res.ElementID[~keep].tolist()}
with open(os.path.join(base, "summary.json"), "w", encoding="utf-8") as f:
    json.dump(summary, f, indent=2)
print(f"Classified {summary['n']} elements: Level 2 {summary['L2']}, Level 1 {summary['L1']}, review {summary['review']}")
