"""Uniclass Ss classifier v1: shared feature construction, training and prediction.

Final model (feature version v14, interoperable and category-free): XGBoost on bounding-box geometry, volume, height
above the lowest element, load-bearing flag and the base quantities Length, Height, Thickness/Width and Area (0 and
missing values treated as missing), averaged with a logistic regression on character n-grams of the family/type name.
Every input has an IFC equivalent (see ifc_extract.py). The Revit category (or the IFC class mapped to it) is not an
input; it is used only by the category check in decide(). Level 1 probabilities are the sums of the Level 2
probabilities within each Level 1 group, so the two levels are always consistent. Confidence gate: Level 2 if
p2 >= tau, else Level 1 if p1 >= tau, else Level 1 flagged for review; objects whose predicted Level 1 group differs
from the usual group of their category are also flagged for review.
Earlier versions (kept for the analyses of the paper): v1/v11/v12 use the Revit category one-hot and the face count.
"""
import re
import unicodedata
import numpy as np
import pandas as pd

CATEGORIES = ["Curtain Panels", "Doors", "Floors", "Roofs", "Stairs", "Structural Columns",
              "Structural Foundations", "Structural Framing", "Walls", "Windows"]
XGB_PARAMS = dict(objective="multi:softprob", eval_metric="mlogloss", tree_method="hist", n_estimators=400,
                  learning_rate=0.05, max_depth=6, min_child_weight=1.0, subsample=0.8, colsample_bytree=0.8,
                  reg_lambda=1.0, reg_alpha=0.0, n_jobs=1)


def _num(s):
    return pd.to_numeric(s, errors="coerce").fillna(0.0).astype(float).values


import os
FEATURE_VERSION = os.environ.get("UNICLASS_FEATURES", "v14")


def _num_nan(s):
    """Numeric value with 0 and missing treated as not available (the exports write 0 when no solid was read)."""
    v = pd.to_numeric(s, errors="coerce").astype(float).to_numpy(copy=True)
    v[v == 0] = np.nan
    return v


def features(df, version=None):
    """Geometry from the bounding box plus Revit category one-hot (one exported model or several).
    v1:  missing centroid, volume and face count coded as 0; height relative to the lowest value in the model.
    v11: missing values kept as missing (XGBoost handles them); height relative to the lowest available centroid."""
    fv = version or FEATURE_VERSION
    W, H, D = _num(df["Bounding_Box_Width"]), _num(df["Bounding_Box_Height"]), _num(df["Bounding_Box_Depth"])
    grp = df["project"] if "project" in df.columns else pd.Series(0, index=df.index)
    if fv == "v12":   # (v13 keeps the box and adds the Revit parameters instead; see below)
        # walls: oriented plan dimensions from the wall length and the type width instead of the axis-aligned box,
        # which is inflated for walls that are not parallel to the project axes
        Lw = _num(df["Length"]) if "Length" in df.columns else np.zeros(len(df))
        Tw = _num(df["Thickness/Width"]) if "Thickness/Width" in df.columns else np.zeros(len(df))
        wall = (df["Category"].astype(str).values == "Walls") & (Lw > 0) & (Tw > 0)
        W, D = np.where(wall, Lw, W), np.where(wall, Tw, D)
    if fv in ("v11", "v12", "v13", "v14"):
        vol, z, faces = _num_nan(df["Volume"]), _num_nan(df["Centroid_Z"]), _num_nan(df["Number_of_Faces"])
        zrel = z - pd.Series(z, index=df.index).groupby(grp.values).transform("min").values
    else:
        vol, z, faces = _num(df["Volume"]), _num(df["Centroid_Z"]), _num(df["Number_of_Faces"])
        zrel = z - pd.Series(z, index=df.index).groupby(grp.values).transform("min").values
    dims = np.sort(np.column_stack([W, H, D]), axis=1)[:, ::-1]
    hmax, hmin = np.maximum(W, D), np.minimum(W, D)
    e = 1e-6
    lb = (df["load_bearing_status"].astype(str).str.strip().str.lower() == "load-bearing").astype(float).values
    geo = np.column_stack([dims, H, hmax, hmin, H / (hmax + e), dims[:, 1] / (dims[:, 0] + e),
                           dims[:, 2] / (dims[:, 0] + e), vol, vol / (W * H * D + e), zrel,
                           faces, lb])
    if fv == "v14":
        # interoperable: category-free and without the face count, which depends on how the authoring tool represents
        # the geometry; every input has an IFC equivalent (bounding box and volume from the geometry, Qto_ base
        # quantities, Pset_*Common.LoadBearing, Name/ObjectType)
        par = np.column_stack([_num_nan(df[c]) if c in df.columns else np.full(len(df), np.nan)
                               for c in ("Length", "Height", "Thickness/Width", "Area")])
        geo = np.delete(geo, 12, axis=1)
        if POSITION != "none":
            return np.column_stack([geo, par, position_features(df, POSITION)])
        return np.column_stack([geo, par])
    if fv == "v13":
        # category-free: the Revit category is not used, so that it can be checked against the prediction;
        # the Revit parameters give oriented dimensions where the family provides them (0 treated as missing)
        par = np.column_stack([_num_nan(df[c]) if c in df.columns else np.full(len(df), np.nan)
                               for c in ("Length", "Height", "Thickness/Width", "Area")])
        if POSITION != "none":
            return np.column_stack([geo, par, position_features(df, POSITION)])
        return np.column_stack([geo, par])
    cat = np.column_stack([(df["Category"].astype(str) == c).astype(float).values for c in CATEGORIES])
    if POSITION != "none":
        return np.column_stack([geo, cat, position_features(df, POSITION)])
    return np.column_stack([geo, cat])


POSITION = os.environ.get("UNICLASS_POSITION", "none")   # exploratory: none, z, zxy, zedge


def _hull_distance(px, py):
    """Distance of each point to the boundary of the convex hull of all points (0 outside or on the hull)."""
    from scipy.spatial import ConvexHull
    ok = ~(np.isnan(px) | np.isnan(py))
    d = np.full(len(px), np.nan)
    pts = np.unique(np.column_stack([px[ok], py[ok]]), axis=0)
    if len(pts) < 3:
        return d
    h = pts[ConvexHull(pts).vertices]
    a, b = h, np.roll(h, -1, axis=0)
    P = np.column_stack([px[ok], py[ok]])[:, None, :]
    ab = (b - a)[None, :, :]
    t = np.clip(((P - a[None]) * ab).sum(2) / np.maximum((ab ** 2).sum(2), 1e-12), 0, 1)
    proj = a[None] + t[..., None] * ab
    d[ok] = np.sqrt(((P - proj) ** 2).sum(2)).min(1)
    return d


def position_features(df, mode):
    """Position relative to each model: normalised height (0 base, 1 top), normalised plan coordinates and the
    distance to the boundary of the plan (convex hull of all element positions). Centroids are used where available,
    otherwise the insertion point (Start); zeros are treated as missing."""
    grp = (df["project"] if "project" in df.columns else pd.Series(0, index=df.index)).values
    x = np.where(np.isnan(_num_nan(df["Centroid_X"])), _num_nan(df["Start_X"]) if "Start_X" in df else np.nan, _num_nan(df["Centroid_X"]))
    y = np.where(np.isnan(_num_nan(df["Centroid_Y"])), _num_nan(df["Start_Y"]) if "Start_Y" in df else np.nan, _num_nan(df["Centroid_Y"]))
    z = _num_nan(df["Centroid_Z"])
    out = np.full((len(df), 5), np.nan)
    for g in np.unique(grp):
        m = grp == g
        for j, v in enumerate((z, x, y)):
            q1, q9 = np.nanpercentile(v[m], [1, 99]) if np.isfinite(v[m]).any() else (np.nan, np.nan)
            out[m, j] = (v[m] - q1) / (q9 - q1) if q9 > q1 else np.nan
        de = _hull_distance(x[m], y[m])
        out[m, 3] = de
        out[m, 4] = de / np.nanmax(de) if np.isfinite(de).any() and np.nanmax(de) > 0 else np.nan
    cols = {"z": [0], "zxy": [0, 1, 2], "zedge": [0, 3, 4]}[mode]
    return out[:, cols]


def norm_name(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"\s+", " ", re.sub(r"[^a-z]+", " ", re.sub(r"[0-9]+", " ", s))).strip()


def train(df, labels, use_text=True, seed=42):
    from xgboost import XGBClassifier
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.linear_model import LogisticRegression
    codes = np.unique(labels)
    y = np.searchsorted(codes, labels)
    m = XGBClassifier(random_state=seed, **XGB_PARAMS).fit(features(df, FEATURE_VERSION), y)
    bundle = {"version": "1.1" if FEATURE_VERSION in ("v13", "v14") else "1.0", "codes": codes, "xgb": m, "use_text": use_text,
              "features": FEATURE_VERSION}
    # category check: most frequent Level 1 group of each Revit category in the training data
    t = pd.DataFrame({"c": df["Category"].astype(str).values, "g": [str(c)[:5] for c in labels]})
    bundle["cat_l1"] = t.groupby("c").g.agg(lambda s: s.value_counts().index[0]).to_dict()
    if use_text:
        vec = TfidfVectorizer(analyzer="char_wb", ngram_range=(3, 5), min_df=2, sublinear_tf=True)
        T = vec.fit_transform(df["Family and Type"].map(norm_name))
        bundle["vec"] = vec
        bundle["lr"] = LogisticRegression(C=10, max_iter=3000).fit(T, y)
    return bundle


def save(bundle, path_stem):
    """Portable files: XGBoost booster in its JSON format; codes and the text model in a joblib file."""
    import joblib
    bundle["xgb"].save_model(path_stem + "_xgb.json")
    rest = {k: v for k, v in bundle.items() if k != "xgb"}
    rest["n_classes"] = len(bundle["codes"])
    joblib.dump(rest, path_stem + ".joblib", compress=3)


def load(path_stem):
    import joblib
    from xgboost import XGBClassifier
    b = joblib.load(path_stem + ".joblib")
    m = XGBClassifier()
    m.load_model(path_stem + "_xgb.json")
    b["xgb"] = m
    return b


def proba(bundle, df):
    nC = len(bundle["codes"])
    P = np.zeros((len(df), nC))
    P[:, :] = bundle["xgb"].predict_proba(features(df, bundle.get("features", FEATURE_VERSION)))
    if bundle["use_text"]:
        pt = np.zeros((len(df), nC))
        pt[:, bundle["lr"].classes_] = bundle["lr"].predict_proba(bundle["vec"].transform(df["Family and Type"].map(norm_name)))
        P = (P + pt) / 2
    return P


def decide(bundle, P, tau=0.8, tau_l1=None, categories=None):
    """Level 2 if p2 >= tau; else Level 1 if p1 >= tau_l1 (default: tau); else Level 1 proposed and flagged for review.
    Category check (when categories are given): if the predicted Level 1 group differs from the most frequent group of
    the stated Revit category in the training data, the object is flagged for review whatever its confidence."""
    tau_l1 = tau if tau_l1 is None else tau_l1
    codes = bundle["codes"]
    par = np.array([c[:5] for c in codes])
    up = np.unique(par)
    M1 = np.zeros((P.shape[0], len(up)))
    for j, c in enumerate(par):
        M1[:, np.searchsorted(up, c)] += P[:, j]
    l2, l1 = codes[P.argmax(1)], up[M1.argmax(1)]
    p2, p1 = P.max(1), M1.max(1)
    level = np.where(p2 >= tau, "L2", np.where(p1 >= tau_l1, "L1", "L1 review"))
    expected = np.full(len(l1), "", dtype=object)
    mismatch = np.zeros(len(l1), bool)
    if categories is not None and bundle.get("cat_l1"):
        expected = np.array([bundle["cat_l1"].get(str(c), "") for c in categories], dtype=object)
        mismatch = (expected != "") & (expected != l1)
        level = np.where(mismatch, "L1 review", level)
    code = np.where(level == "L2", l2, l1)
    return pd.DataFrame({"code": code, "level": level, "L2_pred": l2, "p_L2": p2.round(3), "L1_pred": l1, "p_L1": p1.round(3),
                         "category_group": expected, "category_mismatch": mismatch})
