# -*- coding: utf-8 -*-
# Auto-gerado

import pandas as pd
import numpy as np
import re
import sys
import joblib
import json
import os

# Argumentos: CSV e JSON de entrada
csv_path = sys.argv[1]
json_path = sys.argv[2]
model_path = sys.argv[3]

base_dir = os.path.dirname(json_path)
out_csv = os.path.join(base_dir, "datapreparation.csv")
out_json = os.path.join(base_dir, "classified.json")

# Debug para ver os ficheiros recebidos
print(f"Recebido CSV: {csv_path}")
print(f"Recebido JSON: {json_path}")

# Carregar dados
df = pd.read_json(json_path)

# Eliminar colunas não usadas
drop_cols = [
    'SECClass_Code_EF', 'SECClass_Title_EF', 'SECClasS_Title_Ss',
    'SECClass_Code_Pr', 'SECClass_Title_Pr', 'Family and Type', 'SECClasS_Code_Ss',
    'Total_Surface_Area', 'Total_Edge_Length', 'Base_extension_distance',
    'Start_X', 'Start_Y', 'Start_Z', 'End_X', 'End_Y', 'End_Z',
    'Aspect_Ratio', 'Volume_to_Surface_Area_Ratio',
    'Phase_Created', 'Comments', 'Keynote', 'Description', 'Materials'
]
df.drop(columns=[col for col in drop_cols if col in df.columns], inplace=True)

# Codificar categorias
df['ElementID'] = df['ElementID'].astype('category')
df['Category'] = df['Category'].astype('category')
df = pd.get_dummies(df, columns=['Category'])

# Forçar colunas principais a existirem
for category in ['Doors', 'Floors', 'Roofs', 'Stairs', 'Structural Columns',
                 'Structural Foundations', 'Structural Framing', 'Walls', 'Windows', 'Curtain Panels']:
    col = f'Category_{category}'
    if col not in df.columns:
        df[col] = 0
    df[col] = df[col].astype(int)

# load_bearing_status
df['load_bearing_status'] = df['load_bearing_status'].replace("None", "Non Load-Bearing")
df['load_bearing_status_binary'] = df['load_bearing_status'].astype('category')
df = pd.get_dummies(df, columns=['load_bearing_status_binary'], drop_first=True)
if 'load_bearing_status_binary_Non Load-Bearing' not in df.columns:
    df['load_bearing_status_binary_Non Load-Bearing'] = 0
df['load_bearing_status_binary_Non Load-Bearing'] = df['load_bearing_status_binary_Non Load-Bearing'].astype(float)

# ✅ Mapear Curvature para valores numéricos ANTES de converter para float
if 'Curvature' in df.columns:
    curvature_map = {
        'Bound': 0,
        'Unbound': 1,
        'Cyclic': 2,
        'None': 0  # default
    }
    df['Curvature'] = df['Curvature'].map(curvature_map).fillna(0)

# Variáveis numéricas (força conversão para float, substitui qualquer erro por 0)
numeric_cols = [
    'Length', 'Height', 'Thickness/Width',
    'Base_offset', 'Top_offset', 'Number_of_Faces',
    'Bounding_Box_Width', 'Bounding_Box_Height', 'Bounding_Box_Depth',
    'Centroid_X', 'Centroid_Y', 'Centroid_Z',
    'Orientation_Angle', 'Curvature'
]

for col in numeric_cols:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

# Classificação dos pisos
def extrair_piso(piso):
    if pd.isna(piso):
        return np.nan
    piso = str(piso).upper()
    match = re.search(r'PISO\s*(-?\d+)', piso)
    if match:
        return int(match.group(1))
    elif 'COBERTURA' in piso:
        return 100
    elif 'CAVE' in piso or 'SUBSOLO' in piso:
        return -10
    return np.nan

piso_cols = ['Base_constraint', 'Top_constraint', 'Level', 'Base_level', 'Top_level']
for col in piso_cols:
    if col in df.columns:
        df[f'{col}_num'] = df[col].apply(extrair_piso)

todos_pisos = pd.concat([df[f'{col}_num'] for col in piso_cols if f'{col}_num' in df.columns])
piso_superior = todos_pisos.max(skipna=True)
piso_inferior = todos_pisos.min(skipna=True)

def classificar_piso(valor):
    if pd.isna(valor):
        return 'PISO_INDETERMINADO'
    elif valor == piso_superior:
        return 'PISO_SUPERIOR'
    elif valor == piso_inferior:
        return 'PISO_INFERIOR'
    else:
        return 'PISO_INTERMEDIÁRIO'

for col in piso_cols:
    if f'{col}_num' in df.columns:
        df[f'{col}_cat'] = df[f'{col}_num'].apply(classificar_piso).astype('category')

cat_cols = [f"{col}_cat" for col in piso_cols if f"{col}_cat" in df.columns]
df_dummies = pd.get_dummies(df[cat_cols], prefix=cat_cols)
for col in piso_cols:
    for cat in ['PISO_INDETERMINADO', 'PISO_INFERIOR', 'PISO_INTERMEDIÁRIO', 'PISO_SUPERIOR']:
        dummy_col = f"{col}_cat_{cat}"
        if dummy_col not in df_dummies.columns:
            df_dummies[dummy_col] = 0
df_dummies = df_dummies.astype(int)
df = pd.concat([df, df_dummies], axis=1)

# Remover colunas temporárias
cols_to_remove = (
    ['load_bearing_status'] +
    piso_cols +
    [f"{col}_num" for col in piso_cols] +
    [f"{col}_cat" for col in piso_cols]
)
df.drop(columns=[col for col in cols_to_remove if col in df.columns], inplace=True)

# Guardar CSV
df.to_csv(out_csv, index=False)



# Classificação
model = joblib.load(os.path.join(model_path, "decision_tree_sara_model.pkl"))
encoder = joblib.load(os.path.join(model_path, "label_encoder_sara_train.pkl"))

# Antes de prever
X_new = df.drop(columns=["ElementID"])

# Garante que TODAS as colunas existem
for col in model.feature_names_in_:
    if col not in X_new.columns:
        X_new[col] = 0

# Ordena as colunas como no treino
X_new = X_new[model.feature_names_in_]

# FORÇA todos os valores para numéricos, substituindo strings ou None por 0
X_new = X_new.apply(pd.to_numeric, errors='coerce').fillna(0.0)

# Agora sim previsões
y_pred = model.predict(X_new)
y_pred_labels = encoder.inverse_transform(y_pred)

results = dict(zip(df['ElementID'].astype(str), y_pred_labels))
with open(out_json, 'w') as f:
    json.dump(results, f, indent=2)

print("Classificação concluída com sucesso.")
