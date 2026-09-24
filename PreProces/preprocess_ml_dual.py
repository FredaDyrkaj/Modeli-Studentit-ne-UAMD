"""
Preprocessing dual për projektin «Modeli i studentit optimal në UAMD».

Dy task-e ML:
  1) Regresion  → target MES_PERGJ
  2) Klasifikim → target RREZIK_VONESE (rregull hibrid)

Predictors: demografi + matura/profil + notat e Vitit 1 (pa leakage nga vite të mëvonshme).
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterable

import joblib
import numpy as np
import pandas as pd
from sklearn.compose import ColumnTransformer
from sklearn.impute import SimpleImputer
from sklearn.model_selection import KFold, train_test_split
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder, StandardScaler, TargetEncoder

# ---------------------------------------------------------------------------
# Konfigurim
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = (
    BASE_DIR.parent / "DescriptiveAnalyse" / "Studentet_Uamd_pastruar_spss.xlsx"
)
OUT_DIR = BASE_DIR
RANDOM_STATE = 42
TEST_SIZE = 0.20

# Predictors të lejuar (informacion në fillim + vit i parë)
NUMERIC_FEATURES = [
    "MOSHA",
    "Baze_11",
    "Zgjedhje_11",
    "Baze_12",
    "Zgjedhje_12",
    "MES_VIT1_SEM1",
]
LOW_CARD_CATEGORICAL = ["GJINIA", "LLOJ_RREGJ", "PROFIL_GJIMNAZ"]
HIGH_CARD_CATEGORICAL = ["RRETHI", "GJIMNAZI", "DEGA", "FAKULTETI"]
ALL_FEATURES = NUMERIC_FEATURES + LOW_CARD_CATEGORICAL + HIGH_CARD_CATEGORICAL

# Kolona që shkaktojnë leakage / nuk janë predictors
LEAKAGE_OR_UNUSED = [
    "Nr",
    "DATA_DIPLOMA",
    "VIT_DIPLOMIM",
    "KOHA_DIPLOM",
    "MES_BAZE",
    "MES_ZGJEDHJE",
    "MES_PERGJ_BACH",
    "MES_PERGJ_MAST",
    "MES_PERGJ",
]

# Mapimi i vitit akademik (romak) → vite studimi
VIT_AKADEMIK_MAP = {
    "I": 1,
    "II": 2,
    "III": 3,
    "IV": 4,
    "V": 5,
    "VI": 6,
}


# ===========================================================================
# 1) Ngarkimi & përzgjedhja e kolonave
# ===========================================================================
def load_raw_data(path: Path = DATA_FILE) -> pd.DataFrame:
    """Ngarkon Excel-in burimor."""
    if not path.exists():
        raise FileNotFoundError(f"Nuk u gjet skedari i të dhënave: {path}")
    df = pd.read_excel(path)
    print(f"[LOAD] {path.name} → shape={df.shape}")
    return df


def drop_later_semester_grades(df: pd.DataFrame) -> pd.DataFrame:
    """Hiq notat e viteve të mëvonshme (Baze_/Zgjedhje_21 … 52)."""
    to_drop = [
        c
        for c in df.columns
        if c.startswith(("Baze_", "Zgjedhje_"))
        and c not in {"Baze_11", "Zgjedhje_11", "Baze_12", "Zgjedhje_12"}
    ]
    return df.drop(columns=to_drop, errors="ignore")


# ===========================================================================
# 2) Feature engineering — RREZIK_VONESE
# ===========================================================================
def norma_vite(niveli: object) -> int:
    """NORMA_VITE: 3 për Bachelor, 2 për Master."""
    text = str(niveli).strip().lower()
    if "master" in text:
        return 2
    return 3


def map_vit_akademik(value: object) -> float:
    """Kthen vitin e studimit si numër; NaN për DIPLOMUAR / të panjohura."""
    if pd.isna(value):
        return np.nan
    key = str(value).strip().upper()
    return float(VIT_AKADEMIK_MAP[key]) if key in VIT_AKADEMIK_MAP else np.nan


def build_rrezik_vonese(df: pd.DataFrame) -> pd.DataFrame:
    """
    Rregull hibrid për RREZIK_VONESE.

    1) NORMA_VITE nga NIVELI_STUDIMIT (Bachelor=3, Master=2).
    2) Me KOHA_DIPLOM: 1 nëse KOHA_DIPLOM > NORMA, else 0.
    3) Pa KOHA_DIPLOM: KOHA_AKTUALE = map(VIT_AKADEMIK);
       - > NORMA → 1
       - <= NORMA / NaN / DIPLOMUAR → NaN (përjashtohet nga klasifikimi).
    """
    out = df.copy()
    out["NORMA_VITE"] = out["NIVELI_STUDIMIT"].map(norma_vite)
    out["KOHA_AKTUALE"] = out["VIT_AKADEMIK"].map(map_vit_akademik)

    rrezik = pd.Series(np.nan, index=out.index, dtype="float64")

    has_koha = out["KOHA_DIPLOM"].notna()
    rrezik.loc[has_koha] = (
        out.loc[has_koha, "KOHA_DIPLOM"] > out.loc[has_koha, "NORMA_VITE"]
    ).astype(float)

    no_koha = ~has_koha
    over_norm = no_koha & out["KOHA_AKTUALE"].notna() & (
        out["KOHA_AKTUALE"] > out["NORMA_VITE"]
    )
    rrezik.loc[over_norm] = 1.0
    # rastet no_koha & KOHA_AKTUALE <= NORMA (ose NaN) mbeten NaN → përjashtohen

    out["RREZIK_VONESE"] = rrezik
    return out


# ===========================================================================
# 3) Ndërtimi i dataset-eve të papërpunuara (para encoding/scaling)
# ===========================================================================
def _ensure_feature_columns(df: pd.DataFrame, features: Iterable[str]) -> None:
    missing = [c for c in features if c not in df.columns]
    if missing:
        raise KeyError(f"Mungojnë kolonat e kërkuara: {missing}")


def build_regression_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Dataset regresioni: predictors + MES_PERGJ (pa NaN në target)."""
    _ensure_feature_columns(df, ALL_FEATURES + ["MES_PERGJ"])
    frame = df[ALL_FEATURES + ["MES_PERGJ"]].copy()
    before = len(frame)
    frame = frame.dropna(subset=["MES_PERGJ"]).reset_index(drop=True)
    print(f"[REG] Hiqen {before - len(frame)} rreshta pa MES_PERGJ → n={len(frame)}")
    return frame


def build_classification_frame(df: pd.DataFrame) -> pd.DataFrame:
    """
    Dataset klasifikimi: predictors + RREZIK_VONESE.
    Përjashton studentët aktivë pa vonesë (RREZIK_VONESE NaN).
    """
    _ensure_feature_columns(df, ALL_FEATURES + ["RREZIK_VONESE"])
    frame = df[ALL_FEATURES + ["RREZIK_VONESE"]].copy()
    before = len(frame)
    frame = frame.dropna(subset=["RREZIK_VONESE"]).reset_index(drop=True)
    frame["RREZIK_VONESE"] = frame["RREZIK_VONESE"].astype(int)
    print(
        f"[CLF] Hiqen {before - len(frame)} rreshta me status të pacaktuar "
        f"→ n={len(frame)}"
    )
    return frame


# ===========================================================================
# 4) Pipeline scikit-learn (impute + encode + scale)
# ===========================================================================
def make_preprocessor(task: str) -> ColumnTransformer:
    """
    ColumnTransformer i ripërdorshëm.

    task: 'regression' | 'classification'  (përcakton TargetEncoder.target_type)
    """
    if task not in {"regression", "classification"}:
        raise ValueError("task duhet të jetë 'regression' ose 'classification'")

    target_type = "continuous" if task == "regression" else "binary"

    numeric_pipe = Pipeline(
        steps=[
            ("imputer", SimpleImputer(strategy="median")),
            ("scaler", StandardScaler()),
        ]
    )
    onehot_pipe = Pipeline(
        steps=[
            ("imputer", SimpleImputer(strategy="most_frequent")),
            (
                "onehot",
                OneHotEncoder(
                    drop="first",
                    handle_unknown="ignore",
                    sparse_output=False,
                ),
            ),
        ]
    )
    target_pipe = Pipeline(
        steps=[
            ("imputer", SimpleImputer(strategy="most_frequent")),
            (
                "target_encoder",
                TargetEncoder(
                    target_type=target_type,
                    cv=KFold(
                        n_splits=5, shuffle=True, random_state=RANDOM_STATE
                    ),
                ),
            ),
        ]
    )

    return ColumnTransformer(
        transformers=[
            ("num", numeric_pipe, NUMERIC_FEATURES),
            ("oh", onehot_pipe, LOW_CARD_CATEGORICAL),
            ("te", target_pipe, HIGH_CARD_CATEGORICAL),
        ],
        remainder="drop",
        verbose_feature_names_out=True,
    )


def fit_transform_split(
    frame: pd.DataFrame,
    target_col: str,
    task: str,
    stratify: bool = False,
) -> tuple[pd.DataFrame, Pipeline, dict]:
    """
    Split 80/20 PARA fit-it, pastaj fit pipeline mbi train dhe transform train+test.
    Kthen DataFrame të përpunuar (me kolonën 'set') + pipeline + meta.
    """
    X = frame[ALL_FEATURES].copy()
    y = frame[target_col].copy()

    stratify_y = y if stratify else None
    X_train, X_test, y_train, y_test = train_test_split(
        X,
        y,
        test_size=TEST_SIZE,
        random_state=RANDOM_STATE,
        stratify=stratify_y,
    )

    preprocessor = make_preprocessor(task)
    pipe = Pipeline(steps=[("preprocess", preprocessor)])

    X_train_t = pipe.fit_transform(X_train, y_train)
    X_test_t = pipe.transform(X_test)
    feature_names = pipe.named_steps["preprocess"].get_feature_names_out()

    train_df = pd.DataFrame(X_train_t, columns=feature_names)
    train_df[target_col] = y_train.to_numpy()
    train_df["set"] = "train"

    test_df = pd.DataFrame(X_test_t, columns=feature_names)
    test_df[target_col] = y_test.to_numpy()
    test_df["set"] = "test"

    processed = pd.concat([train_df, test_df], ignore_index=True)
    meta = {
        "n_train": len(train_df),
        "n_test": len(test_df),
        "n_features": len(feature_names),
        "feature_names": list(feature_names),
    }
    return processed, pipe, meta


# ===========================================================================
# 5) Ruajtja & përmbledhja
# ===========================================================================
def export_outputs(
    reg_df: pd.DataFrame,
    clf_df: pd.DataFrame,
    reg_pipe: Pipeline,
    clf_pipe: Pipeline,
    out_dir: Path = OUT_DIR,
) -> None:
    """Ruaj CSV-të e përpunuara dhe pipeline-et joblib."""
    out_dir.mkdir(parents=True, exist_ok=True)

    reg_csv = out_dir / "dataset_ml_regression.csv"
    clf_csv = out_dir / "dataset_ml_classification.csv"
    reg_joblib = out_dir / "preprocessing_pipeline_regression.joblib"
    clf_joblib = out_dir / "preprocessing_pipeline_classification.joblib"

    reg_df.to_csv(reg_csv, index=False)
    clf_df.to_csv(clf_csv, index=False)
    joblib.dump(reg_pipe, reg_joblib)
    joblib.dump(clf_pipe, clf_joblib)

    print(f"[SAVE] {reg_csv.name}")
    print(f"[SAVE] {clf_csv.name}")
    print(f"[SAVE] {reg_joblib.name}")
    print(f"[SAVE] {clf_joblib.name}")


def print_summary(
    reg_raw: pd.DataFrame,
    clf_raw: pd.DataFrame,
    reg_proc: pd.DataFrame,
    clf_proc: pd.DataFrame,
    reg_meta: dict,
    clf_meta: dict,
) -> None:
    """Përmbledhje e formave, kolonave dhe shpërndarjes së target-eve."""
    print("\n" + "=" * 70)
    print("PËRMBLEDHJE — PREPROCESSING DUAL ML")
    print("=" * 70)

    print("\n[REGRESSION]")
    print(f"  Raw (para encoding): {reg_raw.shape}")
    print(f"  Processed:           {reg_proc.shape}")
    print(f"  Train/Test:          {reg_meta['n_train']} / {reg_meta['n_test']}")
    print(f"  # veçori të koduara: {reg_meta['n_features']}")
    print(f"  MES_PERGJ describe:\n{reg_raw['MES_PERGJ'].describe()}")

    print("\n[CLASSIFICATION]")
    print(f"  Raw (para encoding): {clf_raw.shape}")
    print(f"  Processed:           {clf_proc.shape}")
    print(f"  Train/Test:          {clf_meta['n_train']} / {clf_meta['n_test']}")
    print(f"  # veçori të koduara: {clf_meta['n_features']}")
    dist = clf_raw["RREZIK_VONESE"].value_counts().sort_index()
    pct = clf_raw["RREZIK_VONESE"].value_counts(normalize=True).sort_index() * 100
    print("  Shpërndarja e RREZIK_VONESE (0/1):")
    for label in dist.index:
        print(f"    klasa {label}: n={dist[label]} ({pct[label]:.1f}%)")

    print("\n[KOLONAT E PREDICTORS]")
    print(f"  {ALL_FEATURES}")
    print("=" * 70)


# ===========================================================================
# Main
# ===========================================================================
def main() -> None:
    raw = load_raw_data()
    raw = drop_later_semester_grades(raw)
    raw = build_rrezik_vonese(raw)

    # Diagnostikë e shkurtër për target-in e klasifikimit
    n_labeled = raw["RREZIK_VONESE"].notna().sum()
    n_pos = (raw["RREZIK_VONESE"] == 1).sum()
    n_neg = (raw["RREZIK_VONESE"] == 0).sum()
    print(
        f"[RREZIK] të etiketuar={n_labeled} | klasa_1={n_pos} | klasa_0={n_neg} "
        f"| të përjashtuar={raw['RREZIK_VONESE'].isna().sum()}"
    )

    reg_raw = build_regression_frame(raw)
    clf_raw = build_classification_frame(raw)

    reg_proc, reg_pipe, reg_meta = fit_transform_split(
        reg_raw, target_col="MES_PERGJ", task="regression", stratify=False
    )
    clf_proc, clf_pipe, clf_meta = fit_transform_split(
        clf_raw, target_col="RREZIK_VONESE", task="classification", stratify=True
    )

    export_outputs(reg_proc, clf_proc, reg_pipe, clf_pipe)
    print_summary(reg_raw, clf_raw, reg_proc, clf_proc, reg_meta, clf_meta)


if __name__ == "__main__":
    main()
