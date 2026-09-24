"""
Faza D — Modelimi Parashikues dhe Analiza e Mbijetesës (UAMD).

Ekzekuton:
  1) Regresion për MES_PERGJ
  2) Klasifikim për RREZIK_VONESE
  3) Survival Analysis (Kaplan–Meier + Cox PH)
  4) Explainable AI (SHAP) + Early Warning System (EWS)
  5) Dokument Word akademik me tabela, figura dhe interpretim

Daljet ruhen në: PreProces/ML/
"""

from __future__ import annotations

import warnings
from pathlib import Path
from typing import Any

import joblib
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
import shap
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Inches, Pt
from lifelines import CoxPHFitter, KaplanMeierFitter
from lifelines.statistics import multivariate_logrank_test
from sklearn.ensemble import RandomForestClassifier, RandomForestRegressor
from sklearn.linear_model import Lasso, LogisticRegression, Ridge
from sklearn.metrics import (
    accuracy_score,
    confusion_matrix,
    f1_score,
    mean_absolute_error,
    mean_squared_error,
    precision_score,
    r2_score,
    recall_score,
    roc_auc_score,
    roc_curve,
)
from sklearn.svm import SVC, SVR
from sklearn.tree import DecisionTreeClassifier, DecisionTreeRegressor

warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", category=UserWarning)

# ---------------------------------------------------------------------------
# Paths & config
# ---------------------------------------------------------------------------
ML_DIR = Path(__file__).resolve().parent
PREPROCES_DIR = ML_DIR.parent
PROJECT_DIR = PREPROCES_DIR.parent
FIG_DIR = ML_DIR / "figures"
FIG_DIR.mkdir(parents=True, exist_ok=True)

REG_CSV = PREPROCES_DIR / "dataset_ml_regression.csv"
CLF_CSV = PREPROCES_DIR / "dataset_ml_classification.csv"
RAW_XLSX = PROJECT_DIR / "DescriptiveAnalyse" / "Studentet_Uamd_pastruar_spss.xlsx"
RANDOM_STATE = 42
sns.set_theme(style="whitegrid", context="talk")
plt.rcParams["figure.dpi"] = 120
plt.rcParams["savefig.bbox"] = "tight"

VIT_MAP = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6}

# Veçoritë e papërpunuara që pret pipeline-i i klasifikimit (EWS)
RAW_FEATURE_COLS = [
    "MOSHA",
    "Baze_11",
    "Zgjedhje_11",
    "Baze_12",
    "Zgjedhje_12",
    "MES_VIT1_SEM1",
    "GJINIA",
    "LLOJ_RREGJ",
    "PROFIL_GJIMNAZ",
    "RRETHI",
    "GJIMNAZI",
    "DEGA",
    "FAKULTETI",
]

CLF_PIPELINE_PATH = PREPROCES_DIR / "preprocessing_pipeline_classification.joblib"
EWS_BUNDLE_PATH = ML_DIR / "ews_bundle.joblib"

# Pragjet e Early Warning System
EWS_LOW_MAX = 0.40
EWS_MED_MAX = 0.70


# ===========================================================================
# Utilities
# ===========================================================================
def split_xy(df: pd.DataFrame, target: str) -> tuple[pd.DataFrame, pd.DataFrame, pd.Series, pd.Series]:
    """Ndan train/test sipas kolonës 'set' të krijuar në preprocessing."""
    train = df[df["set"] == "train"].copy()
    test = df[df["set"] == "test"].copy()
    feature_cols = [c for c in df.columns if c not in {target, "set"}]
    return (
        train[feature_cols],
        test[feature_cols],
        train[target],
        test[target],
    )


def save_fig(fig: plt.Figure, name: str) -> Path:
    path = FIG_DIR / f"{name}.png"
    fig.savefig(path, dpi=150)
    plt.close(fig)
    print(f"[FIG] {path.name}")
    return path


# ===========================================================================
# 1) REGRESSION
# ===========================================================================
def run_regression(reg_df: pd.DataFrame) -> dict:
    print("\n" + "=" * 70)
    print("1) REGRESION — MES_PERGJ")
    print("=" * 70)

    X_train, X_test, y_train, y_test = split_xy(reg_df, "MES_PERGJ")

    models = {
        "Ridge": Ridge(alpha=1.0, random_state=RANDOM_STATE),
        "Lasso": Lasso(alpha=0.01, random_state=RANDOM_STATE, max_iter=10000),
        "DecisionTree": DecisionTreeRegressor(
            max_depth=8, min_samples_leaf=20, random_state=RANDOM_STATE
        ),
        "RandomForest": RandomForestRegressor(
            n_estimators=200,
            max_depth=12,
            min_samples_leaf=5,
            random_state=RANDOM_STATE,
            n_jobs=-1,
        ),
        "SVR": SVR(kernel="rbf", C=10.0, epsilon=0.1),
    }

    rows = []
    fitted = {}
    for name, model in models.items():
        model.fit(X_train, y_train)
        pred = model.predict(X_test)
        rmse = float(np.sqrt(mean_squared_error(y_test, pred)))
        mae = float(mean_absolute_error(y_test, pred))
        r2 = float(r2_score(y_test, pred))
        rows.append({"Modeli": name, "RMSE": rmse, "MAE": mae, "R2": r2})
        fitted[name] = model
        print(f"  {name:14s}  RMSE={rmse:.4f}  MAE={mae:.4f}  R2={r2:.4f}")

    metrics = pd.DataFrame(rows).sort_values("R2", ascending=False).reset_index(drop=True)
    best_name = metrics.iloc[0]["Modeli"]
    best_model = fitted[best_name]
    print(f"  >> Modeli më i mirë (R²): {best_name}")

    # Feature importance / |coef|
    feature_names = list(X_train.columns)
    if hasattr(best_model, "feature_importances_"):
        importance = pd.Series(best_model.feature_importances_, index=feature_names)
        imp_title = f"Feature Importance — {best_name}"
    elif hasattr(best_model, "coef_"):
        importance = pd.Series(np.abs(np.ravel(best_model.coef_)), index=feature_names)
        imp_title = f"|Koeficientët| — {best_name}"
    else:
        # SVR etj. pa importance direkte → fallback Random Forest
        rf = fitted["RandomForest"]
        importance = pd.Series(rf.feature_importances_, index=feature_names)
        imp_title = "Feature Importance — RandomForest (fallback)"
        print("  (Modeli fitues pa importance direkte; përdoret RF për interpretim)")

    importance = importance.sort_values(ascending=False)
    top_imp = importance.head(12)

    fig, ax = plt.subplots(figsize=(10, 6))
    top_imp.sort_values().plot(kind="barh", ax=ax, color="#2c7fb8")
    ax.set_title(imp_title)
    ax.set_xlabel("Rëndësia / |coef|")
    fig_imp = save_fig(fig, "regression_feature_importance")

    # Comparison bar chart
    fig, axes = plt.subplots(1, 3, figsize=(14, 4.5))
    for ax, col, color in zip(
        axes, ["RMSE", "MAE", "R2"], ["#e34a33", "#fdbb84", "#31a354"]
    ):
        sns.barplot(data=metrics, x="Modeli", y=col, ax=ax, color=color)
        ax.set_title(col)
        ax.tick_params(axis="x", rotation=30)
    fig_cmp = save_fig(fig, "regression_metrics_comparison")

    metrics.to_csv(ML_DIR / "regression_metrics.csv", index=False)
    importance.to_csv(ML_DIR / "regression_feature_importance.csv", header=["importance"])
    joblib.dump(best_model, ML_DIR / f"best_regression_{best_name}.joblib")

    return {
        "metrics": metrics,
        "best_name": best_name,
        "best_model": best_model,
        "importance": importance,
        "fig_importance": fig_imp,
        "fig_comparison": fig_cmp,
        "n_train": len(X_train),
        "n_test": len(X_test),
        "X_train": X_train,
        "X_test": X_test,
        "y_test": y_test,
        "fitted": fitted,
    }


# ===========================================================================
# 2) CLASSIFICATION
# ===========================================================================
def run_classification(clf_df: pd.DataFrame) -> dict:
    print("\n" + "=" * 70)
    print("2) KLASIFIKIM — RREZIK_VONESE")
    print("=" * 70)

    X_train, X_test, y_train, y_test = split_xy(clf_df, "RREZIK_VONESE")

    models = {
        "LogisticRegression": LogisticRegression(
            max_iter=2000,
            class_weight="balanced",
            random_state=RANDOM_STATE,
        ),
        "DecisionTree": DecisionTreeClassifier(
            max_depth=8,
            min_samples_leaf=15,
            class_weight="balanced",
            random_state=RANDOM_STATE,
        ),
        "RandomForest": RandomForestClassifier(
            n_estimators=250,
            max_depth=12,
            min_samples_leaf=5,
            class_weight="balanced",
            random_state=RANDOM_STATE,
            n_jobs=-1,
        ),
        "SVC": SVC(
            kernel="rbf",
            C=1.0,
            probability=True,
            class_weight="balanced",
            random_state=RANDOM_STATE,
        ),
    }

    rows = []
    fitted = {}
    proba_store = {}
    for name, model in models.items():
        model.fit(X_train, y_train)
        pred = model.predict(X_test)
        if hasattr(model, "predict_proba"):
            proba = model.predict_proba(X_test)[:, 1]
        else:
            proba = model.decision_function(X_test)
        auc = float(roc_auc_score(y_test, proba))
        rows.append(
            {
                "Modeli": name,
                "Accuracy": float(accuracy_score(y_test, pred)),
                "Precision": float(precision_score(y_test, pred, zero_division=0)),
                "Recall": float(recall_score(y_test, pred, zero_division=0)),
                "F1": float(f1_score(y_test, pred, zero_division=0)),
                "ROC_AUC": auc,
            }
        )
        fitted[name] = model
        proba_store[name] = (y_test, proba, pred)
        print(
            f"  {name:20s} Acc={rows[-1]['Accuracy']:.3f}  "
            f"F1={rows[-1]['F1']:.3f}  AUC={auc:.3f}"
        )

    metrics = pd.DataFrame(rows).sort_values("ROC_AUC", ascending=False).reset_index(drop=True)
    best_name = metrics.iloc[0]["Modeli"]
    print(f"  >> Modeli më i mirë (ROC-AUC): {best_name}")

    # ROC curves — all models
    fig, ax = plt.subplots(figsize=(8, 6))
    for name, (y_true, proba, _) in proba_store.items():
        fpr, tpr, _ = roc_curve(y_true, proba)
        auc = roc_auc_score(y_true, proba)
        ax.plot(fpr, tpr, lw=2, label=f"{name} (AUC={auc:.3f})")
    ax.plot([0, 1], [0, 1], "k--", lw=1, label="Random")
    ax.set_xlabel("False Positive Rate")
    ax.set_ylabel("True Positive Rate")
    ax.set_title("ROC Curves — RREZIK_VONESE")
    ax.legend(loc="lower right", fontsize=9)
    fig_roc = save_fig(fig, "classification_roc_curves")

    # Confusion matrix — best model
    _, _, best_pred = proba_store[best_name]
    cm = confusion_matrix(y_test, best_pred)
    fig, ax = plt.subplots(figsize=(6, 5))
    sns.heatmap(
        cm,
        annot=True,
        fmt="d",
        cmap="Blues",
        ax=ax,
        xticklabels=["Pa rrezik (0)", "Rrezik (1)"],
        yticklabels=["Pa rrezik (0)", "Rrezik (1)"],
    )
    ax.set_xlabel("Parashikuar")
    ax.set_ylabel("Aktual")
    ax.set_title(f"Confusion Matrix — {best_name}")
    fig_cm = save_fig(fig, "classification_confusion_matrix")

    metrics.to_csv(ML_DIR / "classification_metrics.csv", index=False)
    joblib.dump(fitted[best_name], ML_DIR / f"best_classification_{best_name}.joblib")

    return {
        "metrics": metrics,
        "best_name": best_name,
        "best_model": fitted[best_name],
        "fig_roc": fig_roc,
        "fig_cm": fig_cm,
        "confusion": cm,
        "n_train": len(X_train),
        "n_test": len(X_test),
        "class_dist": clf_df["RREZIK_VONESE"].value_counts().to_dict(),
        "X_train": X_train,
        "X_test": X_test,
        "y_test": y_test,
        "proba_test": proba_store[best_name][1],
        "fitted": fitted,
    }


# ===========================================================================
# 3) EXPLAINABLE AI — SHAP
# ===========================================================================
def _tree_model_for_shap(result: dict, task: str):
    """Zgjedh model tree-based për TreeExplainer (RF / DT)."""
    fitted = result["fitted"]
    preferred = result["best_name"]
    if preferred in {"RandomForest", "DecisionTree"}:
        return preferred, fitted[preferred]
    # SVR / linear / SVC → fallback RandomForest
    print(f"  [SHAP-{task}] {preferred} nuk është tree-based; përdoret RandomForest.")
    return "RandomForest", fitted["RandomForest"]


def _shap_values_positive_class(explainer, X: pd.DataFrame) -> np.ndarray:
    """Kthen SHAP values 2D; për klasifikim merr klasën pozitive (rrezik=1)."""
    values = explainer.shap_values(X)
    if isinstance(values, list):
        return np.asarray(values[1])
    arr = np.asarray(values)
    if arr.ndim == 3:
        # (n_samples, n_features, n_classes)
        return arr[:, :, 1]
    return arr


def run_shap_analysis(reg: dict, clf: dict) -> dict:
    """
    SHAP summary për MES_PERGJ dhe RREZIK_VONESE + waterfall për një student në rrezik.
    """
    print("\n" + "=" * 70)
    print("3) EXPLAINABLE AI — SHAP")
    print("=" * 70)

    # --- Regression SHAP ---
    reg_name, reg_model = _tree_model_for_shap(reg, "REG")
    X_reg = reg["X_train"].sample(n=min(400, len(reg["X_train"])), random_state=RANDOM_STATE)
    reg_explainer = shap.TreeExplainer(reg_model)
    shap_reg = _shap_values_positive_class(reg_explainer, X_reg)
    # for regressor, shap_values is already 2D — helper still works

    # --- Classification SHAP ---
    clf_name, clf_model = _tree_model_for_shap(clf, "CLF")
    X_clf = clf["X_train"].sample(n=min(400, len(clf["X_train"])), random_state=RANDOM_STATE)
    clf_explainer = shap.TreeExplainer(clf_model)
    shap_clf = _shap_values_positive_class(clf_explainer, X_clf)

    # Summary plots (SHAP krijon figuren e vet — ruajmë të ndara dhe një të kombinuar)
    plt.figure(figsize=(9, 7))
    shap.summary_plot(shap_reg, X_reg, show=False, max_display=12)
    plt.title(f"SHAP — MES_PERGJ ({reg_name})")
    fig_reg_path = FIG_DIR / "shap_summary_plot_regression.png"
    plt.savefig(fig_reg_path, dpi=150, bbox_inches="tight")
    plt.close("all")
    print(f"[FIG] {fig_reg_path.name}")

    plt.figure(figsize=(9, 7))
    shap.summary_plot(shap_clf, X_clf, show=False, max_display=12)
    plt.title(f"SHAP — RREZIK_VONESE ({clf_name})")
    fig_clf_path = FIG_DIR / "shap_summary_plot_classification.png"
    plt.savefig(fig_clf_path, dpi=150, bbox_inches="tight")
    plt.close("all")
    print(f"[FIG] {fig_clf_path.name}")

    # Kombinim në shap_summary_plot.png
    img_reg = plt.imread(fig_reg_path)
    img_clf = plt.imread(fig_clf_path)
    fig, axes = plt.subplots(1, 2, figsize=(16, 7))
    axes[0].imshow(img_reg)
    axes[0].axis("off")
    axes[0].set_title("MES_PERGJ")
    axes[1].imshow(img_clf)
    axes[1].axis("off")
    axes[1].set_title("RREZIK_VONESE")
    fig.tight_layout()
    fig_summary = FIG_DIR / "shap_summary_plot.png"
    fig.savefig(fig_summary, dpi=150, bbox_inches="tight")
    plt.close("all")
    print(f"[FIG] {fig_summary.name}")

    # --- Waterfall: një student test në rrezik me probabilitet maksimal ---
    y_test = clf["y_test"].reset_index(drop=True)
    X_test = clf["X_test"].reset_index(drop=True)
    proba = np.asarray(clf["proba_test"])
    risk_idx = np.where(y_test.to_numpy() == 1)[0]
    if len(risk_idx) == 0:
        risk_idx = np.arange(len(X_test))
    best_i = int(risk_idx[np.argmax(proba[risk_idx])])
    x_row = X_test.iloc[[best_i]]
    student_proba = float(proba[best_i])

    # Explanation object for waterfall (SHAP >= 0.40)
    try:
        explanation = clf_explainer(x_row)
        # Binary RF: Explanation may have values shape (1, n_features, 2)
        if getattr(explanation, "values", None) is not None and explanation.values.ndim == 3:
            exp_row = explanation[0, :, 1]
        else:
            exp_row = explanation[0]
        plt.figure(figsize=(10, 6))
        shap.plots.waterfall(exp_row, show=False, max_display=12)
        fig_water = FIG_DIR / "shap_waterfall_plot.png"
        plt.savefig(fig_water, dpi=150, bbox_inches="tight")
        plt.close("all")
    except Exception as exc:
        # Fallback: force plot / bar for single instance
        print(f"  [SHAP] waterfall API dështoi ({exc}); përdoret bar plot.")
        sv_row = _shap_values_positive_class(clf_explainer, x_row)[0]
        order = np.argsort(np.abs(sv_row))[-12:]
        fig, ax = plt.subplots(figsize=(9, 6))
        ax.barh(
            [X_test.columns[i] for i in order],
            sv_row[order],
            color=["#d73027" if v > 0 else "#1a9850" for v in sv_row[order]],
        )
        ax.axvline(0, color="black", lw=0.8)
        ax.set_title(f"SHAP lokal — student në rrezik (p={student_proba:.1%})")
        ax.set_xlabel("SHAP value (ndikimi në P(rrezik))")
        fig_water = save_fig(fig, "shap_waterfall_plot")

    print(f"[FIG] shap_waterfall_plot.png  (student idx={best_i}, p={student_proba:.3f})")

    # Top global drivers (mean |SHAP|)
    mean_abs_reg = pd.Series(np.abs(shap_reg).mean(axis=0), index=X_reg.columns).sort_values(
        ascending=False
    )
    mean_abs_clf = pd.Series(np.abs(shap_clf).mean(axis=0), index=X_clf.columns).sort_values(
        ascending=False
    )
    mean_abs_reg.to_csv(ML_DIR / "shap_mean_abs_regression.csv", header=["mean_abs_shap"])
    mean_abs_clf.to_csv(ML_DIR / "shap_mean_abs_classification.csv", header=["mean_abs_shap"])

    return {
        "fig_summary": fig_summary,
        "fig_waterfall": Path(fig_water) if not isinstance(fig_water, Path) else fig_water,
        "reg_model_name": reg_name,
        "clf_model_name": clf_name,
        "student_index": best_i,
        "student_proba": student_proba,
        "student_features": x_row.iloc[0].to_dict(),
        "top_reg": mean_abs_reg.head(5),
        "top_clf": mean_abs_clf.head(5),
        "clf_explainer": clf_explainer,
        "clf_model": clf_model,
    }


# ===========================================================================
# 4) EARLY WARNING SYSTEM (EWS)
# ===========================================================================
def _risk_category(prob: float) -> str:
    if prob < EWS_LOW_MAX:
        return "I Ultë"
    if prob < EWS_MED_MAX:
        return "I Mesëm"
    return "I Lartë"


def _personalized_recommendation(input_row: pd.Series, prob: float, category: str) -> str:
    """Rekomandim akademik i thjeshtë, i bazuar në notat e vitit 1 dhe nivelin e rrezikut."""
    tips: list[str] = []
    mes = input_row.get("MES_VIT1_SEM1", np.nan)
    b11 = input_row.get("Baze_11", np.nan)
    b12 = input_row.get("Baze_12", np.nan)

    if pd.notna(mes) and mes < 6.0:
        tips.append(
            "Studenti kërkon mbështetje në lëndët bazë sepse notat/mesatarja e semestrit 1 janë kritike"
        )
    elif pd.notna(mes) and mes < 7.0:
        tips.append(
            "Rekomandohet tutoring suplementar për të stabilizuar mesataren e vitit të parë"
        )

    if pd.notna(b11) and b11 < 6.0:
        tips.append("Prioritet: përmirësimi i performancës në lëndët bazë të semestrit 1 (Baze_11)")
    if pd.notna(b12) and b12 < 6.0:
        tips.append("Monitorim i ngushtë i lëndëve bazë të semestrit 2 të vitit 1 (Baze_12)")

    if category == "I Lartë":
        tips.append(
            "Aktivizo plan individual mentoring + takim me koordinatorin e programit brenda 2 javësh"
        )
    elif category == "I Mesëm":
        tips.append(
            "Cakto check-in periodik (çdo muaj) dhe ndiq progresin e notave në semestër"
        )
    else:
        tips.append(
            "Mbaj ritmin aktual; inkurajo pjesëmarrjen në aktivitete akademike shtesë"
        )

    if not tips:
        tips.append(
            f"Probabiliteti i vonesës është {prob:.0%} ({category}); vazhdo monitorimin standard"
        )
    return " | ".join(tips[:3])


def parashiko_rrezikun_studentit(
    input_data: dict[str, Any] | pd.Series | pd.DataFrame,
    pipeline=None,
    model=None,
) -> dict[str, Any]:
    """
    Early Warning System për një student të vitit të parë.

    Parameters
    ----------
    input_data : dict | Series | DataFrame
        Parametrat e papërpunuara (MOSHA, GJINIA, Baze_11, MES_VIT1_SEM1, …).
    pipeline, model : opsionale
        Nëse mungojnë, ngarkohen nga joblib (ews_bundle / preprocessing + best RF).

    Returns
    -------
    dict me:
      - probabilitet_vonese_pct
      - kategoria_rrezikut  (I Ultë / I Mesëm / I Lartë)
      - rekomandim
      - probabilitet_vonese (0–1)
    """
    if pipeline is None or model is None:
        if EWS_BUNDLE_PATH.exists():
            bundle = joblib.load(EWS_BUNDLE_PATH)
            pipeline = pipeline or bundle["pipeline"]
            model = model or bundle["model"]
        else:
            if pipeline is None:
                pipeline = joblib.load(CLF_PIPELINE_PATH)
            if model is None:
                # gjej modelin më të mirë të ruajtur
                candidates = sorted(ML_DIR.glob("best_classification_*.joblib"))
                if not candidates:
                    raise FileNotFoundError("Nuk u gjet modeli i klasifikimit për EWS.")
                model = joblib.load(candidates[0])

    if isinstance(input_data, dict):
        row = pd.DataFrame([input_data])
    elif isinstance(input_data, pd.Series):
        row = input_data.to_frame().T
    else:
        row = input_data.copy()

    missing = [c for c in RAW_FEATURE_COLS if c not in row.columns]
    if missing:
        raise KeyError(f"Mungojnë fushat e detyrueshme për EWS: {missing}")

    X_raw = row[RAW_FEATURE_COLS]
    X_proc = pipeline.transform(X_raw)
    # siguro DataFrame me emra kolonash nëse është array
    if not isinstance(X_proc, pd.DataFrame):
        try:
            names = pipeline.named_steps["preprocess"].get_feature_names_out()
        except Exception:
            names = [f"f{i}" for i in range(np.asarray(X_proc).shape[1])]
        X_proc = pd.DataFrame(X_proc, columns=names)

    proba = float(model.predict_proba(X_proc)[0, 1])
    category = _risk_category(proba)
    recommendation = _personalized_recommendation(X_raw.iloc[0], proba, category)

    return {
        "probabilitet_vonese": proba,
        "probabilitet_vonese_pct": round(proba * 100, 2),
        "kategoria_rrezikut": category,
        "rekomandim": recommendation,
    }


def run_early_warning_demo(clf: dict, raw: pd.DataFrame | None = None) -> dict:
    """Ruaj bundle EWS dhe ekzekuto një shembull demonstrues."""
    print("\n" + "=" * 70)
    print("4) EARLY WARNING SYSTEM")
    print("=" * 70)

    if not CLF_PIPELINE_PATH.exists():
        raise FileNotFoundError(
            f"Mungon pipeline-i i preprocessimit: {CLF_PIPELINE_PATH}"
        )
    pipeline = joblib.load(CLF_PIPELINE_PATH)
    model = clf["best_model"]
    joblib.dump(
        {
            "pipeline": pipeline,
            "model": model,
            "feature_cols": RAW_FEATURE_COLS,
            "thresholds": {"low_max": EWS_LOW_MAX, "med_max": EWS_MED_MAX},
        },
        EWS_BUNDLE_PATH,
    )
    print(f"[SAVE] {EWS_BUNDLE_PATH.name}")

    # Shembull real nga dataset (preferohet) ose fallback sintetik
    demo_input = None
    if raw is not None:
        cand = raw.dropna(subset=["MES_VIT1_SEM1", "Baze_11"]).copy()
        low = cand[cand["MES_VIT1_SEM1"] <= 6.0]
        if len(low):
            r = low.sample(1, random_state=RANDOM_STATE).iloc[0]
            demo_input = {c: r[c] for c in RAW_FEATURE_COLS if c in r.index}

    if demo_input is None:
        demo_input = {
            "MOSHA": 19,
            "Baze_11": 5.0,
            "Zgjedhje_11": 5.5,
            "Baze_12": 5.2,
            "Zgjedhje_12": 5.0,
            "MES_VIT1_SEM1": 5.1,
            "GJINIA": "MASHKULL",
            "LLOJ_RREGJ": "PROGRAM I PARE",
            "PROFIL_GJIMNAZ": "I PERGJITHSHEM",
            "RRETHI": "DURRËS",
            "GJIMNAZI": cand["GJIMNAZI"].mode().iloc[0] if raw is not None else "Unknown",
            "DEGA": "SHKENCA KOMPJUTERIKE",
            "FAKULTETI": "FTI",
        }

    result = parashiko_rrezikun_studentit(demo_input, pipeline=pipeline, model=model)
    print(
        f"  Demo EWS → p={result['probabilitet_vonese_pct']}% | "
        f"{result['kategoria_rrezikut']}"
    )
    print(f"  Rekomandim: {result['rekomandim']}")
    return {"demo_input": demo_input, "demo_result": result}


# ===========================================================================
# 5) SURVIVAL ANALYSIS
# ===========================================================================
def prepare_survival_frame(raw: pd.DataFrame) -> pd.DataFrame:
    """
    Opsioni 1 — kohë deri në diplomim:
      event = 1 nëse diplomuar (KOHA_DIPLOM ose DIPLOMUAR)
      event = 0 censuruar (ende aktiv)
      duration = KOHA_DIPLOM ose viti i studimit (I..VI)
    """
    df = raw.copy()

    is_graduated = df["KOHA_DIPLOM"].notna() | (
        df["VIT_AKADEMIK"].astype(str).str.strip().str.upper() == "DIPLOMUAR"
    )

    vit_num = (
        df["VIT_AKADEMIK"]
        .astype(str)
        .str.strip()
        .str.upper()
        .map(VIT_MAP)
        .astype(float)
    )

    duration = df["KOHA_DIPLOM"].copy()
    # për aktivët: koha e vëzhguar = viti aktual i studimit
    duration = duration.fillna(vit_num)
    # DIPLOMUAR pa KOHA_DIPLOM: përdor medianën e kohës së diplomimit ose 3
    med_grad = df.loc[df["KOHA_DIPLOM"].notna(), "KOHA_DIPLOM"].median()
    if pd.isna(med_grad):
        med_grad = 3.0
    diploma_mask = (
        df["VIT_AKADEMIK"].astype(str).str.strip().str.upper() == "DIPLOMUAR"
    ) & df["KOHA_DIPLOM"].isna()
    duration.loc[diploma_mask] = med_grad

    out = pd.DataFrame(
        {
            "duration": duration,
            "event": is_graduated.astype(int),
            "MOSHA": df["MOSHA"],
            "MES_VIT1_SEM1": df["MES_VIT1_SEM1"],
            "GJINIA": df["GJINIA"],
            "PROFIL_GJIMNAZ": df["PROFIL_GJIMNAZ"].fillna("Unknown"),
            "FAKULTETI": df["FAKULTETI"],
            "NIVELI_STUDIMIT": df["NIVELI_STUDIMIT"],
            "DEGA": df["DEGA"],
        }
    )
    out = out.dropna(subset=["duration"])
    out = out[out["duration"] > 0].reset_index(drop=True)
    print(
        f"[SURV] n={len(out)} | events={out['event'].sum()} "
        f"| censored={(out['event'] == 0).sum()}"
    )
    return out


def run_survival(raw: pd.DataFrame) -> dict:
    print("\n" + "=" * 70)
    print("5) SURVIVAL ANALYSIS — Kaplan–Meier & Cox PH")
    print("=" * 70)

    surv = prepare_survival_frame(raw)

    # --- Kaplan–Meier overall + by PROFIL / FAKULTETI ---
    kmf = KaplanMeierFitter()
    fig, axes = plt.subplots(1, 2, figsize=(14, 5.5))

    # By PROFIL_GJIMNAZ
    ax = axes[0]
    for profil, grp in surv.groupby("PROFIL_GJIMNAZ"):
        if len(grp) < 30:
            continue
        kmf.fit(grp["duration"], grp["event"], label=str(profil))
        kmf.plot_survival_function(ax=ax, ci_show=False)
    ax.set_title("Kaplan–Meier sipas PROFIL_GJIMNAZ")
    ax.set_xlabel("Kohë (vite)")
    ax.set_ylabel("S(t) — probabiliteti i mos-diplomimit ende")
    ax.legend(fontsize=8)

    # By FAKULTETI
    ax = axes[1]
    for fak, grp in surv.groupby("FAKULTETI"):
        kmf.fit(grp["duration"], grp["event"], label=str(fak))
        kmf.plot_survival_function(ax=ax, ci_show=False)
    ax.set_title("Kaplan–Meier sipas FAKULTETI")
    ax.set_xlabel("Kohë (vite)")
    ax.set_ylabel("S(t)")
    ax.legend(fontsize=9)

    fig_km = save_fig(fig, "survival_kaplan_meier")

    # Log-rank (FAKULTETI)
    try:
        lr = multivariate_logrank_test(
            surv["duration"], surv["FAKULTETI"], surv["event"]
        )
        logrank_p = float(lr.p_value)
        print(f"  Log-rank (FAKULTETI) p-value = {logrank_p:.4g}")
    except Exception as exc:
        logrank_p = np.nan
        print(f"  Log-rank dështoi: {exc}")

    # --- Cox PH ---
    cox_df = surv.copy()
    cox_df["MES_VIT1_SEM1"] = cox_df["MES_VIT1_SEM1"].fillna(
        cox_df["MES_VIT1_SEM1"].median()
    )
    cox_df["MOSHA"] = cox_df["MOSHA"].fillna(cox_df["MOSHA"].median())
    # encode kategorike (drop_first)
    cox_model_df = pd.get_dummies(
        cox_df[
            [
                "duration",
                "event",
                "MOSHA",
                "MES_VIT1_SEM1",
                "GJINIA",
                "PROFIL_GJIMNAZ",
                "FAKULTETI",
                "NIVELI_STUDIMIT",
            ]
        ],
        columns=["GJINIA", "PROFIL_GJIMNAZ", "FAKULTETI", "NIVELI_STUDIMIT"],
        drop_first=True,
    )
    # sigurohu që janë numerike float
    cox_model_df = cox_model_df.astype(float)

    cph = CoxPHFitter(penalizer=0.1)
    cph.fit(cox_model_df, duration_col="duration", event_col="event")
    summary = cph.summary.copy()
    summary.index.name = "Faktori"
    summary = summary.reset_index()
    summary["HR"] = np.exp(summary["coef"])
    hr_table = pd.DataFrame(
        {
            "Faktori": summary["Faktori"],
            "coef": summary["coef"],
            "HR": summary["HR"],
            "p_value": summary["p"],
            "HR_low_95": summary["exp(coef) lower 95%"],
            "HR_high_95": summary["exp(coef) upper 95%"],
        }
    ).sort_values("p_value").reset_index(drop=True)
    print(hr_table.to_string(index=False))

    # Forest-style HR plot
    plot_df = hr_table.sort_values("HR")
    fig, ax = plt.subplots(figsize=(9, 6))
    y_pos = np.arange(len(plot_df))
    ax.errorbar(
        plot_df["HR"],
        y_pos,
        xerr=[
            plot_df["HR"] - plot_df["HR_low_95"],
            plot_df["HR_high_95"] - plot_df["HR"],
        ],
        fmt="o",
        color="#2c7fb8",
        ecolor="#7fcdbb",
        capsize=3,
    )
    ax.axvline(1.0, color="red", linestyle="--", lw=1)
    ax.set_yticks(y_pos)
    ax.set_yticklabels(plot_df["Faktori"], fontsize=8)
    ax.set_xlabel("Hazard Ratio (HR)")
    ax.set_title("Cox PH — Hazard Ratios (95% CI)")
    fig_hr = save_fig(fig, "survival_cox_hazard_ratios")

    hr_table.to_csv(ML_DIR / "survival_cox_hazard_ratios.csv", index=False)
    surv.to_csv(ML_DIR / "survival_prepared_data.csv", index=False)

    interpretations = []
    for _, row in hr_table.iterrows():
        factor = row["Faktori"]
        hr = row["HR"]
        p = row["p_value"]
        sig = "statistikisht i rëndësishëm" if p < 0.05 else "jo i rëndësishëm statistikisht"
        if hr > 1:
            meaning = (
                f"HR={hr:.3f} (>1): '{factor}' shoqërohet me rrezik më të lartë "
                f"të ngjarjes së diplomimit në kohë më të shkurtër (hazard më i lartë); {sig} (p={p:.4g})."
            )
        elif hr < 1:
            meaning = (
                f"HR={hr:.3f} (<1): '{factor}' vepron si faktor mbrojtës / vonon "
                f"ngjarjen e diplomimit (hazard më i ulët); {sig} (p={p:.4g})."
            )
        else:
            meaning = f"HR≈1 për '{factor}': efekt neutral; {sig} (p={p:.4g})."
        interpretations.append(meaning)

    return {
        "surv": surv,
        "hr_table": hr_table,
        "interpretations": interpretations,
        "fig_km": fig_km,
        "fig_hr": fig_hr,
        "logrank_p": logrank_p,
        "cph": cph,
        "concordance": float(cph.concordance_index_),
    }


# ===========================================================================
# Main
# ===========================================================================
def main() -> None:
    print("Faza D — Modelimi & Survival Analysis + SHAP/EWS")
    print(f"ML_DIR: {ML_DIR}")

    if not REG_CSV.exists() or not CLF_CSV.exists():
        raise FileNotFoundError(
            "Mungojnë CSV-të e preprocessing. Ekzekuto së pari preprocess_ml_dual.py"
        )
    if not RAW_XLSX.exists():
        raise FileNotFoundError(f"Mungon Excel-i burimor: {RAW_XLSX}")

    reg_df = pd.read_csv(REG_CSV)
    clf_df = pd.read_csv(CLF_CSV)
    raw = pd.read_excel(RAW_XLSX)

    reg_res = run_regression(reg_df)
    clf_res = run_classification(clf_df)
    shap_res = run_shap_analysis(reg_res, clf_res)
    ews_res = run_early_warning_demo(clf_res, raw=raw)
    surv_res = run_survival(raw)

    print("\n" + "=" * 70)
    print("PËRFUNDUAR")
    print(f"  Figurat:   {FIG_DIR}")
    print(f"  EWS demo:  {ews_res['demo_result']}")
    print("=" * 70)


if __name__ == "__main__":
    main()