"""
Sistemi i Paralajmërimit të Hershëm për Studentët e UAMD-së
Aplikacion Streamlit — panel për sekretarinë e departamentit.

Ekzekutimi:
    streamlit run app.py
(nga folderi PreProces/ML, me venv të projektit aktiv)
"""

from __future__ import annotations

from pathlib import Path

import joblib
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ---------------------------------------------------------------------------
# Paths & joblib discovery (CWD + PreProces + models + pranë app.py)
# ---------------------------------------------------------------------------
APP_DIR = Path(__file__).resolve().parent
PREPROCES_DIR = APP_DIR.parent
PROJECT_DIR = PREPROCES_DIR.parent
CWD = Path.cwd().resolve()

# Emra të mundshëm (case-insensitive) për çdo artefakt
EWS_NAME_CANDIDATES = [
    "ews_bundle.joblib",
    "model_pipeline.joblib",
    "best_classification_RandomForest.joblib",
    "best_classification_randomforest.joblib",
]
REG_MODEL_NAME_CANDIDATES = [
    "best_regression_RandomForest.joblib",
    "best_regression_randomforest.joblib",
    "model_rf.joblib",
    "model_regression.joblib",
]
REG_PIPE_NAME_CANDIDATES = [
    "preprocessing_pipeline_regression.joblib",
    "preprocess_pipeline_regression.joblib",
    "pipeline_regression.joblib",
]

CLF_METRICS_NAMES = ["classification_metrics.csv"]
REG_METRICS_NAMES = ["regression_metrics.csv"]
FEAT_IMP_NAMES = ["regression_feature_importance.csv"]


def _search_roots() -> list[Path]:
    """Dosjet ku kërkohen modelet .joblib / metrika."""
    roots = [
        CWD,
        CWD / "PreProces",
        CWD / "preproces",
        CWD / "models",
        CWD / "Models",
        CWD / "ML",
        CWD / "PreProces" / "ML",
        APP_DIR,
        APP_DIR / "models",
        PREPROCES_DIR,
        PREPROCES_DIR / "ML",
        PREPROCES_DIR / "models",
        PROJECT_DIR,
        PROJECT_DIR / "PreProces",
        PROJECT_DIR / "PreProces" / "ML",
        PROJECT_DIR / "models",
    ]
    # unikë + ekzistues
    seen: set[str] = set()
    out: list[Path] = []
    for r in roots:
        try:
            key = str(r.resolve()).lower()
        except Exception:  # noqa: BLE001
            continue
        if key in seen:
            continue
        seen.add(key)
        if r.exists() and r.is_dir():
            out.append(r.resolve())
    return out


def _index_files(roots: list[Path], extensions: tuple[str, ...]) -> dict[str, Path]:
    """
    Indeks case-insensitive: emri.lower() -> Path.
    Skanon root-et dhe një nivel nëndosjesh (përveç .venv).
    """
    index: dict[str, Path] = {}
    ext_set = {e.lower() for e in extensions}
    for root in roots:
        candidates = [root]
        try:
            for child in root.iterdir():
                if child.is_dir() and child.name.lower() not in {".venv", "venv", "__pycache__", ".git"}:
                    candidates.append(child)
        except PermissionError:
            continue
        for folder in candidates:
            try:
                for f in folder.iterdir():
                    if not f.is_file():
                        continue
                    if f.suffix.lower() not in ext_set:
                        continue
                    key = f.name.lower()
                    # prefero hit-in e parë; mos mbishkruaj
                    index.setdefault(key, f.resolve())
            except PermissionError:
                continue
    return index


def find_artifact(
    name_candidates: list[str],
    file_index: dict[str, Path],
) -> Path | None:
    """Gjej skedarin duke injoruar madhësinë e shkronjave."""
    for name in name_candidates:
        hit = file_index.get(name.lower())
        if hit is not None and hit.exists():
            return hit
    return None


def diagnose_missing(label: str, name_candidates: list[str], found: Path | None) -> str:
    roots = _search_roots()
    lines = [
        f"**{label}**",
        f"- CWD (Current Working Directory): `{CWD}`",
        f"- Dosja e `app.py`: `{APP_DIR}`",
        "- Dosjet e kërkuara:",
    ]
    for r in roots:
        lines.append(f"  - `{r}`")
    if found is not None:
        lines.append(f"- ✅ U gjet: `{found}`")
    else:
        lines.append("- ❌ Nuk u gjet. Emrat e pritur (case-insensitive):")
        for n in name_candidates:
            lines.append(f"  - `{n}`")
        lines.append(
            "- Vendosi një nga këta skedarë në root, `PreProces/`, `PreProces/ML/` ose `models/`."
        )
    return "\n".join(lines)


@st.cache_resource
def discover_joblib_paths() -> dict:
    """Zbulon rrugët e modeleve; cache-ohet për sesionin Streamlit."""
    roots = _search_roots()
    joblib_index = _index_files(roots, (".joblib",))
    csv_index = _index_files(roots, (".csv",))
    xlsx_index = _index_files(
        roots
        + ([PROJECT_DIR / "DescriptiveAnalyse"] if (PROJECT_DIR / "DescriptiveAnalyse").exists() else []),
        (".xlsx",),
    )

    ews_path = find_artifact(EWS_NAME_CANDIDATES, joblib_index)
    reg_model_path = find_artifact(REG_MODEL_NAME_CANDIDATES, joblib_index)
    reg_pipe_path = find_artifact(REG_PIPE_NAME_CANDIDATES, joblib_index)

    # Nëse ews_bundle mungon por kemi klasifikues + pipeline klasifikimi, ndërto bundle virtual
    clf_model = find_artifact(
        [
            "best_classification_RandomForest.joblib",
            "best_classification_randomforest.joblib",
            "model_clf.joblib",
        ],
        joblib_index,
    )
    clf_pipe = find_artifact(
        [
            "preprocessing_pipeline_classification.joblib",
            "pipeline_classification.joblib",
        ],
        joblib_index,
    )

    return {
        "roots": roots,
        "cwd": CWD,
        "app_dir": APP_DIR,
        "ews_path": ews_path,
        "reg_model_path": reg_model_path,
        "reg_pipe_path": reg_pipe_path,
        "clf_model_path": clf_model,
        "clf_pipe_path": clf_pipe,
        "joblib_index": joblib_index,
        "clf_metrics": find_artifact(CLF_METRICS_NAMES, csv_index),
        "reg_metrics": find_artifact(REG_METRICS_NAMES, csv_index),
        "feat_imp": find_artifact(FEAT_IMP_NAMES, csv_index),
        "raw_xlsx": find_artifact(
            ["Studentet_Uamd_pastruar_spss.xlsx", "studentet_uamd_pastruar_spss.xlsx"],
            xlsx_index,
        ),
    }


# Compat aliases (përdoren edhe nga metrika / explorer)
def _paths():
    return discover_joblib_paths()


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

# Pragjet e rrezikut sipas kërkesës së UI
RISK_LOW = 0.30
RISK_MED = 0.60


# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="EWS UAMD — Parashikimi i Suksesit",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main-title { font-size: 1.8rem; font-weight: 700; color: #0B3D5C; margin-bottom: 0.2rem; }
    .sub-title { color: #5A6A7A; margin-bottom: 1.2rem; }
    .alert-high { background:#FDEDEC; border-left:6px solid #C0392B; padding:0.9rem 1rem; border-radius:6px; }
    .alert-med  { background:#FEF9E7; border-left:6px solid #F1C40F; padding:0.9rem 1rem; border-radius:6px; }
    .alert-low  { background:#E8F8F5; border-left:6px solid #27AE60; padding:0.9rem 1rem; border-radius:6px; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Cached loaders
# ---------------------------------------------------------------------------
@st.cache_resource
def load_ews_bundle():
    info = discover_joblib_paths()
    path = info["ews_path"]

    # Fallback: ndërto bundle nga model + pipeline klasifikimi
    if path is None and info["clf_model_path"] and info["clf_pipe_path"]:
        try:
            bundle = {
                "pipeline": joblib.load(info["clf_pipe_path"]),
                "model": joblib.load(info["clf_model_path"]),
                "feature_cols": RAW_FEATURE_COLS,
                "_source": "composed",
                "_model_path": str(info["clf_model_path"]),
                "_pipe_path": str(info["clf_pipe_path"]),
            }
            return bundle, None, info
        except Exception as exc:  # noqa: BLE001
            return None, f"Gabim gjatë ndërtimit të EWS nga pjesët: {exc}", info

    if path is None:
        msg = diagnose_missing("EWS / klasifikimi (.joblib)", EWS_NAME_CANDIDATES, None)
        msg += "\n\nAlternativë: `best_classification_RandomForest.joblib` + `preprocessing_pipeline_classification.joblib`."
        return None, msg, info

    try:
        bundle = joblib.load(path)
        # Nëse është vetëm model (jo dict me pipeline), provo ta kompozosh
        if not isinstance(bundle, dict) or "model" not in bundle or "pipeline" not in bundle:
            if info["clf_pipe_path"] is not None:
                bundle = {
                    "pipeline": joblib.load(info["clf_pipe_path"]),
                    "model": bundle if not isinstance(bundle, dict) else bundle.get("model", bundle),
                    "feature_cols": RAW_FEATURE_COLS,
                    "_source": str(path),
                }
            else:
                return (
                    None,
                    diagnose_missing(
                        "Pipeline klasifikimi (për t'u çiftuar me modelin)",
                        [
                            "preprocessing_pipeline_classification.joblib",
                            "ews_bundle.joblib",
                        ],
                        None,
                    ),
                    info,
                )
        bundle["_source"] = str(path)
        return bundle, None, info
    except Exception as exc:  # noqa: BLE001
        return None, f"Gabim gjatë ngarkimit të `{path}`: {exc}", info


@st.cache_resource
def load_regression_artifacts():
    info = discover_joblib_paths()
    model_path = info["reg_model_path"]
    pipe_path = info["reg_pipe_path"]
    missing_msgs = []
    if model_path is None:
        missing_msgs.append(
            diagnose_missing("Modeli i regresionit", REG_MODEL_NAME_CANDIDATES, None)
        )
    if pipe_path is None:
        missing_msgs.append(
            diagnose_missing("Pipeline i regresionit", REG_PIPE_NAME_CANDIDATES, None)
        )
    if missing_msgs:
        return None, None, "\n\n".join(missing_msgs), info
    try:
        model = joblib.load(model_path)
        pipe = joblib.load(pipe_path)
        return model, pipe, None, info
    except Exception as exc:  # noqa: BLE001
        return None, None, f"Gabim ngarkimi regresioni: {exc}", info


@st.cache_data
def load_raw_dataframe() -> pd.DataFrame | None:
    info = discover_joblib_paths()
    path = info.get("raw_xlsx")
    # fallback klasik
    classic = PROJECT_DIR / "DescriptiveAnalyse" / "Studentet_Uamd_pastruar_spss.xlsx"
    if path is None and classic.exists():
        path = classic
    if path is None:
        return None
    try:
        return pd.read_excel(path)
    except Exception:  # noqa: BLE001
        return None


@st.cache_data
def load_choice_lists(raw: pd.DataFrame | None) -> dict:
    defaults = {
        "GJINIA": ["FEMER", "MASHKULL"],
        "PROFIL_GJIMNAZ": ["I PERGJITHSHEM", "PROFESIONAL"],
        "FAKULTETI": ["FB", "FTI", "FSHPJ"],
        "LLOJ_RREGJ": [
            "PROGRAM I PARE",
            "PROGRAM I DYTË",
            "TRANSFERIM STUDIMESH",
            "KUOTË E VEÇANTË",
        ],
        "RRETHI": ["DURRËS", "TIRANË", "DIBËR", "KAVAJË", "LUSHNJË", "BERAT", "KUKËS", "TJETER"],
        "DEGA": [
            "SHKENCA KOMPJUTERIKE",
            "TEKNOLOGJI INFORMACIONI",
            "BANKË FINANCË",
            "FINANCË KONTABILITET",
            "MENAXHIM BIZNESI",
            "SHKENCA POLITIKE",
        ],
        "GJIMNAZI": ["— (tjetër / i panjohur) —"],
    }
    if raw is None:
        return defaults
    out = dict(defaults)
    for col in ["GJINIA", "PROFIL_GJIMNAZ", "FAKULTETI", "LLOJ_RREGJ", "DEGA"]:
        if col in raw.columns:
            vals = sorted(raw[col].dropna().astype(str).unique().tolist())
            if vals:
                out[col] = vals
    if "RRETHI" in raw.columns:
        top = raw["RRETHI"].value_counts().head(25).index.astype(str).tolist()
        out["RRETHI"] = top
    if "GJIMNAZI" in raw.columns:
        top_g = raw["GJIMNAZI"].value_counts().head(40).index.astype(str).tolist()
        out["GJIMNAZI"] = top_g + ["— (tjetër / i panjohur) —"]
    return out


def risk_status(prob: float) -> tuple[str, str, str]:
    """Kthen (kategori, mesazh, css_class)."""
    pct = prob * 100
    if pct < RISK_LOW * 100:
        return (
            "Rrezik i Ulët",
            "Student me performancë të lartë.",
            "alert-low",
        )
    if pct <= RISK_MED * 100:
        return (
            "Rrezik i Mesëm",
            "Rekommandohet këshillim akademik lehtë.",
            "alert-med",
        )
    return (
        "Rrezik i Lartë",
        "ALERT: Studenti kërkon tutorig të menjëhershëm në lëndët bazë.",
        "alert-high",
    )


def gauge_chart(prob_pct: float) -> go.Figure:
    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=prob_pct,
            number={"suffix": "%", "font": {"size": 36}},
            title={"text": "Probabiliteti i Vonesës / Dropout", "font": {"size": 16}},
            gauge={
                "axis": {"range": [0, 100]},
                "bar": {"color": "#2C3E50"},
                "steps": [
                    {"range": [0, 30], "color": "#27AE60"},
                    {"range": [30, 60], "color": "#F1C40F"},
                    {"range": [60, 100], "color": "#E74C3C"},
                ],
                "threshold": {
                    "line": {"color": "#1A252F", "width": 3},
                    "thickness": 0.75,
                    "value": prob_pct,
                },
            },
        )
    )
    fig.update_layout(height=320, margin=dict(l=20, r=20, t=50, b=20))
    return fig


def build_input_dataframe(sidebar_values: dict) -> pd.DataFrame:
    gjimnazi = sidebar_values["GJIMNAZI"]
    if gjimnazi.startswith("—"):
        gjimnazi = "Unknown"
    row = {
        "MOSHA": sidebar_values["MOSHA"],
        "Baze_11": sidebar_values["Baze_11"],
        "Zgjedhje_11": sidebar_values["Zgjedhje_11"],
        "Baze_12": sidebar_values["Baze_12"],
        "Zgjedhje_12": sidebar_values["Zgjedhje_12"],
        "MES_VIT1_SEM1": sidebar_values["MES_VIT1_SEM1"],
        "GJINIA": sidebar_values["GJINIA"],
        "LLOJ_RREGJ": sidebar_values["LLOJ_RREGJ"],
        "PROFIL_GJIMNAZ": sidebar_values["PROFIL_GJIMNAZ"],
        "RRETHI": sidebar_values["RRETHI"],
        "GJIMNAZI": gjimnazi,
        "DEGA": sidebar_values["DEGA"],
        "FAKULTETI": sidebar_values["FAKULTETI"],
    }
    return pd.DataFrame([row])[RAW_FEATURE_COLS]


def predict_all(X_raw: pd.DataFrame, ews_bundle, reg_model, reg_pipe) -> dict:
    """Parashikon MES_PERGJ dhe probabilitetin e vonesës."""
    # Klasifikim
    clf_pipe = ews_bundle["pipeline"]
    clf_model = ews_bundle["model"]
    X_clf = clf_pipe.transform(X_raw)
    if not isinstance(X_clf, pd.DataFrame):
        try:
            names = clf_pipe.named_steps["preprocess"].get_feature_names_out()
        except Exception:  # noqa: BLE001
            names = [f"f{i}" for i in range(np.asarray(X_clf).shape[1])]
        X_clf = pd.DataFrame(X_clf, columns=names)
    proba = float(clf_model.predict_proba(X_clf)[0, 1])

    # Regresion
    X_reg = reg_pipe.transform(X_raw)
    if not isinstance(X_reg, pd.DataFrame):
        try:
            names = reg_pipe.named_steps["preprocess"].get_feature_names_out()
        except Exception:  # noqa: BLE001
            names = [f"f{i}" for i in range(np.asarray(X_reg).shape[1])]
        X_reg = pd.DataFrame(X_reg, columns=names)
    mes_hat = float(reg_model.predict(X_reg)[0])

    return {"mes_pergj": mes_hat, "proba": proba, "proba_pct": proba * 100}


# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.markdown(
    '<div class="main-title">Sistemi i Parashikimit të Suksesit dhe Rrezikut Akademik - UAMD</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="sub-title">Early Warning System & panel kontrolli për sekretarinë e departamentit</div>',
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Sidebar — inputet e studentit
# ---------------------------------------------------------------------------
raw_df = load_raw_dataframe()
choices = load_choice_lists(raw_df)

st.sidebar.header("📋 Të dhënat e studentit")
st.sidebar.caption("Plotëso profilin e vitit të parë për parashikim individual.")

st.sidebar.subheader("Matura")
mes_matura = st.sidebar.slider(
    "Mesatarja e Maturës",
    min_value=5.0,
    max_value=10.0,
    value=7.5,
    step=0.1,
    help="Fushë kontekstuale UI (nuk është kolonë e drejtpërdrejtë në modelin aktual).",
)
profil = st.sidebar.selectbox("Profil i Gjimnazit", choices["PROFIL_GJIMNAZ"])

st.sidebar.subheader("Demografia")
gjinia = st.sidebar.selectbox("Gjinia", choices["GJINIA"])
rrethi = st.sidebar.selectbox("Rrethi", choices["RRETHI"])
fakulteti = st.sidebar.selectbox("Fakulteti", choices["FAKULTETI"])
# Filtro degët sipas fakultetit nëse është e mundur
dega_opts = choices["DEGA"]
if raw_df is not None and {"FAKULTETI", "DEGA"}.issubset(raw_df.columns):
    filtered = (
        raw_df.loc[raw_df["FAKULTETI"] == fakulteti, "DEGA"]
        .dropna()
        .astype(str)
        .value_counts()
        .index.tolist()
    )
    if filtered:
        dega_opts = filtered
dega = st.sidebar.selectbox("Dega", dega_opts)
gjimnazi = st.sidebar.selectbox("Gjimnazi", choices["GJIMNAZI"])
mosha = st.sidebar.number_input("Mosha", min_value=17, max_value=60, value=19, step=1)
lloj_rregj = st.sidebar.selectbox("Lloji i regjistrimit", choices["LLOJ_RREGJ"])

st.sidebar.subheader("Semestri / Viti 1")
mes_vit1 = st.sidebar.slider(
    "Mesatarja e Semestrit 1 (MES_VIT1_SEM1)",
    min_value=5.0,
    max_value=10.0,
    value=7.0,
    step=0.1,
)
baze_11 = st.sidebar.slider("Nota Bazë 11 (Baze_11)", 5.0, 10.0, 7.0, 0.1)
zgjedhje_11 = st.sidebar.slider("Nota Zgjedhje 11 (Zgjedhje_11)", 5.0, 10.0, 7.0, 0.1)
baze_12 = st.sidebar.slider("Nota Bazë 12 (Baze_12)", 5.0, 10.0, 7.0, 0.1)
zgjedhje_12 = st.sidebar.slider("Nota Zgjedhje 12 (Zgjedhje_12)", 5.0, 10.0, 7.0, 0.1)

sidebar_values = {
    "MOSHA": int(mosha),
    "Baze_11": float(baze_11),
    "Zgjedhje_11": float(zgjedhje_11),
    "Baze_12": float(baze_12),
    "Zgjedhje_12": float(zgjedhje_12),
    "MES_VIT1_SEM1": float(mes_vit1),
    "GJINIA": gjinia,
    "LLOJ_RREGJ": lloj_rregj,
    "PROFIL_GJIMNAZ": profil,
    "RRETHI": rrethi,
    "GJIMNAZI": gjimnazi,
    "DEGA": dega,
    "FAKULTETI": fakulteti,
    "MES_MATURA_UI": float(mes_matura),
}

# ---------------------------------------------------------------------------
# Load models (with friendly errors + path diagnostics)
# ---------------------------------------------------------------------------
ews_bundle, ews_err, path_info = load_ews_bundle()
reg_model, reg_pipe, reg_err, _ = load_regression_artifacts()

CLF_METRICS = path_info.get("clf_metrics")
REG_METRICS = path_info.get("reg_metrics")
FEAT_IMP = path_info.get("feat_imp")

with st.sidebar.expander("🔎 Diagnostika e modeleve (.joblib)", expanded=bool(ews_err or reg_err)):
    st.caption(f"CWD: `{path_info['cwd']}`")
    st.caption(f"app.py: `{path_info['app_dir']}`")
    st.write("**Skedarët e gjetur:**")
    st.write(
        {
            "EWS / bundle": str(path_info["ews_path"]) if path_info["ews_path"] else "— mungon —",
            "Reg model": str(path_info["reg_model_path"]) if path_info["reg_model_path"] else "— mungon —",
            "Reg pipeline": str(path_info["reg_pipe_path"]) if path_info["reg_pipe_path"] else "— mungon —",
            "Clf model": str(path_info["clf_model_path"]) if path_info["clf_model_path"] else "— mungon —",
            "Clf pipeline": str(path_info["clf_pipe_path"]) if path_info["clf_pipe_path"] else "— mungon —",
        }
    )
    st.write("**Dosjet e skanuara:**")
    for r in path_info["roots"]:
        st.text(str(r))
    if st.button("Pastro cache & rikërko modelet"):
        st.cache_resource.clear()
        st.cache_data.clear()
        st.rerun()

if ews_err:
    st.error(ews_err)
if reg_err:
    st.warning(reg_err)

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
tab1, tab2, tab3 = st.tabs(
    [
        "🎯 Parashikimi Individual (EWS)",
        "📊 Analiza & Explorer",
        "ℹ️ Informacioni mbi Modelin",
    ]
)

# ========================= TAB 1 =========================
with tab1:
    st.subheader("Early Warning System — Parashikim individual")
    st.write(
        "Fut të dhënat në menynë anësore dhe kliko butonin për të llogaritur "
        "mesataren e parashikuar dhe rrezikun e vonesës."
    )

    col_btn, _ = st.columns([1, 2])
    with col_btn:
        calc = st.button("Kalkulo Parashikimin e Studentit", type="primary", width='content')

    if calc:
        if ews_bundle is None or reg_model is None or reg_pipe is None:
            st.error(
                "Modelet nuk janë të gatshme. Shiko panelin **Diagnostika e modeleve** "
                "në sidebar për CWD-në aktuale dhe emrat e skedarëve që mungojnë.\n\n"
                f"- CWD: `{path_info['cwd']}`\n"
                f"- EWS: `{path_info['ews_path'] or 'MUNGON'}`\n"
                f"- Reg model: `{path_info['reg_model_path'] or 'MUNGON'}`\n"
                f"- Reg pipeline: `{path_info['reg_pipe_path'] or 'MUNGON'}`\n\n"
                "Vendosi `.joblib` në root, `PreProces/`, `PreProces/ML/` ose `models/` "
                "(emrat janë case-insensitive), pastaj kliko **Pastro cache & rikërko**."
            )
        else:
            try:
                X_raw = build_input_dataframe(sidebar_values)
                pred = predict_all(X_raw, ews_bundle, reg_model, reg_pipe)
                kategori, mesazh, css = risk_status(pred["proba"])

                m1, m2, m3 = st.columns(3)
                m1.metric(
                    "Mesatarja e Maturës (input)",
                    f"{sidebar_values['MES_MATURA_UI']:.2f}",
                )
                m2.metric(
                    "MES_PERGJ e parashikuar",
                    f"{pred['mes_pergj']:.2f}",
                    help="Nota mesatare përfundimtare e parashikuar nga Random Forest",
                )
                m3.metric("Probabiliteti i vonesës", f"{pred['proba_pct']:.1f}%")

                g1, g2 = st.columns([1.2, 1])
                with g1:
                    st.plotly_chart(gauge_chart(pred["proba_pct"]), width='content')
                with g2:
                    st.markdown(f"### Statusi: **{kategori}**")
                    st.markdown(
                        f'<div class="{css}">{mesazh}</div>',
                        unsafe_allow_html=True,
                    )
                    st.caption(
                        f"Pragjet: Ulët < {int(RISK_LOW*100)}% · "
                        f"Mesëm {int(RISK_LOW*100)}–{int(RISK_MED*100)}% · "
                        f"Lartë > {int(RISK_MED*100)}%"
                    )
                    with st.expander("Parametrat e dërguar te modeli"):
                        st.dataframe(X_raw.T.rename(columns={0: "Vlera"}), width='content')
            except Exception as exc:  # noqa: BLE001
                st.exception(exc)
    else:
        st.info("Pritet veprimi yt — kliko **Kalkulo Parashikimin e Studentit**.")

# ========================= TAB 2 =========================
with tab2:
    st.subheader("Analiza dhe Explorer i të dhënave UAMD")

    uploaded = st.file_uploader(
        "Ngarko një dataset alternativ (CSV/Excel) — opsionale",
        type=["csv", "xlsx"],
    )

    df_view = None
    source_label = ""
    if uploaded is not None:
        try:
            if uploaded.name.lower().endswith(".csv"):
                df_view = pd.read_csv(uploaded)
            else:
                df_view = pd.read_excel(uploaded)
            source_label = f"Skedar i ngarkuar: {uploaded.name}"
        except Exception as exc:  # noqa: BLE001
            st.error(f"Ngarkimi dështoi: {exc}")
    elif raw_df is not None:
        df_view = raw_df
        source_label = "Dataset zyrtar: Studentet_Uamd_pastruar_spss.xlsx"
    else:
        st.warning("Nuk u gjet dataset-i burimor Excel. Ngarko një skedar manualisht.")

    if df_view is not None:
        st.caption(source_label)
        st.write(f"Dimenzionet: **{df_view.shape[0]}** rreshta × **{df_view.shape[1]}** kolona")
        st.dataframe(df_view.head(100), width='content', height=280)

        c1, c2 = st.columns(2)

        with c1:
            st.markdown("#### Shpërndarja e mesatareve sipas Fakultetit / Degës")
            group_col = st.radio(
                "Grupo sipas",
                ["FAKULTETI", "DEGA"],
                horizontal=True,
                key="group_col",
            )
            value_col = "MES_PERGJ" if "MES_PERGJ" in df_view.columns else None
            if value_col and group_col in df_view.columns:
                plot_df = df_view[[group_col, value_col]].dropna()
                if group_col == "DEGA":
                    top_dega = plot_df[group_col].value_counts().head(12).index
                    plot_df = plot_df[plot_df[group_col].isin(top_dega)]
                fig = px.box(
                    plot_df,
                    x=group_col,
                    y=value_col,
                    color=group_col,
                    points="outliers",
                    title=f"{value_col} sipas {group_col}",
                )
                fig.update_layout(showlegend=False, xaxis_tickangle=-35, height=420)
                st.plotly_chart(fig, width='content')
            else:
                st.info("Mungojnë kolonat e nevojshme për këtë grafik.")

        with c2:
            st.markdown("#### Correlation Heatmap — notat bazë")
            grade_cols = [
                c
                for c in [
                    "Baze_11",
                    "Zgjedhje_11",
                    "Baze_12",
                    "Zgjedhje_12",
                    "MES_VIT1_SEM1",
                    "MES_PERGJ",
                    "MOSHA",
                ]
                if c in df_view.columns
            ]
            if len(grade_cols) >= 2:
                corr = df_view[grade_cols].corr(numeric_only=True)
                fig_h = px.imshow(
                    corr,
                    text_auto=".2f",
                    color_continuous_scale="RdBu_r",
                    zmin=-1,
                    zmax=1,
                    title="Korrelacioni i notave / mesatareve",
                )
                fig_h.update_layout(height=420)
                st.plotly_chart(fig_h, width='content')
            else:
                st.info("Nuk u gjetën mjaftueshëm kolona numerike për heatmap.")

# ========================= TAB 3 =========================
with tab3:
    st.subheader("Metrikat e modeleve dhe Feature Importance")

    left, right = st.columns(2)
    with left:
        st.markdown("#### Klasifikim (RREZIK_VONESE)")
        if CLF_METRICS is not None and Path(CLF_METRICS).exists():
            clf_m = pd.read_csv(CLF_METRICS)
            st.dataframe(clf_m, width='content', hide_index=True)
            best = clf_m.iloc[0]
            k1, k2, k3 = st.columns(3)
            k1.metric("Accuracy", f"{best['Accuracy']:.3f}")
            k2.metric("F1-Score", f"{best['F1']:.3f}")
            k3.metric("ROC-AUC", f"{best['ROC_AUC']:.3f}")
        else:
            st.warning("Nuk u gjet `classification_metrics.csv` në dosjet e kërkuara.")

    with right:
        st.markdown("#### Regresion (MES_PERGJ)")
        if REG_METRICS is not None and Path(REG_METRICS).exists():
            reg_m = pd.read_csv(REG_METRICS)
            st.dataframe(reg_m, width='content', hide_index=True)
            best_r = reg_m.iloc[0]
            r1, r2, r3 = st.columns(3)
            r1.metric("RMSE", f"{best_r['RMSE']:.3f}")
            r2.metric("MAE", f"{best_r['MAE']:.3f}")
            r3.metric("R²", f"{best_r['R2']:.3f}")
        else:
            st.warning("Nuk u gjet `regression_metrics.csv` në dosjet e kërkuara.")

    st.markdown("#### Feature Importance (modeli i regresionit)")
    if FEAT_IMP is not None and Path(FEAT_IMP).exists():
        fi = pd.read_csv(FEAT_IMP, index_col=0)
        fi = fi.reset_index()
        fi.columns = ["feature", "importance"]
        fi["importance"] = pd.to_numeric(fi["importance"], errors="coerce")
        fi = fi.dropna().sort_values("importance", ascending=True)
        fig_fi = px.bar(
            fi.tail(12),
            x="importance",
            y="feature",
            orientation="h",
            title="12 veçoritë më ndikuese për MES_PERGJ",
            color="importance",
            color_continuous_scale="Blues",
        )
        fig_fi.update_layout(height=480, showlegend=False, coloraxis_showscale=False)
        st.plotly_chart(fig_fi, width='content')
        st.dataframe(
            fi.sort_values("importance", ascending=False).reset_index(drop=True),
            width='content',
            hide_index=True,
        )
    else:
        st.warning("Nuk u gjet `regression_feature_importance.csv` në dosjet e kërkuara.")

    st.caption(
        "Modelet: Random Forest (regresion & klasifikim) · Pipeline preprocessing me "
        "imputim, One-Hot / Target Encoding dhe StandardScaler · pa data leakage nga "
        "notat e viteve të mëvonshme."
    )
