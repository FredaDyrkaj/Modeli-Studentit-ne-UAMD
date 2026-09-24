"""
EDA akademike — Modeli i Studentit Optimal në UAMD
===================================================
Lexon Studentet_Uamd_pastruar_spss.xlsx, kryen analizë univariate & bivariate,
ruan grafikët në charts/ (PNG 300 DPI)

Instalim (nga folderi DescriptiveAnalyse):
    ..\\.venv\\Scripts\\python.exe -m pip install -r requirements_eda_optimal.txt

Ekzekutim:
    ..\\.venv\\Scripts\\python.exe eda_studenti_optimal_uamd.py
    ..\\.venv\\Scripts\\python.exe eda_studenti_optimal_uamd.py --no-show
"""

from __future__ import annotations

import argparse
import warnings
from datetime import datetime
from pathlib import Path
from typing import Any

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from scipy import stats
from statsmodels.formula.api import ols
from statsmodels.stats.anova import anova_lm

warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", category=UserWarning)

# ---------------------------------------------------------------------------
# Konfigurim
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT = BASE_DIR / "Studentet_Uamd_pastruar_spss.xlsx"
CHARTS_DIR = BASE_DIR / "charts"

# Emrat e përdoruesit → emrat realë në Excel (nëse ndryshojnë)
COLUMN_ALIASES: dict[str, str] = {
    "PROFIL_GJIMNAZI": "PROFIL_GJIMNAZ",
    "DATA_DIPLOMIMIT": "DATA_DIPLOMA",
    "VIT_DIPLOMIMIT": "VIT_DIPLOMIM",
    "NIVELI_STUDIMEVE": "NIVELI_STUDIMIT",
    "KOHA_DIPLOMIMIT": "KOHA_DIPLOM",
    "MES_PERGJ_BAZE": "MES_PERGJ_BACH",
    "MES_PERGJ_MATURA": "MES_PERGJ_MATURA",  # mund të mungojë
}

CATEGORICAL_CANDIDATES = [
    "GJINIA",
    "RRETHI",
    "GJIMNAZI",
    "PROFIL_GJIMNAZ",
    "LLOJ_RREGJ",
    "NIVELI_STUDIMIT",
    "VIT_AKADEMIK",
    "FAKULTETI",
    "DEGA",
]

NUMERIC_CANDIDATES = [
    "MOSHA",
    "VITI_DTL",
    "VITI_RREGJ",
    "VIT_DIPLOMIM",
    "KOHA_DIPLOM",
    "MES_BAZE",
    "MES_ZGJEDHJE",
    "MES_PERGJ_BACH",
    "MES_PERGJ_MAST",
    "MES_PERGJ_MATURA",
    "MES_VIT1_SEM1",
    "MES_PERGJ",
]

GRADE_COLS_PREFIX = ("Baze_", "Zgjedhje_")
TOP_N_HIGH_CARD = 12
ALPHA = 0.05
PALETTE = "viridis"
DPI = 300

sns.set_theme(style="whitegrid", palette=PALETTE, font_scale=1.05)
plt.rcParams.update(
    {
        "figure.dpi": 120,
        "savefig.dpi": DPI,
        "axes.titlesize": 13,
        "axes.labelsize": 11,
        "figure.facecolor": "white",
    }
)


# ---------------------------------------------------------------------------
# Ndihmësa të përgjithshme
# ---------------------------------------------------------------------------
def resolve_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Riemërton alias-et e kërkuara nëse ekzistojnë me emër alternativ."""
    out = df.copy()
    rename = {}
    for wanted, actual in COLUMN_ALIASES.items():
        if wanted not in out.columns and actual in out.columns:
            # mbajmë emrin real; aliaset përdoren vetëm për lookup
            pass
        if wanted in out.columns and actual not in out.columns and wanted != actual:
            rename[wanted] = actual
    if rename:
        out = out.rename(columns=rename)
    return out


def col(df: pd.DataFrame, *names: str) -> str | None:
    """Kthen emrin e parë të kolonës që ekziston."""
    for n in names:
        if n in df.columns:
            return n
        alias = COLUMN_ALIASES.get(n)
        if alias and alias in df.columns:
            return alias
    return None


def present(df: pd.DataFrame, names: list[str]) -> list[str]:
    found = []
    for n in names:
        c = col(df, n)
        if c and c not in found:
            found.append(c)
    return found


def save_fig(fig: plt.Figure, name: str, show: bool) -> Path:
    CHARTS_DIR.mkdir(parents=True, exist_ok=True)
    path = CHARTS_DIR / name
    fig.savefig(path, bbox_inches="tight", dpi=DPI, facecolor="white")
    if show:
        plt.show()
    else:
        plt.close(fig)
    print(f"  [chart] {path.name}")
    return path


def sig_label(p: float) -> str:
    if pd.isna(p):
        return "n/a"
    if p < 0.001:
        return "p < 0.001 *** (shumë i rëndësishëm)"
    if p < 0.01:
        return f"p = {p:.4f} ** (i rëndësishëm)"
    if p < 0.05:
        return f"p = {p:.4f} * (i rëndësishëm në α=0.05)"
    return f"p = {p:.4f} (jo i rëndësishëm në α=0.05)"


def safe_skew(s: pd.Series) -> float:
    s = pd.to_numeric(s, errors="coerce").dropna()
    if len(s) < 3:
        return float("nan")
    return float(stats.skew(s, bias=False))


# ---------------------------------------------------------------------------
# 1) ANALIZA UNIVARIATE
# ---------------------------------------------------------------------------
def frequency_table(series: pd.Series, top_n: int | None = None) -> pd.DataFrame:
    vc = series.fillna("(mungon)").astype(str).value_counts(dropna=False)
    if top_n is not None and len(vc) > top_n:
        top = vc.head(top_n)
        other = vc.iloc[top_n:].sum()
        vc = pd.concat([top, pd.Series({"(të tjera)": other})])
    n = int(vc.sum())
    out = vc.rename_axis("Kategoria").reset_index(name="Frekuenca")
    out["Përqindja_%"] = (100 * out["Frekuenca"] / n).round(2)
    out["Kumulative_%"] = out["Përqindja_%"].cumsum().round(2)
    return out


def descriptive_stats(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    rows = []
    for c in columns:
        s = pd.to_numeric(df[c], errors="coerce")
        # MOSHA=0 është e pavlefshme në këtë skedar
        if c.upper() == "MOSHA":
            s = s.where(s > 0)
        valid = s.dropna()
        if valid.empty:
            continue
        rows.append(
            {
                "Variabli": c,
                "N": int(valid.count()),
                "Mungesa": int(s.isna().sum()),
                "Mesatarja": round(float(valid.mean()), 3),
                "Mediana": round(float(valid.median()), 3),
                "Std.Dev": round(float(valid.std(ddof=1)), 3),
                "Min": round(float(valid.min()), 3),
                "Max": round(float(valid.max()), 3),
                "Skewness": round(safe_skew(valid), 3),
            }
        )
    return pd.DataFrame(rows)


def plot_categorical_bar(df: pd.DataFrame, column: str, show: bool, top_n: int = 15) -> Path:
    tab = frequency_table(df[column], top_n=top_n)
    fig, ax = plt.subplots(figsize=(10, max(4.5, 0.35 * len(tab))))
    sns.barplot(data=tab, y="Kategoria", x="Frekuenca", hue="Kategoria", legend=False, ax=ax, palette=PALETTE)
    ax.set_title(f"Shpërndarja e frekuencave — {column}")
    ax.set_xlabel("Frekuenca")
    ax.set_ylabel(column)
    fig.tight_layout()
    return save_fig(fig, f"univ_freq_{column}.png", show)


def plot_numeric_hist(df: pd.DataFrame, column: str, show: bool) -> Path:
    s = pd.to_numeric(df[column], errors="coerce")
    if column.upper() == "MOSHA":
        s = s.where(s > 0)
    s = s.dropna()
    fig, ax = plt.subplots(figsize=(8.5, 4.8))
    sns.histplot(s, kde=True, bins=30, color="#1f4e79", ax=ax)
    ax.axvline(s.mean(), color="#c0392b", linestyle="--", label=f"Mesatarja {s.mean():.2f}")
    ax.axvline(s.median(), color="#27ae60", linestyle=":", label=f"Mediana {s.median():.2f}")
    ax.set_title(f"Shpërndarja — {column}")
    ax.set_xlabel(column)
    ax.legend()
    fig.tight_layout()
    return save_fig(fig, f"univ_hist_{column}.png", show)


# ---------------------------------------------------------------------------
# 2) ANALIZA BIVARIATE + teste
# ---------------------------------------------------------------------------
def normality_ok(series: pd.Series, max_n: int = 5000) -> bool:
    s = pd.to_numeric(series, errors="coerce").dropna()
    if len(s) < 8:
        return False
    sample = s.sample(min(len(s), max_n), random_state=42)
    # Shapiro është i ndjeshëm për N të madh; përdorim subsample
    if len(sample) > 500:
        sample = sample.sample(500, random_state=42)
    try:
        _, p = stats.shapiro(sample)
        return bool(p > ALPHA)
    except Exception:
        return False


def compare_two_groups(df: pd.DataFrame, y: str, group_col: str) -> dict[str, Any]:
    data = df[[y, group_col]].dropna()
    levels = data[group_col].astype(str).value_counts()
    levels = levels[levels >= 5].index.tolist()
    if len(levels) < 2:
        return {"test": "n/a", "stat": np.nan, "p": np.nan, "note": "Më pak se 2 grupe me N≥5"}
    # merr dy grupet më të mëdha nëse ka më shumë
    if len(levels) > 2:
        top2 = data[group_col].astype(str).value_counts().head(2).index.tolist()
        data = data[data[group_col].astype(str).isin(top2)]
        levels = top2
        note_extra = f" (krahasim i 2 niveleve kryesore: {levels})"
    else:
        note_extra = ""
    g1 = pd.to_numeric(data.loc[data[group_col].astype(str) == levels[0], y], errors="coerce").dropna()
    g2 = pd.to_numeric(data.loc[data[group_col].astype(str) == levels[1], y], errors="coerce").dropna()
    if normality_ok(g1) and normality_ok(g2):
        stat, p = stats.ttest_ind(g1, g2, equal_var=False)
        test = "Welch t-test"
    else:
        stat, p = stats.mannwhitneyu(g1, g2, alternative="two-sided")
        test = "Mann-Whitney U"
    return {
        "test": test,
        "stat": float(stat),
        "p": float(p),
        "group1": levels[0],
        "group2": levels[1],
        "mean1": float(g1.mean()),
        "mean2": float(g2.mean()),
        "n1": int(len(g1)),
        "n2": int(len(g2)),
        "note": note_extra.strip(),
    }


def anova_or_kruskal(df: pd.DataFrame, y: str, group_col: str, max_levels: int = 12) -> dict[str, Any]:
    data = df[[y, group_col]].dropna().copy()
    data[group_col] = data[group_col].astype(str)
    counts = data[group_col].value_counts()
    keep = counts[counts >= 5].index
    data = data[data[group_col].isin(keep)]
    if data[group_col].nunique() > max_levels:
        top = counts.head(max_levels).index
        data = data[data[group_col].isin(top)]
    groups = [pd.to_numeric(g[y], errors="coerce").dropna() for _, g in data.groupby(group_col)]
    groups = [g for g in groups if len(g) >= 5]
    if len(groups) < 2:
        return {"test": "n/a", "stat": np.nan, "p": np.nan, "note": "Grupe të pamjaftueshme"}
    normal = all(normality_ok(g) for g in groups)
    if normal and len(groups) >= 2:
        # ANOVA me statsmodels për tabelë të plotë
        formula = f"Q('{y}') ~ C(Q('{group_col}'))"
        try:
            model = ols(formula, data=data).fit()
            table = anova_lm(model, typ=2)
            row = table.iloc[0]
            return {
                "test": "One-way ANOVA",
                "stat": float(row["F"]),
                "p": float(row["PR(>F)"]),
                "df": f"{int(row['df'])}",
                "note": f"nivele={data[group_col].nunique()}, N={len(data)}",
                "means": data.groupby(group_col)[y].mean().round(3).to_dict(),
            }
        except Exception as exc:
            return {"test": "ANOVA dështoi", "stat": np.nan, "p": np.nan, "note": str(exc)}
    stat, p = stats.kruskal(*groups)
    return {
        "test": "Kruskal-Wallis H",
        "stat": float(stat),
        "p": float(p),
        "note": f"nivele={len(groups)}, N={len(data)} (jo-parametrike)",
        "means": data.groupby(group_col)[y].mean().round(3).to_dict(),
    }


def correlation_pair(df: pd.DataFrame, x: str, y: str) -> dict[str, Any]:
    sub = df[[x, y]].apply(pd.to_numeric, errors="coerce").dropna()
    if len(sub) < 10:
        return {"n": len(sub), "pearson_r": np.nan, "pearson_p": np.nan, "spearman_r": np.nan, "spearman_p": np.nan}
    pr, pp = stats.pearsonr(sub[x], sub[y])
    sr, sp = stats.spearmanr(sub[x], sub[y])
    return {
        "n": int(len(sub)),
        "pearson_r": float(pr),
        "pearson_p": float(pp),
        "spearman_r": float(sr),
        "spearman_p": float(sp),
        "mean_x": float(sub[x].mean()),
        "mean_y": float(sub[y].mean()),
    }


def plot_boxplot(df: pd.DataFrame, y: str, x: str, title: str, fname: str, show: bool, top_n: int = 10) -> Path:
    data = df[[x, y]].dropna().copy()
    data[x] = data[x].astype(str)
    top = data[x].value_counts().head(top_n).index
    data = data[data[x].isin(top)]
    order = data.groupby(x)[y].median().sort_values(ascending=False).index.tolist()
    fig, ax = plt.subplots(figsize=(11, 5.5))
    sns.boxplot(data=data, x=x, y=y, order=order, hue=x, legend=False, palette=PALETTE, ax=ax)
    ax.set_title(title)
    ax.tick_params(axis="x", rotation=35)
    fig.tight_layout()
    return save_fig(fig, fname, show)


def plot_bar_means(df: pd.DataFrame, y: str, x: str, title: str, fname: str, show: bool, top_n: int = 12) -> Path:
    data = df[[x, y]].dropna().copy()
    data[x] = data[x].astype(str)
    top = data[x].value_counts().head(top_n).index
    data = data[data[x].isin(top)]
    means = data.groupby(x)[y].agg(["mean", "count"]).reset_index()
    means = means.sort_values("mean", ascending=False)
    fig, ax = plt.subplots(figsize=(10, 5.2))
    sns.barplot(data=means, x=x, y="mean", hue=x, legend=False, palette=PALETTE, ax=ax)
    ax.set_title(title)
    ax.set_ylabel(f"Mesatarja e {y}")
    ax.tick_params(axis="x", rotation=35)
    fig.tight_layout()
    return save_fig(fig, fname, show)


def plot_scatter(df: pd.DataFrame, x: str, y: str, title: str, fname: str, show: bool) -> Path:
    sub = df[[x, y]].apply(pd.to_numeric, errors="coerce").dropna()
    fig, ax = plt.subplots(figsize=(7.8, 5.6))
    sns.regplot(
        data=sub,
        x=x,
        y=y,
        ax=ax,
        scatter_kws={"alpha": 0.25, "s": 18, "color": "#1f4e79"},
        line_kws={"color": "#c0392b"},
    )
    ax.set_title(title)
    fig.tight_layout()
    return save_fig(fig, fname, show)


def plot_corr_heatmap(df: pd.DataFrame, columns: list[str], fname: str, show: bool) -> Path | None:
    cols = [c for c in columns if c in df.columns]
    if len(cols) < 2:
        return None
    mat = df[cols].apply(pd.to_numeric, errors="coerce").corr(method="pearson")
    fig, ax = plt.subplots(figsize=(max(8, 0.7 * len(cols)), max(6.5, 0.6 * len(cols))))
    sns.heatmap(mat, annot=True, fmt=".2f", cmap="Blues", vmin=-1, vmax=1, square=True, ax=ax)
    ax.set_title("Matrica e korrelacionit (Pearson) — mesataret & notat bazë")
    fig.tight_layout()
    return save_fig(fig, fname, show)


# ---------------------------------------------------------------------------
# Orkestrimi kryesor
# ---------------------------------------------------------------------------
def run_eda(input_path: Path, show: bool) -> int:
    print("=" * 72)
    print("EDA — Modeli i Studentit Optimal në UAMD")
    print("=" * 72)
    if not input_path.is_file():
        raise FileNotFoundError(f"Nuk u gjet skedari: {input_path}")

    df = pd.read_excel(input_path, sheet_name=0, engine="openpyxl")
    df = resolve_columns(df)
    print(f"Ngarkuar: {input_path.name} | forma={df.shape}")

    notes: list[str] = []
    matura_col = col(df, "MES_PERGJ_MATURA")
    if matura_col is None:
        notes.append(
            "Kolona MES_PERGJ_MATURA nuk ekziston në skedar. Analiza e hyrjes "
            "përdor PROFIL_GJIMNAZ (kategorike) dhe MES_VIT1_SEM1 si tregues i "
            "hershëm i performancës; MES_PERGJ_BACH / MES_PERGJ_MAST përdoren "
            "si mesatare sipas nivelit."
        )
        print("  ! MES_PERGJ_MATURA mungon — përdoren proxy akademike të disponueshme.")

    y = col(df, "MES_PERGJ")
    if y is None:
        raise ValueError("Kolona MES_PERGJ është e detyrueshme dhe mungon.")

    cat_cols = present(df, CATEGORICAL_CANDIDATES)
    num_cols = present(df, NUMERIC_CANDIDATES)
    baze_cols = [c for c in df.columns if str(c).startswith("Baze_")]
    # vetëm notat bazë me mbulim të arsyeshëm (>5% jo-null)
    baze_usable = [c for c in baze_cols if df[c].notna().mean() >= 0.05]

    # ---- Univariate ----
    print("\n[1/4] Analiza univariate...")
    freq_tables: dict[str, pd.DataFrame] = {}
    chart_paths: dict[str, Path | None] = {}
    high_card = {"RRETHI", "GJIMNAZI", "DEGA"}
    for c in cat_cols:
        top = TOP_N_HIGH_CARD if c in high_card else None
        # për LLOJ etj. me pak nivele — pa top
        if c not in high_card and df[c].nunique(dropna=True) > 20:
            top = TOP_N_HIGH_CARD
        tab = frequency_table(df[c], top_n=top)
        freq_tables[c] = tab
        print(f"  frekuenca {c}: {len(tab)} rreshta")
        chart_paths[f"univ_freq_{c}"] = plot_categorical_bar(
            df, c, show=show, top_n=top or 20
        )

    desc = descriptive_stats(df, num_cols)
    print(desc.to_string(index=False))
    # histograme për mesataret kryesore
    hist_targets = [c for c in ["MES_PERGJ", "MES_BAZE", "MES_ZGJEDHJE", "MES_VIT1_SEM1", "MOSHA", "KOHA_DIPLOM"] if c in num_cols]
    if matura_col:
        hist_targets.append(matura_col)
    for c in hist_targets:
        chart_paths[f"univ_hist_{c}"] = plot_numeric_hist(df, c, show=show)

    # ---- Bivariate ----
    print("\n[2/4] Analiza bivariate & teste...")
    bivariate: list[dict[str, Any]] = []
    exec_bits: list[str] = []

    # KPI bazë
    gpa = pd.to_numeric(df[y], errors="coerce").dropna()
    exec_bits.append(
        f"Kampioni: {len(df):,} regjistrime; MES_PERGJ i disponueshëm për {len(gpa):,} "
        f"rreshta (mesatarja={gpa.mean():.2f}, mediana={gpa.median():.2f})."
    )

    # 4.1 PROFIL_GJIMNAZ
    profil = col(df, "PROFIL_GJIMNAZ", "PROFIL_GJIMNAZI")
    if profil:
        chart = plot_boxplot(
            df,
            y=y,
            x=profil,
            title=f"{y} sipas {profil}",
            fname="biv_boxplot_profil.png",
            show=show,
            top_n=10,
        )
        res = anova_or_kruskal(df, y, profil)
        means_txt = ", ".join(f"{k}: {v}" for k, v in (res.get("means") or {}).items())
        bivariate.append(
            {
                "title": f"{y} × {profil} (Boxplot + ANOVA/Kruskal)",
                "commentary": (
                    f"Krahasohet performanca në UAMD ({y}) sipas profilit të gjimnazit. "
                    "Nëse ka ndryshim statistikor mes profileve, profili i shkollës së mesme "
                    "mund të jetë një sinjal i hershëm i suksesit akademik."
                ),
                "chart": chart,
                "stats_text": (
                    f"Testi: {res.get('test')} | Statistika={res.get('stat'):.4f} | "
                    f"{sig_label(res.get('p', np.nan))} | {res.get('note','')} | "
                    f"Mesataret sipas grupit: {means_txt}"
                ),
                "interpretation": (
                    "Ka diferencë të rëndësishme mes profileve; profili i gjimnazit "
                    "duhet përfshirë në modelin e studentit optimal."
                    if (res.get("p") is not None and res["p"] < ALPHA)
                    else "Nuk evidentohet diferencë e rëndësishme mes profileve në këtë kampion; "
                    "profili vetëm nuk mjafton për të shpjeguar MES_PERGJ."
                ),
            }
        )
        if res.get("p") is not None and not pd.isna(res["p"]):
            exec_bits.append(
                f"PROFIL_GJIMNAZ vs {y}: {res.get('test')}, {sig_label(res['p'])}."
            )

    # 4.2 Matura ose proxy
    if matura_col:
        x = matura_col
        title = f"{y} vs {x} (scatter + Pearson/Spearman)"
        commentary = (
            f"Lidhja lineare/monotone mes mesatares së maturës ({x}) dhe performancës në UAMD ({y})."
        )
    else:
        # proxy: MES_BAZE si “bazë akademike”, plus raportojmë qartë mungesën
        x = col(df, "MES_BAZE")
        title = f"{y} vs {x} (proxy për mungesën e MES_PERGJ_MATURA)"
        commentary = (
            "MES_PERGJ_MATURA nuk është në dataset. Si analizë e lidhjes së performancës "
            f"së brendshme, krahasojmë {y} me {x} (mesatarja e lëndëve bazë). "
            "Kur të shtoni notën e maturës, skripti do ta përdorë automatikisht."
        )

    if x:
        chart = plot_scatter(df, x=x, y=y, title=title, fname="biv_scatter_matura_or_proxy.png", show=show)
        corr = correlation_pair(df, x, y)
        exec_bits.append(
            f"Korrelacion {x}–{y}: Pearson r={corr['pearson_r']:.3f} ({sig_label(corr['pearson_p'])})."
        )

    # 4.3 GJINIA
    gjinia = col(df, "GJINIA")
    if gjinia:
        chart = plot_bar_means(
            df, y=y, x=gjinia, title=f"Mesatarja e {y} sipas {gjinia}", fname="biv_bar_gjinia.png", show=show
        )
        # boxplot shtesë
        chart_box = plot_boxplot(
            df, y=y, x=gjinia, title=f"Boxplot {y} × {gjinia}", fname="biv_box_gjinia.png", show=show
        )
        res = compare_two_groups(df, y, gjinia)
        bivariate.append(
            {
                "title": f"{y} × {gjinia} (Barplot + t-test / Mann-Whitney)",
                "commentary": (
                    "Krahasohet performanca mes gjinive. Nëse diferenca është e rëndësishme, "
                    "gjinia duhet kontrolluar në modele multivariate (pa e interpretuar si kauzalitet)."
                ),
                "chart": chart,
                "stats_text": (
                    f"Testi: {res.get('test')} | stat={res.get('stat'):.4f} | {sig_label(res.get('p', np.nan))} | "
                    f"{res.get('group1')} (N={res.get('n1')}, mean={res.get('mean1'):.3f}) vs "
                    f"{res.get('group2')} (N={res.get('n2')}, mean={res.get('mean2'):.3f}) {res.get('note','')}"
                ),
                "interpretation": (
                    f"Diferenca mes {res.get('group1')} dhe {res.get('group2')} është statistikisht e rëndësishme."
                    if res.get("p") is not None and res["p"] < ALPHA
                    else "Nuk ka diferencë statistikisht të rëndësishme mes gjinive për MES_PERGJ."
                ),
            }
        )
        chart_paths["biv_box_gjinia"] = chart_box
        if res.get("p") is not None and not pd.isna(res["p"]):
            exec_bits.append(
                f"GJINIA: {res.get('group1')} mean={res.get('mean1'):.2f} vs "
                f"{res.get('group2')} mean={res.get('mean2'):.2f}; {sig_label(res['p'])}."
            )

    # 4.4 RRETHI (top-N + Kruskal)
    rrethi = col(df, "RRETHI")
    if rrethi:
        chart = plot_bar_means(
            df,
            y=y,
            x=rrethi,
            title=f"Mesatarja e {y} — top {TOP_N_HIGH_CARD} rrethe",
            fname="biv_bar_rrethi.png",
            show=show,
            top_n=TOP_N_HIGH_CARD,
        )
        res = anova_or_kruskal(df, y, rrethi, max_levels=TOP_N_HIGH_CARD)
        bivariate.append(
            {
                "title": f"{y} × {rrethi} (Barplot top-N + ANOVA/Kruskal)",
                "commentary": (
                    "Analizohet nëse rrethi i origjinës shoqërohet me diferenca në MES_PERGJ. "
                    "Për shkak të kardinalitetit të lartë, krahasohen rrethet më të përfaqësuara."
                ),
                "chart": chart,
                "stats_text": (
                    f"Testi: {res.get('test')} | stat={res.get('stat'):.4f} | {sig_label(res.get('p', np.nan))} | "
                    f"{res.get('note','')}"
                ),
                "interpretation": (
                    "Ka diferenca të rëndësishme mes rretheve për performancën."
                    if res.get("p") is not None and res["p"] < ALPHA
                    else "Nuk evidentohet efekt i rëndësishëm i rrethit (top-N) mbi MES_PERGJ."
                ),
            }
        )

    # 4.5 DEGA
    dega = col(df, "DEGA")
    if dega:
        chart = plot_boxplot(
            df,
            y=y,
            x=dega,
            title=f"{y} sipas degës (top 10)",
            fname="biv_box_dega.png",
            show=show,
            top_n=10,
        )
        res = anova_or_kruskal(df, y, dega, max_levels=10)
        bivariate.append(
            {
                "title": f"{y} × {dega} (Boxplot + ANOVA/Kruskal)",
                "commentary": (
                    "Dega/programi është shpesh determinues i ngarkesës akademike dhe i shpërndarjes "
                    "së notave. Krahasimi mes degëve ndihmon të kuptohet ku përqendrohet 'studenti optimal'."
                ),
                "chart": chart,
                "stats_text": (
                    f"Testi: {res.get('test')} | stat={res.get('stat'):.4f} | {sig_label(res.get('p', np.nan))} | "
                    f"{res.get('note','')}"
                ),
                "interpretation": (
                    "Performanca ndryshon në mënyrë të rëndësishme sipas degës — dega është faktor kyç."
                    if res.get("p") is not None and res["p"] < ALPHA
                    else "Nuk u gjet diferencë e rëndësishme mes degëve të analizuara."
                ),
            }
        )
        if res.get("p") is not None and not pd.isna(res["p"]):
            exec_bits.append(f"DEGA vs {y}: {res.get('test')}, {sig_label(res['p'])}.")

    # 4.5b FAKULTETI
    fak = col(df, "FAKULTETI")
    if fak:
        chart = plot_bar_means(
            df, y=y, x=fak, title=f"Mesatarja e {y} sipas fakultetit", fname="biv_bar_fakulteti.png", show=show
        )
        res = anova_or_kruskal(df, y, fak)
        bivariate.append(
            {
                "title": f"{y} × {fak}",
                "commentary": "Krahasim i performancës mes fakulteteve FB, FTI dhe FSHPJ.",
                "chart": chart,
                "stats_text": (
                    f"Testi: {res.get('test')} | stat={res.get('stat'):.4f} | {sig_label(res.get('p', np.nan))} | "
                    f"Mesataret: {res.get('means')}"
                ),
                "interpretation": (
                    "Fakulteti shoqërohet me ndryshime të rëndësishme të MES_PERGJ."
                    if res.get("p") is not None and res["p"] < ALPHA
                    else "Nuk ka diferencë të rëndësishme mes fakulteteve."
                ),
            }
        )
        if res.get("means"):
            best = max(res["means"], key=res["means"].get)
            exec_bits.append(
                f"FAKULTETI: mesataret {res['means']}; më e larta te {best} ({sig_label(res.get('p', np.nan))})."
            )

    # 4.6 MES_VIT1_SEM1 vs MES_PERGJ
    v1 = col(df, "MES_VIT1_SEM1")
    if v1:
        chart = plot_scatter(
            df,
            x=v1,
            y=y,
            title=f"{v1} vs {y} — a parashikon semestri i parë suksesin final?",
            fname="biv_scatter_vit1_sem1.png",
            show=show,
        )
        corr = correlation_pair(df, v1, y)
        exec_bits.append(
            f"MES_VIT1_SEM1–MES_PERGJ: r={corr['pearson_r']:.3f} ({sig_label(corr['pearson_p'])})."
        )

    # 4.7 Heatmap
    print("\n[3/4] Matrica e korrelacionit...")
    corr_cols = []
    for c in [
        "MES_PERGJ",
        "MES_BAZE",
        "MES_ZGJEDHJE",
        "MES_VIT1_SEM1",
        "MES_PERGJ_BACH",
        "MES_PERGJ_MAST",
        "MES_PERGJ_MATURA",
        "MOSHA",
        "KOHA_DIPLOM",
    ]:
        cc = col(df, c)
        if cc and cc not in corr_cols:
            corr_cols.append(cc)
    # shto disa nota bazë me mbulim
    for c in baze_usable[:6]:
        if c not in corr_cols:
            corr_cols.append(c)
    heat = plot_corr_heatmap(df, corr_cols, "biv_corr_heatmap.png", show=show)
    chart_paths["heatmap"] = heat


    
def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="EDA — Studenti Optimal UAMD")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--no-show", action="store_true", help="Mos i hap dritaret; vetëm ruaj PNG")
    args = parser.parse_args(argv)
    return run_eda(args.input, show=not args.no_show)


if __name__ == "__main__":
    raise SystemExit(main())
