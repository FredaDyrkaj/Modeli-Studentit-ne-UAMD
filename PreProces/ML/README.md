# Faza D — Modelimi Parashikues & Survival Analysis

Skripti: `faza_d_modeling_survival.py`

## Ekzekutimi

```powershell
cd C:\Users\esoft\Desktop\Projekt-ModeliStudenteve\FInalWork
.\.venv\Scripts\python.exe PreProces\ML\faza_d_modeling_survival.py
```

## Hyrjet

- `../dataset_ml_regression.csv`
- `../dataset_ml_classification.csv`
- `../../DescriptiveAnalyse/Studentet_Uamd_pastruar_spss.xlsx`

## Daljet

| Artefakt | Përshkrimi |
|----------|------------|
| `Faza_D_Modelimi_dhe_Survival_Analysis_UAMD.docx` | Raporti akademik |
| `figures/*.png` | ROC, CM, KM, HR, **SHAP summary/waterfall**, metrika |
| `ews_bundle.joblib` | Pipeline + model për Early Warning |
| `best_*.joblib` | Modelet fituese |

## Aplikacioni Streamlit (EWS Dashboard)

```powershell
cd C:\Users\esoft\Desktop\Projekt-ModeliStudenteve\FInalWork
.\.venv\Scripts\python.exe -m streamlit run PreProces\ML\app.py
```

Aplikacioni përdor:
- `ews_bundle.joblib` — klasifikimi i rrezikut
- `best_regression_RandomForest.joblib` + `../preprocessing_pipeline_regression.joblib` — MES_PERGJ
- Excel-in burimor për explorer / listat e selectbox

## Survival (opsioni 1)

- `event=1`: diplomuar (`KOHA_DIPLOM` ose `DIPLOMUAR`)
- `event=0`: ende aktiv (censuruar)
- `duration`: `KOHA_DIPLOM` ose viti i studimit I–VI
