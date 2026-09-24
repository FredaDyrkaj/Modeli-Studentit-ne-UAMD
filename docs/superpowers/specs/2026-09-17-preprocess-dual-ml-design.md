# Design: Preprocessing dual ML (Regresion + Klasifikim) — UAMD

**Data:** 2026-09-17  
**Status:** Miratuar  

## Qëllimi
Dy dataset-e ML pa data leakage, bazuar në demografi + vitin e parë.

## Predictors
MOSHA, GJINIA, RRETHI, GJIMNAZI, PROFIL_GJIMNAZ, LLOJ_RREGJ, DEGA, FAKULTETI,  
Baze_11, Zgjedhje_11, Baze_12, Zgjedhje_12, MES_VIT1_SEM1  
(MES_PERGJ_MATURA nuk ekziston në burim — jashtë scope)

## Targets
- Regresion: MES_PERGJ (rreshta me NaN hiqen)
- Klasifikim: RREZIK_VONESE (rregull hibrid)
  - NORMA = 3 Bachelor / 2 Master
  - Me KOHA_DIPLOM: 1 nëse > NORMA else 0
  - Pa KOHA_DIPLOM: KOHA_AKTUALE = map I..VI → 1..6; > NORMA → 1; ≤ NORMA / DIPLOMUAR / NaN → përjashto

## Pipeline
Impute median / most_frequent → OneHot (low-card) → TargetEncoder (high-card) → StandardScaler (num)  
Split 80/20, random_state=42, stratify për klasifikim.

## Artefakte
dataset_ml_regression.csv, dataset_ml_classification.csv,  
preprocessing_pipeline_regression.joblib, preprocessing_pipeline_classification.joblib
