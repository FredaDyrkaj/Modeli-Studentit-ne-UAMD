"""
Rikodifikon kolonat e Tot_Students_FTI_BUSS_FSHPJ.xlsx sipas rregullave në RIKODIFIKIM.txt.

Përdorim:
    python rikodifikim.py
    python rikodifikim.py --input Tot_Students_FTI_BUSS_FSHPJ.xlsx --output Tot_Students_rikodifikuar.xlsx
"""

from __future__ import annotations

import argparse
import logging
import re
from pathlib import Path
from typing import Any

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_RULES = PROJECT_ROOT / "RIKODIFIKIM.txt"
DEFAULT_INPUT = PROJECT_ROOT / "Tot_Students_FTI_BUSS_FSHPJ.xlsx"
DEFAULT_OUTPUT = PROJECT_ROOT / "Tot_Students_FTI_BUSS_FSHPJ_Rikodifikuar.xlsx"
DEFAULT_SHEET = "Tot_Students_FTI_BUSS_FSHPJ"

# (kolona_destinacion, lloji_parsing, kolona_burim — vetëm kur ndryshon)
SECTION_SPECS: dict[str, tuple[str, str, str | None]] = {
    "GJINIA": ("GJINIA", "dash_last", None),
    "Niveli_studimit": ("Niveli_studimit", "dash_last", None),
    "RRETHI": ("RRETHI", "dash_last", None),
    "SHKOLLA E MESME E KRYER": ("SHKOLLA E MESME E KRYER", "tab", None),
    "LLOJI I REGJISTRIMIT": ("LLOJI I REGJISTRIMIT", "tab", None),
    "DEGA": ("DEGA", "tab", None),
    "VITI": ("VITI", "space_dash", None),
    "Niveli_studimit - FAKULTETI": ("FAKULTETI", "space_dash", "Niveli_studimit"),
    "Niveli_studimit": ("Niveli_studimit", "space_dash", None),
}


class RikodifikimRuleSet:
    def __init__(self, target: str, kind: str, source: str, mapping: dict[str, str]):
        self.target = target
        self.kind = kind
        self.source = source
        self.mapping = mapping


def _parse_mapping_line(line: str, kind: str) -> tuple[str, str] | None:
    if kind == "dash_last":
        if "-" not in line:
            return None
        idx = line.rfind("-")
        return line[:idx], line[idx + 1 :]

    if kind == "space_dash":
        if " - " not in line:
            return None
        old, new = line.split(" - ", 1)
        return old.strip(), new.strip()

    if kind == "tab":
        if "\t" in line:
            old, new = line.split("\t", 1)
            return old.strip(), new.strip()
        match = re.match(r"^(.+?)\s{2,}(.+)$", line.strip())
        if match:
            return match.group(1).strip(), match.group(2).strip()
        return None

    raise ValueError(f"Lloj i panjohur parsing: {kind}")


def ngarko_rregullat(path: Path) -> list[RikodifikimRuleSet]:
    text = path.read_text(encoding="utf-8")
    grouped: dict[str, dict[str, Any]] = {}
    current_target: str | None = None

    for raw in text.splitlines():
        stripped = raw.strip()
        if not stripped:
            continue

        if stripped in SECTION_SPECS:
            target, kind, source = SECTION_SPECS[stripped]
            current_target = target
            grouped.setdefault(
                target,
                {
                    "kind": kind,
                    "source": source or target,
                    "map": {},
                },
            )
            continue

        if stripped == "RRETHI-":
            target, kind, source = SECTION_SPECS["RRETHI"]
            current_target = target
            grouped.setdefault(
                target,
                {
                    "kind": kind,
                    "source": source or target,
                    "map": {},
                },
            )
            continue

        if current_target is None:
            continue

        spec = grouped[current_target]
        parsed = _parse_mapping_line(raw, spec["kind"])
        if parsed is None:
            continue

        old, new = parsed
        if current_target == "RRETHI" and old == "RRETHI" and new == "":
            continue
        spec["map"][old] = new

    # FAKULTETI duhet llogaritur para se të ndryshohet Niveli_studimit.
    order = [
        "GJINIA",
        "RRETHI",
        "SHKOLLA E MESME E KRYER",
        "LLOJI I REGJISTRIMIT",
        "DEGA",
        "VITI",
        "FAKULTETI",
        "Niveli_studimit",
    ]
    rules: list[RikodifikimRuleSet] = []
    for name in order:
        if name not in grouped:
            continue
        spec = grouped[name]
        rules.append(
            RikodifikimRuleSet(
                target=name,
                kind=spec["kind"],
                source=spec["source"],
                mapping=spec["map"],
            )
        )
    return rules


def _normalize_key(value: Any) -> str | None:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if text.lower() in {"", "nan", "none"}:
        return None
    return text


def _apliko_harte(ser: pd.Series, mapping: dict[str, str]) -> pd.Series:
    normalized = ser.map(_normalize_key)
    recoded = normalized.map(mapping)
    return recoded.where(recoded.notna(), normalized)


def rikodifiko_dataframe(df: pd.DataFrame, rules: list[RikodifikimRuleSet], log: logging.Logger) -> pd.DataFrame:
    out = df.copy()

    for rule in rules:
        source_col = rule.source
        target_col = rule.target

        if source_col not in out.columns:
            log.warning("Kolona burim '%s' mungon; seksioni '%s' u anashkalua.", source_col, target_col)
            continue

        before = out[source_col].copy()
        recoded = _apliko_harte(before, rule.mapping)
        out[target_col] = recoded

        changed = (before.map(_normalize_key) != recoded) & recoded.notna()
        log.info(
            "Kolona '%s' ← '%s': %d rreshta u rikodifikuan (%d rregulla).",
            target_col,
            source_col,
            int(changed.sum()),
            len(rule.mapping),
        )

    return out


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Rikodifikon kolonat e Tot_Students sipas rIKODIFIKIM.txt."
    )
    parser.add_argument("--rules", type=Path, default=DEFAULT_RULES, help="Skedari i rregullave")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT, help="Excel hyrës")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="Excel dalës")
    parser.add_argument("--sheet", default=DEFAULT_SHEET, help="Emri i fletës Excel")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(message)s",
    )
    log = logging.getLogger("rikodifikim")

    if not args.rules.is_file():
        log.error("Nuk u gjet skedari i rregullave: %s", args.rules)
        return 1
    if not args.input.is_file():
        log.error("Nuk u gjet skedari hyrës: %s", args.input)
        return 1

    rules = ngarko_rregullat(args.rules)
    if not rules:
        log.error("Nuk u lexua asnjë rregull nga %s", args.rules)
        return 1

    log.info("U ngarkuan %d seksione rikodifikimi.", len(rules))

    try:
        df = pd.read_excel(args.input, sheet_name=args.sheet, engine="openpyxl")
    except ValueError:
        df = pd.read_excel(args.input, sheet_name=0, engine="openpyxl")
        log.warning("Fleta '%s' nuk u gjet; u lexua fleta e parë.", args.sheet)

    log.info("U lexuan %d rreshta, %d kolona.", len(df), len(df.columns))
    df_out = rikodifiko_dataframe(df, rules, log)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    df_out.to_excel(args.output, sheet_name=args.sheet, index=False, engine="openpyxl")
    log.info("Skedari u ruajt: %s", args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
