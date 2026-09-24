"""
Merge FTI Bach/Master + Biznesi Bachelor/Master + FSHPJ Bachelor + FSHPJ Master
into one workbook (same column-union rules as merge_db_tot_fti_buss.py).
"""

from __future__ import annotations

import argparse
import logging
from pathlib import Path
from typing import List, Optional, Tuple

from merge_db_tot_fti_buss import (
    DEFAULT_BACH_BUSS,
    DEFAULT_BACH_FTI,
    DEFAULT_MASTER_BUSS,
    DEFAULT_MASTER_FTI,
    merge_workbooks,
)

DEFAULT_BACH_FSHPJ = Path("ResultFiles") / "Db_BACH_FSHPJ.xlsx"
DEFAULT_MASTER_FSHPJ = Path("ResultFiles") / "Db_Master_FSHPJ.xlsx"
DEFAULT_OUTPUT = Path("ResultFiles") / "Db_Tot_Fti_Buss_FSHPJ.xlsx"

SHEET_BACH_FSHPJ = "Db_BACH_FSHPJ"
SHEET_MASTER_FSHPJ = "Db_Master_FSHPJ"
SHEET_OUT = "Db_Tot_Fti_Buss_FSHPJ"

SOURCES_6: Tuple[Tuple[str, str], ...] = (
    ("FTI Bachelor", "Db_Bach_FTI"),
    ("FTI Master", "Db_Master_FTI"),
    ("Biznesi Bachelor", "Db_Bachelor_Buss"),
    ("Biznesi Master", "Db_MASTER_Buss"),
    ("FSHPJ Bachelor", SHEET_BACH_FSHPJ),
    ("FSHPJ Master", SHEET_MASTER_FSHPJ),
)


def main(argv: Optional[List[str]] = None) -> int:
    p = argparse.ArgumentParser(
        description="Merge FTI + Biznesi + FSHPJ Bachelor/Master into Db_Tot_Fti_Buss_FSHPJ.xlsx."
    )
    p.add_argument("--bach-fti", type=Path, default=DEFAULT_BACH_FTI, help="Db_Bach_FTI.xlsx")
    p.add_argument("--master-fti", type=Path, default=DEFAULT_MASTER_FTI, help="Db_Master_FTI.xlsx")
    p.add_argument("--bach-buss", type=Path, default=DEFAULT_BACH_BUSS, help="Db_Bachelor_Buss.xlsx")
    p.add_argument("--master-buss", type=Path, default=DEFAULT_MASTER_BUSS, help="Db_MASTER_Buss.xlsx")
    p.add_argument("--bach-fshpj", type=Path, default=DEFAULT_BACH_FSHPJ, help="Db_BACH_FSHPJ.xlsx")
    p.add_argument("--master-fshpj", type=Path, default=DEFAULT_MASTER_FSHPJ, help="Db_Master_FSHPJ.xlsx")
    p.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="Merged output xlsx")
    p.add_argument(
        "--id-column",
        default="Niveli_studimit",
        help="First column name (row source labels)",
    )
    p.add_argument("-v", "--verbose", action="store_true")
    args = p.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(message)s",
    )
    log = logging.getLogger("merge_db_tot_fti_buss_fshpj")

    paths: List[Path] = [
        args.bach_fti,
        args.master_fti,
        args.bach_buss,
        args.master_buss,
        args.bach_fshpj,
        args.master_fshpj,
    ]
    for path in paths:
        if not path.is_file():
            log.error("Input not found: %s", path)
            return 1

    merge_workbooks(SOURCES_6, paths, args.output, SHEET_OUT, args.id_column, log)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
