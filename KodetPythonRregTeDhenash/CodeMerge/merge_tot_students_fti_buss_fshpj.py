"""
Merge Db_BACH_FSHPJ, Db_Bachelor_Buss, Db_MASTER_Buss, Db_Bach_FTI, Db_Master_FTI
into Tot_Students_FTI_BUSS_FSHPJ.xlsx.

Rules (same as merge_db_tot_fti_buss.merge_workbooks):
- One column per distinct header name (first occurrence sets order; later files add only new names).
- All rows from each file are appended; cells are blank where that file has no such column.
- First column (default Niveli_studimit) records which source each row came from.
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
DEFAULT_OUTPUT = Path("ResultFiles") / "Tot_Students_FTI_BUSS_FSHPJ.xlsx"
SHEET_OUT = "Tot_Students_FTI_BUSS_FSHPJ"

# Order: FSHPJ Bachelor, Biznesi Bach/Master, FTI Bach/Master — (label, sheet name in workbook)
SOURCES_5: Tuple[Tuple[str, str], ...] = (
    ("FSHPJ Bachelor", "Db_BACH_FSHPJ"),
    ("Biznesi Bachelor", "Db_Bachelor_Buss"),
    ("Biznesi Master", "Db_MASTER_Buss"),
    ("FTI Bachelor", "Db_Bach_FTI"),
    ("FTI Master", "Db_Master_FTI"),
)


def main(argv: Optional[List[str]] = None) -> int:
    p = argparse.ArgumentParser(
        description="Merge FSHPJ Bachelor + Biznesi + FTI into Tot_Students_FTI_BUSS_FSHPJ.xlsx."
    )
    p.add_argument("--bach-fshpj", type=Path, default=DEFAULT_BACH_FSHPJ, help="Db_BACH_FSHPJ.xlsx")
    p.add_argument("--bach-buss", type=Path, default=DEFAULT_BACH_BUSS, help="Db_Bachelor_Buss.xlsx")
    p.add_argument("--master-buss", type=Path, default=DEFAULT_MASTER_BUSS, help="Db_MASTER_Buss.xlsx")
    p.add_argument("--bach-fti", type=Path, default=DEFAULT_BACH_FTI, help="Db_Bach_FTI.xlsx")
    p.add_argument("--master-fti", type=Path, default=DEFAULT_MASTER_FTI, help="Db_Master_FTI.xlsx")
    p.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="Merged output xlsx")
    p.add_argument(
        "--id-column",
        default="Niveli_studimit",
        help="First column name (source label per row)",
    )
    p.add_argument("-v", "--verbose", action="store_true")
    args = p.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(message)s",
    )
    log = logging.getLogger("merge_tot_students_fti_buss_fshpj")

    paths: List[Path] = [
        args.bach_fshpj,
        args.bach_buss,
        args.master_buss,
        args.bach_fti,
        args.master_fti,
    ]
    for path in paths:
        if not path.is_file():
            log.error("Input not found: %s", path)
            return 1

    merge_workbooks(SOURCES_5, paths, args.output, SHEET_OUT, args.id_column, log)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
