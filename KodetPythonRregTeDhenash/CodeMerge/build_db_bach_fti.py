"""
Build ResultFiles/Db_Bach_FTI.xlsx from DB_original/FTI_DATABASA E RE 2018-2024.xlsx.

"""

from __future__ import annotations

import argparse
import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

DEFAULT_INPUT = Path("C:\\Users\\esoft\\Desktop\\Projekt-ModeliStudenteve\\FInalWork\\KodetPythonRregTeDhenash\\DB_original") / "FTI_DATABASA E RE 2018-2024.xlsx"
DEFAULT_OUTPUT = Path("C:\\Users\\esoft\\Desktop\\Projekt-ModeliStudenteve\\FInalWork\\KodetPythonRregTeDhenash") / "Db_Bach_FTI.xlsx"

# Output column order (matched fuzzily against the sheet header row).
FIXED_HEADERS: List[str] = [
    "Nr",
    "ID MASH",
    "DATELINDJE",
    "GJINIA",
    "RRETHI",
    "Shkolla e mesme nga vijne",
    "DATE RREGJISTRIMI",
    "LLOJI I REGJISTRIMIT",
    "DATA E DIPLOMES",
    "VITI",
    "DEGA",
]

KURRIKULA_RE = re.compile(r"kurrikula", re.I)
VITI_PAR_RE = re.compile(r"viti\s+i\s+par", re.I)
VITI_DYT_RE = re.compile(r"viti\s+i\s+dyt", re.I)
VITI_TRET_RE = re.compile(r"viti\s+i\s+tret", re.I)
VITI_KAT_RE = re.compile(r"viti\s+i\s+kat", re.I)
SEM_PAR_RE = re.compile(r"semestri\s+par", re.I)
SEM_DYT_RE = re.compile(r"semestri\s+dyt", re.I)


def norm_header(val: Any) -> str:
    if val is None:
        return ""
    t = str(val).lower().replace("\n", " ").strip()
    t = " ".join(t.split())
    for a, b in (("ë", "e"), ("ç", "c")):
        t = t.replace(a, b)
    return t


def find_header_row(ws: Worksheet, max_scan: int = 45) -> Optional[int]:
    for r in range(1, max_scan + 1):
        a = norm_header(ws.cell(r, 1).value).replace(".", "").strip()
        b = norm_header(ws.cell(r, 2).value)
        if a == "nr" and "id" in b and "mas" in b:
            return r
    return None


def effective_value(ws: Worksheet, row: int, col: int) -> Any:
    for m in ws.merged_cells.ranges:
        if m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:
            return ws.cell(m.min_row, m.min_col).value
    return ws.cell(row, col).value


def find_kurrikula_span(ws: Worksheet, max_col: int) -> Optional[Tuple[int, int]]:
    """Return (min_col, max_col) for Kurrikula banner in row 1, if any."""
    for m in ws.merged_cells.ranges:
        if not (m.min_row <= 1 <= m.max_row):
            continue
        v = ws.cell(m.min_row, m.min_col).value
        if v and KURRIKULA_RE.search(str(v)):
            return (m.min_col, m.max_col)
    for c in range(1, max_col + 1):
        v = effective_value(ws, 1, c)
        if v and KURRIKULA_RE.search(str(v)):
            return (c, c)
    return None


def parse_viti_local(s1: str) -> Optional[int]:
    s = s1.lower()
    if VITI_PAR_RE.search(s):
        return 1
    if VITI_DYT_RE.search(s):
        return 2
    if VITI_TRET_RE.search(s):
        return 3
    if VITI_KAT_RE.search(s):
        return 4
    return None


def parse_sem(s2: str) -> Optional[int]:
    s = s2.lower()
    if SEM_DYT_RE.search(s):
        return 2
    if SEM_PAR_RE.search(s):
        return 1
    return None


def detect_year_sem_rows(ws: Worksheet, probe_col: int, kspan: Optional[Tuple[int, int]]) -> Tuple[int, int]:
    """Return (year_row, sem_row) for merged Viti/Sem headers (varies by sheet)."""
    candidates = [(1, 2), (1, 3), (2, 3), (1, 4), (2, 4)]
    for yr_r, sm_r in candidates:
        if yr_r >= sm_r:
            continue
        p = period_for_column(
            probe_col,
            effective_value(ws, yr_r, probe_col),
            effective_value(ws, sm_r, probe_col),
            kspan,
        )
        if p is not None:
            return yr_r, sm_r
    return 1, 2


def period_for_column(
    col: int,
    eff_r1: Any,
    eff_r2: Any,
    kspan: Optional[Tuple[int, int]],
) -> Optional[Tuple[int, int]]:
    if kspan:
        lo, hi = kspan
        if lo <= col <= hi:
            return None
    s1 = str(eff_r1 or "").strip()
    s2 = str(eff_r2 or "").strip()
    if not s1 or not s2:
        return None
    if KURRIKULA_RE.search(s1):
        return None
    vl = parse_viti_local(s1)
    sm = parse_sem(s2)
    if vl is None or sm is None:
        return None
    post = kspan is not None and col > kspan[1]
    if post:
        return (vl + 3, sm)
    return (vl, sm)


def is_zgjedhje_header(cell) -> bool:
    f = cell.fill
    if f.patternType != "solid":
        return False
    fg = f.fgColor
    if fg.type == "theme" and fg.theme == 8:
        return True
    return False


def as_grade(val: Any) -> Optional[float]:
    if val is None or val == "":
        return None
    if isinstance(val, bool):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def resolve_fixed_columns(ws: Worksheet, header_row: int) -> Dict[str, int]:
    """Map display header -> 1-based column index (or -1)."""

    def pick(pred) -> int:
        best_c = -1
        best_score = -1.0
        for c in range(1, ws.max_column + 1):
            h = norm_header(ws.cell(header_row, c).value)
            if not h:
                continue
            if not pred(h):
                continue
            score = 100.0 - len(h) * 0.01
            if score > best_score:
                best_score = score
                best_c = c
        return best_c

    def has_all(h: str, parts: Tuple[str, ...]) -> bool:
        return all(p in h for p in parts)

    out: Dict[str, int] = {}
    out["Nr"] = pick(lambda h: h.rstrip(".") == "nr")
    out["ID MASH"] = pick(lambda h: "id" in h and "mas" in h)
    out["DATELINDJE"] = pick(lambda h: "datelin" in h)
    out["GJINIA"] = pick(lambda h: h == "gjinia")
    out["RRETHI"] = pick(lambda h: "rrethi" in h and "adresa" not in h)
    out["Shkolla e mesme nga vijne"] = pick(lambda h: has_all(h, ("shkolla", "mesme")))
    out["DATE RREGJISTRIMI"] = pick(
        lambda h: ("date" in h or "data" in h)
        and "regjistr" in h
        and "lloji" not in h
        and "nr prog" not in h
        and "nga esht" not in h
        and "vendim" not in h
    )
    out["LLOJI I REGJISTRIMIT"] = pick(
        lambda h: "lloji" in h and "regjistrimit" in h and "te dhena" not in h
    )
    out["DATA E DIPLOMES"] = pick(lambda h: "diplomes" in h or ("data" in h and "diplom" in h))
    out["VITI"] = pick(lambda h: h == "viti")
    out["DEGA"] = pick(lambda h: h == "dega")
    return out


def first_grade_column(fixed_map: Dict[str, int]) -> int:
    dega = fixed_map.get("DEGA", -1)
    if dega > 0:
        return dega + 1
    return 28


def period_sort_key(t: Tuple[int, int]) -> Tuple[int, int]:
    return (t[0], t[1])


def dynamic_headers(all_periods: List[Tuple[int, int]]) -> List[str]:
    out: List[str] = []
    for y, s in sorted(all_periods, key=period_sort_key):
        out.append(f"Baze_{y}-{s}")
        out.append(f"Zgjedhje_{y}-{s}")
    return out


def analyze_sheet(ws: Worksheet, sheet_name: str, log: logging.Logger):
    hr = find_header_row(ws)
    if hr is None:
        log.warning("Sheet %r: no header row (Nr / ID MASH); skipped.", sheet_name)
        return None
    fixed_map = resolve_fixed_columns(ws, hr)
    missing = [k for k, v in fixed_map.items() if v < 0]
    if missing:
        log.warning(
            "Sheet %r: missing fixed columns %s (header row %s).",
            sheet_name,
            ", ".join(missing),
            hr,
        )
    first_g = first_grade_column(fixed_map)
    kspan = find_kurrikula_span(ws, ws.max_column)

    probe_col = None
    for c in range(first_g, min(first_g + 40, ws.max_column + 1)):
        hdr = ws.cell(hr, c).value
        if hdr is not None and str(hdr).strip() != "":
            probe_col = c
            break
    yr_r, sm_r = (1, 2)
    if probe_col is not None:
        yr_r, sm_r = detect_year_sem_rows(ws, probe_col, kspan)

    grade_cols: List[Dict[str, Any]] = []
    unlabeled: List[int] = []

    for c in range(first_g, ws.max_column + 1):
        hdr = ws.cell(hr, c).value
        if hdr is None or str(hdr).strip() == "":
            continue
        p = period_for_column(
            c,
            effective_value(ws, yr_r, c),
            effective_value(ws, sm_r, c),
            kspan,
        )
        if p is None:
            unlabeled.append(c)
            continue
        cell = ws.cell(hr, c)
        grade_cols.append(
            {
                "col": c,
                "period": p,
                "zgjedhje": is_zgjedhje_header(cell),
            }
        )

    if unlabeled:
        from openpyxl.utils import get_column_letter

        sample = ", ".join(get_column_letter(x) for x in unlabeled[:12])
        if len(unlabeled) > 12:
            sample += ", ..."
        log.info(
            "Sheet %r: %d course columns skipped (no Viti/Sem label): %s",
            sheet_name,
            len(unlabeled),
            sample,
        )

    periods_here = sorted({g["period"] for g in grade_cols}, key=period_sort_key)
    return {
        "sheet": sheet_name,
        "header_row": hr,
        "first_data_row": hr + 1,
        "fixed_map": fixed_map,
        "grade_cols": grade_cols,
        "periods": periods_here,
    }


def row_means(
    ws: Worksheet,
    meta: Dict[str, Any],
    r: int,
) -> Dict[Tuple[int, int, str], Optional[float]]:
    """Returns map (y,s,'baze'|'zgjedhje') -> mean."""
    buckets: Dict[Tuple[int, int, str], List[float]] = {}
    for g in meta["grade_cols"]:
        y, s = g["period"]
        kind = "zgjedhje" if g["zgjedhje"] else "baze"
        key = (y, s, kind)
        v = as_grade(ws.cell(r, g["col"]).value)
        if v is None:
            continue
        buckets.setdefault(key, []).append(v)
    out: Dict[Tuple[int, int, str], Optional[float]] = {}
    for k, vals in buckets.items():
        out[k] = sum(vals) / len(vals) if vals else None
    return out


def extract_fixed_values(ws: Worksheet, fixed_map: Dict[str, int], r: int) -> List[Any]:
    row: List[Any] = []
    for h in FIXED_HEADERS:
        col = fixed_map.get(h, -1)
        if col and col > 0:
            row.append(ws.cell(r, col).value)
        else:
            row.append(None)
    return row


def process_workbook(path: Path, log: logging.Logger):
    from openpyxl import load_workbook

    wb = load_workbook(path, data_only=True, read_only=False)
    blocks: List[Tuple[str, Worksheet, Dict[str, Any]]] = []
    all_periods: set = set()
    for name in wb.sheetnames:
        ws = wb[name]
        meta = analyze_sheet(ws, name, log)
        if meta is None:
            continue
        if not meta["grade_cols"]:
            log.info("Sheet %r: no course columns with Viti/Sem; rows still exported.", name)
        blocks.append((name, ws, meta))
        all_periods.update(meta["periods"])

    periods_sorted = sorted(all_periods, key=period_sort_key)
    dyn_headers = dynamic_headers(periods_sorted)
    out_headers = list(FIXED_HEADERS) + dyn_headers

    rows_out: List[List[Any]] = []
    for name, ws, meta in blocks:
        first_r = meta["first_data_row"]
        max_r = ws.max_row
        id_col = meta["fixed_map"].get("ID MASH", 2)
        nr_col = meta["fixed_map"].get("Nr", 1)
        for r in range(first_r, max_r + 1):
            idm = ws.cell(r, id_col).value if id_col > 0 else None
            nr = ws.cell(r, nr_col).value if nr_col > 0 else None
            if (nr is None or str(nr).strip() == "") and (idm is None or str(idm).strip() == ""):
                continue
            fixed_vals = extract_fixed_values(ws, meta["fixed_map"], r)
            means = row_means(ws, meta, r)
            line = list(fixed_vals)
            for y, s in periods_sorted:
                line.append(means.get((y, s, "baze")))
                line.append(means.get((y, s, "zgjedhje")))
            rows_out.append(line)

    wb.close()
    return out_headers, rows_out


def write_output(out_path: Path, headers: List[str], rows: List[List[Any]], log: logging.Logger):
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_wb = Workbook()
    ws = out_wb.active
    assert ws is not None
    ws.title = "Db_Bach_FTI"
    ws.append(headers)
    for row in rows:
        ws.append(row)
    out_wb.save(out_path)
    log.info("Wrote %s (%d rows).", out_path, len(rows))


def main(argv: Optional[List[str]] = None) -> int:
    p = argparse.ArgumentParser(description="Build Db_Bach_FTI.xlsx from FTI workbook.")
    p.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    p.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    p.add_argument("-v", "--verbose", action="store_true")
    args = p.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(message)s",
    )
    log = logging.getLogger("build_db_bach_fti")

    if not args.input.is_file():
        log.error("Input not found: %s", args.input)
        return 1

    headers, rows = process_workbook(args.input, log)
    write_output(args.output, headers, rows, log)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
