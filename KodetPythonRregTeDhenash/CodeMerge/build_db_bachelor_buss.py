"""
Build ResultFiles/Db_Bachelor_Buss.xlsx from DB_original/BiznesiDatabase Notash Bachelor.xlsx.

"""

from __future__ import annotations

import argparse
import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

DEFAULT_INPUT = Path("C:\\Users\\esoft\\Desktop\\Projekt-ModeliStudenteve\\FInalWork\\KodetPythonRregTeDhenash\\DB_original") / "BiznesiDatabase Notash Bachelor.xlsx"
DEFAULT_OUTPUT = Path("C:\\Users\\esoft\\Desktop\\Projekt-ModeliStudenteve\\FInalWork\\KodetPythonRregTeDhenash") / "Db_Bachelor_Buss.xlsx"

FIXED_HEADERS: List[str] = [
    "Nr",
    "NR. MATRIKULLIT",
    "DATELINDJA",
    "GJINIA",
    "RRETHI",
    "SHKOLLA E MESME E KRYER",
    "DATA E REGJISTRIMIT",
    "LLOJI I REGJISTRIMIT",
    "DATA E DIPLOMES",
    "VITI",
    "PROFILI",
]

KURRIKULA_RE = re.compile(r"kurrikula|currikula", re.I)
VITI_PAR_RE = re.compile(r"viti\s*(?:i\s+)?par", re.I)
VITI_DYT_RE = re.compile(r"viti\s*(?:i\s+)?dyt", re.I)
VITI_TRET_RE = re.compile(r"viti\s*(?:i\s+)?tret", re.I)
VITI_KAT_RE = re.compile(r"viti\s*(?:i\s+)?kat", re.I)
SEM_PAR_RE = re.compile(r"semestri(?:\s+i)?\s+par", re.I)
SEM_DYT_RE = re.compile(r"semestri(?:\s+i)?\s+dyt", re.I)


def norm_header(val: Any) -> str:
    if val is None:
        return ""
    t = str(val).lower().replace("\n", " ").strip()
    t = " ".join(t.split())
    for a, b in (("ë", "e"), ("ç", "c")):
        t = t.replace(a, b)
    return t


def effective_value(ws: Worksheet, row: int, col: int) -> Any:
    for m in ws.merged_cells.ranges:
        if m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:
            return ws.cell(m.min_row, m.min_col).value
    return ws.cell(row, col).value


def find_header_row(ws: Worksheet, max_scan: int = 35) -> Optional[int]:
    """Row where col A is Nr and col B (merged-aware) mentions matrikull."""
    for r in range(1, max_scan + 1):
        a = norm_header(effective_value(ws, r, 1)).replace(".", "").strip()
        if a != "nr":
            continue
        b = norm_header(effective_value(ws, r, 2))
        if "matrik" in b:
            return r
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


def meta_blob_for_column(ws: Worksheet, col: int, header_row: int) -> str:
    parts: List[str] = []
    for r in range(1, header_row):
        v = effective_value(ws, r, col)
        if v is None:
            continue
        s = str(v).strip()
        if not s:
            continue
        parts.append(s)
    return "\n".join(parts)


def period_from_meta_blob(blob: str) -> Optional[Tuple[int, int]]:
    """
    Infer (viti, sem) from all merged labels above the header row for this column.
    Handles Biznesi stacks where year and semester appear on different rows.
    """
    if not blob.strip():
        return None
    joined = " ".join(norm_header(x) for x in blob.splitlines())
    if KURRIKULA_RE.search(joined) and "/" in joined:
        tail = joined.split("/", 1)[-1].strip()
        joined_v = tail or joined
    else:
        joined_v = joined

    vl = parse_viti_local(joined_v)
    sm = parse_sem(joined)
    if vl is not None and sm is not None:
        return (vl, sm)

    if parse_viti_local(joined) is not None and parse_sem(joined) is not None:
        return (parse_viti_local(joined), parse_sem(joined))

    if parse_sem(joined) is not None and parse_viti_local(joined) is None:
        sm2 = parse_sem(joined)
        vl2 = parse_viti_local(joined)
        if vl2 is None:
            vl2 = 1
        if sm2 is not None:
            return (vl2, sm2)

    lines = [norm_header(x) for x in blob.splitlines() if str(x).strip()]
    vl = None
    sm = None
    for line in lines:
        v = parse_viti_local(line)
        if v is not None:
            vl = v
        s = parse_sem(line)
        if s is not None:
            sm = s
    if vl is not None and sm is not None:
        return (vl, sm)
    if sm is not None and vl is None:
        return (1, sm)
    return None


def _nearest_sem_from_band(ws: Worksheet, col: int, header_row: int) -> Optional[int]:
    """Some columns sit in horizontal gaps between row-8 semester merges; scan row 8."""
    if header_row <= 8:
        return None
    r = 8
    span = 160
    for cc in range(col, max(0, col - span), -1):
        if cc < 1:
            break
        v = effective_value(ws, r, cc)
        if not v:
            continue
        sm = parse_sem(str(v))
        if sm is not None:
            return sm
    for cc in range(col + 1, min(ws.max_column + 1, col + span + 1)):
        v = effective_value(ws, r, cc)
        if not v:
            continue
        sm = parse_sem(str(v))
        if sm is not None:
            return sm
    return None


def _nearest_viti_from_band(ws: Worksheet, col: int, header_row: int) -> Optional[int]:
    """Scan meta rows 5–7 for the nearest merged Viti label (prefer same column, then left)."""
    top = min(header_row - 1, 7)
    for r in range(top, 4, -1):
        if r < 1:
            break
        for cc in range(col, 0, -1):
            v = effective_value(ws, r, cc)
            if not v:
                continue
            vl = parse_viti_local(str(v))
            if vl is not None:
                return vl
        for cc in range(col + 1, min(ws.max_column + 1, col + 200)):
            v = effective_value(ws, r, cc)
            if not v:
                continue
            vl = parse_viti_local(str(v))
            if vl is not None:
                return vl
    for r in range(4, 0, -1):
        for cc in range(col, 0, -1):
            v = effective_value(ws, r, cc)
            if not v:
                continue
            vl = parse_viti_local(str(v))
            if vl is not None:
                return vl
    return None


def period_for_biznesi_column(ws: Worksheet, col: int, header_row: int) -> Optional[Tuple[int, int]]:
    blob = meta_blob_for_column(ws, col, header_row)
    p = period_from_meta_blob(blob)
    if p is not None:
        return p
    joined = norm_header(blob)
    vl = parse_viti_local(joined)
    sm = parse_sem(joined)
    if vl is None:
        vl = _nearest_viti_from_band(ws, col, header_row)
    if sm is None:
        sm = _nearest_sem_from_band(ws, col, header_row)
    if vl is not None and sm is not None:
        return (vl, sm)
    if sm is not None and vl is None:
        return (1, sm)
    return None


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
    def hdr(r: int, c: int) -> str:
        return norm_header(effective_value(ws, r, c))

    def pick(pred) -> int:
        best_c = -1
        best_score = -1.0
        for c in range(1, ws.max_column + 1):
            h = hdr(header_row, c)
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
    out["NR. MATRIKULLIT"] = pick(
        lambda h: "matrik" in h and "akp" not in h and "identifikimit" not in h
    )
    out["DATELINDJA"] = pick(lambda h: "datelin" in h)
    out["GJINIA"] = pick(lambda h: h == "gjinia")
    out["RRETHI"] = pick(lambda h: "rrethi" in h and "adresa" not in h)
    out["SHKOLLA E MESME E KRYER"] = pick(
        lambda h: ("shkolla" in h and "mesme" in h) or ("shkolla" in h and "kryer" in h)
    )
    out["DATA E REGJISTRIMIT"] = pick(
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
    out["PROFILI"] = pick(lambda h: "profil" in h and "dega" not in h)
    return out


def first_grade_column(fixed_map: Dict[str, int]) -> int:
    prof = fixed_map.get("PROFILI", -1)
    if prof and prof > 0:
        return prof + 1
    return 27


def period_sort_key(t: Tuple[int, int]) -> Tuple[int, int]:
    return (t[0], t[1])


def dynamic_headers(all_periods: List[Tuple[int, int]]) -> List[str]:
    out: List[str] = []
    for y, s in sorted(all_periods, key=period_sort_key):
        out.append(f"Baze_{y}-{s}")
    return out


def analyze_sheet(ws: Worksheet, sheet_name: str, log: logging.Logger):
    hr = find_header_row(ws)
    if hr is None:
        log.warning("Sheet %r: no header row (Nr / NR. MATRIKULLIT); skipped.", sheet_name)
        return None

    fixed_map = resolve_fixed_columns(ws, hr)
    warn_keys = ("LLOJI I REGJISTRIMIT", "DATA E DIPLOMES", "VITI")
    for k in warn_keys:
        if fixed_map.get(k, -1) < 0:
            log.warning("Sheet %r: missing optional column %r - output will be blank.", sheet_name, k)

    missing_core = [k for k in ("Nr", "NR. MATRIKULLIT") if fixed_map.get(k, -1) < 0]
    if missing_core:
        log.warning(
            "Sheet %r: missing core columns %s; sheet may be unusable.",
            sheet_name,
            ", ".join(missing_core),
        )

    first_g = first_grade_column(fixed_map)

    grade_cols: List[Dict[str, Any]] = []
    unlabeled: List[int] = []

    for c in range(first_g, ws.max_column + 1):
        hv = effective_value(ws, hr, c)
        if hv is None or str(hv).strip() == "":
            continue
        p = period_for_biznesi_column(ws, c, hr)
        if p is None:
            unlabeled.append(c)
            continue
        grade_cols.append({"col": c, "period": p})

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


def row_means(ws: Worksheet, meta: Dict[str, Any], r: int) -> Dict[Tuple[int, int], Optional[float]]:
    buckets: Dict[Tuple[int, int], List[float]] = {}
    for g in meta["grade_cols"]:
        y, s = g["period"]
        v = as_grade(ws.cell(r, g["col"]).value)
        if v is None:
            continue
        buckets.setdefault((y, s), []).append(v)
    out: Dict[Tuple[int, int], Optional[float]] = {}
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
    for _name, ws, meta in blocks:
        first_r = meta["first_data_row"]
        max_r = ws.max_row
        nr_col = meta["fixed_map"].get("Nr", 1)
        mat_col = meta["fixed_map"].get("NR. MATRIKULLIT", 2)
        for r in range(first_r, max_r + 1):
            nr = ws.cell(r, nr_col).value if nr_col > 0 else None
            mat = ws.cell(r, mat_col).value if mat_col > 0 else None
            if (nr is None or str(nr).strip() == "") and (mat is None or str(mat).strip() == ""):
                continue
            fixed_vals = extract_fixed_values(ws, meta["fixed_map"], r)
            means = row_means(ws, meta, r)
            line = list(fixed_vals)
            for y, s in periods_sorted:
                line.append(means.get((y, s)))
            rows_out.append(line)

    wb.close()
    return out_headers, rows_out


def write_output(out_path: Path, headers: List[str], rows: List[List[Any]], log: logging.Logger):
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_wb = Workbook()
    ws = out_wb.active
    assert ws is not None
    ws.title = "Db_Bachelor_Buss"
    ws.append(headers)
    for row in rows:
        ws.append(row)
    out_wb.save(out_path)
    log.info("Wrote %s (%d rows).", out_path, len(rows))


def main(argv: Optional[List[str]] = None) -> int:
    p = argparse.ArgumentParser(description="Build Db_Bachelor_Buss.xlsx from Biznesi workbook.")
    p.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    p.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    p.add_argument("-v", "--verbose", action="store_true")
    args = p.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(message)s",
    )
    log = logging.getLogger("build_db_bachelor_buss")

    if not args.input.is_file():
        log.error("Input not found: %s", args.input)
        return 1

    headers, rows = process_workbook(args.input, log)
    write_output(args.output, headers, rows, log)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
