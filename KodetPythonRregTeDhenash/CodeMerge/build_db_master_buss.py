"""
Build ResultFiles/Db_MASTER_Buss.xlsx from DB_original/BiznesiDatabase Notash Master.xlsx.
"""

from __future__ import annotations

import argparse
import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

DEFAULT_INPUT = Path("C:\\Users\\esoft\\Desktop\\Projekt-ModeliStudenteve\\FInalWork\\KodetPythonRregTeDhenash\\DB_original") / "BiznesiDatabase Notash Master.xlsx"
DEFAULT_OUTPUT = Path("C:\\Users\\esoft\\Desktop\\Projekt-ModeliStudenteve\\FInalWork\\KodetPythonRregTeDhenash") / "Db_MASTER_Buss.xlsx"

# Keys for resolve_fixed_columns / extract_fixed_values (match Excel Albanian headers).
FIXED_COLUMN_KEYS: List[str] = [
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

# Row-1 labels in the output workbook (same order as FIXED_COLUMN_KEYS).
OUTPUT_FIXED_HEADERS: List[str] = [
    "No.",
    "MATRICULATION NO.",
    "DATE OF BIRTH",
    "GENDER",
    "DISTRICT",
    "HIGH SCHOOL COMPLETED",
    "DATE OF REGISTRATION",
    "TYPE OF REGISTRATION",
    "DATE OF DIPLOMA",
    "YEAR",
    "PROFILE",
]

MASTER_TYPE_HEADER = "Master Type"

KURRIKULA_RE = re.compile(r"kurrikula|currikula", re.I)
VITI_PAR_RE = re.compile(r"viti\s*(?:i\s+)?par", re.I)
VITI_DYT_RE = re.compile(r"viti\s*(?:i\s+)?dyt", re.I)
VITI_TRET_RE = re.compile(r"viti\s*(?:i\s+)?tret", re.I)
VITI_KAT_RE = re.compile(r"viti\s*(?:i\s+)?kat", re.I)


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
    """
    Biznesi Master: col A ``Nr``; col B is often ``ID UAMD`` (not ``NR. MATRIKULLIT``).
    """
    for r in range(1, max_scan + 1):
        a = norm_header(effective_value(ws, r, 1)).replace(".", "").strip()
        if a != "nr":
            continue
        b = norm_header(effective_value(ws, r, 2))
        if "matrik" in b:
            return r
        if "id" in b and "uamd" in b:
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


def _nearest_viti_from_band(ws: Worksheet, col: int, header_row: int) -> Optional[int]:
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


def viti_for_column(ws: Worksheet, col: int, header_row: int) -> int:
    blob = meta_blob_for_column(ws, col, header_row)
    joined = norm_header(blob)
    vl = parse_viti_local(joined)
    if KURRIKULA_RE.search(joined) and "/" in blob:
        tail = norm_header(blob.split("/")[-1])
        vl = parse_viti_local(tail) or vl
    if vl is None:
        vl = _nearest_viti_from_band(ws, col, header_row)
    if vl is None:
        vl = 1
    return vl


def detect_lloji_masterit(ws: Worksheet) -> Optional[str]:
    """Best-effort label from banner rows (REGJISTRI / MASTER) or sheet title context."""
    for r in range(1, 10):
        for c in range(1, min(ws.max_column, 12) + 1):
            v = effective_value(ws, r, c)
            if v is None:
                continue
            s = str(v).strip()
            if not s:
                continue
            low = s.lower()
            if "regjistri" in low and ":" in s:
                tail = s.split(":", 1)[-1].strip()
                if tail:
                    return tail[:240]
            if "master" in low and len(s) > 6:
                return s[:240]
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

    out: Dict[str, int] = {}
    out["Nr"] = pick(lambda h: h.rstrip(".") == "nr")
    uamd_col = pick(lambda h: "id" in h and "uamd" in h)
    mat_col = pick(
        lambda h: "matrik" in h and "akp" not in h and "identifikimit" not in h and "qsha" not in h
    )
    out["NR. MATRIKULLIT"] = uamd_col if uamd_col > 0 else mat_col
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


def dynamic_headers_for_years(years_sorted: List[int]) -> List[str]:
    return [f"Base_{y}-1" for y in years_sorted]


def analyze_sheet(ws: Worksheet, sheet_name: str, log: logging.Logger):
    hr = find_header_row(ws)
    if hr is None:
        log.warning("Sheet %r: no header row (Nr + ID UAMD / matrikull); skipped.", sheet_name)
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
    lloji = detect_lloji_masterit(ws) or sheet_name

    grade_cols: List[Dict[str, Any]] = []
    for c in range(first_g, ws.max_column + 1):
        hv = effective_value(ws, hr, c)
        if hv is None or str(hv).strip() == "":
            continue
        year = viti_for_column(ws, c, hr)
        grade_cols.append({"col": c, "year": year})

    years_here = {g["year"] for g in grade_cols}
    return {
        "sheet": sheet_name,
        "header_row": hr,
        "first_data_row": hr + 1,
        "fixed_map": fixed_map,
        "grade_cols": grade_cols,
        "years_here": years_here,
        "lloji_masterit": lloji,
    }


def row_means(ws: Worksheet, meta: Dict[str, Any], r: int) -> Dict[int, Optional[float]]:
    buckets: Dict[int, List[float]] = {}
    for g in meta["grade_cols"]:
        y = g["year"]
        v = as_grade(ws.cell(r, g["col"]).value)
        if v is None:
            continue
        buckets.setdefault(y, []).append(v)
    out: Dict[int, Optional[float]] = {}
    for k, vals in buckets.items():
        out[k] = sum(vals) / len(vals) if vals else None
    return out


def extract_fixed_values(ws: Worksheet, fixed_map: Dict[str, int], r: int) -> List[Any]:
    row: List[Any] = []
    for h in FIXED_COLUMN_KEYS:
        col = fixed_map.get(h, -1)
        if col and col > 0:
            row.append(ws.cell(r, col).value)
        else:
            row.append(None)
    return row


def process_workbook(path: Path, log: logging.Logger):
    from openpyxl import load_workbook

    wb = load_workbook(path, data_only=True, read_only=False)
    metas: List[Tuple[str, Worksheet, Dict[str, Any]]] = []
    all_years: Set[int] = set()

    for name in wb.sheetnames:
        ws = wb[name]
        meta = analyze_sheet(ws, name, log)
        if meta is None:
            continue
        all_years.update(meta["years_here"])
        metas.append((name, ws, meta))

    if not metas:
        wb.close()
        log.error("No sheets had a detectable header row (Nr + ID UAMD or matrikull).")
        return list(OUTPUT_FIXED_HEADERS) + [MASTER_TYPE_HEADER], []

    years_sorted = sorted(all_years)
    dyn_headers = dynamic_headers_for_years(years_sorted)
    out_headers = list(OUTPUT_FIXED_HEADERS) + [MASTER_TYPE_HEADER] + dyn_headers

    rows_out: List[List[Any]] = []
    for _name, ws, meta in metas:
        first_r = meta["first_data_row"]
        max_r = ws.max_row
        nr_col = meta["fixed_map"].get("Nr", 1)
        mat_col = meta["fixed_map"].get("NR. MATRIKULLIT", 2)
        lloji = meta.get("lloji_masterit")
        if not meta["grade_cols"]:
            log.info("Sheet %r: no mappable grade columns; rows still exported.", _name)
        for r in range(first_r, max_r + 1):
            nr = ws.cell(r, nr_col).value if nr_col > 0 else None
            mat = ws.cell(r, mat_col).value if mat_col > 0 else None
            if (nr is None or str(nr).strip() == "") and (mat is None or str(mat).strip() == ""):
                continue
            fixed_vals = extract_fixed_values(ws, meta["fixed_map"], r)
            means = row_means(ws, meta, r)
            line = list(fixed_vals) + [lloji]
            for y in years_sorted:
                line.append(means.get(y))
            rows_out.append(line)

    wb.close()
    return out_headers, rows_out


def write_output(out_path: Path, headers: List[str], rows: List[List[Any]], log: logging.Logger):
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_wb = Workbook()
    ws = out_wb.active
    assert ws is not None
    ws.title = "Db_MASTER_Buss"
    ws.append(headers)
    for row in rows:
        ws.append(row)
    out_wb.save(out_path)
    log.info("Wrote %s (%d rows).", out_path, len(rows))


def main(argv: Optional[List[str]] = None) -> int:
    p = argparse.ArgumentParser(description="Build Db_MASTER_Buss.xlsx from Biznesi Master workbook.")
    p.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    p.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    p.add_argument("-v", "--verbose", action="store_true")
    args = p.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(message)s",
    )
    log = logging.getLogger("build_db_master_buss")

    if not args.input.is_file():
        log.error("Input not found: %s", args.input)
        return 1

    headers, rows = process_workbook(args.input, log)
    if not rows:
        log.error("No student rows exported. Check that the workbook matches the Biznesi Master layout.")
        return 1
    write_output(args.output, headers, rows, log)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
