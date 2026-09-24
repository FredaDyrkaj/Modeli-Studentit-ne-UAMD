"""
Build ResultFiles/Db_Master_FSHPJ.xlsx from DB_original/FSHPJ_Databaza master.xlsx.

"""

from __future__ import annotations

import argparse
import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

DEFAULT_INPUT = Path("C:\\Users\\esoft\\Desktop\\Projekt-ModeliStudenteve\\FInalWork\\KodetPythonRregTeDhenash\\DB_original") / "FSHPJ_Databaza master.xlsx"
DEFAULT_OUTPUT = Path("C:\\Users\\esoft\\Desktop\\Projekt-ModeliStudenteve\\FInalWork\\KodetPythonRregTeDhenash") / "Db_Master_FSHPJ.xlsx"

# Output header strings (must match fixed_map keys in resolve_fixed_columns).
SCHOOL_HEADER = "Emri i Shkollës së Mesme ku ka kryer studimet parauniversitare"

FIXED_HEADERS: List[str] = [
    "Nr",
    "ID MASH",
    "DATELINDJA",
    "GJINIA",
    "RRETHI",
    SCHOOL_HEADER,
    "DATA E REGJISTRIMIT",
    "LLOJI I REGJISTRIMIT",
    "DATA E DIPLOMES",
    "VITI",
    "DEGA",
]

NIVELI_STUDIMIT_HEADER = "Niveli_studimit"
NIVELI_STUDIMIT_VALUE = "FSHPJ Master"

# Master output: only years 1–2, semesters 1–2 (fixed column order in the workbook).
MASTER_OUTPUT_PERIODS: List[Tuple[int, int]] = [(1, 1), (1, 2), (2, 1), (2, 2)]
MASTER_OUTPUT_PERIODS_SET = frozenset(MASTER_OUTPUT_PERIODS)

KURRIKULA_RE = re.compile(r"kurrikula", re.I)
VITI_PAR_RE = re.compile(r"viti\s*(?:i\s+)?par", re.I)
VITI_DYT_RE = re.compile(r"viti\s*(?:i\s+)?dyt", re.I)
VITI_TRET_RE = re.compile(r"viti\s*(?:i\s+)?tret", re.I)
VITI_KAT_RE = re.compile(r"viti\s*(?:i\s+)?kat", re.I)
VITI_PEST_RE = re.compile(r"viti\s*(?:i\s+)?pest", re.I)
SEM_PAR_RE = re.compile(r"semestri(?:\s+i)?\s+par", re.I)
SEM_DYT_RE = re.compile(r"semestri(?:\s+i)?\s+dyt", re.I)
SEM_PAR_LOOSE_RE = re.compile(r"semestri(?:\s+i)?\s*$", re.I)
SEM_ROMAN_II_RE = re.compile(r"semestri.*\bii\b", re.I)
ZGJEDHJE_META_RE = re.compile(r"l\.?\s*me\s+zgjedhje|me\s+zgjedhje", re.I)
BASE_META_RE = re.compile(r"lende.*det|det.*lende|detyrueshme|detyruesh", re.I)


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


def is_nr_header_cell(val: Any) -> bool:
    if val is None:
        return False
    t = norm_header(val).replace(".", " ").strip()
    if t == "nr":
        return True
    if t.startswith("nr") and "rendor" in t:
        return True
    return False


def find_header_row(ws: Worksheet, max_scan: int = 60) -> Optional[int]:
    max_c_probe = min(ws.max_column, 45)
    for r in range(1, max_scan + 1):
        if not is_nr_header_cell(effective_value(ws, r, 1)):
            continue
        found_id = False
        for c in range(2, max_c_probe + 1):
            h = norm_header(effective_value(ws, r, c))
            if "id" in h and "mas" in h:
                found_id = True
                break
        if found_id:
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
    if VITI_PEST_RE.search(s):
        return 5
    return None


def parse_sem(s2: str) -> Optional[int]:
    s = s2.lower()
    if SEM_DYT_RE.search(s):
        return 2
    if SEM_ROMAN_II_RE.search(s):
        return 2
    if SEM_PAR_RE.search(s):
        return 1
    if SEM_PAR_LOOSE_RE.search(s) and "dyt" not in s:
        return 1
    return None


def find_kurrikula_span(ws: Worksheet, max_col: int) -> Optional[Tuple[int, int]]:
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
    n1 = norm_header(s1)
    n2 = norm_header(s2)
    if KURRIKULA_RE.search(n1):
        return None
    vl = parse_viti_local(n1)
    sm = parse_sem(n2)
    if vl is None or sm is None:
        return None
    post = kspan is not None and col > kspan[1]
    if post:
        return (vl + 3, sm)
    return (vl, sm)


def _cell_nonempty(val: Any) -> bool:
    return val is not None and str(val).strip() != ""


def semester_label_borrow_left(ws: Worksheet, col: int, sm_r: int, max_left: int = 45) -> Any:
    s = effective_value(ws, sm_r, col)
    if _cell_nonempty(s):
        return s
    lo = max(1, col - max_left)
    for c2 in range(col - 1, lo - 1, -1):
        sb = effective_value(ws, sm_r, c2)
        if _cell_nonempty(sb):
            return sb
    return s


def row_period_for_column(
    ws: Worksheet,
    col: int,
    yr_r: int,
    sm_r: int,
    kspan: Optional[Tuple[int, int]],
) -> Optional[Tuple[int, int]]:
    s1 = effective_value(ws, yr_r, col)
    s2 = semester_label_borrow_left(ws, col, sm_r)
    return period_for_column(col, s1, s2, kspan)


def resolve_period_for_grade_column(
    ws: Worksheet,
    col: int,
    default_yr_r: int,
    default_sm_r: int,
    kspan: Optional[Tuple[int, int]],
    meta_scan_ub: int,
) -> Optional[Tuple[int, int]]:
    """meta_scan_ub: last row index that may carry Viti/Semestri labels (below subject title row)."""
    p = row_period_for_column(ws, col, default_yr_r, default_sm_r, kspan)
    if p is not None:
        return p
    ub = max(1, meta_scan_ub)
    for yr_r in range(1, ub):
        for sm_r in range(yr_r + 1, ub + 1):
            p = row_period_for_column(ws, col, yr_r, sm_r, kspan)
            if p is not None:
                return p
    return None


def detect_year_sem_rows(
    ws: Worksheet,
    probe_col: int,
    kspan: Optional[Tuple[int, int]],
    meta_scan_ub: int,
) -> Tuple[int, int]:
    ub = max(1, meta_scan_ub)
    candidates = [(1, 2), (1, 3), (2, 3), (1, 4), (2, 4)]
    for yr_r, sm_r in candidates:
        if yr_r >= sm_r or sm_r > ub:
            continue
        p = row_period_for_column(ws, probe_col, yr_r, sm_r, kspan)
        if p is not None:
            return yr_r, sm_r
    for yr_r in range(1, ub):
        for sm_r in range(yr_r + 1, ub + 1):
            p = row_period_for_column(ws, probe_col, yr_r, sm_r, kspan)
            if p is not None:
                return yr_r, sm_r
    return 1, 2


def is_zgjedhje_header(cell) -> bool:
    f = cell.fill
    if f.patternType != "solid":
        return False
    fg = f.fgColor
    if fg.type == "theme" and fg.theme == 8:
        return True
    return False


def elective_meta_hint(ws: Worksheet, col: int, scan_until_row: int) -> bool:
    parts: List[str] = []
    for r in range(1, max(1, scan_until_row)):
        v = effective_value(ws, r, col)
        if v is None or str(v).strip() == "":
            continue
        parts.append(norm_header(str(v)))
    blob = " ".join(parts)
    return bool(ZGJEDHJE_META_RE.search(blob))


def as_grade(val: Any) -> Optional[float]:
    if val is None or val == "":
        return None
    if isinstance(val, bool):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(",", ".").rstrip("*").strip()
    try:
        return float(s)
    except ValueError:
        return None


def resolve_fixed_columns(ws: Worksheet, header_row: int) -> Dict[str, int]:
    def hdr(c: int) -> str:
        return norm_header(effective_value(ws, header_row, c))

    def pick(pred) -> int:
        best_c = -1
        best_score = -1.0
        for c in range(1, ws.max_column + 1):
            h = hdr(c)
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
    out["Nr"] = pick(lambda h: h.rstrip(".") == "nr" or ("nr" in h and "rendor" in h))
    out["ID MASH"] = pick(lambda h: "id" in h and "mas" in h)
    out["DATELINDJA"] = pick(lambda h: "datelin" in h)
    out["GJINIA"] = pick(lambda h: h == "gjinia")
    out["RRETHI"] = pick(lambda h: "rrethi" in h and "adresa" not in h and "bashkia" not in h)
    out[SCHOOL_HEADER] = pick(
        lambda h: ("parauniversitar" in h or "parauniversitare" in h)
        or (has_all(h, ("shkoll", "mesme")) and ("kryer" in h or "vijne" in h or "vij" in h))
        or has_all(h, ("shkoll", "mesme", "kryer"))
        or ("shkolla" in h and "mesme" in h)
    )
    out["DATA E REGJISTRIMIT"] = pick(
        lambda h: ("date" in h or "data" in h)
        and "regjistr" in h
        and "lloji" not in h
        and "nr prog" not in h
        and "nga esht" not in h
        and "vendim" not in h
        and "programin" not in h
    )
    out["LLOJI I REGJISTRIMIT"] = pick(
        lambda h: "lloji" in h and "regjistrimit" in h and "te dhena" not in h
    )
    out["DATA E DIPLOMES"] = pick(lambda h: "diplomes" in h or ("data" in h and "diplom" in h))
    out["VITI"] = pick(
        lambda h: h == "viti"
        or h == "viti akademik"
        or ("viti" in h and "akademik" in h)
    )
    out["DEGA"] = pick(
        lambda h: h == "dega"
        or h.startswith("dega ")
        or h.startswith("dege ")
        or (h.startswith("dega/") and len(h) < 40)
    )
    return out


def merge_fixed_column_maps(primary: Dict[str, int], secondary: Dict[str, int]) -> Dict[str, int]:
    out = dict(primary)
    for k, c in secondary.items():
        if out.get(k, -1) < 0 and c > 0:
            out[k] = c
    return out


def find_first_student_data_row(ws: Worksheet, identity_header_row: int, id_col: int, max_scan: int = 120) -> int:
    if id_col < 1:
        return identity_header_row + 1
    last = min(ws.max_row, identity_header_row + max_scan)
    for r in range(identity_header_row + 1, last + 1):
        v = ws.cell(r, id_col).value
        if v is None or str(v).strip() == "":
            continue
        if norm_header(str(v)) == "id mash":
            continue
        return r
    return identity_header_row + 1


def viti_label_borrow_left(ws: Worksheet, col: int, meta_r: int, max_left: int = 120) -> Any:
    s = effective_value(ws, meta_r, col)
    if _cell_nonempty(s):
        return s
    lo = max(1, col - max_left)
    for c2 in range(col - 1, lo - 1, -1):
        sb = effective_value(ws, meta_r, c2)
        if _cell_nonempty(sb):
            return sb
    return s


def infer_period_master_year_sem_split(
    ws: Worksheet,
    col: int,
    meta_scan_ub: int,
    kspan: Optional[Tuple[int, int]],
) -> Optional[Tuple[int, int]]:
    """When only a Viti row exists (no Semestri), split contiguous columns into sem 1 / sem 2."""
    if kspan:
        lo, hi = kspan
        if lo <= col <= hi:
            return None
    if meta_scan_ub < 1:
        return None
    meta_r = None
    for r in range(meta_scan_ub, 0, -1):
        s1 = str(viti_label_borrow_left(ws, col, r) or "").strip()
        if parse_viti_local(norm_header(s1)) is not None:
            meta_r = r
            break
    if meta_r is None:
        return None
    s1 = str(viti_label_borrow_left(ws, col, meta_r) or "").strip()
    n1 = norm_header(s1)
    vl = parse_viti_local(n1)
    if vl is None or KURRIKULA_RE.search(n1):
        return None
    c_lo = col
    while c_lo > 1:
        s_prev = str(viti_label_borrow_left(ws, c_lo - 1, meta_r) or "").strip()
        if parse_viti_local(norm_header(s_prev)) != vl:
            break
        c_lo -= 1
    c_hi = col
    max_c = ws.max_column
    while c_hi < max_c:
        s_next = str(viti_label_borrow_left(ws, c_hi + 1, meta_r) or "").strip()
        if parse_viti_local(norm_header(s_next)) != vl:
            break
        c_hi += 1
    if c_hi == c_lo:
        return (vl, 1)
    mid = c_lo + (c_hi - c_lo + 1) // 2
    sm = 1 if col < mid else 2
    return (vl, sm)


def classify_zgjedhje_column(ws: Worksheet, col: int, identity_hr: int, subject_row: int) -> bool:
    for r in range(identity_hr + 1, subject_row):
        raw = effective_value(ws, r, col)
        t = norm_header(str(raw or ""))
        if BASE_META_RE.search(t):
            return False
        if ZGJEDHJE_META_RE.search(t):
            return True
        cell = ws.cell(r, col)
        if is_zgjedhje_header(cell):
            return True
    return elective_meta_hint(ws, col, subject_row)


def first_grade_column(fixed_map: Dict[str, int]) -> int:
    pos = [c for c in fixed_map.values() if isinstance(c, int) and c > 0]
    if pos:
        return max(pos) + 1
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
        log.warning("Sheet %r: no header row (Nr + ID MASH); skipped.", sheet_name)
        return None

    fixed_map = resolve_fixed_columns(ws, hr)
    id_col = fixed_map.get("ID MASH", -1)
    first_data_row = find_first_student_data_row(ws, hr, id_col)
    subject_row = (first_data_row - 1) if first_data_row > hr + 1 else hr
    if subject_row > hr:
        fixed_map = merge_fixed_column_maps(fixed_map, resolve_fixed_columns(ws, subject_row))

    warn_keys = ("LLOJI I REGJISTRIMIT", "DATA E DIPLOMES", "VITI")
    for k in warn_keys:
        if fixed_map.get(k, -1) < 0:
            log.warning("Sheet %r: missing optional column %r - output will be blank.", sheet_name, k)

    missing_core = [k for k in ("ID MASH",) if fixed_map.get(k, -1) < 0]
    if missing_core:
        log.warning(
            "Sheet %r: missing core columns %s; sheet may be unusable.",
            sheet_name,
            ", ".join(missing_core),
        )

    first_g = first_grade_column(fixed_map)
    kspan = find_kurrikula_span(ws, ws.max_column)
    meta_scan_ub = max(hr - 1, subject_row - 1)

    probe_col = None
    for c in range(first_g, min(first_g + 120, ws.max_column + 1)):
        hdr_sub = effective_value(ws, subject_row, c)
        if hdr_sub is None or str(hdr_sub).strip() == "":
            continue
        probe_col = c
        break
    if probe_col is None:
        for c in range(first_g, min(first_g + 50, ws.max_column + 1)):
            hdr_top = effective_value(ws, hr, c)
            if hdr_top is not None and str(hdr_top).strip() != "":
                probe_col = c
                break

    yr_r, sm_r = (1, 2)
    if probe_col is not None:
        yr_r, sm_r = detect_year_sem_rows(ws, probe_col, kspan, meta_scan_ub)

    grade_cols: List[Dict[str, Any]] = []
    unlabeled: List[int] = []

    for c in range(first_g, ws.max_column + 1):
        hdr = effective_value(ws, subject_row, c)
        if hdr is None or str(hdr).strip() == "":
            continue
        p = resolve_period_for_grade_column(ws, c, yr_r, sm_r, kspan, meta_scan_ub)
        if p is None:
            p = infer_period_master_year_sem_split(ws, c, meta_scan_ub, kspan)
        if p is None:
            unlabeled.append(c)
            continue
        zgj = classify_zgjedhje_column(ws, c, hr, subject_row)
        grade_cols.append(
            {
                "col": c,
                "period": p,
                "zgjedhje": zgj,
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
        "subject_row": subject_row,
        "first_data_row": first_data_row,
        "fixed_map": fixed_map,
        "grade_cols": grade_cols,
        "periods": periods_here,
    }


def row_means(
    ws: Worksheet,
    meta: Dict[str, Any],
    r: int,
) -> Dict[Tuple[int, int, str], Optional[float]]:
    buckets: Dict[Tuple[int, int, str], List[float]] = {}
    for g in meta["grade_cols"]:
        y, s = g["period"]
        kind = "zgjedhje" if g["zgjedhje"] else "baze"
        v = as_grade(ws.cell(r, g["col"]).value)
        if v is None:
            continue
        buckets.setdefault((y, s, kind), []).append(v)
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

    dropped_periods = sorted(
        (p for p in all_periods if p not in MASTER_OUTPUT_PERIODS_SET),
        key=period_sort_key,
    )
    if dropped_periods:
        log.info(
            "Some sheets map grades to periods not exported (Master output is years 1–2 only): %s",
            dropped_periods,
        )

    periods_sorted = list(MASTER_OUTPUT_PERIODS)
    dyn_headers = dynamic_headers(periods_sorted)
    out_headers = list(FIXED_HEADERS) + [NIVELI_STUDIMIT_HEADER] + dyn_headers

    rows_out: List[List[Any]] = []
    for name, ws, meta in blocks:
        first_r = meta["first_data_row"]
        max_r = ws.max_row
        id_col = meta["fixed_map"].get("ID MASH", -1)
        nr_col = meta["fixed_map"].get("Nr", 1)
        for r in range(first_r, max_r + 1):
            idm = ws.cell(r, id_col).value if id_col > 0 else None
            nr = ws.cell(r, nr_col).value if nr_col > 0 else None
            if (nr is None or str(nr).strip() == "") and (idm is None or str(idm).strip() == ""):
                continue
            fixed_vals = extract_fixed_values(ws, meta["fixed_map"], r)
            means = row_means(ws, meta, r)
            line = list(fixed_vals) + [NIVELI_STUDIMIT_VALUE]
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
    ws.title = "Db_Master_FSHPJ"
    ws.append(headers)
    for row in rows:
        ws.append(row)
    out_wb.save(out_path)
    log.info("Wrote %s (%d rows).", out_path, len(rows))


def main(argv: Optional[List[str]] = None) -> int:
    p = argparse.ArgumentParser(description="Build Db_Master_FSHPJ.xlsx from FSHPJ Master workbook.")
    p.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    p.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    p.add_argument("-v", "--verbose", action="store_true")
    args = p.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(message)s",
    )
    log = logging.getLogger("build_db_master_fshpj")

    if not args.input.is_file():
        log.error("Input not found: %s", args.input)
        return 1

    headers, rows = process_workbook(args.input, log)
    write_output(args.output, headers, rows, log)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
