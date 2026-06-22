#!/usr/bin/env python3
"""Parse Build food data.xlsx into foods.json for the macro review tool."""

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path

import openpyxl

MACRO_KEYS = ("cal", "carbs", "fat", "prot", "fib")

# Column indices (1-based) matching the xlsx layout
COL_ORIGINAL = 1
COL_STANDARD = 2
COL_OTHER_NAMES = 3
COL_VARIETIES = 4
COL_SERVING_TYPES = 5
COL_BASE_VOLUME = 6
COL_GEN_MACROS = 7
COL_CORRECTION_MACROS = 8
COL_DB_MACROS = 9
COL_INPUT_MACROS = 10
COL_FINAL_MACROS = 11
COL_DIFF = 12


def _cell(row, col):
    if col - 1 >= len(row):
        return None
    return row[col - 1]


def _str_val(val):
    if val is None:
        return ""
    return str(val).strip()


def _split_lines(val):
    if not val:
        return []
    return [line.strip() for line in str(val).splitlines() if line.strip()]


def _parse_macros(s):
    if not s or not str(s).strip():
        return None
    parts = [p.strip() for p in str(s).split("|")]
    if len(parts) < 5:
        return None
    out = {}
    for i, key in enumerate(MACRO_KEYS):
        try:
            out[key] = float(parts[i])
        except (ValueError, TypeError):
            out[key] = None
    return out


def _macro_raw(val):
    s = _str_val(val)
    return s if s else None


def parse_workbook(path: Path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    foods = []
    food_id = 0

    for row in ws.iter_rows(min_row=3, values_only=True):
        original = _str_val(_cell(row, COL_ORIGINAL))
        if not original:
            continue

        food_id += 1
        input_raw = _macro_raw(_cell(row, COL_INPUT_MACROS))
        final_raw = _macro_raw(_cell(row, COL_FINAL_MACROS))

        diff_val = _cell(row, COL_DIFF)
        if diff_val is not None and diff_val != "":
            try:
                diff_val = float(diff_val)
            except (ValueError, TypeError):
                diff_val = _str_val(diff_val) or None
        else:
            diff_val = None

        foods.append({
            "id": food_id,
            "original_food_name": original,
            "standard_name": _str_val(_cell(row, COL_STANDARD)),
            "other_names": _str_val(_cell(row, COL_OTHER_NAMES)),
            "food_varieties": _split_lines(_cell(row, COL_VARIETIES)),
            "serving_types": _split_lines(_cell(row, COL_SERVING_TYPES)),
            "base_volume": _str_val(_cell(row, COL_BASE_VOLUME)),
            "gen_macros_raw": _macro_raw(_cell(row, COL_GEN_MACROS)),
            "correction_macros_raw": _macro_raw(_cell(row, COL_CORRECTION_MACROS)),
            "db_macros_raw": _macro_raw(_cell(row, COL_DB_MACROS)),
            "input_macros": _parse_macros(input_raw),
            "final_macros": _parse_macros(final_raw),
            "input_macros_raw": input_raw,
            "final_macros_raw": final_raw,
            "diff": diff_val,
        })

    wb.close()
    return foods


def main():
    parser = argparse.ArgumentParser(description="Parse food xlsx to JSON")
    parser.add_argument("input", type=Path, help="Input xlsx file")
    parser.add_argument("-o", "--output", type=Path, default=Path("foods.json"))
    args = parser.parse_args()

    foods = parse_workbook(args.input)
    payload = {
        "meta": {
            "source": args.input.name,
            "count": len(foods),
            "generated_at": datetime.now(timezone.utc).isoformat(),
        },
        "foods": foods,
    }

    args.output.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Wrote {len(foods)} foods to {args.output}")


if __name__ == "__main__":
    main()
