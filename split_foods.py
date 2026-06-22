#!/usr/bin/env python3
"""Split foods.json into N equal parts for parallel validation."""

import argparse
import json
import math
from datetime import datetime, timezone
from pathlib import Path


def split_foods(input_path: Path, output_dir: Path, parts: int = 5):
    data = json.loads(input_path.read_text(encoding="utf-8"))
    foods = data.get("foods", [])
    total = len(foods)
    if total == 0:
        raise SystemExit("No foods found in input file")

    chunk = math.ceil(total / parts)
    output_dir.mkdir(parents=True, exist_ok=True)

    for i in range(parts):
        start = i * chunk
        end = min(start + chunk, total)
        chunk_foods = foods[start:end]
        if not chunk_foods:
            break

        renumbered = []
        for j, food in enumerate(chunk_foods, start=1):
            item = dict(food)
            item["id"] = j
            renumbered.append(item)

        part_num = i + 1
        meta = {
            "source": f"{data.get('meta', {}).get('source', input_path.name)} (part {part_num}/{parts})",
            "part": part_num,
            "parts_total": parts,
            "part_range": f"{start + 1}-{end}",
            "count": len(renumbered),
            "generated_at": datetime.now(timezone.utc).isoformat(),
        }
        payload = {"meta": meta, "foods": renumbered}
        out = output_dir / f"foods_part{part_num}.json"
        out.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"Wrote {len(renumbered)} foods to {out}")


def main():
    parser = argparse.ArgumentParser(description="Split foods.json into equal parts")
    parser.add_argument("input", type=Path, nargs="?", default=Path("foods.json"))
    parser.add_argument("-o", "--output-dir", type=Path, default=Path("parts"))
    parser.add_argument("-n", "--parts", type=int, default=5)
    args = parser.parse_args()
    split_foods(args.input, args.output_dir, args.parts)


if __name__ == "__main__":
    main()
