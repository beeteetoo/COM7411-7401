#!/usr/bin/env python3
"""Allocate the four methodological checks to students, balanced within each module.

Usage:
    python allocate_checks.py roster.csv --seed 20261116 [--id-col StudentID] [--module-col Module]

The roster stays on your machine: this script reads it, writes two files next to it
(<roster>_allocation.csv for Moodle / mail merge, <roster>_allocation_record.txt for
moderation) and never sends anything anywhere. Do not commit either output to the repo.

Method: within each module, students are sorted by ID (so the result does not depend
on roster order), shuffled with a seeded generator, then dealt round-robin to the four
checks so group sizes differ by at most one. The same roster and seed always give the
same result, so the allocation can be re-created for moderation or an appeal.
"""

import argparse
import csv
import hashlib
import random
from collections import Counter
from pathlib import Path

CHECKS = ["Build vs Finding", "Collapse Test", "No/Undefined-Word Test", "To... Test"]


def allocate(rows, id_col, module_col, seed):
    seen = set()
    for r in rows:
        key = (r[module_col].strip(), r[id_col].strip())
        if not key[1]:
            raise SystemExit("Empty student ID in roster")
        if key in seen:
            raise SystemExit(f"Duplicate student in roster: {key[1]}")
        seen.add(key)
    out = []
    for module in sorted({m for m, _ in seen}):
        ids = sorted(i for m, i in seen if m == module)
        rng = random.Random(f"{seed}|{module}")
        rng.shuffle(ids)
        # rotate the starting check per module so odd remainders are not always the same checks
        offset = rng.randrange(len(CHECKS))
        for n, sid in enumerate(ids):
            out.append((module, sid, CHECKS[(n + offset) % len(CHECKS)]))
    return sorted(out)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("roster", type=Path)
    ap.add_argument("--seed", required=True, help="any string/number; record it")
    ap.add_argument("--id-col", default="StudentID")
    ap.add_argument("--module-col", default="Module")
    a = ap.parse_args()

    raw = a.roster.read_bytes()
    with a.roster.open(newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    for col in (a.id_col, a.module_col):
        if not rows or col not in rows[0]:
            raise SystemExit(f"Column '{col}' not found. Columns: {list(rows[0]) if rows else 'none'}")

    result = allocate(rows, a.id_col, a.module_col, a.seed)

    csv_path = a.roster.with_name(a.roster.stem + "_allocation.csv")
    with csv_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Module", "StudentID", "Check"])
        w.writerows(result)

    counts = Counter((m, c) for m, _, c in result)
    lines = [
        "Four checks - allocation record",
        f"Seed: {a.seed}",
        f"Roster SHA-256: {hashlib.sha256(raw).hexdigest()}",
        f"Students: {len(result)}",
        "Method: sort IDs, seeded shuffle per module, deal round-robin (see script docstring).",
        "",
    ]
    for module in sorted({m for m, _, _ in result}):
        lines.append(f"{module}:")
        for c in CHECKS:
            lines.append(f"  {c}: {counts[(module, c)]}")
    rec_path = a.roster.with_name(a.roster.stem + "_allocation_record.txt")
    rec_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Wrote {csv_path}\nWrote {rec_path}")


if __name__ == "__main__":
    main()
