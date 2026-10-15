"""Read-only package checks for the shared dual-track teaching materials."""

from csv import DictReader
from pathlib import Path
from statistics import median
from urllib.parse import unquote
from zipfile import ZipFile
import re


ROOT = Path(__file__).resolve().parents[2]
MASTER = ROOT / "Shared Module Master"
NOTES = MASTER / "Student Notes"
SLIDES = MASTER / "Slides"

errors = []
for num in range(1, 11):
    prefix = f"{num:02d} "
    n = [x for x in NOTES.glob(prefix + "*.md") if not any(k in x.name for k in ("Practice", "Sample"))]
    p = list(SLIDES.glob(prefix + "*.pptx"))
    if len(n) != 1:
        errors.append(f"topic {num}: expected one note; found {len(n)}")
    elif not all(marker in n[0].read_text() for marker in ("**Goal:**", "**Worked example.", "**Glossary:**", "**Your task:**")):
        errors.append(f"{n[0].name}: missing goal, worked example, glossary or task")
    if len(p) != 1:
        errors.append(f"topic {num}: expected one slide deck; found {len(p)}")
    if p:
        with ZipFile(p[0]) as z:
            names = [x for x in z.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", x)]
            if len(names) != 6:
                errors.append(f"{p[0].name}: expected 6 slides; found {len(names)}")
            body = " ".join(z.read(x).decode("utf-8") for x in names)
            for stale in ("COM7302", "6–8 minutes", "six-to-eight-minute"):
                if stale in body:
                    errors.append(f"{p[0].name}: stale phrase {stale}")

if list(NOTES.glob("*.pdf")):
    errors.append("A PDF was placed in student notes; check for staff-only book content")

paths = [
    MASTER / "README.md",
    ROOT / "Facilitator Guide - Dual Track.md",
    ROOT / "Research-Methods-Dual-Track-Development-Plan.md",
    MASTER / "Facilitator/Source Map.md",
    MASTER / "Facilitator/Worked Examples.md",
    ROOT / "Delivery/COM7401/Campus Session Guides.md",
    ROOT / "Delivery/COM7401/Wednesday Zoom Guides.md",
    ROOT / "Delivery/COM7411/Online Session Guides.md",
    ROOT / "Assessment/Assessment Alignment.md",
    ROOT / "Assessment/Scheduled Hours Ledger.md",
    ROOT / "Assessment/Four-Check Reference Guide - PROVISIONAL.md",
    ROOT / "Assessment/Proposed Brief Administration Wording - STAFF REVIEW.md",
    ROOT / "Assessment/Release and Pilot Checklist.md",
    ROOT / "Ethics Pack (Independent Study)/Ethics Self-Assessment Worksheet.md",
]
paths.extend(NOTES.glob("*.md"))
for path in paths:
    text = path.read_text()
    for dest in re.findall(r"(?<!!)\[[^\]]+\]\(([^)]+)\)", text):
        if dest.startswith(("http:", "https:", "mailto:", "#")):
            continue
        target = (path.parent / unquote(dest.split("#")[0])).resolve()
        if not target.exists():
            errors.append(f"{path.relative_to(ROOT)}: broken link to {dest}")

with (NOTES / "07 Sample Data.csv").open(newline="") as f:
    rows = list(DictReader(f))
for label, correct, med in (("Current", 3, 27), ("Proposed", 4, 20)):
    group = [r for r in rows if r["interface"] == label]
    if len(group) != 5 or sum(int(r["correct_flag"]) for r in group) != correct or median(int(r["response_seconds"]) for r in group) != med:
        errors.append(f"{label} worked numeric answer does not match CSV")

if errors:
    print("Material validation failed:")
    for error in errors:
        print("-", error)
    raise SystemExit(1)

print("PASS: 10 numbered notes, 10 six-slide decks, local links, staff-only PDF placement and worked CSV answers")
