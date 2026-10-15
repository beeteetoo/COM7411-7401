# CLAUDE.md — Research Methods 2k26-27-B2 (COM7401 / COM7411)

## What this project is
One shared Research Methods module delivered in **two threads**. Module spec, assessment brief, learning outcomes and slide content are identical; only the *delivery wrapper* differs.

| | **Thread A — COM7401 On Campus** | **Thread B — COM7411 Online** |
|---|---|---|
| Sessions | Week N AM + PM on the **same day**, face to face, ~180 min each | Week N AM = **Monday evening**, Week N PM = **Saturday morning**, Zoom, 3.5 h window each (expect ~2h45 live, fewer students) |
| Extra | One-hour **Zoom** a couple of days after the campus day (mass, cross-cohort, ~130 students) | **No Zoom hour.** Its functions are absorbed into the Saturday sessions, bridge tasks or recordings |
| Pedagogy | Active face-to-face learning; Zoom for scale (demo/briefing/clinic) | Flipped + dialogic: recorded micro-input, breakouts, shared boards, forum, asynchronous bridge tasks, more individual feedback |
| Week 5 | AM = defence workshop, PM = on-campus assignment support | AM = defence workshop, PM = online assignment support (register + drop in) |

Folder names `Week N_AM` / `Week N_PM` are shared by both threads. For Thread B, "AM" means the Monday-evening session and "PM" the Saturday-morning session.

## Scope rules
- **Ignore everything under `26-27-B2-ModuleBox/ARCHBT2/`.** It is superseded; do not read it, link to it or quote it. (The other ModuleBox folders, e.g. `0.2 Assessments&Moderation`, are valid sources.)
- Do **not** open `.pptx` files unless the task needs it. Deck edits happen in Pass 2 only.
- Source of truth for assessment: `26-27-B2-ModuleBox/0.2 Assessments&Moderation/COM7401_Assessment_Brief.docx` and `COM7411_Assessment_Brief.docx` (identical except code and dates).

## Assessment facts that the guides must respect
- Portfolio, 100%: **Part A** written report (2,500 words ±10%, fixed headings: Background ~400 / Critical Literature Review ~900 / Methodology ~900 / Reflection ~300) + **Part B** recorded oral defence (4 min ±30 s).
- Report due end of **Week 4** (COM7401 Fri 27 Nov 2026; COM7411 Sat 28 Nov 2026). Oral defence due end of **Week 5** (Fri 4 Dec; Sat 5 Dec).
- The oral defence tests one of four fixed methodological checks (Build vs Finding, Collapse Test, No/Undefined-Word Test, "To…" Test), issued in Week 4 — the brief says "Thursday Zoom" (campus). COM7411 gets it live in the Week 4 Saturday session with a Moodle fallback.
- GAI: Category B for the report (declare use); **no GAI and no scripted content** for the oral defence.
- The old "6–8 min recorded presentation / Week 5 recording day / full portfolio draft in Week 5" model is **obsolete**.

## Working method (two passes)
1. **Pass 1 — guides.** Facilitator Guide, all ten Session Guides and Zoom notes are updated together so the two threads stay coherent.
2. **Pass 2 — decks.** One deck at a time, `Week 1_AM` → `Week 5_PM`, **pause for the user to check after each**. Take a backup copy of the .pptx (scratchpad) before any in-place edit. `PPTX Edit Instructions.md` is revised per deck as each is done.

## Session Guide conventions
Every Session Guide has the same five parts: **1 Core (identical)** → **2 Thread A run sheet** → **3 Thread B run sheet** → **4 Between sessions / Zoom hook** → **5 Tips and parity check**.
- Change Core and both threads together; if a task changes, say what happens in each thread.
- Refer to slides by **section name, not slide number** (old numbers were unreliable; Pass 2 reconciles them).
- Thread B run sheets plan ~2h45 live and leave the rest of the 3.5 h as open floor / drop-in.
- Never define the four checks from memory; link to the module team's source once the user supplies it.

## Pass 2 progress
- Done: Week 1 AM (17 slides), Week 1 PM (20 slides). Next: Week 2 AM.
- Decks are **forked per module** where content differs (Week 1 AM and PM done). Method: python-pptx in a venv (scratchpad) — build script trims/reorders/adds slides from the original, validate with the pptx skill's `validate.py --original`, render to check. Slides are **student-facing only** (no thread/meta labels); the user keeps their own folder backup, so edit in place.

## Two modules, two Moodle sites — naming convention
COM7401 (On Campus) and COM7411 (Online) each have their **own Moodle site**; students only ever see their own module's files.
- **Anything that differs by module** carries the module in the **filename** — `COM7401 CAMPUS - …` or `COM7411 ONLINE - …` (decks: `… (COM7401 On Campus).pptx` / `… (COM7411 Online).pptx`) — **and** a visible label inside the document (docx header/footer; deck slide text stays student-facing and refers only to the student's own module).
- **Shared content** (identical for both) keeps an unlabelled name and must not mention the other module.
- **Student-facing files never name the other module or "thread".** Lecturer files may name both and say which is which.
- If a deck would otherwise need thread-specific slides, **fork the deck** rather than compromise.

## Philosophical Stance Task (Week 1) — two versions
- **COM7411 Online:** `Week 1_PM/Pre-Session Prep/COM7411 ONLINE - Philosophical Stance Task.docx` — ~2½ h bridge task before the Saturday session; forum post + peer reply.
- **COM7401 On Campus:** `Zoom Sessions/Zoom Day1/COM7401 CAMPUS - Philosophical Stance Task (after Week 1, completed in Zoom 1).docx` — Stage 1 (~70 min) after the Week 1 campus day, Stage 2 live in Thursday's Zoom 1 (run sheet + extracts: `Zoom Day1/Zoom Day1 Guide.md`).
- Lecturer copy for both: `Facilitator Notes (Not for Students)/Lecturer_Copy-Philosophical Stance Task (COM7401 + COM7411).docx`.
- Epistemology-weighted (~55%), MSc level, unmarked. Cheat sheets/readings stay as resources. Consider the same "task around the cheat sheets" approach for later weeks.

## Week 2 additions
- **Literature Search Task** (shared, pre-session, `Week 2_AM/Pre-Session Prep/`): precedes the Funnel and the PM critique; replaces the old "lunch task".
- **Ethics:** COM7401 = Ethics Core live in Zoom 2 (Thursday); COM7411 = `Ethics Pack (Independent Study)/COM7411 ONLINE - Ethics in Research Activity.docx` (+ lecturer copy). Each module has its own labelled `Ethics Self-Assessment Worksheet` (old shared .md removed — it contained ARCHBT2 wording).
- Week 2 AM decks forked (COM7401 / COM7411); the shared question-types pre-reading deck lives in `Pre-Session Prep/To be issues after week 1 PM/`.

## Deck standing rules
- **Keep each deck's closing sequence:** the unique cartoon + "Modern Science" slide ("From this point on… you're a scientific researcher!") followed by Reflection & Writing and Questions. Place new "next steps" slides *before* the Modern Science slide. (Restored in Week 1 PM and Week 2 AM; Week 1 AM ends on that slide.)
- Slides are student-facing only; no thread/meta labels beyond the module code on the cover and footer.
- Fix legacy footers/covers ("COM7302") and PhD-level wording ("doctoral", "thesis") when touching a deck.

## Things to Check files
Each session folder (`Week N_AM` / `Week N_PM`) has a **`Things to Check.md`**: Moodle set-up, assumptions made, decisions awaiting the user, and risks. Create/update it whenever a session is worked on. Start each session's review with it.

## Week 2 PM additions
- **Critical Appraisal Worksheet** (stance-aware, shared, in-class) + lecturer copy; `Comparison_of_Literature_Reviews.docx` kept as a reference table with a corrected stance column. The long `Critical Evaluation Worksheet.docx` is now an optional extension.

## Open items (update as resolved)
- Four-check definitions and allocation mechanism: user to supply.
- Report deadline *time* on the Saturday for COM7411 (same day as Week 4 PM session).
- Literature searching demo (outline in Facilitator Notes) to record; upload to both Moodle sites. COM7411 Moodle needs the "Ethics cases" forum and the link to the university's current ethics guidance.
- Zoom Day1–5 individual guides, Week 5 activity files and `PPTX Edit Instructions.md` still describe the old model.
