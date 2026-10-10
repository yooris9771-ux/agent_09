---
name: mk-ppt
description: A skill exclusive to the agent_09 project that creates, edits, and reads PowerPoint (.pptx) files with python-pptx. Always use this skill when the user mentions PPT, PowerPoint, presentation materials, slides, a deck, or .pptx; asks to turn research results (research/) or report content into a presentation; or requests extracting/editing text, tables, charts, or images in an existing .pptx, or deleting/duplicating/reordering slides. Within this project, prefer this skill even when the user just says "make a presentation" without naming a file format.
---

# mk-ppt

This is the procedure for creating, editing, and analyzing .pptx files with python-pptx (1.0.2) inside this repository. It follows the project rules (CLAUDE.md) as-is: report the TODO list first and execute only after approval.

## 1. Python executable path

Python 3.13 is installed at user scope. If PATH has not been applied to the current shell, `python` resolves to the Microsoft Store shortcut and fails, so check first and use the full path if it doesn't work.

- PowerShell: `& "$env:LOCALAPPDATA\Programs\Python\Python313\python.exe"`
- Bash: `"$LOCALAPPDATA/Programs/Python/Python313/python.exe"`

If you get an error that python-pptx is missing, run `-m pip install python-pptx` with the same interpreter. Replace every `python` example in this document with the path above.

Also put `sys.stdout.reconfigure(encoding="utf-8")` at the top of any generation script you write yourself, because Korean paths and text printed under the Windows console's default encoding (cp949) come out garbled.

## 2. Workflow

1. **Understand the request** — Organize the number of slides, topic, source material (e.g., `research/<topic>/round<N>/*.md`), aspect ratio (16:9 by default), and whether a template exists. If there is source material, read it first and include a per-slide outline in the TODO report.
2. **Write the generation script** — Put one-off generation scripts in the scratchpad (to avoid cluttering the repository). See `references/api.md` for code patterns by feature.
3. **Run and save** — Save outputs to the path given in the storage rules below.
4. **Verify** — Reopen the file with `scripts/inspect_pptx.py <file>` to confirm the slide count, titles, and text are as intended, and summarize the result for the user. If the file won't open or text is empty, fix the cause and verify again.

## 3. Output storage rules

- Path: `ppt/<topic>/<presentation name>_<N>차[_YYYYMMDD].pptx` (folder names in lowercase English + hyphens; file names may be Korean; 차 = round)
- Do not overwrite existing files; increment the round number.
- When creating the `ppt/` folder or a new topic folder for the first time, update the "Folder Structure" section of CLAUDE.md and `translate/CLAUDE.md` together.

## 4. Design defaults

For consistent results, use the following unless asked otherwise.

- Slide size 16:9 (`Inches(13.333)` × `Inches(7.5)`)
- Korean font `맑은 고딕` (Malgun Gothic) — if not specified, Korean text falls back to a default Western font and looks awkward. Set `font.name` per run.
- Titles 32–40pt, body 18–24pt, no more than 5–6 bullets per slide — split the slide if there are more.
- First slide is a title slide (title, subtitle, date); last slide is a summary or conclusion.
- When there is numeric data, show it as a table or chart instead of listing sentences.

## 5. Feature scope summary

Detailed code is in `references/api.md`, organized by area. Read only the section you need.

| Area | Supported tasks |
|---|---|
| File | Create, open, save (path/BytesIO), slide size, document properties, start from a template |
| Slides | Add with a chosen layout, iterate, background, speaker notes, find layout by name |
| Shapes | Auto shapes, text boxes, pictures (cropping), tables, charts, movies, connectors, groups, freeforms, OLE, placeholders, rotation, shadow, hyperlinks |
| Text | Paragraphs/runs, fonts, alignment, spacing, margins, auto-fit, hyperlinks |
| Tables | Cell values, merge/split, column width/row height, cell fill, style options |
| Charts | Bar, line, pie, doughnut, area, scatter, bubble, radar; replace data, axes, legend, data labels, series formatting |
| Formatting | Solid, gradient, and pattern fills; theme colors; line formatting |
| Reading/analysis | Extract text, images, table and chart data; find and replace |
| Low level | Edit XML directly via `._element` |

## 6. Tasks without an official API

- **Delete, duplicate, move slides** → use `scripts/slide_ops.py` (`python` is the full path from section 1; numbers start at 1)
  - `python .claude/skills/mk-ppt/scripts/slide_ops.py delete <in> <out> <number>`
  - `python .claude/skills/mk-ppt/scripts/slide_ops.py duplicate <in> <out> <number>` — the copy is inserted right after the original. Charts are copied independently, and speaker notes are copied as text only.
  - `python .claude/skills/mk-ppt/scripts/slide_ops.py move <in> <out> <old number> <new number>`
- **Animations and transitions** → these require direct XML editing, so first ask the user whether doing it manually is preferable.
- **SmartArt, sections, comments, .ppt (legacy format)** → not supported. Substitute with combinations of shapes, or guide the user to convert to .pptx.
- **Rendering to images/PDF** → requires an external tool such as LibreOffice. If none is installed, use text verification with `inspect_pptx.py` instead.
