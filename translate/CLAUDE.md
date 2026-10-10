# CLAUDE.md

Name: Chunsik

Rules:
- Always respond in Korean.
- For every .md file written in this repository, save a translation in the `translate` folder at the same relative path.
- If an original .md file is later modified or deleted, check the change (modification/deletion) and apply the same change to its translation in the `translate` folder.

## Work Process Rules

1. **Report the TODO list first**: when the user requests work through a prompt, always create the related TODO list and report it to the user before executing.
2. **Execute only after approval**: execute the work only after the user has read and approved the TODO list. Before approval, do not create, modify, delete, or move files. Read-only lookups needed to build the TODO list are allowed.
3. **Finish by organizing the structure**: if folders or files were added during the work, always reorganize them into the file and folder structure that Claude Code can recognize best before finishing. Follow the "Folder Structure" and "Output Storage Rules" sections below, and update the "Folder Structure" section whenever the structure changes.

## Folder Structure

```
agent_09/
├── CLAUDE.md                     # Project rules (this file)
├── .claude/skills/mk-ppt/        # Project-only PPT creation skill (python-pptx)
│   ├── SKILL.md
│   ├── references/api.md         # Code reference by feature
│   └── scripts/                  # inspect_pptx.py, slide_ops.py
├── research/<topic>/<round>/     # Research result .md files
│   └── ai-development/
│       ├── round1/               # Round 1 research (01–03)
│       └── round2/               # Round 2 research (01–04)
├── report/<topic>/               # Report .docx files
│   └── ai-development/
│       ├── AI_발전_보고서_1차.docx
│       └── AI_발전_보고서_2차_20261004.docx
├── ppt/<topic>/                  # Presentation .pptx files (created by the mk-ppt skill)
│   └── anthropic/
│       └── 클로드_회사소개_1차_20261010.pptx
└── translate/                    # .md translations (same relative paths as the originals)
    ├── CLAUDE.md
    ├── .claude/skills/mk-ppt/    # SKILL.md, references/api.md
    └── research/ai-development/round1/, round2/
```

## Output Storage Rules

- Name folders with lowercase English letters and hyphens (`-`). Example: `ai-development`
- File names may be in Korean; within the same round, prefix them with numbers such as `01_`, `02_` to show order.
- Research results: `research/<topic>/round<N>/NN_<title>.md`
- Reports: `report/<topic>/<report name>_<N>차[_YYYYMMDD].docx` (차 = round)
- Presentations: `ppt/<topic>/<presentation name>_<N>차[_YYYYMMDD].pptx` (created with the `mk-ppt` skill)
- When researching the same topic again, do not overwrite existing files; create a new round folder (`round<N+1>`).
- Exclude Word temporary lock files (`~$*`) via `.gitignore`.
