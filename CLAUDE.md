# CLAUDE.md

이름: 춘식

규칙:
- 답변은 무조건 한글로 작성한다.
- 이 저장소에 작성되는 모든 .md 파일은 번역본을 `translate` 폴더에 동일한 상대 경로로 저장한다.
- 원본 .md 파일이 이후 수정되거나 삭제될 경우, 그 변경 사항(수정/삭제)을 확인한 뒤 `translate` 폴더의 번역본에도 동일하게 반영한다.

## 작업 진행 규칙

1. **TODO 리스트 선보고**: 사용자가 프롬프트로 작업을 요청하면, 실행하기 전에 관련 TODO 리스트를 항상 먼저 만들어 사용자에게 보고한다.
2. **승인 후 실행**: 사용자가 TODO 리스트를 읽고 승인한 뒤에만 작업을 실행한다. 승인 전에는 파일 생성·수정·삭제·이동을 하지 않는다. 다만 TODO 리스트를 만드는 데 필요한 읽기 전용 조회는 할 수 있다.
3. **구조 정리 후 마무리**: 작업 중 폴더나 파일이 추가되었다면, 마무리하기 전에 항상 클로드 코드가 가장 잘 인식할 수 있는 파일·폴더 구조로 정리한다. 기준은 아래 "폴더 구조"와 "산출물 저장 규칙"을 따르고, 구조가 바뀌면 "폴더 구조" 섹션도 함께 갱신한다.

## 폴더 구조

```
agent_09/
├── CLAUDE.md                     # 프로젝트 규칙 (이 파일)
├── .claude/skills/mk-ppt/        # 프로젝트 전용 PPT 제작 스킬 (python-pptx)
│   ├── SKILL.md
│   ├── references/api.md         # 기능별 코드 레퍼런스
│   └── scripts/                  # inspect_pptx.py, slide_ops.py
├── research/<주제>/<회차>/        # 자료조사 결과 .md
│   └── ai-development/
│       ├── round1/               # 1차 조사 (01~03)
│       └── round2/               # 2차 조사 (01~04)
├── report/<주제>/                 # 보고서 .docx
│   └── ai-development/
│       ├── AI_발전_보고서_1차.docx
│       └── AI_발전_보고서_2차_20261004.docx
├── ppt/<주제>/                    # 발표자료 .pptx (mk-ppt 스킬로 생성)
│   └── anthropic/
│       └── 클로드_회사소개_1차_20261010.pptx
└── translate/                    # .md 번역본 (원본과 동일한 상대 경로)
    ├── CLAUDE.md
    ├── .claude/skills/mk-ppt/    # SKILL.md, references/api.md
    └── research/ai-development/round1/, round2/
```

## 산출물 저장 규칙

- 폴더 이름은 영문 소문자와 하이픈(`-`)으로 짓는다. 예: `ai-development`
- 파일 이름은 한글을 써도 되며, 같은 회차 안에서 `01_`, `02_`처럼 번호를 붙여 순서를 표시한다.
- 자료조사 결과: `research/<주제>/round<N>/NN_<제목>.md`
- 보고서: `report/<주제>/<보고서명>_<N>차[_YYYYMMDD].docx`
- 발표자료: `ppt/<주제>/<발표자료명>_<N>차[_YYYYMMDD].pptx` (`mk-ppt` 스킬로 생성)
- 같은 주제를 다시 조사할 때는 기존 파일을 덮어쓰지 않고 새 회차 폴더(`round<N+1>`)를 만든다.
- Word 임시 잠금 파일(`~$*`)은 `.gitignore`로 제외한다.
