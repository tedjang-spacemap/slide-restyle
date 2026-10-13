# slide-restyle

샘플 PPT의 디자인을 DESIGN.md로 추출하고, 다른 PPT를 그 디자인으로 변환하는 Claude Code 스킬.
스킬 본체: `.claude/skills/slide-restyle/` (scripts/, reference/).
- 계획·마일스톤(M1–M4)·진행 상태: [docs/PLAN.md](docs/PLAN.md) — 현재 M3 완료(사용자 확인 대기), M4 대기
- 주요 결정과 이유: [docs/DECISIONS.md](docs/DECISIONS.md) — 새 결정은 여기에 추가, 디자인 값 자체는 DESIGN.md

## Commands

```bash
S=.claude/skills/slide-restyle/scripts
uv run $S/extract_design.py samples/*.pptx --out design     # 측정 → design/{design.json,summary.md,assets/,base.pptx,base.potx}
uv run $S/design_tokens.py DESIGN.md                        # DESIGN.md의 yaml design-tokens 블록 검증
uv run $S/extract_content.py inputs/x.pptx --render         # → work/x/{content.json,media/,src_png/}
uv run $S/check_plan.py work/x/plan.json --md work/x/plan.md # plan.json 원문 축자·누락·한도 검사 + 분류표
uv run $S/build_deck.py work/x/plan.json [--variant srmc] --render  # → outputs/x.pptx + x.report.md, work/x/build/<variant>/{build.json,png/}
$S/render.sh <deck.pptx> <out_dir> [dpi]                    # PDF + slide-N.png + 폰트 대체 경고
```
스크립트는 PEP 723 헤더로 의존성을 선언 → 항상 `uv run`으로 실행 (전역 pip 설치 없음).

## 역할 분담
- 스크립트: 측정·추출·생성·렌더 (결정적). Claude: DESIGN.md 서술, 레이아웃 분류, 요약/분할 판단, 렌더 검수.
- 디자인 변형 2개: `collision`(기본) / `srmc`. 레이아웃 좌표는 native 변형 기준 실측값으로 저장하고, 다른 변형에서는 content_area 선형 변환(→ 자동으로 estimated).
- DESIGN.md 토큰은 모두 `source: measured`(+evidence `C:S5`) 또는 `estimated`(+basis). 추정값은 본문에도 "(추정)".
- 원본 내용은 바꾸지 않는다. plan.json 텍스트는 `{"t","src"}` 참조만(축자 부분 문자열) — check_plan.py가 검사.
  요약/분할/수작업 필요는 outputs/의 변환 보고서에 기록.

## Gotchas
- LibreOffice headless는 번들 fontconfig라 macOS 폰트를 못 본다 → 반드시 `render.sh` 사용(폰트 설정을 생성해 주입). 직접 `soffice` 호출 금지.
- Homebrew `font-roboto`는 가변 폰트라 "Roboto Light"로 매칭 안 됨 → 정적 Roboto를 `./.fonts/`에 둠(gitignore). 없으면 googlefonts/roboto v2.138 roboto-unhinted.zip.
- 출력 한글 폰트는 Pretendard(제목 Bold / 본문 Light, 사용자 결정). 샘플 테마 ea는 맑은 고딕이라 런마다 `<a:ea>`를 명시해야 함.
- 맑은 고딕은 macOS 설치 불가 → 샘플 렌더 시 render.sh가 Apple SD Gothic Neo로 alias(샘플 한글 QA만 근사치).
- 샘플 테마 색은 Office 기본값이라 무의미. 실제 색은 런/도형 직접 지정값(extract_design이 집계).
- 샘플 슬라이드는 placeholder가 아닌 자유 텍스트상자(Google Slides 출신) → chrome은 반복 요소 탐지로 찾음.
- 제외 이미지는 `design/exclude.yaml` (dHash로 해상도 달라도 매칭). MINISTRY 배너는 사용자 요청으로 제외.
- 공식 pptx 스킬의 validate.py는 `python -I`로 실행하면 형제 모듈 import 실패 → `-I` 없이 실행.

## Git
- `samples/`, `inputs/`, `outputs/`, `work/`, `.fonts/`는 커밋하지 않는다(원본 자료·산출물).
