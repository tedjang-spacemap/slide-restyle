---
name: slide-restyle
description: 샘플 PPT로 DESIGN.md를 만들고 다른 PPT를 그 디자인으로 변환한다. `/slide-restyle design <sample.pptx…>`, `/slide-restyle convert <input.pptx> [--variant srmc]`. PPT 디자인 추출, 템플릿 적용, 슬라이드 리스타일 요청 시 사용한다.
argument-hint: design <sample.pptx…> | convert <input.pptx> [--variant collision|srmc]
---

# slide-restyle — 샘플 PPT 디자인으로 다른 PPT 변환

`$ARGUMENTS`의 첫 단어로 흐름을 고른다: `design` 또는 `convert`. 없거나 모호하면 사용자에게 묻는다.

## 공통 규칙

- 저장소 루트에서 실행한다. `S=.claude/skills/slide-restyle/scripts`
- 스크립트는 항상 `uv run $S/<name>.py`. 전역 pip 설치 금지.
- 렌더는 반드시 `$S/render.sh <deck.pptx> <out_dir> [dpi]`. 직접 `soffice` 호출 금지(폰트 설정 주입 필요).
- 폰트: 정적 Roboto는 `./.fonts/`에, 한글은 Pretendard. 렌더 결과에 폰트 대체 경고가 있으면 보고한다.
- 디렉터리: `samples/`(디자인 원본) · `inputs/`(변환 대상) · `design/`(추출 산출물) · `work/<deck>/`(중간물) · `outputs/`(결과). `samples/ inputs/ work/ outputs/`는 커밋하지 않는다.
- **원본 내용은 바꾸지 않는다.** 요약·분할·수작업 필요는 `outputs/<deck>.report.md`에 기록한다.
- 새 결정(사용자가 고른 것)은 `docs/DECISIONS.md`에 번호를 이어 추가한다. 디자인 값 자체는 DESIGN.md에만 둔다.

## `design <sample.pptx…>`

1. 측정: `uv run $S/extract_design.py <samples…> --out design`
   → `design/{design.json,summary.md,assets/,base.pptx,base.potx}`. 대표 덱이 필요하면 `--primary`.
2. `design/summary.md`를 읽는다. 제외할 요소(로고·배너 등)는 사용자에게 확인하고 `design/exclude.yaml`에 적은 뒤 1을 다시 실행한다.
3. DESIGN.md 작성: 골격과 토큰 규칙은 [reference/design-md-template.md](reference/design-md-template.md).
   - yaml `design-tokens` 블록의 모든 값에 `source: measured`(+`evidence`) 또는 `estimated`(+`basis`)를 단다. 추정은 본문에도 "(추정)".
   - 서술(분위기·역할·Do/Don't·9장 분류 규칙)은 Claude가 쓰고, 값은 design.json 측정값에서 가져온다.
4. 검증: `uv run $S/design_tokens.py DESIGN.md` 통과할 때까지 수정.
5. 보고: 레이아웃 중 실측/추정 개수, 필요한 폰트 목록, 사용자 결정이 필요한 것(폰트, 제외 요소, 변형).

## `convert <input.pptx> [--variant collision|srmc]`

0. 선결: `DESIGN.md`와 `design/base.pptx`가 없으면 `design`을 먼저 하라고 안내하고 멈춘다.
1. 추출: `uv run $S/extract_content.py inputs/x.pptx --render` → `work/x/{content.json,media/,src_png/}`
2. 분류: content.json과 `src_png/`를 보고 [DESIGN.md 9장](../../../DESIGN.md) 규칙으로 `work/x/plan.json`을 쓴다.
   스키마는 [reference/plan-schema.md](reference/plan-schema.md), 입력 구조는 [reference/content-schema.md](reference/content-schema.md).
   텍스트는 `{"t","src"}` 참조만(원문 축자 부분 문자열).
3. 검사: `uv run $S/check_plan.py work/x/plan.json --md work/x/plan.md` — 통과할 때까지 plan 수정.
4. **분류표 확인**: `plan.md`를 사용자에게 보여 주고 승인받은 뒤 빌드한다. 변형은 사용자가 지정한 값만 쓴다(기본 collision, 임의 변경 금지).
5. 생성: `uv run $S/build_deck.py work/x/plan.json [--variant srmc] --render`
   → `outputs/x.pptx`(+ `.report.md`), `work/x/build/<variant>/{build.json,png/}`
6. 검수: `png/`를 슬라이드별로 직접 보고 넘침·겹침·잘림·폰트 대체를 확인한다. 문제가 있으면 plan.json을 고쳐 3→5를 반복한다
   (넘침은 축소 → 재배치 → 분할 → 요약 순, DESIGN.md 8장).
7. 보고: 출력 경로, 사용 변형, 추정으로 그려진 레이아웃, 요약·분할 내역, 수작업 필요 장, 렌더 폰트 대체.

미구현 레이아웃(toc, intro_columns, visual_kv, image_pair_kv)이 필요하면 build_deck이 중단한다. 우회하지 말고 사용자에게 알린다.
