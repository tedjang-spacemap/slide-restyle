# PLAN — slide-restyle

샘플 PPT의 디자인을 awesome-design-md 형식의 DESIGN.md로 추출하고, 그 DESIGN.md를 기준으로 다른 PPT를
같은 디자인으로 변환하는 Claude Code 스킬. 원래 요구사항은 2026-10-13 세션에서 사용자가 제시했다.
결정과 그 이유는 [DECISIONS.md](DECISIONS.md)에 있다.

## 워크플로

```
samples/*.pptx ─extract_design.py→ design/design.json ─(Claude 서술)→ DESIGN.md + design/base.pptx
inputs/x.pptx  ─extract_content.py→ work/x/content.json ─(Claude 분류·넘침 판단)→ work/x/plan.json
plan.json + DESIGN.md ─build_deck.py→ outputs/x.pptx ─render.sh→ work/x/png/ ─(Claude 검수)→ 재생성 or 완료
                                                                                → outputs/x.report.md
```

호출 형태: `/slide-restyle design <sample.pptx>`, `/slide-restyle convert <input.pptx> [--variant srmc]`

## 진행 방식

- 마일스톤 단위로 진행하고, 마일스톤이 끝날 때마다 사용자 확인을 받은 뒤 다음으로 넘어간다.
- 계획을 먼저 보여 주고, 승인받은 뒤 코드를 쓴다.

## 마일스톤

### M1 — extract_design.py + DESIGN.md 초안   [완료 2026-10-13]
- 산출물: `extract_design.py`, `design_tokens.py`(검증기), `render.sh`, `reference/design-md-template.md`,
  `DESIGN.md`, `design/`(design.json, summary.md, assets/, base.pptx, base.potx, exclude.yaml)
- 확인 기준: 토큰이 샘플 실측값과 맞는지, 실측/추정이 구분되는지, 필요한 폰트 목록
- 결과: 토큰만으로 C:S5와 표지를 다시 그려 원본과 비교(`work/specimen/compare.png`). 레이아웃 13개 중
  6개 실측, 7개 추정. 사용자 결정 3건 반영 후 변형 2개(collision/srmc)로 확장.

### M2 — extract_content.py + 레이아웃 분류   [완료 2026-10-13]
- 산출물: `extract_content.py`, `check_plan.py`, `reference/content-schema.md`·`plan-schema.md`,
  레이아웃 분류 결과 `work/<deck>/plan.json`(+ `plan.md` 분류표)
- 확인 기준: 입력 PPT 슬라이드별 분류표, 변환 불가 후보(SmartArt, 자유 배치 도식 등)
- 첫 입력: `inputs/spacemap-flow.pptx` (10장 → 출력 11장)
- [x] `extract_content.py` — 슬라이드별 텍스트 계층(런 서식), 이미지, 표, 차트, SmartArt, 노트 + 특징값(역할 후보, 격자, 연결 경로) (결정적)
- [x] `content.json` / `plan.json` 스키마 (`reference/`)
- [x] 카탈로그 확장 (사용자 결정, DECISIONS #13): `io_steps`, `card_grid`(+`stats`), `cover.flow`/`legend` — 모두 추정
- [x] DESIGN.md 9장 규칙으로 분류 → `plan.json`, `check_plan.py`로 원문 축자·누락·한도 검사 통과
- [x] 분류표 사용자 확인
- 결과: 10장 → 11장. io_steps 5, image_full 3(수작업 필요 S2·S4·S7, 원본 렌더라 밝은 배경), card_grid 2, cover 1.
  도식 3장 처리와 태그 색(D 파랑·H 노랑·A 분홍·N 회색)은 현재 안대로 두고 M3 렌더 검수에서 다시 본다.

### M3 — build_deck.py + render 검수 루프   [완료 2026-10-13, 사용자 확인 대기]
- 산출물: `build_deck.py`, 검수 루프, `outputs/<deck>.report.md` 형식
- 확인 기준: 변환 결과 PPT와 PNG, 보고서
- [x] `base.pptx` 위에 슬라이드 추가, `--variant`로 chrome 선택 (collision/srmc 둘 다 생성·렌더 확인)
- [x] 베이스 테마의 ea 폰트 자체를 Pretendard로 교체 (+ Hang 스크립트, theme3의 맑은 고딕 latin → Arial/Roboto Light)
- [x] 실제 폰트로 줄바꿈을 계산해 넘침 판단 → 축소 → 재배치 → 분할 (io_steps는 행 단위 자동 분할. 요약은 Claude 판단)
- [x] `render.sh` → PNG → Claude 검수(넘침, 겹침, 폰트 대체) → 재생성
- [x] 보고서: 사용 변형, 추정 레이아웃 목록, 요약·분할 내역, 수작업 필요 슬라이드, 테마 폰트 교체, 렌더 폰트 대체
- [x] (선택) 공식 pptx 스킬 `validate.py` — 생성 슬라이드는 통과. `ppt/authors.xml` 오류 1건은 base.pptx에서 온 것(검사기에 2018 스키마 없음)
- 결과: spacemap-flow 10장 → 11장, 두 변형 모두 대체 폰트 0. 축소·자동 분할이 필요한 장 없음.
  M2 plan.json이 워크트리 삭제로 사라져 PLAN의 M2 결과대로 다시 작성(check_plan 통과).
  검수에서 고친 것: latin 폰트에 없는 글자(①, →) → ea 폰트 런으로 분리, 불릿 `buFont` 명시, 상태 배지 폭을 글자에 맞춤,
  카드 제목은 원본 색 대신 스타일 색.
- 도식 3장(S2·S4·S7): 처음엔 image_full(원본 렌더)이었으나 다크 덱에서 흰 판으로 보여, 사용자 결정으로 `diagram_restyle`
  (원본 도형 XML 복제 → 배치 유지, 색·글꼴만 다시 칠함, DECISIONS #21)로 바꿨다. 배율 collision 1.03/0.84/0.82, srmc 0.99/0.92/0.89.
  수작업 필요 장은 없어졌다.
- 상자 스타일(DECISIONS #22): 샘플 상자를 측정해 `components.box_styles`로 옮기고, 입력 상자 12종을 `box_map`으로 대응(실측 7종 99곳,
  추정 5종 21곳). 색 검토 후 B안: H 사람 = 상태 상자 노랑(유일한 채움), D·A·N·핵심 파일 = 색 테두리 패널(추정), 외부·미정 = 그림 테두리 점선.
  상자 안 본문은 흰색. 패널:상태 채움 면적 10:1(샘플 9–13:1).
- 남은 판단: `card_grid.stats`는 카드 한 줄이라 아래 절반이 빈다. S4 "승인 큐" 상자의 `accept · reject · skip`이 두 줄로 꺾인다
  (Roboto가 맑은 고딕보다 넓음 — 높이 검사는 통과).
- 구현하지 않은 레이아웃: toc, intro_columns, visual_kv, image_pair_kv (입력에 없음 — 쓰이면 build_deck이 중단하고 알림).
  section/references/bullets/bullets_image/screenshot_steps는 영역 상자에 텍스트만 넣는 기본형.

### M4 — SKILL.md로 묶기   [완료 2026-10-13, 사용자 확인 대기]
- [x] `.claude/skills/slide-restyle/SKILL.md` — `design` / `convert` 흐름
- [x] description에 "샘플 PPT로 DESIGN.md를 만들고 다른 PPT를 그 디자인으로 변환" 명시
