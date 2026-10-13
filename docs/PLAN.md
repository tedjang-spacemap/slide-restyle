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

### M2 — extract_content.py + 레이아웃 분류   [대기]
- 산출물: `extract_content.py`, `content.json` 스키마, 레이아웃 분류 결과 `work/<deck>/plan.json`
- 확인 기준: 입력 PPT 슬라이드별 분류표, 변환 불가 후보(SmartArt, 자유 배치 도식 등)
- 첫 입력: `inputs/spacemap-flow.pptx` (10장)
- [ ] `extract_content.py` — 슬라이드별 텍스트 계층, 이미지, 표, 도형 수, 노트 (결정적)
- [ ] `content.json` 스키마
- [ ] DESIGN.md 9장 규칙으로 분류 → `plan.json` (넘침 시 축소/분할/요약 판단 포함)
- [ ] 분류표 사용자 확인 — 카탈로그에 없는 레이아웃이 필요한지 점검

### M3 — build_deck.py + render 검수 루프   [예정]
- 산출물: `build_deck.py`, 검수 루프, `outputs/<deck>.report.md` 형식
- 확인 기준: 변환 결과 PPT와 PNG, 보고서
- [ ] `base.pptx` 위에 슬라이드 추가, `--variant`로 chrome 선택
- [ ] 베이스 테마의 ea 폰트 자체를 Pretendard로 교체 (런 단위 `<a:ea>` 지정이 빠져도 맑은 고딕으로 돌아가지 않게)
- [ ] 실제 폰트로 줄바꿈을 계산해 넘침 판단 → 축소 → 재배치 → 분할 (요약은 Claude 판단)
- [ ] `render.sh` → PNG → Claude 검수(넘침, 겹침, 폰트 대체) → 재생성
- [ ] 보고서: 사용 변형, 추정 레이아웃 목록, 요약·분할 내역, 수작업 필요 슬라이드
- [ ] (선택) 공식 pptx 스킬 `validate.py`로 OOXML 무결성 검사

### M4 — SKILL.md로 묶기   [예정]
- [ ] `.claude/skills/slide-restyle/SKILL.md` — `design` / `convert` 흐름
- [ ] description에 "샘플 PPT로 DESIGN.md를 만들고 다른 PPT를 그 디자인으로 변환" 명시
