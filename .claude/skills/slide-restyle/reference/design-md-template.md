# DESIGN.md 템플릿 (PPT용 awesome-design-md 변형)

`/slide-restyle design`이 만드는 DESIGN.md의 골격이다. awesome-design-md의 9개 섹션 순서를 유지하되
4·5·8장을 슬라이드에 맞게 바꿨다. 스크립트는 **`yaml design-tokens` 펜스 블록 하나만** 파싱하고,
나머지 서술은 Claude와 사람이 읽는다. 서술과 토큰이 다르면 토큰이 우선한다.

## 작성 규칙

- **출처 표시**: 모든 토큰에 `source`를 단다.
  - `measured` — 샘플에서 실측. `evidence`에 근거 슬라이드를 적는다 (`C:S5` = 덱 약칭 C의 5번 슬라이드).
  - `estimated` — 샘플에 없어서 추정. `basis`에 무엇에서 유도했는지 적는다. 서술에도 **(추정)** 이라고 쓴다.
  - 섹션 단위 기본값 `source:`를 두고, 다른 항목만 개별로 덮어써도 된다.
- **단위**: 길이는 inch(소수 2–3자리), 폰트는 pt, 색은 `#` 없는 6자리 hex, 투명도는 `alpha`(0–1, 1=불투명).
- **좌표계**: `meta.canonical_slide` 기준 `[x, y, w, h]`. 다른 크기의 샘플은 extract_design.py가 환산한 값을 쓴다.
- 색과 폰트는 이름(`accent`, `heading`)으로 정의하고 스타일·레이아웃에서는 이름으로 참조한다.
- 제외 요소(`design/exclude.yaml`)는 토큰에 넣지 않고 `meta.excluded`에 사유만 남긴다.

## 섹션 골격

```
# <디자인 이름> — DESIGN.md
(한 문단 요약: 출처 샘플, 기준 크기, 제외 요소, 미결정 사항)

```yaml design-tokens
meta: {name, version, samples: [{id, file, slides, size_in, scale}], canonical_slide: {w, h},
       units: {length: in, font: pt}, default_chrome, excluded: [...]}
colors:   {<name>: {hex, alpha?, role, source, evidence|basis}}
fonts:    {<role>: {latin, ea, qa_render, qa_reliable, source}}
type:     {<style>: {font, size, bold?, italic?, color, align?, caps?, line_spacing?, source, evidence}}
chrome:   {<variant>: {background: {...}, elements: [{id, kind, box, style|asset|line, slot?, text?}]}}
components: {table, chart, callout, bullets, image_frame, highlight, ...}
layouts:  {<id>: {use_for, chrome, source, evidence|basis, regions: {<name>: {box, style, ...}}, capacity: {...}}}
overflow: {order, min_size, split, summarize, per_layout: {...}}
```

## 1. Visual Theme & Atmosphere      — 분위기, 배경 처리, 밀도
## 2. Color Palette & Roles          — 표: 이름 / hex / 역할 / 출처
## 3. Typography Rules               — 폰트 역할, 크기 위계 표, 한글(ea) 폰트, QA 렌더 대체
## 4. Component Stylings             — 표, 차트, 콜아웃(패널/배지), 불릿, 이미지 프레임, 강조 표시
## 5. Layout Principles              — 레이아웃 카탈로그: 레이아웃마다 용도·영역 좌표·용량·출처
## 6. Depth & Elevation              — 그림자/테두리/투명 패널 규칙
## 7. Do's and Don'ts                — 샘플에서 관찰된 규칙과 금지 사항
## 8. Responsive Behavior (Overflow) — 넘침 처리 순서: 축소 → 분할 → 요약, 레이아웃별 한도
## 9. Agent Prompt Guide             — 변환 시 Claude가 따를 판단 규칙 (레이아웃 선택, 분류 기준)
```
