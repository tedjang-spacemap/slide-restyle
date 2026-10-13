# plan.json — `slide-restyle/plan@1`

Claude가 content.json과 원본 렌더(`src_png/`)를 보고 DESIGN.md 9장 규칙으로 작성하는 변환 계획. `build_deck.py`(M3)의 입력이다.
작성 후 `check_plan.py`로 검사한다 — 원문 축자 여부, 원본 텍스트 누락, 레이아웃 한도.

## 원칙

- **텍스트는 참조로만 쓴다**: 모든 텍스트는 `{"t": "...", "src": "s3.7"}`. `t`는 그 shape 텍스트의 축자 부분 문자열이어야 한다
  (공백만 정규화). 나누고 옮기는 것은 되고, 고쳐 쓰는 것은 안 된다. 요약이 필요하면 `overflow.plan`에 `summarize`로 적고
  M3에서 보고서·노트에 원문을 남긴다.
- **원본 텍스트는 모두 어딘가에 있어야 한다**: 텍스트 참조로 쓰이거나, `image_full`·`diagram_restyle` 장의 `covers`에 들어가거나,
  `dropped`에 이유와 함께 적힌다(예: 페이지 번호 → chrome이 다시 매김).
- 런 서식(code, 색)은 plan에 옮기지 않는다. build_deck이 `src`로 원본 런을 찾아 쓴다. 원본 색을 디자인 색으로 바꾸는
  대응표만 `deck.color_map`에 적는다(판단은 plan, 적용은 build_deck). 대응이 `text`이거나 표에 없는 색은 영역 스타일 색을 쓴다.
  셀·각주·배너·칩 본문에만 적용하고, 제목·카드 제목·키커는 항상 스타일 색이다.

## 구조

```jsonc
{
  "schema": "slide-restyle/plan@1",
  "source": "inputs/x.pptx",
  "content": "content.json",          // plan.json 기준 상대 경로
  "design": "DESIGN.md",
  "variant": "collision",             // --variant 값. 레이아웃 선택은 변형과 무관, 실측/추정 표시만 달라짐
  "deck": {
    "title": {"t": "…", "src": "s1.2"},
    "tags": {"D": {"color": "status_blue", "src_hex": "1F5F8B", "label": {"t": "…", "src": "…"}, "note?": "…"}},
    "color_map": {"17212B": "text", "5B6876": "text_muted", "1F5F8B": "tag:D"},   // 원본 런 색 → 디자인 색 이름 / tag:<키>
    "line_map": {"17212B": "text_label", "5B6876": "rule"},     // diagram_restyle 선 색 (없으면 color_map, 그래도 없으면 rule_subtle)
    "box_map": [                                                 // 입력 상자 종류(채움·선 hex, null = 없음) → DESIGN.md components.box_styles
      {"match": {"fill": "E3EEF6", "line": "1F5F8B"}, "style": "status_blue", "why": "D 처리"},
      {"match": {"fill": "EFEAF8", "line": "6B4FA3"}, "style": "panel_line", "line": "tag:A", "why": "…"}
    ],
    "footer_left": "SPACEMAP Platform / …"
  },
  "slides": [Slide, ...]              // 출력 순서
}
```

### Slide

| 필드 | 설명 |
|---|---|
| `out`, `src` | 출력 장 번호, 원본 장 번호 (한 원본이 여러 출력이 될 수 있음) |
| `layout`, `layout_variant` | DESIGN.md `layouts` 키와 그 `variants` 키 |
| `split` | `{group, part, of}` — 한 원본을 나눈 경우. 키커 끝에 `· part/of` |
| `slots` | chrome 슬롯: `kicker`, `title`, `status`(+`color`: 상태색 이름), `lead`, `source_line` |
| `regions` | 레이아웃별 내용 (아래) |
| `covers` | `image_full`이 이미지로 보여 주거나 `diagram_restyle`이 원문 그대로 다시 그리는 shape id 목록 (또는 `"all"`) |
| `dropped` | `[{src, why}]` |
| `overflow` | `{risk: none/low/medium/high, plan: [shrink, reflow, split, summarize], note}` — M2의 판단은 추정, M3에서 실제 폰트로 다시 계산 |
| `manual` | 수작업 필요 사유 문자열, 없으면 null |
| `reason` | 분류 근거 (9장 규칙 번호) |

### 레이아웃별 regions

| layout | regions |
|---|---|
| `cover` | slots `kicker`/`title`/`meta`; `flow: [{text, dashed}]` (추정 영역), `legend: [{tag, label}]` (`tag`도 텍스트 참조) |
| `image_full` | `image: {file: "src_png/slide-NN.png", crop_in: [x, y, w, h], fit}` — crop은 원본 슬라이드 inch 좌표 |
| `diagram_restyle` | `diagram: {exclude: [shape id]}` — 빼고 나머지 최상위 도형을 다시 칠함, `footnote?`; `covers`에 다시 칠한 텍스트 도형 id |
| `io_steps` | `header: [in, step, out]`, `rows: [{in: {lines, secondary?}, step: {no?, headline, tags, lines, note, kind, dashed}, out: {lines, secondary?}}]`, `footnote?` — `secondary`/`note`는 io_note 스타일 |
| `card_grid` | `banner?: {heading, body, dashed}`, `cards: [{heading, body: [..]}]`, `legend?` |
| `card_grid.stats` | `cards: [{number, label, detail: [..]}]` |
| 그 밖 | DESIGN.md 해당 레이아웃의 `regions` 이름을 키로, 값은 텍스트 참조 또는 그 목록 |

`step.kind`는 대표 태그(테두리 색 결정), `dashed`는 원본 도형이 점선이었는지(미확정·사람 단계 표시 유지).
