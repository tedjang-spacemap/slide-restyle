# content.json — `slide-restyle/content@1`

`extract_content.py`가 입력 덱에서 만드는 결정적 추출 결과. 레이아웃은 정하지 않고, 분류(plan.json)에 필요한 사실만 담는다.
길이는 모두 EMU(914400 = 1″), 글자 크기는 pt, 색은 6자리 hex.

```jsonc
{
  "schema": "slide-restyle/content@1",
  "source": "inputs/x.pptx",
  "slide_size": {"w": 12192000, "h": 6858000, "w_in": 13.33, "h_in": 7.5},
  "units": {"box": "EMU [x, y, w, h]", "emu_per_in": 914400, "size": "pt"},
  "meta": {"title": null, "author": null, "slides": 10},
  "slides": [{
    "n": 1, "layout_name": "Blank", "hidden": false,
    "notes": "발표자 노트 원문" | null,
    "shapes": [Shape, ...],          // z 순서, 그룹은 풀어서 (group 필드로 소속 표시)
    "features": Features
  }],
  "deck_features": {"repeated_texts": [{"text": "...", "at": ["1:s1.11", "10:s10.9"]}]}   // 범례·푸터 후보
}
```

## Shape

| 필드 | 설명 |
|---|---|
| `id` | `s<슬라이드>.<순번>` — plan.json의 `src`가 이 값을 가리킨다 |
| `xml_id`, `name`, `z` | 원본 cNvPr id·이름, 그리기 순서 |
| `box` | `[x, y, w, h]` EMU, 그룹 변환 적용 후 슬라이드 좌표 |
| `group` | 소속 그룹 shape의 id (그룹 자체도 `kind: group`으로 기록) |
| `ph`, `ph_idx` | placeholder 종류(`TITLE`, `BODY` …)와 idx |
| `kind` | `text`(채움·선 없는 텍스트상자) · `shape`(보이는 도형, 텍스트 있을 수 있음) · `connector` · `picture` · `table` · `chart` · `smartart` · `ole` · `group` · `empty` |
| `geom` | prstGeom 이름(`roundRect` …) 또는 `custom` |
| `fill`, `line` | `{hex, alpha?}` / `{hex, width_pt, dash}` — 직접 지정 또는 테마 style 참조를 해석한 값 |
| `text` | 아래 Text |
| `connector` | `{begin, end, from, to, headEnd?, tailEnd?}` — 끝점(EMU)과 닿은 shape id(가장 가까운 테두리 0.15″ 이내 또는 stCxn/endCxn). 화살촉이 begin에만 있으면 from/to를 뒤집는다 |
| `picture` | `{file: "media/<md5[:8]>.<ext>", md5, ext, px, crop}` |
| `table` | `{n_rows, n_cols, col_w, first_row_header, rows: [{h, cells: [{text, fill?, gridSpan?, rowSpan?, merged_into_prev?}]}]}` — 셀 텍스트는 자르지 않는다 |
| `chart` | `{chart_type, title, series: [{name, categories, values}]}` |
| `smartart` | `{texts: [...]}` — 데이터 모델 순서의 모든 텍스트 |

### Text

```jsonc
{"plain": "문단은 \n으로 이은 전체 텍스트",
 "anchor": "t", "wrap": "square", "insets": {"lIns": 0}, "autofit": "spAutoFit", "chars": 42,
 "paragraphs": [{"level": 0, "align": "l", "bullet": "•" | "autonum:arabicPeriod" | null, "marL": 0, "indent": 0,
                 "runs": [{"text": "…", "font": "Consolas", "ea": "Malgun Gothic", "size": 11.5,
                           "bold": false, "italic": false, "color": {"hex": "1F5F8B"}, "code": true}]}]}
```
`size`는 런 → 문단 → lstStyle → 마스터 txStyles 순으로 상속을 따라간 값이다(`extract_design.read_text`). `code: true`는 고정폭 글꼴 런(스크립트명·경로).

## Features (판단 아님, 측정된 사실)

| 필드 | 설명 |
|---|---|
| `counts` | kind별 개수 |
| `chars` | 슬라이드 전체 글자 수 |
| `roles.title` | `{id, by, text, split?, section?}` — placeholder › 이름 `Title` › 위 45% 안의 가장 큰 글자 순. `split` = 끝부분 런이 크기·색이 달라 `{main, status, status_color}`로 나뉠 때. `section` = `① 섹션 — 소제목` 꼴의 `{marker, section, sub?}` |
| `roles.lead` | 제목 바로 아래(0.3″ 이내, 왼쪽 정렬 일치) 텍스트상자, 또는 이름 `Lead`/subtitle placeholder |
| `roles.page_number` | 아래 15% 안의 숫자만 있는 텍스트 |
| `grids` | 같은 크기 상자의 직사각 배열(2×2 이상, 또는 3개 이상 한 줄): `{rows, cols, col_x, col_w, row_y, row_h, cells: [[id]], headers: [id|null]}`. 같은 열의 셀은 종류·폭·높이가 같아야 한다. `headers` = 첫 행 바로 위(0.25″ 이내)에 같은 왼쪽 선으로 놓인 텍스트 |
| `chains` | 연결선으로 끝과 끝이 이어진 shape 경로(3개 이상), 예: 표지의 ①→②→③→④ |
| `connectors_linked` / `connectors_dangling` | 양 끝이 shape에 닿은 연결선 수 / 아닌 수 |
| `overlaps` | 보이는 상자끼리 겹친 쌍의 수 |
| `loose_shapes` | 격자·역할에 속하지 않은 텍스트/도형 id — 자유 도식 정도의 지표 |
| `style_classes` | `(geom, fill, line, dash)`가 같은 도형 묶음, 많은 순 — 색으로 구분한 종류(D/N/A/H 등) 탐지용 |
