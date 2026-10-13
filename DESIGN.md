# SPACEMAP Dark Training Deck — DESIGN.md

SPACEMAP 교육용 덱 두 개(`C` = Collision Avoidance User Guide, 21장 / `R` = SRMC Interference Cases, 9장)에서
실측한 디자인 명세. 기준 캔버스는 C의 13.335″×7.5″(16:9)이며, R(10″×5.625″)의 좌표와 폰트 크기는
×1.3335로 환산했다. 측정 원자료는 `design/design.json`, 요약은 `design/summary.md`, 기준 마스터는
`design/base.pptx`·`base.potx`, 로고·배경 이미지는 `design/assets/`에 있다.

- **제외 요소**: "MINISTRY OF SCIENCE, TECHNOLOGY AND INNOVATION / MALAYSIAN SPACE AGENCY" 배너(21곳)는
  사용자 요청으로 디자인에서 뺐다(`design/exclude.yaml`). 배너가 있던 자리(좌상단)는 비워 둔다.
- **출처 표기**: `measured` = 샘플 실측(근거 `C:S5` 등), `estimated` = 샘플에 없어 추정(근거 `basis`).
  본문에서는 추정 항목에 **(추정)** 을 붙인다.
- **사용자 결정 (2026-10-13)**: ① 변형 2개 — **collision**(기본)과 **srmc**. 색·폰트·컴포넌트·표지는 공유하고 슬라이드
  틀(배경·로고·키커/제목·푸터)만 다르다. 변환 시 `--variant srmc`로 고른다. ② 표지·마지막 장의 초록색
  "Professional Training Programme / SSA | SDA | STM / 날짜" 이벤트 블록은 배너와 같이 **제외**.
  ③ 한글(ea) 폰트는 **Pretendard** — 샘플 테마의 맑은 고딕 대신 쓰는 사용자 지정값이다.

```yaml design-tokens
meta:
  name: spacemap-dark-training
  version: 0.1-draft
  samples:
    - {id: C, file: "Collision_Avoidance_User_Guide_EN_SPACEMAP(MYSA).pptx", slides: 21, size_in: [13.335, 7.5], scale: 1.0}
    - {id: R, file: SRMC_Interference_Cases.pptx, slides: 9, size_in: [10.0, 5.625], scale: 1.3335}
  canonical_slide: {w: 13.335, h: 7.5}
  units: {length: in, font: pt, color: hex6, alpha: "0-1 (1=opaque)"}
  base_template: design/base.pptx      # masters/layouts kept: "Custom Layout"(C M2), "빈 화면"(C M1)
  base_layout: Custom Layout           # every generated slide starts from this layout
  default_chrome: collision            # variant used when --variant is not given
  variants: [collision, srmc]          # = keys of `chrome`; choose with /slide-restyle convert … --variant srmc
  excluded:
    - {what: "MINISTRY OF SCIENCE… / MALAYSIAN SPACE AGENCY banner", where: "C:S1,S3–S21 top-left; R:S1 top-left", reason: user request 2026-10-13}
    - {what: "green event block 'Professional Training Programme / SSA | SDA | STM / dates' (00B050)", where: "C:S1,S21; R:S1 box [0.55, 5.84, 6.33, 1.19]", reason: user request 2026-10-13}

colors:
  bg:            {hex: "030509", role: slide background (opaque full-slide rect), source: measured, evidence: "C:S1,S3–S21; R:S1"}
  text:          {hex: FFFFFF, role: primary text, source: measured, evidence: "C, R all"}
  text_muted:    {hex: C6C6C6, role: secondary text, footer, notes, source: measured, evidence: "C:S3–S20"}
  text_label:    {hex: B8C0CC, role: table header / key labels / captions / source line, source: measured, evidence: "R:S2–S9"}
  accent:        {hex: FFC000, role: titles, numbers, page number, highlights, source: measured, evidence: "C all; R kicker"}
  rule:          {hex: "808080", role: header divider line, source: measured, evidence: "C:S1,S3–S21"}
  rule_title:    {hex: "848484", role: srmc title underline 2.25pt, source: measured, evidence: "R:S2–S9 (layout line)"}
  rule_subtle:   {hex: 4B5563, role: row separators, image frame, source: measured, evidence: "R:S2–S8"}
  panel:         {hex: 0B1426, alpha: 0.82, role: visual panel fill (roundRect), source: measured, evidence: "R:S4–S8"}
  panel_alt:     {hex: "182233", alpha: 0.82, role: timeline panel fill, source: measured, evidence: "R:S7"}
  status_amber:    {hex: FFC000, fill: 4D3A0A, row_tint: {hex: 2E2610, alpha: 0.82}, role: "status: unconfirmed / proximity", source: measured, evidence: "R:S2,S3,S5"}
  status_red:      {hex: F87171, fill: 5A1E22, row_tint: {hex: 2E1517, alpha: 0.82}, role: "status: satellite / alert / 'after'", source: measured, evidence: "R:S2,S3,S4,S6"}
  status_blue:     {hex: 60A5FA, fill: 1E3A66, role: "status: ground / neutral info", source: measured, evidence: "R:S2,S6,S7,S8"}
  diagram_blue:  {hex: 7FA8F5, role: diagram — own asset, source: measured, evidence: "R:S4–S8"}
  diagram_pink:  {hex: FFB4B6, role: diagram — other asset label, source: measured, evidence: "R:S4,S5"}
  diagram_label: {hex: B8C4DA, role: diagram — ground label, source: measured, evidence: "R:S4–S8"}
  diagram_faint: {hex: 7C8BA8, role: diagram — annotation, source: measured, evidence: "R:S4,S5"}
  diagram_arc:   {hex: 2A3A5C, role: diagram — orbit arc, source: measured, evidence: "R:S4"}
  highlight_red: {hex: FF0000, role: screenshot callout box (3pt roundRect), source: measured, evidence: "C:S2"}
  chart_series:  {hexes: [FFC000, 60A5FA, F87171, 7FA8F5, C6C6C6, FFB4B6], source: estimated, basis: "accent + status/diagram colors in measured frequency order"}

fonts:
  heading: {latin: Arial, ea: Pretendard, bold: true, qa_render: "Arial + Pretendard Bold (installed)", qa_reliable: true, source: measured, evidence: "C titles/numbers; R all", ea_source: "user choice 2026-10-13 (sample theme ea = 맑은 고딕)"}
  body:    {latin: Roboto Light, ea: Pretendard Light, qa_render: "Roboto Light (static, ./.fonts) + Pretendard Light", qa_reliable: true, source: measured, evidence: "C body/kicker/footer", ea_source: "user choice 2026-10-13; Light weight matches Roboto Light"}
  heading_regular: {latin: Arial, ea: Pretendard, bold: false, qa_reliable: true, source: measured, evidence: "R 본문 전체 (Arial regular)", ea_source: "user choice 2026-10-13"}
  ea_note: "한글은 런마다 ea 폰트를 Pretendard로 명시한다(테마 ea는 맑은 고딕이라 명시하지 않으면 그쪽으로 렌더됨). 받는 PC에 Pretendard가 없으면 대체 폰트로 표시된다."

type:
  # --- chrome / titles (C) ---
  kicker:        {font: body, size: 12, italic: true, color: text, align: l, source: measured, evidence: "C:S3–S20", note: "형식 '< 섹션 · 하위 · n/m >'"}
  title:         {font: heading, size: 26, bold: true, color: accent, align: l, source: measured, evidence: "C:S3–S20"}
  footer:        {font: body, size: 9.5, color: text_muted, source: measured, evidence: "C:S3–S20"}
  page_number:   {font: heading, size: 14, bold: true, color: accent, align: r, source: measured, evidence: "C:S3–S20"}
  # --- cover / closing ---
  cover_kicker:  {font: body, size: 20, italic: true, color: text, line_spacing: 1.1, source: measured, evidence: "C:S1; R:S1"}
  cover_title:   {font: heading, ea_first: true, size: 32, bold: true, color: accent, bullet: "•", marL: 0.5, indent: -0.5, line_spacing: 1.1, source: measured, evidence: "C:S1; R:S1", note: "샘플은 런 폰트가 +mn-ea(맑은 고딕). 변환 시 heading(Arial Bold + Pretendard)로 지정"}
  cover_meta:    {font: body, size: 16, color: text, line_spacing: 1.1, source: measured, evidence: "C:S1; R:S1", note: "마지막 줄 URL은 italic"}
  closing_title: {font: heading, size: 60, bold: true, color: accent, source: measured, evidence: "C:S21"}
  closing_sub:   {font: body, size: 28, color: text, source: measured, evidence: "C:S21"}
  closing_url:   {font: body, size: 16, italic: true, color: text_muted, source: measured, evidence: "C:S21"}
  section_title: {font: heading, size: 44, bold: true, color: accent, source: estimated, basis: "closing_title(60)과 title(26) 사이; 샘플에 섹션 구분 슬라이드 없음"}
  # --- content (C) ---
  lead:          {font: body, size: 23, color: text, source: measured, evidence: "C:S4"}
  col_heading:   {font: heading, size: 22, bold: true, color: accent, source: measured, evidence: "C:S4"}
  col_body:      {font: body, size: 20, color: text, source: measured, evidence: "C:S4"}
  toc_number:    {font: heading, size: 23, bold: true, color: accent, source: measured, evidence: "C:S3"}
  toc_heading:   {font: heading, size: 20.5, bold: true, color: text, source: measured, evidence: "C:S3"}
  toc_desc:      {font: body, size: 16.5, color: text_muted, source: measured, evidence: "C:S3"}
  step_number:   {font: heading, size: 34, bold: true, color: accent, source: measured, evidence: "C:S5–S9,S11–S20"}
  step_headline: {font: heading, size: 17, bold: true, color: text, source: measured, evidence: "C:S5–S9,S11–S20"}
  body:          {font: body, size: 17, color: text, source: measured, evidence: "C:S5–S20 (가장 많이 쓰인 본문 스타일)"}
  note:          {font: body, size: 15.5, color: text_muted, source: measured, evidence: "C:S5–S20"}
  list_number:   {font: heading, size: 23, bold: true, color: accent, source: measured, evidence: "C:S10"}
  bullet_l1:     {font: body, size: 20, color: text, source: estimated, basis: "col_body 크기 + C:S1 불릿 문자"}
  bullet_l2:     {font: body, size: 17, color: text_muted, source: estimated, basis: "body/note 위계"}
  # --- data (R, scaled) ---
  table_header:  {font: heading, size: 13.3, bold: true, color: text_label, source: measured, evidence: "R:S2"}
  table_cell:    {font: heading, size: 12, color: text, source: measured, evidence: "R:S2"}
  table_key:     {font: heading, size: 13.3, bold: true, color: text, source: measured, evidence: "R:S2 (첫 열)"}
  kv_label:      {font: heading, size: 13.3, bold: true, color: text_label, source: measured, evidence: "R:S3–S8"}
  kv_value:      {font: heading, size: 14.7, color: text, emphasis_bold: true, source: measured, evidence: "R:S3–S8 (핵심 행은 bold)"}
  pill:          {font: heading, size: 12, bold: true, align: ctr, color: "status.hex", source: measured, evidence: "R:S2–S8"}
  caption:       {font: heading, size: 10.7, color: text_label, source: measured, evidence: "R:S3"}
  source_line:   {font: heading, size: 10.7, color: text_label, source: measured, evidence: "R:S2–S8"}
  panel_heading: {font: heading, size: 14.7, bold: true, color: text, source: measured, evidence: "R:S3,S7"}
  diagram_label: {font: heading, size: 12, bold: true, color: diagram_blue, source: measured, evidence: "R:S4–S8"}
  srmc_kicker:   {font: heading, size: 13.3, bold: true, color: accent, caps: true, source: measured, evidence: "R:S2–S9"}
  srmc_title:    {font: heading, size: 28, bold: true, color: text, line_spacing: 0.9, source: measured, evidence: "R:S2–S9"}
  reference:     {font: heading, size: 13.3, color: text, label_bold: true, para_space_before: 7.1, source: measured, evidence: "R:S9"}
  # --- process / cards (estimated, added M2 for diagram-style inputs) ---
  lead_compact:  {font: body, size: 16, color: text_muted, source: estimated, basis: "lead(23)는 도식 위 한 줄 설명에 너무 큼. note(15.5)보다 약간 크게"}
  io_step_no:    {font: heading, size: 12, bold: true, color: accent, source: estimated, basis: "step_number(34)의 색·굵기를 셀 크기로 축소"}
  io_headline:   {font: heading, size: 13.3, bold: true, color: text, source: estimated, basis: "table_key(13.3)와 같은 크기"}
  io_line:       {font: body, size: 11.5, color: text, source: estimated, basis: "table_cell(12)보다 한 단계 작게; 셀 안 3줄을 0.63″에 넣기 위함"}
  io_note:       {font: body, size: 10.5, color: text_muted, source: estimated, basis: "caption(10.7) 근처, overflow.floors.table_cell(10.5)"}
  tag:           {font: heading, size: 11, bold: true, color: "tag.hex", source: estimated, basis: "pill 글자 규칙(Bold, 상태색)을 글자 태그에 적용"}
  card_heading:  {font: heading, size: 17, bold: true, color: accent, source: estimated, basis: "col_heading(22)을 6칸 그리드용으로 축소, step_headline(17) 크기"}
  card_body:     {font: body, size: 14, color: text, source: estimated, basis: "overflow.floors.body(14)"}
  stat_number:   {font: heading, size: 34, bold: true, color: accent, source: estimated, basis: "step_number(34) 실측 스타일 재사용"}
  stat_label:    {font: heading, size: 15, bold: true, color: text, source: estimated, basis: "panel_heading(14.7) 근처"}
  legend:        {font: body, size: 10.5, color: text_muted, source: estimated, basis: "footer(9.5)보다 약간 크게; 글자 태그는 tag 스타일"}
  chip:          {font: heading, size: 13, bold: true, color: text, align: ctr, source: estimated, basis: "pill(12)보다 약간 크게, 표지 흐름 칩"}

chrome:
  # Each key is a design variant. Shared: colors, fonts, type, components, cover/closing.
  # A variant swaps the slide frame (background, logo, kicker/title, footer) and may override styles.
  collision:   # DEFAULT — C:S3–S20 (18/21 slides)
    source: measured
    evidence: "C:S3–S20"
    background:
      - {kind: rect, box: [0, 0, 13.335, 7.5], fill: bg}   # covers master earth image
    text_box_defaults: {insets: [0, 0, 0, 0], autofit: none, anchor: t, para_space: 0}
    elements:
      - {id: logo,        kind: picture, asset: design/assets/a572a534.png, box: [11.09, 0.35, 1.82, 0.25]}
      - {id: header_rule, kind: line, box: [0.42, 0.92, 12.5, 0], color: rule, width_pt: 0.75}
      - {id: kicker,      kind: text, slot: kicker, box: [0.42, 1.04, 10.26, 0.28], style: kicker, single_line: true, format: "< {text} >"}
      - {id: title,       kind: text, slot: title,  box: [0.42, 1.33, 12.5, 0.5],  style: title}
      - {id: status,      kind: pill, slot: status, box: [8.79, 1.36, 4.13, 0.43], optional: true, title_w_when_present: 8.2, source: estimated, basis: "R의 헤더 우상단 배지를 C 제목 줄 오른쪽 끝으로 이동"}
      - {id: footer_l,    kind: text, slot: footer_left,  box: [0.42, 7.16, 6.77, 0.19], style: footer, align: l, default: "SPACEMAP Platform / <deck title>"}
      - {id: footer_r,    kind: text, slot: footer_right, box: [8.8, 7.16, 3.23, 0.19], style: footer, align: r, default: platform.spacemap42.com}
      - {id: page_no,     kind: text, slot: page_number,  box: [12.42, 7.08, 0.5, 0.29], style: page_number, format: "{n:02d}"}
    content_area: [0.42, 2.12, 12.5, 4.66]   # top = screenshot top C:S5–S20, bottom = max screenshot bottom 6.78
    style_overrides: {}
  srmc:        # ALTERNATIVE — R:S2–S9 (scaled ×1.3335)
    source: measured
    evidence: "R:S2–S9"
    background:
      - {kind: picture, asset: design/assets/6f74de44.jpg, box: [0, 0, 13.335, 7.5]}   # dark earth image, no overlay
    text_box_defaults: {insets: [0, 0, 0, 0], autofit: none, anchor: ctr, para_space: 0}
    elements:
      - {id: kicker,     kind: text, slot: kicker, box: [0.67, 0.16, 8.0, 0.4], style: srmc_kicker, format: "{TEXT}"}
      - {id: title,      kind: text, slot: title,  box: [0.67, 0.5, 12.0, 0.66], style: srmc_title, insets: [0, 0.05, 0.1, 0.05]}
      - {id: title_rule, kind: line, box: [0.56, 1.23, 12.22, 0], color: rule_title, width_pt: 2.25}
      - {id: status,     kind: pill, slot: status, box: [8.53, 0.16, 4.13, 0.43], optional: true, kicker_w_when_present: 6.0, evidence: "R:S3–S8"}
      - {id: logo,       kind: picture, asset: design/assets/c980e851.png, box: [0.13, 7.03, 1.43, 0.32]}
      - {id: source,     kind: text, slot: source_line, box: [1.75, 7.02, 10.9, 0.33], style: source_line, optional: true}
    content_area: [0.67, 1.63, 12.0, 5.1]    # top = table header R:S2, bottom = panel bottom R:S4–S8 (6.73)
    no_page_number: true                      # R has none
    style_overrides:                          # applied to every layout when this variant is chosen
      source: estimated
      basis: "R는 모든 텍스트가 Arial이고 보조 텍스트가 B8C0CC. C 전용 스타일을 R 관례로 치환"
      fonts: {body: heading_regular}          # Roboto Light → Arial regular (+ Pretendard Regular)
      colors: {text_muted: text_label}        # C6C6C6 → B8C0CC

variant_mapping:
  # How a layout measured under one variant is placed under another.
  method: linear content_area map
  rule: "x' = A.x + (x − N.x)·A.w/N.w, y' = A.y + (y − N.y)·A.h/N.h; widths ×A.w/N.w, heights ×A.h/N.h (N = layout.native 변형의 content_area, A = 선택한 변형의 content_area). box, grid, columns, flow, col_template, row_pitch, label_w/value_x/value_w/row_h 등 모든 길이에 적용. 폰트 크기는 그대로."
  provenance: "native 변형에서는 레이아웃의 source 그대로, 다른 변형에서는 항상 estimated"
  applies_to: "native가 collision 또는 srmc인 레이아웃. native: none(표지·마지막 장·전체 이미지)은 변형과 무관하게 같은 좌표"
components:
  table:
    source: measured
    evidence: "R:S2 (텍스트 상자로 그린 표 → 네이티브 표로 옮길 때 같은 시각 규칙 적용)"
    header: {style: table_header, fill: none, rule_below: {color: text, width_pt: 0.86}}
    row: {style: table_cell, first_col_style: table_key, fill: none, rule_below: {color: rule_subtle, width_pt: 0.86}, anchor: ctr, cell_insets: [0, 0, 0, 0]}
    row_tint: {optional: true, use: "상태가 있는 행만 status.row_tint", evidence: "R:S2 rows 1–3"}
    vertical_rules: none
    status_column: {component: pill}
  key_value:
    source: measured
    evidence: "R:S3–S8"
    label: {style: kv_label, width_ratio: 0.29}
    value: {style: kv_value}
    rule_below: {color: rule_subtle, width_pt: 0.86}
    anchor: ctr
  pill:
    source: measured
    evidence: "R:S2–S8"
    shape: roundRect
    corner_adj: {header: 18750, in_table: 17857}   # OOXML roundRect adj (0–50000), R:S4 / R:S2
    fill: "status.fill"
    text: {style: pill, color: "status.hex"}
    sizes: {header: [4.13, 0.43], in_table: [1.6, 0.51]}
    fit_width: {min_w: 1.6, pad: 0.25, align: r, source: estimated, basis: "R:S3–S8 배지는 긴 문구라 4.13″. 짧은 상태 문구(‘활성’)는 글자 폭 + 여백, 최소 in_table 폭 1.6″, 오른쪽 끝은 슬롯에 맞춤 (M3)"}
  panel:
    source: measured
    evidence: "R:S4–S8"
    shape: roundRect
    corner_adj: 2222          # nearly square corners, R:S4
    fill: panel
    line: {color: panel, width_pt: 0.86}
  chart:
    source: estimated
    basis: "샘플에 차트 없음. 팔레트·라벨 규칙을 표/다이어그램에서 유도"
    series_colors: chart_series
    axis_labels: {font: body, size: 12, color: text_muted}
    gridlines: {color: rule_subtle, width_pt: 0.75}
    data_labels: {font: heading, size: 12, bold: true, color: text}
    legend: {font: body, size: 12, color: text_muted, position: bottom}
    background: none
  bullets:
    source: estimated
    basis: "본문에 불릿 슬라이드 없음. 문자 '•'는 C:S1 표지 제목에서 실측"
    char: "•"
    color: accent
    l1: {style: bullet_l1, marL: 0.3, indent: -0.3, space_before: 8}
    l2: {style: bullet_l2, marL: 0.6, indent: -0.25, space_before: 4}
  image_frame:
    screenshot: {border: none, fit: contain, align: top-left, source: measured, evidence: "C:S5–S20"}
    figure:     {border: {color: rule_subtle, width_pt: 0.65}, caption: {style: caption, gap: 0.05}, header: {style: panel_heading, gap: 0.06}, source: measured, evidence: "R:S3"}
  highlight:
    rect:    {shape: roundRect, line: {color: highlight_red, width_pt: 3}, fill: none, source: measured, evidence: "C:S2"}
    ellipse: {shape: ellipse, line: {color: status_red, width_pt: 1.72, dash: dash}, fill: none, source: measured, evidence: "R:S3"}
  tag_chip:
    source: estimated
    basis: "입력 덱의 글자 태그(D/N/A/H 등)를 색 글자로 옮김. 색은 상태색 순서로 배정"
    shape: none                       # bold colored letter, no fill (source decks use colored letters)
    palette: [status_blue, status_amber, diagram_pink, text_label, status_red, diagram_blue]
    note: "태그→색 대응은 덱마다 plan.json의 deck.tags에 적는다. 원본 색과 가까운 색을 고른다"
  step_cell:
    source: estimated
    basis: "panel 컴포넌트(R:S4–S8 실측)를 표 셀 크기로 사용. 칸 종류별 색은 box_styles + plan.json deck.box_map (M3)"
    shape: roundRect
    corner_adj: 6000
    fill: panel
    line: {color: rule_subtle, width_pt: 0.75}
    insets: [0.1, 0.04, 0.1, 0.04]
    anchor: ctr
  box_styles:
    # 샘플 슬라이드의 상자(도형) 스타일 — design.json box_styles 실측. 입력 덱의 상자 종류는 plan.json deck.box_map으로
    # 이 중 하나에 대응시킨다(도식·단계 셀·카드·표지 칩 공통). fill/line은 색 이름 또는 `<색>.fill`, text는 상자 안 기본 글자색
    # (원본에서 본문색이던 런만 바뀜). 모양(geom)은 도식에서는 원본 유지, 새로 그리는 칸에서는 이 값.
    # dash(선 모양 규칙, 스타일마다): keep = 원본 상자가 점선이면 점선 유지, none = 항상 실선, dash = 항상 점선. 생략하면 keep.
    #   샘플 상자에는 점선이 없어(점선은 R:S3 타원·R:S4 호뿐) 이 규칙 자체가 추정이다. 점선 처리를 바꾸려면 여기만 고친다.
    # 상태 상자 안 본문은 흰색·보조색 그대로 둔다: 샘플의 상태색 글자는 12pt Bold 짧은 배지 문구뿐이고, 본문에 쓰면 7장 Don't(노랑 본문)와
    # 대비 기준(60A5FA/1E3A66 = 4.46 < 4.5)에 걸린다. 상태 채움은 예외 단계에만 써서 샘플의 면적 비율(패널:상태색 ≈ 9–13:1)을 지킨다 (M3).
    panel:        {geom: roundRect, corner_adj: 2222, fill: panel, line: panel, line_pt: 0.86, dash: keep, source: measured, evidence: "R:S4–S8"}
    panel_alt:    {geom: roundRect, corner_adj: 2222, fill: panel_alt, line: rule_subtle, line_pt: 0.86, dash: keep, source: measured, evidence: "R:S7"}
    status_blue:  {geom: roundRect, corner_adj: 17857, fill: status_blue.fill, line: status_blue.fill, line_pt: 0.86, dash: keep, source: measured, evidence: "R:S2,S6–S8 (Arial 12 Bold 60A5FA)"}
    status_amber: {geom: roundRect, corner_adj: 17857, fill: status_amber.fill, line: status_amber.fill, line_pt: 0.86, dash: keep, source: measured, evidence: "R:S2,S3,S5 (Arial 12 Bold FFC000)"}
    status_red:   {geom: roundRect, corner_adj: 17857, fill: status_red.fill, line: status_red.fill, line_pt: 0.86, dash: keep, source: measured, evidence: "R:S2,S4 (Arial 12 Bold F87171)"}
    frame:        {geom: rect, fill: none, line: rule_subtle, line_pt: 0.65, dash: keep, source: measured, evidence: "R:S3 (그림 테두리)"}
    callout:      {geom: rect, fill: FFFFFF, line: status_red, line_pt: 0.65, text: status_red, dash: keep, source: measured, evidence: "R:S3 (흰 바탕 라벨)"}
    panel_line:   {geom: roundRect, corner_adj: 2222, fill: panel, line: "<box_map.line>", line_pt: 1.25, dash: keep, source: estimated, basis: "샘플에 색 테두리 상자 없음. 상태색이 없는 종류(보라·초록·강조 파일)는 panel + 색 테두리"}
  arrow:
    source: estimated
    basis: "샘플에 화살표 없음. rule 색 선 + 삼각 화살촉"
    line: {color: rule, width_pt: 1.0}
    head: triangle
  device_art:
    box: [10.22, 4.71, 2.29, 2.41]
    pieces:   # z-order: back to front
      - {asset: design/assets/ba59c3ca.png, box: [11.31, 4.71, 1.2, 1.5]}
      - {asset: design/assets/97609c37.png, box: [10.22, 4.95, 2.14, 2.17]}
    source: measured
    evidence: "C:S1,S21; R:S1 (그룹: 태블릿 2대 이미지)"

layouts:
  cover:
    use_for: 덱 첫 장
    native: none
    source: measured
    evidence: "C:S1; R:S1"
    background: [{kind: rect, box: [0, 0, 13.335, 7.5], fill: bg}]
    regions:
      logo:   {kind: picture, asset: design/assets/a572a534.png, box: [8.77, 0.85, 3.74, 0.52]}
      rule:   {kind: line, box: [0.89, 2.05, 11.55, 0], color: rule, width_pt: 0.75}
      kicker: {box: [0.55, 2.44, 7.13, 0.45], style: cover_kicker, insets: [0.1, 0.05, 0.1, 0.05], autofit: shape}
      title:  {box: [0.55, 3.13, 8.47, 1.24], style: cover_title, insets: [0.05, 0.05, 0.05, 0.05], autofit: shape, max_lines: 2}
      meta:   {box: [0.55, 4.6, 7.13, 0.97], style: cover_meta, insets: [0.1, 0.05, 0.1, 0.05], autofit: shape, max_lines: 3}
      art:    {component: device_art}
      flow:   {box: [0.55, 5.85, 9.45, 0.5], optional: true, source: estimated, basis: "meta 아래 빈 띠(5.57–6.6)에 칩 행 배치, device_art(x 10.22) 왼쪽에서 끝남", item: {style: chip, component: step_cell, max_items: 5, gap: 0.35, connector: arrow}}
      legend: {box: [0.55, 6.62, 9.45, 0.3], optional: true, style: legend, source: estimated, basis: "flow 아래 한 줄"}
    capacity: {title_lines: 2, title_chars_per_line: 30, meta_lines: 3, flow_items: 5}
  closing:
    use_for: 마지막 감사/연락처 장
    native: none
    source: measured
    evidence: "C:S21"
    background: [{kind: rect, box: [0, 0, 13.335, 7.5], fill: bg}]
    regions:
      logo:     {kind: picture, asset: design/assets/a572a534.png, box: [8.77, 0.85, 3.74, 0.52]}
      rule:     {kind: line, box: [0.89, 2.05, 11.55, 0], color: rule, width_pt: 0.75}
      title:    {box: [0.55, 2.5, 9.3, 1.09], style: closing_title}
      subtitle: {box: [0.55, 3.67, 9.3, 0.52], style: closing_sub}
      url:      {box: [0.55, 4.25, 9.3, 0.36], style: closing_url}
      art:      {component: device_art}
  section:
    use_for: 장 구분(섹션 머리) 슬라이드
    native: collision
    omit_slots: [title]
    source: estimated
    basis: "샘플에 섹션 슬라이드 없음. closing의 제목 위치/위계를 콘텐츠 크롬 안으로 옮김"
    regions:
      number: {box: [0.42, 2.9, 3.0, 0.6], style: toc_number}
      title:  {box: [0.42, 3.5, 12.5, 1.0], style: section_title}
      desc:   {box: [0.42, 4.6, 9.0, 0.8], style: toc_desc}
  toc:
    use_for: 목차 / 개요 목록 (항목 = 번호 + 제목 + 한 줄 설명)
    native: collision
    source: measured
    evidence: "C:S3"
    grid: {cols: 2, rows: 3, col_x: [0.42, 6.88], row_y: [2.26, 3.81, 5.36]}
    item:
      number:  {dx: 0, dy: 0, w: 1.04, h: 0.42, style: toc_number}
      heading: {dx: 1.25, dy: 0, w: 4.73, h: 0.47, style: toc_heading}
      desc:    {dx: 1.25, dy: 0.53, w: 4.67, h: 0.75, style: toc_desc, max_lines: 2}
    capacity: {items: 6}
  intro_columns:
    use_for: 리드 문장 + 3개 병렬 항목(단계/특징)
    native: collision
    source: measured
    evidence: "C:S4"
    regions:
      lead: {box: [0.42, 2.19, 12.19, 0.8], style: lead, max_lines: 2}
    columns: {count: 3, x: [0.42, 4.71, 9.0], w: 3.91, heading: {y: 4.03, h: 0.77, style: col_heading}, body: {y: 4.96, h: 1.51, style: col_body}}
    variants:
      two_columns: {source: estimated, basis: "3단 그리드(열 3.91 + 간격 0.38)를 2단으로 재분할", x: [0.42, 6.86], w: 6.06}
  screenshot_steps:
    use_for: 큰 이미지(스크린샷/그림) + 오른쪽 설명. 이미지+텍스트 슬라이드의 기본형
    native: collision
    source: measured
    evidence: "C:S5–S20 (16장)"
    regions:
      image:    {box: [0.42, 2.12, 8.54, 4.66], fit: contain, align: top-left, component: image_frame.screenshot}
      step_no:  {box: [9.32, 2.07, 3.57, 0.62], style: step_number, optional: true}
      headline: {box: [9.32, 2.87, 3.57, 0.6], style: step_headline, optional: true, max_lines: 3}
      body:     {box: [9.32, 3.66, 3.57, 1.73], style: body, flow: true}
      note:     {box: [9.32, 4.72, 3.57, 0.81], style: note, optional: true, flow: true}
    flow: {column: [9.32, 2.07, 3.57, 4.71], gap: 0.15, order: [step_no, headline, body, note], note: "본문/노트 높이는 텍스트에 맞춰 위에서부터 쌓음 (C:S6,S9,S17 실측 y가 내용 따라 이동)"}
    variants:
      numbered_list: {source: measured, evidence: "C:S10", number: {x: 9.32, w: 0.33, style: list_number}, text: {x: 9.75, w: 3.15, style: body}, row_pitch: 0.83, start_y: 2.12, max_items: 4}
    capacity: {body_chars: 170, note_chars: 95}
  image_full:
    use_for: 세로로 긴 표/문서 이미지 한 장을 원본 그대로 보여줄 때, 변환 불가 슬라이드의 원본 이미지 삽입
    native: none
    source: measured
    evidence: "C:S2"
    background: [{kind: base_layout_bg}]   # earth image from master (b387d858), no dark rect
    regions:
      image: {box: [0, 0, 13.335, 7.5], fit: contain, align: center}
  table:
    use_for: 행 3–6개의 비교/목록 표
    native: srmc
    source: measured
    evidence: "R:S2 (텍스트 상자로 그린 6행 표)"
    regions:
      header_row: {box: [0.67, 1.63, 12.0, 0.47]}
      first_row:  {box: [0.67, 2.17, 12.0, 0.61]}
    row_pitch: 0.67
    col_template: {x: [0.8, 2.2, 5.07, 8.53, 11.0], w: [1.2, 2.67, 3.27, 2.33, 1.6]}
    component: table
    capacity: {rows: 6, cols: 6}
  visual_kv:
    use_for: 왼쪽 도식/그림 패널 + 오른쪽 속성(키-값) 목록
    native: srmc
    source: measured
    evidence: "R:S4–S8"
    regions:
      panel: {box: [0.67, 1.93, 5.87, 4.8], component: panel}
      kv:    {box: [6.93, 1.93, 5.73, 4.8], component: key_value, label_w: 1.67, value_x: 8.67, value_w: 4.0, row_h: 0.8}
    uses_slots: [status]
    capacity: {kv_rows: 6}
  image_pair_kv:
    use_for: 전/후 비교 이미지 2장 + 속성 목록
    native: srmc
    source: measured
    evidence: "R:S3"
    regions:
      header_1:  {box: [0.67, 1.64, 3.45, 0.37], style: panel_heading}
      image_1:   {box: [0.67, 2.07, 3.45, 4.6], component: image_frame.figure, fit: contain}
      caption_1: {box: [0.67, 6.72, 3.45, 0.29], style: caption}
      header_2:  {box: [4.38, 1.64, 3.45, 0.37], style: panel_heading}
      image_2:   {box: [4.38, 2.07, 3.45, 4.6], component: image_frame.figure, fit: contain}
      caption_2: {box: [4.38, 6.72, 3.45, 0.29], style: caption}
      kv:        {box: [9.13, 1.67, 3.53, 5.25], component: key_value, label_w: 0.93, value_x: 10.13, value_w: 2.53, row_h: 1.05}
    uses_slots: [status]
    capacity: {kv_rows: 5}
  bullets:
    use_for: 제목 + 불릿 목록 (입력 PPT에서 가장 흔한 형태)
    native: collision
    source: estimated
    basis: "샘플에 불릿 본문 없음. content_area + components.bullets"
    regions:
      body: {box: [0.42, 2.12, 12.5, 4.66], component: bullets}
    capacity: {lines: 9, l1_items: 7}
  bullets_image:
    use_for: 불릿 + 이미지 1장
    native: collision
    source: estimated
    basis: "screenshot_steps의 좌우를 바꿔 텍스트를 왼쪽에 둠"
    regions:
      body:  {box: [0.42, 2.12, 5.9, 4.66], component: bullets}
      image: {box: [6.62, 2.12, 6.3, 4.66], fit: contain, align: top-right}
    capacity: {lines: 9}
  io_steps:
    use_for: "입력 → 처리 → 출력 단계 행(3열 × 2–5행). 처리 칸 = 단계 번호 + 제목 + 글자 태그 + 보조 줄"
    native: collision
    source: estimated
    basis: "샘플에 없음. 입력 덱(spacemap-flow S3,S5,S6,S8,S9)의 반복 구조를 collision content_area에 맞춤. 셀은 step_cell(panel), 머리글은 table_header 실측 스타일"
    regions:
      lead:     {box: [0.42, 2.12, 12.5, 0.4], style: lead_compact, optional: true, max_lines: 1}
      footnote: {box: [0.42, 6.5, 12.5, 0.28], style: io_note, optional: true, max_lines: 1}
    header: {y: 2.56, h: 0.26, style: table_header}
    columns:
      in:   {x: 0.42, w: 3.2}
      step: {x: 3.92, w: 5.0}
      out:  {x: 9.22, w: 3.7}
    arrows: {component: arrow, between: [[in, step], [step, out]], gap: 0.3}
    rows: {top: 2.88, bottom: 6.42, gap: 0.08, max: 5, bottom_without_footnote: 6.78}
    cell:
      component: step_cell
      in:   {lines: io_line, secondary: io_note}
      step: {number: io_step_no, headline: io_headline, tags: tag, lines: io_line, note: io_note}
      out:  {lines: io_line, secondary: io_note}
    capacity: {rows: 5, step_lines: 3, in_lines: 2, out_lines: 3}
  card_grid:
    use_for: "같은 형식의 카드 4–6개(3×2 / 2×2) — 원칙·특징 목록. 위에 배너 1개, 아래 범례 1줄을 선택적으로"
    native: collision
    source: estimated
    basis: "intro_columns 3단 열(C:S4 실측 x/w)을 2행 그리드로 확장. 카드는 panel 컴포넌트"
    regions:
      banner: {box: [0.42, 2.12, 12.5, 0.72], component: panel, optional: true, line_dash_from_source: true, heading: {style: card_heading}, body: {style: card_body}}
      legend: {box: [0.42, 6.84, 12.5, 0.24], style: legend, optional: true}
    columns: {count: 3, x: [0.42, 4.71, 9.0], w: 3.91}
    rows: {with_banner: {y: [3.04, 4.95], h: 1.73}, without_banner: {y: [2.12, 4.5], h: 2.2}}
    card: {component: panel, insets: [0.18, 0.14, 0.18, 0.14], heading: {style: card_heading, h: 0.36}, body: {style: card_body, gap: 0.08}}
    capacity: {cards: 6, card_body_lines: 3}
    variants:
      two_by_two: {source: estimated, basis: "intro_columns.two_columns(x/w)와 같은 2단", columns: {count: 2, x: [0.42, 6.86], w: 6.06}, capacity: {cards: 4}}
      stats:
        source: estimated
        basis: "숫자 카드 1행(3–4개). 숫자는 step_number 실측 스타일"
        columns: {count: 4, x: [0.42, 3.64, 6.86, 10.08], w: 2.84}
        rows: {y: [2.12], h: 1.6}
        card: {number: {style: stat_number, h: 0.6}, label: {style: stat_label, h: 0.3}, detail: {style: io_note, max_lines: 2}}
        capacity: {cards: 4}
  diagram_restyle:
    use_for: "자유 배치 도식(상자 + 연결선) — 원본 도형의 배치·연결을 그대로 두고 색·글꼴만 이 디자인으로 다시 칠함"
    native: collision
    source: estimated
    basis: "샘플에 도식 없음(M3, 사용자 결정). 원본 도형 XML을 복제해 본문 영역에 균등 비율로 맞추고, 채움=panel, 선=deck.line_map/color_map, 글자=굵게→heading·보통→body. 그림자 등 효과 제거(6장 평면 원칙)"
    regions:
      lead:     {box: [0.42, 2.12, 12.5, 0.4], style: lead_compact, optional: true, max_lines: 1}
      diagram:  {box: [0.42, 2.62, 12.5, 4.16], fit: contain, align: top-center, bottom_with_footnote: 6.42}
      footnote: {box: [0.42, 6.5, 12.5, 0.28], style: io_note, optional: true, max_lines: 1}
    shape: {fill: panel, line_default: rule_subtle, keep: [geometry, dash, width, arrowheads], effects: none}
    text: {bold: heading, regular: body, size: "원본 × 배율", min_size: 12, color: "deck.color_map (없으면 text)"}
    capacity: {shapes: 40}
  references:
    use_for: 출처/참고문헌 목록
    native: srmc
    source: measured
    evidence: "R:S9"
    regions:
      body: {box: [0.67, 1.87, 12.0, 4.67], style: reference}
      note: {box: [0.67, 6.6, 12.0, 0.4], style: caption, italic: true}
    capacity: {items: 10}

overflow:
  measure: "실제 폰트 파일 기준 줄바꿈 계산 (build_deck.py), 이후 LibreOffice 렌더로 확인"
  order: [shrink, reflow, split, summarize]
  shrink:
    min_ratio: 0.85
    floors: {title: 22, body: 14, note: 13, bullet_l1: 16, bullet_l2: 14, table_cell: 10.5, kv_value: 12.5, step_headline: 15, io_line: 10.5, io_note: 9.5, card_body: 12}
    source: estimated
    basis: "floors는 샘플 내 같은 역할의 최소 실측 크기 근처 (note 15.5, R 셀 12, R 캡션 10.7)"
  reflow:
    title: "collision: 1줄 초과 시 22pt에서 2줄까지 허용 (1.33+0.61 < content top 2.12). srmc: 제목 상자 0.66″ 안에서 24pt까지 축소 후 2줄 불가 → split 없이 보고서에 기록 (추정)"
    screenshot_steps: "flow 열 안에서 note를 아래로 밀고, 그래도 넘치면 split"
  split:
    when: "shrink 하한에서도 넘칠 때"
    how: "의미 단위(불릿 묶음/표 행/키-값 행)로 나누고 제목은 그대로, kicker 끝에 ' · n/m' 추가"
    evidence: "C:S5–S9 kicker '< 01 · Getting started · 1/5 >' 실측 관례"
    table: "헤더 행을 각 분할 슬라이드에 반복"
  summarize:
    when: "분할하면 한 장에 1–2줄만 남는 등 분할이 부자연스러울 때만"
    rule: "원문을 바꾸지 않는 것이 기본. 요약 시 원문 전체를 보고서에 기록하고 노트(slide notes)에도 원문 보존"
  per_layout:
    cover:            {title_lines: 2, meta_lines: 3}
    toc:              {items: 6, desc_lines: 2}
    intro_columns:    {lead_lines: 2, column_body_lines: 4}
    screenshot_steps: {headline_lines: 3, column_height: 4.71}
    table:            {rows: 6}
    visual_kv:        {rows: 6}
    image_pair_kv:    {rows: 5}
    bullets:          {lines: 9}
    io_steps:         {rows: 5, step_lines: 3}
    card_grid:        {cards: 6, card_body_lines: 3}
    diagram_restyle:  {scale_min: 0.7, note: "배율이 0.7 미만이면 글자가 너무 작아 split(도식을 둘로) 또는 수작업"}
```

## 1. Visual Theme & Atmosphere

검은 우주 배경 위에 노랑(#FFC000)과 흰색만으로 구성한 **고대비 다크 테마**다. 배경은 마스터의 어두운 지구 이미지를
`#030509` 불투명 사각형으로 완전히 덮은 순수 검정이고(C:S1, S3–S21), 지구 이미지는 원본 그대로 보여주는 전체 이미지
슬라이드(C:S2)에서만 드러난다. 장식은 얇은 회색 구분선과 SPACEMAP 로고뿐이며, 시각적 무게는 큰 스크린샷과 노란 제목이
담당한다. 정보 밀도는 중간 — 한 슬라이드에 이미지 하나와 짧은 설명 블록.

**srmc 변형**은 같은 브랜드의 데이터 지향 버전이다. 배경은 어두운 지구 이미지를 덮지 않고 그대로 쓰고, 노란 대문자
키커 → 흰 Arial 제목 → 굵은 회색 밑줄 순서로 위를 짜며, 로고는 좌하단, 하단 중앙에 출처 줄이 온다. 본문까지 모두 Arial이고,
반투명 남색 패널·상태 배지·키-값 표로 정보를 압축한다.

## 2. Color Palette & Roles

| 이름 | hex | 역할 | 출처 |
|---|---|---|---|
| bg | `030509` | 슬라이드 배경 | 실측 C:S1,S3–S21 |
| text | `FFFFFF` | 본문, 소제목 | 실측 |
| text_muted | `C6C6C6` | 보조 설명, 푸터 | 실측 C |
| text_label | `B8C0CC` | 표 헤더, 키 라벨, 캡션, 출처 | 실측 R |
| accent | `FFC000` | 제목, 번호, 페이지 번호 | 실측 C·R |
| rule | `808080` | 헤더 구분선 0.75pt | 실측 C |
| rule_subtle | `4B5563` | 행 구분선, 이미지 테두리 | 실측 R |
| panel | `0B1426` @82% | 도식 패널 | 실측 R:S4–S8 |
| status amber / red / blue | `FFC000` on `4D3A0A` · `F87171` on `5A1E22` · `60A5FA` on `1E3A66` | 상태 배지 | 실측 R |
| chart_series | FFC000 → 60A5FA → F87171 → 7FA8F5 → C6C6C6 → FFB4B6 | 차트 계열 | **(추정)** |

테마 색(theme1–5)은 모두 Office 기본값이라 쓰지 않는다. 실제 색은 런·도형에 직접 지정되어 있다.

## 3. Typography Rules

- **두 가족만 쓴다**: 제목·숫자·데이터는 **Arial Bold**, 본문·키커·푸터는 **Roboto Light**. 키커와 URL만 이탤릭.
- 위계(C 기준): 제목 26 → 단계 번호 34 / 리드 23 / 열 제목 22 → 본문 17–20 → 노트 15.5 → 푸터 9.5.
  표지 제목 32, 감사 제목 60. R의 데이터 텍스트는 환산 후 12–14.7pt.
- 텍스트 상자는 **여백 0, 자동 맞춤 없음, 위쪽 정렬, 단락 간격 0**이 기본(C 콘텐츠 전체). 표지만 여백 0.05–0.1″,
  줄간격 110%, 도형 크기 자동 맞춤.
- 한글: **Pretendard**(제목 Bold, 본문 Light). 샘플 테마의 ea 폰트는 맑은 고딕이지만 사용자 결정으로 바꿨다.
  런마다 ea 폰트를 명시해야 테마 값(맑은 고딕)으로 돌아가지 않는다. Pretendard는 로컬에 설치되어 있어 한글 넘침 QA도 신뢰할 수 있다.
- 섹션 제목 44pt와 불릿 20/17pt는 **(추정)**.

## 4. Component Stylings

- **표**: 세로선 없음, 배경 없음. 헤더는 라벨색 Bold, 행 아래 `4B5563` 0.86pt 구분선, 첫 열 Bold, 셀 세로 가운데 정렬.
  상태가 있는 행만 상태색 82% 틴트(R:S2). 네이티브 표에도 같은 규칙 적용.
- **키-값 목록**: 라벨(라벨색 Bold) / 값(흰색, 핵심 행만 Bold), 행마다 아래 구분선(R:S3–S8).
- **상태 배지(pill)**: 둥근 사각형, 어두운 상태색 채움 + 밝은 상태색 Bold 글자, 가운데 정렬.
- **패널**: `0B1426` 82% 둥근 사각형, 같은 색 0.86pt 테두리 — 도식을 담는 그릇.
- **차트 (추정)**: 배경 없음, 계열색은 chart_series 순서, 축 라벨 Roboto Light 12 회색, 격자 `4B5563`, 데이터 라벨 Arial Bold 12 흰색.
- **불릿 (추정)**: `•` 노란색, L1 Roboto Light 20 흰색, L2 17 회색, 들여쓰기 0.3″/0.6″.
- **이미지 프레임**: 스크린샷은 테두리 없이 좌상단 정렬(C). 비교용 그림은 `4B5563` 0.65pt 테두리 + 위 헤더 + 아래 캡션(R:S3).
- **강조 표시**: 스크린샷 위 빨간 3pt 둥근 사각형(C:S2), 그림 위 빨간 점선 타원 + 흰 바탕 라벨(R:S3).
- **상자 스타일**: 샘플 실측 상자는 패널(`0B1426` 82%), **상태 상자**(어두운 상태색 채움 + 같은 색 선 + 밝은 상태색 Bold 글자,
  R:S2–S8), 그림 테두리(`4B5563` 0.65pt, 채움 없음), 흰 바탕 라벨(R:S3). 입력 덱의 상자 종류는 plan.json `deck.box_map`으로
  이 중 하나에 대응시켜 도식·단계 셀·카드·표지 칩에 똑같이 쓴다. 상태 채움은 **예외 단계에만**(샘플에서 상태색은 배지 크기 강조, 패널
  대비 면적 약 1/10), 흔한 종류는 패널 + 색 테두리 **(추정)**. 상태 상자 안 본문은 흰색 — 상태색 글자는 짧은 배지 문구 전용.
- **글자 태그 (추정)**: D/N/A/H 같은 한 글자 표시는 채움 없이 상태색 Bold 글자로. 태그→색 대응은 덱마다 정한다(plan.json).
- **화살표 (추정)**: `808080` 1pt 선 + 삼각 화살촉.

## 5. Layout Principles

**두 변형의 슬라이드 틀**

| | collision (기본) | srmc |
|---|---|---|
| 배경 | `#030509` 불투명 사각형 | 어두운 지구 이미지 그대로 |
| 상단 | 로고 우상단 → 0.92″ 구분선 → 이탤릭 키커 `< … >` → 노란 제목 26pt | 노란 대문자 키커 13.3pt → 흰 제목 28pt → 굵은 밑줄 |
| 상태 배지 | 제목 줄 오른쪽 끝 **(추정)** | 헤더 우상단 (실측) |
| 하단 | 좌우 푸터 + 노란 2자리 페이지 번호 | 좌하단 로고 + 출처 줄, 페이지 번호 없음 |
| 본문 영역 | [0.42, 2.12, 12.5 × 4.66] | [0.67, 1.63, 12.0 × 5.1] |
| 본문 글꼴 | Roboto Light / 보조 `C6C6C6` | Arial / 보조 `B8C0CC` **(치환 규칙은 추정)** |

레이아웃 좌표는 **원래 나온 샘플의 변형(native) 기준 실측값**으로 저장한다. 다른 변형으로 변환할 때는 두 본문 영역 사이의
선형 변환(`variant_mapping`)으로 옮기며, 그렇게 옮긴 레이아웃은 항상 **추정**이다. 표지·마지막 장·전체 이미지는 틀이 없어
두 변형에서 같다.

| 레이아웃 | 용도 | native | collision | srmc |
|---|---|---|---|---|
| cover | 표지 | – | 실측 C:S1 | 실측 R:S1 |
| closing | 마지막 장 | – | 실측 C:S21 | 실측(공유) |
| image_full | 원본 이미지 전체 | – | 실측 C:S2 | 실측(공유) |
| toc | 목차 2×3 | collision | 실측 C:S3 | 추정 |
| intro_columns | 리드 + 3단 (2단은 추정) | collision | 실측 C:S4 | 추정 |
| screenshot_steps | 큰 이미지 + 오른쪽 설명 / 번호 목록 | collision | 실측 C:S5–S20 | 추정 |
| table | 비교 표 | srmc | 추정 | 실측 R:S2 |
| visual_kv | 도식 패널 + 키-값 | srmc | 추정 | 실측 R:S4–S8 |
| image_pair_kv | 전/후 이미지 + 키-값 | srmc | 추정 | 실측 R:S3 |
| references | 출처 목록 | srmc | 추정 | 실측 R:S9 |
| section | 섹션 구분 | collision | 추정 | 추정 |
| bullets / bullets_image | 불릿 본문 (+이미지) | collision | 추정 | 추정 |
| io_steps | 입력 → 처리 → 출력 단계 행 (3열 × 최대 5행) | collision | 추정 | 추정 |
| card_grid | 카드 3×2 / 2×2 (+배너, 범례), 숫자 카드 1×4(`stats`) | collision | 추정 | 추정 |
| diagram_restyle | 자유 도식을 배치 그대로 다시 칠함 | collision | 추정 | 추정 |

표지의 단계 칩 행(`flow`)과 범례(`legend`)도 **(추정)** 선택 영역이다(M2, 입력 덱에 표지 칩이 있어 추가).

## 6. Depth & Elevation

완전히 평면이다. 슬라이드에 그림자·광선·반사 효과가 없다(그림자는 쓰이지 않는 레이아웃에만 있음). 깊이는
① 82% 불투명 남색 패널과 ② 상태색 행 틴트로만 표현한다. 테두리는 0.65–0.86pt 회색 계열만 쓴다.

## 7. Do's and Don'ts

**Do**
- 제목은 항상 노랑 Arial Bold 한 줄, 그 위에 이탤릭 키커로 위치 표시(`< 01 · Getting started · 1/5 >`).
- 이미지는 크게(콘텐츠 폭의 2/3), 설명은 오른쪽 좁은 열에 위에서부터 쌓는다.
- 숫자·번호·핵심 값은 노랑 Bold로 띄운다.
- 분할된 연속 슬라이드는 키커 끝에 `n/m`을 단다.

**Don't**
- 흰 배경, 밝은 패널, 그라데이션, 그림자를 넣지 않는다.
- 제외된 MINISTRY 배너나 다른 기관 로고를 좌상단에 넣지 않는다. 그 자리는 비워 둔다.
- 노랑을 본문 단락 전체에 쓰지 않는다(제목·숫자·짧은 강조만). 상태 상자 안 본문도 흰색이다.
- 상태색 채움을 흔한 요소에 쓰지 않는다. 예외·주의 단계만 채우고 나머지는 패널 + 얇은 색 테두리.
- 표에 세로선이나 셀 배경 격자를 쓰지 않는다.
- Office 테마 색(파랑 4472C4 등)을 그대로 쓰지 않는다.

## 8. Responsive Behavior (Overflow)

넘침 처리 순서: **축소 → 재배치 → 분할 → 요약**(토큰 `overflow`).
1. **축소**: 원래 크기의 85% 또는 역할별 하한(본문 14, 노트 13, 표 셀 10.5…)까지만 줄인다. **(하한값은 추정)**
2. **재배치**: 제목은 22pt에서 2줄까지 허용. screenshot_steps는 오른쪽 열 안에서 노트를 아래로 민다.
3. **분할**: 불릿 묶음·표 행·키-값 행 단위로 나누고, 제목은 유지하고 키커에 `· n/m`을 붙인다(C:S5–S9 관례). 표는 헤더를 반복한다.
4. **요약**: 분할이 부자연스러울 때만. 원문 전체를 변환 보고서와 슬라이드 노트에 남긴다.

레이아웃별 한도: 표 6행, 키-값 6행(image_pair_kv 5행), 목차 6항목, 불릿 9줄, 표지 제목 2줄, 단계 행 5행, 카드 6개.

## 9. Agent Prompt Guide

변환 시 슬라이드마다 다음 순서로 판단한다.
1. 첫 장 → `cover`, 마지막 감사/Q&A 장 → `closing`, 목차 → `toc`, 장 제목만 있는 슬라이드 → `section`.
2. 큰 이미지 1장 + 짧은 설명 → `screenshot_steps`(단계 번호가 있으면 `step_no` 사용, 번호 목록이면 `numbered_list` 변형).
3. 이미지 1장 + 불릿 여러 개 → `bullets_image`. 불릿만 → `bullets`.
4. 표 → `table`(6행 초과 시 분할). 2열 표(항목/값)는 `visual_kv`의 kv 단독 또는 `table`.
   입력/처리/출력처럼 열 머리글 아래 같은 모양 상자가 행마다 화살표로 이어진 격자 → `io_steps`(5행 초과 시 분할).
5. 병렬 항목 3개(또는 2개) → `intro_columns`. 같은 형식 카드 4–6개 → `card_grid`, 큰 숫자 + 이름 카드 → `card_grid.stats`.
6. 출처/참고문헌 → `references`.
7. 자유 배치 도식(상자·연결선만으로 된 다이어그램) → `diagram_restyle`로 원본 배치를 유지한 채 다시 칠한다(사용자 결정 M3).
   그림·SmartArt·차트가 섞여 다시 칠할 수 없는 도식만 원본 렌더 이미지를 `image_full`로 넣고 보고서에 **수작업 필요**로 표시.
8. 원문 텍스트는 바꾸지 않는다. 키커는 입력에 섹션 정보가 있을 때만 채우고, 없으면 덱 제목을 쓴다.
   제목이 "① 섹션 — 소제목" 꼴이면 키커 `< ① 섹션 >` + 제목 "소제목"으로 나눈다(글자는 그대로, 자리만 옮김. 사용자 결정 2026-10-13).
   제목 끝에 다른 크기·색의 상태 문구가 붙어 있으면 `status` 배지로 옮긴다.
   도식 슬라이드에 표지 칩·범례·숫자 카드처럼 넣을 자리가 있는 요소는 노트로 숨기지 말고 해당 영역(`cover.flow`, `legend`,
   `card_grid.stats`)으로 옮긴다. 자유 도식 부분만 `image_full`로 남기고, 한 슬라이드가 둘로 나뉘면 키커에 `· n/m`.
9. 변형은 사용자가 지정한다(`--variant`, 기본 collision). 지정이 없을 때 임의로 바꾸지 않는다. 변환 보고서에 사용한
   변형과, 그 변형에서 **추정**으로 그려진 레이아웃 목록을 적는다.
10. srmc 변형에서는 출처 줄(`source_line`)을 채울 수 있으면 채우고(입력의 출처/각주 텍스트), 없으면 비운다. 페이지 번호는 넣지 않는다.
