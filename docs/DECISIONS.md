# DECISIONS — slide-restyle

주요 결정과 그 이유. 새 결정은 아래에 추가하고, 뒤집힌 결정은 지우지 않고 `상태: 대체됨 → #N`으로 남긴다.
디자인 값 자체(색·폰트·좌표)는 [DESIGN.md](../DESIGN.md)가 기준이다. 진행 계획은 [PLAN.md](PLAN.md).

---

## 1. 공식 pptx 스킬은 베이스로 쓰지 않고 참고만 한다 — 2026-10-13
- 이유:
  - 생성 방식이 pptxgenjs(Node)와 XML 직접 편집이라 "Python + python-pptx" 제약과 맞지 않는다.
  - "대담한 팔레트", "같은 레이아웃 반복 금지" 같은 새 디자인용 지침이 샘플 재현이라는 목적과 충돌한다.
  - 라이선스가 Proprietary라 공개 레포에 코드를 복사할 수 없다.
- 가져온 것: python-pptx 주의점(런 단위 입력, 슬라이드 복제 없음, SVG/EMF 불가), 템플릿 상속 방식,
  렌더 절차와 시각 검수 체크리스트, 선택적 `validate.py` 검사.

## 2. 스크립트는 PEP 723 + `uv run`으로 실행한다 — 2026-10-13
- 이유: 전역 pip 설치 없이 의존성(python-pptx, PyYAML, Pillow)을 스크립트별로 선언할 수 있다.

## 3. design 단계 산출물은 DESIGN.md + base 템플릿 두 가지 — 2026-10-13
- 결정: `design/base.pptx`·`base.potx`(샘플에서 슬라이드만 지우고 마스터·레이아웃·테마 유지)를 함께 만든다.
  build_deck은 이 위에 슬라이드를 추가한다.
- 이유: 배경 이미지, 로고, 마스터 장식은 토큰만으로 재현할 수 없다.

## 4. 단계 사이 중간 산출물은 JSON 파일로 남긴다 — 2026-10-13
- 결정: design.json → content.json → plan.json. 스크립트(결정적)와 Claude(판단)의 경계를 파일 단위로 나눈다.

## 5. DESIGN.md 토큰은 yaml 펜스 블록 하나, 모든 값에 출처를 단다 — 2026-10-13
- 결정: `yaml design-tokens` 블록만 스크립트가 파싱한다. 각 값은 `measured`(+evidence `C:S5`) 또는
  `estimated`(+basis). 길이 단위는 inch, 기준 캔버스는 C의 13.335″×7.5″, R은 ×1.3335로 환산.
- 이유: 사용자 요구 — 샘플에 없어 추정한 레이아웃은 실측값과 구분해 명시.
- 메모: 초기 계획은 좌표를 EMU 정수로 저장하는 것이었으나, DESIGN.md 가독성을 위해 inch로 바꿨다
  (design.json은 EMU 유지).

## 6. 샘플은 PPTX를 쓴다, PDF 지원은 보류 — 2026-10-13
- 결정: 샘플이 PDF에서 PPTX 2개로 교체되어 PDF 분기(pdfplumber)는 구현하지 않았다.
- 메모: 필요해지면 pdfplumber(MIT)를 쓴다. PyMuPDF는 AGPL이라 공개 레포에 맞지 않는다.

## 7. MINISTRY 배너는 디자인에서 제외 — 2026-10-13 (사용자 요청)
- 결정: 샘플 원본은 수정하지 않고, `design/exclude.yaml`에 dHash로 등록해 추출 단계에서 뺀다.
  오른쪽 위 SPACEMAP 로고는 남긴다. 배너가 있던 좌상단은 비워 둔다.
- 이유: 배너는 마스터가 아니라 각 슬라이드에 직접 붙어 있고, 두 샘플의 해상도가 달라 내용 해시가 필요하다.

## 8. 표지·마지막 장의 초록 행사 블록도 제외 — 2026-10-13 (사용자 결정)
- 결정: "Professional Training Programme / SSA | SDA | STM / 날짜" 블록과 색 `00B050`을 토큰에서 뺀다.
- 이유: 배너와 같은 행사 전용 문구.

## 9. 한글(ea) 폰트는 Pretendard — 2026-10-13 (사용자 결정)
- 결정: 제목 Pretendard Bold, 본문 Pretendard Light(Roboto Light와 굵기 맞춤). 런마다 `<a:ea>`를 명시한다.
- 이유: 샘플 테마의 맑은 고딕은 macOS에 설치할 수 없어 한글 넘침 검수가 근사치가 된다.
- 영향: 받는 PC에 Pretendard가 없으면 대체 폰트로 보인다. M3에서 베이스 테마 ea도 Pretendard로 바꾼다.

## 10. 슬라이드 틀은 collision 단일 방식 — 2026-10-13
- 상태: 대체됨 → #11

## 11. 디자인 변형 2개: collision(기본) / srmc — 2026-10-13 (사용자 결정)
- 결정: 색·폰트·컴포넌트·표지는 공유하고 슬라이드 틀(배경·로고·키커/제목·푸터)만 다르다. `--variant srmc`로 고른다.
  레이아웃 좌표는 원래 나온 샘플의 변형(native) 기준 실측값으로 저장하고, 다른 변형에서는 본문 영역 간
  선형 변환(`variant_mapping`)으로 옮긴다. 옮긴 레이아웃은 항상 estimated.
- 이유: 사용자가 SRMC 방식 디자인도 쓸 일이 있다고 판단.

## 12. 렌더는 반드시 render.sh로, 정적 Roboto는 `./.fonts/`에 — 2026-10-13
- 결정: `soffice`를 직접 호출하지 않는다. render.sh가 macOS 폰트 폴더와 `./.fonts/`를 포함한 fontconfig를 만들어
  주입한다. 맑은 고딕은 Apple SD Gothic Neo로 alias(샘플 한글 QA만 근사치).
- 이유: LibreOffice headless의 번들 fontconfig는 macOS 폰트를 못 보고, Homebrew Roboto는 가변 폰트라
  "Roboto Light"로 매칭되지 않는다. 시스템 폰트 폴더는 건드리지 않는다.
