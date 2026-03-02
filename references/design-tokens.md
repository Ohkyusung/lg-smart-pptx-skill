# LG Group Design Tokens Reference

## Color Palette

### Primary Colors

| Token | Name | HEX | RGB | 용도 |
|-------|------|-----|-----|------|
| `LG_RED` | LG Red | `#A50034` | `(165, 0, 52)` | 브랜드 컬러. L-브래킷, 액센트 바, TOC 번호, 카테고리 라벨, 강조 |
| `BLACK` | Black | `#000000` | `(0, 0, 0)` | 제목, 본문 텍스트 |
| `WHITE` | White | `#FFFFFF` | `(255, 255, 255)` | 슬라이드 배경 |

### Secondary Colors (텍스트)

| Token | Name | HEX | RGB | 용도 |
|-------|------|-----|-----|------|
| `DARK_GRAY` | Dark Gray | `#333333` | `(51, 51, 51)` | 부제목, 보조 텍스트 |
| `MEDIUM_GRAY` | Medium Gray | `#666666` | `(102, 102, 102)` | 섹션명, 캡션, 비활성 텍스트 |

### Surface Colors (배경)

| Token | Name | HEX | RGB | 용도 |
|-------|------|-----|-----|------|
| `LIGHT_GRAY` | Light Gray | `#F2F2F2` | `(242, 242, 242)` | 콘텐츠 박스 배경, 테이블 교대행 |
| `CHARCOAL` | Charcoal | `#3C3C3C` | `(60, 60, 60)` | 타임라인 헤더, 테이블 헤더 |
| `BORDER_GRAY` | Border Gray | `#CCCCCC` | `(204, 204, 204)` | 구분선, 테두리 |

### Accent Colors (강조)

| Token | Name | HEX | RGB | 용도 |
|-------|------|-----|-----|------|
| `GREEN` | Accent Green | `#2E7D32` | `(46, 125, 50)` | 미래 계획 항목, 신규 기능 |
| `ORANGE` | Accent Orange | `#D4760A` | `(212, 118, 10)` | 하이라이트 항목, 주의 사항 |

### Affiliate Tag Colors (계열사 뱃지)

| 계열사 | HEX | 용도 |
|--------|-----|------|
| LGES (에너지솔루션) | `#1565C0` | 파란색 텍스트 태그 |
| LGD (디스플레이) | `#D4760A` | 주황색 텍스트 태그 |
| LGC (화학) | `#2E7D32` | 초록색 텍스트 태그 |

## Typography

### Font Family

| 우선순위 | 폰트명 | 비고 |
|----------|--------|------|
| 1 | LG Smart (LG스마트체) | 기본 폰트 (한글+영문) |
| 2 | 맑은 고딕 (Malgun Gothic) | 폴백 폰트 |
| 3 | Arial | 최종 폴백 (영문) |

### Font Weight Structure

LG Smart는 표준 패밀리 구조를 사용합니다:

| Weight | 패밀리명 | bold 속성 | 파일명 |
|--------|----------|-----------|--------|
| Regular | `LG Smart` | `False` | LG Smart Regular.ttf |
| SemiBold | `LG Smart Light` | `True` | LG Smart SemiBold.ttf |
| Bold | `LG Smart` | `True` | LG Smart Bold.ttf |

> 에이투지체와의 차이: 에이투지체는 weight별 별도 패밀리명("에이투지체 4 Regular", "에이투지체 7 Bold")이라 `font.bold=False` 고정이었지만,
> LG Smart는 동일 패밀리 "LG Smart"에서 `font.bold` 속성으로 Bold를 구분합니다.
> SemiBold는 "LG Smart Light" 패밀리의 Bold 스타일로 등록되어 있습니다.

### Type Scale

| 용도 | Size | Weight | Line Height | 색상 |
|------|------|--------|-------------|------|
| 표지 제목 | 32pt | Bold | 1.3 | Black |
| 표지 부제 | 14pt | Regular | 1.4 | Dark Gray |
| TOC 타이틀 | 28pt | Regular | 1.3 | Black |
| TOC 항목 | 16pt | Bold | 1.3 | LG Red |
| TOC 하위항목 | 13pt | Regular | 1.3 | Dark Gray |
| 섹션 제목 | 28pt | Bold | 1.2 | Black |
| 슬라이드 부제 | 16pt | SemiBold | 1.3 | Dark Gray |
| 본문 제목 | 18pt | Bold | 1.3 | Black |
| 본문 | 12pt | Regular | 1.5 | Black |
| 본문 (소) | 11pt | Regular | 1.5 | Black/Dark Gray |
| 표 헤더 | 10pt | Bold | 1.2 | White (on Charcoal) |
| 표 본문 | 10pt | Regular | 1.3 | Black |
| 캡션/주석 | 9pt | Regular | 1.3 | Medium Gray |
| 라벨 뱃지 | 9pt | Bold | 1.2 | White (on LG Red) |

### Korean Font Setup (python-pptx)

python-pptx에서 한글 폰트를 올바르게 적용하려면 Latin과 East Asian 폰트를 모두 설정해야 합니다:

```python
# LG Smart 폰트 설정 예시
# font.bold = True/False 로 Bold를 제어 (표준 패밀리 구조)
# SemiBold는 "LG Smart Light" 패밀리 + bold=True 로 접근

from pptx.oxml.ns import qn

def set_korean_font(run, font_name, bold=False):
    font = run.font
    font.bold = bold  # LG Smart는 bold 속성으로 weight 제어

    rPr = run._r.get_or_add_rPr()
    # East Asian font
    for ea in rPr.findall(qn('a:ea')):
        rPr.remove(ea)
    ea = OxmlElement('a:ea')
    ea.set('typeface', font_name)
    ea.set('charset', '-127')
    rPr.append(ea)
    # Complex Script font
    for cs in rPr.findall(qn('a:cs')):
        rPr.remove(cs)
    cs = OxmlElement('a:cs')
    cs.set('typeface', font_name)
    cs.set('charset', '-127')
    rPr.append(cs)
```

## Layout Dimensions

### Slide Size
- **비율**: 16:9 와이드스크린
- **크기**: 13.333" x 7.5" (33.867cm x 19.05cm)

### Margins & Spacing

| 영역 | 값 | 비고 |
|------|------|------|
| 좌측 마진 (표지) | 0.8cm | L-브래킷 시작점 |
| 좌측 마진 (내용) | 1.5cm | 액센트 바 뒤 |
| 우측 마진 | 1.0cm | |
| 상단 마진 | 1.2cm | |
| 하단 마진 | 1.0cm | |
| 콘텐츠 시작점 | 2.0cm (좌) | 액센트 바 + 여백 |
| 제목-본문 간격 | 1.5cm | |
| 불릿 항목 간격 | 6pt (space_after) | |
| 표 셀 내부 패딩 | 0.2cm (좌우), 0.1cm (상하) | |

### Component Dimensions

| 컴포넌트 | 폭 | 높이 | 비고 |
|----------|------|------|------|
| 액센트 바 (좌) | 0.3cm | 슬라이드 전체 | x=0 |
| L-브래킷 팔 길이 | 2.5cm | - | |
| L-브래킷 두께 | 0.4cm | - | |
| 섹션 인디케이터 점 | 0.35cm | 0.35cm | 원형 |
| 타임라인 헤더 높이 | 1.2cm | - | |
| 라벨 뱃지 | 4.0cm | 0.8cm | 가변 폭 |
| 로드맵 좌측 라벨 | 2.5cm | 가변 | |
| 테이블 행 높이 | ~1.0cm | - | |

## Slide Type Specifications

### Cover (표지)
```
┌──────────────────────────────┐
│ ┌──┐                        │
│ │  │                        │
│ │                            │
│                              │
│      [제목 - 32pt Bold]      │
│                              │
│                              │
│      [부제 - 14pt]           │
│      [날짜 - 14pt Bold]   ┌─┤
│                        LG │ │
│                            └─┤
└──────────────────────────────┘
```

### TOC (목차)
```
┌──────────────────────────────┐
│  ─────────────────────────── │ ← 상단 회색선
│                              │
│  Contents                    │
│  ───                         │ ← 빨간 밑줄 바
│  ─────────────────────────── │ ← 회색 구분선
│                              │
│    I.  Summary               │ ← 빨간색
│    II. 시스템 소개           │ ← 빨간색
│        - 항목 1              │ ← 검정 들여쓰기
│        - 항목 2              │
│    III. 첨부자료             │
│                              │
└──────────────────────────────┘
```

### Content (내용)
```
┌──────────────────────────────┐
│▌ 1.1 제목               섹션 ●│ ← 좌: 레드바, 우: 섹션+점
│▌                              │
│▌ [본문 내용 영역]             │
│▌  • 불릿 항목 1               │
│▌  • 불릿 항목 2               │
│▌  • 불릿 항목 3               │
│▌                              │
│▌                              │
│▌                              │
└──────────────────────────────┘
```

### Roadmap (로드맵)
```
┌──────────────────────────────┐
│▌ [SPC] 로드맵           섹션 ●│
│▌ 설명 텍스트                  │
│▌                              │
│▌     ◄ 2025 ►◄ 2026 ►◄ 2027 ►│ ← 차콜 쉐브론
│▌ ┌──┐┌────┐┌────┐┌────┐      │
│▌ │ 라││    ││    ││    │      │
│▌ │ 벨││ 셀 ││ 셀 ││ 셀 │     │
│▌ └──┘└────┘└────┘└────┘      │
│▌                              │
│▌ ■ 계열사별 현황              │
│▌ ┌────┬────┬────┬────┐       │
│▌ │헤더│헤더│헤더│헤더│        │ ← 차콜 헤더
│▌ ├────┼────┼────┼────┤       │
│▌ │데이│터  │    │    │       │
│▌ └────┴────┴────┴────┘       │
└──────────────────────────────┘
```
