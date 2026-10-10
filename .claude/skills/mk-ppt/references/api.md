# python-pptx 1.0.2 기능별 코드 레퍼런스

필요한 절만 골라 읽는다.

## 목차
1. 파일
2. 슬라이드
3. 도형
4. 텍스트
5. 표
6. 차트
7. 서식(DML)
8. 읽기·분석
9. 저수준 확장(XML)

공통 import:

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.chart.data import CategoryChartData, XyChartData, BubbleChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
```

---

## 1. 파일

```python
prs = Presentation()                      # 새 파일 (기본 4:3)
prs.slide_width = Inches(13.333)          # 16:9로 변경
prs.slide_height = Inches(7.5)

prs = Presentation("template.pptx")       # 기존 파일 또는 템플릿에서 시작
prs.save("out.pptx")                      # 저장

import io
buf = io.BytesIO(); prs.save(buf)         # 메모리에 저장

cp = prs.core_properties                  # 문서 속성
cp.title, cp.author, cp.subject, cp.keywords = "제목", "작성자", "주제", "키워드"
```

## 2. 슬라이드

기본 템플릿의 레이아웃 인덱스: 0 제목, 1 제목+내용, 2 구역 머리글, 3 콘텐츠 2개, 4 비교, 5 제목만, 6 빈 화면, 7 캡션 있는 콘텐츠, 8 캡션 있는 그림.

```python
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = "제목"
slide.placeholders[1].text = "본문"

layout = prs.slide_layouts.get_by_name("Title Only")   # 이름으로 찾기
for i, s in enumerate(prs.slides, 1): ...              # 순회
s = prs.slides.get(slide_id)                           # ID로 찾기

fill = slide.background.fill                           # 배경
fill.solid(); fill.fore_color.rgb = RGBColor(0xF5, 0xF7, 0xFA)

slide.notes_slide.notes_text_frame.text = "발표자 노트"
```

삭제·복제·이동은 공식 API가 없다 → `scripts/slide_ops.py`.

## 3. 도형

```python
sh = slide.shapes
box = sh.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(1))
box.adjustments[0] = 0.2                  # 둥근 모서리 정도
box.rotation = 15
box.shadow.inherit = False                # 그림자 끄기

tb = sh.add_textbox(Inches(1), Inches(2), Inches(5), Inches(1))

pic = sh.add_picture("img.png", Inches(1), Inches(3), width=Inches(4))  # height 생략 시 비율 유지
pic.crop_left = 0.1                       # 비율(0~1)로 자르기

conn = sh.add_connector(MSO_CONNECTOR.STRAIGHT, 0, 0, 0, 0)
conn.begin_connect(box, 3); conn.end_connect(other, 1)   # 연결점 인덱스

grp = sh.add_group_shape()
grp.shapes.add_shape(MSO_SHAPE.OVAL, 0, 0, Inches(1), Inches(1))

ff = sh.build_freeform(Inches(1), Inches(1), scale=Inches(1))
ff.add_line_segments([(2, 0), (1, 1.5)], close=True)
ff.convert_to_shape()

sh.add_movie("a.mp4", Inches(1), Inches(1), Inches(4), Inches(3), poster_frame_image="p.png")
sh.add_ole_object("data.xlsx", "Excel.Sheet.12", Inches(1), Inches(1))

box.click_action.hyperlink.address = "https://example.com"
box.click_action.target_slide = prs.slides[2]            # 슬라이드 이동 링크

# 플레이스홀더 채우기
for ph in slide.placeholders:
    print(ph.placeholder_format.idx, ph.placeholder_format.type, ph.name)
slide.placeholders[1].insert_picture("img.png")           # 그림 플레이스홀더

# 종류 판별
shape.shape_type == MSO_SHAPE_TYPE.PICTURE
shape.has_text_frame, shape.has_table, shape.has_chart
```

## 4. 텍스트

```python
tf = tb.text_frame
tf.word_wrap = True
tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
tf.margin_left = tf.margin_right = Inches(0.1)

p = tf.paragraphs[0]
p.text = "첫 문단"
p.alignment = PP_ALIGN.CENTER
p.line_spacing = 1.2
p.space_after = Pt(6)

p2 = tf.add_paragraph(); p2.level = 1     # 들여쓰기 수준
r = p2.add_run(); r.text = "강조"
f = r.font
f.name = "맑은 고딕"; f.size = Pt(20); f.bold = True; f.italic = False; f.underline = True
f.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
r.hyperlink.address = "https://example.com"

tf.fit_text(font_file="C:/Windows/Fonts/malgun.ttf", max_size=24)   # 상자에 맞게 축소. 글꼴 파일 경로가 필요하다
# 더 간단한 대안: tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE (PowerPoint가 열 때 축소)
```

한글 글꼴은 run마다 `font.name`을 지정해야 확실하다. 도형 전체 텍스트를 바꾼 뒤에는 모든 paragraph의 run을 돌며 서식을 다시 적용한다.

## 5. 표

```python
gf = sh.add_table(rows=4, cols=3, left=Inches(1), top=Inches(1.5), width=Inches(8), height=Inches(2))
tbl = gf.table
tbl.columns[0].width = Inches(3)
tbl.rows[0].height = Inches(0.5)
tbl.first_row = True; tbl.horz_banding = True

c = tbl.cell(0, 0)
c.text = "항목"
c.fill.solid(); c.fill.fore_color.rgb = RGBColor(0x1F, 0x4E, 0x79)
c.vertical_anchor = MSO_ANCHOR.MIDDLE
c.margin_left = Inches(0.05)

tbl.cell(1, 0).merge(tbl.cell(2, 0))      # 병합
tbl.cell(1, 0).split()                    # 분할
```

## 6. 차트

```python
cd = CategoryChartData()
cd.categories = ["2024", "2025", "2026"]
cd.add_series("매출", (10, 14, 21))
gf = sh.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1.5), Inches(8), Inches(5), cd)
chart = gf.chart

chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.BOTTOM
chart.legend.include_in_layout = False
chart.has_title = True; chart.chart_title.text_frame.text = "연도별 매출"
chart.font.name = "맑은 고딕"; chart.font.size = Pt(12)

va = chart.value_axis
va.minimum_scale, va.maximum_scale = 0, 25
va.has_major_gridlines = True
va.tick_labels.number_format = '0"억"'; va.tick_labels.number_format_is_linked = False

plot = chart.plots[0]
plot.has_data_labels = True
plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END

s = plot.series[0]
s.format.fill.solid(); s.format.fill.fore_color.rgb = RGBColor(0x2E, 0x75, 0xB6)
s.points[2].format.fill.solid()            # 특정 막대만 강조

chart.replace_data(new_chart_data)         # 데이터 교체
```

주요 차트 종류: `COLUMN_CLUSTERED`, `BAR_CLUSTERED`, `COLUMN_STACKED`, `LINE_MARKERS`, `PIE`, `DOUGHNUT`, `AREA`, `XY_SCATTER`(XyChartData), `BUBBLE`(BubbleChartData), `RADAR`.

## 7. 서식(DML)

```python
fill = shape.fill
fill.solid(); fill.fore_color.rgb = RGBColor(0xFF, 0xC0, 0x00)
fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
fill.fore_color.brightness = 0.4           # 밝게(-1~1)

fill.gradient(); fill.gradient_angle = 90
st = fill.gradient_stops
st[0].color.rgb = RGBColor(0xFF, 0xFF, 0xFF); st[1].color.rgb = RGBColor(0x1F, 0x4E, 0x79)

from pptx.enum.dml import MSO_PATTERN
fill.patterned(); fill.pattern = MSO_PATTERN.LIGHT_DOWNWARD_DIAGONAL

fill.background()                          # 투명

from pptx.enum.dml import MSO_LINE_DASH_STYLE
ln = shape.line
ln.color.rgb = RGBColor(0, 0, 0); ln.width = Pt(1.5); ln.dash_style = MSO_LINE_DASH_STYLE.DASH
ln.fill.background()                       # 선 없음
```

## 8. 읽기·분석

```python
for i, slide in enumerate(prs.slides, 1):
    for shape in slide.shapes:
        if shape.has_text_frame:
            print(i, shape.name, shape.text_frame.text)
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            open(f"img_{i}.{shape.image.ext}", "wb").write(shape.image.blob)
        if shape.has_table:
            rows = [[c.text for c in r.cells] for r in shape.table.rows]
        if shape.has_chart:
            ch = shape.chart
            cats = list(ch.plots[0].categories)
            vals = [(s.name, list(s.values)) for s in ch.plots[0].series]
```

그룹 도형 안쪽은 `shape.shape_type == MSO_SHAPE_TYPE.GROUP`이면 `shape.shapes`를 재귀 순회한다. 빠른 전체 확인은 `scripts/inspect_pptx.py`.

**찾아 바꾸기** — 서식을 유지하려면 run 단위로 바꾼다. 단어가 여러 run에 걸쳐 있으면 찾지 못하므로, 그런 경우 paragraph 텍스트를 합쳐 비교한 뒤 첫 run에 결과를 넣고 나머지 run을 비운다.

```python
for shape in slide.shapes:
    if shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                r.text = r.text.replace("{{날짜}}", "2026-10-10")
```

## 9. 저수준 확장(XML)

```python
el = shape._element                        # lxml 요소
from lxml import etree
print(etree.tostring(el, pretty_print=True).decode())

from pptx.oxml.ns import qn
sldIdLst = prs.slides._sldIdLst            # 슬라이드 순서 목록 (slide_ops.py가 사용)
```

단위: `Inches(1) == Cm(2.54) == Pt(72) == Emu(914400)`.
