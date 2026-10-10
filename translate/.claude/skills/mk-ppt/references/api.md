# python-pptx 1.0.2 Code Reference by Feature

Read only the sections you need.

## Table of contents
1. File
2. Slides
3. Shapes
4. Text
5. Tables
6. Charts
7. Formatting (DML)
8. Reading/analysis
9. Low-level extension (XML)

Common imports:

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

## 1. File

```python
prs = Presentation()                      # new file (4:3 by default)
prs.slide_width = Inches(13.333)          # change to 16:9
prs.slide_height = Inches(7.5)

prs = Presentation("template.pptx")       # start from an existing file or template
prs.save("out.pptx")                      # save

import io
buf = io.BytesIO(); prs.save(buf)         # save to memory

cp = prs.core_properties                  # document properties
cp.title, cp.author, cp.subject, cp.keywords = "제목", "작성자", "주제", "키워드"
```

## 2. Slides

Layout indexes in the default template: 0 Title, 1 Title and Content, 2 Section Header, 3 Two Content, 4 Comparison, 5 Title Only, 6 Blank, 7 Content with Caption, 8 Picture with Caption.

```python
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = "제목"
slide.placeholders[1].text = "본문"

layout = prs.slide_layouts.get_by_name("Title Only")   # find by name
for i, s in enumerate(prs.slides, 1): ...              # iterate
s = prs.slides.get(slide_id)                           # find by ID

fill = slide.background.fill                           # background
fill.solid(); fill.fore_color.rgb = RGBColor(0xF5, 0xF7, 0xFA)

slide.notes_slide.notes_text_frame.text = "발표자 노트"
```

Delete, duplicate, and move have no official API → `scripts/slide_ops.py`.

## 3. Shapes

```python
sh = slide.shapes
box = sh.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(1))
box.adjustments[0] = 0.2                  # corner rounding amount
box.rotation = 15
box.shadow.inherit = False                # turn off shadow

tb = sh.add_textbox(Inches(1), Inches(2), Inches(5), Inches(1))

pic = sh.add_picture("img.png", Inches(1), Inches(3), width=Inches(4))  # omitting height keeps aspect ratio
pic.crop_left = 0.1                       # crop by ratio (0–1)

conn = sh.add_connector(MSO_CONNECTOR.STRAIGHT, 0, 0, 0, 0)
conn.begin_connect(box, 3); conn.end_connect(other, 1)   # connection point indexes

grp = sh.add_group_shape()
grp.shapes.add_shape(MSO_SHAPE.OVAL, 0, 0, Inches(1), Inches(1))

ff = sh.build_freeform(Inches(1), Inches(1), scale=Inches(1))
ff.add_line_segments([(2, 0), (1, 1.5)], close=True)
ff.convert_to_shape()

sh.add_movie("a.mp4", Inches(1), Inches(1), Inches(4), Inches(3), poster_frame_image="p.png")
sh.add_ole_object("data.xlsx", "Excel.Sheet.12", Inches(1), Inches(1))

box.click_action.hyperlink.address = "https://example.com"
box.click_action.target_slide = prs.slides[2]            # link that jumps to a slide

# fill placeholders
for ph in slide.placeholders:
    print(ph.placeholder_format.idx, ph.placeholder_format.type, ph.name)
slide.placeholders[1].insert_picture("img.png")           # picture placeholder

# identify shape type
shape.shape_type == MSO_SHAPE_TYPE.PICTURE
shape.has_text_frame, shape.has_table, shape.has_chart
```

## 4. Text

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

p2 = tf.add_paragraph(); p2.level = 1     # indent level
r = p2.add_run(); r.text = "강조"
f = r.font
f.name = "맑은 고딕"; f.size = Pt(20); f.bold = True; f.italic = False; f.underline = True
f.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
r.hyperlink.address = "https://example.com"

tf.fit_text(font_file="C:/Windows/Fonts/malgun.ttf", max_size=24)   # shrink to fit the box; requires a font file path
# simpler alternative: tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE (PowerPoint shrinks it on open)
```

To be sure Korean fonts apply, set `font.name` on every run. After replacing a shape's whole text, loop over every paragraph's runs and reapply the formatting.

## 5. Tables

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

tbl.cell(1, 0).merge(tbl.cell(2, 0))      # merge
tbl.cell(1, 0).split()                    # split
```

## 6. Charts

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
s.points[2].format.fill.solid()            # highlight a single bar

chart.replace_data(new_chart_data)         # replace data
```

Main chart types: `COLUMN_CLUSTERED`, `BAR_CLUSTERED`, `COLUMN_STACKED`, `LINE_MARKERS`, `PIE`, `DOUGHNUT`, `AREA`, `XY_SCATTER` (XyChartData), `BUBBLE` (BubbleChartData), `RADAR`.

## 7. Formatting (DML)

```python
fill = shape.fill
fill.solid(); fill.fore_color.rgb = RGBColor(0xFF, 0xC0, 0x00)
fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
fill.fore_color.brightness = 0.4           # lighten (-1 to 1)

fill.gradient(); fill.gradient_angle = 90
st = fill.gradient_stops
st[0].color.rgb = RGBColor(0xFF, 0xFF, 0xFF); st[1].color.rgb = RGBColor(0x1F, 0x4E, 0x79)

from pptx.enum.dml import MSO_PATTERN
fill.patterned(); fill.pattern = MSO_PATTERN.LIGHT_DOWNWARD_DIAGONAL

fill.background()                          # transparent

from pptx.enum.dml import MSO_LINE_DASH_STYLE
ln = shape.line
ln.color.rgb = RGBColor(0, 0, 0); ln.width = Pt(1.5); ln.dash_style = MSO_LINE_DASH_STYLE.DASH
ln.fill.background()                       # no line
```

## 8. Reading/analysis

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

For the inside of group shapes, if `shape.shape_type == MSO_SHAPE_TYPE.GROUP`, traverse `shape.shapes` recursively. For a quick full check, use `scripts/inspect_pptx.py`.

**Find and replace** — to preserve formatting, replace at the run level. A word that spans multiple runs won't be found; in that case, compare against the joined paragraph text, put the result in the first run, and empty the remaining runs.

```python
for shape in slide.shapes:
    if shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                r.text = r.text.replace("{{날짜}}", "2026-10-10")
```

## 9. Low-level extension (XML)

```python
el = shape._element                        # lxml element
from lxml import etree
print(etree.tostring(el, pretty_print=True).decode())

from pptx.oxml.ns import qn
sldIdLst = prs.slides._sldIdLst            # slide order list (used by slide_ops.py)
```

Units: `Inches(1) == Cm(2.54) == Pt(72) == Emu(914400)`.
