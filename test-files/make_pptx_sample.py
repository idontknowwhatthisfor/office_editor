"""Writes test-files/pptx-sample-raw.pptx with python-pptx (pip install python-pptx; MIT).

    python3 test-files/make_pptx_sample.py
    osascript test-files/pptx_resave.applescript "$PWD/test-files/pptx-sample-raw.pptx" pptx-sample

The second step re-saves the deck in PowerPoint (so the file is as PowerPoint writes it) and
exports PowerPoint's own picture of every slide, the reference for the viewer.
"""
import struct
import zlib
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

here = Path(__file__).parent


def png(width: int, height: int) -> bytes:
    """A gradient test picture, written without any imaging library."""
    rows = b""
    for y in range(height):
        rows += b"\0" + b"".join(
            bytes((int(255 * x / width), int(255 * y / height), 160)) for x in range(width))
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b""))


prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
L = prs.slide_layouts

# 1. Title slide
s = prs.slides.add_slide(L[0])
s.shapes.title.text = "Quarterly Review"
s.placeholders[1].text = "Pocket Office test deck · November 2026"
s.notes_slide.notes_text_frame.text = "Speaker notes for the title slide."

# 2. Bullets with levels and run formatting
s = prs.slides.add_slide(L[1])
s.shapes.title.text = "Highlights"
tf = s.placeholders[1].text_frame
tf.text = "Revenue grew 12% year on year"
for text, level in [("Strongest in the North region", 1), ("New customers: 1,240", 1),
                    ("Costs held flat", 0), ("Hiring paused until Q2", 1), ("Contractors down 8%", 2),
                    ("Next steps", 0)]:
    p = tf.add_paragraph()
    p.text = text
    p.level = level
p = tf.add_paragraph()
r = p.add_run(); r.text = "Mixed runs: "
r = p.add_run(); r.text = "bold"; r.font.bold = True
r = p.add_run(); r.text = ", "
r = p.add_run(); r.text = "italic red"; r.font.italic = True; r.font.color.rgb = RGBColor(0xC0, 0, 0)
r = p.add_run(); r.text = ", "
r = p.add_run(); r.text = "large"; r.font.size = Pt(32)

# 3. Shapes
s = prs.slides.add_slide(L[5])
s.shapes.title.text = "Shapes"
def shape(kind, x, y, w, h, text="", fill=None, line=None, rot=0):
    sh = s.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if text:
        sh.text_frame.text = text
    if fill:
        sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor.from_string(fill)
    if line:
        sh.line.color.rgb = RGBColor.from_string(line); sh.line.width = Pt(3)
    sh.rotation = rot
    return sh
shape(MSO_SHAPE.RECTANGLE, 0.6, 1.8, 2.4, 1.4, "Rectangle")
shape(MSO_SHAPE.ROUNDED_RECTANGLE, 3.4, 1.8, 2.4, 1.4, "Rounded", fill="70AD47", line="375623")
shape(MSO_SHAPE.OVAL, 6.2, 1.8, 2.4, 1.4, "Ellipse", fill="FFC000")
shape(MSO_SHAPE.ISOSCELES_TRIANGLE, 9.0, 1.8, 1.8, 1.4)
shape(MSO_SHAPE.RIGHT_ARROW, 0.6, 3.8, 2.6, 1.2, "Arrow", fill="ED7D31")
shape(MSO_SHAPE.DIAMOND, 3.6, 3.7, 1.6, 1.6)
shape(MSO_SHAPE.PENTAGON, 5.6, 3.8, 2.4, 1.2, "Chevron-ish")
shape(MSO_SHAPE.RECTANGLE, 8.6, 3.8, 2.4, 1.2, "Rotated 20°", fill="5B9BD5", rot=20)
g = shape(MSO_SHAPE.RECTANGLE, 0.6, 5.6, 4.0, 1.2, "Gradient")
g.fill.gradient(); g.fill.gradient_angle = 0
g.fill.gradient_stops[0].color.rgb = RGBColor(0x44, 0x72, 0xC4)
g.fill.gradient_stops[1].color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(5.2), Inches(6.2), Inches(9.5), Inches(5.6))
c.line.width = Pt(2.5); c.line.color.rgb = RGBColor(0x7F, 0x7F, 0x7F)
shape(MSO_SHAPE.STAR_5_POINT, 10.4, 5.4, 1.4, 1.4, fill="FFD966")

# 4. Picture
s = prs.slides.add_slide(L[5])
s.shapes.title.text = "A picture"
img = here / "pptx-sample-picture.png"
img.write_bytes(png(240, 160))
s.shapes.add_picture(str(img), Inches(1), Inches(1.8), Inches(6), Inches(4))
tb = s.shapes.add_textbox(Inches(7.5), Inches(2), Inches(5), Inches(3))
tb.text_frame.word_wrap = True
tb.text_frame.text = ("Caption text box. This paragraph is long enough to wrap onto several lines "
                      "inside its box, as a caption beside a picture would.")

# 5. Table
s = prs.slides.add_slide(L[5])
s.shapes.title.text = "Table"
rows, cols = 5, 4
t = s.shapes.add_table(rows, cols, Inches(1), Inches(1.8), Inches(11), Inches(3.5)).table
data = [["Region", "Q1", "Q2", "Change"], ["North", "120", "134", "+12%"], ["South", "98", "101", "+3%"],
        ["East", "143", "139", "−3%"], ["West", "77", "90", "+17%"]]
for i, row in enumerate(data):
    for j, text in enumerate(row):
        t.cell(i, j).text = text
        if j:
            t.cell(i, j).text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT

# 6. Group, anchors, alignment
s = prs.slides.add_slide(L[6])
grp = s.shapes.add_group_shape()
a = grp.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(2))
a.text_frame.text = "Grouped A"
b = grp.shapes.add_shape(MSO_SHAPE.OVAL, Inches(3), Inches(2.2), Inches(3), Inches(2))
b.text_frame.text = "Grouped B"
for i, (anchor, align, name) in enumerate([(MSO_ANCHOR.TOP, PP_ALIGN.LEFT, "Top left"),
                                           (MSO_ANCHOR.MIDDLE, PP_ALIGN.CENTER, "Middle centre"),
                                           (MSO_ANCHOR.BOTTOM, PP_ALIGN.RIGHT, "Bottom right")]):
    box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7 + i * 2), Inches(1), Inches(1.8), Inches(3))
    box.fill.solid(); box.fill.fore_color.rgb = RGBColor(0xF2, 0xF2, 0xF2)
    box.line.color.rgb = RGBColor(0x40, 0x40, 0x40)
    box.text_frame.text = name
    box.text_frame.vertical_anchor = anchor
    box.text_frame.paragraphs[0].alignment = align
    box.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(0, 0, 0)
tb = s.shapes.add_textbox(Inches(1), Inches(5), Inches(11), Inches(1.5))
tb.text_frame.word_wrap = True
p = tb.text_frame.paragraphs[0]
p.text = "Text box on a blank slide, 24 pt Georgia, centred."
p.alignment = PP_ALIGN.CENTER
p.runs[0].font.size = Pt(24); p.runs[0].font.name = "Georgia"

# 7. Chart (not drawn by the viewer yet; it should say so)
s = prs.slides.add_slide(L[5])
s.shapes.title.text = "Chart"
cd = CategoryChartData()
cd.categories = ["Q1", "Q2", "Q3", "Q4"]
cd.add_series("Sales", (19.2, 21.4, 16.7, 23.0))
s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1.5), Inches(1.8), Inches(10), Inches(5), cd)

# 8. Two content + coloured background
s = prs.slides.add_slide(L[3])
s.background.fill.solid(); s.background.fill.fore_color.rgb = RGBColor(0x1F, 0x38, 0x64)
s.shapes.title.text = "Two columns on a dark background"
s.shapes.title.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
for ph, items in ((s.placeholders[1], ["Left one", "Left two"]), (s.placeholders[2], ["Right one", "Right two"])):
    ph.text_frame.text = items[0]
    ph.text_frame.add_paragraph().text = items[1]
    for p in ph.text_frame.paragraphs:
        p.runs[0].font.color.rgb = RGBColor(0xDD, 0xEB, 0xF7)

prs.save(here / "pptx-sample-raw.pptx")
print("wrote", here / "pptx-sample-raw.pptx")
