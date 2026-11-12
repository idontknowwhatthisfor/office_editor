"""Generates small sample files for device testing (standard library only).

    python3 test-files/make_samples.py

Real-world files from Word/Excel/Acrobat belong in this folder too; these only cover the basics.
"""
import struct
import zipfile
import zlib
from pathlib import Path

here = Path(__file__).parent

LOREM = (
    "The quick brown fox jumps over the lazy dog. Pocket editors should open fast, save without "
    "breaking the file, work offline and cost nothing. This paragraph repeats so that every page "
    "has enough text to select, copy and search."
).split()


def make_pdf(path: Path, pages: int) -> None:
    """A text PDF using the non-embedded Helvetica font; page 3 is landscape."""
    objects: list[bytes] = []

    def add(body: bytes) -> int:
        objects.append(body)
        return len(objects)

    catalog = add(b"")  # filled in below
    pages_obj = add(b"")
    font = add(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
    bold = add(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>")
    kids = []
    for n in range(1, pages + 1):
        w, h = (842, 595) if n == 3 else (595, 842)
        lines = [f"BT /F2 22 Tf 56 {h - 80} Td (Page {n} of {pages}) Tj ET"]
        if n % 50 == 0:
            lines.append(f"BT /F2 14 Tf 56 {h - 110} Td (Milestone marker for page {n}) Tj ET")
        y = h - 150
        words = LOREM * 6
        per_line = 11 if w < 800 else 16
        for i in range(0, len(words), per_line):
            if y < 60:
                break
            text = " ".join(words[i:i + per_line])
            lines.append(f"BT /F1 12 Tf 56 {y} Td ({text}) Tj ET")
            y -= 18
        stream = "\n".join(lines).encode("latin-1")
        content = add(b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"\nendstream")
        kids.append(add(
            b"<< /Type /Page /Parent %d 0 R /MediaBox [0 0 %d %d] /Contents %d 0 R "
            b"/Resources << /Font << /F1 %d 0 R /F2 %d 0 R >> >> >>"
            % (pages_obj, w, h, content, font, bold)
        ))
    objects[catalog - 1] = b"<< /Type /Catalog /Pages %d 0 R >>" % pages_obj
    objects[pages_obj - 1] = b"<< /Type /Pages /Count %d /Kids [%s] >>" % (
        pages, b" ".join(b"%d 0 R" % k for k in kids))

    out = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = []
    for i, body in enumerate(objects, start=1):
        offsets.append(len(out))
        out += b"%d 0 obj\n" % i + body + b"\nendobj\n"
    xref = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objects) + 1)
    for off in offsets:
        out += b"%010d 00000 n \n" % off
    out += b"trailer\n<< /Size %d /Root %d 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (
        len(objects) + 1, catalog, xref)
    path.write_bytes(out)


def png(width: int, height: int) -> bytes:
    """A small gradient PNG (red to blue), built by hand."""
    rows = b"".join(
        b"\x00" + b"".join(bytes((255 - 255 * x // width, 80, 255 * x // width)) for x in range(width))
        for _ in range(height)
    )

    def chunk(kind: bytes, data: bytes) -> bytes:
        body = kind + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body))

    return (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(rows))
            + chunk(b"IEND", b""))


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def run(text: str, props: str = "") -> str:
    rpr = f"<w:rPr>{props}</w:rPr>" if props else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'


def para(content: str, props: str = "") -> str:
    ppr = f"<w:pPr>{props}</w:pPr>" if props else ""
    return f"<w:p>{ppr}{content}</w:p>"


def make_docx(path: Path) -> None:
    """A Word document with the features the viewer must show (PRD 4.2 P0): headings, bold /
    italic / underline, bullet and numbered lists (nested), a table, an image and a link."""
    def item(text: str, num_id: int, level: int = 0) -> str:
        return para(run(text), f'<w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="{num_id}"/></w:numPr>')

    def cell(text: str, bold: bool = False) -> str:
        return f'<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>{para(run(text, "<w:b/>" if bold else ""))}</w:tc>'

    def row(*cells: str, header: bool = False) -> str:
        return "<w:tr>" + "".join(cell(c, header) for c in cells) + "</w:tr>"

    image = (
        '<w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
        '<wp:extent cx="2743200" cy="914400"/><wp:docPr id="1" name="Gradient" descr="A red to blue gradient"/>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:nvPicPr><pic:cNvPr id="1" name="gradient.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="914400"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'
    )
    body = "".join([
        para(run("Group project notes"), '<w:pStyle w:val="Heading1"/>'),
        para(run("This paragraph has ") + run("bold", "<w:b/>") + run(", ") + run("italic", "<w:i/>")
             + run(", ") + run("underlined", '<w:u w:val="single"/>') + run(" and ")
             + run("bold italic", "<w:b/><w:i/>") + run(" text, plus a ")
             + '<w:hyperlink r:id="rId4">' + run("link", '<w:rStyle w:val="Hyperlink"/>') + "</w:hyperlink>"
             + run(". Characters that need escaping: &lt;tag&gt; &amp; “quotes” — café, naïve, 日本語.")),
        para(run("Shopping list"), '<w:pStyle w:val="Heading2"/>'),
        item("Apples", 1), item("Bananas", 1), item("Ripe ones only", 1, 1), item("Cherries", 1),
        para(run("Steps"), '<w:pStyle w:val="Heading2"/>'),
        item("Open the file", 2), item("Fix the typo", 2), item("Share it back", 2),
        para(run("Budget"), '<w:pStyle w:val="Heading2"/>'),
        '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
        + row("Item", "Qty", "Price", header=True) + row("Notebook", "3", "45.000")
        + row("Pens", "10", "20.000") + "</w:tbl>",
        para(run("Picture"), '<w:pStyle w:val="Heading2"/>'),
        para(f"<w:r>{image}</w:r>"),
        para(run("Centred closing line."), '<w:jc w:val="center"/>'),
    ] + [para(run(" ".join(LOREM))) for _ in range(12)])

    def numbering(abstract_id: int, fmt: str, texts: tuple[str, str]) -> str:
        levels = "".join(
            f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/><w:lvlText w:val="{t}"/>'
            f'<w:pPr><w:ind w:left="{720 * (i + 1)}" w:hanging="360"/></w:pPr></w:lvl>'
            for i, t in enumerate(texts))
        return f'<w:abstractNum w:abstractNumId="{abstract_id}">{levels}</w:abstractNum>'

    def style(style_id: str, name: str, kind: str = "paragraph", props: str = "") -> str:
        return f'<w:style w:type="{kind}" w:styleId="{style_id}"><w:name w:val="{name}"/>{props}</w:style>'

    borders = "".join(f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                      for side in ("top", "left", "bottom", "right", "insideH", "insideV"))

    xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    parts = {
        "[Content_Types].xml":
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Default Extension="png" ContentType="image/png"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
            '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
            '</Types>',
        "_rels/.rels":
            f'<Relationships xmlns="{REL}"><Relationship Id="rId1" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>',
        "word/_rels/document.xml.rels":
            f'<Relationships xmlns="{REL}">'
            f'<Relationship Id="rId1" Type="{R}/styles" Target="styles.xml"/>'
            f'<Relationship Id="rId2" Type="{R}/numbering" Target="numbering.xml"/>'
            f'<Relationship Id="rId3" Type="{R}/image" Target="media/gradient.png"/>'
            f'<Relationship Id="rId4" Type="{R}/hyperlink" Target="https://example.com/" TargetMode="External"/>'
            '</Relationships>',
        "word/styles.xml":
            f'<w:styles xmlns:w="{W}">'
            '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
            + style("Heading1", "heading 1", props='<w:pPr><w:spacing w:before="240"/></w:pPr><w:rPr><w:b/><w:color w:val="2F5496"/><w:sz w:val="36"/></w:rPr>')
            + style("Heading2", "heading 2", props='<w:pPr><w:spacing w:before="160"/></w:pPr><w:rPr><w:b/><w:color w:val="2F5496"/><w:sz w:val="28"/></w:rPr>')
            + style("Hyperlink", "Hyperlink", "character", '<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>')
            + style("TableGrid", "Table Grid", "table", f"<w:tblPr><w:tblBorders>{borders}</w:tblBorders></w:tblPr>")
            + "</w:styles>",
        "word/numbering.xml":
            f'<w:numbering xmlns:w="{W}">' + numbering(0, "bullet", ("•", "o")) + numbering(1, "decimal", ("%1.", "%2."))
            + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
            "</w:numbering>",
        "word/document.xml":
            f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>{body}'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>',
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, content in parts.items():
            z.writestr(name, (xml + content).encode("utf-8"))
        z.writestr("word/media/gradient.png", png(300, 100))


make_pdf(here / "sample-200-pages.pdf", 200)
make_pdf(here / "sample-5-pages.pdf", 5)
make_docx(here / "sample.docx")
(here / "not-a-pdf.pdf").write_text("this is not a PDF\n")
print("wrote", ", ".join(sorted(p.name for p in here.iterdir() if p.suffix != ".py")))
