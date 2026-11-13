#!/usr/bin/env python3
"""Builds the blank files behind "New document" and "New spreadsheet" on the home screen.

Writes app/src/main/res/raw/blank_document.docx and blank_spreadsheet.xlsx. They are written by
hand rather than saved from Word/Excel so they hold no author name, template path or other
metadata, only what Word and Excel need: A4 paper with 2.54 cm margins, Calibri 11 pt, Word's
usual paragraph spacing; one empty sheet.

Run from anywhere: python3 app/templates/make_blank_templates.py
"""
import zipfile
from pathlib import Path

RAW = Path(__file__).resolve().parent.parent / 'src' / 'main' / 'res' / 'raw'

XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
OFFICE_DOC = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
STYLES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
SHEET = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
SML = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

DOCX = {
    '[Content_Types].xml': XML + (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '</Types>'
    ),
    '_rels/.rels': XML + (
        f'<Relationships xmlns="{PKG_REL}">'
        f'<Relationship Id="rId1" Type="{OFFICE_DOC}" Target="word/document.xml"/>'
        '</Relationships>'
    ),
    'word/_rels/document.xml.rels': XML + (
        f'<Relationships xmlns="{PKG_REL}">'
        f'<Relationship Id="rId1" Type="{STYLES}" Target="styles.xml"/>'
        '</Relationships>'
    ),
    'word/document.xml': XML + (
        f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>'
        '<w:p/>'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>'
        '<w:cols w:space="708"/></w:sectPr>'
        '</w:body></w:document>'
    ),
    'word/styles.xml': XML + (
        f'<w:styles xmlns:w="{W}">'
        '<w:docDefaults>'
        '<w:rPrDefault><w:rPr>'
        '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>'
        '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>'
        '</w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
        '</w:docDefaults>'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
        '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">'
        '<w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
        '</w:styles>'
    ),
}

XLSX = {
    '[Content_Types].xml': XML + (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>'
    ),
    '_rels/.rels': XML + (
        f'<Relationships xmlns="{PKG_REL}">'
        f'<Relationship Id="rId1" Type="{OFFICE_DOC}" Target="xl/workbook.xml"/>'
        '</Relationships>'
    ),
    'xl/_rels/workbook.xml.rels': XML + (
        f'<Relationships xmlns="{PKG_REL}">'
        f'<Relationship Id="rId1" Type="{SHEET}" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId2" Type="{STYLES}" Target="styles.xml"/>'
        '</Relationships>'
    ),
    'xl/workbook.xml': XML + (
        f'<workbook xmlns="{SML}" xmlns:r="{R}">'
        '<bookViews><workbookView/></bookViews>'
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
        '<calcPr calcId="191029"/>'
        '</workbook>'
    ),
    'xl/worksheets/sheet1.xml': XML + (
        f'<worksheet xmlns="{SML}" xmlns:r="{R}">'
        '<dimension ref="A1"/>'
        '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        '<sheetData/>'
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        '</worksheet>'
    ),
    'xl/styles.xml': XML + (
        f'<styleSheet xmlns="{SML}">'
        '<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>'
    ),
}


def write(path: Path, parts: dict[str, str]) -> None:
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, text in parts.items():
            # A fixed date keeps the output identical between runs.
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, text)
    print(f'wrote {path} ({path.stat().st_size} bytes)')


RAW.mkdir(parents=True, exist_ok=True)
write(RAW / 'blank_document.docx', DOCX)
write(RAW / 'blank_spreadsheet.xlsx', XLSX)
