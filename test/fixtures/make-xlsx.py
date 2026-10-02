"""Minimal real .xlsx generator (deep-scan regression fixture, OCR 🔴-1).

Runs through the officemcp RunPython channel: reads {"path": ...} from `data`,
writes a smallest-valid xlsx (inline strings, one sheet named 数据) with the
header row [指标, 数值, 备注] and one data row. Excel opens this file for real.
"""
import json
import zipfile

args = json.loads(data)
p = args["path"]

ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
      '</Types>')
rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>')
wb = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
      '<sheets><sheet name="数据" sheetId="1" r:id="rId1"/></sheets></workbook>')
wbrels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
          '</Relationships>')
sheet = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
         '<sheetData>'
         '<row r="1">'
         '<c r="A1" t="inlineStr"><is><t>指标</t></is></c>'
         '<c r="B1" t="inlineStr"><is><t>数值</t></is></c>'
         '<c r="C1" t="inlineStr"><is><t>备注</t></is></c>'
         '</row>'
         '<row r="2">'
         '<c r="A2" t="inlineStr"><is><t>深扫标记</t></is></c>'
         '<c r="B2"><v>42</v></c>'
         '</row>'
         '</sheetData></worksheet>')

with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", ct)
    z.writestr("_rels/.rels", rels)
    z.writestr("xl/workbook.xml", wb)
    z.writestr("xl/_rels/workbook.xml.rels", wbrels)
    z.writestr("xl/worksheets/sheet1.xml", sheet)

print(json.dumps({"ok": True, "path": p, "bytes": len(open(p, "rb").read())}))
