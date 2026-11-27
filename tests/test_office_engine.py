# MetaForge - Pruebas unitarias de office_engine
# Copyright (C) 2026 Kevin R. Clas (ClasK7) - GPLv3

import os
import tempfile
import unittest
from unittest import mock
import xml.etree.ElementTree as ET
import zipfile

from metaforge import FileProcessingError, ValidationError
from metaforge.office_engine import (
    CORE_NS,
    FIXED_ZIP_DATE_TIME,
    OfficeMetadataEngine,
    normalize_w3cdtf,
    parse_edit_time_minutes,
    validate_revision_number,
)


CORE_XML_SAMPLE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Borrador</dc:title>
  <dc:creator>Ana</dc:creator>
  <cp:lastModifiedBy>Luis</cp:lastModifiedBy>
  <cp:revision>4</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">2025-01-10T10:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-02-15T12:00:00Z</dcterms:modified>
</cp:coreProperties>
"""

APP_XML_SAMPLE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office Word</Application>
  <Company>Acme Corp</Company>
  <TotalTime>35</TotalTime>
</Properties>
"""

CUSTOM_XML_SAMPLE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="ClienteSecreto">
    <vt:lpwstr>ProyectoX</vt:lpwstr>
  </property>
</Properties>
"""

DOC_XML_WITH_MC_IGNORABLE = (
    "\ufeff"
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'mc:Ignorable="w14">'
    '<w:body><w:p><w:r>'
    '<w:t xml:space="preserve">Hola\u200bmundo&#x2060;&#08203;\ufeff'
    ' con\u202fespacio y emoji\u200d y guion\u00ad y zwnj\u200c</w:t>'
    '</w:r></w:p></w:body></w:document>'
)

COMMENTS_XML_SAMPLE = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:comment w:id="0" w:author="Ana"><w:p><w:r><w:t>Nota\u200boculta</w:t></w:r></w:p></w:comment>'
    '</w:comments>'
)


def _build_zip(path: str, files: dict[str, bytes | str]) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, content in files.items():
            raw = content.encode("utf-8") if isinstance(content, str) else content
            zf.writestr(name, raw)


class TestOfficeValidators(unittest.TestCase):
    def test_normalize_w3cdtf_python310_compatible(self):
        self.assertEqual(normalize_w3cdtf(""), "")
        self.assertEqual(normalize_w3cdtf("2026:11:27 14:30:00Z"), "2026-11-27T14:30:00Z")
        self.assertEqual(normalize_w3cdtf("2026:11:27 14:30:00-05:00"), "2026-11-27T14:30:00-05:00")
        self.assertEqual(normalize_w3cdtf("2026-11-27T14:30:00.250Z"), "2026-11-27T14:30:00.250Z")
        with self.assertRaises(ValidationError):
            normalize_w3cdtf("2024-02-30T10:00:00Z")
        with self.assertRaises(ValidationError):
            normalize_w3cdtf("no-es-fecha")

    def test_parse_edit_time_minutes(self):
        self.assertEqual(parse_edit_time_minutes(""), "0")
        self.assertEqual(parse_edit_time_minutes("45"), "45")
        self.assertEqual(parse_edit_time_minutes("0:05:00"), "5")
        self.assertEqual(parse_edit_time_minutes("1 hour 20 minutes"), "80")
        self.assertEqual(parse_edit_time_minutes("1 day 2 hours"), "1560")
        with self.assertRaises(ValidationError):
            parse_edit_time_minutes("mucho tiempo")

    def test_validate_revision_number(self):
        self.assertEqual(validate_revision_number("3"), "3")
        for bad in ("", "0", "-1", "abc"):
            with self.assertRaises(ValidationError):
                validate_revision_number(bad)


class TestOfficeEngine(unittest.TestCase):
    def test_edit_and_wipe_preserves_ooxml_structure_and_clears_custom_xml(self):
        with tempfile.TemporaryDirectory() as tmp:
            docx_path = os.path.join(tmp, "informe.docx")
            _build_zip(
                docx_path,
                {
                    "[Content_Types].xml": "<Types/>",
                    "_rels/.rels": "<Relationships/>",
                    "docProps/core.xml": CORE_XML_SAMPLE,
                    "docProps/app.xml": APP_XML_SAMPLE,
                    "docProps/custom.xml": CUSTOM_XML_SAMPLE,
                    "word/document.xml": DOC_XML_WITH_MC_IGNORABLE,
                    "word/comments.xml": COMMENTS_XML_SAMPLE,
                },
            )

            # Edición manual de fecha con formato ExifTool y duración en horas
            OfficeMetadataEngine.update_tags_batch(
                docx_path,
                {
                    "XML:CreateDate": "2026:03:15 09:00:00Z",
                    "XML:TotalEditTime": "1 hour 20 minutes",
                },
            )
            with zipfile.ZipFile(docx_path, "r") as zf:
                core_root = ET.fromstring(zf.read("docProps/core.xml"))
                created = core_root.find(f"{{{CORE_NS['dcterms']}}}created")
                self.assertEqual(created.text, "2026-03-15T09:00:00Z")
                self.assertEqual(created.attrib.get(f"{{{CORE_NS['xsi']}}}type"), "dcterms:W3CDTF")
                self.assertIn(b"80", zf.read("docProps/app.xml"))

            # Escaneo Unicode detecta U+200B (directo + entidad), U+2060 (entidad) y U+FEFF interno (sin BOM 0)
            hits = OfficeMetadataEngine.scan_invisible_unicode(docx_path)
            self.assertEqual(hits.get("Zero-Width Space (U+200B)"), 2)
            self.assertEqual(hits.get("Word Joiner (U+2060)"), 1)
            self.assertEqual(hits.get("Byte Order Mark interno (U+FEFF)"), 1)

            # Limpieza total
            OfficeMetadataEngine.wipe_all(docx_path, strip_invisible_unicode=True)

            with zipfile.ZipFile(docx_path, "r") as zf:
                # ZipInfo con fecha fija
                for info in zf.infolist():
                    self.assertEqual(info.date_time, FIXED_ZIP_DATE_TIME)

                # Fechas eliminadas en core.xml
                core_root = ET.fromstring(zf.read("docProps/core.xml"))
                self.assertIsNone(core_root.find(f"{{{CORE_NS['dcterms']}}}created"))
                self.assertIsNone(core_root.find(f"{{{CORE_NS['dcterms']}}}modified"))
                self.assertEqual(core_root.find(f"{{{CORE_NS['cp']}}}revision").text, "1")

                # custom.xml vaciado a <Properties/> válido
                custom_root = ET.fromstring(zf.read("docProps/custom.xml"))
                self.assertEqual(len(list(custom_root)), 0)

                # word/document.xml preserva BOM inicial, prefijos w:/mc:Ignorable y caracteres legítimos
                doc_text = zf.read("word/document.xml").decode("utf-8")
                self.assertTrue(doc_text.startswith("\ufeff"))
                self.assertIn('mc:Ignorable="w14"', doc_text)
                self.assertNotIn("ns0:", doc_text)
                self.assertIn("Holamundo con\u202fespacio y emoji\u200d y guion\u00ad y zwnj\u200c", doc_text)

                # word/comments.xml NO se modifica (coherente con el aviso al usuario)
                self.assertEqual(zf.read("word/comments.xml").decode("utf-8"), COMMENTS_XML_SAMPLE)

            self.assertEqual(OfficeMetadataEngine.scan_invisible_unicode(docx_path), {})

    def test_wipe_all_succeeds_without_app_xml_and_manual_edit_errors(self):
        with tempfile.TemporaryDirectory() as tmp:
            minimal = os.path.join(tmp, "minimal.docx")
            _build_zip(minimal, {"docProps/core.xml": CORE_XML_SAMPLE})

            with self.assertRaises(FileProcessingError):
                OfficeMetadataEngine.update_tag(minimal, "Company", "X")

            # wipe_all debe tolerar ausencia de app.xml y custom.xml
            OfficeMetadataEngine.wipe_all(minimal)

    def test_alias_conflict_creator_vs_author(self):
        with tempfile.TemporaryDirectory() as tmp:
            docx_path = os.path.join(tmp, "alias.docx")
            _build_zip(docx_path, {"docProps/core.xml": CORE_XML_SAMPLE})
            with self.assertRaises(ValidationError):
                OfficeMetadataEngine.update_tags_batch(docx_path, {"Creator": "Ana", "Author": "Beto"})

    def test_xlsx_and_pptx_unicode_scan_and_clean(self):
        with tempfile.TemporaryDirectory() as tmp:
            xlsx = os.path.join(tmp, "hoja.xlsx")
            _build_zip(
                xlsx,
                {
                    "docProps/core.xml": CORE_XML_SAMPLE,
                    "xl/sharedStrings.xml": "<sst><si><t>Celda\u200bUno</t></si></sst>",
                },
            )
            self.assertEqual(OfficeMetadataEngine.scan_invisible_unicode(xlsx), {"Zero-Width Space (U+200B)": 1})
            OfficeMetadataEngine.wipe_all(xlsx, strip_invisible_unicode=True)
            self.assertEqual(OfficeMetadataEngine.scan_invisible_unicode(xlsx), {})

            pptx = os.path.join(tmp, "slides.pptx")
            _build_zip(
                pptx,
                {
                    "docProps/core.xml": CORE_XML_SAMPLE,
                    "ppt/slides/slide1.xml": '<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><a:t>Diapo&#8203;1</a:t></p:sld>',
                },
            )
            self.assertEqual(OfficeMetadataEngine.scan_invisible_unicode(pptx), {"Zero-Width Space (U+200B)": 1})
            OfficeMetadataEngine.wipe_all(pptx, strip_invisible_unicode=True)
            self.assertEqual(OfficeMetadataEngine.scan_invisible_unicode(pptx), {})

    def test_security_rejects_doctype_utf8_utf16_badzip_and_size_limit(self):
        with tempfile.TemporaryDirectory() as tmp:
            xxe_utf8 = os.path.join(tmp, "xxe8.docx")
            _build_zip(
                xxe_utf8,
                {"docProps/core.xml": '<?xml version="1.0"?><!DOCTYPE foo [<!ENTITY x "y">]><cp:coreProperties/>'},
            )
            with self.assertRaises(FileProcessingError):
                OfficeMetadataEngine.update_tag(xxe_utf8, "Title", "Test")

            xxe_utf16 = os.path.join(tmp, "xxe16.docx")
            _build_zip(
                xxe_utf16,
                {
                    "docProps/core.xml": '<?xml version="1.0" encoding="utf-16"?><!DOCTYPE a><cp:coreProperties/>'.encode(
                        "utf-16"
                    )
                },
            )
            with self.assertRaises(FileProcessingError):
                OfficeMetadataEngine.update_tag(xxe_utf16, "Title", "Test")

            bad_zip = os.path.join(tmp, "corrupt.docx")
            with open(bad_zip, "wb") as f:
                f.write(b"not a zip")
            with self.assertRaises(FileProcessingError):
                OfficeMetadataEngine.scan_invisible_unicode(bad_zip)

            # Límite de lectura acotado por bytes reales descomprimidos
            big_xml = os.path.join(tmp, "big.docx")
            _build_zip(
                big_xml,
                {"word/document.xml": "<w:t>" + ("A" * 2048) + "</w:t>"},
            )
            with mock.patch("metaforge.office_engine.MAX_XML_BYTES", 512):
                with self.assertRaises(FileProcessingError):
                    OfficeMetadataEngine.scan_invisible_unicode(big_xml)


if __name__ == "__main__":
    unittest.main()
