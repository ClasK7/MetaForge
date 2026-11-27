# MetaForge - Pruebas unitarias de session
# Copyright (C) 2026 Kevin R. Clas (ClasK7) - GPLv3

import errno
import hashlib
import json
import os
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from metaforge import (
    DestinationNotWritableError,
    FileProcessingError,
    ValidationError,
)
from metaforge.exiftool_client import ExifToolClient
from metaforge.scanner import UNICODE_FINDING_TAG
from metaforge.session import DocumentSession, resolve_output_path
from tests.test_office_engine import (
    APP_XML_SAMPLE,
    CORE_XML_SAMPLE,
    CUSTOM_XML_SAMPLE,
    DOC_XML_WITH_MC_IGNORABLE,
    _build_zip,
)


def _sha256(path: str) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


class _StubExifToolClient(ExifToolClient):
    def __init__(self, metadata_by_path=None, fail_write=False):
        super().__init__("exiftool")
        self.metadata_by_path = metadata_by_path or {}
        self.fail_write = fail_write
        self.write_calls = []
        self.wipe_calls = []

    def read_metadata(self, file_path: str):
        if file_path in self.metadata_by_path:
            return dict(self.metadata_by_path[file_path])
        return dict(self.metadata_by_path.get("*", {"SourceFile": file_path, "EXIF:Artist": "Original"}))

    def write_tags(self, file_path: str, updates: dict[str, str]):
        if self.fail_write:
            raise FileProcessingError("fallo simulado de escritura")
        self.write_calls.append((file_path, dict(updates)))
        # Simular mutación en el temporal
        Path(file_path).write_bytes(Path(file_path).read_bytes() + b"_MOD")
        cur = self.read_metadata(file_path)
        cur.update(updates)
        self.metadata_by_path[file_path] = cur
        return []

    def wipe_all(self, file_path: str):
        if self.fail_write:
            raise FileProcessingError("fallo simulado de purga")
        self.wipe_calls.append(file_path)
        Path(file_path).write_bytes(b"WIPED")
        self.metadata_by_path[file_path] = {"SourceFile": file_path, "File:FileType": "JPEG"}
        return []


class TestResolveOutputPath(unittest.TestCase):
    def test_naming_sequences_dotfiles_multi_dots_and_normalization(self):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            self.assertEqual(resolve_output_path(str(t / "foto.jpg")), t / "foto_procesado.jpg")
            (t / "foto_procesado.jpg").write_bytes(b"1")
            self.assertEqual(resolve_output_path(str(t / "foto.jpg")), t / "foto_procesado_1.jpg")
            (t / "foto_procesado_1.jpg").write_bytes(b"2")
            self.assertEqual(resolve_output_path(str(t / "foto.jpg")), t / "foto_procesado_2.jpg")

            # Normalización de sufijo _procesado incluso en target_dir vacío
            empty_dir = t / "vacia"
            empty_dir.mkdir()
            self.assertEqual(
                resolve_output_path(str(t / "doc_procesado.docx"), str(empty_dir)),
                empty_dir / "doc_procesado_1.docx",
            )
            self.assertEqual(
                resolve_output_path(str(t / "doc_procesado_4.docx"), str(empty_dir)),
                empty_dir / "doc_procesado_1.docx",
            )

            # Sin extensión, archivos ocultos y múltiples puntos
            self.assertEqual(resolve_output_path(str(t / "README")), t / "README_procesado")
            self.assertEqual(resolve_output_path(str(t / ".gitignore")), t / ".gitignore_procesado")
            self.assertEqual(resolve_output_path(str(t / ".foto.jpg")), t / ".foto_procesado.jpg")
            self.assertEqual(resolve_output_path(str(t / "a.b.c")), t / "a.b_procesado.c")


class TestDocumentSession(unittest.TestCase):
    def test_lazy_temp_copy_save_preserves_original_sha256_and_returns_to_original(self):
        with tempfile.TemporaryDirectory() as tmp:
            orig = os.path.join(tmp, "foto.jpg")
            Path(orig).write_bytes(b"ORIGINAL_IMAGE_BYTES")
            orig_hash = _sha256(orig)

            client = _StubExifToolClient(
                {orig: {"SourceFile": orig, "File:FileType": "JPEG", "EXIF:Artist": "Ana"}}
            )
            session = DocumentSession(client)
            session.open_file(orig)

            # Cero copias temporales al inspeccionar
            self.assertIsNone(session.working_temp_path)
            self.assertFalse(session.has_pending_changes)

            # Primer cambio crea copia temporal perezosa
            session.apply_edit("EXIF:Artist", "Anonimo")
            self.assertIsNotNone(session.working_temp_path)
            temp_dir = session._session_temp_dir
            self.assertTrue(os.path.isdir(temp_dir))
            self.assertTrue(session.has_pending_changes)
            self.assertEqual(session.metadata["EXIF:Artist"], "Anonimo")

            # Guardar copia exporta _procesado, borra temporal y recarga metadatos originales
            saved = session.save_copy()
            self.assertEqual(saved, os.path.join(tmp, "foto_procesado.jpg"))
            self.assertTrue(os.path.exists(saved))
            self.assertEqual(_sha256(orig), orig_hash)
            self.assertFalse(os.path.exists(temp_dir))
            self.assertFalse(session.has_pending_changes)
            self.assertEqual(session.metadata["EXIF:Artist"], "Ana")

    def test_first_change_failure_rolls_back_temp_dir(self):
        with tempfile.TemporaryDirectory() as tmp:
            orig = os.path.join(tmp, "foto.jpg")
            Path(orig).write_bytes(b"DATA")
            client = _StubExifToolClient(
                {orig: {"SourceFile": orig, "EXIF:Artist": "Ana"}},
                fail_write=True,
            )
            session = DocumentSession(client)
            session.open_file(orig)
            with self.assertRaises(FileProcessingError):
                session.apply_edit("EXIF:Artist", "Nuevo")
            self.assertIsNone(session.working_temp_path)
            self.assertIsNone(session._session_temp_dir)
            self.assertFalse(session.has_pending_changes)

    def test_office_selective_clean_and_wipe_all_on_working_copy(self):
        with tempfile.TemporaryDirectory() as tmp:
            orig = os.path.join(tmp, "informe.docx")
            _build_zip(
                orig,
                {
                    "docProps/core.xml": CORE_XML_SAMPLE,
                    "docProps/app.xml": APP_XML_SAMPLE,
                    "docProps/custom.xml": CUSTOM_XML_SAMPLE,
                    "word/document.xml": DOC_XML_WITH_MC_IGNORABLE,
                },
            )
            orig_hash = _sha256(orig)

            meta = {
                "SourceFile": orig,
                "File:FileType": "DOCX",
                "XMP:Creator": "Ana",
                "XML:LastModifiedBy": "Luis",
                "XML:Words": 400,
                "XML:TotalEditTime": "1 minute",
                "XML:RevisionNumber": 1,
            }
            client = _StubExifToolClient({"*": meta})
            session = DocumentSession(client)
            session.open_file(orig)
            self.assertTrue(session.is_office)
            self.assertFalse(session.is_pdf)

            # Escaneo bajo demanda detecta Unicode y hallazgos editables
            findings, _, uni_hits = session.run_scan()
            self.assertTrue(any(f["tag"] == UNICODE_FINDING_TAG for f in findings))
            self.assertTrue(bool(uni_hits))

            cleaned = session.apply_selective_clean()
            self.assertGreaterEqual(cleaned, 2)
            # Re-escaneo sobre working_temp_path ya no debe encontrar Unicode
            _, _, uni_after = session.run_scan()
            self.assertEqual(uni_after, {})

            # Limpieza total + guardar copia deja intacto el original
            session.apply_wipe_all()
            saved = session.save_copy()
            self.assertEqual(_sha256(orig), orig_hash)
            with zipfile.ZipFile(saved, "r") as zf:
                self.assertIn(b"<Properties", zf.read("docProps/custom.xml"))
                self.assertNotIn(b"ProyectoX", zf.read("docProps/custom.xml"))

    def test_selective_clean_returns_zero_without_creating_temp_when_nothing_editable(self):
        with tempfile.TemporaryDirectory() as tmp:
            orig = os.path.join(tmp, "foto.jpg")
            Path(orig).write_bytes(b"JPG")
            client = _StubExifToolClient({orig: {"SourceFile": orig, "File:FileType": "JPEG"}})
            session = DocumentSession(client)
            session.open_file(orig)
            session.findings = [
                {"severity": "ai_high", "tag": "JUMBF:C2PA", "value": "x", "reason": "r", "editable": False}
            ]
            self.assertEqual(session.apply_selective_clean(), 0)
            self.assertIsNone(session.working_temp_path)

    def test_pdf_detected_by_filetype_without_extension_and_export_json(self):
        with tempfile.TemporaryDirectory() as tmp:
            orig = os.path.join(tmp, "evidencia_sin_ext")
            Path(orig).write_bytes(b"%PDF-1.7")
            client = _StubExifToolClient(
                {orig: {"SourceFile": orig, "File:FileType": "PDF", "PDF:Author": "Ana"}}
            )
            session = DocumentSession(client)
            session.open_file(orig)
            self.assertTrue(session.is_pdf)

            out_json = os.path.join(tmp, "meta.json")
            session.export_json(out_json)
            loaded = json.loads(Path(out_json).read_text(encoding="utf-8"))
            self.assertEqual(loaded["PDF:Author"], "Ana")

            with mock.patch("builtins.open", side_effect=OSError("disk error")):
                with self.assertRaises(FileProcessingError):
                    session.export_json(out_json)

    def test_save_copy_readonly_enospc_and_samefile_guard(self):
        with tempfile.TemporaryDirectory() as tmp:
            orig = os.path.join(tmp, "foto.jpg")
            Path(orig).write_bytes(b"DATA")
            client = _StubExifToolClient({orig: {"SourceFile": orig, "EXIF:Artist": "Ana"}})
            session = DocumentSession(client)
            session.open_file(orig)
            session.apply_edit("EXIF:Artist", "Beto")

            # Carpeta solo lectura -> DestinationNotWritableError
            with mock.patch("metaforge.session.tempfile.mkstemp", side_effect=PermissionError("denied")):
                with self.assertRaises(DestinationNotWritableError):
                    session.save_copy()

            # ENOSPC -> FileProcessingError sin temporales huérfanos
            with mock.patch(
                "metaforge.session.shutil.copy2",
                side_effect=OSError(errno.ENOSPC, "No space left on device"),
            ):
                with self.assertRaises(FileProcessingError):
                    session.save_copy()
            leftovers = [p for p in os.listdir(tmp) if p.startswith(".metaforge_")]
            self.assertEqual(leftovers, [])

            # Guarda samefile impide sobrescribir original
            with mock.patch("metaforge.session.resolve_output_path", return_value=Path(orig)):
                with self.assertRaises(ValidationError):
                    session.save_copy()

            session.discard_changes()
            self.assertFalse(session.has_pending_changes)
            self.assertEqual(session.metadata["EXIF:Artist"], "Ana")


if __name__ == "__main__":
    unittest.main()
