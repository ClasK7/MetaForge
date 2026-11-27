# MetaForge - Pruebas unitarias de exiftool_client
# Copyright (C) 2026 Kevin R. Clas (ClasK7) - GPLv3

import json
import os
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from metaforge import ExifToolUnavailableError, FileProcessingError, ValidationError
from metaforge.exiftool_client import (
    ExifToolClient,
    resolve_exiftool_path,
    validate_tag_name,
    verify_exiftool,
)


class _FakeProc:
    def __init__(self, returncode=0, stdout="", stderr="", timeout_on_comm=False):
        self.returncode = returncode
        self._stdout = stdout
        self._stderr = stderr
        self._timeout = timeout_on_comm
        self.terminated = False
        self.killed = False
        self._alive = True

    def poll(self):
        return None if self._alive else self.returncode

    def communicate(self, timeout=None):
        if self._timeout:
            raise subprocess.TimeoutExpired(cmd="exiftool", timeout=timeout)
        self._alive = False
        return self._stdout, self._stderr

    def terminate(self):
        self.terminated = True
        self._alive = False

    def kill(self):
        self.killed = True
        self._alive = False

    def wait(self, timeout=None):
        return self.returncode


class TestExifToolValidationAndResolution(unittest.TestCase):
    def test_validate_tag_name_accepts_valid_tags(self):
        self.assertEqual(validate_tag_name("Artist"), "Artist")
        self.assertEqual(validate_tag_name(" XMP:Creator "), "XMP:Creator")
        self.assertEqual(validate_tag_name("EXIF:CreateDate"), "EXIF:CreateDate")
        self.assertEqual(validate_tag_name("XMP-dc:Title"), "XMP-dc:Title")

    def test_validate_tag_name_rejects_injections_and_reserved(self):
        invalid_tags = [
            "",
            "   ",
            "-config",
            "Comment<",
            "Tag>",
            "Artist=foo",
            "Tag\n-config",
            "Artist\n",
            "@argfile",
            "SourceFile",
            "FileName",
            "Directory",
            "EXIF:All",
            "XMP:Geotag",
            "HardLink",
            "TestName",
        ]
        for bad in invalid_tags:
            with self.subTest(tag=bad):
                with self.assertRaises(ValidationError):
                    validate_tag_name(bad)

    def test_resolve_exiftool_path_ignores_bin_on_linux(self):
        with tempfile.TemporaryDirectory() as tmp:
            bin_dir = Path(tmp) / "bin"
            bin_dir.mkdir()
            (bin_dir / "exiftool.exe").write_bytes(b"MZ")
            with mock.patch("metaforge.exiftool_client.sys.platform", "linux"):
                with mock.patch("metaforge.exiftool_client.shutil.which", return_value="/usr/bin/exiftool"):
                    self.assertEqual(resolve_exiftool_path(Path(tmp)), "/usr/bin/exiftool")

    def test_resolve_exiftool_path_prefers_local_bin_on_windows(self):
        with tempfile.TemporaryDirectory() as tmp:
            bin_dir = Path(tmp) / "bin"
            bin_dir.mkdir()
            exe = bin_dir / "exiftool.exe"
            exe.write_bytes(b"MZ")
            with mock.patch("metaforge.exiftool_client.sys.platform", "win32"):
                with mock.patch("metaforge.exiftool_client.getattr", side_effect=getattr):
                    self.assertEqual(resolve_exiftool_path(Path(tmp)), str(exe))

    def test_verify_exiftool_success_and_missing(self):
        completed = subprocess.CompletedProcess(["exiftool", "-ver"], 0, stdout="12.76\n", stderr="")
        with mock.patch("metaforge.exiftool_client.subprocess.run", return_value=completed) as m_run:
            path, ver = verify_exiftool("/usr/bin/exiftool")
            self.assertEqual((path, ver), ("/usr/bin/exiftool", "12.76"))
            self.assertEqual(m_run.call_args.kwargs["stdin"], subprocess.DEVNULL)

        with mock.patch("metaforge.exiftool_client.sys.platform", "linux"):
            with self.assertRaises(ExifToolUnavailableError) as ctx:
                verify_exiftool(None)
            self.assertIn("libimage-exiftool-perl", str(ctx.exception))

        with mock.patch("metaforge.exiftool_client.sys.platform", "win32"):
            with mock.patch("metaforge.exiftool_client.subprocess.run", side_effect=FileNotFoundError):
                with self.assertRaises(ExifToolUnavailableError) as ctx:
                    verify_exiftool("bin/exiftool.exe")
                self.assertIn("exiftool_files", str(ctx.exception))


class TestExifToolClientExecution(unittest.TestCase):
    def test_write_tags_uses_json_and_argfile_with_multiline_and_hash_values(self):
        captured_argfile = []
        captured_json = []
        created_paths = []

        def fake_popen(cmd, **kwargs):
            self.assertEqual(kwargs["stdin"], subprocess.DEVNULL)
            self.assertIn("-@", cmd)
            argfile_path = cmd[cmd.index("-@") + 1]
            created_paths.append(argfile_path)
            content = Path(argfile_path).read_text(encoding="utf-8").splitlines()
            captured_argfile.extend(content)
            json_opt = next(line for line in content if line.startswith("-json="))
            json_path = json_opt.split("=", 1)[1]
            created_paths.append(json_path)
            captured_json.extend(json.loads(Path(json_path).read_text(encoding="utf-8")))
            return _FakeProc(returncode=0, stdout="    1 image files updated\n", stderr="")

        with tempfile.TemporaryDirectory() as tmp:
            tricky_file = os.path.join(tmp, "-foto con ñ y tilde á.jpg")
            client = ExifToolClient("exiftool", work_dir=Path(tmp))
            with mock.patch("metaforge.exiftool_client.subprocess.Popen", side_effect=fake_popen):
                warnings = client.write_tags(
                    tricky_file,
                    {"EXIF:ImageDescription": "línea 1\n-config=malo.cfg\n#comentario"},
                )

            self.assertEqual(warnings, [])
            self.assertIn("--", captured_argfile)
            self.assertEqual(captured_argfile[-1], os.path.abspath(tricky_file))
            self.assertEqual(captured_json[0]["SourceFile"], "*")
            self.assertEqual(
                captured_json[0]["EXIF:ImageDescription"],
                "línea 1\n-config=malo.cfg\n#comentario",
            )
            # Todos los temporales (-@ y -json) deben quedar borrados.
            for p in created_paths:
                self.assertFalse(os.path.exists(p))

    def test_temp_files_cleaned_on_exception(self):
        created_paths = []

        def fake_popen(cmd, **kwargs):
            argfile_path = cmd[cmd.index("-@") + 1]
            created_paths.append(argfile_path)
            return _FakeProc(timeout_on_comm=True)

        with tempfile.TemporaryDirectory() as tmp:
            client = ExifToolClient("exiftool", work_dir=Path(tmp))
            with mock.patch("metaforge.exiftool_client.subprocess.Popen", side_effect=fake_popen):
                with self.assertRaises(FileProcessingError):
                    client.read_metadata(os.path.join(tmp, "a.jpg"))
            for p in created_paths:
                self.assertFalse(os.path.exists(p))

    def test_read_metadata_errors(self):
        client = ExifToolClient("exiftool")
        with mock.patch(
            "metaforge.exiftool_client.subprocess.Popen",
            return_value=_FakeProc(returncode=0, stdout="not-json"),
        ):
            with self.assertRaises(FileProcessingError):
                client.read_metadata("sample.jpg")

        with mock.patch(
            "metaforge.exiftool_client.subprocess.Popen",
            return_value=_FakeProc(
                returncode=0,
                stdout=json.dumps([{"SourceFile": "sample.jpg", "ExifTool:Error": "File format error"}]),
            ),
        ):
            with self.assertRaises(FileProcessingError):
                client.read_metadata("sample.jpg")

        with mock.patch(
            "metaforge.exiftool_client.subprocess.Popen",
            side_effect=FileNotFoundError("no binary"),
        ):
            with self.assertRaises(ExifToolUnavailableError):
                client.read_metadata("sample.jpg")

    def test_write_idempotent_vs_false_success(self):
        client = ExifToolClient("exiftool")
        # Archivo ya limpio -> éxito idempotente
        with mock.patch(
            "metaforge.exiftool_client.subprocess.Popen",
            return_value=_FakeProc(
                returncode=0,
                stdout="    0 image files updated\n    1 image files unchanged\n",
                stderr="",
            ),
        ):
            self.assertEqual(client.wipe_all("clean.jpg"), [])

        # Falso éxito con advertencia/error en stderr -> debe elevar FileProcessingError
        with mock.patch(
            "metaforge.exiftool_client.subprocess.Popen",
            return_value=_FakeProc(
                returncode=0,
                stdout="    0 image files updated\n",
                stderr="Warning: Sorry, FakeTag is not writable\n",
            ),
        ):
            with self.assertRaises(FileProcessingError):
                client.write_tags("photo.jpg", {"EXIF:FakeTag": "val"})

    def test_terminate_active_stops_running_process(self):
        client = ExifToolClient("exiftool")
        proc = _FakeProc()
        client._active_proc = proc
        client.terminate_active()
        self.assertTrue(proc.terminated)


if __name__ == "__main__":
    unittest.main()
