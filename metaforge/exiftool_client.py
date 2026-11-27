# MetaForge - Inspector y Editor Forense de Metadatos
# Copyright (C) 2026 Kevin R. Clas (ClasK7)
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

"""Adaptador seguro de ExifTool con resolución multiplataforma y sanitización."""

from __future__ import annotations

import json
import logging
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
from pathlib import Path
from typing import Any

from metaforge import (
    ExifToolUnavailableError,
    FileProcessingError,
    ValidationError,
)

logger = logging.getLogger(__name__)

TAG_NAME_RE = re.compile(r"[A-Za-z][A-Za-z0-9_-]*(?::[A-Za-z][A-Za-z0-9_-]*)*")

# Pseudo-etiquetas de ExifTool que alteran nombres/rutas en disco o borran grupos completos.
RESERVED_TAG_SEGMENTS = {
    "sourcefile",
    "filename",
    "directory",
    "all",
    "hardlink",
    "testname",
    "geotag",
}


def _windows_startupinfo() -> subprocess.STARTUPINFO | None:
    if os.name == "nt":
        info = subprocess.STARTUPINFO()
        info.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        return info
    return None


def _install_hint() -> str:
    if sys.platform == "win32":
        return (
            "No se encontró un ejecutable funcional de ExifTool.\n\n"
            "En Windows 11:\n"
            "1. Descarga ExifTool para Windows desde https://exiftool.org/\n"
            "2. Extrae el paquete, renombra 'exiftool(-k).exe' a 'exiftool.exe'\n"
            "3. Coloca 'exiftool.exe' junto con la carpeta 'exiftool_files/' dentro de 'bin/'."
        )
    return (
        "No se encontró 'exiftool' en el sistema.\n\n"
        "En Kali Linux / Debian / Ubuntu ejecuta:\n"
        "sudo apt update && sudo apt install -y libimage-exiftool-perl python3-tk"
    )


def resolve_exiftool_path(project_root: Path) -> str | None:
    """Resuelve la ruta del ejecutable de ExifTool según el sistema operativo.

    En Windows 11 prioriza ``bin/exiftool.exe`` (con soporte para PyInstaller
    vía ``sys._MEIPASS``) y recurre al ``PATH`` como respaldo. En Linux/macOS
    ignora ``bin/`` y utiliza únicamente ``exiftool`` instalado en el sistema.

    Args:
        project_root: Directorio raíz del proyecto MetaForge.

    Returns:
        Ruta al ejecutable de ExifTool o ``None`` si no se localizó.
    """
    if sys.platform == "win32":
        candidates: list[Path] = []
        if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
            candidates.append(Path(sys._MEIPASS) / "bin" / "exiftool.exe")
        candidates.append(Path(project_root) / "bin" / "exiftool.exe")
        for candidate in candidates:
            if candidate.is_file():
                return str(candidate)
        return shutil.which("exiftool")

    return shutil.which("exiftool")


def verify_exiftool(binary_path: str | None) -> tuple[str, str]:
    """Comprueba que el binario de ExifTool exista y responda a ``-ver``.

    Usa ``stdin=subprocess.DEVNULL`` para evitar que una copia sin renombrar de
    ``exiftool(-k).exe`` en Windows quede bloqueada esperando confirmación.

    Args:
        binary_path: Ruta resuelta del binario o ``None``.

    Returns:
        Tupla ``(resolved_path, version_str)``.

    Raises:
        ExifToolUnavailableError: Si el binario falta, no responde o falla.
    """
    if not binary_path:
        raise ExifToolUnavailableError(_install_hint())

    try:
        res = subprocess.run(
            [binary_path, "-ver"],
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=5,
            startupinfo=_windows_startupinfo(),
        )
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError) as exc:
        logger.error("Fallo al verificar ExifTool (%s): %s", binary_path, exc)
        raise ExifToolUnavailableError(f"{_install_hint()}\n\nDetalle técnico: {exc}") from exc

    version = res.stdout.strip()
    if res.returncode != 0 or not version:
        detail = res.stderr.strip() or f"código de salida {res.returncode}"
        logger.error("ExifTool devolvió error en -ver: %s", detail)
        raise ExifToolUnavailableError(f"{_install_hint()}\n\nDetalle técnico: {detail}")

    return binary_path, version


def validate_tag_name(raw_tag: str) -> str:
    """Valida que un nombre de etiqueta cumpla la sintaxis permitida y segura.

    Rechaza operadores de ExifTool, saltos de línea y pseudo-etiquetas
    reservadas (comprobando el nombre completo y el segmento final tras ``:``).

    Args:
        raw_tag: Nombre de etiqueta con o sin prefijo de grupo.

    Returns:
        Nombre de etiqueta limpio y validado.

    Raises:
        ValidationError: Si la etiqueta está vacía, mal formada o es reservada.
    """
    if not isinstance(raw_tag, str):
        raise ValidationError("El nombre de la etiqueta debe ser una cadena de texto.")

    if any(ch in raw_tag for ch in ("\r", "\n", "\x00")):
        raise ValidationError("El nombre de la etiqueta contiene saltos de línea o caracteres de control no permitidos.")

    clean_tag = raw_tag.strip()
    if not clean_tag or not TAG_NAME_RE.fullmatch(clean_tag):
        raise ValidationError(
            f"Nombre de etiqueta inválido: '{raw_tag}'. "
            "Usa solo letras, números, guiones y opcionalmente 'Grupo:Etiqueta'."
        )

    lower_full = clean_tag.lower()
    lower_last = clean_tag.split(":")[-1].lower()
    if lower_full in RESERVED_TAG_SEGMENTS or lower_last in RESERVED_TAG_SEGMENTS:
        raise ValidationError(
            f"La etiqueta '{clean_tag}' está reservada por el sistema y no puede editarse."
        )

    return clean_tag


class ExifToolClient:
    """Cliente seguro para invocar ExifTool en lectura, escritura y purga total.

    La escritura de etiquetas emplea ``-json=<archivo_utf8.json>`` para que
    ningún valor (incluyendo saltos de línea o ``#``) pueda interpretarse como
    opción de línea de comandos, y todas las operaciones pasan la ruta absoluta
    mediante un archivo temporal UTF-8 ``-@`` precedido por ``--``.
    """

    def __init__(self, binary_path: str, work_dir: Path | None = None) -> None:
        """Inicializa el cliente con la ruta verificada de ExifTool.

        Args:
            binary_path: Ruta al ejecutable de ExifTool.
            work_dir: Directorio privado opcional para archivos temporales.
        """
        self.binary_path = binary_path
        self.work_dir = Path(work_dir) if work_dir is not None else None
        self._proc_lock = threading.Lock()
        self._active_proc: subprocess.Popen[str] | None = None

    @staticmethod
    def _validate_file_path(file_path: str) -> str:
        if not file_path or any(ch in file_path for ch in ("\r", "\n", "\x00")):
            raise ValidationError("La ruta del archivo contiene caracteres de control no permitidos.")
        return os.path.abspath(file_path)

    def _create_temp_file(self, suffix: str) -> str:
        target_dir = str(self.work_dir) if self.work_dir and self.work_dir.is_dir() else None
        fd, path = tempfile.mkstemp(prefix="mf_et_", suffix=suffix, dir=target_dir)
        os.close(fd)
        return path

    def _run_with_argfile(
        self,
        flags: list[str],
        file_path: str,
        timeout: int,
        json_payload: list[dict[str, Any]] | None = None,
    ) -> tuple[int, str, str]:
        abspath = self._validate_file_path(file_path)
        argfile_path = self._create_temp_file(".args")
        json_path: str | None = None

        try:
            arg_lines = ["-charset", "utf8", "-charset", "filename=utf8", *flags]
            if json_payload is not None:
                json_path = self._create_temp_file(".json")
                with open(json_path, "w", encoding="utf-8", newline="\n") as jf:
                    json.dump(json_payload, jf, ensure_ascii=False)
                # -json= importa valores estructurados evitando que saltos de línea
                # o '#' en los datos se interpreten como opciones CLI de ExifTool.
                arg_lines.append(f"-json={json_path}")

            arg_lines.extend(["--", abspath])
            with open(argfile_path, "w", encoding="utf-8", newline="\n") as af:
                af.write("\n".join(arg_lines) + "\n")

            # -charset filename=utf8 antes de -@ garantiza abrir rutas UTF-8 en Windows.
            cmd = [self.binary_path, "-charset", "filename=utf8", "-@", argfile_path]
            try:
                proc = subprocess.Popen(
                    cmd,
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    startupinfo=_windows_startupinfo(),
                )
            except FileNotFoundError as exc:
                raise ExifToolUnavailableError(_install_hint()) from exc
            except OSError as exc:
                raise FileProcessingError(f"No se pudo iniciar ExifTool: {exc}") from exc

            with self._proc_lock:
                self._active_proc = proc

            try:
                stdout, stderr = proc.communicate(timeout=timeout)
                return proc.returncode, stdout, stderr
            except subprocess.TimeoutExpired as exc:
                self._kill_process(proc)
                raise FileProcessingError(
                    f"ExifTool superó el tiempo límite ({timeout} s) al procesar el archivo."
                ) from exc
            finally:
                with self._proc_lock:
                    if self._active_proc is proc:
                        self._active_proc = None
        finally:
            for tmp in (argfile_path, json_path):
                if tmp and os.path.exists(tmp):
                    try:
                        os.remove(tmp)
                    except OSError:
                        pass

    @staticmethod
    def _kill_process(proc: subprocess.Popen[str]) -> None:
        try:
            proc.terminate()
            proc.wait(timeout=2)
        except subprocess.TimeoutExpired:
            try:
                proc.kill()
                proc.wait(timeout=2)
            except OSError:
                pass
        except OSError:
            pass

    @staticmethod
    def _evaluate_write_result(returncode: int, stdout: str, stderr: str) -> list[str]:
        warnings = [line.strip() for line in stderr.splitlines() if line.strip()]
        if returncode != 0:
            detail = "\n".join(warnings) or stdout.strip() or f"código de salida {returncode}"
            raise FileProcessingError(f"ExifTool falló al escribir metadatos:\n{detail}")

        stdout_lower = stdout.lower()
        if "0 image files updated" in stdout_lower:
            # Archivo ya limpio o valor idéntico sin errores reales -> éxito idempotente.
            if "image files unchanged" in stdout_lower and not warnings:
                return []
            if warnings:
                raise FileProcessingError(
                    "ExifTool no pudo actualizar el archivo:\n" + "\n".join(warnings)
                )

        return warnings

    def read_metadata(self, file_path: str) -> dict[str, Any]:
        """Lee todos los metadatos del archivo en formato JSON con grupos.

        Args:
            file_path: Ruta al archivo objetivo.

        Returns:
            Diccionario ``{Grupo:Etiqueta: valor}`` devuelto por ExifTool.

        Raises:
            FileProcessingError: Si ExifTool falla, expira el tiempo o devuelve
                JSON inválido o un error interno de lectura.
        """
        returncode, stdout, stderr = self._run_with_argfile(["-j", "-G"], file_path, timeout=30)
        if returncode != 0:
            detail = stderr.strip() or stdout.strip() or f"código de salida {returncode}"
            raise FileProcessingError(f"No se pudieron leer los metadatos:\n{detail}")

        try:
            parsed = json.loads(stdout)
        except json.JSONDecodeError as exc:
            raise FileProcessingError(f"Salida JSON inválida de ExifTool: {exc}") from exc

        if not isinstance(parsed, list) or not parsed or not isinstance(parsed[0], dict):
            raise FileProcessingError("ExifTool no devolvió metadatos válidos para el archivo.")

        metadata = parsed[0]
        if "ExifTool:Error" in metadata:
            raise FileProcessingError(f"Error reportado por ExifTool: {metadata['ExifTool:Error']}")

        return metadata

    def write_tags(self, file_path: str, updates: dict[str, str]) -> list[str]:
        """Escribe o borra una o varias etiquetas en ``file_path`` vía ``-json=``.

        Args:
            file_path: Ruta al archivo de trabajo a modificar.
            updates: Mapa ``{etiqueta: nuevo_valor}`` (cadena vacía para borrar).

        Returns:
            Lista de advertencias no fatales emitidas por ``stderr``.

        Raises:
            ValidationError: Si alguna etiqueta no pasa ``validate_tag_name``.
            FileProcessingError: Si ExifTool falla o no actualiza el archivo.
        """
        if not updates:
            return []

        validated: dict[str, str] = {}
        for raw_tag, value in updates.items():
            clean_tag = validate_tag_name(raw_tag)
            validated[clean_tag] = "" if value is None else str(value)

        payload = [{"SourceFile": "*", **validated}]
        returncode, stdout, stderr = self._run_with_argfile(
            ["-overwrite_original"],
            file_path,
            timeout=60,
            json_payload=payload,
        )
        return self._evaluate_write_result(returncode, stdout, stderr)

    def wipe_all(self, file_path: str) -> list[str]:
        """Elimina todos los metadatos editables mediante ``exiftool -all=``.

        Args:
            file_path: Ruta al archivo de trabajo a limpiar.

        Returns:
            Lista de advertencias no fatales emitidas por ``stderr``.

        Raises:
            FileProcessingError: Si ExifTool falla al ejecutar la limpieza.
        """
        returncode, stdout, stderr = self._run_with_argfile(
            ["-overwrite_original", "-all="],
            file_path,
            timeout=60,
        )
        return self._evaluate_write_result(returncode, stdout, stderr)

    def terminate_active(self) -> None:
        """Termina el subproceso activo de ExifTool si hay uno en ejecución."""
        with self._proc_lock:
            proc = self._active_proc
        if proc is not None and proc.poll() is None:
            logger.info("Terminando proceso activo de ExifTool.")
            self._kill_process(proc)
