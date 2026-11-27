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

"""Gestor de sesión no destructiva con copia temporal perezosa y rutas _procesado."""

from __future__ import annotations

import atexit
import errno
import json
import logging
import os
import re
import shutil
import tempfile
from pathlib import Path
from typing import Any

from metaforge import (
    DestinationNotWritableError,
    FileProcessingError,
    ValidationError,
)
from metaforge.exiftool_client import ExifToolClient
from metaforge.office_engine import (
    APP_TAG_MAP,
    CORE_TAG_MAP,
    OFFICE_EXTENSIONS,
    OfficeMetadataEngine,
)
from metaforge.scanner import (
    ForensicScanner,
    UNICODE_FINDING_TAG,
    is_tag_editable as _scanner_is_tag_editable,
)

logger = logging.getLogger(__name__)

_PROCESADO_SUFFIX_RE = re.compile(r"_procesado(?:_\d+)?$")
_OFFICE_FILETYPES = {"DOCX", "XLSX", "PPTX"}


def _split_stem_and_ext(filename: str) -> tuple[str, str]:
    """Separa nombre base y extensión respetando archivos ocultos estilo Unix."""
    if filename.startswith(".") and filename.count(".") == 1:
        return filename, ""
    return os.path.splitext(filename)


def resolve_output_path(original_path: str, target_dir: str | None = None) -> Path:
    """Calcula la ruta libre con sufijo ``_procesado`` sin crear archivos en disco.

    Si el nombre original ya termina en ``_procesado`` o ``_procesado_<N>``,
    normaliza la raíz y busca desde ``_procesado_1`` (incluso en una carpeta
    vacía). En caso contrario prueba ``<base>_procesado<ext>`` y luego
    ``_procesado_1``, ``_procesado_2``, etc.

    Args:
        original_path: Ruta del archivo original cargado en sesión.
        target_dir: Carpeta destino opcional; si se omite usa la del original.

    Returns:
        Ruta ``Path`` del primer nombre disponible en el directorio destino.
    """
    src = Path(original_path)
    dest_folder = Path(target_dir) if target_dir is not None else src.parent
    stem, ext = _split_stem_and_ext(src.name)

    suffix_match = _PROCESADO_SUFFIX_RE.search(stem)
    if suffix_match:
        base = stem[: suffix_match.start()]
        idx = 1
        while True:
            candidate = dest_folder / f"{base}_procesado_{idx}{ext}"
            if not candidate.exists():
                return candidate
            idx += 1

    first_candidate = dest_folder / f"{stem}_procesado{ext}"
    if not first_candidate.exists():
        return first_candidate

    idx = 1
    while True:
        candidate = dest_folder / f"{stem}_procesado_{idx}{ext}"
        if not candidate.exists():
            return candidate
        idx += 1


class DocumentSession:
    """Coordina el ciclo de vida no destructivo de un único documento abierto.

    La lectura inicial se realiza directamente sobre ``original_path`` sin
    generar copias en disco. Al aplicar el primer cambio se crea un directorio
    temporal privado (``0700`` en POSIX) con la copia de trabajo. Al guardar
    la copia final se escribe atómicamente en ``_procesado[_N]``, se borra el
    temporal y se recargan los metadatos originales desde ``original_path``.
    """

    def __init__(self, exiftool_client: ExifToolClient) -> None:
        """Inicializa una sesión vacía asociada a un cliente ExifTool.

        Args:
            exiftool_client: Instancia verificada de ``ExifToolClient``.
        """
        self.exiftool_client = exiftool_client
        self.original_path: str | None = None
        self.working_temp_path: str | None = None
        self._session_temp_dir: str | None = None
        self.pending_change_count: int = 0
        self.metadata: dict[str, Any] = {}
        self.unicode_hits: dict[str, int] = {}
        self.findings: list[dict[str, Any]] = []
        self.flagged_tags: dict[str, str] = {}
        self.is_office: bool = False
        self.is_pdf: bool = False

    @property
    def active_path(self) -> str | None:
        """Ruta actualmente activa: la copia temporal si existe, o el archivo original."""
        return self.working_temp_path or self.original_path

    @property
    def has_pending_changes(self) -> bool:
        """Indica si existen cambios sin exportar en la copia temporal de trabajo."""
        return self.working_temp_path is not None and self.pending_change_count > 0

    def is_tag_editable(self, full_tag: str, raw_value: Any = "") -> bool:
        """Fachada única de editabilidad para la interfaz y las mutaciones de sesión.

        Args:
            full_tag: Nombre de etiqueta con o sin grupo.
            raw_value: Valor actual devuelto por ExifTool (si existe).

        Returns:
            ``True`` si el campo puede editarse para el documento abierto.
        """
        if not self.active_path:
            return False
        effective_path = self.active_path
        if self.is_office and os.path.splitext(effective_path.lower())[1] not in OFFICE_EXTENSIONS:
            effective_path = effective_path + ".docx"
        return _scanner_is_tag_editable(effective_path, full_tag, raw_value)

    def _detect_format_flags(self, file_path: str, metadata: dict[str, Any]) -> None:
        _, ext = os.path.splitext(file_path.lower())
        filetype = str(metadata.get("File:FileType", "")).strip().upper()
        self.is_office = ext in OFFICE_EXTENSIONS or filetype in _OFFICE_FILETYPES
        self.is_pdf = ext == ".pdf" or filetype == "PDF"

    def open_file(self, file_path: str) -> dict[str, Any]:
        """Abre un archivo en modo solo lectura inicial sin generar copias en disco.

        Args:
            file_path: Ruta al archivo a inspeccionar.

        Returns:
            Diccionario de metadatos leídos desde ``file_path``.
        """
        self.cleanup()
        metadata = self.exiftool_client.read_metadata(file_path)
        self.original_path = file_path
        self.metadata = metadata
        self.unicode_hits = {}
        self.findings = []
        self.flagged_tags = {}
        self.pending_change_count = 0
        self._detect_format_flags(file_path, metadata)
        return metadata

    def ensure_working_copy(self) -> str:
        """Crea de forma perezosa la copia temporal de trabajo dentro de un directorio privado.

        Returns:
            Ruta absoluta de la copia temporal de trabajo.

        Raises:
            ValidationError: Si no hay archivo cargado en sesión.
            FileProcessingError: Si falla la copia (por ejemplo ``ENOSPC``).
        """
        if not self.original_path:
            raise ValidationError("No hay ningún archivo cargado en la sesión.")
        if self.working_temp_path and os.path.exists(self.working_temp_path):
            return self.working_temp_path

        temp_dir = tempfile.mkdtemp(prefix="metaforge_")
        if os.name != "nt":
            try:
                os.chmod(temp_dir, 0o700)
            except OSError:
                pass

        _, ext = _split_stem_and_ext(Path(self.original_path).name)
        temp_file = os.path.join(temp_dir, f"working{ext}")
        try:
            shutil.copy2(self.original_path, temp_file)
        except OSError as exc:
            shutil.rmtree(temp_dir, ignore_errors=True)
            if exc.errno == errno.ENOSPC:
                raise FileProcessingError(
                    "No hay espacio suficiente en el directorio temporal del sistema para crear la copia de trabajo."
                ) from exc
            raise FileProcessingError(f"No se pudo crear la copia temporal de trabajo: {exc}") from exc

        self._session_temp_dir = temp_dir
        self.working_temp_path = temp_file
        self.exiftool_client.work_dir = Path(temp_dir)
        atexit.register(self.cleanup)
        return temp_file

    def _reload_after_mutation(self) -> None:
        assert self.working_temp_path is not None
        self.metadata = self.exiftool_client.read_metadata(self.working_temp_path)
        self.unicode_hits = {}
        self.findings = []
        self.flagged_tags = {}
        self.pending_change_count += 1

    def apply_edit(self, tag: str, value: str) -> list[str]:
        """Aplica una edición manual sobre la copia temporal de trabajo.

        Si falla al intentar aplicar el primer cambio, revierte la creación del
        directorio temporal dejando ``has_pending_changes`` en ``False``.

        Args:
            tag: Nombre de etiqueta a modificar o vaciar.
            value: Nuevo valor textual.

        Returns:
            Lista de advertencias no fatales devueltas por ExifTool (vacía en Office).

        Raises:
            ValidationError: Si no hay archivo abierto o la etiqueta es de solo lectura.
        """
        if not self.original_path:
            raise ValidationError("Primero selecciona un archivo.")

        existing_val = self.metadata.get(tag, "")
        if not self.is_tag_editable(tag, existing_val):
            raise ValidationError(f"La etiqueta '{tag}' es de solo lectura para este archivo.")

        had_temp_before = self.working_temp_path is not None
        working_path = self.ensure_working_copy()
        warnings: list[str] = []
        try:
            if self.is_office:
                OfficeMetadataEngine.update_tag(working_path, tag, value)
            else:
                warnings = self.exiftool_client.write_tags(working_path, {tag: value})
            self._reload_after_mutation()
        except Exception:
            if not had_temp_before:
                self.cleanup()
            raise
        return warnings

    def apply_selective_clean(self) -> int:
        """Limpia en la copia temporal únicamente los hallazgos editables del último escaneo.

        En documentos Office asigna ``RevisionNumber="1"`` y ``TotalEditTime="0"``
        para mantener validez OOXML, omite la clave sintética ``[Contenido Office]``
        y aplica saneamiento Unicode si existían hallazgos invisibles.

        Returns:
            Número de elementos saneados (etiquetas limpiadas más 1 si se limpió
            Unicode invisible). Devuelve ``0`` sin crear temporal si no hay nada
            editable que limpiar.
        """
        if not self.original_path or not self.findings:
            return 0

        has_unicode = bool(self.unicode_hits) or any(
            f.get("tag") == UNICODE_FINDING_TAG for f in self.findings
        )

        if self.is_office:
            updates: dict[str, str] = {}
            for finding in self.findings:
                tag = finding.get("tag", "")
                if not finding.get("editable") or tag == UNICODE_FINDING_TAG:
                    continue
                short = tag.split(":")[-1].strip()
                if short == "Author":
                    short = "Creator"
                if short in CORE_TAG_MAP or short in APP_TAG_MAP:
                    if short == "RevisionNumber":
                        updates[short] = "1"
                    elif short == "TotalEditTime":
                        updates[short] = "0"
                    else:
                        updates[short] = ""

            if not updates and not has_unicode:
                return 0

            had_temp_before = self.working_temp_path is not None
            working_path = self.ensure_working_copy()
            try:
                OfficeMetadataEngine.update_tags_batch(
                    working_path,
                    updates,
                    strip_invisible_unicode=has_unicode,
                    allow_missing_parts=True,
                )
                self._reload_after_mutation()
            except Exception:
                if not had_temp_before:
                    self.cleanup()
                raise
            return len(updates) + (1 if has_unicode else 0)

        exif_updates: dict[str, str] = {}
        for finding in self.findings:
            tag = finding.get("tag", "")
            if not finding.get("editable") or tag == UNICODE_FINDING_TAG:
                continue
            if tag not in exif_updates and self.is_tag_editable(tag, self.metadata.get(tag, "")):
                exif_updates[tag] = ""

        if not exif_updates:
            return 0

        had_temp_before = self.working_temp_path is not None
        working_path = self.ensure_working_copy()
        try:
            self.exiftool_client.write_tags(working_path, exif_updates)
            self._reload_after_mutation()
        except Exception:
            if not had_temp_before:
                self.cleanup()
            raise
        return len(exif_updates)

    def apply_wipe_all(self) -> list[str]:
        """Aplica la limpieza total de metadatos (y Unicode/custom.xml en Office) sobre el temporal.

        Returns:
            Lista de advertencias no fatales de ExifTool (vacía en Office).
        """
        if not self.original_path:
            raise ValidationError("Primero selecciona un archivo.")

        had_temp_before = self.working_temp_path is not None
        working_path = self.ensure_working_copy()
        warnings: list[str] = []
        try:
            if self.is_office:
                OfficeMetadataEngine.wipe_all(working_path, strip_invisible_unicode=True)
            else:
                warnings = self.exiftool_client.wipe_all(working_path)
            self._reload_after_mutation()
        except Exception:
            if not had_temp_before:
                self.cleanup()
            raise
        return warnings

    def run_scan(self) -> tuple[list[dict[str, Any]], dict[str, str], dict[str, int]]:
        """Ejecuta el escáner forense sobre la sesión activa (con escaneo Unicode bajo demanda).

        Returns:
            Tupla ``(findings, flagged_tags, unicode_hits)``.
        """
        if not self.active_path:
            raise ValidationError("Primero selecciona un archivo para escanear.")

        unicode_hits: dict[str, int] = {}
        if self.is_office:
            unicode_hits = OfficeMetadataEngine.scan_invisible_unicode(self.active_path)

        findings, flagged_tags = ForensicScanner.analyze(
            self.metadata,
            unicode_hits,
            self.active_path,
        )
        self.unicode_hits = unicode_hits
        self.findings = findings
        self.flagged_tags = flagged_tags
        return findings, flagged_tags, unicode_hits

    def save_copy(self, target_dir: str | None = None) -> str:
        """Exporta la copia temporal de forma atómica hacia ``_procesado[_N]`` y regresa al original.

        Detecta carpetas de solo lectura mediante EAFP elevando
        ``DestinationNotWritableError``. Impide sobrescribir ``original_path``
        comprobando ``os.path.samefile``. Tras guardar con éxito, elimina el
        directorio temporal y recarga los metadatos originales desde
        ``original_path`` reiniciando el estado de cambios pendientes.

        Args:
            target_dir: Carpeta destino alternativa (opcional).

        Returns:
            Ruta absoluta de la copia creada en disco.

        Raises:
            ValidationError: Si no hay cambios pendientes o el destino coincide con el original.
            DestinationNotWritableError: Si la carpeta destino carece de permisos de escritura.
            FileProcessingError: Si falla la escritura o no hay espacio (``ENOSPC``).
        """
        if not self.original_path or not self.has_pending_changes or not self.working_temp_path:
            raise ValidationError("No existen cambios pendientes para guardar.")

        dest_dir = Path(target_dir) if target_dir is not None else Path(self.original_path).parent
        tmp_dest: str | None = None
        try:
            fd, tmp_dest = tempfile.mkstemp(prefix=".metaforge_", dir=str(dest_dir))
            os.close(fd)
        except OSError as exc:
            if isinstance(exc, PermissionError) or exc.errno in (errno.EACCES, errno.EROFS):
                raise DestinationNotWritableError(
                    f"La carpeta '{dest_dir}' es de solo lectura o no permite escritura."
                ) from exc
            if exc.errno == errno.ENOSPC:
                raise FileProcessingError("No hay espacio suficiente en la carpeta de destino.") from exc
            raise FileProcessingError(f"No se pudo preparar la escritura en destino: {exc}") from exc

        try:
            shutil.copy2(self.working_temp_path, tmp_dest)
            final_path = resolve_output_path(self.original_path, str(dest_dir))

            if os.path.abspath(str(final_path)) == os.path.abspath(self.original_path) or (
                final_path.exists() and os.path.samefile(final_path, self.original_path)
            ):
                raise ValidationError(
                    "El destino elegido coincide con el archivo original; el archivo original no puede sobrescribirse."
                )

            os.replace(tmp_dest, final_path)
            tmp_dest = None
        except (ValidationError, DestinationNotWritableError):
            raise
        except OSError as exc:
            if exc.errno == errno.ENOSPC:
                raise FileProcessingError("No hay espacio suficiente en la carpeta de destino.") from exc
            if isinstance(exc, PermissionError) or exc.errno in (errno.EACCES, errno.EROFS):
                raise DestinationNotWritableError(
                    f"Permiso denegado al guardar en '{dest_dir}'."
                ) from exc
            raise FileProcessingError(f"Error al guardar la copia procesada: {exc}") from exc
        finally:
            if tmp_dest and os.path.exists(tmp_dest):
                try:
                    os.remove(tmp_dest)
                except OSError:
                    pass

        saved_path = str(final_path)
        original = self.original_path
        self.cleanup()
        self.metadata = self.exiftool_client.read_metadata(original)
        self.unicode_hits = {}
        self.findings = []
        self.flagged_tags = {}
        self._detect_format_flags(original, self.metadata)
        return saved_path

    def discard_changes(self) -> None:
        """Descarta la copia temporal de trabajo y recarga los metadatos originales."""
        if not self.original_path:
            return
        original = self.original_path
        self.cleanup()
        self.metadata = self.exiftool_client.read_metadata(original)
        self.unicode_hits = {}
        self.findings = []
        self.flagged_tags = {}
        self._detect_format_flags(original, self.metadata)

    def export_json(self, dest_path: str) -> None:
        """Exporta los metadatos actuales de la sesión a un archivo JSON UTF-8.

        Args:
            dest_path: Ruta destino para el archivo JSON.

        Raises:
            ValidationError: Si no hay metadatos cargados.
            FileProcessingError: Si falla la escritura en disco.
        """
        if not self.metadata:
            raise ValidationError("No hay metadatos cargados para exportar.")
        try:
            with open(dest_path, "w", encoding="utf-8") as handle:
                json.dump(self.metadata, handle, indent=4, ensure_ascii=False)
        except OSError as exc:
            raise FileProcessingError(f"No se pudo guardar el archivo JSON: {exc}") from exc

    def cleanup(self) -> None:
        """Elimina el directorio temporal privado de la sesión y reinicia punteros."""
        temp_dir = self._session_temp_dir
        self._session_temp_dir = None
        self.working_temp_path = None
        self.pending_change_count = 0
        self.exiftool_client.work_dir = None
        try:
            atexit.unregister(self.cleanup)
        except Exception:
            pass
        if temp_dir and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)
