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

"""Motor seguro de metadatos OOXML (.docx, .xlsx, .pptx) y saneamiento Unicode."""

from __future__ import annotations

import os
import re
import shutil
import tempfile
import xml.etree.ElementTree as ET
from xml.parsers import expat
import zipfile
import zlib
from datetime import datetime
from typing import Callable

from metaforge import FileProcessingError, ValidationError

OFFICE_EXTENSIONS = {".docx", ".xlsx", ".pptx"}

# Mapeo entre etiquetas mostradas por ExifTool y nodos XML reales de Office.
CORE_TAG_MAP = {
    "Title": ("dc", "title"),
    "Subject": ("dc", "subject"),
    "Creator": ("dc", "creator"),
    "Author": ("dc", "creator"),
    "Description": ("dc", "description"),
    "Keywords": ("cp", "keywords"),
    "LastModifiedBy": ("cp", "lastModifiedBy"),
    "RevisionNumber": ("cp", "revision"),
    "LastPrinted": ("cp", "lastPrinted"),
    "CreateDate": ("dcterms", "created"),
    "ModifyDate": ("dcterms", "modified"),
    "Category": ("cp", "category"),
}

APP_TAG_MAP = {
    "Application": "Application",
    "AppVersion": "AppVersion",
    "Company": "Company",
    "Manager": "Manager",
    "Template": "Template",
    "TotalEditTime": "TotalTime",
}

DATE_XML_LOCAL_NAMES = {"created", "modified", "lastPrinted"}

CORE_NS = {
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "dcmitype": "http://purl.org/dc/dcmitype/",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
}

APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

# Vaciar docProps/custom.xml a <Properties/> preserva las relaciones OPC sin romper Office.
EMPTY_CUSTOM_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>'
).encode("utf-8")

# Catálogo cerrado de caracteres Unicode invisibles sospechosos (sin U+202F, U+200D, U+2009, U+00AD ni U+200C).
INVISIBLE_UNICODE_MAP = {
    "\u200b": ("Zero-Width Space (U+200B)", ""),
    "\u2060": ("Word Joiner (U+2060)", ""),
    "\ufeff": ("Byte Order Mark interno (U+FEFF)", ""),
}

_ENTITY_CODE_TO_CHAR = {ord(ch): ch for ch in INVISIBLE_UNICODE_MAP}
_TEXT_NODE_RE = re.compile(r"(<(?:w:|a:)?t(?:\s[^>]*)?>)(.*?)(</(?:w:|a:)?t>)", re.DOTALL)
_NUMERIC_ENTITY_RE = re.compile(r"&#(?:[xX]0*([0-9a-fA-F]+)|0*([0-9]+));")

_EXIF_DATE_RE = re.compile(
    r"^(\d{4}):(\d{2}):(\d{2})[ T](\d{2}:\d{2}:\d{2}(?:\.\d+)?)(Z|z|[+-]\d{2}:\d{2})?$"
)
_ISO_DATE_RE = re.compile(
    r"^(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2}:\d{2}(?:\.\d+)?)(Z|z|[+-]\d{2}:\d{2})?$"
)
_HMS_RE = re.compile(r"^(\d+):(\d{1,2}):(\d{1,2})$")
_DURATION_TOKEN_RE = re.compile(
    r"(\d+)\s*(days?|días?|dias?|hours?|hrs?|horas?|minutes?|mins?|minutos?|seconds?|secs?|segundos?)\b",
    re.IGNORECASE,
)

MAX_XML_BYTES = 20 * 1024 * 1024
MAX_TOTAL_TEXT_BYTES = 200 * 1024 * 1024
MAX_ZIP_ENTRIES = 10_000
FIXED_ZIP_DATE_TIME = (1980, 1, 1, 0, 0, 0)
_READ_CHUNK_BYTES = 64 * 1024


def normalize_w3cdtf(value: str) -> str:
    """Valida y normaliza una fecha a formato ISO 8601 / W3CDTF compatible con Office.

    Acepta cadena vacía (para eliminar el nodo XML de fecha), formato ISO 8601
    (``YYYY-MM-DDTHH:MM:SS[.fff][Z|±HH:MM]``) y el formato devuelto por
    ExifTool (``YYYY:MM:DD HH:MM:SS[.fff][Z|±HH:MM]``). Compatible con
    Python 3.10+ normalizando ``Z`` a ``+00:00`` antes de validar el calendario.

    Args:
        value: Cadena de fecha ingresada por el usuario o ExifTool.

    Returns:
        Cadena ``W3CDTF`` normalizada, o ``""`` si el valor estaba vacío.

    Raises:
        ValidationError: Si la sintaxis o la fecha de calendario son inválidas.
    """
    text = value.strip()
    if not text:
        return ""

    exif_match = _EXIF_DATE_RE.fullmatch(text)
    if exif_match:
        year, month, day, time_part, tz_part = exif_match.groups()
        iso_candidate = f"{year}-{month}-{day}T{time_part}{tz_part or 'Z'}"
    else:
        iso_match = _ISO_DATE_RE.fullmatch(text)
        if not iso_match:
            raise ValidationError(
                f"Fecha inválida para Office: '{value}'. "
                "Usa formato ISO 8601 (ej. 2026-11-27T14:30:00Z)."
            )
        date_part, time_part, tz_part = iso_match.groups()
        iso_candidate = f"{date_part}T{time_part}{tz_part or 'Z'}"

    # En Python 3.10 datetime.fromisoformat no acepta el sufijo literal 'Z'.
    parse_target = (
        iso_candidate[:-1] + "+00:00"
        if iso_candidate.endswith(("Z", "z"))
        else iso_candidate
    )
    try:
        datetime.fromisoformat(parse_target)
    except ValueError as exc:
        raise ValidationError(
            f"La fecha '{value}' no corresponde a una fecha de calendario válida."
        ) from exc

    if iso_candidate.endswith("z"):
        iso_candidate = iso_candidate[:-1] + "Z"
    return iso_candidate


def parse_edit_time_minutes(value: str) -> str:
    """Convierte una duración de edición en minutos enteros para ``TotalTime``.

    Soporta cadena vacía (``"0"``), enteros puros (``"45"``), formato
    ``H:MM:SS`` y expresiones con días, horas y minutos (``"1 day 2 hours"``,
    ``"1 hour 20 minutes"``).

    Args:
        value: Texto con el tiempo total de edición.

    Returns:
        Cantidad entera de minutos en formato texto.

    Raises:
        ValidationError: Si el texto no representa una duración válida.
    """
    text = value.strip()
    if not text:
        return "0"

    if text.isdigit():
        return str(int(text))

    hms_match = _HMS_RE.fullmatch(text)
    if hms_match:
        hours, mins, _secs = (int(part) for part in hms_match.groups())
        return str(hours * 60 + mins)

    matches = list(_DURATION_TOKEN_RE.finditer(text))
    if not matches:
        raise ValidationError(
            f"Valor inválido para TotalEditTime: '{value}'. "
            "Usa minutos enteros (ej. 45) o formato como '1 hour 20 minutes'."
        )

    total_minutes = 0
    for match in matches:
        amount = int(match.group(1))
        unit = match.group(2).lower()
        if unit.startswith(("day", "día", "dia")):
            total_minutes += amount * 1440
        elif unit.startswith(("hour", "hr", "hora")):
            total_minutes += amount * 60
        elif unit.startswith(("min",)):
            total_minutes += amount
    return str(total_minutes)


def validate_revision_number(value: str) -> str:
    """Valida que ``RevisionNumber`` sea un entero mayor o igual a 1.

    Args:
        value: Valor propuesto para ``cp:revision``.

    Returns:
        Entero positivo normalizado como cadena.

    Raises:
        ValidationError: Si el valor está vacío, no es numérico o es menor que 1.
    """
    text = value.strip()
    if not text.isdigit() or int(text) < 1:
        raise ValidationError(
            f"RevisionNumber inválido: '{value}'. Debe ser un número entero mayor o igual a 1."
        )
    return str(int(text))


def _is_unicode_target_xml(ext: str, member_name: str) -> bool:
    """Determina si una entrada ZIP contiene texto editable del documento Office.

    Excluye intencionadamente ``comments*.xml`` para mantener coherencia con el
    aviso al usuario de que los comentarios no se modifican.
    """
    if not member_name.endswith(".xml"):
        return False
    if ext == ".docx":
        return member_name in (
            "word/document.xml",
            "word/footnotes.xml",
            "word/endnotes.xml",
        ) or member_name.startswith(("word/header", "word/footer"))
    if ext == ".xlsx":
        return member_name == "xl/sharedStrings.xml" or member_name.startswith("xl/worksheets/sheet")
    if ext == ".pptx":
        return member_name.startswith(("ppt/slides/slide", "ppt/notesSlides/notesSlide"))
    return False


def _read_limited(zin: zipfile.ZipFile, item: zipfile.ZipInfo, max_bytes: int) -> bytes:
    """Lee una entrada ZIP contando bytes reales descomprimidos para frenar bombas ZIP."""
    if item.flag_bits & 0x1:
        raise FileProcessingError(
            f"El archivo Office está cifrado o protegido con contraseña ({item.filename})."
        )
    chunks: list[bytes] = []
    total = 0
    try:
        with zin.open(item, "r") as stream:
            while True:
                chunk = stream.read(_READ_CHUNK_BYTES)
                if not chunk:
                    break
                total += len(chunk)
                if total > max_bytes:
                    raise FileProcessingError(
                        f"La entrada XML '{item.filename}' excede el límite seguro de tamaño ({max_bytes} bytes)."
                    )
                chunks.append(chunk)
    except (zipfile.BadZipFile, RuntimeError, zlib.error, EOFError, NotImplementedError) as exc:
        raise FileProcessingError(f"No se pudo descomprimir '{item.filename}': {exc}") from exc
    return b"".join(chunks)


def _parse_safe_xml(xml_bytes: bytes) -> ET.Element:
    """Parsea XML bloqueando declaraciones DOCTYPE y ENTITY mediante Expat (UTF-8 y UTF-16)."""
    def _reject_xml_entity(*_args: object) -> None:
        raise FileProcessingError(
            "El archivo Office contiene declaraciones XML DOCTYPE/ENTITY no permitidas."
        )

    checker = expat.ParserCreate()
    checker.StartDoctypeDeclHandler = _reject_xml_entity
    checker.EntityDeclHandler = _reject_xml_entity
    try:
        checker.Parse(xml_bytes, True)
        return ET.fromstring(xml_bytes)
    except FileProcessingError:
        raise
    except (expat.ExpatError, ET.ParseError) as exc:
        raise FileProcessingError(f"XML interno de Office mal formado: {exc}") from exc


def _decode_xml_text(xml_bytes: bytes) -> tuple[str, bool]:
    try:
        text = xml_bytes.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise FileProcessingError(f"Codificación UTF-8 inválida en XML de Office: {exc}") from exc
    has_bom = text.startswith("\ufeff")
    return (text[1:] if has_bom else text), has_bom


def _count_or_strip_in_segment(segment: str, counter: dict[str, int] | None = None) -> str:
    """Cuenta o elimina caracteres invisibles y sus entidades numéricas dentro de un nodo de texto."""
    def _replace_entity(match: re.Match[str]) -> str:
        hex_part, dec_part = match.groups()
        code = int(hex_part, 16) if hex_part is not None else int(dec_part, 10)
        ch = _ENTITY_CODE_TO_CHAR.get(code)
        if ch is None:
            return match.group(0)
        label, replacement = INVISIBLE_UNICODE_MAP[ch]
        if counter is not None:
            counter[label] = counter.get(label, 0) + 1
            return match.group(0)
        return replacement

    updated = _NUMERIC_ENTITY_RE.sub(_replace_entity, segment)
    for ch, (label, replacement) in INVISIBLE_UNICODE_MAP.items():
        hits = updated.count(ch)
        if hits > 0:
            if counter is not None:
                counter[label] = counter.get(label, 0) + hits
            else:
                updated = updated.replace(ch, replacement)
    return updated


class OfficeMetadataEngine:
    """Motor nativo para editar, purgar metadatos y sanear Unicode en Office OOXML."""

    @staticmethod
    def _register_core_ns() -> None:
        for prefix, uri in CORE_NS.items():
            ET.register_namespace(prefix, uri)

    @staticmethod
    def _register_app_ns() -> None:
        ET.register_namespace("", APP_NS)
        ET.register_namespace("vt", VT_NS)

    @classmethod
    def _prepare_updates(cls, updates: dict[str, str]) -> tuple[dict[str, str], dict[str, str]]:
        core_updates: dict[str, str] = {}
        app_updates: dict[str, str] = {}
        creator_vals: dict[str, str] = {}

        for raw_tag, raw_val in updates.items():
            clean_tag = raw_tag.split(":")[-1].strip()
            val_str = "" if raw_val is None else str(raw_val)

            if clean_tag in CORE_TAG_MAP:
                if clean_tag in ("Creator", "Author"):
                    creator_vals[clean_tag] = val_str
                _, local_name = CORE_TAG_MAP[clean_tag]
                if local_name in DATE_XML_LOCAL_NAMES:
                    val_str = normalize_w3cdtf(val_str)
                elif local_name == "revision":
                    val_str = validate_revision_number(val_str)
                core_updates[clean_tag] = val_str
            elif clean_tag in APP_TAG_MAP:
                if APP_TAG_MAP[clean_tag] == "TotalTime":
                    val_str = parse_edit_time_minutes(val_str)
                app_updates[clean_tag] = val_str
            else:
                valid = ", ".join(list(CORE_TAG_MAP.keys()) + list(APP_TAG_MAP.keys()))
                raise ValidationError(
                    f"La etiqueta '{clean_tag}' no es editable en archivos de Office.\n\n"
                    f"Etiquetas soportadas:\n{valid}"
                )

        if "Creator" in creator_vals and "Author" in creator_vals:
            if creator_vals["Creator"] != creator_vals["Author"]:
                raise ValidationError(
                    "Conflicto de alias: 'Creator' y 'Author' apuntan al mismo campo dc:creator con valores distintos."
                )

        return core_updates, app_updates

    @classmethod
    def scan_invisible_unicode(cls, file_path: str) -> dict[str, int]:
        """Escanea caracteres Unicode invisibles dentro de nodos de texto Office.

        Analiza únicamente el contenido interior de ``<w:t>``, ``<t>`` y ``<a:t>``
        en ``.docx``, ``.xlsx`` y ``.pptx``, cubriendo caracteres directos y
        entidades numéricas e ignorando el BOM inicial en posición 0.

        Args:
            file_path: Ruta al documento Office.

        Returns:
            Diccionario ``{etiqueta_caracter: cantidad}``.

        Raises:
            FileProcessingError: Si el ZIP está corrupto, cifrado o supera los límites.
        """
        _, ext = os.path.splitext(file_path.lower())
        if ext not in OFFICE_EXTENSIONS:
            return {}

        counts: dict[str, int] = {}
        total_text_bytes = 0

        try:
            with zipfile.ZipFile(file_path, "r") as zin:
                entries = zin.infolist()
                if len(entries) > MAX_ZIP_ENTRIES:
                    raise FileProcessingError(
                        f"El archivo Office excede el máximo de entradas permitidas ({MAX_ZIP_ENTRIES})."
                    )
                for item in entries:
                    if not _is_unicode_target_xml(ext, item.filename):
                        continue
                    data = _read_limited(zin, item, MAX_XML_BYTES)
                    total_text_bytes += len(data)
                    if total_text_bytes > MAX_TOTAL_TEXT_BYTES:
                        raise FileProcessingError(
                            "El contenido XML descomprimido del documento supera el límite seguro de inspección."
                        )
                    body, _ = _decode_xml_text(data)
                    for match in _TEXT_NODE_RE.finditer(body):
                        _count_or_strip_in_segment(match.group(2), counter=counts)
        except FileProcessingError:
            raise
        except (zipfile.BadZipFile, RuntimeError, zlib.error, EOFError, OSError) as exc:
            raise FileProcessingError(f"No se pudo leer el documento Office: {exc}") from exc

        return counts

    @staticmethod
    def _strip_invisible_chars_from_xml(xml_bytes: bytes) -> bytes:
        """Limpia caracteres invisibles solo dentro de <w:t>/<t>/<a:t> preservando prefijos OOXML."""
        body, has_bom = _decode_xml_text(xml_bytes)

        def _clean_match(match: re.Match[str]) -> str:
            return f"{match.group(1)}{_count_or_strip_in_segment(match.group(2))}{match.group(3)}"

        cleaned_body = _TEXT_NODE_RE.sub(_clean_match, body)
        final_text = ("\ufeff" + cleaned_body) if has_bom else cleaned_body
        return final_text.encode("utf-8")

    @classmethod
    def _modify_core_xml(cls, xml_bytes: bytes, updates_dict: dict[str, str]) -> bytes:
        cls._register_core_ns()
        root = _parse_safe_xml(xml_bytes)

        for clean_tag, new_value in updates_dict.items():
            ns_prefix, local_name = CORE_TAG_MAP[clean_tag]
            full_tag = f"{{{CORE_NS[ns_prefix]}}}{local_name}"
            elem = root.find(full_tag)

            # En fechas XML (created, modified, lastPrinted), un valor vacío corrompe Office -> eliminar nodo.
            if local_name in DATE_XML_LOCAL_NAMES and not new_value:
                if elem is not None:
                    root.remove(elem)
                continue

            if elem is None:
                elem = ET.SubElement(root, full_tag)
            if local_name in ("created", "modified"):
                elem.set(f"{{{CORE_NS['xsi']}}}type", "dcterms:W3CDTF")
            elem.text = new_value

        return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    @classmethod
    def _modify_app_xml(cls, xml_bytes: bytes, updates_dict: dict[str, str]) -> bytes:
        cls._register_app_ns()
        root = _parse_safe_xml(xml_bytes)

        for clean_tag, new_value in updates_dict.items():
            xml_name = APP_TAG_MAP[clean_tag]
            full_tag = f"{{{APP_NS}}}{xml_name}"
            elem = root.find(full_tag)
            if elem is None:
                elem = ET.SubElement(root, full_tag)
            elem.text = new_value

        return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    @classmethod
    def update_tags_batch(
        cls,
        file_path: str,
        updates: dict[str, str],
        strip_invisible_unicode: bool = False,
        clear_custom_xml: bool = False,
        allow_missing_parts: bool = False,
    ) -> None:
        """Aplica cambios de metadatos, vaciado de ``custom.xml`` y saneado Unicode en un ciclo ZIP.

        Args:
            file_path: Ruta al documento Office de trabajo.
            updates: Diccionario de etiquetas y nuevos valores.
            strip_invisible_unicode: Si es ``True``, limpia caracteres invisibles en nodos de texto.
            clear_custom_xml: Si es ``True``, vacía ``docProps/custom.xml`` dejando ``<Properties/>``.
            allow_missing_parts: Si es ``False``, eleva error si falta ``core.xml`` o ``app.xml``
                al intentar editar propiedades pertenecientes a esa parte.

        Raises:
            ValidationError: Si alguna etiqueta o valor no supera la validación.
            FileProcessingError: Si el archivo es inválido, cifrado o faltan partes requeridas.
        """
        _, ext = os.path.splitext(file_path.lower())
        core_updates, app_updates = cls._prepare_updates(updates)

        parent_dir = os.path.dirname(os.path.abspath(file_path)) or "."
        tmp_fd, tmp_zip_path = tempfile.mkstemp(prefix=".mf_ooxml_", suffix=".zip", dir=parent_dir)
        os.close(tmp_fd)

        found_core = False
        found_app = False
        total_text_bytes = 0

        try:
            with zipfile.ZipFile(file_path, "r") as zin, zipfile.ZipFile(tmp_zip_path, "w") as zout:
                entries = zin.infolist()
                if len(entries) > MAX_ZIP_ENTRIES:
                    raise FileProcessingError(
                        f"El archivo Office excede el máximo de entradas permitidas ({MAX_ZIP_ENTRIES})."
                    )

                for item in entries:
                    clean_info = zipfile.ZipInfo(
                        filename=item.filename,
                        date_time=FIXED_ZIP_DATE_TIME,
                    )
                    clean_info.compress_type = item.compress_type

                    if item.filename == "docProps/core.xml" and core_updates:
                        found_core = True
                        raw_xml = _read_limited(zin, item, MAX_XML_BYTES)
                        zout.writestr(clean_info, cls._modify_core_xml(raw_xml, core_updates))
                    elif item.filename == "docProps/app.xml" and app_updates:
                        found_app = True
                        raw_xml = _read_limited(zin, item, MAX_XML_BYTES)
                        zout.writestr(clean_info, cls._modify_app_xml(raw_xml, app_updates))
                    elif item.filename == "docProps/custom.xml" and clear_custom_xml:
                        zout.writestr(clean_info, EMPTY_CUSTOM_XML)
                    elif strip_invisible_unicode and _is_unicode_target_xml(ext, item.filename):
                        raw_xml = _read_limited(zin, item, MAX_XML_BYTES)
                        total_text_bytes += len(raw_xml)
                        if total_text_bytes > MAX_TOTAL_TEXT_BYTES:
                            raise FileProcessingError(
                                "El contenido XML descomprimido supera el límite seguro de escritura."
                            )
                        zout.writestr(clean_info, cls._strip_invisible_chars_from_xml(raw_xml))
                    else:
                        if item.flag_bits & 0x1:
                            raise FileProcessingError(
                                f"El archivo Office está cifrado o protegido con contraseña ({item.filename})."
                            )
                        # Copiar partes binarias no modificadas en streaming para no saturar RAM.
                        with zin.open(item, "r") as src_stream, zout.open(clean_info, "w") as dst_stream:
                            shutil.copyfileobj(src_stream, dst_stream, length=_READ_CHUNK_BYTES)

            if not allow_missing_parts:
                if core_updates and not found_core:
                    raise FileProcessingError(
                        "El documento Office no contiene 'docProps/core.xml' para editar esa etiqueta."
                    )
                if app_updates and not found_app:
                    raise FileProcessingError(
                        "El documento Office no contiene 'docProps/app.xml' para editar esa etiqueta."
                    )

            os.replace(tmp_zip_path, file_path)
        except (ValidationError, FileProcessingError):
            if os.path.exists(tmp_zip_path):
                try:
                    os.remove(tmp_zip_path)
                except OSError:
                    pass
            raise
        except (zipfile.BadZipFile, RuntimeError, zlib.error, EOFError, NotImplementedError, OSError) as exc:
            if os.path.exists(tmp_zip_path):
                try:
                    os.remove(tmp_zip_path)
                except OSError:
                    pass
            raise FileProcessingError(f"Error al procesar el documento Office: {exc}") from exc

    @classmethod
    def update_tag(cls, file_path: str, raw_tag: str, new_value: str) -> None:
        """Actualiza una única etiqueta de Office delegando en ``update_tags_batch``.

        Args:
            file_path: Ruta al documento Office de trabajo.
            raw_tag: Etiqueta con o sin prefijo de grupo.
            new_value: Nuevo valor a escribir.
        """
        cls.update_tags_batch(file_path, {raw_tag: new_value})

    @classmethod
    def wipe_all(cls, file_path: str, strip_invisible_unicode: bool = True) -> None:
        """Purga metadatos estándar, fechas, ``custom.xml`` y Unicode invisible de un Office.

        Args:
            file_path: Ruta al documento Office de trabajo.
            strip_invisible_unicode: Si es ``True``, limpia caracteres invisibles en nodos de texto.
        """
        updates = {
            "Title": "",
            "Subject": "",
            "Creator": "",
            "Description": "",
            "Keywords": "",
            "LastModifiedBy": "",
            "RevisionNumber": "1",
            "LastPrinted": "",
            "CreateDate": "",
            "ModifyDate": "",
            "Category": "",
            "Application": "",
            "AppVersion": "",
            "Company": "",
            "Manager": "",
            "Template": "",
            "TotalEditTime": "0",
        }
        cls.update_tags_batch(
            file_path,
            updates,
            strip_invisible_unicode=strip_invisible_unicode,
            clear_custom_xml=True,
            allow_missing_parts=True,
        )
