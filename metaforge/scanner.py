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

"""Motor puro de análisis forense/IA y política única de editabilidad de etiquetas."""

from __future__ import annotations

import os
import re
from datetime import datetime, timezone
from typing import Any

from metaforge import ValidationError
from metaforge.office_engine import (
    APP_TAG_MAP,
    CORE_TAG_MAP,
    OFFICE_EXTENSIONS,
    parse_edit_time_minutes,
)

UNICODE_FINDING_TAG = "[Contenido Office]"

READONLY_GROUPS = {"ExifTool", "File", "System", "ZIP", "Composite", "JUMBF"}
OFFICE_EDITABLE_GROUPS = {"XML", "XMP", ""}
RESERVED_UNGROUPED_TAGS = {"sourcefile", "filename", "directory"}

AI_GROUP_NAMES = {"c2pa", "jumbf"}
AI_TAG_EXACT_NAMES = {
    "c2pa",
    "jumbf",
    "parameters",
    "workflow",
    "prompt",
    "negativeprompt",
    "digitalsourcetype",
    "sd-metadata",
    "synthid",
    "contentcredentials",
    "aigenerated",
    "ai_generated",
}

# Firmas explícitas de IA generativa y credenciales C2PA (sin LibreOffice, Google Docs, Canva ni "imagen").
AI_VALUE_SIGNATURES: list[tuple[re.Pattern[str], str]] = [
    (re.compile(r"\bopenai\b", re.IGNORECASE), "OpenAI"),
    (re.compile(r"\bchatgpt\b", re.IGNORECASE), "ChatGPT"),
    (re.compile(r"\bgpt-[345]\b", re.IGNORECASE), "GPT-3/4/5"),
    (re.compile(r"\bdall-?e\b", re.IGNORECASE), "DALL-E"),
    (re.compile(r"\bmidjourney\b", re.IGNORECASE), "Midjourney"),
    (re.compile(r"\bstable\s*diffusion\b", re.IGNORECASE), "Stable Diffusion"),
    (re.compile(r"\bcomfyui\b", re.IGNORECASE), "ComfyUI"),
    (re.compile(r"\bautomatic1111\b", re.IGNORECASE), "Automatic1111"),
    (re.compile(r"\bfirefly\b", re.IGNORECASE), "Adobe Firefly"),
    (re.compile(r"\bgemini\b", re.IGNORECASE), "Gemini"),
    (re.compile(r"\bclaude\b", re.IGNORECASE), "Claude"),
    (re.compile(r"\banthropic\b", re.IGNORECASE), "Anthropic"),
    (re.compile(r"\bperplexity\b", re.IGNORECASE), "Perplexity"),
    (re.compile(r"\bcopilot\b", re.IGNORECASE), "Copilot"),
    (re.compile(r"\bsynthid\b", re.IGNORECASE), "SynthID"),
    (re.compile(r"\bc2pa\b", re.IGNORECASE), "C2PA"),
    (re.compile(r"\btrainedalgorithmicmedia\b", re.IGNORECASE), "TrainedAlgorithmicMedia"),
    (
        re.compile(r"\bcompositewithtrainedalgorithmicmedia\b", re.IGNORECASE),
        "CompositeWithTrainedAlgorithmicMedia",
    ),
    (re.compile(r"\bsteps:\s*\d+,\s*sampler:", re.IGNORECASE), "Parámetros de difusión (Steps/Sampler)"),
    (re.compile(r"\bcfg\s*scale:", re.IGNORECASE), "Parámetros de difusión (CFG Scale)"),
]

AUTOMATED_EXPORT_SIGNATURES: list[tuple[re.Pattern[str], str]] = [
    (re.compile(r"\bpython-docx\b", re.IGNORECASE), "python-docx"),
    (re.compile(r"\bpandoc\b", re.IGNORECASE), "Pandoc"),
    (re.compile(r"\baspose\b", re.IGNORECASE), "Aspose"),
    (re.compile(r"\breportlab\b", re.IGNORECASE), "ReportLab"),
    (re.compile(r"\bweasyprint\b", re.IGNORECASE), "WeasyPrint"),
]

OPSEC_TAGS = {
    "creator",
    "author",
    "lastmodifiedby",
    "artist",
    "ownername",
    "description",
    "company",
    "manager",
    "lastprinted",
    "serialnumber",
    "internalserialnumber",
    "lensserialnumber",
    "hostcomputer",
    "usercomment",
    "gpslatitude",
    "gpslongitude",
    "gpsposition",
    "gpsaltitude",
    "make",
    "model",
}

_DATE_GROUPS = ("XML", "PDF", "EXIF", "XMP")
_EXIF_DT_RE = re.compile(
    r"^(\d{4}):(\d{2}):(\d{2})(?:[ T](\d{2}:\d{2}:\d{2}(?:\.\d+)?)(Z|z|[+-]\d{2}:\d{2})?)?"
)
_ISO_DT_RE = re.compile(
    r"^(\d{4}-\d{2}-\d{2})(?:[ T](\d{2}:\d{2}:\d{2}(?:\.\d+)?)(Z|z|[+-]\d{2}:\d{2})?)?"
)
_CAMEL_SPLIT_RE = re.compile(r"[A-Z]+(?=[A-Z][a-z]|\b)|[A-Z]?[a-z]+|\d+")


def _split_tag(full_tag: str) -> tuple[str, str]:
    clean = full_tag.strip()
    if ":" in clean:
        group, short = clean.split(":", 1)
        return group.strip(), short.strip()
    return "", clean


def _is_office_path(file_path: str) -> bool:
    if not file_path:
        return False
    _, ext = os.path.splitext(file_path.lower())
    return ext in OFFICE_EXTENSIONS


def is_tag_editable(file_path: str, full_tag: str, raw_value: Any = "") -> bool:
    """Determina de forma unificada si una etiqueta es editable para el archivo dado.

    Comprueba primero grupos de solo lectura (incluyendo ``ZIP`` y ``JUMBF``),
    etiquetas reservadas sin grupo, valores binarios y estructuras compuestas
    (``list``/``dict``). Para archivos Office restringe la edición a los grupos
    ``XML``/``XMP`` y a las claves soportadas por ``OfficeMetadataEngine``.

    Args:
        file_path: Ruta del archivo analizado.
        full_tag: Etiqueta con o sin prefijo de grupo.
        raw_value: Valor crudo obtenido de ExifTool (opcional).

    Returns:
        ``True`` si el campo puede editarse; ``False`` si es solo lectura.
    """
    if not full_tag or not full_tag.strip():
        return False

    group, short_tag = _split_tag(full_tag)
    if group in READONLY_GROUPS:
        return False
    if not group and short_tag.lower() in RESERVED_UNGROUPED_TAGS:
        return False
    if isinstance(raw_value, (list, dict)):
        return False
    if isinstance(raw_value, str) and raw_value.startswith("(Binary data"):
        return False

    if _is_office_path(file_path):
        if group not in OFFICE_EDITABLE_GROUPS:
            return False
        return short_tag in CORE_TAG_MAP or short_tag in APP_TAG_MAP

    return True


def _matches_ai_tag(group: str, short_tag: str) -> bool:
    group_lower = group.lower()
    short_lower = short_tag.lower()
    if group_lower in AI_GROUP_NAMES or short_lower in AI_TAG_EXACT_NAMES:
        return True

    # Tokenización camelCase/snake_case: evita falsos positivos por subcadena como Thumbnail o Detail.
    tokens = [tok.lower() for tok in _CAMEL_SPLIT_RE.findall(short_tag)]
    if not tokens:
        return False
    joined_underscore = "_".join(tokens)
    if joined_underscore in AI_TAG_EXACT_NAMES:
        return True
    if "ai" in tokens and any(t in tokens for t in ("generated", "generator", "model", "prompt", "metadata")):
        return True
    return False


def _parse_datetime_utc(raw_value: Any) -> datetime | None:
    if raw_value is None:
        return None
    text = str(raw_value).strip()
    if not text:
        return None

    m_exif = _EXIF_DT_RE.match(text)
    if m_exif:
        year, month, day, time_part, tz_part = m_exif.groups()
        iso_str = f"{year}-{month}-{day}T{time_part or '00:00:00'}{tz_part or '+00:00'}"
    else:
        m_iso = _ISO_DT_RE.match(text)
        if not m_iso:
            return None
        date_part, time_part, tz_part = m_iso.groups()
        iso_str = f"{date_part}T{time_part or '00:00:00'}{tz_part or '+00:00'}"

    if iso_str.endswith(("Z", "z")):
        iso_str = iso_str[:-1] + "+00:00"
    try:
        dt = datetime.fromisoformat(iso_str)
    except ValueError:
        return None
    if dt.tzinfo is None:
        return dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)


def _normalize_person_value(raw_value: Any) -> str:
    if raw_value is None:
        return ""
    if isinstance(raw_value, list):
        return ", ".join(str(item).strip() for item in raw_value if str(item).strip())
    return str(raw_value).strip()


class ForensicScanner:
    """Analiza metadatos en memoria para detectar firmas IA/C2PA, anomalías y OPSEC."""

    @staticmethod
    def _extract_int(value: Any) -> int | None:
        if value is None:
            return None
        match = re.search(r"\d+", str(value))
        return int(match.group()) if match else None

    @classmethod
    def analyze(
        cls,
        metadata: dict[str, Any],
        unicode_hits: dict[str, int],
        file_path: str,
    ) -> tuple[list[dict[str, Any]], dict[str, str]]:
        """Ejecuta el análisis forense sobre los metadatos y hallazgos Unicode ya leídos.

        No accede al disco. Evalúa reglas de IA/C2PA, exportadores automatizados,
        anomalías temporales/estructurales y fugas OPSEC de forma independiente.

        Args:
            metadata: Diccionario ``{Grupo:Etiqueta: valor}`` devuelto por ExifTool.
            unicode_hits: Conteos de caracteres invisibles obtenidos de ``OfficeMetadataEngine``.
            file_path: Ruta lógica del archivo para calcular editabilidad.

        Returns:
            Tupla ``(findings, flagged_tags)`` donde cada hallazgo contiene las
            claves ``severity``, ``tag``, ``value``, ``reason`` y ``editable``.
        """
        findings: list[dict[str, Any]] = []
        flagged_tags: dict[str, str] = {}
        priority = {"ai_high": 3, "anomaly_med": 2, "opsec_low": 1}

        def add_finding(
            severity: str,
            tag_key: str,
            value: Any,
            reason: str,
            editable: bool | None = None,
        ) -> None:
            is_edit = (
                is_tag_editable(file_path, tag_key, value)
                if editable is None
                else editable
            )
            findings.append(
                {
                    "severity": severity,
                    "tag": tag_key,
                    "value": str(value),
                    "reason": reason,
                    "editable": is_edit,
                }
            )
            if tag_key and (
                tag_key not in flagged_tags
                or priority[severity] > priority[flagged_tags[tag_key]]
            ):
                flagged_tags[tag_key] = severity

        # 1. Análisis etiqueta por etiqueta (evaluación independiente de IA, exportadores y OPSEC).
        for full_key, raw_val in metadata.items():
            if full_key == "SourceFile":
                continue
            val_str = _normalize_person_value(raw_val) if isinstance(raw_val, list) else str(raw_val).strip()
            if not val_str:
                continue

            group, short_tag = _split_tag(full_key)
            short_lower = short_tag.lower()
            editable = is_tag_editable(file_path, full_key, raw_val)

            if _matches_ai_tag(group, short_tag):
                add_finding(
                    "ai_high",
                    full_key,
                    val_str,
                    "Etiqueta asociada a generación por IA o credenciales C2PA/JUMBF.",
                    editable,
                )
            else:
                for pattern, label in AI_VALUE_SIGNATURES:
                    if pattern.search(val_str):
                        add_finding(
                            "ai_high",
                            full_key,
                            val_str,
                            f"Firma explícita de herramienta de IA o C2PA detectada ({label}).",
                            editable,
                        )
                        break
                else:
                    for pattern, label in AUTOMATED_EXPORT_SIGNATURES:
                        if pattern.search(val_str):
                            add_finding(
                                "anomaly_med",
                                full_key,
                                val_str,
                                f"Generador o exportador programático detectado ({label}).",
                                editable,
                            )
                            break

            if short_lower in OPSEC_TAGS or "gps" in short_lower:
                add_finding(
                    "opsec_low",
                    full_key,
                    val_str,
                    "Metadato de identidad, ubicación, dispositivo o trazabilidad personal (OPSEC).",
                    editable,
                )

        # 2. Heurísticas estructurales de documentos Office.
        words = cls._extract_int(metadata.get("XML:Words"))
        revisions = cls._extract_int(metadata.get("XML:RevisionNumber"))
        edit_raw = metadata.get("XML:TotalEditTime")
        edit_mins: int | None = None
        if edit_raw is not None and str(edit_raw).strip():
            try:
                edit_mins = int(parse_edit_time_minutes(str(edit_raw)))
            except ValidationError:
                edit_mins = None

        if words is not None and words >= 250:
            if edit_mins is not None and edit_mins <= 2:
                add_finding(
                    "anomaly_med",
                    "XML:TotalEditTime",
                    edit_raw,
                    f"Anomalía: {words} palabras con solo {edit_mins} min de edición (patrón común en archivos exportados o copiados en bloque).",
                )
            if revisions is not None and revisions <= 1:
                add_finding(
                    "anomaly_med",
                    "XML:RevisionNumber",
                    metadata.get("XML:RevisionNumber"),
                    f"Anomalía: documento extenso ({words} palabras) con una sola revisión registrada.",
                )

        # 3. Consistencia temporal comparando pares homogéneos dentro del mismo grupo.
        for grp in _DATE_GROUPS:
            create_key = f"{grp}:CreateDate"
            modify_key = f"{grp}:ModifyDate"
            if create_key not in metadata or modify_key not in metadata:
                continue
            dt_create = _parse_datetime_utc(metadata[create_key])
            dt_mod = _parse_datetime_utc(metadata[modify_key])
            if dt_create is None or dt_mod is None:
                continue
            if dt_mod < dt_create:
                add_finding(
                    "anomaly_med",
                    modify_key,
                    metadata[modify_key],
                    f"Inconsistencia temporal: ModifyDate ({metadata[modify_key]}) es anterior a CreateDate ({metadata[create_key]}).",
                )
            elif abs(dt_mod.year - dt_create.year) >= 2:
                add_finding(
                    "anomaly_med",
                    create_key,
                    metadata[create_key],
                    f"Inconsistencia temporal: creado en {dt_create.year} pero modificado en {dt_mod.year} (plantilla antigua o fecha residual).",
                )

        # 4. HeadingPairs bilingüe (soporta list o str).
        heading_pairs = metadata.get("XML:HeadingPairs")
        if heading_pairs is not None:
            joined = (
                " ".join(str(x) for x in heading_pairs)
                if isinstance(heading_pairs, list)
                else str(heading_pairs)
            )
            if "Título" in joined and "Title" in joined:
                add_finding(
                    "anomaly_med",
                    "XML:HeadingPairs",
                    joined,
                    "Mezcla de idiomas internos ('Título' y 'Title'): frecuente al generar un documento en inglés y editarlo en Office en español.",
                )

        # 5. Discrepancia de autoría unificada.
        creator_key = None
        creator_val = ""
        for candidate in ("XML:Creator", "XMP:Creator", "PDF:Author"):
            normalized = _normalize_person_value(metadata.get(candidate))
            if normalized:
                creator_key = candidate
                creator_val = normalized
                break
        last_mod = _normalize_person_value(metadata.get("XML:LastModifiedBy"))
        if creator_key and creator_val and last_mod and creator_val != last_mod:
            add_finding(
                "anomaly_med",
                creator_key,
                f"{creator_val} ≠ {last_mod}",
                "El autor original difiere del último editor registrado (LastModifiedBy).",
            )

        # 6. Hallazgos de caracteres Unicode invisibles (recibidos en memoria).
        for label, count in unicode_hits.items():
            if count > 0:
                add_finding(
                    "ai_high",
                    UNICODE_FINDING_TAG,
                    f"{count} ocurrencias de {label}",
                    "Carácter Unicode invisible detectado en el contenido de texto del documento Office.",
                    editable=True,
                )

        return findings, flagged_tags
