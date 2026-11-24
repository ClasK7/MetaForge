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

import os
import re
import sys
import json
import shutil
import zipfile
import tempfile
import subprocess
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

OFFICE_EXTENSIONS = {".docx", ".xlsx", ".pptx"}

# Mapeo entre etiquetas mostradas por ExifTool y nodos XML reales de Office
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

# Firmas para el Escáner Forense / Detector de IA
AI_TAG_KEYWORDS = {
    "c2pa", "jumbf", "parameters", "workflow", "prompt", "digitalsourcetype",
    "ai", "dream", "sd-metadata", "generation", "seed", "negativeprompt",
    "synthid", "contentcredentials", "aigenerated"
}

AI_VALUE_PATTERNS = [
    r"\bopenai\b", r"\bchatgpt\b", r"\bgpt-[345]\b", r"\bdall-?e\b",
    r"\bmidjourney\b", r"\bstable\s*diffusion\b", r"\bcomfyui\b",
    r"\bautomatic1111\b", r"\bfirefly\b", r"\bgemini\b", r"\bimagen\b",
    r"\bclaude\b", r"\banthropic\b", r"\bperplexity\b", r"\bcopilot\b",
    r"\bsynthid\b", r"\bc2pa\b", r"\btrainedalgorithmicmedia\b",
    r"\bcompositewithtrainedalgorithmicmedia\b", r"\bpython-docx\b",
    r"\bpandoc\b", r"\baspose\b", r"\breportlab\b", r"\bweasyprint\b",
    r"\bgamma\.app\b", r"\bcanva\b", r"\blibreoffice\b", r"\bgoogle\s*docs\b",
    r"\bsteps:\s*\d+,\s*sampler:", r"\bcfg\s*scale:"
]

OPSEC_TAGS = {
    "creator", "author", "lastmodifiedby", "artist", "ownername",
    "description", "company", "manager", "lastprinted", "serialnumber",
    "internalserialnumber", "lensserialnumber", "hostcomputer", "usercomment",
    "gpslatitude", "gpslongitude", "gpsposition", "gpsaltitude", "make", "model"
}

READONLY_GROUPS = {"ExifTool", "File", "ZIP", "Composite"}

INVISIBLE_UNICODE_MAP = {
    "\u200b": ("Zero-Width Space (U+200B)", ""),
    "\u200c": ("Zero-Width Non-Joiner (U+200C)", ""),
    "\u200d": ("Zero-Width Joiner (U+200D)", ""),
    "\u2060": ("Word Joiner (U+2060)", ""),
    "\u00ad": ("Soft Hyphen (U+00AD)", ""),
    "\u202f": ("Narrow No-Break Space (U+202F)", " "),
    "\u2009": ("Thin Space (U+2009)", " "),
}


def get_exiftool_path():
    if getattr(sys, "frozen", False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))

    win_path = os.path.join(base_path, "bin", "exiftool.exe")
    if os.path.exists(win_path):
        return win_path
    return "exiftool"


EXIFTOOL_BIN = get_exiftool_path()


class OfficeMetadataEngine:
    """Motor nativo para editar, purgar metadatos y sanear Unicode oculto en Office."""

    @staticmethod
    def _register_core_ns():
        for prefix, uri in CORE_NS.items():
            ET.register_namespace(prefix, uri)

    @staticmethod
    def _register_app_ns():
        ET.register_namespace("", APP_NS)
        ET.register_namespace("vt", VT_NS)

    @classmethod
    def update_tags_batch(cls, file_path, updates: dict, strip_invisible_unicode=False):
        """Aplica múltiples cambios en un solo ciclo de escritura ZIP."""
        core_updates = {}
        app_updates = {}

        for raw_tag, new_val in updates.items():
            clean_tag = raw_tag.split(":")[-1].strip()
            if clean_tag in CORE_TAG_MAP:
                core_updates[clean_tag] = new_val
            elif clean_tag in APP_TAG_MAP:
                app_updates[clean_tag] = new_val
            else:
                valid = ", ".join(list(CORE_TAG_MAP.keys()) + list(APP_TAG_MAP.keys()))
                raise ValueError(
                    f"La etiqueta '{clean_tag}' no es editable en archivos de Office.\n\n"
                    f"Etiquetas soportadas:\n{valid}"
                )

        tmp_fd, tmp_zip_path = tempfile.mkstemp(suffix=".zip")
        os.close(tmp_fd)

        try:
            with zipfile.ZipFile(file_path, "r") as zin, zipfile.ZipFile(tmp_zip_path, "w") as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == "docProps/core.xml" and core_updates:
                        data = cls._modify_core_xml(data, core_updates)
                    elif item.filename == "docProps/app.xml" and app_updates:
                        data = cls._modify_app_xml(data, app_updates)
                    elif strip_invisible_unicode and item.filename.startswith("word/") and item.filename.endswith(".xml"):
                        data = cls._strip_invisible_chars_from_xml(data)
                    zout.writestr(item, data)

            shutil.copyfile(tmp_zip_path, file_path)
        finally:
            if os.path.exists(tmp_zip_path):
                os.remove(tmp_zip_path)

    @classmethod
    def update_tag(cls, file_path, raw_tag, new_value):
        cls.update_tags_batch(file_path, {raw_tag: new_value})

    @classmethod
    def wipe_all(cls, file_path, strip_invisible_unicode=True):
        updates = {
            "Title": "",
            "Subject": "",
            "Creator": "",
            "Description": "",
            "Keywords": "",
            "LastModifiedBy": "",
            "RevisionNumber": "1",
            "LastPrinted": "",
            "Category": "",
            "Application": "",
            "AppVersion": "",
            "Company": "",
            "Manager": "",
            "Template": "",
            "TotalEditTime": "0",
        }
        cls.update_tags_batch(file_path, updates, strip_invisible_unicode=strip_invisible_unicode)

    @classmethod
    def scan_invisible_unicode(cls, file_path):
        """Busca caracteres Unicode invisibles/atípicos dentro del contenido XML de un .docx."""
        counts = {}
        if not file_path.lower().endswith(".docx"):
            return counts
        try:
            with zipfile.ZipFile(file_path, "r") as zin:
                for name in zin.namelist():
                    if name in ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml") or name.startswith(("word/header", "word/footer")):
                        text = zin.read(name).decode("utf-8", errors="ignore")
                        for ch, (label, _) in INVISIBLE_UNICODE_MAP.items():
                            c = text.count(ch)
                            if c > 0:
                                counts[label] = counts.get(label, 0) + c
        except Exception:
            pass
        return counts

    @staticmethod
    def _strip_invisible_chars_from_xml(xml_bytes):
        """Limpia caracteres invisibles preservando intactos los prefijos XML de Word."""
        text = xml_bytes.decode("utf-8", errors="ignore")
        has_bom = text.startswith("\ufeff")
        body = text[1:] if has_bom else text
        for ch, (_, replacement) in INVISIBLE_UNICODE_MAP.items():
            body = body.replace(ch, replacement)
        final_text = ("\ufeff" + body) if has_bom else body
        return final_text.encode("utf-8")

    @classmethod
    def _modify_core_xml(cls, xml_bytes, updates_dict):
        cls._register_core_ns()
        root = ET.fromstring(xml_bytes)

        for clean_tag, new_value in updates_dict.items():
            ns_prefix, local_name = CORE_TAG_MAP[clean_tag]
            full_tag = f"{{{CORE_NS[ns_prefix]}}}{local_name}"
            elem = root.find(full_tag)

            # En fechas XML (created, modified, lastPrinted), un valor vacío corrompe Word -> eliminar el nodo
            if local_name in DATE_XML_LOCAL_NAMES and not new_value.strip():
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
    def _modify_app_xml(cls, xml_bytes, updates_dict):
        cls._register_app_ns()
        root = ET.fromstring(xml_bytes)

        for clean_tag, new_value in updates_dict.items():
            xml_name = APP_TAG_MAP[clean_tag]
            full_tag = f"{{{APP_NS}}}{xml_name}"
            elem = root.find(full_tag)
            if elem is None:
                elem = ET.SubElement(root, full_tag)

            if xml_name == "TotalTime":
                digits = re.findall(r"\d+", new_value)
                elem.text = digits[0] if digits else "0"
            else:
                elem.text = new_value

        return ET.tostring(root, encoding="utf-8", xml_declaration=True)


class ForensicScanner:
    """Analiza metadatos y estructura para detectar firmas de IA, anomalías y fugas OPSEC."""

    @staticmethod
    def analyze(file_path, metadata: dict):
        findings = []
        flagged_tags = {}  # {full_tag: severity ('ai_high' | 'anomaly_med' | 'opsec_low')}

        def add_finding(severity, tag_key, value, reason, editable=True):
            findings.append({
                "severity": severity,
                "tag": tag_key,
                "value": str(value),
                "reason": reason,
                "editable": editable,
            })
            # Prioridad de color si una etiqueta dispara múltiples reglas
            priority = {"ai_high": 3, "anomaly_med": 2, "opsec_low": 1}
            if tag_key and (tag_key not in flagged_tags or priority[severity] > priority[flagged_tags[tag_key]]):
                flagged_tags[tag_key] = severity

        # 1. Análisis etiqueta por etiqueta
        for full_key, raw_val in metadata.items():
            if full_key == "SourceFile":
                continue
            val_str = str(raw_val).strip()
            if not val_str:
                continue

            group = full_key.split(":")[0] if ":" in full_key else ""
            short_tag = full_key.split(":")[-1].lower()
            is_editable = group not in READONLY_GROUPS

            # A) Etiquetas típicas de generación IA / C2PA
            if any(kw in short_tag or kw in group.lower() for kw in AI_TAG_KEYWORDS):
                add_finding("ai_high", full_key, val_str, "Etiqueta asociada a generación por IA o credenciales C2PA/JUMBF.", is_editable)
                continue

            # B) Patrones de IA / generadores en los valores
            val_lower = val_str.lower()
            matched_ai = False
            for pattern in AI_VALUE_PATTERNS:
                if re.search(pattern, val_lower):
                    add_finding("ai_high", full_key, val_str, f"Firma de herramienta de IA o generador automatizado detectada ({pattern.strip(r'\\b')}).", is_editable)
                    matched_ai = True
                    break
            if matched_ai:
                continue

            # C) Fugas de privacidad / OPSEC
            if short_tag in OPSEC_TAGS or "gps" in short_tag:
                add_finding("opsec_low", full_key, val_str, "Metadato de identidad, ubicación, dispositivo o trazabilidad personal (OPSEC).", is_editable)

        # 2. Heurísticas y anomalías de documentos Office / PDF
        words = ForensicScanner._extract_int(metadata.get("XML:Words"))
        edit_mins = ForensicScanner._extract_int(metadata.get("XML:TotalEditTime"))
        revisions = ForensicScanner._extract_int(metadata.get("XML:RevisionNumber"))

        if words is not None and words >= 250:
            if edit_mins is not None and edit_mins <= 2:
                add_finding(
                    "anomaly_med", "XML:TotalEditTime", metadata.get("XML:TotalEditTime"),
                    f"Anomalía: {words} palabras con solo {edit_mins} min de edición (patrón común en archivos exportados o copiados en bloque).",
                    True
                )
            if revisions is not None and revisions <= 1:
                add_finding(
                    "anomaly_med", "XML:RevisionNumber", metadata.get("XML:RevisionNumber"),
                    f"Anomalía: documento extenso ({words} palabras) con una sola revisión registrada.",
                    True
                )

        # Brecha de años entre CreateDate y ModifyDate
        create_date = str(metadata.get("XML:CreateDate") or metadata.get("PDF:CreateDate") or "")
        modify_date = str(metadata.get("XML:ModifyDate") or metadata.get("PDF:ModifyDate") or "")
        if len(create_date) >= 4 and len(modify_date) >= 4:
            try:
                y_create = int(create_date[:4])
                y_mod = int(modify_date[:4])
                if abs(y_mod - y_create) >= 2:
                    key_name = "XML:CreateDate" if "XML:CreateDate" in metadata else "PDF:CreateDate"
                    add_finding(
                        "anomaly_med", key_name, create_date,
                        f"Inconsistencia temporal: creado en {y_create} pero modificado en {y_mod} (plantilla antigua o fecha residual).",
                        True
                    )
            except ValueError:
                pass

        # HeadingPairs bilingüe (ej. "Título" y "Title" juntos)
        heading_pairs = metadata.get("XML:HeadingPairs")
        if isinstance(heading_pairs, list):
            joined = " ".join(str(x) for x in heading_pairs)
            if "Título" in joined and "Title" in joined:
                add_finding(
                    "anomaly_med", "XML:HeadingPairs", joined,
                    "Mezcla de idiomas interno ('Título' + 'Title'): típico al generar un documento en un entorno en inglés y editarlo en Word en español.",
                    False
                )

        # Discrepancia entre Creator y LastModifiedBy
        creator = str(metadata.get("XMP:Creator", "")).strip()
        last_mod = str(metadata.get("XML:LastModifiedBy", "")).strip()
        if creator and last_mod and creator != last_mod:
            add_finding(
                "anomaly_med", "XMP:Creator", f"{creator} ≠ {last_mod}",
                "El autor original (Creator) difiere del último editor (LastModifiedBy).",
                True
            )

        # 3. Caracteres Unicode invisibles dentro de .docx
        unicode_hits = OfficeMetadataEngine.scan_invisible_unicode(file_path)
        for label, count in unicode_hits.items():
            add_finding(
                "ai_high", "[Contenido DOCX]", f"{count} ocurrencias de {label}",
                "Carácter Unicode invisible/atípico detectado en el texto interno (frecuente al copiar o exportar texto desde LLMs).",
                True
            )

        return findings, flagged_tags, unicode_hits

    @staticmethod
    def _extract_int(value):
        if value is None:
            return None
        match = re.search(r"\d+", str(value))
        return int(match.group()) if match else None


class MetaForgeApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("MetaForge - Inspector, Escáner IA y Editor Forense de Metadatos")
        self.geometry("1160x700")
        self.current_file = None
        self.metadata_dict = {}
        self.last_findings = []
        self.last_flagged_tags = {}
        self.last_unicode_hits = {}
        self._setup_ui()

    def _setup_ui(self):
        top_frame = ctk.CTkFrame(self)
        top_frame.pack(fill="x", padx=15, pady=(15, 8))

        self.btn_open = ctk.CTkButton(top_frame, text="Seleccionar Archivo", command=self.load_file, width=155)
        self.btn_open.pack(side="left", padx=8, pady=10)

        self.lbl_filepath = ctk.CTkLabel(top_frame, text="Ningún archivo cargado...", text_color="gray", anchor="w")
        self.lbl_filepath.pack(side="left", padx=8, fill="x", expand=True)

        self.btn_scan = ctk.CTkButton(
            top_frame, text="Escanear IA / Forense", command=self.run_forensic_scan,
            fg_color="#D35400", hover_color="#A04000", width=175
        )
        self.btn_scan.pack(side="right", padx=6, pady=10)

        self.btn_export = ctk.CTkButton(
            top_frame, text="Exportar JSON", command=self.export_json,
            fg_color="#2B7A78", hover_color="#205C5A", width=115
        )
        self.btn_export.pack(side="right", padx=6, pady=10)

        self.btn_wipe = ctk.CTkButton(
            top_frame, text="Limpieza Forense (Todo)", command=self.wipe_metadata,
            fg_color="#B22222", hover_color="#8B0000", width=180
        )
        self.btn_wipe.pack(side="right", padx=6, pady=10)

        # Leyenda visual del escáner
        legend_frame = ctk.CTkFrame(self, fg_color="transparent")
        legend_frame.pack(fill="x", padx=20, pady=(0, 4))
        ctk.CTkLabel(legend_frame, text="Leyenda del Escáner:", font=("Segoe UI", 11, "bold")).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(legend_frame, text=" ■ Rastro IA / Generador ", fg_color="#6E1D1D", corner_radius=4, font=("Segoe UI", 11)).pack(side="left", padx=4)
        ctk.CTkLabel(legend_frame, text=" ■ Anomalía Forense ", fg_color="#7A5214", corner_radius=4, font=("Segoe UI", 11)).pack(side="left", padx=4)
        ctk.CTkLabel(legend_frame, text=" ■ Privacidad / OPSEC ", fg_color="#2D2B6B", corner_radius=4, font=("Segoe UI", 11)).pack(side="left", padx=4)

        mid_frame = ctk.CTkFrame(self)
        mid_frame.pack(fill="both", expand=True, padx=15, pady=5)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#2b2b2b", foreground="white", fieldbackground="#2b2b2b", rowheight=28, borderwidth=0)
        style.map("Treeview", background=[("selected", "#1f538d")])
        style.configure("Treeview.Heading", background="#1f1f1f", foreground="white", font=("Segoe UI", 10, "bold"))

        self.tree = ttk.Treeview(mid_frame, columns=("tag", "value"), show="headings", selectmode="browse")
        self.tree.heading("tag", text="Etiqueta / Campo (Tag)")
        self.tree.heading("value", text="Valor Actual")
        self.tree.column("tag", width=330, anchor="w")
        self.tree.column("value", width=750, anchor="w")

        # Colores para filas detectadas
        self.tree.tag_configure("ai_high", background="#6E1D1D", foreground="#ffffff")
        self.tree.tag_configure("anomaly_med", background="#7A5214", foreground="#ffffff")
        self.tree.tag_configure("opsec_low", background="#2D2B6B", foreground="#ffffff")

        scrollbar = ctk.CTkScrollbar(mid_frame, orientation="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        scrollbar.pack(side="right", fill="y", padx=(0, 10), pady=10)
        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)

        bottom_frame = ctk.CTkFrame(self)
        bottom_frame.pack(fill="x", padx=15, pady=(5, 15))

        ctk.CTkLabel(bottom_frame, text="Etiqueta (ej. Creator, LastModifiedBy, CreateDate, Software):").grid(row=0, column=0, padx=10, pady=(10, 2), sticky="w")
        ctk.CTkLabel(bottom_frame, text="Nuevo Valor (dejar vacío para borrar contenido):").grid(row=0, column=1, padx=10, pady=(10, 2), sticky="w")

        self.entry_tag = ctk.CTkEntry(bottom_frame, width=310, placeholder_text="Nombre de la etiqueta...")
        self.entry_tag.grid(row=1, column=0, padx=10, pady=(0, 12), sticky="ew")

        self.entry_val = ctk.CTkEntry(bottom_frame, width=520, placeholder_text="Valor a inyectar...")
        self.entry_val.grid(row=1, column=1, padx=10, pady=(0, 12), sticky="ew")

        self.btn_save = ctk.CTkButton(
            bottom_frame, text="Aplicar Cambio", command=self.update_tag,
            width=160, fg_color="#1E8449", hover_color="#145A32"
        )
        self.btn_save.grid(row=1, column=2, padx=10, pady=(0, 12))
        bottom_frame.columnconfigure(1, weight=1)

    def _is_office_file(self):
        if not self.current_file:
            return False
        _, ext = os.path.splitext(self.current_file.lower())
        return ext in OFFICE_EXTENSIONS

    def run_exiftool(self, args):
        cmd = [EXIFTOOL_BIN, "-charset", "utf8"] + args
        startupinfo = None
        if os.name == "nt":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        return subprocess.run(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            text=True, encoding="utf-8", startupinfo=startupinfo
        )

    def _safe_exiftool_write(self, exif_args):
        _, ext = os.path.splitext(self.current_file)
        tmp_fd, tmp_path = tempfile.mkstemp(suffix=ext)
        os.close(tmp_fd)

        try:
            shutil.copy2(self.current_file, tmp_path)
            res = self.run_exiftool(["-overwrite_original"] + exif_args + [tmp_path])
            if res.returncode != 0:
                return False, res.stderr
            shutil.copyfile(tmp_path, self.current_file)
            return True, ""
        except PermissionError:
            return False, (
                "Permiso denegado al sobrescribir el archivo.\n"
                "Verifica que no esté abierto en otro programa."
            )
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    def load_file(self):
        file_path = filedialog.askopenfilename(title="Selecciona cualquier archivo para inspeccionar")
        if not file_path:
            return
        self.current_file = file_path
        self.last_findings = []
        self.last_flagged_tags = {}
        self.last_unicode_hits = {}
        self.lbl_filepath.configure(text=file_path, text_color="white")
        self.refresh_metadata()

    def refresh_metadata(self, keep_highlights=False):
        if not self.current_file:
            return

        res = self.run_exiftool(["-j", "-G", self.current_file])
        if res.returncode != 0:
            messagebox.showerror("Error", f"No se pudo leer el archivo:\n{res.stderr}")
            return

        try:
            data = json.loads(res.stdout)[0]
            self.metadata_dict = data
            self.tree.delete(*self.tree.get_children())

            if not keep_highlights:
                self.last_flagged_tags = {}

            for key, val in data.items():
                if key == "SourceFile":
                    continue
                row_tag = (self.last_flagged_tags[key],) if key in self.last_flagged_tags else ()
                self.tree.insert("", "end", values=(key, str(val)), tags=row_tag)
        except Exception as e:
            messagebox.showerror("Error de parseo", str(e))

    def run_forensic_scan(self):
        if not self.current_file or not self.metadata_dict:
            messagebox.showwarning("Aviso", "Primero selecciona un archivo para escanear.")
            return

        findings, flagged_tags, unicode_hits = ForensicScanner.analyze(self.current_file, self.metadata_dict)
        self.last_findings = findings
        self.last_flagged_tags = flagged_tags
        self.last_unicode_hits = unicode_hits

        # Re-renderizar la tabla aplicando colores a las filas sospechosas
        self.refresh_metadata(keep_highlights=True)
        self._show_scan_report_window(findings, unicode_hits)

    def _show_scan_report_window(self, findings, unicode_hits):
        win = ctk.CTkToplevel(self)
        win.title("Reporte de Escaneo Forense e IA")
        win.geometry("760x520")
        win.grab_set()

        if not findings:
            ctk.CTkLabel(
                win,
                text="Archivo limpio: no se detectaron firmas de IA, anomalías temporales ni fugas OPSEC.",
                font=("Segoe UI", 14, "bold"),
                text_color="#2ECC71",
                wraplength=700
            ).pack(pady=40, padx=20)
            ctk.CTkButton(win, text="Cerrar", command=win.destroy).pack(pady=10)
            return

        ai_count = sum(1 for f in findings if f["severity"] == "ai_high")
        anom_count = sum(1 for f in findings if f["severity"] == "anomaly_med")
        opsec_count = sum(1 for f in findings if f["severity"] == "opsec_low")

        header_text = (
            f"Hallazgos detectados: {len(findings)}   |   "
            f"Rastros IA: {ai_count}   |   Anomalías: {anom_count}   |   OPSEC: {opsec_count}"
        )
        ctk.CTkLabel(win, text=header_text, font=("Segoe UI", 13, "bold")).pack(pady=(15, 8), padx=15)

        scroll_box = ctk.CTkScrollableFrame(win, width=710, height=360)
        scroll_box.pack(fill="both", expand=True, padx=15, pady=5)

        badge_map = {
            "ai_high": ("[RASTRO IA]", "#E74C3C"),
            "anomaly_med": ("[ANOMALÍA]", "#F39C12"),
            "opsec_low": ("[OPSEC]", "#7F8CFF"),
        }

        for item in findings:
            card = ctk.CTkFrame(scroll_box)
            card.pack(fill="x", padx=6, pady=5)
            badge_txt, badge_color = badge_map[item["severity"]]

            top_line = f"{badge_txt}  {item['tag']}  →  {item['value'][:90]}"
            ctk.CTkLabel(card, text=top_line, font=("Segoe UI", 12, "bold"), text_color=badge_color, anchor="w").pack(fill="x", padx=10, pady=(6, 2))
            ctk.CTkLabel(card, text=item["reason"], font=("Segoe UI", 11), text_color="#DCDCDC", anchor="w", wraplength=670, justify="left").pack(fill="x", padx=10, pady=(0, 6))

        btn_bar = ctk.CTkFrame(win, fg_color="transparent")
        btn_bar.pack(fill="x", padx=15, pady=12)

        ctk.CTkButton(
            btn_bar,
            text="Limpiar Solo Hallazgos Editables",
            command=lambda: self._clean_detected_only(win),
            fg_color="#D35400",
            hover_color="#A04000",
            width=230
        ).pack(side="left", padx=6)

        ctk.CTkLabel(
            btn_bar,
            text="También puedes cerrar esta ventana y editar cada fila resaltada manualmente.",
            font=("Segoe UI", 11),
            text_color="gray"
        ).pack(side="left", padx=10)

        ctk.CTkButton(btn_bar, text="Cerrar", command=win.destroy, width=100).pack(side="right", padx=6)

    def _clean_detected_only(self, report_window):
        if not self.last_findings:
            return

        if self._is_office_file():
            updates = {}
            for f in self.last_findings:
                if not f["editable"] or f["tag"] == "[Contenido DOCX]":
                    continue
                short = f["tag"].split(":")[-1]
                if short in CORE_TAG_MAP or short in APP_TAG_MAP:
                    updates[short] = "1" if short == "RevisionNumber" else ("0" if short == "TotalEditTime" else "")

            strip_uni = bool(self.last_unicode_hits)
            if not updates and not strip_uni:
                messagebox.showinfo("Aviso", "Los hallazgos actuales son solo informativos y no requieren borrado automático.")
                return

            try:
                OfficeMetadataEngine.update_tags_batch(self.current_file, updates, strip_invisible_unicode=strip_uni)
                report_window.destroy()
                self.refresh_metadata()
                messagebox.showinfo("Limpieza selectiva completada", "Se limpiaron los campos marcados y caracteres invisibles detectados.")
            except PermissionError:
                messagebox.showerror("Archivo bloqueado", "Cierra Word/Excel/PowerPoint antes de limpiar.")
            except Exception as e:
                messagebox.showerror("Error", str(e))
        else:
            exif_args = []
            seen = set()
            for f in self.last_findings:
                if not f["editable"]:
                    continue
                tag = f["tag"]
                if tag not in seen:
                    seen.add(tag)
                    exif_args.append(f"-{tag}=")

            if not exif_args:
                messagebox.showinfo("Aviso", "No hay etiquetas editables para limpiar selectivamente.")
                return

            ok, err = self._safe_exiftool_write(exif_args)
            if ok:
                report_window.destroy()
                self.refresh_metadata()
                messagebox.showinfo("Limpieza selectiva completada", f"Se eliminaron {len(exif_args)} etiquetas sospechosas.")
            else:
                messagebox.showerror("Error", err)

    def on_row_select(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        tag, val = self.tree.item(selected[0], "values")
        self.entry_tag.delete(0, tk.END)
        self.entry_tag.insert(0, tag)
        self.entry_val.delete(0, tk.END)
        self.entry_val.insert(0, val)

    def update_tag(self):
        if not self.current_file:
            messagebox.showwarning("Aviso", "Primero selecciona un archivo.")
            return

        tag = self.entry_tag.get().strip()
        val = self.entry_val.get()

        if not tag:
            messagebox.showwarning("Aviso", "Debes especificar el nombre de una etiqueta.")
            return

        if self._is_office_file():
            try:
                OfficeMetadataEngine.update_tag(self.current_file, tag, val)
                self.refresh_metadata()
                messagebox.showinfo("Éxito", f"Metadato de Office '{tag}' actualizado correctamente.")
            except PermissionError:
                messagebox.showerror("Archivo bloqueado", "Cierra el documento en Microsoft Office antes de guardar.")
            except Exception as e:
                messagebox.showerror("Error en documento Office", str(e))
        else:
            ok, err = self._safe_exiftool_write([f"-{tag}={val}"])
            if ok:
                self.refresh_metadata()
                messagebox.showinfo("Éxito", f"Etiqueta '{tag}' actualizada correctamente.")
            else:
                messagebox.showerror("Error al escribir", f"ExifTool reportó un problema:\n{err}")

    def wipe_metadata(self):
        if not self.current_file:
            return
        confirm = messagebox.askyesno(
            "Confirmar Limpieza Forense",
            "¿Estás seguro de purgar TODOS los metadatos editables y caracteres invisibles de este archivo?"
        )
        if not confirm:
            return

        if self._is_office_file():
            try:
                OfficeMetadataEngine.wipe_all(self.current_file, strip_invisible_unicode=True)
                self.refresh_metadata()
                messagebox.showinfo("Limpio", "Metadatos internos de Office y caracteres ocultos purgados.")
            except PermissionError:
                messagebox.showerror("Archivo bloqueado", "Cierra el documento en Office antes de limpiar.")
            except Exception as e:
                messagebox.showerror("Error", str(e))
        else:
            ok, err = self._safe_exiftool_write(["-all="])
            if ok:
                self.refresh_metadata()
                messagebox.showinfo("Limpio", "Todos los metadatos han sido purgados.")
            else:
                messagebox.showerror("Error", err)

    def export_json(self):
        if not self.metadata_dict:
            return
        save_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if save_path:
            with open(save_path, "w", encoding="utf-8") as f:
                json.dump(self.metadata_dict, f, indent=4, ensure_ascii=False)
            messagebox.showinfo("Exportado", "Reporte de metadatos guardado en JSON.")


if __name__ == "__main__":
    app = MetaForgeApp()
    app.mainloop()