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

"""Interfaz gráfica responsiva CustomTkinter, diálogo de reporte e hilo serializado."""

from __future__ import annotations

import logging
import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk
from typing import Any, Callable

import customtkinter as ctk

from metaforge import DestinationNotWritableError, MetaForgeError
from metaforge.office_engine import CORE_TAG_MAP, DATE_XML_LOCAL_NAMES
from metaforge.session import DocumentSession

logger = logging.getLogger(__name__)

_UI_FONT_FAMILY = "Segoe UI" if sys.platform == "win32" else "DejaVu Sans"

_SEVERITY_LABELS = {
    "ai_high": "[IA]",
    "anomaly_med": "[Anomalía]",
    "opsec_low": "[OPSEC]",
}

_PDF_WARNING_TEXT = (
    "ADVERTENCIA PDF: ExifTool aplica actualizaciones incrementales en archivos PDF. "
    "Los metadatos anteriores pueden seguir recuperables dentro del archivo binario."
)


def _fixed_treeview_map(style: ttk.Style, option: str) -> list[tuple[Any, ...]]:
    """Corrige el bug de colores de fila de ``ttk.Treeview`` en Tk 8.6.9+."""
    return [
        elm
        for elm in style.map("Treeview", query_opt=option)
        if elm[:2] != ("!disabled", "!selected")
    ]


class ReportDialog(ctk.CTkToplevel):
    """Ventana modal responsiva con los hallazgos del escáner forense e IA."""

    def __init__(
        self,
        parent: "MetaForgeApp",
        findings: list[dict[str, Any]],
        unicode_hits: dict[str, int],
        on_clean_callback: Callable[["ReportDialog"], None],
    ) -> None:
        """Construye la ventana de reporte con ajuste dinámico de ancho y grab diferido.

        Args:
            parent: Ventana principal ``MetaForgeApp``.
            findings: Lista de hallazgos generada por ``ForensicScanner.analyze``.
            unicode_hits: Conteos Unicode detectados en Office.
            on_clean_callback: Callback ejecutado al pulsar «Limpiar Solo Hallazgos Editables».
        """
        super().__init__(parent)
        self.title("Reporte de Escaneo Forense e IA")
        self.geometry("800x540")
        self.minsize(640, 420)
        self.transient(parent)

        self._reason_labels: list[ctk.CTkLabel] = []
        self._build_contents(findings, unicode_hits, on_clean_callback)
        self.bind("<Configure>", self._on_resize)
        self.after(20, lambda: self._safe_grab_set(attempts=5))

    def _safe_grab_set(self, attempts: int) -> None:
        # En Kali Linux (X11/Wayland) grab_set falla si la ventana aún no es visible.
        if not self.winfo_exists():
            return
        try:
            self.grab_set()
        except tk.TclError:
            if attempts > 1:
                self.after(20, lambda: self._safe_grab_set(attempts - 1))

    def _build_contents(
        self,
        findings: list[dict[str, Any]],
        unicode_hits: dict[str, int],
        on_clean_callback: Callable[["ReportDialog"], None],
    ) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        if not findings:
            ctk.CTkLabel(
                self,
                text="Archivo sin hallazgos: no se detectaron firmas de IA/C2PA, anomalías ni fugas OPSEC.",
                font=ctk.CTkFont(family=_UI_FONT_FAMILY, size=14, weight="bold"),
                text_color="#2ECC71",
                wraplength=700,
            ).grid(row=0, column=0, pady=40, padx=20, sticky="ew")
            ctk.CTkButton(self, text="Cerrar", command=self.destroy, width=120).grid(
                row=2, column=0, pady=14
            )
            return

        ai_count = sum(1 for f in findings if f["severity"] == "ai_high")
        anom_count = sum(1 for f in findings if f["severity"] == "anomaly_med")
        opsec_count = sum(1 for f in findings if f["severity"] == "opsec_low")
        editable_count = sum(1 for f in findings if f.get("editable"))

        header_text = (
            f"Hallazgos: {len(findings)}   |   IA/C2PA: {ai_count}   |   "
            f"Anomalías: {anom_count}   |   OPSEC: {opsec_count}   |   Editables: {editable_count}"
        )
        ctk.CTkLabel(
            self,
            text=header_text,
            font=ctk.CTkFont(family=_UI_FONT_FAMILY, size=13, weight="bold"),
        ).grid(row=0, column=0, padx=16, pady=(14, 8), sticky="w")

        scroll_box = ctk.CTkScrollableFrame(self)
        scroll_box.grid(row=1, column=0, padx=16, pady=4, sticky="nsew")
        scroll_box.grid_columnconfigure(0, weight=1)

        badge_map = {
            "ai_high": ("[RASTRO IA]", "#E74C3C"),
            "anomaly_med": ("[ANOMALÍA]", "#F39C12"),
            "opsec_low": ("[OPSEC]", "#7F8CFF"),
        }

        for idx, item in enumerate(findings):
            card = ctk.CTkFrame(scroll_box)
            card.grid(row=idx, column=0, padx=6, pady=5, sticky="ew")
            card.grid_columnconfigure(0, weight=1)

            badge_txt, badge_color = badge_map[item["severity"]]
            edit_badge = "[Editable]" if item.get("editable") else "[Solo informativo]"
            top_line = f"{badge_txt} {edit_badge}  {item['tag']}  →  {item['value'][:90]}"
            ctk.CTkLabel(
                card,
                text=top_line,
                font=ctk.CTkFont(family=_UI_FONT_FAMILY, size=12, weight="bold"),
                text_color=badge_color,
                anchor="w",
            ).grid(row=0, column=0, padx=10, pady=(6, 2), sticky="ew")

            reason_lbl = ctk.CTkLabel(
                card,
                text=item["reason"],
                font=ctk.CTkFont(family=_UI_FONT_FAMILY, size=11),
                text_color="#DCDCDC",
                anchor="w",
                justify="left",
                wraplength=680,
            )
            reason_lbl.grid(row=1, column=0, padx=10, pady=(0, 6), sticky="ew")
            self._reason_labels.append(reason_lbl)

        btn_bar = ctk.CTkFrame(self, fg_color="transparent")
        btn_bar.grid(row=2, column=0, padx=16, pady=12, sticky="ew")
        btn_bar.grid_columnconfigure(1, weight=1)

        can_clean = editable_count > 0 or bool(unicode_hits)
        self.btn_clean = ctk.CTkButton(
            btn_bar,
            text="Limpiar Solo Hallazgos Editables",
            command=lambda: on_clean_callback(self),
            fg_color="#D35400",
            hover_color="#A04000",
            width=235,
            state="normal" if can_clean else "disabled",
        )
        self.btn_clean.grid(row=0, column=0, padx=(0, 10))

        ctk.CTkLabel(
            btn_bar,
            text="Los cambios se aplican sobre la copia temporal; pulsa «Guardar copia» para exportar.",
            font=ctk.CTkFont(family=_UI_FONT_FAMILY, size=11),
            text_color="gray",
            anchor="w",
        ).grid(row=0, column=1, sticky="ew", padx=6)

        ctk.CTkButton(btn_bar, text="Cerrar", command=self.destroy, width=100).grid(
            row=0, column=2, padx=(10, 0)
        )

    def _on_resize(self, _event: Any) -> None:
        if not self._reason_labels:
            return
        wrap = max(420, self.winfo_width() - 95)
        for lbl in self._reason_labels:
            if lbl.winfo_exists():
                lbl.configure(wraplength=wrap)


class MetaForgeApp(ctk.CTk):
    """Ventana principal responsiva de MetaForge con ejecución en segundo plano."""

    def __init__(self, session: DocumentSession, exiftool_version: str) -> None:
        """Inicializa la interfaz principal y enlaza la sesión de documento.

        Args:
            session: Instancia de ``DocumentSession`` para operar sobre archivos.
            exiftool_version: Versión verificada de ExifTool al arranque.
        """
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")
        super().__init__()

        self.session = session
        self.exiftool_version = exiftool_version
        self._busy = False
        self._result_queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self._poll_after_id: str | None = None
        self._filter_after_id: str | None = None
        self._pending_on_success: Callable[[Any], None] | None = None

        self.title("MetaForge - Inspector, Escáner IA y Editor Forense de Metadatos")
        self.geometry("1200x740")
        self.minsize(960, 640)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self._setup_ui()
        self._update_controls_state()

    def _setup_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # Zona 1: Barra superior de archivo y acciones
        top_frame = ctk.CTkFrame(self)
        top_frame.grid(row=0, column=0, padx=14, pady=(12, 6), sticky="ew")
        top_frame.grid_columnconfigure(1, weight=1)

        self.btn_open = ctk.CTkButton(
            top_frame, text="Seleccionar Archivo", command=self._on_open_file, width=160
        )
        self.btn_open.grid(row=0, column=0, padx=(10, 8), pady=(10, 6), sticky="w")

        self.lbl_filepath = ctk.CTkLabel(
            top_frame,
            text="Ningún archivo cargado...",
            text_color="gray",
            anchor="w",
            font=ctk.CTkFont(family=_UI_FONT_FAMILY, size=12),
        )
        self.lbl_filepath.grid(row=0, column=1, padx=8, pady=(10, 6), sticky="ew")

        self.btn_discard = ctk.CTkButton(
            top_frame,
            text="Descartar cambios",
            command=self._on_discard_changes,
            fg_color="#566573",
            hover_color="#34495E",
            width=145,
        )
        self.btn_discard.grid(row=0, column=2, padx=6, pady=(10, 6))

        self.btn_save_copy = ctk.CTkButton(
            top_frame,
            text="Guardar copia",
            command=self._on_save_copy,
            fg_color="#1E8449",
            hover_color="#145A32",
            width=165,
        )
        self.btn_save_copy.grid(row=0, column=3, padx=(6, 10), pady=(10, 6))

        # Segunda fila compacta de acciones forenses para evitar recortes en 960 px
        actions_bar = ctk.CTkFrame(top_frame, fg_color="transparent")
        actions_bar.grid(row=1, column=0, columnspan=4, padx=10, pady=(0, 8), sticky="ew")
        actions_bar.grid_columnconfigure(3, weight=1)

        self.btn_scan = ctk.CTkButton(
            actions_bar,
            text="Escanear IA / Forense",
            command=self._on_run_scan,
            fg_color="#D35400",
            hover_color="#A04000",
            width=175,
        )
        self.btn_scan.grid(row=0, column=0, padx=(0, 8))

        self.btn_wipe = ctk.CTkButton(
            actions_bar,
            text="Limpieza Forense (Todo)",
            command=self._on_wipe_all,
            fg_color="#B22222",
            hover_color="#8B0000",
            width=185,
        )
        self.btn_wipe.grid(row=0, column=1, padx=6)

        self.btn_export = ctk.CTkButton(
            actions_bar,
            text="Exportar JSON",
            command=self._on_export_json,
            fg_color="#2B7A78",
            hover_color="#205C5A",
            width=130,
        )
        self.btn_export.grid(row=0, column=2, padx=6)

        self.lbl_meta_info = ctk.CTkLabel(
            actions_bar,
            text=f"ExifTool v{self.exiftool_version}",
            text_color="gray",
            anchor="e",
            font=ctk.CTkFont(family=_UI_FONT_FAMILY, size=11),
        )
        self.lbl_meta_info.grid(row=0, column=3, padx=(8, 0), sticky="ew")

        # Zona 2: Franja fija de advertencia para PDF (oculta por defecto)
        self.pdf_banner = ctk.CTkFrame(self, fg_color="#7A5214", corner_radius=6)
        self.pdf_banner.grid(row=1, column=0, padx=14, pady=(0, 6), sticky="ew")
        self.pdf_banner.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            self.pdf_banner,
            text=_PDF_WARNING_TEXT,
            text_color="#FFFFFF",
            font=ctk.CTkFont(family=_UI_FONT_FAMILY, size=12, weight="bold"),
            anchor="w",
        ).grid(row=0, column=0, padx=12, pady=7, sticky="ew")
        self.pdf_banner.grid_remove()

        # Zona 3: Barra de búsqueda, filtro rápido y leyenda
        filter_frame = ctk.CTkFrame(self)
        filter_frame.grid(row=2, column=0, padx=14, pady=(0, 6), sticky="ew")
        filter_frame.grid_columnconfigure(0, weight=1)

        self.entry_search = ctk.CTkEntry(
            filter_frame,
            placeholder_text="Buscar por etiqueta o valor...",
        )
        self.entry_search.grid(row=0, column=0, padx=(10, 6), pady=8, sticky="ew")
        self.entry_search.bind("<KeyRelease>", self._schedule_filter_refresh)

        self.filter_mode = ctk.CTkSegmentedButton(
            filter_frame,
            values=["Todos", "Editables", "Solo lectura", "Hallazgos"],
            command=lambda _v: self._populate_table(),
        )
        self.filter_mode.set("Todos")
        self.filter_mode.grid(row=0, column=1, padx=6, pady=8)

        self.btn_clear_filter = ctk.CTkButton(
            filter_frame,
            text="Limpiar",
            command=self._on_clear_filter,
            width=80,
            fg_color="#4A5568",
            hover_color="#2D3748",
        )
        self.btn_clear_filter.grid(row=0, column=2, padx=(4, 10), pady=8)

        legend_row = ctk.CTkFrame(filter_frame, fg_color="transparent")
        legend_row.grid(row=1, column=0, columnspan=3, padx=10, pady=(0, 7), sticky="ew")
        legend_row.grid_columnconfigure(5, weight=1)

        ctk.CTkLabel(
            legend_row,
            text="Leyenda:",
            font=ctk.CTkFont(family=_UI_FONT_FAMILY, size=11, weight="bold"),
        ).grid(row=0, column=0, padx=(0, 8))
        ctk.CTkLabel(
            legend_row, text=" ■ IA / C2PA ", fg_color="#6E1D1D", corner_radius=4,
            font=ctk.CTkFont(family=_UI_FONT_FAMILY, size=11),
        ).grid(row=0, column=1, padx=3)
        ctk.CTkLabel(
            legend_row, text=" ■ Anomalía ", fg_color="#7A5214", corner_radius=4,
            font=ctk.CTkFont(family=_UI_FONT_FAMILY, size=11),
        ).grid(row=0, column=2, padx=3)
        ctk.CTkLabel(
            legend_row, text=" ■ OPSEC ", fg_color="#2D2B6B", corner_radius=4,
            font=ctk.CTkFont(family=_UI_FONT_FAMILY, size=11),
        ).grid(row=0, column=3, padx=3)
        ctk.CTkLabel(
            legend_row, text=" ■ Solo lectura ", fg_color="#3A3A3A", text_color="#B8B8B8",
            corner_radius=4, font=ctk.CTkFont(family=_UI_FONT_FAMILY, size=11),
        ).grid(row=0, column=4, padx=3)

        self.lbl_count = ctk.CTkLabel(
            legend_row,
            text="Mostrando 0 de 0 etiquetas",
            text_color="gray",
            anchor="e",
            font=ctk.CTkFont(family=_UI_FONT_FAMILY, size=11),
        )
        self.lbl_count.grid(row=0, column=5, sticky="ew")

        # Zona 4: Tabla principal responsiva de 3 columnas
        mid_frame = ctk.CTkFrame(self)
        mid_frame.grid(row=3, column=0, padx=14, pady=(0, 6), sticky="nsew")
        mid_frame.grid_columnconfigure(0, weight=1)
        mid_frame.grid_rowconfigure(0, weight=1)

        row_font = tkfont.Font(family=_UI_FONT_FAMILY, size=10)
        row_height = max(26, row_font.metrics("linespace") + 10)

        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Treeview",
            background="#2b2b2b",
            foreground="#ffffff",
            fieldbackground="#2b2b2b",
            rowheight=row_height,
            borderwidth=0,
            font=(_UI_FONT_FAMILY, 10),
        )
        style.map(
            "Treeview",
            foreground=_fixed_treeview_map(style, "foreground"),
            background=[("selected", "#1f538d")] + _fixed_treeview_map(style, "background"),
        )
        style.configure(
            "Treeview.Heading",
            background="#1f1f1f",
            foreground="#ffffff",
            font=(_UI_FONT_FAMILY, 10, "bold"),
        )

        self.tree = ttk.Treeview(
            mid_frame,
            columns=("status", "tag", "value"),
            show="headings",
            selectmode="browse",
        )
        self.tree.heading("status", text="Estado")
        self.tree.heading("tag", text="Etiqueta / Campo")
        self.tree.heading("value", text="Valor Actual")
        self.tree.column("status", width=170, minwidth=120, stretch=False, anchor="w")
        self.tree.column("tag", width=310, minwidth=180, stretch=True, anchor="w")
        self.tree.column("value", width=620, minwidth=240, stretch=True, anchor="w")

        # Precedencia de tags: solo lectura primero para que las severidades conserven texto blanco legible.
        self.tree.tag_configure("readonly_row", background="#262626", foreground="#A0A0A0")
        self.tree.tag_configure("opsec_low", background="#2D2B6B", foreground="#FFFFFF")
        self.tree.tag_configure("anomaly_med", background="#7A5214", foreground="#FFFFFF")
        self.tree.tag_configure("ai_high", background="#6E1D1D", foreground="#FFFFFF")

        v_scroll = ctk.CTkScrollbar(mid_frame, orientation="vertical", command=self.tree.yview)
        h_scroll = ctk.CTkScrollbar(mid_frame, orientation="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        self.tree.grid(row=0, column=0, padx=(8, 0), pady=(8, 0), sticky="nsew")
        v_scroll.grid(row=0, column=1, padx=(2, 8), pady=(8, 0), sticky="ns")
        h_scroll.grid(row=1, column=0, padx=(8, 0), pady=(2, 8), sticky="ew")
        self.tree.bind("<<TreeviewSelect>>", self._on_row_select)

        # Zona 5: Panel inferior de edición
        edit_frame = ctk.CTkFrame(self)
        edit_frame.grid(row=4, column=0, padx=14, pady=(0, 6), sticky="ew")
        edit_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            edit_frame,
            text="Etiqueta (selecciona una fila o escríbela manualmente):",
            font=ctk.CTkFont(family=_UI_FONT_FAMILY, size=11),
        ).grid(row=0, column=0, padx=10, pady=(8, 2), sticky="w")
        ctk.CTkLabel(
            edit_frame,
            text="Nuevo Valor (dejar vacío para borrar contenido):",
            font=ctk.CTkFont(family=_UI_FONT_FAMILY, size=11),
        ).grid(row=0, column=1, padx=10, pady=(8, 2), sticky="w")

        self.entry_tag = ctk.CTkEntry(edit_frame, width=300, placeholder_text="Ej. Creator, EXIF:Artist...")
        self.entry_tag.grid(row=1, column=0, padx=10, pady=(0, 4), sticky="ew")
        self.entry_tag.bind("<KeyRelease>", lambda _e: self._sync_edit_controls_for_tag())

        self.entry_val = ctk.CTkEntry(edit_frame, placeholder_text="Valor a aplicar en la copia temporal...")
        self.entry_val.grid(row=1, column=1, padx=10, pady=(0, 4), sticky="ew")

        self.btn_apply = ctk.CTkButton(
            edit_frame,
            text="Aplicar cambio",
            command=self._on_apply_edit,
            width=155,
            fg_color="#1F618D",
            hover_color="#154360",
        )
        self.btn_apply.grid(row=1, column=2, padx=10, pady=(0, 4))

        self.lbl_edit_hint = ctk.CTkLabel(
            edit_frame,
            text="Selecciona un archivo para inspeccionar o editar sus metadatos.",
            text_color="gray",
            anchor="w",
            font=ctk.CTkFont(family=_UI_FONT_FAMILY, size=11),
        )
        self.lbl_edit_hint.grid(row=2, column=0, columnspan=3, padx=10, pady=(0, 8), sticky="ew")

        # Zona 6: Barra de estado inferior
        status_frame = ctk.CTkFrame(self, height=30)
        status_frame.grid(row=5, column=0, padx=14, pady=(0, 10), sticky="ew")
        status_frame.grid_columnconfigure(0, weight=1)

        self.lbl_status = ctk.CTkLabel(
            status_frame,
            text="Listo",
            anchor="w",
            font=ctk.CTkFont(family=_UI_FONT_FAMILY, size=11),
        )
        self.lbl_status.grid(row=0, column=0, padx=10, pady=4, sticky="ew")

        self.progress_bar = ctk.CTkProgressBar(status_frame, mode="indeterminate", width=180)
        self.progress_bar.grid(row=0, column=1, padx=10, pady=6)
        self.progress_bar.grid_remove()

    # ------------------------------------------------------------------
    # Concurrencia serializada (un único hilo trabajador activo a la vez)
    # ------------------------------------------------------------------

    def _run_in_background(
        self,
        task_fn: Callable[[], Any],
        on_success: Callable[[Any], None],
        status_text: str,
        on_error: Callable[[Exception], None] | None = None,
    ) -> None:
        if self._busy:
            return
        self._busy = True
        self._pending_on_success = on_success
        self._pending_on_error = on_error
        self.lbl_status.configure(text=status_text)
        self.progress_bar.grid()
        self.progress_bar.start()
        self._update_controls_state()

        def _worker() -> None:
            try:
                result = task_fn()
                self._result_queue.put(("ok", result))
            except Exception as exc:  # noqa: BLE001 - propagado al hilo principal
                logger.exception("Error durante operación en segundo plano: %s", status_text)
                self._result_queue.put(("err", exc))

        threading.Thread(target=_worker, daemon=True).start()
        self._poll_queue()

    def _poll_queue(self) -> None:
        # Sondea la cola únicamente mientras existe un hilo trabajador activo.
        if not self.winfo_exists():
            return
        try:
            status, payload = self._result_queue.get_nowait()
        except queue.Empty:
            if self._busy:
                self._poll_after_id = self.after(50, self._poll_queue)
            return

        self._poll_after_id = None
        on_success = self._pending_on_success
        on_error = getattr(self, "_pending_on_error", None)
        self._pending_on_success = None
        self._pending_on_error = None
        self._busy = False
        self.progress_bar.stop()
        self.progress_bar.grid_remove()
        self.lbl_status.configure(text="Listo")
        self._update_controls_state()

        if status == "ok" and on_success is not None:
            on_success(payload)
        elif status == "err":
            if on_error is not None:
                on_error(payload)
            else:
                messagebox.showerror("Error", str(payload))

    # ------------------------------------------------------------------
    # Sincronización visual y tabla filtrada en memoria
    # ------------------------------------------------------------------

    def _update_controls_state(self) -> None:
        has_file = self.session.original_path is not None
        has_pending = self.session.has_pending_changes

        if self._busy:
            for btn in (
                self.btn_open,
                self.btn_scan,
                self.btn_wipe,
                self.btn_export,
                self.btn_discard,
                self.btn_save_copy,
                self.btn_apply,
            ):
                btn.configure(state="disabled")
            self.entry_tag.configure(state="disabled")
            self.entry_val.configure(state="disabled")
            return

        self.btn_open.configure(state="normal")
        self.btn_scan.configure(state="normal" if has_file else "disabled")
        self.btn_wipe.configure(state="normal" if has_file else "disabled")
        self.btn_export.configure(state="normal" if has_file else "disabled")
        self.btn_discard.configure(state="normal" if has_pending else "disabled")
        self.btn_save_copy.configure(state="normal" if has_pending else "disabled")
        self.entry_tag.configure(state="normal" if has_file else "disabled")

        if has_pending:
            self.btn_save_copy.configure(
                text=f"Guardar copia ({self.session.pending_change_count})"
            )
        else:
            self.btn_save_copy.configure(text="Guardar copia")

        if has_file and self.session.is_pdf:
            self.pdf_banner.grid()
        else:
            self.pdf_banner.grid_remove()

        if has_file:
            fmt_label = "Office OOXML" if self.session.is_office else ("PDF" if self.session.is_pdf else "Estándar")
            pending_txt = (
                f" | Cambios pendientes: {self.session.pending_change_count}"
                if has_pending
                else " | Sin cambios pendientes"
            )
            self.lbl_meta_info.configure(
                text=f"Formato: {fmt_label}{pending_txt} | ExifTool v{self.exiftool_version}"
            )
        else:
            self.lbl_meta_info.configure(text=f"ExifTool v{self.exiftool_version}")

        self._sync_edit_controls_for_tag()

    def _sync_edit_controls_for_tag(self) -> None:
        if self._busy or not self.session.original_path:
            self.entry_val.configure(state="disabled")
            self.btn_apply.configure(state="disabled")
            return

        tag = self.entry_tag.get().strip()
        if not tag:
            self.entry_val.configure(state="normal")
            self.btn_apply.configure(state="normal")
            self.lbl_edit_hint.configure(
                text="Selecciona una fila o escribe una etiqueta editable para modificarla en la copia temporal.",
                text_color="gray",
            )
            return

        raw_val = self.session.metadata.get(tag, "")
        editable = self.session.is_tag_editable(tag, raw_val)
        if not editable:
            self.entry_val.configure(state="disabled")
            self.btn_apply.configure(state="disabled")
            self.lbl_edit_hint.configure(
                text=f"'{tag}' es de solo lectura para este archivo (no modificable).",
                text_color="#E67E22",
            )
            return

        self.entry_val.configure(state="normal")
        self.btn_apply.configure(state="normal")
        short = tag.split(":")[-1].strip()
        if self.session.is_office and short in CORE_TAG_MAP and CORE_TAG_MAP[short][1] in DATE_XML_LOCAL_NAMES:
            self.lbl_edit_hint.configure(
                text="Formato de fecha Office (W3CDTF): YYYY-MM-DDTHH:MM:SSZ (o vacío para eliminar).",
                text_color="#5DADE2",
            )
        elif self.session.is_office and short == "TotalEditTime":
            self.lbl_edit_hint.configure(
                text="Tiempo total de edición en minutos enteros (ej. 45 o '1 hour 20 minutes').",
                text_color="#5DADE2",
            )
        elif self.session.is_office and short == "RevisionNumber":
            self.lbl_edit_hint.configure(
                text="Número de revisión: entero mayor o igual a 1.",
                text_color="#5DADE2",
            )
        else:
            self.lbl_edit_hint.configure(
                text="El cambio se aplicará sobre la copia temporal; el archivo original no se modifica.",
                text_color="gray",
            )

    def _schedule_filter_refresh(self, _event: Any) -> None:
        if self._filter_after_id is not None:
            self.after_cancel(self._filter_after_id)
        self._filter_after_id = self.after(150, self._populate_table)

    def _on_clear_filter(self) -> None:
        self.entry_search.delete(0, tk.END)
        self.filter_mode.set("Todos")
        self._populate_table()

    def _populate_table(self) -> None:
        self._filter_after_id = None
        self.tree.delete(*self.tree.get_children())

        metadata = self.session.metadata
        if not metadata:
            self.lbl_count.configure(text="Mostrando 0 de 0 etiquetas")
            return

        query = self.entry_search.get().strip().lower()
        mode = self.filter_mode.get()
        total = 0
        shown = 0

        for key, val in metadata.items():
            if key == "SourceFile":
                continue
            total += 1
            val_str = str(val)
            editable = self.session.is_tag_editable(key, val)
            severity = self.session.flagged_tags.get(key)

            if mode == "Editables" and not editable:
                continue
            if mode == "Solo lectura" and editable:
                continue
            if mode == "Hallazgos" and not severity:
                continue
            if query and query not in key.lower() and query not in val_str.lower():
                continue

            status_parts: list[str] = []
            if severity:
                status_parts.append(_SEVERITY_LABELS[severity])
            if not editable:
                status_parts.append("[Solo lectura]")
            status_text = " ".join(status_parts)

            row_tags: list[str] = []
            if not editable:
                row_tags.append("readonly_row")
            if severity:
                row_tags.append(severity)

            self.tree.insert("", "end", values=(status_text, key, val_str), tags=tuple(row_tags))
            shown += 1

        self.lbl_count.configure(text=f"Mostrando {shown} de {total} etiquetas")

    def _refresh_view(self) -> None:
        if self.session.original_path:
            full_path = self.session.original_path
            basename = os.path.basename(full_path)
            display = full_path if len(full_path) <= 72 else f"...{os.sep}{basename}"
            self.lbl_filepath.configure(text=display, text_color="white")
        else:
            self.lbl_filepath.configure(text="Ningún archivo cargado...", text_color="gray")
        self._populate_table()
        self._update_controls_state()

    # ------------------------------------------------------------------
    # Manejadores de eventos de usuario
    # ------------------------------------------------------------------

    def _confirm_discard_if_pending(self, action_desc: str) -> bool:
        if not self.session.has_pending_changes:
            return True
        return bool(
            messagebox.askyesno(
                "Cambios pendientes sin guardar",
                f"Tienes {self.session.pending_change_count} cambio(s) sin exportar.\n"
                f"¿Deseas descartarlos para {action_desc}?",
            )
        )

    def _on_open_file(self) -> None:
        if not self._confirm_discard_if_pending("abrir otro archivo"):
            return
        file_path = filedialog.askopenfilename(title="Selecciona un archivo para inspeccionar")
        if not file_path:
            return

        def _after_open(_meta: dict[str, Any]) -> None:
            self.entry_tag.delete(0, tk.END)
            self.entry_val.delete(0, tk.END)
            self._on_clear_filter()
            self._refresh_view()

        self._run_in_background(
            lambda: self.session.open_file(file_path),
            _after_open,
            f"Leyendo metadatos de {os.path.basename(file_path)}...",
        )

    def _on_row_select(self, _event: Any) -> None:
        selected = self.tree.selection()
        if not selected:
            return
        _status, tag, val = self.tree.item(selected[0], "values")
        self.entry_tag.configure(state="normal")
        self.entry_tag.delete(0, tk.END)
        self.entry_tag.insert(0, tag)
        self.entry_val.configure(state="normal")
        self.entry_val.delete(0, tk.END)
        self.entry_val.insert(0, val)
        self._sync_edit_controls_for_tag()

    def _on_apply_edit(self) -> None:
        tag = self.entry_tag.get().strip()
        val = self.entry_val.get()
        if not tag:
            messagebox.showwarning("Aviso", "Debes especificar el nombre de una etiqueta.")
            return

        def _after_edit(warnings: list[str]) -> None:
            self._refresh_view()
            if warnings:
                messagebox.showwarning(
                    "Cambio aplicado con advertencias",
                    "Cambio aplicado en la copia temporal con advertencias:\n" + "\n".join(warnings[:5]),
                )
            else:
                self.lbl_status.configure(
                    text=f"Cambio aplicado en copia temporal ({tag}). Pulsa «Guardar copia» para exportar."
                )

        self._run_in_background(
            lambda: self.session.apply_edit(tag, val),
            _after_edit,
            f"Aplicando cambio a {tag}...",
        )

    def _on_run_scan(self) -> None:
        def _after_scan(result: tuple[list[dict[str, Any]], dict[str, str], dict[str, int]]) -> None:
            findings, _flagged, unicode_hits = result
            self._refresh_view()
            ReportDialog(self, findings, unicode_hits, self._on_selective_clean_from_dialog)

        self._run_in_background(
            self.session.run_scan,
            _after_scan,
            "Ejecutando escaneo forense e IA...",
        )

    def _on_selective_clean_from_dialog(self, dialog: ReportDialog) -> None:
        if self.session.is_pdf:
            confirm_pdf = messagebox.askyesno(
                "Confirmar limpieza en PDF",
                "En archivos PDF, ExifTool realiza un borrado incremental y los metadatos "
                "anteriores siguen siendo recuperables dentro del archivo binario.\n\n"
                "¿Deseas aplicar igualmente la limpieza selectiva sobre la copia temporal?",
            )
            if not confirm_pdf:
                return

        def _after_clean(cleaned_count: int) -> None:
            if dialog.winfo_exists():
                dialog.destroy()
            self._refresh_view()
            if cleaned_count == 0:
                messagebox.showinfo(
                    "Aviso",
                    "Los hallazgos actuales son solo informativos y no requieren borrado automático.",
                )
            else:
                messagebox.showinfo(
                    "Limpieza selectiva aplicada",
                    f"Se sanearon {cleaned_count} elemento(s) en la copia temporal.\n"
                    "Pulsa «Guardar copia» para exportar el archivo procesado.",
                )

        self._run_in_background(
            self.session.apply_selective_clean,
            _after_clean,
            "Aplicando limpieza selectiva...",
        )

    def _on_wipe_all(self) -> None:
        if self.session.is_pdf:
            msg = (
                "ADVERTENCIA: En archivos PDF, ExifTool realiza un borrado incremental y los "
                "metadatos anteriores pueden seguir recuperables dentro del archivo binario.\n\n"
                "¿Confirmas aplicar la limpieza total sobre la copia temporal?"
            )
        elif self.session.is_office:
            msg = (
                "Se purgarán en la copia temporal: propiedades estándar (core.xml y app.xml), "
                "fechas CreateDate/ModifyDate/LastPrinted, propiedades personalizadas (custom.xml "
                "vaciado a <Properties/>) y caracteres Unicode invisibles en nodos de texto.\n\n"
                "Nota: los metadatos del sistema de archivos (File:*), miniaturas y comentarios no se modifican.\n\n"
                "¿Deseas continuar?"
            )
        else:
            msg = (
                "¿Deseas purgar todos los metadatos editables sobre la copia temporal?\n"
                "El archivo original no se modificará."
            )

        if not messagebox.askyesno("Confirmar Limpieza Forense", msg):
            return

        def _after_wipe(warnings: list[str]) -> None:
            self._refresh_view()
            extra = ("\n\nAdvertencias:\n" + "\n".join(warnings[:5])) if warnings else ""
            messagebox.showinfo(
                "Limpieza aplicada en copia temporal",
                "Metadatos purgados en la copia temporal.\n"
                f"Pulsa «Guardar copia» para exportar el archivo procesado.{extra}",
            )

        self._run_in_background(
            self.session.apply_wipe_all,
            _after_wipe,
            "Ejecutando limpieza forense total...",
        )

    def _on_save_copy(self, target_dir: str | None = None) -> None:
        def _after_save(saved_path: str) -> None:
            self.entry_tag.delete(0, tk.END)
            self.entry_val.delete(0, tk.END)
            self._refresh_view()
            messagebox.showinfo(
                "Copia guardada",
                f"Archivo procesado guardado en:\n{saved_path}\n\n"
                "El archivo original permanece intacto y la vista regresó a su estado original.",
            )

        def _on_save_error(exc: Exception) -> None:
            if isinstance(exc, DestinationNotWritableError):
                chosen = filedialog.askdirectory(
                    title="Carpeta de solo lectura — Seleccionar carpeta de destino..."
                )
                if chosen:
                    self._on_save_copy(target_dir=chosen)
                return
            messagebox.showerror("Error al guardar copia", str(exc))

        self._run_in_background(
            lambda: self.session.save_copy(target_dir=target_dir),
            _after_save,
            "Guardando copia procesada...",
            on_error=_on_save_error,
        )

    def _on_discard_changes(self) -> None:
        if not self._confirm_discard_if_pending("volver al estado original"):
            return

        def _after_discard(_res: Any) -> None:
            self.entry_tag.delete(0, tk.END)
            self.entry_val.delete(0, tk.END)
            self._refresh_view()
            self.lbl_status.configure(text="Cambios descartados. Vista restaurada al archivo original.")

        self._run_in_background(
            self.session.discard_changes,
            _after_discard,
            "Restaurando metadatos originales...",
        )

    def _on_export_json(self) -> None:
        if not self.session.metadata:
            return
        save_path = filedialog.asksaveasfilename(
            title="Exportar metadatos a JSON",
            defaultextension=".json",
            filetypes=[("Archivos JSON", "*.json")],
        )
        if not save_path:
            return
        try:
            self.session.export_json(save_path)
            messagebox.showinfo("Exportado", f"Reporte de metadatos guardado en:\n{save_path}")
        except MetaForgeError as exc:
            messagebox.showerror("Error al exportar", str(exc))

    def _on_close(self) -> None:
        if not self._confirm_discard_if_pending("cerrar MetaForge"):
            return
        if self._poll_after_id is not None:
            try:
                self.after_cancel(self._poll_after_id)
            except tk.TclError:
                pass
        if self._filter_after_id is not None:
            try:
                self.after_cancel(self._filter_after_id)
            except tk.TclError:
                pass
        self.session.exiftool_client.terminate_active()
        self.session.cleanup()
        self.destroy()
