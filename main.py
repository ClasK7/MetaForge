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

"""Punto de entrada de MetaForge con verificación previa de ExifTool y Tk."""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from metaforge import ExifToolUnavailableError
from metaforge.exiftool_client import (
    ExifToolClient,
    resolve_exiftool_path,
    verify_exiftool,
)
from metaforge.session import DocumentSession

logger = logging.getLogger("metaforge")


def _configure_logging(debug: bool) -> None:
    level = logging.DEBUG if debug else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )


def _get_project_root() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def _show_startup_error(title: str, message: str) -> None:
    """Muestra el error en consola y, si Tk está disponible, en un diálogo gráfico."""
    print(f"\n[{title}]\n{message}\n", file=sys.stderr)
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(title, message)
        root.destroy()
    except Exception:
        pass


def main(argv: list[str] | None = None) -> int:
    """Valida el entorno (ExifTool y Tkinter) e inicia la interfaz de MetaForge.

    Args:
        argv: Argumentos opcionales de línea de comandos (``--debug``).

    Returns:
        Código de salida del proceso (``0`` éxito, ``1`` error de entorno).
    """
    parser = argparse.ArgumentParser(description="MetaForge - Inspector y Editor Forense de Metadatos")
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Activa el registro técnico detallado (sin volcar valores de metadatos).",
    )
    args = parser.parse_args(argv)
    _configure_logging(args.debug)

    project_root = _get_project_root()
    binary_path = resolve_exiftool_path(project_root)

    try:
        resolved_bin, exiftool_version = verify_exiftool(binary_path)
    except ExifToolUnavailableError as exc:
        _show_startup_error("ExifTool no disponible", str(exc))
        return 1

    logger.info("MetaForge iniciado en %s (Python %s, ExifTool %s)", sys.platform, sys.version.split()[0], exiftool_version)

    # Importación diferida de UI para interceptar falta de python3-tk en Kali Linux sin traceback crudo.
    try:
        from metaforge.ui import MetaForgeApp
    except (ImportError, OSError, RuntimeError) as exc:
        hint = (
            "No se pudo cargar la interfaz gráfica (Tkinter / CustomTkinter).\n\n"
            "En Kali Linux / Debian ejecuta:\n"
            "sudo apt update && sudo apt install -y libimage-exiftool-perl python3-tk\n"
            "pip install -r requirements.txt\n\n"
            f"Detalle técnico: {exc}"
        )
        print(f"\n[Error de interfaz]\n{hint}\n", file=sys.stderr)
        return 1

    client = ExifToolClient(resolved_bin)
    session = DocumentSession(client)
    try:
        app = MetaForgeApp(session=session, exiftool_version=exiftool_version)
        app.mainloop()
    except Exception as exc:
        logger.exception("Fallo inesperado en el bucle gráfico.")
        _show_startup_error("Error al iniciar MetaForge", str(exc))
        return 1
    finally:
        client.terminate_active()
        session.cleanup()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
