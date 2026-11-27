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

"""Paquete principal de MetaForge y jerarquía unificada de excepciones."""

__version__ = "2.0.0"


class MetaForgeError(Exception):
    """Excepción base para todos los errores controlados de MetaForge."""


class ExifToolUnavailableError(MetaForgeError):
    """Se eleva cuando ExifTool no está instalado o no responde al arranque."""


class DestinationNotWritableError(MetaForgeError):
    """Se eleva cuando la carpeta de destino no tiene permisos de escritura."""


class ValidationError(MetaForgeError):
    """Se eleva ante etiquetas, fechas, valores o rutas de salida inválidos."""


class FileProcessingError(MetaForgeError):
    """Se eleva ante fallos de lectura, escritura, ZIP/XML o falta de espacio."""


__all__ = [
    "__version__",
    "MetaForgeError",
    "ExifToolUnavailableError",
    "DestinationNotWritableError",
    "ValidationError",
    "FileProcessingError",
]
