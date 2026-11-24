# MetaForge 

MetaForge es una herramienta forense de código abierto construida en Python para inspeccionar, inyectar y purgar metadatos de imágenes, documentos de Office (.docx, .xlsx, .pptx) y multimedia. Incluye un escáner heurístico para detectar anomalías OPSEC y rastros de herramientas de Inteligencia Artificial (C2PA, metadatos generativos, y caracteres Unicode invisibles).

![Interfaz de MetaForge](assets/Cap_metaforge.jpg)

## Requisitos Previos

* **Python 3.8+** instalado en el sistema.
* **ExifTool**: 
  * En **Windows**, el ejecutable ya está incluido en la carpeta `bin/`.
  * En **Linux/Debian/Kali**, debes instalarlo a nivel de sistema:
    ```bash
    sudo apt update && sudo apt install libimage-exiftool-perl
    ```

## Instalación y Ejecución

1. **Clonar el repositorio:**
   ```bash
   git clone [https://github.com/ClasK7/MetaForge.git](https://github.com/ClasK7/MetaForge.git)
   cd MetaForge
2. **Crear y activar un entorno virtual:**

   * **En Windows (PowerShell/CMD):**
     ```cmd
     python -m venv venv
     .\venv\Scripts\activate
     ```

   * **En Linux:**
     ```bash
     python3 -m venv venv
     source venv/bin/activate
     ```

3. **Instalar dependencias:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Lanzar la aplicación:**
   ```bash
   python main.py
   ```