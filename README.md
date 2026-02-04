# Consolidador de Documentos PDF y WORD

Aplicación para consolidar archivos PDF y Word en un único documento DOCX ordenados numéricamente.

## Características

- ✅ Lee archivos PDF y Word (.docx, .doc)
- ✅ Ordena automáticamente por número al inicio del nombre (1.pdf, 2.docx, 3.pdf, etc)
- ✅ Consolida en DOCX, PDF o DOCX+PDF
- ✅ Añade saltos de página entre documentos
- ✅ Interfaz gráfica amigable
- ✅ Barra de progreso durante procesamiento

## Instalación

### 1. Instalar dependencias

Ejecuta el archivo `instalar_dependencias.bat` o en PowerShell:

```powershell
pip install -r requirements.txt
```

**Nota importante para Windows:** Necesitas tener instalado **Ghostscript** para procesar PDFs:
- Descarga desde: https://www.ghostscript.com/download/gsdnld.html
- Instala la versión para Windows
- Reinicia la aplicación después de instalar Ghostscript

## Uso

1. Ejecuta `ejecutar.bat` o corre desde PowerShell:
   ```powershell
   python consolidador.py
   ```

2. Haz clic en "Seleccionar Carpeta de Entrada"

3. Selecciona la carpeta que contiene los archivos a consolidar

4. Los archivos se listarán en orden numérico automáticamente

5. (Opcional) Cambia el nombre del archivo de salida

6. Elige el **formato de salida** (DOCX, PDF o DOCX+PDF)

7. Haz clic en "Consolidar Documentos"

8. El archivo consolidado se guardará en la misma carpeta

## Estructura de archivos

```
CONSOLIDADOR/
├── consolidador.py              # Punto de entrada (UI)
├── ui.py                        # Interfaz gráfica (Tkinter)
├── core.py                      # Lógica de negocio (consolidación)
├── requirements.txt             # Dependencias de Python
├── instalar_dependencias.bat    # Script para instalar dependencias
├── ejecutar.bat                 # Script para ejecutar la aplicación
├── tests/
│   └── quick_test.py            # Prueba unitaria desechable
└── README.md                    # Este archivo
```

## Ejemplo

**Entrada:**
- 1.pdf
- 2.docx
- 3.pdf
- documento_4.pdf

**Salida:**
- consolidado.docx / consolidado.pdf (según el formato elegido)

## Requisitos del Sistema

- Python 3.7 o superior
- Windows (recomendado)
- Ghostscript (para procesar PDFs)
- Microsoft Word instalado (requerido por docx2pdf para convertir Word a PDF)
- Para conversión Word → PDF: se recomienda Python 3.10–3.12 por compatibilidad de pywin32

## Solución de problemas

**Error "No module named 'pdf2image'":**
- Ejecuta: `pip install -r requirements.txt`

**Error con PDFs:**
- Instala Ghostscript: https://www.ghostscript.com/download/gsdnld.html
- Reinicia después de instalar

## Prueba unitaria desechable

Puedes ejecutar una prueba rápida para validar que la UI y la lógica de negocio se integran correctamente:

```powershell
python tests/quick_test.py
```

La prueba crea archivos temporales, ejecuta una consolidación básica y luego limpia los recursos.

**Error "Document is damaged":**
- Asegúrate de que los PDFs sean válidos y no estén corrupto

## Autor

Consolidador de Documentos v1.0
