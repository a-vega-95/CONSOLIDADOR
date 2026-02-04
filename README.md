# Consolidador de Documentos PDF y WORD

Aplicación para consolidar archivos PDF y Word en un único documento DOCX ordenados numéricamente.

## Características

- ✅ Lee archivos PDF y Word (.docx, .doc)
- ✅ Ordena automáticamente por número al inicio del nombre (1.pdf, 2.docx, 3.pdf, etc)
- ✅ Consolida todo en un documento DOCX final
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

6. Haz clic en "Consolidar Documentos"

7. El archivo consolidado se guardará en la misma carpeta

## Estructura de archivos

```
CONSOLIDADOR/
├── consolidador.py              # Script principal
├── requirements.txt             # Dependencias de Python
├── instalar_dependencias.bat    # Script para instalar dependencias
├── ejecutar.bat                 # Script para ejecutar la aplicación
└── README.md                    # Este archivo
```

## Ejemplo

**Entrada:**
- 1.pdf
- 2.docx
- 3.pdf
- documento_4.pdf

**Salida:**
- consolidado.docx (contiene todo el contenido en orden: 1 → 2 → 3 → 4)

## Requisitos del Sistema

- Python 3.7 o superior
- Windows (recomendado)
- Ghostscript (para procesar PDFs)

## Solución de problemas

**Error "No module named 'pdf2image'":**
- Ejecuta: `pip install -r requirements.txt`

**Error con PDFs:**
- Instala Ghostscript: https://www.ghostscript.com/download/gsdnld.html
- Reinicia después de instalar

**Error "Document is damaged":**
- Asegúrate de que los PDFs sean válidos y no estén corrupto

## Autor

Consolidador de Documentos v1.0
