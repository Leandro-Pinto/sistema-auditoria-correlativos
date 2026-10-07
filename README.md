# Sistema de Auditoría de Correlativos – Documentos de Gestión

Programa de escritorio para Windows que audita la numeración correlativa de documentos
de gestión (resoluciones, oficios, etc.) a partir de un Excel con las columnas
`id`, `Tipo de norma` y `Título` (ej. `045-2024-GOREMAD/GR`).

## Qué hace
- Lee el número y el año desde el título de cada documento.
- Detecta números **faltantes**, **repetidos**, títulos **sin número**, **sin año** o con **formato** irregular.
- Permite corregir registros; las correcciones se guardan en `<archivo>.correcciones.json`.
- Exporta un **reporte Excel** (resumen, faltantes, por revisar, una hoja por año).
- Exporta un **informe PDF** en A4 horizontal con gráficos, hallazgos por año, nivel de riesgo,
  conclusiones, recomendaciones y firmas.

## Requisitos
- Windows 10/11
- Python 3.9 o superior
- Librerías: `openpyxl`, `reportlab` (ver `requirements.txt`)

## Uso
1. Doble clic en `Iniciar_Sistema.bat` (instala las librerías la primera vez).
2. «Abrir Excel…» y elegir el archivo a auditar.

Para crear un ejecutable `.exe` que no necesita Python: doble clic en `Crear_EXE.bat`
(el resultado queda en `dist\SistemaAuditoria.exe`).

## Archivos
| Archivo | Contenido |
|---|---|
| `SistemaAuditoria.py` | Interfaz gráfica (Tkinter) |
| `auditoria_core.py` | Lógica de auditoría y exportación a Excel |
| `reporte_pdf.py` | Informe de auditoría en PDF |
| `Iniciar_Sistema.bat` | Lanzador para Windows |
| `Crear_EXE.bat` | Genera el ejecutable con PyInstaller |

## Estados
CONFIRMADO · REPETIDO · FALTA · SIN NÚMERO · SIN AÑO · AÑO INCOMPLETO · AÑO INFERIDO · FORMATO · CORREGIDO
