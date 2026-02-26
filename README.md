# Generador de Informes Clínicos

Aplicación para la elaboración automatizada de informes clínicos de **audiometría** y **espirometría**.

## Descripción

Esta aplicación permite a profesionales de salud del Centro de Atención Integral Terapéutico Panamá (CAIT Panamá) digitalizar y estandarizar el proceso de generación de informes clínicos, reduciendo tiempo de elaboración y evitando errores de transcripción.

## Características principales

✅ Crear y editar informes de audiometría y espirometría  
✅ Cargar logo de la empresa  
✅ Registrar datos de personas evaluadas  
✅ Adjuntar documentos relacionados (PDF)  
✅ Generar informes en PDF con logo de fondo opaco  
✅ Exportar paquetes ZIP con informe y anexos  
✅ Interfaz gráfica moderna con colores verde naturaleza  
✅ Entrada manual de datos - sin base de datos  

## Estructura del proyecto

```
practica profecional/
├── src/
│   ├── core/           # Lógica central
│   ├── models/         # Modelos de datos
│   ├── services/       # Servicios (PDF, ZIP, archivos)
│   ├── utils/          # Funciones auxiliares
│   ├── ui/             # Interfaz gráfica
│   ├── templates/      # Plantillas de informes
│   └── assets/         # Logos e imágenes
├── tests/              # Pruebas unitarias
├── config/             # Configuración
├── data/               # Datos, anexos y exportaciones
│   ├── attachments/    # Adjuntos de prueba y anexos
│   ├── databases/      # Catálogos (ej. evaluadores)
│   ├── exports/        # PDFs/ZIPs generados
│   └── reports/        # Informes generados
├── scripts/            # Scripts de demostración
├── docs/               # Documentación
├── logs/               # Registros de aplicación
├── build/              # Salida de PyInstaller
├── CAIT_Informes.spec  # Configuración del ejecutable
└── main.py             # Punto de entrada

```

## Requisitos

- Python 3.x con `pip`
- `tkinter` habilitado (incluido en la mayoría de instalaciones de Python en Windows)

## Instalación

1. Clonar el repositorio
2. Instalar dependencias:
```bash
pip install -r requirements.txt
```

## Uso

Ejecutar la aplicación:
```bash
python main.py
```

## Descarga de la aplicación

Puedes descargar la aplicación lista para usar en el ZIP del release:

- https://github.com/AngelBroce/Generador-de-Informes-Cait/raw/main/resources/CAIT_Informes.zip

### Flujo de trabajo
1. **Crear informe**: Selecciona tipo, completa empresa, ubicación y evaluador
2. **Cargar logo**: Carga la imagen del logo de tu institución
3. **Agregar evaluados**: Registra las personas evaluadas
4. **Generar PDF**: El sistema crea el informe con logo de fondo
5. **Exportar ZIP**: Empaqueta el informe y anexos

## Dependencias

Instaladas desde [requirements.txt](requirements.txt):

- reportlab==4.0.9
- PyPDF2==3.0.1
- Pillow==11.0.0
- pyinstaller==6.18.0
- pytest==7.4.3
- pytest-cov==4.1.0
- tkcalendar==1.6.1
- matplotlib==3.9.2
- seaborn==0.13.2

## Scripts útiles

- [scripts/demo_export_zip.py](scripts/demo_export_zip.py): genera un ZIP de ejemplo con anexos y certificados.
- [scripts/generate_long_pdf.py](scripts/generate_long_pdf.py): crea un PDF extenso de demostración.
- [scripts/generate_protocol_test_pdf.py](scripts/generate_protocol_test_pdf.py): prueba el orden de anexos del protocolo.

## Pruebas

Tipos de pruebas contempladas:

- Unitarias: [tests/unit/](tests/unit/)
- Integracion: [tests/integration/](tests/integration/)

Nota: actualmente las carpetas estan preparadas, pero no incluyen casos automatizados.

```bash
pytest
```

Cobertura:
```bash
pytest --cov=src --cov-report=term-missing
```

## Construcción del ejecutable

```bash
pyinstaller CAIT_Informes.spec
```

- **reportlab** - Generación de PDFs profesionales
- **PyPDF2** - Manipulación de PDFs
- **Pillow** - Procesamiento de imágenes y watermarks
- **tkinter** - Interfaz gráfica moderna

## Características del PDF

- Encabezado institucional
- Información general del informe
- Lista de personas evaluadas
- Conclusiones y recomendaciones
- Logo de fondo opaco en todas las páginas
- Pie de página con fecha y número de página

## Autor

Angel  Broce 

## Licencia

Todos los derechos reservados © 2026
