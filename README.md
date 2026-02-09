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
├── data/               # Datos y exportaciones
├── docs/               # Documentación
├── logs/               # Registros de aplicación
└── main.py             # Punto de entrada

```

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

### Flujo de trabajo
1. **Crear informe**: Selecciona tipo, completa empresa, ubicación y evaluador
2. **Cargar logo**: Carga la imagen del logo de tu institución
3. **Agregar evaluados**: Registra las personas evaluadas
4. **Generar PDF**: El sistema crea el informe con logo de fondo
5. **Exportar ZIP**: Empaqueta el informe y anexos

## Dependencias principales

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
