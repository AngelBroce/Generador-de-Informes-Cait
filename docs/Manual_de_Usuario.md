# Manual de Usuario — Generador de Informes CAIT

## 1. Objetivo
Este sistema permite crear informes técnicos CAIT, adjuntar evidencias y exportar el resultado en formato PDF y ZIP.

---

## 2. Requisitos
### Opción A: Uso con ejecutable (recomendado para usuarios finales)
- Descargar `CAIT_Informes_Setup_vX.Y.Z.exe` o `CAIT_Informes_Portable_vX.Y.Z.zip` desde el release.
- Si usas Setup: ejecutar el instalador y seguir el asistente.
- Si usas Portable: descomprimir el ZIP completo y ejecutar `CAIT_Informes.exe`.
- Verificar `SHA256SUMS.txt` para confirmar integridad de descarga.

> Importante: no mover el `.exe` por separado. Debe mantenerse junto con la carpeta `_internal` y el resto de archivos extraídos.

### Opción B: Uso con código fuente (usuarios técnicos)
- Python 3.11+
- Instalar dependencias con:

```bash
pip install -r requirements.txt
```

- Ejecutar:

```bash
python main.py
```

---

## 3. Estructura general de la pantalla
La aplicación se organiza en secciones principales:
- **Datos generales del informe**
- **Datos de empresa y contraparte**
- **Evaluador(es) técnico(s)**
- **Cuerpo del informe**
- **Adjuntos y exportación**

Botones principales:
- **Nuevo informe**
- **Guardar borrador**
- **Cargar borrador**
- **Generar PDF**
- **Exportar ZIP**

---

## 4. Flujo recomendado de uso
1. Abrir la aplicación.
2. Completar los datos de presentación (empresa, fecha de evaluación y contraparte).
3. En fechas, elegir modo **Una fecha** o **Varias fechas** según el caso.
4. Seleccionar tipo de informe.
5. Completar el contenido técnico.
6. Cargar adjuntos requeridos.
7. Guardar borrador (opcional).
8. Generar PDF.
9. Exportar ZIP final para entrega.

---

## 5. Tipos de informe y evaluadores
## 5.1 Informe normal
- Se utiliza un evaluador principal.

## 5.2 Informe combinado (audiometría + espirometría)
- El sistema solicita **dos evaluadores**:
  - Evaluador de Audiometría
  - Evaluador de Espirometría
- Ambos nombres se muestran en la sección **PREPARADO POR** del PDF.
- Deben seleccionarse ambos para poder generar correctamente el informe.

---

## 6. Borradores
### Guardar borrador
- Guarda el estado actual del formulario para continuar después.

### Cargar borrador
- Usar el botón **Cargar borrador**.
- Se abrirá un listado interno con los borradores disponibles.
- Seleccionar el borrador deseado para cargarlo.

---

## 7. Generación de PDF
Al generar el PDF, el sistema incluye:
- Portada y datos de presentación.
- Contenido técnico del informe.
- Gráficos y estadísticas cuando aplican.
- Firmas y sección **PREPARADO POR** según tipo de informe.

Recomendaciones:
- Verificar ortografía y datos antes de generar.
- Confirmar que evaluadores y contraparte estén correctos.

---

## 8. Exportación ZIP para entrega
La exportación ZIP agrupa:
- PDF del informe.
- Adjuntos seleccionados.
- Estructura de carpetas para compartir fácilmente.

Uso recomendado:
- Entregar este ZIP a cliente o contraparte.
- No editar manualmente la estructura interna del ZIP.

---

## 9. Adjuntos
Puedes incluir archivos de soporte como:
- Audiometrías
- Espirometrías
- Calibración
- Idoneidad
- Otros respaldos definidos por el flujo interno

Buenas prácticas:
- Usar nombres de archivo claros.
- Evitar duplicados con nombres iguales.

---

## 10. Solución de problemas
## 10.1 La app no abre en otra PC
- Si se usó Setup: reinstalar con permisos de administrador.
- Si se usó Portable: verificar que se descomprimió todo el ZIP.
- Ejecutar siempre desde la carpeta completa, no desde el `.exe` aislado.

## 10.5 El navegador bloquea la descarga
- Descargar el archivo `CAIT_Informes_Portable_vX.Y.Z.zip` en lugar del `.exe`.
- Validar el archivo con `SHA256SUMS.txt`.
- Si el navegador aplica bloqueo preventivo, conservar manualmente y luego verificar hash antes de ejecutar.

## 10.2 Faltan datos al generar PDF
- Revisar campos obligatorios.
- En informe combinado, confirmar selección de los dos evaluadores.

## 10.3 No aparece un borrador
- Confirmar que el borrador fue guardado correctamente.
- Volver a abrir **Cargar borrador** y refrescar selección.

## 10.4 Problemas con exportación ZIP
- Verificar permisos de escritura en la carpeta destino.
- Reintentar con una ruta local corta (por ejemplo, Escritorio).

---

## 11. Recomendaciones operativas
- Guardar borrador antes de generar documentos finales.
- Mantener una copia del ZIP final en respaldo.
- Estandarizar nombres de informes por fecha y cliente.

---

## 12. Soporte interno
Si se detecta un error:
1. Registrar qué acción se estaba realizando.
2. Guardar captura de pantalla del mensaje.
3. Compartir el detalle con el responsable técnico.

---

**Versión del manual:** 1.0  
**Fecha:** 25/02/2026
