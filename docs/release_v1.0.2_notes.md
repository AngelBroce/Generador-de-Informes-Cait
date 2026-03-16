v1.0.2 - Distribución oficial por instalador.

- Descarga e instalación mediante CAIT_Informes_Setup.exe.
- Se elimina el paquete ZIP antiguo del release.
- Mejoras responsive y estabilidad de inicio en entorno instalado.

Instalación:
1) Descargar CAIT_Informes_Setup.exe
2) Ejecutar el asistente
3) Abrir CAIT Informes desde el acceso directo

v1.0.3 - Correcciones de usabilidad y mantenimiento.

- Corregido el flujo para eliminar evaluadores desde el formulario principal.
- Mejorada la legibilidad general en pantallas pequeñas aumentando la escala visual.

Actualización:
1) Cerrar CAIT Informes si está abierto.
2) Ejecutar el nuevo instalador CAIT_Informes_Setup.exe (v1.0.3).
3) Completar el asistente de instalación para reemplazar la versión anterior.
4) Abrir la aplicación y verificar la eliminación de evaluadores y el tamaño de letra.

v1.0.4 - Mejora de accesibilidad visual.

- Aumentado el tamaño de letra y la escala visual general de la interfaz.
- Se mantiene la estructura y distribución de pantallas existente.
- Versionado del instalador centralizado para reflejar correctamente la versión en Windows.

Actualización:
1) Cerrar CAIT Informes si está abierto.
2) Ejecutar el nuevo instalador CAIT_Informes_Setup.exe (v1.0.4).
3) Completar el asistente para reemplazar la versión anterior.
4) Verificar que en "Aplicaciones instaladas" aparezca la versión 1.0.4.

v1.0.6 - Fechas flexibles y distribución de release reforzada.

- Nuevo selector para elegir "Una fecha" o "Varias fechas" en fecha de evaluación.
- Nuevo selector para elegir "Una fecha" o "Varias fechas" en fechas del estudio.
- Los borradores ahora guardan y restauran el modo de fechas seleccionado.
- Se agrega script `scripts/prepare_release_assets.bat` para generar:
	- instalador versionado,
	- ZIP portable versionado,
	- archivo `SHA256SUMS.txt`.
- Documentación actualizada para descargar desde Releases usando assets que reducen bloqueos del navegador.

Actualización:
1) Cerrar CAIT Informes si está abierto.
2) Instalar `CAIT_Informes_Setup_v1.0.6.exe` o usar `CAIT_Informes_Portable_v1.0.6.zip`.
3) Si usas descarga web, validar hash con `SHA256SUMS.txt` antes de ejecutar.
4) Verificar en la pantalla de presentación que puedes alternar entre "Una fecha" y "Varias fechas".

v1.0.7 - Corrección de distribución visual y recompilación de instalador.

- Se elimina la franja de estado inferior para que el contenido use todo el alto de la ventana.
- Se recompila el instalador con los cambios recientes de fechas (una/varias).

Actualización:
1) Cerrar CAIT Informes.
2) Instalar `CAIT_Informes_Setup_v1.0.7.exe` o usar `CAIT_Informes_Portable_v1.0.7.zip`.
3) Verificar en Presentación del informe el selector de fechas y que el contenido llega hasta abajo.
