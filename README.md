# MyAdmon — Partes horarios (versión de escritorio)

Una aplicación de calendario inspirada en el flujo de partes de AITEX. Está desarrollada con **Python 3.10+**, **Tkinter** y **SQLite**, todas herramientas incluidas con Python en Windows. **No necesita pip.** Solo el asistente Copilot opcional necesita Internet y GitHub Copilot CLI.

## Inicio rápido en Windows

1. Clona este repositorio o descarga y descomprime su ZIP en una carpeta permanente.
2. Instala Python 3.10 o superior desde python.org si no lo tienes, marcando la opción **Add Python to PATH**.
3. Ejecuta `INICIAR.bat` (doble clic). También puedes abrir una terminal en la carpeta y escribir `py -3 partes_aitex.py`.
4. Se abrirá una ventana con el calendario de partes vacío. Pulsa **Importar Excel original** y selecciona tu libro (por ejemplo `Horas MyAdmon.xlsx`) para cargar los registros.

## Funcionalidades

- Calendario mensual con colores por proyecto, totales diarios y mensuales, navegación entre meses y selector de proyecto.
- Selección de un día, alta de partes, edición de tareas, eliminación con confirmación y lista de partes del día.
- Proyectos nuevos y categorías de actividad editables.
- Pendientes de completar o sin fecha (incluye actividades importadas con horas faltantes).
- Exportación a un libro `xlsx` con **Registros** y **Resumen**. Guarda las horas como `HH:MM`, y también en horas decimales correctas, más el valor original del Excel cuando se importó.
- Importación del formato original del Excel (**una pestaña por proyecto**). Los registros ya importados de la misma pestaña y fila se omiten para evitar reintroducirlos.
- **Escritura en el Excel importado**: cada parte que crees, edites o borres en la app se escribe en las casillas de la pestaña de su proyecto (fecha, mes, tarea y horas, con el formato de la pestaña). Los proyectos nuevos crean su propia pestaña. Si el Excel está abierto, los cambios quedan pendientes y se escriben con **Escribir en Excel ahora** tras cerrarlo.
- Base de datos SQLite persistente y copias de seguridad bajo demanda.
- **Asistente Copilot** (panel lateral derecho, se oculta con el botón de la cabecera): propone qué poner en el parte del día seleccionado a partir de tus últimos partes del proyecto y de lo ya imputado ese día. Con **Crear parte con la última sugerencia** se abre el formulario ya relleno. Requiere [GitHub Copilot CLI](https://docs.github.com/copilot/how-tos/copilot-cli) (`npm install -g @github/copilot`) y una cuenta con Copilot; el asistente no tiene acceso a archivos ni herramientas, solo recibe ese contexto.

## Dónde se guardan los datos

En Windows, la base de datos está en:

`%LOCALAPPDATA%\MyAdmonPartes\partes.sqlite3`

La base de datos se crea vacía en el primer inicio; los registros se cargan con **Importar Excel original**. Los datos se almacenan fuera de la carpeta del programa: puedes actualizar la aplicación sin perderlos **siempre que no borres la carpeta de datos de usuario**.

Haz una copia periódica con el botón **Copia de seguridad** y guarda el archivo en otra ubicación. Antes de cada escritura en el Excel se guarda una copia del libro en `%LOCALAPPDATA%\MyAdmonPartes\copias_excel` (se conservan las 20 últimas). Las horas se escriben con la convención del libro: `8.15` = 08:15, `8.5` = 08:30 y `8.45` = 08:45.

## Importante: interpretación de horas del Excel

El libro contiene mezclados valores como `8.5`, `8.15`, `8.45` y `9.27`. Para obtener totales reales sin confundir bases 10 y 60, **la importación inicial aplica esta hipótesis**:

| Excel | Interpretación adoptada |
|---|---|
| 8.5 | 08:30 (media hora) |
| 8.15 | 08:15 |
| 8.3 | 08:30 |
| 8.45 | 08:45 |
| 9.27 | 09:27, **marcado para revisar** |

Esta interpretación **no puede garantizarse para todas las anotaciones históricas** porque el libro no documenta la convención. Se conserva en la base de datos la cifra original junto con la normalizada para poder auditarlo. Los registros nuevos siempre se introducen por separado como **horas y minutos**.

El repositorio no incluye históricos personales, bases de datos, libros Excel ni capturas con partes reales. Importa tu propio Excel desde la aplicación.

## Límites de esta versión

- Esta herramienta es de **gestión personal**, no es una integración con la plataforma oficial de AITEX. No inicia sesión allí ni presenta partes en su nombre.
- Las ideas del asistente son orientativas: revisa siempre el texto y las horas antes de guardar el parte.
- No incluye control de festivos autonómicos, convenio, ausencias, jornadas teóricas ni validaciones corporativas: requieren reglas internas verificadas.
- Un registro histórico procedente de una hoja con día/mes pero sin año se ha fechado provisionalmente en **2024**. Comprueba esos registros si tu libro original abarca otros años.
- El botón **Duplicar** abre una nueva actividad precargada desde la seleccionada. Antes de guardarla podrás cambiar la fecha y el tiempo.
- Una importación posterior desde el mismo Excel no sobrescribe cambios hechos en la aplicación; las filas que escribe la app ya quedan vinculadas y no se duplican. Los cambios que hagas a mano en filas ya importadas no se trasladan a la app.

## Para ejecutar el autodiagnóstico

Consola: `py -3 comprobar.py`

Comprueba la base inicial y la importación/exportación a Excel sin necesidad de interfaz gráfica. Si no existe el histórico local `data/horas_iniciales.json`, utiliza datos de prueba sintéticos en una carpeta temporal.
