# Cruce diario de uso y finalización de un servicio (3 informes → 1 hoja)

Script para Google Apps Script que cruza automáticamente tres informes Excel diarios recibidos por correo (uso del servicio, datos de cliente y datos de póliza), los combina por un identificador común y guarda el resultado en un Google Sheet histórico, enviando un correo de resumen con tabla HTML.

## Qué hace

1. **Descarga los tres informes del día anterior**: busca por correo (asunto + adjunto con un prefijo determinado) los tres Excel que llegan cada día:
   - **Usos**: registros de uso del servicio (archivo "maestro").
   - **Cliente**: datos de contacto (email, número de cuenta) de cada cliente.
   - **Asistencia**: datos de póliza (ramo, número de póliza) de cada cliente.
2. **Filtra los usos de ayer**: de entre todos los registros de uso, se queda solo con los de la fecha de ayer y cuyo valor de una columna de control coincide con un valor objetivo configurable (por ejemplo, un estado o tipo de servicio concreto).
3. **Cruza los datos**: para cada registro de uso filtrado, busca su email/cuenta en el informe de Cliente y su ramo/póliza en el informe de Asistencia, usando un identificador común ("Mapfre ID").
4. **Guarda el resultado**: añade las filas combinadas a una hoja de Google Sheets histórica (crea la hoja y las cabeceras si no existen).
5. **Envía un resumen por correo**: con el número de registros procesados y una tabla HTML con el detalle (limitada a 50 filas para no saturar el correo), o un aviso de error si falta alguno de los tres archivos.

## Requisitos previos

- Un Google Sheet donde se registra el histórico consolidado.
- Los tres informes deben llegar cada día por correo, cada uno con un asunto identificable y un adjunto Excel (`.xlsx`/`.xls`) cuyo nombre empiece por un prefijo constante.
- El servicio avanzado **Drive API** habilitado en el proyecto de Apps Script (necesario para convertir los adjuntos Excel a Google Sheets temporalmente).

## Instalación

1. Crea un proyecto nuevo en [script.google.com](https://script.google.com).
2. Copia el contenido de `Codigo.gs` en el editor.
3. Habilita el servicio avanzado **Drive API**: Editor → Servicios (icono "+") → busca "Drive API" → Añadir.
4. Rellena el objeto `CONFIG` al principio del archivo:
   - `EMAIL_DESTINO` / `EMAIL_CC`: destinatarios del correo de resumen y de errores.
   - `SHEET_DESTINO_ID`: ID del Google Sheet donde se guarda el histórico.
   - En cada uno de `FILES.USOS`, `FILES.CLIENTE` y `FILES.ASISTENCIA`: el `SUBJECT` (asunto del correo) y `PREFIX` (prefijo del nombre del adjunto) reales de tus informes.
5. Ajusta los índices de columna (`COL_..._ORIGINAL`) y `SKIP_ROWS` (número de filas de cabecera a saltar) de cada informe a la estructura real de tus Excel — estos valores dependen de cómo esté maquetado cada informe y son lo primero que hay que revisar si el cruce no encuentra datos.
6. En `FILES.USOS`, ajusta `TARGET_VAL` al valor que debe tener la columna de control para que una fila se considere válida (por ejemplo, el nombre de un servicio o un estado "OK").
7. La primera vez que ejecutes la función, Google te pedirá autorizar permisos de Gmail, Sheets y Drive.

## Activador (ejecución automática)

En el editor de Apps Script, ve a **Activadores** (icono de reloj) y crea uno para `procesarCruceDiario`, con periodicidad diaria (después de la hora a la que sueles recibir los tres informes). También puedes ejecutarlo manualmente con el botón ▶️ para probarlo.

## Notas

- El código incluido en este repositorio usa **valores de ejemplo** (`TU_SHEET_ID_DESTINO`, `tu_correo@tudominio.com`, asuntos y prefijos de adjunto, etc.) en lugar de los datos reales del entorno original. Sustitúyelos por los tuyos antes de ejecutar.
- El script incluye unos logs de depuración (las primeras 5 filas con ID) para poder comprobar rápidamente si la fecha y el valor de control se están interpretando bien — muy útil la primera vez que lo configures con tu propio Excel.
- Si un cliente no aparece en el informe de Cliente o de Asistencia, el script no falla: rellena esos campos con "No encontrado" y sigue con el resto de registros.
