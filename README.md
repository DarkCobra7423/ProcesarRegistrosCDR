# ProcesarRegistrosCDR

Macro de Microsoft Excel VBA para la limpieza, normalización,
clasificación y análisis estructurado de registros telefónicos (CDR)
provenientes de exportaciones de Telcel.

## Descripción

La macro automatiza parte de la preparación de una sábana telefónica
para facilitar su revisión y análisis. Crea una copia de los datos
originales, depura filas que no cumplen el formato telefónico esperado,
agrega campos derivados y genera tablas de resultados.

Sus funciones principales son:

-   Validar la hoja activa y comprobar que existan los encabezados
    esperados.
-   Crear la hoja `Procesado` sin modificar directamente la hoja
    original.
-   Eliminar filas que no correspondan al encabezado o a registros cuyo
    campo de teléfono tenga 10 dígitos.
-   Clasificar eventos de voz, mensajes, datos, tránsito y
    transferencia.
-   Separar el sentido entrante y saliente cuando corresponde.
-   Retirar el prefijo `52` de números que tengan 12 dígitos y comiencen
    con ese prefijo.
-   Extraer los primeros seis dígitos de la contraparte como campo
    `NUMBIFETEL`.
-   Calcular el día de la semana, la hora en formato de 24 horas, la
    etapa del día y la duración en minutos.
-   Limpiar el valor textual `null` del campo IMEI.
-   Consolidar coordenadas de antenas y combinarlas con el azimut.
-   Crear la hoja `Frecuencias de comunicacion` con una tabla dinámica
    de comunicaciones y una tabla de conexiones agrupadas por
    coordenadas y azimut.
-   Ofrecer la generación opcional de un archivo KML con antenas y
    sectores direccionales estimados.

## Tecnología

-   Microsoft Excel.
-   Visual Basic for Applications (VBA).
-   `Scripting.Dictionary` mediante enlace tardío.
-   `ADODB.Stream` para escribir el archivo KML en UTF-8.
-   Formato KML para visualizar datos geográficos en aplicaciones
    compatibles, como Google Earth.

## Requisitos y formato de entrada

1.  Abrir el archivo en una versión de escritorio de Microsoft Excel que
    permita ejecutar macros.
2.  Guardar el libro de trabajo en formato habilitado para macros
    (`.xlsm`).
3.  Seleccionar la hoja que contiene la exportación original de Telcel
    antes de ejecutar la macro.
4.  La macro busca una fila de encabezados en la que la columna A sea
    `Telefono` y la columna B sea `Tipo`.
5.  Se esperan al menos 11 columnas en la fila de encabezados y al menos
    una fila de datos debajo.
6.  El campo de teléfono de la columna A debe contener el encabezado
    `Telefono` o un número de 10 dígitos para que la fila se conserve
    durante la limpieza inicial.

El código está preparado para una estructura concreta de exportación de
Telcel. No debe asumirse que funcionará sin ajustes con archivos de
otros operadores o con exportaciones cuyo esquema haya cambiado.

## Cómo ejecutar

1.  Abre el archivo `.xlsm` que contiene el módulo VBA.
2.  Importa el módulo `mdlProcesamientoTelcel` al editor de VBA si aún
    no está incluido en el libro.
3.  Vuelve a Excel y selecciona la hoja que contiene los datos
    originales.
4.  Ejecuta la macro `ProcesarRegistrosTelcel` desde **Desarrollador →
    Macros** o desde el editor de VBA.
5.  Lee los mensajes de validación y confirma si deseas reemplazar la
    hoja `Procesado` cuando ya exista.
6.  Al terminar el procesamiento, responde si deseas generar el archivo
    KML y elige su ubicación de guardado.

**Importante:** la macro puede eliminar hojas de resultados existentes.
Revisa los avisos antes de confirmar y conserva una copia de respaldo
del libro original.

## Hojas y resultados generados

### `Procesado`

Copia de la exportación original con filas iniciales ajenas a la tabla
eliminadas y columnas de trabajo agregadas.

Entre los campos calculados o reorganizados se encuentran:

  -----------------------------------------------------------------------
  Campo                               Descripción
  ----------------------------------- -----------------------------------
  `SUBTIPO`                           Clasificación simplificada del
                                      evento, por ejemplo `LL - E`,
                                      `LL - S`, `M - E`, `M - S`, `Datos`
                                      o `Datos WiFi`.

  `SENTIDO`                           Identifica eventos entrantes (`E`)
                                      o salientes (`S`) cuando la
                                      clasificación lo permite.

  `COLOR`                             Etiqueta de color sugerida para
                                      apoyar una semaforización visual.

  `NUMA` / `NUMB`                     Números telefónicos asociados al
                                      registro, con eliminación
                                      condicional del prefijo `52`.

  `NUMBIFETEL`                        Primeros seis dígitos del número de
                                      contraparte.

  `DIA`                               Día de la semana derivado de la
                                      fecha del evento.

  `HORA 24`                           Hora normalizada en formato de 24
                                      horas.

  `ETAPA`                             Franja del día: madrugada, mañana,
                                      tarde o noche.

  `MIN`                               Duración en minutos calculada a
                                      partir de `SEG`.

  `BTS`                               Coordenadas de la antena
                                      concatenadas como latitud y
                                      longitud.

  `BTS / AZIMUTH`                     Coordenadas concatenadas con el
                                      azimut.
  -----------------------------------------------------------------------

El módulo también conserva y utiliza otros campos de la exportación,
como fecha, hora original, duración, IMEI, red, proveedor de servicios,
municipio, estado y azimut.

### `Frecuencias de comunicacion`

Contiene resultados de análisis generados por la macro:

-   **Tabla dinámica de comunicaciones:** cuenta eventos por número de
    contraparte y subtipo. Incluye filtros por fecha, etapa del día, día
    de la semana y hora.
-   **Tabla de conexiones:** agrupa los registros por coordenadas y
    azimut y calcula la frecuencia de aparición de cada combinación.

La tabla dinámica se configura para mostrar llamadas y mensajes
entrantes y salientes. La tabla de conexiones considera los subtipos de
llamadas, mensajes y datos definidos en el código.

### Archivo KML opcional

Si se solicita, la macro permite guardar `Antenas_CDR.kml`. El archivo
incluye marcadores para las antenas interpretadas correctamente y,
cuando hay un azimut numérico, puede añadir sectores direccionales
estimados.

Parámetros definidos actualmente en el código:

-   Radio del sector: **500 metros**.
-   Apertura del sector: **120 grados**.
-   Incremento angular para construir el contorno: **10 grados**.
-   El KML se escribe en codificación UTF-8.

Los sectores son representaciones geométricas calculadas a partir de las
coordenadas y el azimut de entrada; no deben interpretarse por sí solos
como mediciones exactas de cobertura radioeléctrica.

## Clasificación de eventos

El código contempla las siguientes etiquetas originales de Telcel:

  Tipo original        Subtipo asignado   Sentido
  -------------------- ------------------ -------------
  `DATOS`              `Datos`            Saliente
  `DATOS WiFi`         `Datos WiFi`       Saliente
  `MENSAJE ENTRANTE`   `M - E`            Entrante
  `MENSAJE SALIENTE`   `M - S`            Saliente
  `VOZ ENTRANTE`       `LL - E`           Entrante
  `VOZ SALIENTE`       `LL - S`           Saliente
  `VOZ TRANSITO`       `Transito`         No asignado
  `VOZ TRANSFER`       `Transfer`         No asignado

Los tipos que no coincidan exactamente con las cadenas contempladas
quedan sin subtipo, sentido ni color asignados. Si cambia la
nomenclatura de la exportación, habrá que actualizar la función de
clasificación.

## Validaciones y manejo de errores

La macro comprueba, entre otras cosas:

-   Que la hoja activa no sea una de las hojas de salida.
-   Que se localicen los encabezados esperados.
-   Que haya al menos 11 columnas y filas de datos.
-   Que queden registros después de la limpieza inicial.
-   Que las coordenadas puedan interpretarse para generar el KML.

Si ocurre un error durante el procesamiento, se muestra un mensaje con
la descripción y el código del error. La hoja `Procesado` podría quedar
parcialmente transformada; el código recomienda eliminarla manualmente
antes de volver a ejecutar la macro.

## Limitaciones y consideraciones técnicas

-   **Compatibilidad de entrada:** está diseñada para una estructura
    específica de exportación de Telcel; no es un importador CDR
    universal.
-   **Formato de teléfono:** la limpieza inicial conserva filas cuyo
    campo de teléfono sea el encabezado esperado o un valor numérico de
    10 dígitos. Los formatos distintos pueden excluirse.
-   **Prefijo `52`:** se elimina únicamente cuando el número tiene 12
    caracteres y comienza con `52`.
-   **Fechas y horas:** la interpretación depende de que Excel/VBA
    reconozca los valores de entrada. Conviene validar muestras con
    formatos atípicos.
-   **Día de la semana:** se calcula con `Format(..., "dddd")`, por lo
    que el idioma del resultado puede depender de la configuración
    regional del sistema.
-   **Prefijo de seis dígitos:** `NUMBIFETEL` se obtiene tomando los
    primeros seis caracteres de `NUMB`; el código no consulta una base
    externa para verificar operador o región.
-   **Coordenadas:** la conversión a decimal supone coordenadas en
    formato de grados, minutos y segundos (DMS) que la función pueda
    interpretar. Las coordenadas inválidas se omiten del KML.
-   **Sectores KML:** el radio y la apertura son constantes configuradas
    en el código, no parámetros medidos automáticamente.
-   **Restauración de Excel:** la rutina de optimización desactiva
    temporalmente actualización de pantalla, eventos y cálculo
    automático. Si se modifica el código, conviene asegurar que esos
    valores se restauren también ante errores.
-   **Hoja de frecuencias:** el código elimina una hoja existente con el
    nombre `Frecuencias de comunicacion` antes de crearla de nuevo.
    Guarda los resultados previos si necesitas conservarlos.
-   **Privacidad:** los CDR pueden contener información personal y datos
    sensibles. Trabaja con autorización, limita el acceso, protege el
    archivo y evita compartir datos reales en repositorios o
    documentación pública.

## Estructura del módulo

El código se encuentra en el módulo `mdlProcesamientoTelcel`. Sus
procedimientos principales son:

-   `ProcesarRegistrosTelcel`: coordina la validación, el procesamiento
    y la generación de resultados.
-   `ValidacionPrevia`: revisa la hoja y la estructura de entrada.
-   `PrepararHojaProcesado`: copia y depura los datos.
-   `InsertarColumnasDeTrabajo`: agrega las columnas calculadas.
-   `EnriquecerDatos`: normaliza y completa los registros.
-   `ClasificarTipo`: asigna subtipo, sentido y color.
-   `CrearHojaFrecuencias`: construye la hoja de resultados.
-   `CrearPivotComunicaciones`: crea la tabla dinámica de
    comunicaciones.
-   `GenerarTablaConexiones`: agrupa coordenadas y azimut.
-   `GenerarKML`: exporta antenas y sectores a KML.

## Recomendaciones antes de usar en producción

1.  Prueba la macro con una copia de una exportación y conserva intacto
    el archivo de origen.
2.  Verifica manualmente una muestra de números, fechas, horas,
    duraciones y coordenadas antes y después del proceso.
3.  Confirma que las etiquetas de eventos de la exportación coincidan
    exactamente con las previstas por la función `ClasificarTipo`.
4.  Valida las coordenadas y los azimuts en una herramienta geográfica
    antes de interpretar el KML.
5.  Mantén una versión de respaldo del libro y del módulo VBA antes de
    modificar la lógica.
6.  No uses las frecuencias ni los sectores generados como conclusión
    automática: requieren contexto, revisión y corroboración.

## Uso responsable

Esta herramienta organiza registros para facilitar su revisión. La
generación de frecuencias, coordenadas o sectores no demuestra por sí
misma la identidad de una persona, su ubicación exacta ni la comisión de
una conducta. Las interpretaciones deben realizarse por personal
autorizado, conforme a los procedimientos aplicables y con validación de
los datos.

------------------------------------------------------------------------

**Nombre sugerido del proyecto:** `ProcesarRegistrosCDR`\
**Nombre actual de la macro en el código:** `ProcesarRegistrosTelcel`\
**Módulo VBA:** `mdlProcesamientoTelcel`\
**Formato de libro recomendado:** `.xlsm`

### Desarrollador:
- [Carlos Daniel Angel Padilla (DarkCobra7423)](https://github.com/DarkCobra7423)