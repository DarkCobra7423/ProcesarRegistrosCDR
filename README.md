# 📊 Procesar Registros CDR (Telcel)

[![Excel VBA](https://shields.io)](https://microsoft.com)
[![KML Enabled](https://shields.io)](https://google.com)

Macro de Microsoft Excel VBA diseñada para la limpieza, normalización, clasificación y análisis estructurado de registros telefónicos (CDR) provenientes de exportaciones de **Telcel**.

---

## 📝 Descripción

La macro automatiza de extremo a extremo la preparación de una sábana telefónica para facilitar su revisión técnica y análisis forense o estadístico. Crea una copia de seguridad de los datos originales, depura filas inválidas, calcula campos derivados y genera visualizaciones estructuradas.

### 🌟 Funciones Principales
* **Validación inicial:** Comprueba la integridad de la hoja activa y la presencia de los encabezados requeridos.
* **Procesamiento seguro:** Genera una nueva hoja llamada `Procesado` sin alterar ni escribir sobre los datos originales.
* **Depuración de registros:** Filtra y elimina filas huérfanas, dejando únicamente encabezados válidos y registros con números telefónicos de 10 dígitos.
* **Clasificación inteligente:** Agrupa eventos de voz, mensajes, datos, tránsito y transferencia, aislando el sentido (entrante/saliente).
* **Normalización numérica:** Remueve automáticamente el prefijo internacional `52` en números de 12 dígitos y extrae la clave `NUMBIFETEL` (primeros 6 dígitos).
* **Campos temporales:** Calcula de forma automática el día de la semana, hora en formato militar (24h), etapa del día (madrugada, mañana, tarde, noche) y la duración exacta en minutos.
* **Limpieza de metadatos:** Reemplaza los valores de texto `"null"` del campo IMEI para evitar ruido en el análisis.
* **Geolocalización:** Consolida coordenadas de antenas (Latitud/Longitud) y las concatena con su Azimut para generar sectores direccionales.
* **Análisis dinámico:** Crea reportes dinámicos de frecuencias por contraparte y mapas de calor de conexiones por antena.
* **Exportación geográfica:** Ofrece la generación opcional de un archivo **KML** con sectores estimados para abrir directamente en Google Earth.

---

## 🛠️ Tecnología y Dependencias

* **Motor:** Microsoft Excel con soporte para Visual Basic for Applications (VBA).
* **Rendimiento:** Uso de `Scripting.Dictionary` mediante enlace tardío (*Late Binding*) para optimizar búsquedas en memoria.
* **Codificación:** Utiliza `ADODB.Stream` para forzar la escritura del archivo KML en formato `UTF-8` real.
* **Compatibilidad:** Formato estándar KML compatible con Google Earth y software GIS (QGIS, ArcGIS).

---

## 📋 Requisitos y Formato de Entrada

1. **Entorno de ejecución:** Debe ejecutarse en una versión de escritorio de Microsoft Excel compatible con macros.
2. **Formato del archivo:** El libro de trabajo debe guardarse obligatoriamente con la extensión `.xlsm`.
3. **Punto de partida:** Seleccionar la pestaña de la exportación original de Telcel antes de presionar el botón de inicio.
4. **Estructura del encabezado:** La macro busca de manera estricta que la columna **A** se llame `Telefono` y la columna **B** sea `Tipo`. El layout requiere al menos **11 columnas** activas con datos.

⚠️ **Nota importante:** Este desarrollo fue modelado exclusivamente para la estructura y nomenclatura de exportación propia de **Telcel**. No funcionará de forma nativa ni correcta con archivos de otros operadores (Movistar, AT&T, etc.) o layouts modificados.

---

## 🚀 Cómo Ejecutar

1. Abre el archivo `.xlsm` que contiene el módulo VBA o importa el componente `mdlProcesamientoTelcel` dentro del editor (Alt + F11).
2. Regresa a la interfaz de Excel y posiciónate sobre la hoja de los datos crudos.
3. Abre el menú de macros mediante **Desarrollador → Macros** (o Alt + F8) y ejecuta: `ProcesarRegistrosTelcel`.
4. Sigue las instrucciones en pantalla. Si detecta una hoja previa de un proceso anterior, confirmará si deseas sobrescribirla.
5. Al finalizar el procesamiento de datos, el sistema te preguntará si deseas exportar el archivo geoespacial KML y dónde guardarlo.

---

## 🗂️ Hojas y Resultados Generados

### 1. Hoja: `Procesado`
Es el núcleo limpio de los datos. Contiene los registros originales depurados junto con las siguientes columnas analíticas añadidas a la derecha:

| Campo | Descripción |
| :--- | :--- |
| `SUBTIPO` | Clasificación simplificada del evento, por ejemplo `LL - E`, `LL - S`, `M - E`, `M - S`, `Datos` o `Datos WiFi`. |
| `SENTIDO` | Identifica eventos entrantes (`E`) o salientes (`S`) cuando la clasificación lo permite. |
| `COLOR` | Etiqueta de color sugerida para apoyar una semaforización visual. |
| `NUMA` / `NUMB` | Números telefónicos asociados al registro, con eliminación condicional del prefijo `52`. |
| `NUMBIFETEL` | Primeros seis dígitos del número de contraparte. |
| `DIA` | Día de la semana derivado de la fecha del evento. |
| `HORA 24` | Hora normalizada en formato de 24 horas. |
| `ETAPA` | Franja del día: madrugada, mañana, tarde o noche. |
| `MIN` | Duración en minutos calculada a partir de `SEG`. |
| `BTS` | Coordenadas de la antena concatenadas como latitud y longitud. |
| `BTS / AZIMUTH` | Coordenadas concatenadas con el azimut. |

### 2. Hoja: `Frecuencias de comunicacion`
Contiene la inteligencia del negocio y resúmenes analíticos estructurados en dos bloques:
* **Tabla Dinámica de Comunicaciones:** Conteo de interacciones agrupadas por contraparte y subtipo. Cuenta con segmentadores nativos por fecha, etapa del día, día de la semana y rango horario.
* **Tabla de Conexiones por Celda:** Matriz que consolida la frecuencia de uso de las antenas, ideal para identificar zonas de mayor permanencia o movilidad del objetivo.

### 3. Archivo KML Opcional (`Antenas_CDR.kml`)
Genera un gemelo digital del comportamiento geográfico. Dibuja los nodos de comunicación y proyecta los conos de cobertura teóricos basándose en los siguientes parámetros quemados en código:

* **Radio del sector estimado:** 500 metros.
* **Apertura de haz (Beamwidth):** 120 grados.
* **Precisión de arco:** Incrementos de 10 grados para renderizado suave.

📌 *Aviso legal: Los sectores son proyecciones geométricas matemáticas ideales para facilitar la lectura visual. No representan mapas de cobertura radioeléctrica reales ni mediciones de espectro técnico.*

---

## 🗃️ Clasificación de Eventos Map

El diccionario interno mapea las siguientes etiquetas nativas de Telcel hacia la estructura normalizada:

| Tipo Original | Subtipo Asignado | Sentido |
| :--- | :--- | :--- |
| `DATOS` | `Datos` | Saliente |
| `DATOS WiFi` | `Datos WiFi` | Saliente |
| `MENSAJE ENTRANTE` | `M - E` | Entrante |
| `MENSAJE SALIENTE` | `M - S` | Saliente |
| `VOZ ENTRANTE` | `LL - E` | Entrante |
| `VOZ SALIENTE` | `LL - S` | Saliente |
| `VOZ TRANSITO` | `Transito` | *No asignado* |
| `VOZ TRANSFER` | `Transfer` | *No asignado* |

*Cualquier tipo de registro que no coincida de forma exacta con la lista anterior se mantendrá con sus campos analíticos vacíos para auditoría manual.*

---

## 🛡️ Validaciones y Manejo de Errores

El sistema cuenta con un controlador de eventos básico que interrumpe la macro de forma segura si:
* El usuario intenta correr el proceso estando posicionado sobre una hoja de resultados (`Procesado` o `Frecuencias de comunicacion`).
* El layout original no cumple con el mínimo de 11 columnas o las cabeceras `Telefono`/`Tipo`.
* Las coordenadas geográficas de origen vienen corruptas o vacías durante la generación del KML.

En caso de fallo general, saltará un cuadro de diálogo con el código nativo del error de VBA. Si la hoja `Procesado` quedó construida a medias, se recomienda borrarla manualmente antes de reintentar.

---

## 🛑 Limitaciones Técnicas

* **Formato Estricto:** Si los números de teléfono tienen caracteres extraños, espacios o formatos internacionales mal formados (distintos a 10 dígitos o 12 dígitos con prefijo `52`), podrían ser descartados en el filtro inicial.
* **Dependencia Regional de Fecha:** La conversión de horas y la asignación de la franja del día dependen enteramente de la configuración regional del sistema operativo donde corra Excel.

### Desarrollador:
- [Carlos Daniel Angel Padilla (DarkCobra7423)](https://github.com/DarkCobra7423)
