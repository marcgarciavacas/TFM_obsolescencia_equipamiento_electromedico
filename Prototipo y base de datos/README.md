# TFM_obsolescencia_equipamiento_electromedico
Material complementario del TFM sobre evaluación de obsolescencia del equipamiento electromédico.

## Instalación y ejecución

Se necesita Windows, Python 3 con el lanzador "py" y conexión a Internet para
instalar las bibliotecas.

1. Descarga y descomprime los archivos, conservando la estructura de carpetas,
   incluida ".streamlit".
2. En la carpeta del prototipo, ejecuta "Instalar_librerias.bat" y espera a
   que termine la instalación.
3. Ejecuta "Iniciar.bat". La aplicación se abrirá en el navegador.
4. Mantén abierta la ventana de comandos mientras utilizas la aplicación.
   Para detenerla, cierra esa ventana.

La instalación de bibliotecas solo es necesaria la primera vez.
Después, basta con ejecutar "Iniciar.bat".

## Uso del prototipo

La aplicación permite consultar el resumen de resultados, el listado de
equipos, las fichas individuales y las alertas. También permite filtrar
por familia, servicio, estado y prioridad.

Los datos se leen del archivo "Base_datos.xlsx", que debe estar en la misma
carpeta que "app.py".

## Actualización de los datos

Detén la aplicación y modifica "Base_datos.xlsx" en Excel, conservando su
estructura y sus fórmulas. Recalcula y guarda el libro antes de volver a
iniciar la aplicación.

## Análisis de sensibilidad

El archivo "Analisis_sensibilidad.xlsx" se utiliza de forma independiente
al prototipo. Incluye los escenarios y un resumen de los cambios respecto
a la configuración original.

El análisis del TFM utiliza variaciones del ±10 % en pesos y umbrales de
correctivos, y desplazamientos de ±5 puntos en los límites entre niveles.

Este libro contiene su propia copia de los datos y no se actualiza
automáticamente al modificar "Base_datos.xlsx".
