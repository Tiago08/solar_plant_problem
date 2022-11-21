# Solar Plants
En esta oportunidad, se cuenta con 2 archivos (.xlsx), donde cada uno representa la generación de energía en una planta solar de 1 día completo, dividida por inversor. La relación entre inversor y planta es de muchos a uno, es decir, una planta está relacionada a 1 o más inversores.

El desafío es poder procesar los archivos tomando en cuenta que en la realidad pueden haber 1000 archivos a procesarse.

Para cada archivo, se solicita lo siguiente:

1) Generar un gráfico, un line chart donde el eje x sea la fecha y el eje y sea el active power para finalmente guardarlo.

2) Guardar en un (.txt) la suma por día del active power, el valor máximo y mínimo del active energy y por último, la ruta donde se encuentra el grafico generado en 1).

3) Imprimir en consola la suma total del active power por día, de todas las plantas.

## Mi solución
He creado un script que te preguntará si quieres procesar todos los archivos excel que esten en la misma carpeta que el script o procesar solo uno, y hará todo los pasos establecidos anteriormente.
Primero deberá instalar las dependencias encontradas en el archivo 'requirements.txt' con el siguiente comando:
`pip install -r requirements.txt`

Luego correr el programa con:
`python script.py`
El script creara un archivo .txt y una imagen con el gráfico por cada archivo excel procesado.
