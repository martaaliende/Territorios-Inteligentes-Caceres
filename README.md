# TFG — Agrupación de Municipios de la Provincia de Cáceres

## Visualización del notebook

Los notebooks de este repositorio no se pueden previsualizar directamente en GitHub debido a que contienen widgets interactivos.

Para **ver el notebook con las salidas** (para visualizar las tablas y mapas interactivos debes ejecutar el notebook en tu propio entorno):
(https://nbviewer.org/github/martaaliende/Territorios-Inteligentes-Caceres/blob/main/MainCodigo.ipynb)

Para **ejecutarlo** en Colab:

[![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/martaaliende/Territorios-Inteligentes-Caceres/blob/main/MainCodigo.ipynb)

## Estructura de carpetas en Google Drive

Crea esta carpeta en tu Drive y copia dentro todos los archivos del repositorio:

```
MyDrive/
└── Colab Notebooks/
    └── TFG/
        ├── MainCodigo.ipynb
        ├── crear_matrices.ipynb
        ├── descarga_datos.py
        ├── scoring_comparativa.py
        ├── DATOS_JEX.xlsx
        ├── 04-04.xlsx
        ├── 09-26_.geojson
        ├── datosgoogleapi.xlsx
        ├── comparativa_indicadores.xlsx
        ├── matriz_distancias.xlsx
        ├── matriz_tiempos.xlsx
        ├── matriz_adyacencia.xlsx
        └── municipios por MancoM y por GAL.xlsx
```

Los archivos `09-13.xlsx`, `09-25.xlsx` y `09-26.xlsx` se generan automáticamente al ejecutar el notebook.

- El notebook requiere el shapefile de límites municipales. Descárgalo desde el [Centro de Descargas del CNIG](https://centrodedescargas.cnig.es/CentroDescargas/limites-municipales-provinciales-autonomicos) y colócalo en una carpeta llamada `SHP_ETRS89` dentro de `TFG/`, con los archivos:
  - `recintos_municipales_inspire_peninbal_etrs89.*`
  - `ll_municipales_inspire_peninbal_etrs89.*`

  La estructura final de la carpeta quedará así:

  ```
  TFG/
  ├── SHP_ETRS89/
  │   ├── recintos_municipales_inspire_peninbal_etrs89.*
  │   └── ll_municipales_inspire_peninbal_etrs89.*
  ├── MainCodigo.ipynb
  └── ... (resto de archivos)
  ```
## Cómo ejecutar

### 1. Monta Google Drive en Colab

Al abrir cualquier notebook en Colab, ejecuta primero esta celda (aparece al inicio de `MainCodigo.ipynb`):

```python
from google.colab import drive
drive.mount('/content/drive')
```

Colab pedirá permiso para acceder a tu Drive, acéptalo.

### 2. Verifica la ruta base

La ruta base está definida al inicio del notebook:

```python
BASE = '/content/drive/MyDrive/Colab Notebooks/TFG'
```

Si la carpeta está guardada en otra ubicación, cambia esta línea para que apunte a donde está.

### 3. Ejecuta

Abre `MainCodigo.ipynb` en Google Colab y ejecuta las celdas **en orden de arriba a abajo** (`Entorno de ejecución → Ejecutar todo`).

---

## Descripción de los archivos

| Archivo | Descripción |
|---|---|
| `MainCodigo.ipynb` | Notebook principal, ejecutar este |
| `crear_matrices.ipynb` | Genera las matrices de distancia, tiempos y adyacencia (solo si quieres regenerarlas) |
| `descarga_datos.py` | Script de carga e importación de indicadores |
| `scoring_comparativa.py` | Lógica de scoring multicriterio y comparativa de escenarios |
| `DATOS_JEX.xlsx` | Datos socioeconómicos de la Junta de Extremadura |
| `04-04.xlsx` | Cobertura 5G por municipio |
| `09-26_.geojson` | Polígonos industriales (OpenStreetMap) |
| `datosgoogleapi.xlsx` | POIs descargados de Google Places API |
| `matriz_distancias.xlsx` | Matriz de distancias en km entre municipios |
| `matriz_tiempos.xlsx` | Matriz de tiempos en minutos entre municipios |
| `matriz_adyacencia.xlsx` | Matriz de adyacencia entre municipios |
| `municipios por MancoM y por GAL.xlsx` | Agrupaciones de referencia por mancomunidad y GAL |

---

## Notas

- La generación de las matrices (`crear_matrices.ipynb`) puede tardar varios minutos porque descarga el grafo de carreteras de la provincia de Cáceres. Si ya tienes los archivos `.xlsx` de las matrices no es necesario ejecutarlo.

- El script `descarga_datos.py` requiere conexión a internet para geocodificación inversa (Nominatim/OpenStreetMap) y opcionalmente para Google Places API. Los datos de Google Places (`datosgoogleapi.xlsx`) se descargan por separado; si ya tienes el archivo no es necesario volver a ejecutar esa parte.
