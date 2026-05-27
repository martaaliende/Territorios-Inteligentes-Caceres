"""
DESCARGA E IMPORTACIÓN DE DATOS PARA EL SCORING

ESTE CÓDIGO GENERA LOS FICHEROS:
  09-13.xlsx        ← Espacios innovación/talento/emprendimiento (inventario propio)
  09-25.xlsx        ← Oficinas Acelera Pyme Rural (inventario propio)
  09-26.xlsx        ← Polígonos industriales (desde 09-26_.geojson via Nominatim)
  datosgoogleapi.xlsx ← POIs Google Places API (o carga si ya existe)

FICHEROS QUE YA TENGO EN Drive:
  09-24_.geojson    ← Descargado desde overpass-turbo.eu con código ISO3166-2=ES-CC
  DATOS_JEX.xlsx
  04-04.xlsx

USO EN COLAB:
  BASE = 'ruta donde se encuentre este fichero'
  exec(open(f'{BASE}/descarga_datos.py').read())
  # Resultado: variable DATOS_CARGADOS disponible para scoring_agrupaciones.py
"""

import os, re, json, time
import numpy as np
import pandas as pd
import requests
import warnings
warnings.filterwarnings('ignore')

# 0. RUTAS

if 'BASE' not in dir():
    BASE = input('Introduce la ruta base del proyecto (ej: /content/drive/MyDrive/...): ').strip()
    print(f'✅  BASE definida como: {BASE}')

SI = f'{BASE}'
os.makedirs(SI, exist_ok=True)

RUTA_JEX       = f'{SI}/DATOS_JEX.xlsx'
RUTA_5G        = f'{SI}/04-04.xlsx'
RUTA_GEOJSON   = f'{SI}/09-26_.geojson'
RUTA_09_13     = f'{SI}/09-13.xlsx'
RUTA_09_25     = f'{SI}/09-25.xlsx'
RUTA_09_26     = f'{SI}/09-26.xlsx'
RUTA_GOOGLEAPI = f'{SI}/datosgoogleapi.xlsx'

if 'GOOGLE_PLACES_API_KEY' not in dir():
    GOOGLE_PLACES_API_KEY = "AIzaSyADmDD0peCy1iEPnYQlwzVTqOirfbeJHdc"

# UTILIDADES

def normalizar_nombre(nombre):
    """
    Normalizo los nombres de municipios para que coincidan entre Excel y notebook.
    Correcciones aplicadas:
      1. 'Aldea (La)' → 'La Aldea'
      2. 'de La Vera' → 'de la Vera' (artículos/preposiciones en minúscula SOLO si no son inicio)
      3. 'Arroyomolinosde La Vera' → 'Arroyomolinos de la Vera' (espacio faltante)
      4. Espacios extra al inicio/final
    """
    nombre = str(nombre).strip()

    # 3. Corrección  para errores concretos del Excel
    nombre = nombre.replace('Arroyomolinosde', 'Arroyomolinos de')

    # Añadir espacio antes de una mayúscula pegada a una preposición: 'deGata' → 'de Gata'
    nombre = re.sub(r'(de|del)((?:[A-ZÁÉÍÓÚÑ])[a-záéíóúñ]+)', r'\1 \2', nombre)
    nombre = re.sub(r'([a-záéíóúñ])(de|del)([A-ZÁÉÍÓÚÑ])', r'\1 \2 \3', nombre)

    # 1. Artículos al final que aparecen entre paréntesis: 'Aldea (La)' → 'La Aldea'
    m = re.match(r'(.+)\s\((El|La|Los|Las)\)$', nombre)
    if m:
        nombre = f'{m.group(2)} {m.group(1)}'

    # 2. Forzar minúsculas en artículos y preposiciones INTERMEDIOS (no al inicio)
    # 'de La Vera' → 'de la Vera', pero 'El Gordo' → 'El Gordo' (se respeta)
    nombre = re.sub(
        r'(?<=\s)(De|Del|El|La|Las|Los|Y|En|A|Al)\b',
        lambda x: x.group(0).lower(),
        nombre
    )

    return nombre.strip()

def _ok(msg):   print(f'✅  {msg}')
def _warn(msg): print(f'⚠️  {msg}')
def _err(msg):  print(f'❌  {msg}')


# 1. DATOS_JEX — 02-01, 02-03, 02-04, 02-11, 09-23

def cargar_juntaex(municipios):
    global RUTA_JEX
    codigos = ['02-01', '02-03', '02-04', '02-11', '09-23']
    vacio = {c: pd.Series(np.nan, index=municipios) for c in codigos}

    for nombre in ['DATOS_JEX.xlsx']:
        for carpeta in [SI, BASE]:
            candidato = f'{carpeta}/{nombre}'
            if os.path.exists(candidato):
                RUTA_JEX = candidato
                break
        if RUTA_JEX:
            break

    if not RUTA_JEX:
        _warn('[JuntaEx] DATOS_JEX.xlsx no encontrado.')
        _warn(f'   Ruta esperada: {SI}/DATOS_JEX.xlsx')
        return vacio

    try:
        df = pd.read_excel(RUTA_JEX, sheet_name='Datos',
                           usecols=[0, 168, 176, 247, 248, 258, 487], header=0)
        df.columns = ['municipio', 'FM_eso', 'FU_bach', 'IN_salud',
                      'IO_hospital', 'IY_farmacia', 'RT_empleo']
        df['municipio'] = df['municipio'].map(normalizar_nombre)
        df = df.dropna(subset=['municipio']).set_index('municipio')

        resultado = {
            '02-01': (df['FM_eso'].fillna(0) + df['FU_bach'].fillna(0)).reindex(municipios),
            '02-03': df['IN_salud'].reindex(municipios),
            '02-04': df['IO_hospital'].reindex(municipios),
            '02-11': df['IY_farmacia'].reindex(municipios),
            '09-23': df['RT_empleo'].reindex(municipios),
        }
        _ok(f'[JuntaEx] Cargados: {codigos} ({df.shape[0]} municipios en fichero)')
        _ok(f'[JuntaEx] Nombres en fichero normalizados (artículos en minúsculas)')

        for cod, serie in resultado.items():
            faltantes = serie[serie.isna()].index.tolist()
            if faltantes:
                print(f'   ⚠ [{cod}] Sin datos: {len(faltantes)} municipios:')
                print(f'      {faltantes}')
                print(f'      → Comprueba que estos nombres coinciden exactamente en DATOS_JEX.xlsx')
        return resultado
    except Exception as e:
        _err(f'[JuntaEx] Error: {e}')
        return vacio


# 2. COBERTURA 5G — 04-04

def cargar_5g(municipios):
    if not os.path.exists(RUTA_5G):
        _warn(f'[04-04] {RUTA_5G} no encontrado.')
        return pd.Series(np.nan, index=municipios)
    try:
        df = pd.read_excel(RUTA_5G, sheet_name='Municipio_%hogar',
                           usecols=[0, 1, 3, 17], header=0)
        df.columns = ['ccaa', 'provincia', 'municipio', '5g']
        df = df[df['provincia'].astype(str).str.contains(
            'áceres|Caceres|CACERES', case=False, na=False)].copy()
        df['municipio'] = df['municipio'].map(normalizar_nombre)
        resultado = df.set_index('municipio')['5g'].reindex(municipios)
        _ok(f'[04-04] 5G cargado ({resultado.notna().sum()}/{len(municipios)} municipios)')
        return resultado
    except Exception as e:
        _err(f'[04-04] Error: {e}')
        return pd.Series(np.nan, index=municipios)


# 3. CIRCULAR FAB LAB — 09-24

def cargar_circularfab(municipios):
    try:
        resp = requests.get('https://circularfab.es/', timeout=30)
        resp.raise_for_status()
        detectado = any(t in resp.text.lower() for t in ['cáceres', 'caceres', 'extremadura'])
        s = pd.Series(0.0, index=municipios)
        if 'Cáceres' in municipios:
            s['Cáceres'] = float(detectado)
        _ok(f'[09-24] Fab Lab detectado: {detectado}')
        return s
    except Exception as e:
        _warn(f'[09-24] Error accediendo circularfab.es: {e}')
        return pd.Series(np.nan, index=municipios)


# 4. POLÍGONOS INDUSTRIALES — 09-26
#    Genera 09-26.xlsx desde 09-26_.geojson via geocodificación inversa Nominatim

def _geocodificar_municipio(lat, lon, intento=1):
    url = 'https://nominatim.openstreetmap.org/reverse'
    params = {'lat': lat, 'lon': lon, 'format': 'json', 'zoom': 10, 'accept-language': 'es'}
    headers = {'User-Agent': 'TFG_Caceres_Poligonos/1.0'}
    try:
        r = requests.get(url, params=params, headers=headers, timeout=10)
        r.raise_for_status()
        addr = r.json().get('address', {})
        mun = (addr.get('municipality') or addr.get('town')
               or addr.get('village') or addr.get('city', ''))
        return normalizar_nombre(mun)
    except Exception as e:
        if intento < 3:
            time.sleep(5)
            return _geocodificar_municipio(lat, lon, intento + 1)
        return ''


def generar_09_26():
    """Convierte 09-26_.geojson → 09-26.xlsx via geocodificación inversa Nominatim."""
    if not os.path.exists(RUTA_GEOJSON):
        _warn(f'[09-26] GeoJSON no encontrado: {RUTA_GEOJSON}')
        _warn('   Descárgalo desde overpass-turbo.eu con el filtro ISO3166-2=ES-CC')
        _warn('   y guárdalo como 09-26_.geojson en la carpeta Scoring Indicadores.')
        return False

    print('[09-26] Generando 09-26.xlsx desde GeoJSON')
    with open(RUTA_GEOJSON, encoding='utf-8') as f:
        data = json.load(f)
    features = data.get('features', [])
    print(f'   GeoJSON cargado: {len(features)} polígonos')

    filas = []
    for i, feat in enumerate(features):
        lon, lat = feat['geometry']['coordinates']
        props = feat.get('properties', {})
        municipio = _geocodificar_municipio(lat, lon)
        filas.append({
            'municipio':       municipio,
            'nombre_poligono': props.get('name', f"Industrial {feat.get('id', i)}"),
            'landuse':         props.get('landuse', 'industrial'),
            'lat':             lat,
            'lon':             lon,
            'osm_id':          feat.get('id', ''),
        })
        time.sleep(1.1)
        if (i + 1) % 20 == 0 or (i + 1) == len(features):
            print(f'   {i+1}/{len(features)} procesados...')

    df = pd.DataFrame(filas)
    df.to_excel(RUTA_09_26, index=False)
    _ok(f'[09-26] Guardado en: {RUTA_09_26}')
    _ok('   Fuente: © OpenStreetMap contributors, ODbL 1.0. https://osm.org/copyright')
    return True


def cargar_09_26(municipios):
    if not os.path.exists(RUTA_09_26):
        _warn(f'[09-26] {RUTA_09_26} no encontrado. Generando desde GeoJSON...')
        if not generar_09_26():
            return pd.Series(np.nan, index=municipios)
    try:
        df = pd.read_excel(RUTA_09_26)
        col_mun = next(c for c in df.columns if 'unicip' in c.lower() or 'munic' in c.lower())
        df[col_mun] = df[col_mun].map(normalizar_nombre)
        conteos = df.groupby(col_mun).size().reindex(municipios, fill_value=0)
        _ok(f'[09-26] Polígonos industriales cargados ({conteos.sum():.0f} polígonos, OSM)')
        return conteos.astype(float)
    except Exception as e:
        _err(f'[09-26] Error: {e}')
        return pd.Series(np.nan, index=municipios)


# 5. OFICINAS ACELERA PYME — 09-25
#    Genera 09-25.xlsx desde datos hardcodeados (inventario propio)

def generar_09_25():
    """Genera 09-25.xlsx — inventario propio de oficinas Acelera Pyme Rural."""
    oficinas = [
        # Diputación de Cáceres — Acelera Pyme Rural
        {'municipio': 'Arroyo de la Luz',      'nombre_oficina': 'Oficina Acelera Pyme Rural — Arroyo de la Luz',        'organismo': 'Diputación de Cáceres',                               'programa': 'Acelera Pyme Rural',     'tipo': 'Presencial', 'direccion': 'Centro Circular FAB, Arroyo de la Luz',           'lat': 39.4833, 'lon': -6.5833, 'fuente': 'acelerapyme.dip-caceres.es/localizacion.php', 'fecha_consulta': 'mayo 2026'},
        {'municipio': 'Miajadas',              'nombre_oficina': 'Oficina Acelera Pyme Rural — Miajadas',                'organismo': 'Diputación de Cáceres',                               'programa': 'Acelera Pyme Rural',     'tipo': 'Presencial', 'direccion': 'Centro Circular FAB, Miajadas',                   'lat': 39.1500, 'lon': -6.0167, 'fuente': 'acelerapyme.dip-caceres.es/localizacion.php', 'fecha_consulta': 'mayo 2026'},
        {'municipio': 'Moraleja',              'nombre_oficina': 'Oficina Acelera Pyme Rural — Moraleja',                'organismo': 'Diputación de Cáceres',                               'programa': 'Acelera Pyme Rural',     'tipo': 'Presencial', 'direccion': 'Centro Circular FAB, Moraleja',                   'lat': 40.0667, 'lon': -6.6833, 'fuente': 'acelerapyme.dip-caceres.es/localizacion.php', 'fecha_consulta': 'mayo 2026'},
        {'municipio': 'Trujillo',              'nombre_oficina': 'Oficina Acelera Pyme Rural — Trujillo',                'organismo': 'Diputación de Cáceres',                               'programa': 'Acelera Pyme Rural',     'tipo': 'Presencial', 'direccion': 'Centro Circular FAB, Trujillo',                   'lat': 39.4603, 'lon': -5.8808, 'fuente': 'acelerapyme.dip-caceres.es/localizacion.php', 'fecha_consulta': 'mayo 2026'},
        {'municipio': 'Valencia de Alcántara', 'nombre_oficina': 'Oficina Acelera Pyme Rural — Valencia de Alcántara',  'organismo': 'Diputación de Cáceres',                               'programa': 'Acelera Pyme Rural',     'tipo': 'Presencial', 'direccion': 'Centro Circular FAB, Valencia de Alcántara',      'lat': 39.4167, 'lon': -7.2333, 'fuente': 'acelerapyme.dip-caceres.es/localizacion.php', 'fecha_consulta': 'mayo 2026'},
        {'municipio': 'Cáceres',               'nombre_oficina': 'Oficina Acelera Pyme Rural — Virtual (Cáceres)',       'organismo': 'Diputación de Cáceres',                               'programa': 'Acelera Pyme Rural',     'tipo': 'Virtual',    'direccion': 'Sede Diputación de Cáceres',                      'lat': 39.4753, 'lon': -6.3724, 'fuente': 'acelerapyme.dip-caceres.es/localizacion.php', 'fecha_consulta': 'mayo 2026'},
        # CREEX
        {'municipio': 'Cáceres',               'nombre_oficina': 'Oficina Acelera Pyme CREEX — Cáceres',                 'organismo': 'CREEX — Confederación Regional Empresarial Extremeña', 'programa': 'Acelera Pyme CREEX',     'tipo': 'Presencial', 'direccion': 'C/ Obispo Segura Sáez 8, Cáceres',               'lat': 39.4753, 'lon': -6.3724, 'fuente': 'regiondigital.com — feb. 2026',              'fecha_consulta': 'mayo 2026'},
        {'municipio': 'Plasencia',             'nombre_oficina': 'Oficina Acelera Pyme CREEX — Plasencia',               'organismo': 'CREEX — Confederación Regional Empresarial Extremeña', 'programa': 'Acelera Pyme CREEX',     'tipo': 'Presencial', 'direccion': 'Avenida Juan Carlos I, 15 entreplanta, Plasencia', 'lat': 40.0303, 'lon': -6.0894, 'fuente': 'regiondigital.com — feb. 2026',              'fecha_consulta': 'mayo 2026'},
        # Cámara de Comercio
        {'municipio': 'Cáceres',               'nombre_oficina': 'Oficina Acelera Pyme — Cámara de Comercio de Cáceres', 'organismo': 'Cámara de Comercio de Cáceres',                       'programa': 'Acelera Pyme — Cámaras', 'tipo': 'Presencial', 'direccion': 'Plaza del Dr. Durán 2, 10003 Cáceres',            'lat': 39.4753, 'lon': -6.3724, 'fuente': 'camaracaceres.com/oficina-acelerapyme',       'fecha_consulta': 'mayo 2026'},
        {'municipio': 'Plasencia',             'nombre_oficina': 'Oficina Acelera Pyme — Cámara de Comercio (Plasencia)', 'organismo': 'Cámara de Comercio de Cáceres',                      'programa': 'Acelera Pyme — Cámaras', 'tipo': 'Presencial', 'direccion': 'Avda. Dolores Ibarruri 34, Plasencia',            'lat': 40.0303, 'lon': -6.0894, 'fuente': 'oapcamaracaceres.com',                        'fecha_consulta': 'mayo 2026'},
    ]
    pd.DataFrame(oficinas).to_excel(RUTA_09_25, index=False)
    _ok(f'[09-25] Generado y guardado en: {RUTA_09_25}')


def cargar_09_25(municipios):
    if not os.path.exists(RUTA_09_25):
        _warn(f'[09-25] {RUTA_09_25} no encontrado. Generando...')
        generar_09_25()
    try:
        df = pd.read_excel(RUTA_09_25)
        df['municipio'] = df['municipio'].map(normalizar_nombre)
        conteos = df.groupby('municipio').size().reindex(municipios, fill_value=0)
        _ok(f'[09-25] Oficinas Acelera Pyme cargadas ({conteos.sum():.0f} registros)')
        return conteos.astype(float)
    except Exception as e:
        _err(f'[09-25] Error: {e}')
        return pd.Series(np.nan, index=municipios)


# 6. ESPACIOS DE INNOVACIÓN — 09-13
#    Genero 09-13.xlsx desde datos hardcodeados (inventario propio)
#    Fuente: aldealab.es, fundecyt-pctex.es, camaracaceres.com, innovacion.dip-caceres.es

def generar_09_13():
    """Genera 09-13.xlsx — inventario propio de espacios de innovación."""
    espacios = [
        {'municipio': 'Cáceres',               'nombre_espacio': 'AldeaLab — Edificio Embarcadero',          'tipo': 'Coworking / Preincubadora',              'organismo': 'Ayuntamiento de Cáceres',                                        'descripcion': 'Preincubadora coworking para emprendedores en fase inicial. Top 10 Ranking Funcas 2025.',                     'direccion': 'Edificio Embarcadero, Cáceres',                     'lat': 39.4753, 'lon': -6.3724, 'fuente': 'aldealab.es / ayto-caceres.es',              'fecha_consulta': 'mayo 2026'},
        {'municipio': 'Cáceres',               'nombre_espacio': 'AldeaLab — Garaje 2.0 (CEEI Extremadura)', 'tipo': 'Incubadora tecnológica / Hub innovación', 'organismo': 'Ayuntamiento de Cáceres / FUNDECYT-PCTEX / Junta de Extremadura', 'descripcion': 'Sede del CEEI Extremadura. 30 despachos para empresas innovadoras y Factoría de Innovación de 812m².',        'direccion': 'C/ Santa Cristina s/n, Aldea Moret, 10195 Cáceres',  'lat': 39.4753, 'lon': -6.3724, 'fuente': 'aldealab.es / diputacionimpulsa.com',         'fecha_consulta': 'mayo 2026'},
        {'municipio': 'Cáceres',               'nombre_espacio': 'FUNDECYT-PCTEX — Sede Cáceres',            'tipo': 'Parque científico-tecnológico',           'organismo': 'FUNDECYT — Parque Científico y Tecnológico de Extremadura',       'descripcion': 'Parque científico-tecnológico con incubación de empresas de base tecnológica y transferencia de conocimiento.', 'direccion': 'Edificio Garaje 2.0, C/ Santa Cristina s/n, Cáceres', 'lat': 39.4753, 'lon': -6.3724, 'fuente': 'fundecyt-pctex.es / extremaduraempresarial.es', 'fecha_consulta': 'mayo 2026'},
        {'municipio': 'Plasencia',             'nombre_espacio': 'Vivero de Empresas — Cámara de Comercio',  'tipo': 'Vivero de empresas',                     'organismo': 'Cámara de Comercio de Cáceres',                                  'descripcion': '15 despachos equipados para empresas de nueva creación (máx. 2 años). Salas de reuniones y asesoramiento.',   'direccion': 'Plasencia (consultar camaracaceres.com)',             'lat': 40.0303, 'lon': -6.0894, 'fuente': 'camaracaceres.com/viveros-de-empresas',       'fecha_consulta': 'mayo 2026'},
        {'municipio': 'Navalmoral de la Mata', 'nombre_espacio': 'Vivero de Empresas — Cámara de Comercio',  'tipo': 'Vivero de empresas',                     'organismo': 'Cámara de Comercio de Cáceres',                                  'descripcion': '6 despachos equipados para empresas de nueva creación.',                                                      'direccion': 'Navalmoral de la Mata (consultar camaracaceres.com)', 'lat': 39.8833, 'lon': -5.5500, 'fuente': 'camaracaceres.com/viveros-de-empresas',       'fecha_consulta': 'mayo 2026'},
        {'municipio': 'Riolobos',              'nombre_espacio': 'Coworking Rural — Pajares de la Rivera',   'tipo': 'Coworking rural',                        'organismo': 'Ayuntamiento de Riolobos / Diputación de Cáceres',                'descripcion': 'Aldea digital reconvertida en espacio de coworking rural para teletrabajadores y emprendedores digitales.',    'direccion': 'Pajares de la Rivera, Riolobos, Cáceres',            'lat': 39.9667, 'lon': -6.2833, 'fuente': 'innovacion.dip-caceres.es/jornada-coworking', 'fecha_consulta': 'mayo 2026'},
    ]
    pd.DataFrame(espacios).to_excel(RUTA_09_13, index=False)
    _ok(f'[09-13] Generado y guardado en: {RUTA_09_13}')


def cargar_09_13(municipios):
    if not os.path.exists(RUTA_09_13):
        _warn(f'[09-13] {RUTA_09_13} no encontrado. Generando...')
        generar_09_13()
    try:
        df = pd.read_excel(RUTA_09_13)
        df['municipio'] = df['municipio'].map(normalizar_nombre)
        conteos = df.groupby('municipio').size().reindex(municipios, fill_value=0)
        _ok(f'[09-13] Espacios innovación cargados ({conteos.sum():.0f} registros)')
        return conteos.astype(float)
    except Exception as e:
        _err(f'[09-13] Error: {e}')
        return pd.Series(np.nan, index=municipios)


# 7. GOOGLE PLACES API — 02-05..02-13, 09-27
#    Si existe datosgoogleapi.xlsx lo carga directamente.
#    Si no, descarga desde la API y guarda el xlsx.

PLACES_TIPOS = {
    '02-05': [{'type': 'police'}, {'type': 'police', 'keyword': 'guardia civil'}, {'type': 'police', 'keyword': 'policia local'}],
    '02-06': [{'type': 'fire_station'}],
    '02-07': [{'type': 'shopping_mall'}, {'type': 'department_store'}, {'type': 'shopping_mall', 'keyword': 'retail park'}, {'type': 'shopping_mall', 'keyword': 'parque comercial'}],
    '02-08': [{'type': 'clothing_store'}, {'type': 'shoe_store'}, {'type': 'book_store'}, {'type': 'jewelry_store'}, {'type': 'store', 'keyword': 'calle comercial'}],
    '02-09': [{'type': 'lodging'}],
    '02-10': [{'type': 'restaurant'}],
    '02-12': [{'type': 'supermarket'}, {'type': 'grocery_or_supermarket'}],
    '02-13': [{'type': 'bank'}, {'type': 'atm'}],
    '09-27': [{'type': 'storage'}, {'type': 'store', 'keyword': 'almacen'}, {'type': 'store', 'keyword': 'comercio industrial'}],
}


def _radio_desde_area(area_km2, minimo=1000, maximo=50000):
    return int(np.clip(np.sqrt(area_km2 / np.pi) * 1000, minimo, maximo))


def _places_count(lat, lon, radio_m, type_=None, keyword=None, api_key=None):
    url = 'https://maps.googleapis.com/maps/api/place/nearbysearch/json'
    params = {'location': f'{lat},{lon}', 'radius': radio_m, 'key': api_key}
    if type_:   params['type']    = type_
    if keyword: params['keyword'] = keyword
    total, pagina = 0, 0
    while pagina < 3:
        try:
            data = requests.get(url, params=params, timeout=15).json()
        except Exception as exc:
            print(f'      ❌ Error red: {exc}')
            break
        status = data.get('status', 'UNKNOWN')
        if status == 'ZERO_RESULTS': break
        if status == 'REQUEST_DENIED':
            _err(f'REQUEST_DENIED — API key inválida. {data.get("error_message", "")}')
            return -1
        if status != 'OK':
            print(f'      ⚠ status={status} | type={type_} keyword={keyword}')
            break
        resultados = data.get('results', [])
        total  += len(resultados)
        pagina += 1
        token   = data.get('next_page_token')
        if not token or len(resultados) < 20: break
        time.sleep(2)
        params = {'pagetoken': token, 'key': api_key}
    return total


def cargar_google_places(municipios, gdf_wgs84=None, api_key=None):
    TODOS = list(PLACES_TIPOS.keys())
    resultado = {cod: pd.Series(np.nan, index=municipios) for cod in TODOS}

    # PASO 1: caché 
    if os.path.exists(RUTA_GOOGLEAPI):
        try:
            df = pd.read_excel(RUTA_GOOGLEAPI, index_col=0)
            df.index = df.index.map(normalizar_nombre)
            for cod in TODOS:
                if cod in df.columns:
                    resultado[cod] = df[cod].reindex(municipios)
            ok = [c for c in TODOS if not resultado[c].isna().all()]
            _ok(f'[POIs] Caché cargada desde {RUTA_GOOGLEAPI}')
            print(f'   Indicadores con datos: {ok}')
            for cod in ok:
                faltantes = [m for m in municipios if pd.isna(resultado[cod].get(m))]
                if faltantes:
                    print(f'   ⚠ [{cod}] Sin datos: {len(faltantes)} municipios: '
                          f'{faltantes[:5]}{"..." if len(faltantes) > 5 else ""}')
            return resultado
        except Exception as e:
            _err(f'[POIs] Error leyendo caché: {e}')

    # PASO 2: descarga desde API
    key = api_key or GOOGLE_PLACES_API_KEY
    if not key or gdf_wgs84 is None:
        _warn('[POIs] Sin caché y sin API key o gdf_wgs84. Skipping.')
        return resultado

    print('   [POIs] Verificando API key...')
    test = _places_count(39.4753, -6.3724, 5000, type_='restaurant', api_key=key)
    if test == -1:
        _err('[POIs] API key rechazada.')
        return resultado
    _ok(f'[POIs] Key válida (test: {test} restaurantes en Cáceres)')

    gdf_idx = {normalizar_nombre(row['NAMEUNIT']): (row['lat'], row['lon'], row['area_km2'])
               for _, row in gdf_wgs84.iterrows()}

    print(f'\n   [POIs] Descargando {len(municipios)} municipios × {len(TODOS)} indicadores...\n')
    errores_fat = 0
    for i, municipio in enumerate(municipios, 1):
        if municipio not in gdf_idx:
            print(f'  ⚠ [{i:>3}] {municipio}: sin centroide en gdf_wgs84')
            continue
        lat, lon, area = gdf_idx[municipio]
        radio = _radio_desde_area(area)
        print(f'  [{i:>3}/{len(municipios)}] {municipio:<35} radio={radio:>6,}m', end='')

        for cod, busquedas in PLACES_TIPOS.items():
            subtotal = 0
            for b in busquedas:
                n = _places_count(lat, lon, radio, type_=b.get('type'),
                                  keyword=b.get('keyword'), api_key=key)
                if n == -1:
                    errores_fat += 1
                    subtotal = np.nan
                    break
                subtotal += n
                time.sleep(0.25)
            resultado[cod][municipio] = subtotal

        if errores_fat >= 5:
            _err('[POIs] 5 errores fatales — deteniendo.')
            break
        else:
            errores_fat = 0

        resumen = ' | '.join(
            f"{c.split('-')[-1]}:{'?' if pd.isna(resultado[c][municipio]) else int(resultado[c][municipio])}"
            for c in TODOS)
        print(f'  →  {resumen}')

    try:
        pd.DataFrame(resultado, index=municipios).to_excel(RUTA_GOOGLEAPI)
        _ok(f'[POIs] Guardado en: {RUTA_GOOGLEAPI}')
    except Exception as e:
        _warn(f'[POIs] No se pudo guardar caché: {e}')

    return resultado


# 8. FUNCIÓN PRINCIPAL

def cargar_todos_los_datos(municipios, gdf_wgs84=None, api_key=None):
    """
    Carga y genera todos los datos necesarios para el scoring.
    Devuelve DataFrame municipios × indicadores.
    """
    print('\n' + '='*60)
    print('DESCARGA / IMPORTACIÓN DE DATOS')
    print('='*60)

    data = {}

    data.update(cargar_juntaex(municipios))
    data['04-04'] = cargar_5g(municipios)
    data['09-24'] = cargar_circularfab(municipios)
    data['09-26'] = cargar_09_26(municipios)
    data['09-25'] = cargar_09_25(municipios)
    data['09-13'] = cargar_09_13(municipios)
    data.update(cargar_google_places(municipios, gdf_wgs84=gdf_wgs84, api_key=api_key))

    df = pd.DataFrame(data, index=municipios)

    print('\n' + '='*60)
    print('COBERTURA POR INDICADOR')
    print('='*60)
    for col in sorted(df.columns):
        pct  = df[col].notna().mean() * 100
        n_ok = df[col].notna().sum()
        flag = '✅' if pct > 90 else ('⚠️ ' if pct > 50 else '❌')
        print(f'  {flag} {col:<12} {pct:5.1f}%  ({n_ok}/{len(municipios)} municipios)')

    sin_datos = df[df.isna().all(axis=1)].index.tolist()
    if sin_datos:
        print(f'\n  ⚠️  Municipios sin ningún dato ({len(sin_datos)}): {sin_datos[:10]}')

    return df


# ESTADO AL EJECUTAR

print('\n' + '='*60)
print('ESTADO DE FICHEROS DE DATOS')
print('='*60)

ficheros = {
    'DATOS_JEX.xlsx':  f'{SI}/DATOS_JEX.xlsx',
    '04-04.xlsx (5G)':         RUTA_5G,
    '09-26_.geojson (→09-26)': RUTA_GEOJSON,
    '09-13.xlsx':              RUTA_09_13,
    '09-25.xlsx':              RUTA_09_25,
    '09-26.xlsx':              RUTA_09_26,
    'datosgoogleapi.xlsx':     RUTA_GOOGLEAPI,
}

for nombre, ruta in ficheros.items():
    existe = os.path.exists(ruta)
    auto   = nombre in ('09-13.xlsx', '09-25.xlsx','09-26.xlsx')
    genera = nombre == '09-26_.geojson (→09-26)'
    if existe:
        print(f'  ✅  {nombre}')
    elif auto:
        print(f'  🔧  {nombre:<35} — se generará automáticamente')
    elif genera:
        print(f'  ⬇️  {nombre:<35} — descárgalo desde overpass-turbo.eu con ISO3166-2=ES-CC')
    elif nombre == 'datosgoogleapi.xlsx':
        print(f'  🔧  {nombre:<35} — se descargará automáticamente con la API key')
    else:
        print(f'  ❌  {nombre:<35} — FALTA: necesitas este archivo para continuar')


print(f'\n📁  Carpeta: {SI}')
print('='*60)
print('\nℹ️  Este exec() ha cargado todas las funciones de descarga.')
print('   Para cargar todos los datos en el notebook llama a:')
print('   df_indicadores = cargar_todos_los_datos(municipios_lista, gdf_wgs84=gdf_wgs84)')
