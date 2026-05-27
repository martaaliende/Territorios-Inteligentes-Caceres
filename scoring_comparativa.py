"""
SCORING Y COMPARATIVA DE AGRUPACIONES Cáceres
Contiene:
  1. Metadatos de indicadores (INDICADORES_META)
  2. Indicador compuesto 09-EMP
  3. Función de scoring por agrupación: calcular_scoring_agrupaciones()
  4. Carga de GAL y mancomunidades: cargar_agrupaciones_territoriales()
  5. Comparativa escenarios vs GAL vs mancomunidades: main_comparativa()

REQUISITO: descarga_datos.py debe ejecutarse antes.

USO EN COLAB:
  exec(open(f'{BASE}').read())
  resultados_comparativa, scores_gal, scores_manco, resumen_df = main_comparativa(
      escenarios_dict       = ESCENARIOS,
      A=A, T=T, P=P,
      d_star                = d_star_ref,
      municipios            = municipios,
      df_indicadores        = df_indicadores,
      indicadores_meta      = INDICADORES_META,
      calcular_scoring_fn   = calcular_scoring_agrupaciones,
      agrupar_municipios_fn = agrupar_municipios,
      forzar_asignacion_fn  = forzar_asignacion,
      pop_file              = POP_FILE,
      normalizar_fn         = normalizar,
      base_output           = BASE,
      max_time              = 45.0,)
"""

import warnings
import re
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

warnings.filterwarnings('ignore')


# PARTE 1: SCORING DE AGRUPACIONES (indicadores, 09-EMP, calcular_scoring_agrupaciones)

# 1. CABECERAS Y ESCENARIOS

CABECERAS_8  = ['Cáceres', 'Plasencia', 'Navalmoral de la Mata', 'Coria',
                'Trujillo', 'Moraleja', 'Jaraíz de la Vera', 'Valencia de Alcántara']
CABECERAS_10 = CABECERAS_8  + ['Miajadas', 'Guadalupe']
CABECERAS_12 = CABECERAS_10 + ['Cabezuela del Valle', 'Caminomorisco']
CABECERAS_14 = CABECERAS_12 + ['Hervás', 'Montánchez']

ESCENARIOS = {
    '8 cabeceras':  CABECERAS_8,
    '10 cabeceras': CABECERAS_10,
    '12 cabeceras': CABECERAS_12,
    '14 cabeceras': CABECERAS_14,}


# 2. METADATOS DE INDICADORES

INDICADORES_META = {
    # ACCESIBILIDAD (Junta de Extremadura)
    '02-01': dict(
        nombre='Acc. centros educativos (secundaria + bachillerato)',
        peso=1.0, sentido='pos',
        fuente='Junta de Extremadura — DATOS_JEX.xlsx cols FM+FU'),
    '02-03': dict(
        nombre='Acc. centros de salud primaria y PAS',
        peso=1.0, sentido='pos',
        fuente='Junta de Extremadura — DATOS_JEX.xlsx col IN'),
    '02-04': dict(
        nombre='Acc. hospitales',
        peso=1.0, sentido='pos',
        fuente='Junta de Extremadura — DATOS_JEX.xlsx col IO'),
    '02-11': dict(
        nombre='Acc. farmacias',
        peso=1.0, sentido='pos',
        fuente='Junta de Extremadura — DATOS_JEX.xlsx col IY'),

    # ACCESIBILIDAD (Google Places API)
    '02-05': dict(
        nombre='Acc. Policía (Local + Autonómica + Guardia Civil)',
        peso=1.0, sentido='pos',
        fuente='Google Places API — type=police'),
    '02-06': dict(
        nombre='Acc. Bomberos',
        peso=1.0, sentido='pos',
        fuente='Google Places API — type=fire_station'),
    '02-07': dict(
        nombre='Acc. centros comerciales (mall + retail park)',
        peso=1.0, sentido='pos',
        fuente='Google Places API — type=shopping_mall / department_store'),
    '02-08': dict(
        nombre='Acc. comercio a pie de calle',
        peso=1.0, sentido='pos',
        fuente='Google Places API — type=clothing_store / shoe_store / etc.'),
    '02-09': dict(
        nombre='Acc. hostelería y alojamientos',
        peso=1.0, sentido='pos',
        fuente='Google Places API — type=lodging'),
    '02-10': dict(
        nombre='Acc. restaurantes',
        peso=1.0, sentido='pos',
        fuente='Google Places API — type=restaurant'),
    '02-12': dict(
        nombre='Acc. alimentación (supermercados)',
        peso=1.0, sentido='pos',
        fuente='Google Places API — type=supermarket / grocery_or_supermarket'),
    '02-13': dict(
        nombre='Acc. servicios bancarios',
        peso=1.0, sentido='pos',
        fuente='Google Places API — type=bank / atm'),

    # CONECTIVIDAD
    '04-04': dict(
        nombre='Cobertura 5G (%)',
        peso=1.0, sentido='pos',
        fuente='Ministerio Transformación Digital — 04-04.xlsx Municipio_%hogar col 5G'),

    # ECONOMÍA
    '09-EMP': dict(
        nombre='Acc. recursos empleo e innovación [norm(09-23)×0.50 + norm(09-25)×0.25 + norm(09-24)×0.25]',
        peso=1.0, sentido='pos',
        fuente='09-23: DATOS_JEX col RT | 09-25: inventario Acelera Pyme | 09-24: circularfab.es'),
    '09-13': dict(
        nombre='Nº espacios innovación, talento y emprendimiento digital',
        peso=1.0, sentido='pos',
        fuente='Inventario propio — 09-13.xlsx | aldealab.es, fundecyt-pctex.es, camaracaceres.com — mayo 2026'),
    '09-25': dict(
        nombre='Acc. oficinas de innovación y Acelera Pyme Rural',
        peso=1.0, sentido='pos',
        fuente='Inventario propio — 09-25.xlsx | acelerapyme.dip-caceres.es, acelerapyme-creex.es, camaracaceres.com — mayo 2026'),
    '09-26': dict(
        nombre='Acc. polígonos industriales',
        peso=1.0, sentido='pos',
        fuente='OpenStreetMap contributors, ODbL 1.0 — descarga via Overpass Turbo — 09-26.xlsx'),
    '09-27': dict(
        nombre='Acc. comercio industrial',
        peso=1.0, sentido='pos',
        fuente='Google Places API — type=storage / keyword=almacen'),}


# 3. INDICADOR COMPUESTO 09-EMP
def construir_09_EMP(ind_09_23, ind_09_24, ind_09_25):
    """
    09-EMP = norm(09-23) × 0.50
           + norm(09-25) × 0.25
           + norm(09-24) × 0.25

    09-23: nº oficinas de empleo (JuntaEx)
    09-25: nº oficinas Acelera Pyme Rural (inventario propio)
    09-24: presencia Red Circular Fab Lab (circularfab.es)

    Si algún componente no tiene datos se redistribuye su peso
    proporcionalmente entre los disponibles.
    """
    def safe_norm(s):
        s = s.fillna(0)
        mn, mx = s.min(), s.max()
        return pd.Series(0.5, index=s.index) if mx == mn else (s - mn) / (mx - mn)

    tiene_23 = not ind_09_23.isna().all()
    tiene_25 = not ind_09_25.isna().all()
    tiene_24 = not ind_09_24.isna().all()

    if not tiene_23 and not tiene_25 and not tiene_24:
        print('❌  [09-EMP] Sin datos para ningún componente (09-23, 09-25, 09-24)')
        return pd.Series(np.nan, index=ind_09_23.index)

    # Pesos base
    pesos = {'09-23': 0.50 if tiene_23 else 0.0,
             '09-25': 0.25 if tiene_25 else 0.0,
             '09-24': 0.25 if tiene_24 else 0.0}
    total = sum(pesos.values())
    pesos = {k: v / total for k, v in pesos.items()}  # renormalizar si falta alguno

    r = pd.Series(0.0, index=ind_09_23.index)
    if tiene_23: r += safe_norm(ind_09_23) * pesos['09-23']
    if tiene_25: r += safe_norm(ind_09_25) * pesos['09-25']
    if tiene_24: r += safe_norm(ind_09_24) * pesos['09-24']

    componentes = (
        f"norm(09-23)×{pesos['09-23']:.2f}" +
        (f" + norm(09-25)×{pesos['09-25']:.2f}" if tiene_25 else '') +
        (f" + norm(09-24)×{pesos['09-24']:.2f}" if tiene_24 else '')
    )
    print(f'✅  [09-EMP] {componentes}')
    return r


def preparar_indicadores(df_indicadores, municipios):
    """
    Recibe el DataFrame de descarga_datos.py y añade 09-EMP.
    09-EMP usa 09-23, 09-25 y 09-24 como componentes.
    Como 09-25 ya entra en 09-EMP, se elimina de INDICADORES_META
    para no contabilizarla dos veces en el scoring global.
    Devuelve el DataFrame completo listo para el scoring.
    """
    df = df_indicadores.copy()
    ind_09_23 = df.get('09-23', pd.Series(np.nan, index=municipios))
    ind_09_24 = df.get('09-24', pd.Series(np.nan, index=municipios))
    ind_09_25 = df.get('09-25', pd.Series(np.nan, index=municipios))
    df['09-EMP'] = construir_09_EMP(ind_09_23, ind_09_24, ind_09_25)
    return df

# 4. SCORING

def _normalizar(serie, sentido='pos'):
    mn, mx = serie.min(), serie.max()
    if mx == mn:
        return pd.Series(0.5, index=serie.index)
    norm = (serie - mn) / (mx - mn)
    return norm if sentido == 'pos' else (1 - norm)


def calcular_scoring_agrupaciones(df_asig, df_indicadores):
    """
    Scoring 0-100 por agrupación: ponderación global única sin distinción por ámbito.
    Todos los indicadores tienen el mismo peso (1.0).
    
    Parámetros:
      df_asig        : DataFrame con columnas ['municipio', 'cabecera']
      df_indicadores : DataFrame municipios × indicadores (salida de descarga_datos.py)

    Devuelve (df_scores, df_raw).
      df_scores : score global por cabecera
      df_raw    : valores brutos medios por indicador y cabecera
    """
    municipios = df_asig['municipio'].unique().tolist()

    # Añadir 09-EMP si no está ya (incluye 09-23, 09-25 y 09-24)
    if '09-EMP' not in df_indicadores.columns:
        df_indicadores = preparar_indicadores(df_indicadores, municipios)

    # Forzar todas las columnas a numérico para evitar TypeError por strings
    df_indicadores = df_indicadores.apply(pd.to_numeric, errors='coerce')

    # 09-25 entra dentro de 09-EMP, no se puntúa por separado
    indicadores_scoring = {k: v for k, v in INDICADORES_META.items() if k != '09-25'}

    grupos = df_asig.groupby('cabecera')['municipio'].apply(list).to_dict()

    # Paso 1: media por agrupación
    ind_por_grupo = {}
    for cod in indicadores_scoring:
        ind_por_grupo[cod] = {}
        for cab, munis in grupos.items():
            if cod in df_indicadores.columns:
                vals = df_indicadores[cod].reindex(munis).dropna()
                ind_por_grupo[cod][cab] = float(vals.mean()) if len(vals) > 0 else np.nan
            else:
                ind_por_grupo[cod][cab] = np.nan

    df_raw = pd.DataFrame(ind_por_grupo, index=list(grupos.keys()))

    # Paso 2: normalizar entre agrupaciones (todos sentido 'pos')
    df_norm = pd.DataFrame(index=df_raw.index)
    for cod, meta in indicadores_scoring.items():
        col = df_raw[cod].astype(float) if cod in df_raw.columns \
              else pd.Series(np.nan, index=df_raw.index)
        df_norm[cod] = _normalizar(col, meta['sentido'])

    # Paso 3: score global: media simple (peso 1.0 igual para todos)
    vals = df_norm.values.astype(float)
    score_global = []
    for row in vals:
        mask = ~np.isnan(row)
        score_global.append(float(row[mask].mean()) * 100 if mask.sum() > 0 else np.nan)

    scores = pd.DataFrame(index=df_raw.index)
    scores['score_global']      = np.array(score_global)
    scores['n_municipios']      = [len(grupos.get(c, [])) for c in scores.index]
    scores['cobertura_datos_%'] = (
        df_norm.notna().sum(axis=1) / len(indicadores_scoring) * 100).round(1)
    return scores.round(2), df_raw.round(4)


# 5. INFORME DE FUENTES
def imprimir_informe_fuentes():
    print('\n' + '='*70)
    print('INDICADORES ACTIVOS Y FUENTES')
    print('='*70)
    print(f'{"CÓDIGO":<10} {"PESO":<6} FUENTE')
    print('-'*70)
    for cod, meta in INDICADORES_META.items():
        nota = ' [→ entra en 09-EMP]' if cod == '09-25' else ''
        print(f'{cod:<10} {meta["peso"]:<6} {meta["fuente"][:55]}{nota}')
    print()
    print('SCORING: media ponderada global única de todos los indicadores.')
    print('  09-EMP = norm(09-23)×0.50 + norm(09-25)×0.25 + norm(09-24)×0.25')
    print()


# PARTE 2 COMPARATIVA ESCENARIOS vs GAL vs MANCOMUNIDADES

# UTILIDADES

def _norm_nombre(nombre):
    nombre = str(nombre).strip()
    nombre = nombre.replace('Arroyomolinosde', 'Arroyomolinos de')
    nombre = re.sub(r'(de|del)((?:[A-ZÁÉÍÓÚÑ])[a-záéíóúñ]+)', r'\1 \2', nombre)
    nombre = re.sub(r'([a-záéíóúñ])(de|del)([A-ZÁÉÍÓÚÑ])', r'\1 \2 \3', nombre)
    m = re.match(r'(.+)\s\((El|La|Los|Las)\)$', nombre)
    if m:
        nombre = f'{m.group(2)} {m.group(1)}'
    nombre = re.sub(r'(?<=\s)(De|Del|El|La|Las|Los|Y|En|A|Al)\b',
                    lambda x: x.group(0).lower(), nombre)
    return nombre.strip()


def _normalizar_serie(serie):
    mn, mx = serie.min(), serie.max()
    if mx == mn:
        return pd.Series(0.5, index=serie.index)
    return (serie - mn) / (mx - mn)


# CARGAR GAL Y MANCOMUNIDADES
def cargar_agrupaciones_territoriales(pop_file, municipios_validos):
    try:
        df = pd.read_excel(pop_file, sheet_name='Municipios', header=1)
    except Exception as e:
        print(f'❌ Error leyendo {pop_file}: {e}')
        return {}, {}

    df['mun'] = df['Municipio'].map(_norm_nombre)
    df = df[df['mun'].isin(municipios_validos)].copy()

    # Nombres cortos para GAL que no tienen acrónimo entre paréntesis
    NOMBRES_CORTOS_MANUALES = {
        'Asociación para el Desarrollo de la Comarca de la Sierra San Pedro - Los Baldíos': 'Sierra San Pedro',
    }

    def nombre_corto_gal(nombre):
        nombre = str(nombre).strip()
        if nombre in NOMBRES_CORTOS_MANUALES:
            return NOMBRES_CORTOS_MANUALES[nombre]
        m = re.search(r'\(([^)]+)\)$', nombre)
        return m.group(1) if m else nombre

    df['gal_corto'] = df['Grupo de acción local'].apply(nombre_corto_gal)

    grupos_gal = (df[~df['Grupo de acción local'].isin(['null', 'nan', None]) &
                    df['Grupo de acción local'].notna()]
                  .groupby('gal_corto')['mun'].apply(list).to_dict())

    grupos_manco = (df[df['Mancomunidad'] != 'Sin Mancomunar']
                    .groupby('Mancomunidad')['mun'].apply(list).to_dict())

    print(f'✅  GAL: {len(grupos_gal)} grupos | '
          f'Mancomunidades: {len(grupos_manco)} grupos')
    return grupos_gal, grupos_manco


# CALCULAR MEDIAS POR INDICADOR PARA UN CONJUNTO DE GRUPOS
def _medias_por_indicador(df_indicadores, grupos, indicadores_meta):
    """
    Para cada grupo calcula la media de sus municipios en cada indicador.
    Devuelve DataFrame grupos × indicadores con valores brutos.
    """
    # 09-25 entra en 09-EMP, no se puntúa por separado
    indicadores = [k for k in indicadores_meta if k != '09-25']

    # Forzar todas las columnas a numérico para evitar TypeError por strings
    df_num = df_indicadores.apply(pd.to_numeric, errors='coerce')

    filas = {}
    for grp, munis in grupos.items():
        fila = {}
        for cod in indicadores:
            if cod in df_num.columns:
                vals = df_num[cod].reindex(munis).dropna()
                fila[cod] = float(vals.mean()) if len(vals) > 0 else np.nan
            else:
                fila[cod] = np.nan
        filas[grp] = fila

    return pd.DataFrame(filas).T  # grupos × indicadores



# FUNCIÓN PRINCIPAL

def main_comparativa(escenarios_dict, A, T, P, d_star, municipios,
                     df_indicadores, indicadores_meta,
                     calcular_scoring_fn, agrupar_municipios_fn,
                     forzar_asignacion_fn, pop_file, normalizar_fn,
                     base_output, max_time=45.0, peso_ambito=None):
    """
    Compara escenarios de cabeceras, GAL y Mancomunidades indicador a indicador.
    peso_ambito se acepta por compatibilidad pero no se usa.

    Devuelve (resultados_escenarios, scores_gal, scores_manco, resumen_df).
    """
    print('\n' + '='*65)
    print('COMPARATIVA POR INDICADOR: ESCENARIOS vs GAL vs MANCOMUNIDADES')
    print('='*65)

    indicadores = [k for k in indicadores_meta if k != '09-25']

    # construir 09-EMP en df_indicadores antes de calcular medias.
    if '09-EMP' not in df_indicadores.columns:
        df_indicadores = preparar_indicadores(df_indicadores, list(df_indicadores.index))

    # 1. Medias por indicador para cada escenario
    resultados_escenarios = {}
    grupos_escenario_todos = {}  # acumular todos los grupos de todos los escenarios

    for nombre_esc, cabeceras in escenarios_dict.items():
        df_asig, _ = agrupar_municipios_fn(cabeceras, A, T, P, d_star, max_time=max_time)
        df_asig, _ = forzar_asignacion_fn(
            df_asig, [h for h in cabeceras if h in municipios], T, A=A)
        grupos = df_asig.groupby('cabecera')['municipio'].apply(list).to_dict()
        df_med = _medias_por_indicador(df_indicadores, grupos, indicadores_meta)
        df_med['tipo']      = 'Escenario'
        df_med['agrupacion'] = nombre_esc
        resultados_escenarios[nombre_esc] = df_med
        grupos_escenario_todos.update(grupos)

    # 2. Medias por indicador para GAL y Mancomunidades
    grupos_gal, grupos_manco = cargar_agrupaciones_territoriales(pop_file, municipios)

    df_gal   = _medias_por_indicador(df_indicadores, grupos_gal,   indicadores_meta)
    df_manco = _medias_por_indicador(df_indicadores, grupos_manco, indicadores_meta)
    df_gal['tipo']       = 'GAL'
    df_gal['agrupacion'] = 'GAL'
    df_manco['tipo']       = 'Mancomunidad'
    df_manco['agrupacion'] = 'Mancomunidades'

    scores_gal   = df_gal
    scores_manco = df_manco

    # 3. Tabla unificada con normalización consistente
    # tomando como regrencia de min max los valores de los GAL + Mancomunidades (territorios reales y estables).
    # Los escenarios se proyectan sobre esa escala haciendo los scores comparables entre sí.

    df_referencia = pd.concat([df_gal, df_manco])[indicadores].apply(pd.to_numeric, errors='coerce')

    def _normalizar_con_referencia(serie_grupo, serie_ref):
        """Normaliza serie_grupo usando min/max de serie_ref (universo fijo)."""
        mn = serie_ref.min()
        mx = serie_ref.max()
        if mx == mn:
            return pd.Series(0.5, index=serie_grupo.index)
        norm = (serie_grupo - mn) / (mx - mn)
        return norm.clip(0, 1)  # los escenarios pueden superar el rango de referencia

    def _score_bloque(df_raw_bloque):
        """Devuelve df normalizado + score_global para un bloque de grupos."""
        df_num = df_raw_bloque[indicadores].apply(pd.to_numeric, errors='coerce')
        df_n = pd.DataFrame(index=df_num.index)
        for cod in indicadores:
            df_n[cod] = _normalizar_con_referencia(df_num[cod], df_referencia[cod])
        df_n['score_global'] = df_n[indicadores].mean(axis=1) * 100
        return df_n

    partes_norm = []

    for nombre_esc, df_med in resultados_escenarios.items():
        df_n = _score_bloque(df_med)
        df_n['tipo']         = 'Escenario'
        df_n['agrupacion']   = nombre_esc
        df_n['grupo']        = df_med.index
        df_n['n_municipios'] = [len(grupos_escenario_todos.get(g, [])) for g in df_med.index]
        partes_norm.append(df_n)

    df_n_gal = _score_bloque(df_gal)
    df_n_gal['tipo']         = 'GAL'
    df_n_gal['agrupacion']   = 'GAL'
    df_n_gal['grupo']        = df_gal.index
    df_n_gal['n_municipios'] = [len(grupos_gal.get(g, [])) for g in df_gal.index]
    partes_norm.append(df_n_gal)

    df_n_manco = _score_bloque(df_manco)
    df_n_manco['tipo']         = 'Mancomunidad'
    df_n_manco['agrupacion']   = 'Mancomunidades'
    df_n_manco['grupo']        = df_manco.index
    df_n_manco['n_municipios'] = [len(grupos_manco.get(g, [])) for g in df_manco.index]
    partes_norm.append(df_n_manco)

    # Reordenar columnas
    cols_front = ['tipo', 'agrupacion', 'grupo', 'score_global', 'n_municipios']
    resumen_df = pd.concat(partes_norm)[cols_front + indicadores].reset_index(drop=True)

    # 4. Imprimir resumen 
    print('\n=== SCORE GLOBAL POR GRUPO ===')
    print(resumen_df[['tipo','agrupacion','grupo','score_global','n_municipios']]
          .sort_values(['agrupacion','score_global'], ascending=[True,False])
          .to_string(index=False))

    print('\n=== MEDIA POR TIPO DE AGRUPACIÓN ===')
    print(resumen_df.groupby('tipo')['score_global'].agg(['mean','min','max']).round(2))

    print('\n=== SCORE POR ESCENARIO DE AGRUPACIÓN ===')
    esc_df = (resumen_df[resumen_df['tipo'] == 'Escenario']
              .groupby('agrupacion')['score_global']
              .agg(media='mean', minimo='min', maximo='max',
                   grupos_sobre_media_gal=lambda x: (x > resumen_df[resumen_df['tipo']=='GAL']['score_global'].mean()).sum()))
    # Ordenar por número de cabeceras
    orden = [k for k in escenarios_dict.keys() if k in esc_df.index]
    esc_df = esc_df.reindex(orden)
    esc_df.columns = ['Score medio', 'Score mín.', 'Score máx.', 'Grupos > media GAL']
    print(esc_df.round(2).to_string())

    # 5. Guardar Excel
    ruta_out = f'{base_output}/comparativa_indicadores.xlsx'
    try:
        with pd.ExcelWriter(ruta_out, engine='openpyxl') as writer:
            resumen_df.to_excel(writer, sheet_name='Todos', index=False)
            resumen_df[resumen_df['tipo']=='GAL'].to_excel(
                writer, sheet_name='GAL', index=False)
            resumen_df[resumen_df['tipo']=='Mancomunidad'].to_excel(
                writer, sheet_name='Mancomunidades', index=False)
            for nombre_esc in escenarios_dict:
                resumen_df[resumen_df['agrupacion']==nombre_esc].to_excel(
                    writer, sheet_name=nombre_esc[:31], index=False)
        print(f'\n✅  Excel guardado en: {ruta_out}')
    except Exception as e:
        print(f'⚠️  No se pudo guardar: {e}')

    # 6. Gráfico boxplot
    fig, ax = plt.subplots(figsize=(13, 5))
    etiquetas, datos, colores = [], [], []
    palette = plt.cm.Blues(np.linspace(0.4, 0.85, len(escenarios_dict)))

    for i, nombre_esc in enumerate(escenarios_dict):
        sub = resumen_df[resumen_df['agrupacion'] == nombre_esc]['score_global'].dropna()
        etiquetas.append(nombre_esc)
        datos.append(sub.values)
        colores.append(palette[i])

    etiquetas.append('GAL')
    datos.append(resumen_df[resumen_df['tipo']=='GAL']['score_global'].dropna().values)
    colores.append('#e07b39')

    etiquetas.append('Mancomunidades')
    datos.append(resumen_df[resumen_df['tipo']=='Mancomunidad']['score_global'].dropna().values)
    colores.append('#5a9e6f')

    bp = ax.boxplot(datos, patch_artist=True,
                    medianprops=dict(color='black', linewidth=2))
    for patch, color in zip(bp['boxes'], colores):
        patch.set_facecolor(color)
        patch.set_alpha(0.8)

    ax.set_xticks(range(1, len(etiquetas) + 1))
    ax.set_xticklabels(etiquetas, rotation=15, ha='right', fontsize=9)
    ax.set_ylabel('Score Global (0–100)')
    ax.set_title('Comparativa por indicador: Escenarios vs GAL vs Mancomunidades')
    ax.grid(axis='y', alpha=0.3)
    plt.tight_layout()

    ruta_fig = f'{base_output}/comparativa_boxplot.png'
    try:
        plt.savefig(ruta_fig, dpi=150)
        print(f'✅  Gráfico guardado en: {ruta_fig}')
    except Exception:
        pass
    plt.show()

    return resultados_escenarios, scores_gal, scores_manco, resumen_df


print('✅  scoring_comparativa.py cargado.')
print('   Funciones: calcular_scoring_agrupaciones(), main_comparativa()')
