# -*- coding: utf-8 -*-
"""
generar_rutas.py
Lee el Excel de trabajos del portal (formato direcciones.xlsx) y genera
rutas.json para el dashboard Rutas-tecnicos, SIN datos de clientes.

Novedad: ahora calcula el RGU de cada orden (regla de COMPONENTES, la misma
calibrada contra Power BI) y lo emite en el campo "rgu" de cada trabajo.
Asi el dashboard publicado muestra el panel RGU sin tener que cargar el Excel
a mano, y NO se recalcula en el navegador (fuente unica de verdad).

Uso:
    python generar_rutas.py                     -> usa el Excel mas reciente de CARPETA_EXCEL
    python generar_rutas.py archivo.xlsx        -> usa un archivo especifico
    python generar_rutas.py --subir             -> ademas lo sube a GitHub
    python generar_rutas.py archivo.xlsx --subir
"""
import pandas as pd
import json
import re
import sys
import os
import base64
import unicodedata
from datetime import datetime

# ================== CONFIGURACION (ajustar) ==================
CARPETA_EXCEL = r"C:\Users\Luis Sepulveda\Rutas tecnicos\descargas"   # carpeta donde cae el Excel del portal
ARCHIVO_SALIDA = "rutas.json"

GITHUB_USUARIO = "luis-sepulveda"
GITHUB_REPO = "Rutas-tecnicos"
GITHUB_ARCHIVO = "rutas.json"
# Rama SOLO de datos, separada de la que publica GitHub Pages.
# Pages reconstruye el sitio con cada commit de su rama y tiene un tope
# de 10 construcciones por hora; subiendo cada 5 min se pasa del tope y
# los cambios reales del index.html quedan encolados detras de
# reconstrucciones que no sirven de nada, porque el dashboard lee este
# archivo por raw.githubusercontent y no por Pages.
GITHUB_RAMA = "data"
# El token se lee de la variable de entorno GITHUB_TOKEN (igual que subir_github.py)
# =============================================================

# Columnas OBLIGATORIAS (si falta una, se aborta)
COLUMNAS = ['Técnico', 'Orden de Trabajo', 'Tipo de Actividad', 'Franja',
            'Inicio', 'Fin', 'Dirección', 'Ciudad', 'Coordenada X', 'Coordenada Y', 'Estado']
# Columnas OPCIONALES para el RGU (si el export las trae; si no, se rellenan vacias)
COLUMNAS_RGU = ['Subtipo', 'Pasos', 'Fecha']
# Columnas OPCIONALES para la regla "GSA cuenta como Completado"
# (si el export reducido no las trae, la regla simplemente no aplica)
COLUMNAS_CIERRE = ['Código de Cierre', 'Area derivación']

# Tipos que son trabajo de verdad. Todo lo demas que trae el export (Colación,
# No inicia ruta, Gestion Con SOPTECC, etc.) no es una orden de terreno y no
# se anota como parte del bucket.
TIPOS_REALES = ['Reparación', 'Alta', 'Migración',
                'Modificación de Servicio', 'Upgrade promoción']

# OT que quedaron fuera del filtro DOMI en la ultima pasada de preparar_df().
# Lo llena preparar_df() y lo lee generar() para meterlo en rutas.json.
OTS_FUERA_DOMI = []

# Columna OPCIONAL con el numero de peticion/solicitud. El portal la nombra
# distinto segun el export; se busca por varios alias y si no viene, queda
# vacia y el dashboard muestra la OT en su lugar.
COLUMNAS_PET = ['Petición']


def limpiar_tecnico(t):
    """Quita prefijos tipo FS_MM_NFTT_DOMI_ / QA_B2B_MM_NFTT_DOMI_"""
    return re.sub(r'^(QA_)?(B2B_)?(FS_)?MM_NFTT_(DOMI|TRAZ)_', '', str(t)).strip()


# ============ NORMALIZACION DE TEXTO Y COLUMNAS ============
def norm_txt(s):
    """minusculas, sin tildes, sin espacios dobles ni de los extremos."""
    s = str(s or '')
    s = ''.join(c for c in unicodedata.normalize('NFD', s)
                if unicodedata.category(c) != 'Mn')
    return re.sub(r'\s+', ' ', s).strip().lower()


def buscar_col(df, *alias):
    """Devuelve el nombre REAL de la columna del Excel que calza con alguno de
    los alias, comparando sin tildes / sin mayusculas / sin espacios sobrantes.
    Asi da lo mismo si el portal manda 'Area derivación' o 'Área Derivacion '."""
    mapa = {norm_txt(c): c for c in df.columns}
    for a in alias:
        real = mapa.get(norm_txt(a))
        if real:
            return real
    return None


# Estados canonicos: el dashboard compara con estos strings EXACTOS.
_ESTADOS = {
    'completado': 'Completado', 'completada': 'Completado',
    'no realizada': 'No Realizada', 'no realizado': 'No Realizada',
    'pendiente': 'Pendiente', 'iniciado': 'Iniciado', 'iniciada': 'Iniciado',
    'en ruta': 'en ruta', 'enruta': 'en ruta',
    'cancelado': 'Cancelado', 'cancelada': 'Cancelado',
    'suspendido': 'Suspendido', 'suspendida': 'Suspendido',
}


def norm_estado(e):
    return _ESTADOS.get(norm_txt(e), str(e or '').strip())
# ===========================================================


# ================== CALCULO DE RGU ==================
# REGLA DE COMPONENTES (calibrada contra Power BI). Debe quedar IDENTICA a la
# funcion rguComponentes()/rguOrden() del index.html -> mantener en sincronia.
#   Reparacion                 = 1 (plano)
#   Alta/Traslado/Migracion/Upgrade = suma de componentes de los Pasos:
#       D-Box IPTV : 1o = 1, cada adicional = 0,5
#       Extensor   : 0,5 c/u
#       Plan Banda Ancha : 1
#       Plan de telefonia: 1
#       Gateway NextGen y ONT GPON : NO cuentan
#       (solo pasos "Instalar"/"Actualizar"; los "Desinstalar" no cuentan)
#   Modificacion de Servicio   = 0,1 si hay una Alta/Migracion COMPLETADA en la
#                                misma direccion y fecha; si no, 1
#   Colacion / otros           = 0
def _descripciones_utiles(pasos):
    descs = re.findall(r'<descripcion>(.*?)</descripcion>', str(pasos or ''), re.I | re.S)
    return [d for d in descs if re.match(r'\s*(instalar|actualizar)\b', d, re.I)]


def rgu_componentes(pasos):
    descs = _descripciones_utiles(pasos)
    dbox = sum(1 for d in descs if re.search(r'd-?box', d, re.I))
    ext = sum(1 for d in descs if re.search(r'extensor', d, re.I))
    ba = sum(1 for d in descs if re.search(r'plan banda ancha', d, re.I))
    tel = sum(1 for d in descs if re.search(r'plan de telefon|telefon[ií]a', d, re.I))
    r = 0.0
    if dbox:
        r += 1 + 0.5 * (dbox - 1)
    r += 0.5 * ext + ba + tel
    return r


def rgu_orden(tipo, subtipo, pasos):
    t = str(tipo or '').strip().lower()
    if 'reparaci' in t:
        return 1
    if t == 'alta' or 'traslado' in t or 'migraci' in t or 'upgrade' in t:
        r = rgu_componentes(pasos)
        if r > 0:
            return r
        # respaldo si el export no trae Pasos: N Play del Subtipo
        m = re.search(r'(\d+)\s*play', str(subtipo or ''), re.I)
        return int(m.group(1)) if m else 1
    if 'modifica' in t:
        return 1   # se ajusta a 0,1 mas abajo si corresponde
    return 0       # Colacion, No inicia, etc.


def _dirbase(d):
    s = str(d or '').upper()
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    return re.sub(r'\s+', ' ', s).strip()


def _play_sub(sub):
    m = re.search(r'(\d+)\s*play', str(sub or ''), re.I)
    return m.group(1) if m else ''


def _cats_migr(pasos):
    descs = _descripciones_utiles(pasos)
    c = 0
    if any(re.search(r'plan banda ancha', d, re.I) for d in descs):
        c += 1
    if any(re.search(r'd-?box', d, re.I) for d in descs):
        c += 1
    if any(re.search(r'plan de telefon|telefon[ií]a', d, re.I) for d in descs):
        c += 1
    return c or 1


def etiqueta(tipo, sub, pasos):
    """Etiqueta de tipo de actividad con su tramo de Play (para el detalle del panel)."""
    t = str(tipo or '').strip().lower()
    if 'reparaci' in t:
        return 'Reparación'
    if t == 'alta':
        pl = _play_sub(sub)
        return f'Alta {pl} Play' if pl else 'Alta'
    if 'traslado' in t:
        pl = _play_sub(sub)
        return f'Alta Traslado {pl} Play' if pl else 'Alta Traslado'
    if 'migraci' in t:
        return f'Migración {_cats_migr(pasos)} Play'
    if 'upgrade' in t:
        return 'Upgrade promoción'
    if 'modifica' in t:
        return 'Modificación de Servicio'
    return str(tipo or 'Otros')


def calcular_rgu(df):
    """Agrega la columna 'rgu' al DataFrame ya filtrado."""
    # generadoras completadas (Alta/Migracion/Traslado) por direccion+fecha
    gener = set()
    for r in df.itertuples():
        t = str(r.t).strip().lower()
        if r.e == 'Completado' and (t == 'alta' or 'traslado' in t or 'migraci' in t):
            gener.add(_dirbase(r.dir) + '|' + str(r.fecha))

    valores = []
    for r in df.itertuples():
        t = str(r.t).strip().lower()
        if 'modifica' in t:
            val = 0.1 if (_dirbase(r.dir) + '|' + str(r.fecha)) in gener else 1
        else:
            val = rgu_orden(r.t, r.sub, r.pasos)
        valores.append(round(float(val), 2))
    df['rgu'] = valores
    return df
# ===================================================


def fecha_de_datos(df):
    """Dia al que PERTENECEN los datos del export (columna 'Fecha' del portal),
    en formato 'dd-mm-yyyy'.

    Esto NO es lo mismo que la hora en que corrio el script. A primera hora el
    portal todavia puede entregar la jornada de AYER; si marcamos el snapshot
    solo con datetime.now(), las alertas creen que ayer y hoy son el mismo dia
    y marcan como "caidas" todas las ordenes abiertas de ayer.

    Se usa la fecha MAS REPETIDA (moda) porque el export puede traer alguna
    orden suelta de otro dia. Si el export no trae la columna, cae a hoy.
    """
    try:
        s = pd.to_datetime(df['fecha'], errors='coerce', dayfirst=True).dropna()
        if len(s):
            return s.dt.strftime('%d-%m-%Y').mode().iloc[0]
    except Exception as e:
        print(f"AVISO: no pude leer la columna Fecha ({e}). Se asume hoy.")
    return datetime.now().strftime('%d-%m-%Y')


def excel_mas_reciente(carpeta):
    archivos = [os.path.join(carpeta, f) for f in os.listdir(carpeta)
                if f.lower().endswith('.xlsx') and not f.startswith('~$')]
    if not archivos:
        raise FileNotFoundError(f"No hay archivos .xlsx en {carpeta}")
    return max(archivos, key=os.path.getmtime)


def preparar_df(path_excel):
    """Lee el Excel del portal y devuelve el DataFrame ya normalizado,
    filtrado (DOMI + Quilpue/Villa Alemana) y con la regla GSA aplicada.

    Se separo de generar() para que historial_dir.py use EXACTAMENTE la
    misma preparacion. Duplicar esta logica en dos scripts es justo el
    problema que ya tuvimos con el RGU en transformador.py / procesar.py:
    se tocaba uno y el otro quedaba desalineado sin que nadie lo notara.

    Devuelve None si tras los filtros no quedo ninguna orden."""
    print(f"Leyendo: {path_excel}")
    df = pd.read_excel(path_excel, engine="openpyxl")

    faltantes = [c for c in COLUMNAS if c not in df.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas en el Excel: {faltantes}")

    # columnas opcionales de RGU: si no vienen en el export, se crean vacias
    for c in COLUMNAS_RGU:
        if c not in df.columns:
            df[c] = ''
            print(f"AVISO: el export no trae la columna '{c}'. El RGU usara respaldo (N Play).")

    # columnas de cierre para la regla GSA: se buscan SIN tildes ni mayusculas,
    # asi da lo mismo como venga escrito el encabezado en el export.
    col_cierre = buscar_col(df, 'Código de Cierre', 'Codigo de Cierre',
                            'Cod. de Cierre', 'Código Cierre', 'Motivo de Cierre')
    col_deriv = buscar_col(df, 'Area derivación', 'Área derivación',
                           'Area de derivación', 'Área de derivación',
                           'Area derivacion', 'Derivación')

    if col_cierre and col_cierre != 'Código de Cierre':
        print(f"OK: columna de cierre detectada como '{col_cierre}'")
        df = df.rename(columns={col_cierre: 'Código de Cierre'})
    elif not col_cierre:
        df['Código de Cierre'] = ''
        print("*** ATENCION: el export NO trae columna de Codigo de Cierre. "
              "La regla GSA NO se aplicara. ***")

    # numero de peticion (opcional): mismo truco de alias sin tildes
    col_pet = buscar_col(df, 'Petición', 'Peticion', 'N° Petición', 'Nro Petición',
                         'Numero de Petición', 'Número de Petición', 'ID Petición',
                         'Solicitud', 'N° Solicitud')
    if col_pet and col_pet != 'Petición':
        print(f"OK: columna de peticion detectada como '{col_pet}'")
        df = df.rename(columns={col_pet: 'Petición'})
    elif not col_pet:
        df['Petición'] = ''

    if col_deriv and col_deriv != 'Area derivación':
        print(f"OK: columna de derivacion detectada como '{col_deriv}'")
        df = df.rename(columns={col_deriv: 'Area derivación'})
    elif not col_deriv:
        df['Area derivación'] = ''
        print("*** ATENCION: el export NO trae columna de Area de derivacion. "
              "La regla GSA NO se aplicara. ***")

    df = df[COLUMNAS + COLUMNAS_RGU + COLUMNAS_CIERRE + COLUMNAS_PET].copy()
    df.columns = ['tec', 'o', 't', 'f', 'hi', 'hf', 'dir', 'ciudad',
                  'lng', 'lat', 'e', 'sub', 'pasos', 'fecha', 'cierre', 'deriv',
                  'pet']

    # ================== FILTROS DE NEGOCIO ==================
    # Antes de filtrar se anotan las OT de actividades REALES que quedan
    # fuera del filtro DOMI: son las que estan en el bucket de la plataforma
    # ("V Quilpue" y similares), todavia sin tecnico asignado.
    #
    # No se dibujan ni se cuentan en el tablero, pero alertas_rutas.py las
    # necesita: sin ellas, cuando despacho reparte el bucket a los tecnicos
    # —tipico entre 08:30 y 09:00— esas ordenes aparecen por primera vez en el
    # snapshot filtrado y se cuentan como novedad del dia, cuando en realidad
    # ya estaban ahi desde temprano, solo que sin dueno.
    global OTS_FUERA_DOMI
    _es_domi = df['tec'].astype(str).str.contains('_DOMI_', case=False, na=False)
    _real = df['t'].astype(str).isin(TIPOS_REALES)
    OTS_FUERA_DOMI = sorted(set(df.loc[~_es_domi & _real, 'o'].dropna().astype(str)))
    if OTS_FUERA_DOMI:
        print(f"[bucket] {len(OTS_FUERA_DOMI)} orden(es) sin tecnico DOMI "
              f"(bucket de la plataforma). Se anotan para no contarlas como "
              f"nuevas cuando se asignen.")

    # 1) solo tecnicos DOMI (excluye buckets como "V Quilpue", TRAZ, HFC, etc.)
    #    Se usa _DOMI_ con guiones para no capturar nombres que contengan "DOMI"
    df = df[df['tec'].astype(str).str.contains('_DOMI_', case=False, na=False)]
    # 2) solo Quilpue y Villa Alemana
    df = df[df['ciudad'].astype(str).str.strip().str.upper().isin(['QUILPUE', 'VILLA ALEMANA'])]

    # Corte temprano: si no quedo nada tras los filtros (tipico a primera hora,
    # cuando todavia no hay rutas asignadas), NO se regenera el dashboard. Asi
    # se conserva intacto el ultimo reporte valido en vez de pisarlo con ceros.
    if df.empty:
        print("Sin ordenes DOMI en Quilpue/Villa Alemana en este export. "
              "Se conserva el dashboard anterior.")
        return None
    # ========================================================

    # ============== REGLA GSA (cuenta como Completado) ==============
    # Toda orden "No Realizada" cuyo Codigo de Cierre sea "Falla
    # Aprovisionamiento" Y su Area de derivacion sea "GSA" se cuenta como
    # COMPLETADA. Se reetiqueta el Estado ANTES de calcular el RGU para que
    # sume en el panel y para que las generadoras (Alta/Migracion) entren en
    # la logica de Modificacion=0,1.
    # Se normaliza PRIMERO el estado (sin tildes, sin espacios sobrantes) para
    # que un "Completado " o un "COMPLETADO" no rompan las comparaciones.
    df['e'] = df['e'].map(norm_estado)

    # .astype(str) es obligatorio: si el df llega vacio, .map() no ejecuta
    # norm_txt ni una vez y la Serie conserva el dtype float64 que traia la
    # columna en NaN desde el Excel. Sobre float64 el accesor .str revienta con
    # "Can only use .str accessor with string values". Forzando el tipo, la
    # regla GSA funciona igual con 0 filas que con datos reales.
    _c = df['cierre'].map(norm_txt).astype(str)
    _d = df['deriv'].map(norm_txt).astype(str)
    # Match por CONTENIDO, no por igualdad exacta: el portal a veces manda
    # "Falla de Aprovisionamiento", "FALLA APROVISIONAMIENTO CTO" o "GSA - ...".
    es_falla = _c.str.contains('falla', na=False) & _c.str.contains('aprovision', na=False)
    es_gsa = _d.str.contains(r'\bgsa\b', na=False, regex=True)
    es_nr = df['e'].eq('No Realizada')

    mask_gsa = es_nr & es_falla & es_gsa
    n_gsa = int(mask_gsa.sum())
    df.loc[mask_gsa, 'e'] = 'Completado'
    df['gsa'] = mask_gsa.astype(int)   # marca para poder auditarla despues

    # --- diagnostico: para cachar al tiro si la regla no esta pegando ---
    n_nr = int(es_nr.sum())
    print(f"Regla GSA: {n_gsa} de {n_nr} 'No Realizada' reclasificada(s) como Completado")
    if n_nr and not n_gsa:
        casi = df[es_nr & (es_falla | es_gsa)]
        if len(casi):
            print("  OJO: hay No Realizadas que calzan a medias. Revisa como viene escrito:")
            for r in casi.head(8).itertuples():
                print(f"    OT {r.o} | cierre='{r.cierre}' | deriv='{r.deriv}'")
        else:
            valores = sorted(set(f"cierre='{c}' deriv='{d}'"
                                 for c, d in zip(df.loc[es_nr, 'cierre'],
                                                 df.loc[es_nr, 'deriv'])))[:8]
            print("  Ninguna No Realizada calza. Valores presentes en el export:")
            for v in valores:
                print(f"    {v}")
    # -------------------------------------------------------------------
    # ===============================================================

    return df


def generar(path_excel):
    """Genera rutas.json a partir del Excel del portal."""
    df = preparar_df(path_excel)
    if df is None:
        return None

    # RGU (antes de limpiar el nombre del tecnico; usa dir/fecha/sub/pasos)
    df = calcular_rgu(df)

    # limpiar
    df['tec'] = df['tec'].apply(limpiar_tecnico)
    df['hi'] = df['hi'].fillna('').astype(str)
    df['hf'] = df['hf'].fillna('').astype(str)
    df = df.dropna(subset=['lat', 'lng'])           # sin coordenadas no se puede mapear
    df['d'] = df['dir'].astype(str).str.strip() + ', ' + df['ciudad'].astype(str).str.strip()

    # ordenar por tecnico y hora de inicio (vacios al final)
    df['orden_hora'] = df['hi'].replace('', '99:99')
    df = df.sort_values(['tec', 'orden_hora'])

    tecnicos = {}
    for tec, g in df.groupby('tec'):
        filas = []
        for r in g.itertuples():
            item = {'o': r.o, 't': r.t, 'f': r.f, 'hi': r.hi, 'hf': r.hf,
                    'd': r.d, 'lat': round(float(r.lat), 7),
                    'lng': round(float(r.lng), 7),
                    'e': r.e, 'rgu': r.rgu, 'cat': etiqueta(r.t, r.sub, r.pasos)}
            # numero de peticion, solo si el export lo trae (columna opcional)
            pet = str(getattr(r, 'pet', '') or '').strip()
            if pet and pet.lower() not in ('nan', 'none'):
                item['pet'] = pet
            # Motivo de cierre: SOLO en las No Realizadas. En las demas no
            # aporta y multiplicaria por dos el peso del rutas.json, que se
            # baja cada 5 min. Aca sirve para que el aviso de recaida pueda
            # decir por que se cayo antes: no es lo mismo "sin moradores"
            # (llamar al cliente) que "ductos tapados" (escalar a planta).
            # OJO con la regla GSA: esas filas ya vienen con e='Completado'
            # aunque en el Excel eran No Realizada, asi que hay que mirar
            # tambien r.gsa o el motivo nunca se guardaria para ellas.
            if r.e == 'No Realizada' or r.gsa:
                cierre = str(getattr(r, 'cierre', '') or '').strip()
                deriv = str(getattr(r, 'deriv', '') or '').strip()
                if cierre and cierre.lower() not in ('nan', 'none'):
                    item['cie'] = cierre
                if deriv and deriv.lower() not in ('nan', 'none'):
                    item['der'] = deriv
            if r.gsa:
                item['gsa'] = 1      # se cerro por regla GSA, no por el tecnico
            filas.append(item)
        tecnicos[tec] = filas

    # Dia real de los datos (columna Fecha del export), no la hora de corrida.
    # Lo consume alertas_rutas.py para no comparar jornadas distintas.
    f_datos = fecha_de_datos(df)
    hoy = datetime.now().strftime('%d-%m-%Y')
    if f_datos != hoy:
        print(f"*** OJO: el export es de la jornada {f_datos} y hoy es {hoy}. "
              f"El portal todavia no publica las rutas de hoy. ***")

    # ---- direcciones que YA se habian caido antes (campo 'prev' por orden) ----
    # Cruce contra el historico local (historial_dir_estado.json). Va aca, con
    # el rutas.json todavia en memoria, para que el dato viaje adentro y el
    # dashboard no tenga que bajar otro archivo. Blindado: si falla, rutas.json
    # sale igual sin el campo y el boton del dashboard queda en cero.
    try:
        import historial_dir
        historial_dir.marcar_repetidas(tecnicos, f_datos)
        # OJO CON EL ORDEN: primero marcar, DESPUES ingerir. Al reves, las OTs
        # de hoy quedarian en el indice antes de la consulta y TODA la bandeja
        # saldria marcada como "reagendada" contra si misma.
        historial_dir.ingerir_ots(tecnicos, f_datos)
    except Exception as e:
        print(f"[historial] no se pudieron marcar las repetidas: {e}")

    salida = {
        'actualizado': datetime.now().strftime('%d-%m-%Y %H:%M'),
        'fecha_datos': f_datos,
        'tecnicos': tecnicos
    }
    # Las OT que estan en el bucket sin tecnico asignado viajan aparte, solo
    # para que alertas_rutas.py las anote como ya vistas. No se dibujan ni
    # entran a las caidas: el tablero sigue siendo de ordenes con dueno.
    if OTS_FUERA_DOMI:
        salida['fuera_domi'] = OTS_FUERA_DOMI

    with open(ARCHIVO_SALIDA, 'w', encoding='utf-8') as fh:
        json.dump(salida, fh, ensure_ascii=False)

    n_trab = sum(len(v) for v in tecnicos.values())
    rgu_tot = round(sum(w['rgu'] for v in tecnicos.values() for w in v), 1)
    print(f"OK -> {ARCHIVO_SALIDA}: {len(tecnicos)} tecnicos, {n_trab} trabajos, "
          f"{rgu_tot} RGU total ({salida['actualizado']})")

    # ---- espejo en consola del panel "Sin completar" del dashboard ----
    sin_comp = []
    for tec, v in tecnicos.items():
        act = [w for w in v if w['e'] != 'Cancelado']
        if act and not any(w['e'] == 'Completado' for w in act):
            sin_comp.append((tec, len(act)))
    if sin_comp:
        print(f"Sin completadas ({len(sin_comp)}): " +
              ", ".join(f"{t} ({n})" for t, n in sorted(sin_comp)))
    else:
        print("Sin completadas: ninguno, todos tienen al menos una")
    # -------------------------------------------------------------------
    return salida


def subir_github():
    """Sube rutas.json al repo via Contents API (mismo patron de subir_github.py)."""
    import urllib.request
    import urllib.error

    token = os.environ.get('GITHUB_TOKEN')
    if not token:
        print("ERROR: falta la variable de entorno GITHUB_TOKEN")
        sys.exit(1)

    url = (f"https://api.github.com/repos/{GITHUB_USUARIO}/{GITHUB_REPO}"
           f"/contents/{GITHUB_ARCHIVO}")
    headers = {
        'Authorization': f'Bearer {token}',
        'Accept': 'application/vnd.github+json',
        'User-Agent': 'generar-rutas'
    }

    # 1) obtener sha actual si el archivo ya existe
    sha = None
    try:
        req = urllib.request.Request(url + f"?ref={GITHUB_RAMA}", headers=headers)
        with urllib.request.urlopen(req) as r:
            sha = json.loads(r.read())['sha']
    except urllib.error.HTTPError as e:
        if e.code != 404:
            raise

    # 2) subir contenido
    with open(ARCHIVO_SALIDA, 'rb') as fh:
        contenido = base64.b64encode(fh.read()).decode()

    cuerpo = {
        'message': f"actualizacion rutas {datetime.now().strftime('%d-%m %H:%M')}",
        'content': contenido,
        'branch': GITHUB_RAMA
    }
    if sha:
        cuerpo['sha'] = sha

    req = urllib.request.Request(url, data=json.dumps(cuerpo).encode(),
                                 headers=headers, method='PUT')
    with urllib.request.urlopen(req) as r:
        if r.status in (200, 201):
            print("Subido a GitHub OK")
        else:
            print(f"Respuesta inesperada: {r.status}")

ARCHIVO_TOA = "estados_toa.json"


def subir_toa():
    """Sube estados_toa.json al mismo repo y rama que rutas.json.
    Si el archivo no existe o no cambio, no hace nada."""
    import hashlib
    import urllib.request
    import urllib.error

    if not os.path.exists(ARCHIVO_TOA):
        print("[toa] no hay estados_toa.json, se omite")
        return

    token = os.environ.get('GITHUB_TOKEN')
    if not token:
        print("[toa] falta GITHUB_TOKEN, se omite")
        return

    url = (f"https://api.github.com/repos/{GITHUB_USUARIO}/{GITHUB_REPO}"
           f"/contents/{ARCHIVO_TOA}")
    headers = {
        'Authorization': f'Bearer {token}',
        'Accept': 'application/vnd.github+json',
        'User-Agent': 'generar-rutas'
    }

    with open(ARCHIVO_TOA, 'rb') as fh:
        datos = fh.read()

    sha = None
    try:
        req = urllib.request.Request(url + f"?ref={GITHUB_RAMA}", headers=headers)
        with urllib.request.urlopen(req) as r:
            sha = json.loads(r.read())['sha']
    except urllib.error.HTTPError as e:
        if e.code != 404:
            print(f"[toa] no pude leer el sha: {e.code}")
            return

    if sha:
        cabecera = f"blob {len(datos)}\0".encode()
        if hashlib.sha1(cabecera + datos).hexdigest() == sha:
            print("[toa] sin cambios, no se sube")
            return

    cuerpo = {
        'message': f"estados toa {datetime.now().strftime('%d-%m %H:%M')}",
        'content': base64.b64encode(datos).decode(),
        'branch': GITHUB_RAMA
    }
    if sha:
        cuerpo['sha'] = sha

    req = urllib.request.Request(url, data=json.dumps(cuerpo).encode(),
                                 headers=headers, method='PUT')
    try:
        with urllib.request.urlopen(req) as r:
            if r.status in (200, 201):
                print("[toa] estados_toa.json subido OK")
    except Exception as e:
        print(f"[toa] no se pudo subir: {e}")



if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if a != '--subir']
    subir = '--subir' in sys.argv

    path = args[0] if args else excel_mas_reciente(CARPETA_EXCEL)

    # --- snapshot ANTERIOR: se lee ANTES de que generar() sobreescriba
    #     rutas.json, para poder comparar y detectar ordenes caidas.
    try:
        import alertas_rutas
        _previo = alertas_rutas.cargar_json(ARCHIVO_SALIDA)
    except Exception:
        alertas_rutas = None
        _previo = None

    salida = generar(path)   # escribe rutas.json (sin 'alertas'), o None si vacio

    # --- acumulado del dia + inyeccion de la llave 'alertas' en rutas.json ---
    #     Si generar() devolvio None, rutas.json NO se toco: no hay que comparar.
    #     procesar_dia reescribe rutas.json ya con las alertas, ANTES de subir.
    if salida and alertas_rutas is not None:
        alertas_rutas.procesar_dia(_previo, salida, ARCHIVO_SALIDA)

    # --- historico de direcciones con visitas No Realizadas ---
    #     Se alimenta del EXCEL (no del rutas.json) porque ahi esta el Codigo
    #     de Cierre en crudo. Blindado: si falla, no frena nada.
    try:
        import historial_dir
        historial_dir.ingerir(path)
        if subir:
            historial_dir.subir_github()   # solo sube si cambio
    except Exception as e:
        print(f"[historial] se omite este ciclo: {e}")

    if subir:
        subir_github()
        subir_toa()
