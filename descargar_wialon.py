#!/usr/bin/env python3
"""
descargar_wialon.py — desplazamiento diario de la flota, del GPS de Wialon.

Alimenta las filas "Desplaz. skidder" y "Desplaz. shovel" del Informe de Faena.

INCREMENTAL A PROPÓSITO. Bajar el mes completo de las 25 unidades toma ~20 min y el pipeline
entero corre en 2. Cada corrida baja solo los últimos VENTANA_DIAS y los MEZCLA sobre
`wialon_km.json`, que va versionado en el repo. Un día ya guardado se sobrescribe con la
lectura nueva (los mensajes pueden llegar tarde), y se conservan el mes en curso y el anterior.

SOLO SG Y HM (gerencia 2026-07-31). Son las dos máquinas para las que DESPLAZARSE ES EL
TRABAJO: el skidder arrastra a cancha, la shovel acomoda por el frente. Quedan fuera los AM
(Timbermax/Falcon) porque son ASISTENCIA —en julio dieron 0,0 y 0,3 km/día contra 8-34 de los
skidders— y los FM (feller), que cortan avanzando, que es otro trabajo.

El nombre de la unidad en Wialon ya trae faena y tipo:
    CEMA-Z.CONS-MILL-MILLALEMU 1.3-SKIDDER-T5869
…así que el crosswalk sale de ahí y no de un CSV a mano que hay que mantener.

NO filtra los traslados en cama baja: guarda el crudo y el informe filtra al calcular, para
poder auditar qué se descartó. En julio fueron 13 días, y solos inflaban a M1.1 de 20,1 a
33,5 km/día.

Requiere WIALON_TOKEN en el entorno. Sin él no hace nada y el informe muestra "rep.".
"""
import json
import math
import os
import re
import sys
import urllib.parse
import urllib.request
from datetime import datetime, timedelta
from pathlib import Path

BASE = Path(__file__).parent
SALIDA = BASE / "wialon_km.json"
API = "https://hst-api.wialon.us/wialon/ajax.html"

# Cuántos días hacia atrás se refrescan en cada corrida. 3 da margen para mensajes que llegan
# tarde y para una corrida que se saltó, sin pagar el mes entero.
VENTANA_DIAS = 3

# Wialon nombra al SG como SKIDDER o GRAPPLE según la unidad (el T5882 de M1.1 es el SG-13).
TIPOS = {'SKIDDER', 'GRAPPLE', 'SHOVEL'}
RX_UNIDAD = re.compile(
    r'MILLALEMU ?([\d.]+)\s*-\s*(SKIDDER|GRAPPLE|SHOVEL)-T(\d+)$', re.I)

# Chile es UTC-4. Las fechas del informe son días de faena, no días UTC.
TZ_OFFSET = 4 * 3600


def call(svc, params, sid=None, timeout=180):
    d = {'svc': svc, 'params': json.dumps(params)}
    if sid:
        d['sid'] = sid
    url = API + '?' + urllib.parse.urlencode(d)
    return json.loads(urllib.request.urlopen(url, timeout=timeout).read())


def haversine_km(lat1, lon1, lat2, lon2):
    R = 6371.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    x = (math.sin(math.radians(lat2 - lat1) / 2) ** 2
         + math.cos(p1) * math.cos(p2) * math.sin(math.radians(lon2 - lon1) / 2) ** 2)
    return 2 * R * math.asin(math.sqrt(x))


def dia_faena(ts):
    """Día de faena de un timestamp (hora de Chile)."""
    return datetime.fromtimestamp(ts - TZ_OFFSET).strftime('%Y-%m-%d')


def main():
    token = os.environ.get('WIALON_TOKEN')
    if not token:
        print("  ⚠️  WIALON_TOKEN no está; el desplazamiento queda 'por reportar'")
        return 0

    previo = {}
    if SALIDA.exists():
        try:
            previo = json.loads(SALIDA.read_text(encoding='utf-8'))
        except Exception as e:
            print(f"  ⚠️  wialon_km.json no legible ({e}); se reconstruye desde cero")

    try:
        sid = call('token/login', {'token': token})['eid']
    except Exception as e:
        print(f"  ⚠️  Wialon no responde al login ({e}); se conserva el archivo anterior")
        return 0

    try:
        items = call('core/search_items', {
            'spec': {'itemsType': 'avl_unit', 'propName': 'sys_name',
                     'propValueMask': '*', 'sortType': 'sys_name'},
            'force': 1, 'flags': 1, 'from': 0, 'to': 0}, sid).get('items') or []
    except Exception as e:
        print(f"  ⚠️  Wialon no listó unidades ({e}); se conserva el archivo anterior")
        return 0

    flota = [(u, RX_UNIDAD.search(u['nm'])) for u in items]
    flota = [(u, m) for u, m in flota if m]
    if not flota:
        print("  ⚠️  el token no ve ninguna máquina SG/HM de Millalemu")
        return 0

    hoy = datetime.now()
    desde = (hoy - timedelta(days=VENTANA_DIAS)).replace(hour=4, minute=0, second=0, microsecond=0)
    t0, t1 = int(desde.timestamp()), int(hoy.timestamp())
    print(f"  Wialon: {len(flota)} máquinas SG/HM · refrescando desde {desde:%d-%m}")

    nuevos = 0
    for u, m in flota:
        faena, tipo, disp = f"M{m.group(1)}", m.group(2).upper(), 'T' + m.group(3)
        reg = previo.setdefault(disp, {'faena': faena, 'tipo': tipo, 'unidad': u['nm'], 'dias': {}})
        reg['faena'], reg['tipo'], reg['unidad'] = faena, tipo, u['nm']
        try:
            call('messages/unload', {}, sid)
        except Exception:
            pass
        try:
            r = call('messages/load_interval', {
                'itemId': u['id'], 'timeFrom': t0, 'timeTo': t1,
                'flags': 0, 'flagsMask': 0, 'loadCount': 0xffffffff}, sid)
        except Exception as e:
            print(f"    ! {disp}: {e}")
            continue
        pts = [(x['t'], x['pos']['y'], x['pos']['x'], x['pos'].get('s', 0))
               for x in (r.get('messages') or []) if x.get('pos')]
        frescos = {}
        for a, b in zip(pts, pts[1:]):
            dh = (b[0] - a[0]) / 3600
            if dh <= 0 or dh > 1:      # hueco largo: no se puede atribuir a un día
                continue
            km = haversine_km(a[1], a[2], b[1], b[2])
            e = frescos.setdefault(dia_faena(b[0]), {'km': 0.0, 'h_mov': 0.0, 'n': 0})
            if km < 2:                 # salto absurdo entre dos puntos = ruido de GPS
                e['km'] += km
            if (b[3] or 0) > 1 or km > 0.005:
                e['h_mov'] += dh
            e['n'] += 1
        # Los días recién leídos PISAN a los guardados: un mensaje puede llegar tarde y la
        # lectura nueva es siempre la más completa.
        reg['dias'].update(frescos)
        nuevos += len(frescos)

    try:
        call('core/logout', {}, sid)
    except Exception:
        pass

    # Se conservan el mes en curso y el anterior: el informe nunca mira más atrás y el archivo
    # no crece sin control en un repo.
    corte = (hoy.replace(day=1) - timedelta(days=1)).replace(day=1).strftime('%Y-%m-%d')
    for reg in previo.values():
        reg['dias'] = {d: v for d, v in reg['dias'].items() if d >= corte}

    SALIDA.write_text(json.dumps(previo, sort_keys=True), encoding='utf-8')
    dias_tot = sum(len(r['dias']) for r in previo.values())
    print(f"  ✅ wialon_km.json — {len(previo)} máquinas, {dias_tot} días-máquina "
          f"({nuevos} refrescados)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
