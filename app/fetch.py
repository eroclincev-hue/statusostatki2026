#!/usr/bin/env python3
"""Выгрузка неотгруженных поставок FBO (WB supplies-api) по двум кабинетам -> data.json.

Токены берутся из .env рядом со скриптом:
  TOKEN_1 = Аванти, TOKEN_2 = ИП Ероклинцев.
Детали и состав поставок кешируются по (id, updatedDate), чтобы не упираться в лимит WB (~10 запросов / ~20 с).
"""
import json, os, sys, time, urllib.request, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
B = 'https://supplies-api.wildberries.ru/api/v1'
# 1 Не запланировано, 2 Запланировано, 3 Отгрузка разрешена,
# 4 Идёт приёмка, 5 Принято, 6 Отгружено на воротах
STATUSES = {2: 'Запланировано', 3: 'Отгрузка разрешена', 4: 'Идёт приёмка'}
BOX = {0: 'Виртуальная', 1: 'Без коробов', 2: 'Короба', 5: 'Монопаллеты', 6: 'Суперсейф'}
CABINETS = [('erokl', 'ИП Ероклинцев', 'TOKEN_2'), ('avanti', 'Аванти', 'TOKEN_1')]
# Для кабинета показываем только поставки, где есть артикул с этой меткой (товар TESFE в Аванти помечен «_тф»)
ARTICLE_FILTER = {'avanti': ['_тф', 'пылесос_вертикальный_gt804_белый', 'пылесос_вертикальный_gt804_черный',
                             'пылесос_вертикальный_gt917', 'пылесос_вертикальный_nh016']}
# Бренд по артикулу, когда в карточке WB он не заполнен (со слов Лили)
BRAND_OVERRIDE = {'HF-2000': 'ALURAM', 'HF-2501': 'ALURAM', 'VF-1500': 'ALURAM'}
# С какого ФФ идёт поставка FBO (по таблице Лили)
SUPPLY_FF = {
    41605037: 'ФФ2Б', 41605041: 'ФФ №1', 41605016: 'ФФ №1', 41605001: 'ФФ2Б', 41605029: 'ФФ2Б',
    41615296: 'ФФ2Б', 41562436: 'ФФ №1', 41615329: 'ФФ2Б', 41623432: 'ФФ Краснодар', 41643794: 'ONYX FF',
    41622707: 'ФФ Краснодар', 41643589: 'ФФ2Б', 41615363: 'ФФ2Б', 41656409: 'Флаконы', 41656637: 'ФФ2Б',
    41605020: 'ФФ2Б', 41605027: 'ФФ №1', 41605028: 'ФФ2Б',
    41659459: 'Самара один квадрат', 41659541: 'Самара один квадрат', 41659593: 'Самара один квадрат',
    41660463: 'Граф', 41660522: 'Граф', 41660714: 'Граф',
    41674147: 'ФФ №1', 41674160: 'ФФ2Б', 41674040: 'ФФ2Б', 41674063: 'ФФ2Б',
    41540350: 'ФФ №1', 41539317: 'ФФ №1', 41571557: 'ФФ №1', 41585261: '3PL', 41550332: 'Самара один квадрат',
    41570920: 'ФФ2Б', 41589250: 'ФФ2Б', 41550383: 'Самара один квадрат', 41576705: 'ФФ Казань',
    41550424: 'Самара один квадрат', 41501471: 'ONYX FF', 41509141: 'ONYX FF',
}
# Артикулы, которые не показываем нигде на сайте
HIDE_ARTICLES = {'губка_тест', 'наушники_с_экраном'}
CONTENT = 'https://content-api.wildberries.ru/content/v2/get/cards/list'
MSK = datetime.timezone(datetime.timedelta(hours=3))


def env():
    """Токены: из переменных окружения (секреты GitHub) или из файла .env рядом со скриптом."""
    e = {k: v for k, v in os.environ.items() if k.startswith('TOKEN_') and v}
    p = os.path.join(HERE, '.env')
    if os.path.exists(p):
        for l in open(p):
            if '=' in l:
                k, v = l.strip().split('=', 1); e.setdefault(k, v)
    return e


def call(token, method, url, body=None, tries=8):
    for a in range(tries):
        req = urllib.request.Request(url, method=method, headers={'Authorization': token, 'Content-Type': 'application/json'},
                                     data=json.dumps(body).encode() if body is not None else None)
        try:
            r = urllib.request.urlopen(req, timeout=90)
            return json.loads(r.read() or b'null')
        except urllib.error.HTTPError as e:
            if e.code == 429 or e.code >= 500:
                wait = int(e.headers.get('X-Ratelimit-Retry') or 0) or int(e.headers.get('X-Ratelimit-Reset') or 0) or 5
                time.sleep(min(max(wait, 1), 30) + 0.5)
                continue
            raise RuntimeError(f'{method} {url} -> {e.code} {e.read()[:300]!r}')
        except Exception:
            time.sleep(3)
    raise RuntimeError(f'{method} {url}: не удалось после {tries} попыток')


def list_all(token):
    out, off = [], 0
    while True:
        page = call(token, 'POST', f'{B}/supplies?limit=1000&offset={off}',
                    {'statusIDs': list(STATUSES)})
        out += page or []
        if not page or len(page) < 1000:
            return out
        off += 1000


def brands_for(token, nms, bcache):
    for nm in nms:
        if str(nm) in bcache: continue
        r = call(token, 'POST', CONTENT, {'settings': {'cursor': {'limit': 100}, 'filter': {'textSearch': str(nm), 'withPhoto': -1}}})
        for c in (r or {}).get('cards', []):
            bcache[str(c['nmID'])] = c.get('brand') or ''
        bcache.setdefault(str(nm), '')


def main(out_path):
    e = env()
    cache_path = os.path.join(HERE, 'cache.json')
    cache = json.load(open(cache_path)) if os.path.exists(cache_path) else {}
    new_cache, supplies, errors = {}, [], []
    bcache = cache.get('_brands', {})
    today = datetime.datetime.now(MSK).strftime('%Y-%m-%d')
    for cab, cab_name, key in CABINETS:
        tok = e.get(key)
        if not tok:
            errors.append(f'{cab_name}: нет токена'); continue
        try:
            lst = list_all(tok)
        except Exception as ex:
            errors.append(f'{cab_name}: {ex}'); continue
        # статусы из STATUSES; просроченные запланированные / с разрешённой отгрузкой убираем
        lst = [x for x in lst if x.get('statusID') in STATUSES
               and not (x['statusID'] in (2, 3) and x.get('supplyDate') and x['supplyDate'][:10] < today)]
        print(cab_name, 'поставок:', len(lst), file=sys.stderr)
        for x in lst:
            pre = not x.get('supplyID')
            sid = x['preorderID'] if pre else x['supplyID']
            ck = f"{cab}:{'p' if pre else 's'}{sid}"
            c = cache.get(ck)
            if not c or c.get('upd') != x.get('updatedDate'):
                q = '?isPreorderID=true' if pre else ''
                try:
                    det = call(tok, 'GET', f'{B}/supplies/{sid}{q}')
                    goods = []
                    off = 0
                    while True:
                        g = call(tok, 'GET', f'{B}/supplies/{sid}/goods?limit=1000&offset={off}'
                                 + ('&isPreorderID=true' if pre else ''))
                        goods += g or []
                        if not g or len(g) < 1000: break
                        off += 1000
                except Exception as ex:
                    errors.append(f'{cab_name} {sid}: {ex}')
                    if c: new_cache[ck] = c
                    continue
                det = {k: v for k, v in (det or {}).items() if k != 'phone'}
                c = {'upd': x.get('updatedDate'), 'det': det, 'goods': goods}
            new_cache[ck] = c
            d, goods = c['det'] or {}, c['goods'] or []
            af = ARTICLE_FILTER.get(cab)
            if af and not any(m in (gg.get('vendorCode') or '').lower() for gg in goods for m in af):
                continue
            try:
                brands_for(tok, {gg.get('nmID') for gg in goods if gg.get('nmID')}, bcache)
            except Exception as ex:
                errors.append(f'{cab_name} бренды: {ex}')

            def brand(gg):
                if (gg.get('vendorCode') or '') in BRAND_OVERRIDE:
                    return BRAND_OVERRIDE[gg['vendorCode']]
                if af and any(m in (gg.get('vendorCode') or '').lower() for m in af):
                    return 'TESFE'  # в Аванти бренд в карточках не заполнен, TESFE узнаём по артикулу
                return (bcache.get(str(gg.get('nmID'))) or '').strip() or 'Без бренда'
            supplies.append({
                'cab': cab, 'cabName': cab_name, 'ff': SUPPLY_FF.get(x.get('supplyID')) or '',
                'id': x.get('supplyID'), 'preorderID': x.get('preorderID'),
                'status': x['statusID'], 'statusName': STATUSES[x['statusID']],
                'box': BOX.get(d.get('boxTypeID', x.get('boxTypeID')), str(d.get('boxTypeID'))),
                'created': x.get('createDate'), 'planned': x.get('supplyDate'), 'updated': x.get('updatedDate'),
                'warehouse': d.get('warehouseName') or '', 'transit': d.get('transitWarehouseName') or '',
                'qty': d.get('quantity') if d.get('quantity') is not None else sum(gg.get('quantity') or 0 for gg in goods),
                'accepted': d.get('acceptedQuantity') or 0, 'ready': d.get('readyForSaleQuantity') or 0,
                'factDate': d.get('factDate'),
                'acceptanceCost': d.get('acceptanceCost'), 'coef': d.get('paidAcceptanceCoefficient'),
                'storageCoef': d.get('storageCoef'), 'deliveryCoef': d.get('deliveryCoef'),
                'goods': [{'barcode': gg.get('barcode'), 'article': gg.get('vendorCode'), 'nm': gg.get('nmID'),
                           'size': gg.get('techSize'), 'color': gg.get('color'), 'qty': gg.get('quantity') or 0,
                           'accepted': gg.get('acceptedQuantity') or 0, 'brand': brand(gg)}
                          for gg in goods if (gg.get('vendorCode') or '').lower() not in HIDE_ARTICLES],
            })
        new_cache['_brands'] = bcache
        json.dump(new_cache, open(cache_path, 'w'), ensure_ascii=False)
    data = {'updated': datetime.datetime.now(MSK).strftime('%d.%m.%Y %H:%M'),
            'supplies': supplies, 'errors': errors}
    json.dump(data, open(out_path, 'w'), ensure_ascii=False)
    print('итого', len(supplies), 'ошибок', len(errors), file=sys.stderr)
    for er in errors[:10]: print(' ', er, file=sys.stderr)


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'data.json'))
