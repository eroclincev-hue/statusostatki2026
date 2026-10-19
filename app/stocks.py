#!/usr/bin/env python3
"""Остатки FBS на складах фулфилментов -> stocks.json.

Ероклинцев: все карточки. Аванти: только TESFE (артикулы из ARTICLE_FILTER в fetch.py).
Склады раскладываются по ФФ по имени склада. Сырые остатки по каждому складу сохраняются,
сводка «на ФФ» считается на странице.
"""
import json, os, sys, time, datetime, re
from concurrent.futures import ThreadPoolExecutor
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fetch import call, env, MSK, ARTICLE_FILTER, BRAND_OVERRIDE, CONTENT, HIDE_ARTICLES

HERE = os.path.dirname(os.path.abspath(__file__))
MP = 'https://marketplace-api.wildberries.ru'

# (регулярка по имени склада, ФФ). Склады, не попавшие ни в одну группу, не показываются.
FF_RULES = [
    (r'^(граф)', 'Граф'),
    (r'ff2b|фф2б', 'ФФ2Б'),
    (r'^(фулфилмент №1|фф№1|фф №1)', 'ФФ №1'),
    (r'^самара один квадрат', 'Самара один квадрат'),
    (r'^onyx', 'ONYX FF'),
    (r'^доски', 'Доски'),
    (r'^флаконы', 'Флаконы'),
    (r'^3pl', '3PL'),
    (r'tesfe', 'TESFE'),
]
EXCLUDE = re.compile(r'sellerkeller|^dbs|термокружка', re.I)


def ff_of(name):
    if EXCLUDE.search(name):
        return None
    for rx, ff in FF_RULES:
        if re.search(rx, name, re.I):
            return ff
    return None


def all_cards(tok):
    out, cur = [], {'limit': 100}
    while True:
        r = call(tok, 'POST', CONTENT, {'settings': {'cursor': cur, 'filter': {'withPhoto': -1}}})
        out += r['cards']
        c = r['cursor']
        if c['total'] < 100:
            return out
        cur = {'limit': 100, 'updatedAt': c['updatedAt'], 'nmID': c['nmID']}


def stocks_for(tok, wid, skus):
    res = {}
    for i in range(0, len(skus), 1000):
        r = call(tok, 'POST', f'{MP}/api/v3/stocks/{wid}', {'skus': skus[i:i + 1000]}, tries=20)
        for s in (r or {}).get('stocks', []):
            res[s['sku']] = s.get('amount') or 0
    return res


SALES_DAYS = 14
CANCEL_WB = {'declined_by_client', 'canceled', 'canceled_by_client', 'canceled_by_missed_call'}


def fbs_sales(e, days):
    """{артикул в нижнем регистре: заказов FBS за days дней}"""
    out = {}
    now = int(time.time())
    for cab, key in [('erokl', 'TOKEN_2'), ('avanti', 'TOKEN_1')]:
        tok = e.get(key)
        if not tok: continue
        orders, start = [], now - days * 86400
        while start < now:  # окна до 7 дней
            end = min(start + 7 * 86400, now); nxt = 0
            while True:
                r = call(tok, 'GET', f'{MP}/api/v3/orders?limit=1000&next={nxt}&dateFrom={start}&dateTo={end}')
                orders += r.get('orders', [])
                if len(r.get('orders', [])) < 1000: break
                nxt = r['next']
            start = end
        orders += [o for o in call(tok, 'GET', f'{MP}/api/v3/orders/new').get('orders', [])]
        seen, uniq = set(), []
        for o in orders:
            if o['id'] not in seen: seen.add(o['id']); uniq.append(o)
        st = {}
        ids = [o['id'] for o in uniq]
        for i in range(0, len(ids), 1000):
            for x in call(tok, 'POST', f'{MP}/api/v3/orders/status', {'orders': ids[i:i + 1000]}).get('orders', []):
                st[x['id']] = x
        for o in uniq:
            x = st.get(o['id'], {})
            if x.get('supplierStatus') == 'cancel': continue
            if x.get('wbStatus') in CANCEL_WB and x.get('supplierStatus') in (None, 'new', 'confirm'): continue
            a = (o.get('article') or '').lower()
            out[a] = out.get(a, 0) + 1
    return out


def main(out_path):
    e = env()
    cache_path = os.path.join(HERE, 'cards_cache.json')
    ccache = json.load(open(cache_path)) if os.path.exists(cache_path) else {}
    items, whs, errors = [], [], []
    for cab, cab_name, key in [('erokl', 'ИП Ероклинцев', 'TOKEN_2'), ('avanti', 'Аванти', 'TOKEN_1')]:
        tok = e[key]
        # карточки: Ероклинцев небольшой, тянем каждый раз; Аванти 18 тыс. — кешируем на сутки
        c = ccache.get(cab)
        if not c or c['day'] != datetime.datetime.now(MSK).strftime('%Y-%m-%d'):
            try:
                cards = all_cards(tok)
            except Exception as ex:
                errors.append(f'{cab_name} карточки: {ex}')
                cards = (c or {}).get('cards', [])
            af = ARTICLE_FILTER.get(cab)
            if af:
                cards = [cd for cd in cards if any(m in (cd.get('vendorCode') or '').lower() for m in af)]
            slim = [{'nm': cd['nmID'], 'vc': cd.get('vendorCode') or '', 'brand': (cd.get('brand') or '').strip(),
                     'title': cd.get('title') or '', 'subject': cd.get('subjectName') or '',
                     'skus': [s for sz in cd.get('sizes', []) for s in sz.get('skus', [])]} for cd in cards]
            c = {'day': datetime.datetime.now(MSK).strftime('%Y-%m-%d'), 'cards': slim}
            ccache[cab] = c
        af = ARTICLE_FILTER.get(cab)
        sku2item = {}
        for cd in c['cards']:
            if cd['vc'].lower() in HIDE_ARTICLES: continue
            brand = BRAND_OVERRIDE.get(cd['vc']) or ('TESFE' if af else cd['brand']) or 'Без бренда'
            it = {'cab': cab, 'cabName': cab_name, 'nm': cd['nm'], 'article': cd['vc'], 'brand': brand,
                  'title': cd['title'], 'subject': cd['subject'], 'st': {}, 'bc': (cd['skus'] or [''])[0]}
            items.append(it)
            for s in cd['skus']:
                sku2item[s] = it
        try:
            wl = call(tok, 'GET', f'{MP}/api/v3/warehouses')
        except Exception as ex:
            errors.append(f'{cab_name} склады: {ex}'); continue
        wl = [(w['id'], w['name'], ff_of(w['name'])) for w in wl]
        wl = [w for w in wl if w[2]]
        skus = list(sku2item)
        with ThreadPoolExecutor(4) as ex:
            futs = {w[0]: ex.submit(stocks_for, tok, w[0], skus) for w in wl}
        for wid, name, ff in wl:
            try:
                st = futs[wid].result()
            except Exception as er:
                errors.append(f'{cab_name} {name}: {er}'); continue
            whs.append({'cab': cab, 'id': wid, 'name': name, 'ff': ff})
            for sku, amt in st.items():
                if amt and sku in sku2item:
                    d = sku2item[sku]['st']
                    d[str(wid)] = d.get(str(wid), 0) + amt
        print(cab_name, 'товаров', len(c['cards']), 'складов', len(wl), file=sys.stderr)
    json.dump(ccache, open(cache_path, 'w'), ensure_ascii=False)
    # реальные остатки ФФ по выгрузкам ВМС
    real_meta = {}
    import ffreal
    try:
        real = ffreal.compute()
    except Exception as ex:
        errors.append(f'выгрузки ВМС: {ex}'); real = {}
    by_art = {}
    for it in items:
        by_art.setdefault(it['article'].lower(), it)
    for ff, v in real.items():
        real_meta[ff] = {'time': v['time'], 'file': v['file'], 'unmatched': v['unmatched'], 'exclusive': v.get('exclusive', False), 'label': v.get('label'), 'fbs_ok': v.get('fbs_ok', False), 'file': v['file'], 'components': v.get('components', [])}
        for r in v['rows']:
            if r['article'].lower() in HIDE_ARTICLES: continue
            it = by_art.get(r['article'].lower())
            if not it:
                it = {'cab': 'erokl', 'cabName': 'ИП Ероклинцев', 'nm': None, 'article': r['article'],
                      'brand': BRAND_OVERRIDE.get(r['article'], 'Без бренда'), 'title': r['title'], 'subject': '', 'st': {}}
                items.append(it); by_art[r['article'].lower()] = it
            it.setdefault('real', {})[ff] = {'now': r['now'], 'base': r['base'], 'orders': r['orders'], 'fbo': r.get('fbo')}
    # продажи FBS за SALES_DAYS дней (оба кабинета, все склады), без отменённых до сборки
    sales = {}
    try:
        sales = fbs_sales(e, SALES_DAYS)
    except Exception as ex:
        errors.append(f'продажи FBS: {ex}')
    for it in items:
        it['sold'] = sales.get(it['article'].lower(), 0)
        if it['cab'] == 'erokl':
            cu = ffreal.canister_units(it['article'])
            if cu:
                base = next(a for f, _, a in ffreal.CANISTER_FAMILIES if f == cu[0])
                it['fam'] = [base, cu[1]]
    # справочник баркодов / артикулов WB и Ozon от Лили: дополняет то, чего нет в карточках
    rp = os.path.join(HERE, 'ref_codes.json')
    ref = json.load(open(rp)) if os.path.exists(rp) else {}
    for it in items:
        rs = ref.get(it['article'].lower())
        if not rs: continue
        r = next((x for x in rs if (x['le'] == 'ИП Ероклинцев') == (it['cab'] == 'erokl')), rs[0])
        if not it.get('bc') and r['bc']: it['bc'] = r['bc']
        if not it.get('nm') and r['nm']:
            try: it['nm'] = int(float(r['nm']))
            except ValueError: pass
        oz = [x['oz'] for x in rs if x['oz']]
        if oz: it['oz'] = oz[0]
    data = {'updated': datetime.datetime.now(MSK).strftime('%d.%m.%Y %H:%M'),
            'warehouses': whs, 'items': [i for i in items if i['st'] or i.get('real') or i['sold']],
            'real': real_meta, 'salesDays': SALES_DAYS, 'errors': errors,
            'costs': json.load(open(os.path.join(HERE, 'costs.json'))) if os.path.exists(os.path.join(HERE, 'costs.json')) else {}}
    json.dump(data, open(out_path, 'w'), ensure_ascii=False)
    print('позиций с остатком', len(data['items']), 'ошибок', len(errors), file=sys.stderr)
    for er in errors[:10]: print(' ', er, file=sys.stderr)


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'stocks.json'))
