#!/usr/bin/env python3
"""Bouwt de kleuterdoelen-database (minimumdoelen + GO!-doelen per vak/onderdeel)
uit de 'Doelenset BaO *.xlsx' bestanden. Enkel kleuterdoelen (t/m 5-6 jaar)."""
import openpyxl, glob, json, re, unicodedata
from collections import OrderedDict

AGE_COLS = {'3-4':13, '4-5':14, '5-6':15}  # kleuterleeftijden

def slug(s):
    s = unicodedata.normalize('NFKD', s).encode('ascii','ignore').decode()
    s = re.sub(r'[^a-zA-Z0-9]+','-', s).strip('-').lower()
    return s

def kleuter_age(r):
    return any(str(r[i])=='True' for i in AGE_COLS.values())

def ages_of(r):
    return [a for a,i in AGE_COLS.items() if str(r[i])=='True']

def clean(v):
    if v is None: return None
    s = str(v).replace('\r\n','\n').strip()
    return s or None

vakken = []
_seen_ids = set()
def uniq_id(base):
    """Garandeer globaal unieke id (bron lijst soms zelfde nummer onder 2 onderdelen)."""
    if base not in _seen_ids:
        _seen_ids.add(base); return base
    i = 2
    while f"{base}~{i}" in _seen_ids:
        i += 1
    nid = f"{base}~{i}"; _seen_ids.add(nid); return nid

for f in sorted(glob.glob("Doelenset BaO *.xlsx")):
    wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
    ws = wb.active
    vaknaam = ws.title.strip()
    vakcode = slug(vaknaam)
    rows = [r for r in ws.iter_rows(values_only=True)]
    data = [r for r in rows[1:] if not all(c is None for c in r)]
    has_ond = any(str(r[0])=='onderwerp' for r in data)
    group_type = 'onderwerp' if has_ond else 'subthema'

    onderdelen = OrderedDict()  # naam -> {min:{}, go:{}}
    cur = None
    def ensure(naam):
        if naam not in onderdelen:
            onderdelen[naam] = {'naam':naam, 'min_map':OrderedDict(), 'go_map':OrderedDict()}
        return onderdelen[naam]

    for r in data:
        t = str(r[0])
        if t == group_type:
            cur = clean(r[2]) or 'Algemeen'
            ensure(cur)
            continue
        if cur is None:
            cur = 'Algemeen'; ensure(cur)
        if t == 'minimumdoel' and r[3] is not None:
            nummer = clean(r[3]); tekst = clean(r[4])
            if not tekst: continue
            extra = clean(r[5])
            node = ensure(cur)
            key = nummer
            if key not in node['min_map']:
                node['min_map'][key] = {
                    'id': uniq_id(f"{vakcode}:min:{nummer}"),
                    'nummer': nummer, 'tekst': tekst,
                    'detail': extra,
                }
        elif t == 'doelzin' and (r[3] is not None or kleuter_age(r)):
            tekst = clean(r[2])
            if not tekst: continue
            gonr = clean(r[1])
            konum = clean(r[3])
            node = ensure(cur)
            key = tekst.lower()
            if key in node['go_map']:
                g = node['go_map'][key]
                if gonr and gonr not in g['go_nrs']: g['go_nrs'].append(gonr)
                for a in ages_of(r):
                    if a not in g['leeftijden']: g['leeftijden'].append(a)
                if konum:
                    for k in konum.split(','):
                        k=k.strip()
                        if k and k not in g['min_refs']: g['min_refs'].append(k)
            else:
                node['go_map'][key] = {
                    'id': uniq_id(f"{vakcode}:go:{gonr or slug(tekst)[:20]}"),
                    'go_nrs': [gonr] if gonr else [],
                    'tekst': tekst,
                    'leeftijden': ages_of(r),
                    'min_refs': [k.strip() for k in konum.split(',')] if konum else [],
                }
    # leeftijd per minimumdoel-nummer afleiden uit de gekoppelde GO!-doelen
    ages_by_num = {}  # KO nummer -> set(leeftijden)
    for node in onderdelen.values():
        for g in node['go_map'].values():
            for k in g.get('min_refs', []):
                ages_by_num.setdefault(k, set()).update(g['leeftijden'])

    # finalize
    order={'3-4':0,'4-5':1,'5-6':2}
    out_onderdelen = []
    for naam, node in onderdelen.items():
        mins = list(node['min_map'].values())
        gos = list(node['go_map'].values())
        if not mins and not gos: continue
        for g in gos:
            g['leeftijden'].sort(key=lambda a:order.get(a,9))
        for m in mins:
            ages = ages_by_num.get(m['nummer'], set())
            m['leeftijden'] = sorted(ages, key=lambda a:order.get(a,9))
        out_onderdelen.append({'naam':naam,'minimumdoelen':mins,'go_doelen':gos})
    vakken.append({'naam':vaknaam,'code':vakcode,'onderdelen':out_onderdelen})
    wb.close()

db = {'versie':1, 'beschrijving':'Kleuterdoelen GO! (t/m 5-6 jaar) per vak en onderdeel','vakken':vakken}
with open('doelen_database.json','w',encoding='utf-8') as fh:
    json.dump(db, fh, ensure_ascii=False, indent=1)

# stats
tot_min=tot_go=0
for v in vakken:
    vm=sum(len(o['minimumdoelen']) for o in v['onderdelen'])
    vg=sum(len(o['go_doelen']) for o in v['onderdelen'])
    tot_min+=vm; tot_go+=vg
    print(f"{v['naam']:28s} onderdelen={len(v['onderdelen']):2d}  min={vm:3d}  go={vg:3d}")
print(f"\nTOTAAL: {len(vakken)} vakken, {tot_min} minimumdoelen, {tot_go} GO!-doelen")
