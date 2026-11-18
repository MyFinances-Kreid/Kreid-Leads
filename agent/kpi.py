# Onglet « KPI Distance » — exécuté dans le contexte de build.py après build_base.py
# Sources Python déjà chargées : _LEADS (source, nom, date, clé tél.), _best (fiches CRM retenues), _EXTRA (montant simulé…),
# _rows_bd (lignes de bordereaux rapprochées), onglets fournisseurs (prix). Valeurs figées à chaque génération (2 fois par jour).
import statistics as _st

_PRIXCOL = {'Weendeal': (W, 3, 7, 11), 'Finanzen': (Fz, 2, 10, 13), 'Qwamplify': (Q, 2, 6, 21), 'Kreid.fr': (Kf, 2, 4, 6)}
_price = {}
for _src, (_ws, _r0, _ct, _cp) in _PRIXCOL.items():
    for _r in range(_r0, _ws.max_row + 1):
        _t = re.sub(r'\D', '', str(_ws.cell(_r, _ct).value or ''))[-9:]
        _v = _ws.cell(_r, _cp).value
        if _t and isinstance(_v, (int, float)): _price[(_src, _t)] = float(_v)

_pay = {}   # clé -> [date 1er paiement, CA, MAF]
for _b in _rows_bd:
    _k = _b[6]
    if not _k: continue
    p = _pay.setdefault(_k, [None, 0.0, 0.0])
    if isinstance(_b[1], _dt.datetime) and str(_b[0]).startswith('MY'): p[0] = min(p[0], _b[1]) if p[0] else _b[1]
    p[1] += _b[5] or 0; p[2] = max(p[2], _b[4] or 0)

STEPS_K = ['recep', 'simu', 'dossier', 'depot', 'accord', 'paye']
_L = []
_seen = set()
for src, nom, d, tel in _LEADS:
    if (src, tel) in _seen and tel: continue
    _seen.add((src, tel))
    c = _best[tel][1] if tel in _best else None
    ex = _EXTRA.get(tel, (None,) * 4)
    pay = _pay.get(tel)
    acc = c[9] if c else None
    if pay and pay[0] and not acc: acc = pay[0]
    _L.append(dict(src=src, date=d, tel=tel, cost=_price.get((src, tel), 0.0), coll=(c[2] if c else '(absent du CRM)'),
                   recep=c[5] if c else None, simu=c[6] if c else None, dossier=c[7] if c else None, depot=c[8] if c else None,
                   accord=acc, refus=bool(c and c[10]), paye=pay[0] if pay else None, ca=pay[1] if pay else 0.0, maf=pay[2] if pay else 0.0,
                   mts=ex[0], incrm=bool(c)))

K = wb.create_sheet('KPI Distance')
setup(K); K.sheet_properties.tabColor = GREEN
band(K, 1, 'KPI PÔLE DISTANCE — entonnoir, délais, conseillers, cohortes, back office',
     f"Calculé le {_dt.datetime.now():%d/%m/%Y à %H:%M} à partir du CRM (extraction du {TODAY:%d/%m/%Y} ou antérieure), des onglets fournisseurs et des bordereaux Budgeo · cible : 6 % des leads financés", 14)


def _hdr(r, c0, labels, widths=None):
    for i, t in enumerate(labels):
        x = K.cell(r, c0 + i, t); x.font = font(b=True, color='FFFFFF'); x.fill = fill(HEAD); x.border = box
        x.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    K.row_dimensions[r].height = 30


def _row(r, c0, vals, fmts, bold=False, shade=False):
    for i, v in enumerate(vals):
        x = K.cell(r, c0 + i, v); x.border = box; x.font = font(b=bold)
        if fmts[i]: x.number_format = fmts[i]
        if shade: x.fill = fill(LIGHT)
        if i: x.alignment = Alignment(horizontal='center')


def _pct(a, b): return (a / b) if b else None
def _med(v): return _st.median(v) if v else None
def _days(a, b):
    return (b - a).days if (isinstance(a, _dt.datetime) and isinstance(b, _dt.datetime) and b >= a) else None

PCT, EUR, NB, J = '0.0%', '#,##0 €', '#,##0', '0" j"'
PAYANTS = ('Weendeal', 'Finanzen', 'Qwamplify')

# ---------- 1. Synthèse
r = 4; section(K, r, 1, '1 · Synthèse (leads payants Weendeal + Finanzen + Qwamplify)', 14); r += 1
pay_l = [x for x in _L if x['src'] in PAYANTS]
n = len(pay_l); n_acc = sum(1 for x in pay_l if x['accord']); n_pay = sum(1 for x in pay_l if x['paye'])
cost = sum(x['cost'] for x in pay_l); ca = sum(x['ca'] for x in pay_l); maf = sum(x['maf'] for x in pay_l)
l30 = sum(1 for x in pay_l if (TODAY - x['date']).days <= 30)
cards = [('Leads achetés', n, NB), ('dont 30 derniers jours', l30, NB), ('Taux lead → accord', _pct(n_acc, n), PCT),
         ('Taux lead → payé', _pct(n_pay, n), PCT), ('Cible', 0.06, PCT), ('Coût des leads', cost, EUR), ('CA réel', ca, EUR),
         ('ROI (CA / coût)', _pct(ca, cost), '0.00"x"'), ('MAF financé', maf, EUR), ('Coût par dossier payé', (cost / n_pay) if n_pay else None, EUR)]
for i, (lab, v, f) in enumerate(cards):
    c = 1 + (i % 5) * 3; rr = r + (i // 5) * 2
    a = K.cell(rr, c, lab); a.font = font(sz=8, color=GREY)
    b = K.cell(rr + 1, c, v); b.font = font(sz=14, b=True, color=NAVY); b.number_format = f
r += 5

# ---------- 2. Entonnoir par source
section(K, r, 1, '2 · Entonnoir par source (nombre de leads ayant atteint l’étape · % depuis le lead)', 14); r += 1
_hdr(r, 1, ['Source', 'Leads', 'Dans le CRM', 'Simulation', 'Dossier reçu', 'Dépôt banque', 'Accord', 'Payé (bordereau)',
            '% simu', '% dossier', '% dépôt', '% accord', '% payé', 'Dossier → dépôt']); r += 1
for src in ('Weendeal', 'Finanzen', 'Qwamplify', 'Kreid.fr', 'Total payants'):
    xs = pay_l if src == 'Total payants' else [x for x in _L if x['src'] == src]
    N = len(xs); cnt = [sum(1 for x in xs if x['incrm'])] + [sum(1 for x in xs if x[k]) for k in ('simu', 'dossier', 'depot', 'accord', 'paye')]
    _row(r, 1, [src, N] + cnt + [_pct(cnt[1], N), _pct(cnt[2], N), _pct(cnt[3], N), _pct(cnt[4], N), _pct(cnt[5], N), _pct(cnt[3], cnt[2])],
         [None, NB] + [NB] * 6 + [PCT] * 6, bold=(src == 'Total payants'), shade=(src == 'Total payants'))
    r += 1
K.cell(r, 1, "Lecture : « Dossier → dépôt » = part des dossiers reçus effectivement déposés en banque (fuite principale à surveiller).").font = font(i=True, color=GREY, sz=8)
r += 2

# ---------- 3. Délais
section(K, r, 1, '3 · Délais entre étapes (jours calendaires · tous leads suivis)', 14); r += 1
_hdr(r, 1, ['Étape', 'Nb de cas', 'Médiane', '75 % des cas sous', 'Lecture']); r += 1
pairs = [('Achat → réception CRM', 'date', 'recep', 'Arrivée du lead dans le CRM'), ('Réception → simulation', 'recep', 'simu', 'Réactivité du premier traitement'),
         ('Simulation → dossier reçu', 'simu', 'dossier', 'Capacité à obtenir les pièces'), ('Dossier reçu → dépôt', 'dossier', 'depot', 'KPI back office (cible ≤ 2 j)'),
         ('Dépôt → accord', 'depot', 'accord', 'Délai banque'), ('Accord → paiement Budgeo', 'accord', 'paye', 'Déblocage + commissionnement'),
         ('Réception → accord', 'recep', 'accord', 'Cycle commercial complet')]
for lab, a_, b_, txt in pairs:
    v = sorted(d for d in (_days(x[a_], x[b_]) for x in _L) if d is not None)
    _row(r, 1, [lab, len(v), _med(v), v[int(len(v) * .75)] if v else None, txt], [None, NB, J, J, None])
    K.cell(r, 5).alignment = Alignment(horizontal='left'); r += 1
r += 1

# ---------- 4. Conseillers
section(K, r, 1, '4 · Conseillers (fiches CRM des leads suivis · toutes sources)', 14); r += 1
_hdr(r, 1, ['Collaborateur', 'Leads', 'Simulation', 'Dossier reçu', 'Dépôt', 'Accord', 'Payé', 'Montant simulé', '% simu',
            '% simu → dossier', '% dossier → dépôt', '% lead → accord', 'Leads 30 j', 'Dernière étape']); r += 1
colls = {}
for x in _L: colls.setdefault(x['coll'], []).append(x)
for cname, xs in sorted(colls.items(), key=lambda kv: (kv[0].startswith('('), -len(kv[1]))):
    N = len(xs); c = [sum(1 for x in xs if x[k]) for k in ('simu', 'dossier', 'depot', 'accord', 'paye')]
    last = max([x[k] for x in xs for k in ('recep', 'simu', 'dossier', 'depot', 'accord') if isinstance(x[k], _dt.datetime)] or [None])
    _row(r, 1, [cname, N] + c + [sum(x['mts'] or 0 for x in xs) or None, _pct(c[0], N), _pct(c[1], c[0]), _pct(c[2], c[1]), _pct(c[3], N),
                 sum(1 for x in xs if x['recep'] and (TODAY - x['recep']).days <= 30), last],
         [None, NB, NB, NB, NB, NB, NB, EUR, PCT, PCT, PCT, PCT, NB, 'dd/mm/yyyy'])
    if cname.startswith('À réattribuer'): K.cell(r, 1).font = font(b=True, color='C00000')
    r += 1
K.cell(r, 1, "Montant simulé : disponible à partir de la prochaine lecture du CRM avec le favori mis à jour. « À réattribuer » = fiches de personnes sorties des effectifs.").font = font(i=True, color=GREY, sz=8)
r += 2

# ---------- 5. Cohortes
section(K, r, 1, '5 · Cohortes par mois d’achat (leads payants) · % d’accords obtenus à âge égal', 14); r += 1
_hdr(r, 1, ['Mois d’achat', 'Leads', 'Coût', '% accord à 30 j', '% accord à 60 j', '% accord à 90 j', '% accord à date', 'Payés', 'CA réel', 'ROI', 'MAF']); r += 1
months = sorted({x['date'].strftime('%Y-%m') for x in pay_l})
for m in months:
    xs = [x for x in pay_l if x['date'].strftime('%Y-%m') == m]
    age = (TODAY - min(x['date'] for x in xs)).days
    def at(nd):
        if age < nd + 28: return '—'   # cohorte trop jeune pour être comparée
        return _pct(sum(1 for x in xs if x['accord'] and _days(x['date'], x['accord']) is not None and _days(x['date'], x['accord']) <= nd), len(xs))
    co = sum(x['cost'] for x in xs); ca_ = sum(x['ca'] for x in xs)
    _row(r, 1, [m, len(xs), co, at(30), at(60), at(90), _pct(sum(1 for x in xs if x['accord']), len(xs)), sum(1 for x in xs if x['paye']), ca_,
                _pct(ca_, co), sum(x['maf'] for x in xs)], [None, NB, EUR, PCT, PCT, PCT, PCT, NB, EUR, '0.00"x"', EUR])
    r += 1
K.cell(r, 1, "« — » : cohorte trop récente pour mesurer ce délai. Comparer les mois sur une même colonne (à 30 j, 60 j, 90 j) et non sur « à date ».").font = font(i=True, color=GREY, sz=8)
r += 2

# ---------- 6. Back office
section(K, r, 1, '6 · Back office (base des KPI de pilotage du pôle)', 14); r += 1
_hdr(r, 1, ['Indicateur', 'Valeur', 'Cible proposée (à valider)', 'Lecture']); r += 1
dd = [_days(x['dossier'], x['depot']) for x in _L if x['dossier'] and x['depot']]
open_doss = [x for x in _L if x['dossier'] and not x['depot'] and not x['refus'] and (TODAY - x['dossier']).days > 7]
acc_np = [x for x in _L if x['accord'] and not x['paye'] and not x['refus']]
dep_wait = [x for x in _L if x['depot'] and not x['accord'] and not x['refus'] and (TODAY - x['depot']).days > 15]
bo = [('Dossiers déposés en ≤ 2 j après réception', _pct(sum(1 for d in dd if d is not None and d <= 2), len(dd)), 0.8, PCT, 'Réactivité du montage'),
      ('Dossiers reçus non déposés depuis > 7 j', len(open_doss), 0, NB, 'À déposer ou à classer avec motif'),
      ('Dépôts sans décision depuis > 15 j', len(dep_wait), 0, NB, 'Relancer la banque'),
      ('Accords non encore payés', len(acc_np), None, NB, 'Suivre déblocage et commissionnement'),
      ('Ancienneté médiane des accords non payés', _med([(TODAY - x['accord']).days for x in acc_np if isinstance(x['accord'], _dt.datetime)]), 30, J, 'Au-delà : risque de désistement'),
      ('Leads absents du CRM (hors Qwamplify)', sum(1 for x in _L if not x['incrm'] and x['src'] != 'Qwamplify'), 0, NB, 'Échec d’envoi API ou saisie manquante')]
for lab, v, tgt, f, txt in bo:
    _row(r, 1, [lab, v, tgt, txt], [None, f, f, None]); K.cell(r, 4).alignment = Alignment(horizontal='left'); r += 1

for c, w in enumerate([30, 9, 11, 11, 11, 11, 11, 13, 9, 11, 11, 11, 9, 13], 1):
    K.column_dimensions[L(c)].width = w
K.column_dimensions['E'].width = 14
K.freeze_panes = 'A4'
