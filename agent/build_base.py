# ---------- CRM (extraction lecture seule de l'extranet IOBSP, rapprochée par téléphone)
import datetime as _dt
def _d(s):
    try: return _dt.datetime.strptime(s, '%d/%m/%Y')
    except Exception: return None
_rows = [l.split(';') for l in open(CRM_TXT, encoding='utf-8').read().strip().split('\n')]
_best = {}
_cnt = {}
_COLL = {'DE': 'De Temmerman', 'MATON': 'Maton', 'BOUIN': 'Bouin', 'MORGAN': 'Morgan', 'LAZARE': 'Lazare', 'PETYT': 'Petyt', 'BATON': 'Baton'}
def _std(p):   # ancien format (9 champs) ou nouveau (11 champs avec nom / prénom)
    return p if len(p) == 11 else p[:3] + ['', ''] + p[3:9]
_rows = [_std(p) for p in _rows if len(p) in (9, 11)]
for ref, coll, tel, nom, prenom, recep, simu, retour, depot, acc, refus in _rows:
  for key in [t[-9:] for t in tel.split(',') if t]:
    _cnt[key] = _cnt.get(key, 0) + 1
    score = (5 if acc else 4 if depot else 3 if retour else 2 if simu else 1, _d(recep) or _dt.datetime(2000, 1, 1))
    if key not in _best or score > _best[key][0]:
        _best[key] = (score, [key, ref, _COLL.get(coll, coll.title()), nom.title(), prenom.title(), _d(recep), _d(simu), _d(retour), _d(depot), _d(acc),
                              (_d(refus) or (refus.strip() + '…' if refus.strip() else None))])
C = wb.create_sheet('CRM')
setup(C); C.sheet_properties.tabColor = HEAD
band(C, 1, 'CRM — extraction des fiches leads (extranet IOBSP, lecture seule)',
     "Extraction du " + TODAY.strftime('%d/%m/%Y') + " par l'agent · 1 ligne = 1 téléphone (fiche la plus avancée si doublon) · ne rien saisir ici", 11)
CH = ['Clé téléphone', 'Réf. CRM', 'Collaborateur', 'Nom', 'Prénom', 'Réception contact', 'Simulation envoyée', 'Retour dossier',
      'Dépôt banque', 'Dossier accepté', 'Refus (date ou motif)', 'Nb fiches CRM', 'Extrait le']
for c, t in enumerate(CH, 1):
    x = C.cell(3, c, t); x.font = font(b=True, color='FFFFFF'); x.fill = fill(HEAD); x.border = box
    x.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
C.row_dimensions[3].height = 30
EXTRACT = TODAY
for i, (key, (_, v)) in enumerate(sorted(_best.items(), key=lambda kv: kv[1][1][5] or _dt.datetime(2000, 1, 1), reverse=True)):
    r = 4 + i
    for c, val in enumerate(v + [_cnt[key], EXTRACT], 1):
        x = C.cell(r, c, val); x.font = font(); x.border = box
        if isinstance(val, _dt.datetime): x.number_format = 'dd/mm/yyyy'
        if c == 1: x.number_format = '@'
        x.alignment = Alignment(horizontal='center')
CRM_LAST = 4 + len(_best) + 400   # marge pour les prochaines extractions
for c, w in enumerate([13, 11, 15, 18, 15, 13, 13, 13, 13, 13, 18, 9, 11], 1):
    C.column_dimensions[L(c)].width = w
C.freeze_panes = 'B4'
C.auto_filter.ref = f'A3:M{3 + len(_best)}'
C.cell(3, 1).comment = Comment("Téléphone sur 9 chiffres (sans le 0) : clé de rapprochement avec les onglets fournisseurs.", 'Claude')


# ---------- BORDEREAUX BUDGEO (source unique du CA réel jusqu'au passage MNE)
BD = wb.create_sheet('Bordereaux Budgeo')
setup(BD); BD.sheet_properties.tabColor = GREEN
band(BD, 1, 'BORDEREAUX BUDGEO — CA réellement perçu par dossier (source du CA réel)',
     "Règle : le CA réel d'un lead = montant « Paiement » des bordereaux Budgeo (jusqu'au passage MNE) · 1 ligne = 1 ligne de bordereau", 9)
BH_ = ['N° bordereau', 'Date', 'Client', 'Banque', 'MAF (€)', 'CA perçu (€)', 'Clé téléphone lead', 'Fournisseur du lead', 'Rapprochement']
for c, t in enumerate(BH_, 1):
    x = BD.cell(3, c, t); x.font = font(b=True, color='FFFFFF'); x.fill = fill(HEAD); x.border = box
    x.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
def _num(x):
    x = x.replace('€', '').replace('\u00a0', ' ').replace(' ', '').replace(',', '.').strip()
    try: return float(x)
    except Exception: return None
def _norm(s): return ''.join(c for c in unicodedata.normalize('NFD', str(s or '').upper()) if unicodedata.category(c) != 'Mn').replace('-', ' ').replace("'", ' ').strip()
# 1) lecture des bordereaux (PDF -> texte)
_bl = []
for pdf in sorted(glob.glob(os.path.join(BX_DIR, 'Bordereau MY*.pdf'))):
    txt = subprocess.run(['pdftotext', '-layout', pdf, '-'], capture_output=True, text=True).stdout
    m = re.search(r'le (\d\d) (\d\d) (\d{4})', os.path.basename(pdf))
    bdate = _dt.datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    for block in re.split(r'BORDEREAU DE COMMISSIONNEMENT ', txt)[1:]:
        num = block.split()[0]
        if num.startswith('MYPROD'): continue
        for line in block.split('\n'):
            if not line.strip().endswith('€') or 'TOTAL' in line: continue
            parts = re.split(r'\s{2,}', line.strip())
            if len(parts) < 3 or not re.match(r"[A-Za-z]", parts[0]): continue
            pay = _num(parts[-1])
            if pay is None: continue
            sep = num.startswith('MYSEP')
            maf = None if sep else _num(parts[2]) if len(parts) > 3 else None
            _bl.append([num, bdate, parts[0].upper(), 'SEP' if sep else parts[1], maf, pay])
# dédoublonnage (un même bordereau peut être présent dans 2 fichiers)
_seen = set(); _tmp = []
for x in _bl:
    k = (x[0], x[2], x[5])
    if k not in _seen: _seen.add(k); _tmp.append(x)
_bl = _tmp
# 2) leads achetés (nom, date, téléphone) pour le rapprochement automatique
def _leads():
    out = []
    for name, ws, r0, cn, cd, ct in (('Weendeal', W, 3, 2, 1, 7), ('Finanzen', Fz, 2, 5, 1, 10), ('Qwamplify', Q, 2, 4, 1, 6), ('Kreid.fr', Kf, 2, 2, 1, 4)):
        for r in range(r0, ws.max_row + 1):
            d = ws.cell(r, cd).value
            if d in (None, ''): continue
            if isinstance(d, str):
                try: d = _dt.datetime(int(d[:4]), int(d[5:7]), int(d[8:10]))
                except Exception: continue
            tel = re.sub(r'\D', '', str(ws.cell(r, ct).value or ''))[-9:]
            out.append((name, _norm(ws.cell(r, cn).value), d, tel))
    return out
_LEADS = _leads()
# leads du site : l'alerte « Absent du CRM » ne s'active que si l'extraction CRM a bien cherché leurs téléphones
KREID_CRM = any(src == 'Kreid.fr' and tel in _best for src, _, _, tel in _LEADS)
_acc = {k: v[1][9] for k, v in _best.items()}   # date d'accord CRM par clé téléphone
def _match(name, bdate):
    nm = _norm(name); res = []
    for src, ln, d, tel in _LEADS:
        if not ln or not (nm == ln or nm.startswith(ln + ' ') or ln.startswith(nm + ' ')): continue
        a = _acc.get(tel)
        if a and d <= a <= bdate and (bdate - a).days <= 120 and tel not in [x[0] for x in res]: res.append((tel, src))
    return res[0] if len(res) == 1 else (None, None)
# rapprochements manuels (numéro différent entre fournisseur et CRM, homonymes…)
_MAN = {}
_MAN_CSV = os.path.join(os.path.dirname(ADJ_CSV), 'rapprochements_manuels.csv')
if os.path.exists(_MAN_CSV):
    for l in open(_MAN_CSV, encoding='utf-8').read().strip().split('\n')[1:]:
        p = l.split(';')
        if len(p) >= 4: _MAN[(p[0].strip(), _norm(p[1]))] = (p[2].strip(), p[3].strip())
# 3) ajustements manuels (gestes commerciaux, reprises…)
_adj = []
if os.path.exists(ADJ_CSV):
    for i, l in enumerate(open(ADJ_CSV, encoding='utf-8').read().strip().split('\n')[1:]):
        p = l.split(';')
        if len(p) >= 5 and p[4].strip():
            _adj.append(['AJUSTEMENT', _dt.datetime.fromisoformat(p[0]), p[1].upper(), 'Ajustement', None, float(p[4]), p[3].strip(), p[2].strip(), p[5].strip() if len(p) > 5 else ''])
_rows_bd = []
for num, bdate, name, bank, maf, pay in _bl:
    key, src = _MAN.get((num, _norm(name))) or ((None, None) if num.startswith('MYSEP') else _match(name, bdate))
    _rows_bd.append([num, bdate, name, bank, maf, pay, key, src, 'Lead rattaché' if key else 'Hors leads suivis'])
_rows_bd += [a[:8] + ['Ajustement : ' + a[8]] for a in _adj]
for i, vals in enumerate(_rows_bd):
    r = 4 + i
    key = vals[6]
    for c, v in enumerate(vals, 1):
        x = BD.cell(r, c, v); x.font = font(b=bool(key) and c in (3, 6), color=('C00000' if vals[0] == 'AJUSTEMENT' else GREEN if key and c == 9 else '000000')); x.border = box
        x.number_format = {2: 'dd/mm/yyyy', 5: '#,##0 €', 6: '#,##0.00 €', 7: '@'}.get(c, 'General')
        x.alignment = Alignment(horizontal='center' if c in (1, 2, 4, 7, 8) else None)
_bl = _rows_bd
BD_LAST = 2000
for c, w in enumerate([15, 11, 22, 12, 12, 13, 15, 14, 40], 1):
    BD.column_dimensions[L(c)].width = w
BD.freeze_panes = 'A4'
BD.auto_filter.ref = f'A3:I{3 + len(_bl)}'
BD.cell(3, 7).comment = Comment("Renseigné uniquement quand la ligne correspond à un lead acheté (rapprochement nom + téléphone CRM). Sans clé, la ligne ne compte pas dans la rentabilité des leads.", 'Claude')
note_r = 5 + len(_bl)
BD.cell(note_r, 1, "Rapprochement automatique : même nom + accord CRM du lead (téléphone) entre la date d'achat et le bordereau (≤ 120 j). Lignes « AJUSTEMENT » = _agent/ajustements.csv.").font = font(sz=8, i=True, color=GREY)

# ---------- BASE LEADS (consolidation propre, alimentée automatiquement)
B = wb.create_sheet('Base leads')
setup(B); B.sheet_properties.tabColor = GREY
band(B, 1, 'BASE LEADS — consolidation automatique des 4 sources + avancement CRM (source des TCD)',
     "Ne rien saisir ici : alimenté par Weendeal, Finanzen, Qwamplify, Kreid.fr et l'onglet CRM. 1 ligne = 1 lead. Filtre « Alerte suivi » = liste de relances.", 26)
HDR = ['Fournisseur', "Date d'achat", 'Année', 'Trimestre', 'Mois', 'Nom', 'Prénom', 'Prix lead (€)', 'Accord (1/0)',
       'CA si accord (€)', 'Financé (1/0)', 'CA réel encaissé (€)', 'Résultat lead (€)', 'Lead',
       'Dans CRM (1/0)', 'Simulation (1/0)', 'Dossier reçu (1/0)', 'Dépôt (1/0)', 'Accord CRM (1/0)', 'Refus CRM (1/0)',
       'Étape CRM', 'Alerte suivi', 'Collaborateur', 'CA en attente (€)', 'Téléphone (clé)', 'Ligne CRM', 'Ligne source', 'N° relance', 'MAF financé (€)']
HR = 3
for c, t in enumerate(HDR, 1):
    x = B.cell(HR, c, t); x.font = font(b=True, color='FFFFFF'); x.fill = fill(HEAD if c <= 24 or c == 29 else GREY); x.border = box
    x.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
B.row_dimensions[HR].height = 30
# compteurs (cellules de contrôle) en AD/AE
B['AD3'] = 'Contrôle'; B['AD3'].font = font(b=True, color=NAVY)
ctl = [('Weendeal', '=COUNTA(Weendeal!$A$3:$A$1000)'), ('Finanzen', '=COUNTA(Finanzen!$A$2:$A$1000)'),
       ('Qwamplify', '=COUNTA(Qwamplify!$A$2:$A$1000)'), ('Total', '=SUM(AE4:AE6)+AE10'),
       ('Lignes base', '=COUNTIF($A$4:$A$1203,"?*")'), ('Relances', '=COUNTIF($V$4:$V$1203,"?*")')]
for i, (k, f) in enumerate(ctl):
    B.cell(4 + i, 30, k).font = font(b=(k in ('Total', 'Lignes base')))
    x = B.cell(4 + i, 31, f); x.font = font(b=(k in ('Total', 'Lignes base'))); x.border = box
B['AD10'] = 'Kreid.fr'; B['AE10'] = "=COUNTA('Kreid.fr'!$A$2:$A$1000)"; B['AE10'].border = box
B['AD12'] = '« Total » doit être égal à « Lignes base ».'; B['AD12'].font = font(sz=8, i=True, color=GREY)
nW, nF, nQ = '$AE$4', '$AE$5', '$AE$6'
SRC = {  # colonne dans chaque onglet source : (Weendeal, Finanzen, Qwamplify, Kreid.fr)
    'date': ('A', 'A', 'A', 'A'), 'nom': ('B', 'E', 'D', 'B'), 'prenom': ('C', 'D', 'E', 'C'), 'prix': ('K', 'M', 'U', 'F'),
    'acc': ('L', 'N', 'V', 'G'), 'ca': ('M', 'O', 'W', 'H'), 'careel': ('O', 'Q', 'Y', 'J'), 'tel': ('G', 'J', 'F', 'D')}
def pick(key, r, wrap='{}'):
    w, f, q, kk = SRC[key]
    a = wrap.format(f'INDEX(Weendeal!${w}$3:${w}$1000,$AA{r})')
    b = wrap.format(f'INDEX(Finanzen!${f}$2:${f}$1000,$AA{r})')
    c = wrap.format(f'INDEX(Qwamplify!${q}$2:${q}$1000,$AA{r})')
    d = wrap.format(f"INDEX('Kreid.fr'!${kk}$2:${kk}$1000,$AA{r})")
    return f'IF($A{r}="Weendeal",{a},IF($A{r}="Finanzen",{b},IF($A{r}="Qwamplify",{c},{d})))'
FIRST, LAST = HR + 1, HR + 1200
EUR = '#,##0.00 €;-#,##0.00 €;-'
CR = lambda col, r: f'INDEX(CRM!${col}$1:${col}${CRM_LAST},$Z{r})'
for r in range(FIRST, LAST + 1):
    k = f'(ROW()-{HR})'
    wd = f'INDEX(Weendeal!$A$3:$A$1000,$AA{r})'
    has = lambda col: f'IF($Z{r}=0,0,IF({CR(col, r)}<>"",1,0))'
    KX = '' if KREID_CRM else f'$A{r}<>"Kreid.fr",'
    fx = {
        1: f'=IF({k}<={nW},"Weendeal",IF({k}<={nW}+{nF},"Finanzen",IF({k}<={nW}+{nF}+{nQ},"Qwamplify",IF({k}<={nW}+{nF}+{nQ}+$AE$10,"Kreid.fr",""))))',
        27: f'=IF($A{r}="","",IF($A{r}="Weendeal",{k},IF($A{r}="Finanzen",{k}-{nW},IF($A{r}="Qwamplify",{k}-{nW}-{nF},{k}-{nW}-{nF}-{nQ}))))',
        2: (f'=IF($A{r}="","",IF($A{r}="Weendeal",IF(ISNUMBER({wd}),INT({wd}),DATE(VALUE(LEFT({wd},4)),VALUE(MID({wd},6,2)),VALUE(MID({wd},9,2)))),'
            f'INT(IF($A{r}="Finanzen",INDEX(Finanzen!$A$2:$A$1000,$AA{r}),IF($A{r}="Qwamplify",INDEX(Qwamplify!$A$2:$A$1000,$AA{r}),INDEX(\'Kreid.fr\'!$A$2:$A$1000,$AA{r}))))))'),
        3: f'=IF($A{r}="","",YEAR(B{r}))',
        4: f'=IF($A{r}="","",YEAR(B{r})&" T"&ROUNDUP(MONTH(B{r})/3,0))',
        5: f'=IF($A{r}="","",YEAR(B{r})&"-"&TEXT(MONTH(B{r}),"00"))',
        6: f'=IF($A{r}="","",PROPER({pick("nom", r)}))',
        7: f'=IF($A{r}="","",PROPER({pick("prenom", r)}))',
        8: f'=IF($A{r}="","",{pick("prix", r, "N({})")})',
        9: f'=IF($A{r}="","",IF(LOWER(TRIM({pick("acc", r)}))="oui",1,0))',
        10: f'=IF($A{r}="","",{pick("ca", r, "N({})")})',
        12: f'=IF($A{r}="","",IF($Y{r}="",0,SUMIFS(\'Bordereaux Budgeo\'!$F$4:$F${BD_LAST},\'Bordereaux Budgeo\'!$G$4:$G${BD_LAST},$Y{r},\'Bordereaux Budgeo\'!$H$4:$H${BD_LAST},$A{r})))',
        29: f'=IF($A{r}="","",IF($Y{r}="",0,_xlfn.MAXIFS(\'Bordereaux Budgeo\'!$E$4:$E${BD_LAST},\'Bordereaux Budgeo\'!$G$4:$G${BD_LAST},$Y{r},\'Bordereaux Budgeo\'!$H$4:$H${BD_LAST},$A{r})))',
        11: f'=IF($A{r}="","",IF(L{r}>0,1,0))',
        13: f'=IF($A{r}="","",L{r}-H{r})',
        14: f'=IF($A{r}="","",1)',
        25: f'=IF($A{r}="","",RIGHT(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({pick("tel", r)}&""," ",""),".",""),"-",""),9))',
        26: f'=IF($A{r}="","",IFERROR(MATCH($Y{r},CRM!$A$1:$A${CRM_LAST},0),0))',
        15: f'=IF($A{r}="","",IF($Z{r}>0,1,0))',
        16: f'=IF($A{r}="","",{has("G")})',
        17: f'=IF($A{r}="","",{has("H")})',
        18: f'=IF($A{r}="","",{has("I")})',
        19: f'=IF($A{r}="","",{has("J")})',
        20: f'=IF($A{r}="","",IF(S{r}=1,0,{has("K")}))',
        21: (f'=IF($A{r}="","",IF(O{r}=0,"0 · Absent du CRM",IF(S{r}=1,"5 · Accord banque",IF(T{r}=1,"× Refus / sans suite",'
             f'IF(R{r}=1,"4 · Dépôt banque",IF(Q{r}=1,"3 · Dossier reçu",IF(P{r}=1,"2 · Simulation envoyée","1 · Contact reçu")))))))'),
        22: (f'=IF($A{r}="","",IF(TODAY()-B{r}>120,"",IF(O{r}=0,IF(AND($A{r}<>"Qwamplify",{KX}L{r}=0,TODAY()-B{r}>1),"Absent du CRM",""),'
             f'IF(S{r}=1,IF(L{r}=0,"Accord : suivre déblocage",""),IF(T{r}=1,"",'
             f'IF(R{r}=1,IF(TODAY()-{CR("I", r)}>15,"Dépôt > 15 j : relancer banque",""),'
             f'IF(Q{r}=1,IF(TODAY()-{CR("H", r)}>7,"Dossier reçu > 7 j : déposer",""),'
             f'IF(P{r}=1,IF(TODAY()-{CR("G", r)}>10,"Simulation > 10 j : relancer pièces",""),'
             f'IF(TODAY()-B{r}>2,"Contact non traité",""))))))))'),
        23: f'=IF(OR($A{r}="",$Z{r}=0),"",{CR("C", r)})',
        24: f'=IF($A{r}="","",IF(AND(L{r}=0,OR(I{r}=1,S{r}=1)),J{r},0))',
        28: f'=IF(V{r}="","",COUNTIF($V${FIRST}:V{r},"?*"))',
    }
    for c in range(1, 30):
        x = B.cell(r, c, fx[c]); x.font = font(color=('808080' if 25 <= c <= 28 else '000000'))
        x.number_format = {2: 'dd/mm/yyyy', 8: EUR, 10: EUR, 12: EUR, 13: EUR, 24: EUR, 29: '#,##0 €;-#,##0 €;-'}.get(c, '0' if c in (3, 9, 11, 14, 15, 16, 17, 18, 19, 20, 26, 27, 28) else 'General')
        if c == 25: x.number_format = '@'
        x.alignment = Alignment(horizontal='center' if c in (1, 3, 4, 5, 9, 11, 14, 15, 16, 17, 18, 19, 20, 25, 26, 27, 28) else None)
for c, w in enumerate([11, 11, 7, 9, 9, 16, 14, 11, 8, 12, 8, 13, 13, 6, 8, 9, 9, 8, 9, 8, 20, 28, 14, 12, 11, 7, 7, 8, 3, 12, 8], 1):
    B.column_dimensions[L(c)].width = w
B.freeze_panes = 'B4'
B.auto_filter.ref = f'A{HR}:AC{LAST}'
B.column_dimensions['AC'].width = 12
B.conditional_formatting.add(f'A{FIRST}:X{LAST}', openpyxl.formatting.rule.FormulaRule(
    formula=[f'AND($A{FIRST}<>"",MOD(ROW(),2)=0)'], fill=fill('F3F6FA')))
B.conditional_formatting.add(f'M{FIRST}:M{LAST}', openpyxl.formatting.rule.CellIsRule(operator='greaterThan', formula=['0'], font=Font(name=F, b=True, color=GREEN)))
B.conditional_formatting.add(f'V{FIRST}:V{LAST}', openpyxl.formatting.rule.FormulaRule(formula=[f'$V{FIRST}<>""'], font=Font(name=F, b=True, color='C00000')))
B.cell(HR, 27).comment = Comment('Numéro de ligne dans l\'onglet fournisseur (technique, pour retrouver le lead).', 'Claude')
B.cell(HR, 11).comment = Comment('1 = CA réel encaissé > 0 (même règle que l\'ancienne synthèse « Ventes »).', 'Claude')
B.cell(HR, 13).comment = Comment('CA réel encaissé − prix du lead. Hors salaire et commission (volontairement, à ce stade).', 'Claude')
B.cell(HR, 22).comment = Comment("Règles de relance (leads de moins de 120 jours) : absent du CRM ; contact sans simulation > 2 j ; simulation sans retour de dossier > 10 j ; dossier reçu non déposé > 7 j ; dépôt sans décision > 15 j ; accord sans CA encaissé. Seuils modifiables avec Claude.", 'Claude')
B.cell(HR, 12).comment = Comment("CA réel = somme des montants « Paiement » des bordereaux Budgeo rattachés à ce lead (onglet Bordereaux Budgeo). Les montants saisis à la main dans les onglets fournisseurs ne sont plus utilisés.", 'Claude')
B.cell(HR, 29).comment = Comment("MAF (montant à financer) du bordereau Budgeo rattaché au lead : base de l'objectif annuel en MAF.", 'Claude')
B.cell(HR, 24).comment = Comment("CA saisi à l'accord (onglet fournisseur) pour les dossiers accordés pas encore présents dans un bordereau Budgeo : CA attendu, non encaissé.", 'Claude')
B.cell(HR, 9).comment = Comment("Accord saisi à la main dans l'onglet fournisseur (historique). La référence est désormais « Accord CRM ».", 'Claude')

# ---------- RELANCES (liste dynamique pour Corentin, sans origine des contacts)
RL = wb.create_sheet('Relances')
setup(RL); RL.sheet_properties.tabColor = 'C00000'
band(RL, 1, 'RELANCES DU JOUR — clients ayant formulé une demande',
     "Liste recalculée à chaque ouverture à partir du CRM · traiter les relances dans le CRM, rien à saisir ici", 9)
RH = ['N°', 'Relance à faire', 'Nom', 'Prénom', 'Téléphone', 'Étape CRM', 'Collaborateur', 'Reçu le', 'Ancienneté (j)']
for c, t in enumerate(RH, 1):
    x = RL.cell(3, c, t); x.font = font(b=True, color='FFFFFF'); x.fill = fill(HEAD); x.border = box
    x.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
RL['K3'] = 'Total relances'; RL['K3'].font = font(b=True, color=NAVY)
RL['L3'] = "='Base leads'!AE9"; RL['L3'].font = font(b=True, color='C00000')
BL_ = lambda col: f"'Base leads'!${col}${FIRST}:${col}${LAST}"
for i in range(1, 301):
    r = 3 + i
    m = f'MATCH({i},{BL_("AB")},0)'
    vals = [f'=IF({i}>$L$3,"",{i})',
            f'=IF($A{r}="","",INDEX({BL_("V")},{m}))', f'=IF($A{r}="","",INDEX({BL_("F")},{m}))',
            f'=IF($A{r}="","",INDEX({BL_("G")},{m}))', f'=IF($A{r}="","","0"&INDEX({BL_("Y")},{m}))',
            f'=IF($A{r}="","",INDEX({BL_("U")},{m}))', f'=IF($A{r}="","",INDEX({BL_("W")},{m}))',
            f'=IF($A{r}="","",INDEX({BL_("B")},{m}))', f'=IF($A{r}="","",TODAY()-H{r})']
    for c, v in enumerate(vals, 1):
        x = RL.cell(r, c, v); x.font = font(); x.alignment = Alignment(horizontal='center' if c != 2 else None)
        if c == 8: x.number_format = 'dd/mm/yyyy'
        if c == 9: x.number_format = '0'
RL.conditional_formatting.add('A4:I303', openpyxl.formatting.rule.FormulaRule(formula=['$A4<>""'], border=box))
for c, w in enumerate([5, 32, 16, 14, 13, 22, 15, 11, 10, 3, 13, 8], 1):
    RL.column_dimensions[L(c)].width = w
RL.freeze_panes = 'A4'
