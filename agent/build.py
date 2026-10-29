import openpyxl, copy
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L

import sys, os, re, csv, glob, subprocess, unicodedata, datetime as _dt
SRC, CRM_TXT, BX_DIR, ADJ_CSV, OUT = sys.argv[1:6]
TODAY = _dt.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
wb = openpyxl.load_workbook(SRC)
for s in ('Rentabilité Qwamplify', 'Optimisation contacts', 'Leads', 'Relances', 'CRM', 'Bordereaux Budgeo', 'Base leads', 'Méthode & sources', 'Journal imports'):
    if s in wb.sheetnames: del wb[s]

NAVY, HEAD, GREEN, LIGHT, INPUT, GREY = '17365D', '1F4E78', '2E8B57', 'F3F6FA', 'FFF2CC', '666666'
F = 'Arial'
font = lambda **k: Font(name=F, size=k.pop('sz', 9), **k)
fill = lambda c: PatternFill('solid', fgColor=c)
thin = Side(style='thin', color='D9DEE7'); box = Border(left=thin, right=thin, top=thin, bottom=thin)
FOOT = ("Kreid est une marque commerciale de My Finances, SAS au capital de 1 000 € – RCS Lille Métropole 911 118 255 – "
        "ORIAS n° 22002570 – Document interne de pilotage – données personnelles : usage strictement interne (RGPD)")

def setup(ws):
    ws.sheet_view.showGridLines = False
    ws.oddFooter.center.text = FOOT; ws.oddFooter.center.size = 7
    ws.oddHeader.right.text = 'KREID · Pôle Distance'; ws.oddHeader.right.size = 8
    ws.page_setup.orientation = 'landscape'; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

def band(ws, r, text, sub, width):
    for c in range(1, width + 1):
        ws.cell(r, c).fill = fill(NAVY); ws.cell(r + 1, c).fill = fill(GREEN)
    ws.cell(r, 1, text).font = font(sz=15, b=True, color='FFFFFF'); ws.row_dimensions[r].height = 26
    ws.cell(r + 1, 1, sub).font = font(i=True, color='FFFFFF')

def section(ws, r, c0, text, width):
    for c in range(c0, c0 + width):
        ws.cell(r, c).border = Border(bottom=Side(style='medium', color=GREEN))
    ws.cell(r, c0, text).font = font(sz=11, b=True, color=NAVY); ws.row_dimensions[r].height = 20

# ---------- Sources : on ne modifie jamais les onglets de saisie (Weendeal, Finanzen, Qwamplify)
W, Fz, Q = wb['Weendeal'], wb['Finanzen'], wb['Qwamplify']
if 'Kreid.fr' not in wb.sheetnames:   # 1re création : leads du site connus au 28/10/2026, ensuite onglet de saisie
    _seed = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'kreid_seed.csv')   # données personnelles : hors Git
    SEED = [l.split(';') for l in open(_seed, encoding='utf-8').read().strip().split('\n')[1:]] if os.path.exists(_seed) else []
    K = wb.create_sheet('Kreid.fr', index=wb.sheetnames.index('Qwamplify') + 1)
    KH = ['Date de la demande', 'Nom', 'Prénom', 'Téléphone', 'Canal', 'Prix (€)', 'Accord', 'CA', 'Transfo', 'CA réel', 'Commentaire']
    for c, t in enumerate(KH, 1):
        x = K.cell(1, c, t); x.font = font(b=True, color='FFFFFF'); x.fill = fill(HEAD); x.border = box
        x.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for i, (d, n, p, t, canal) in enumerate(SEED):
        r = 2 + i
        vals = [_dt.datetime.fromisoformat(d), n, p, t, canal, 0, 'non', None, 'non', None, None]
        for c, v in enumerate(vals, 1):
            x = K.cell(r, c, v); x.font = font(); x.border = box
            if c == 1: x.number_format = 'dd/mm/yyyy'
            if c == 4: x.number_format = '@'
            if c == 6: x.number_format = '#,##0.00 €'
    for c, w in enumerate([12, 18, 14, 13, 28, 9, 8, 11, 8, 11, 30], 1):
        K.column_dimensions[L(c)].width = w
    K.freeze_panes = 'B2'
    K.cell(1, 1).comment = Comment("Leads du site kreid.fr (formulaire « Demande d'informations depuis le site internet Kreid »). Coût d'achat = 0. L'agent ajoute chaque nouvelle demande reçue par mail ; CA réel = bordereaux Budgeo (automatique).", 'Claude')
Kf = wb['Kreid.fr']

exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'build_base.py'), encoding='utf-8').read())

# ---------- LEADS (TCD)
S = wb.create_sheet('Leads', index=3)
setup(S); S.sheet_properties.tabColor = NAVY
band(S, 1, 'RENTABILITÉ DES LEADS — par fournisseur et par période',
     "TCD actualisés automatiquement à l'ouverture · après une saisie : Données › Actualiser tout · CA rattaché au mois d'achat du lead", 24)
S['A3'] = ("Lecture : Résultat = CA réel encaissé − coût d'achat des leads (hors salaire et commission conseiller). "
           "Filtres Année / Trimestre en tête de chaque TCD. Données : onglet « Base leads ».")
S['A3'].font = font(sz=8, i=True, color=GREY)
section(S, 5, 1, '1. RENTABILITÉ PAR FOURNISSEUR', 11)
section(S, 5, 13, 'Indicateurs complémentaires (toutes périodes)', 7)
section(S, 17, 1, '2. ENTONNOIR CRM PAR FOURNISSEUR (avancement réel des dossiers)', 11)
section(S, 31, 1, '3. RENTABILITÉ PAR PÉRIODE ET PAR FOURNISSEUR', 8)
section(S, 31, 10, '4. RÉSULTAT PAR MOIS × FOURNISSEUR', 6)
# bloc ROI (formules sur la base, non filtrées)
hdr = ['Fournisseur', 'ROI (CA réel / coût)', 'Coût lead par financé', 'CA réel par lead', 'Dernier lead', 'MAF financé', 'MAF par € de lead']
for i, t in enumerate(hdr):
    x = S.cell(7, 13 + i, t); x.font = font(b=True, color='FFFFFF'); x.fill = fill(HEAD); x.border = box
    x.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
S.row_dimensions[7].height = 30
BA, BH, BK, BL, BB = "'Base leads'!$A$4:$A$1203", "'Base leads'!$H$4:$H$1203", "'Base leads'!$K$4:$K$1203", "'Base leads'!$L$4:$L$1203", "'Base leads'!$B$4:$B$1203"
BAC = "'Base leads'!$AC$4:$AC$1203"
for i, src in enumerate(['Weendeal', 'Finanzen', 'Qwamplify', 'Kreid.fr', 'Total payants']):
    r = 8 + i; tot = src.startswith('Total')
    crit = f'"?*",{BA},"<>Kreid.fr"' if tot else f'M{r}'
    vals = [src,
            f'=IFERROR(SUMIFS({BL},{BA},{crit})/SUMIFS({BH},{BA},{crit}),"coût 0")',
            f'=IFERROR(SUMIFS({BH},{BA},{crit})/SUMIFS({BK},{BA},{crit}),0)',
            f'=IFERROR(SUMIFS({BL},{BA},{crit})/COUNTIFS({BA},{crit}),0)',
            f'=_xlfn.MAXIFS({BB},{BA},{crit})',
            f'=SUMIFS({BAC},{BA},{crit})',
            f'=IFERROR(SUMIFS({BAC},{BA},{crit})/SUMIFS({BH},{BA},{crit}),"coût 0")']
    for c, v in enumerate(vals):
        x = S.cell(r, 13 + c, v); x.border = box; x.font = font(b=tot)
        x.number_format = ['General', '0.00"x"', '#,##0 €', '#,##0.00 €', 'dd/mm/yyyy', '#,##0 €', '#,##0 €'][c]
        if tot: x.fill = fill(LIGHT)
S['M14'] = "ROI < 1 : le fournisseur ne rembourse même pas le prix de ses leads. Kreid.fr : coût d'achat nul (référence), exclu du total payants."
S['M14'].font = font(sz=8, i=True, color=GREY)
S.conditional_formatting.add('N8:N10', openpyxl.formatting.rule.CellIsRule(operator='lessThan', formula=['1'], font=Font(name=F, b=True, color='C00000')))
S.conditional_formatting.add('N8:N10', openpyxl.formatting.rule.CellIsRule(operator='greaterThanOrEqual', formula=['3'], font=Font(name=F, b=True, color=GREEN)))
widths = {'A': 14, 'B': 14, 'C': 15, 'D': 17, 'E': 11, 'F': 13, 'G': 15, 'H': 17, 'I': 15, 'J': 14, 'K': 14, 'L': 14,
          'M': 14, 'N': 14, 'O': 14, 'P': 14, 'Q': 13, 'R': 14, 'S': 14}
for k, v in widths.items(): S.column_dimensions[k].width = v

# ---------- MÉTHODE & SOURCES
M = wb.create_sheet('Méthode & sources')
setup(M); M.page_setup.orientation = 'portrait'
band(M, 1, 'MÉTHODE, SOURCES ET DÉCISIONS', 'Documentation — fichier régénéré automatiquement le ' + TODAY.strftime('%d/%m/%Y'), 6)
M.column_dimensions['A'].width = 30
for _c in 'BCDEF': M.column_dimensions[_c].width = 22
lines = [
    ('Structure du fichier', ''),
    ('Weendeal / Finanzen / Qwamplify', "Onglets de saisie, 1 ligne par lead. Ajouter les nouveaux leads à la suite, sans ligne vide ni ligne de total (les anciennes lignes de total ont été supprimées)."),
    ('Base leads', "Consolidation automatique des 3 onglets (formules), triée par fournisseur. Ne rien y saisir. Contrôle des volumes en AE4:AE9. Colonne « Alerte suivi » = règles de relance. Capacité 1 200 leads (étendre les formules au-delà)."),
    ('Leads', "3 TCD sur la Base leads, actualisés à l'ouverture : (1) par fournisseur, (2) par mois × fournisseur, (3) matrice résultat mois × fournisseur. Bloc d'indicateurs complémentaires (ROI, coût par financé) en formules."),
    ('CRM', "Extraction en lecture seule de la liste clients de l'extranet IOBSP (dates : réception, simulation, retour dossier, dépôt, accord, refus), rapprochée des leads par le téléphone. Mise à jour par l'agent Claude (aucune saisie). 1re extraction : 28/10/2026."),
    ('Relances', "Liste du jour calculée depuis la Base leads (sans mention de l'origine des contacts) : à traiter dans le CRM."),
    ('Définitions', ''),
    ('Coût leads', 'Somme des prix d\'achat (Prix TTC).'),
    ('Accord', 'Colonne « Accord » = oui dans l\'onglet fournisseur.'),
    ('Financé', 'CA réel encaissé > 0.'),
    ('Taux de transfo', 'Financés / leads achetés (moyenne de la colonne Financé 1/0).'),
    ('Résultat', 'CA réel encaissé − coût des leads. Volontairement hors salaire et commission à ce stade.'),
    ('Période', "Tout est rattaché au mois d'achat du lead (logique de cohorte), pas au mois d'encaissement."),
    ('Sources des chiffres', ''),
    ('Qwamplify', "198 leads repris de SUIVI-DISTANCE.xlsx (Suivi Leads), achetés du 31/03 au 20/05/2026."),
    ('Mise à jour automatique', "Chaque jour ouvré, l'agent Claude lit le CRM (lecture seule), intègre les nouveaux bordereaux Budgeo du dossier Commission/Perçues/2026/Budgéo et régénère ce fichier. Scripts et procédure : dossier _agent (README.md). Ajustements manuels du CA (gestes commerciaux…) : _agent/ajustements.csv."),
    ('Kreid.fr', "Leads du site (coût 0) : demandes reçues par le formulaire « Demande d'informations depuis le site internet Kreid ». Base initiale reconstituée le 28/10/2026 depuis la messagerie (36 demandes depuis le 31/03) + 2 leads saisis « KREID » dans SUIVI-DISTANCE. L'agent ajoute les nouvelles demandes à chaque passage."),
    ('MAF', "MAF financé = montant « MAF » du bordereau Budgeo rattaché au lead (une seule fois par dossier)."),
    ('CA réel', "Règle (Yohan, 28/10/2026) : le CA réel = montant « Paiement » des bordereaux Budgeo, jusqu'au passage MNE. Onglet Bordereaux Budgeo, rapproché des leads par nom + téléphone CRM. Les CA saisis à la main dans les onglets fournisseurs ne servent plus que de « CA en attente » pour les accords pas encore payés."),
    ('Prix Qwamplify', "20 € HT (info Yohan, 28/10/2026), saisi 24 € TTC pour rester homogène avec Finanzen (Prix TTC). Prix Weendeal (35 €) : base HT/TTC à confirmer."),
    ('Ajouter un fournisseur', "Créer son onglet, ajouter son compteur dans Base leads (R4:R6) et sa branche dans les formules de la base (à faire avec Claude)."),
    ('Reporté (2e temps)', "Rentabilité détaillée Qwamplify (coût complet par dossier) et optimisation des contacts : retirés de cette version, à reprendre."),
]
r = 4
for k, v in lines:
    if not v:
        section(M, r, 1, k, 6)
    else:
        a = M.cell(r, 1, k); a.font = font(b=True); a.fill = fill(LIGHT); a.border = box; a.alignment = Alignment(wrap_text=True, vertical='top')
        M.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        b = M.cell(r, 2, v); b.font = font(color='C00000' if 'NON TROUVÉ' in v else '000000'); b.alignment = Alignment(wrap_text=True, vertical='top')
        M.row_dimensions[r].height = max(18, 14 * (len(v) // 95 + 1))
    r += 1

# ---------- JOURNAL DES IMPORTS (traçabilité : _agent/journal_imports.csv, le plus récent en haut)
J = wb.create_sheet('Journal imports')
setup(J); J.sheet_properties.tabColor = GREY
band(J, 1, 'JOURNAL DES IMPORTS — ce que chaque passage a lu, ajouté et vu changer',
     "Exports Weendeal / Finanzen, extractions CRM : fichiers lus (empreinte), leads ajoutés, statuts portail, étapes CRM franchies · 2 000 derniers événements", 8)
JH = ['Horodatage', 'Source', 'Fichier', 'Événement', 'Clé', 'Nom', 'Prénom', 'Détail']
for c, t in enumerate(JH, 1):
    x = J.cell(3, c, t); x.font = font(b=True, color='FFFFFF'); x.fill = fill(HEAD); x.border = box
_jp = os.environ.get('JOURNAL_FOR_BUILD') or os.path.join(os.path.dirname(ADJ_CSV), 'journal_imports.csv')
_jr = [l for l in csv.reader(open(_jp, encoding='utf-8'), delimiter=';')][1:] if os.path.exists(_jp) else []
_COLORS = {'AJOUT_ONGLET': GREEN, 'SYNTHESE': NAVY, 'TELEPHONE_MANQUANT': 'C00000', 'FICHE_DISPARUE': 'C00000', 'DISPARU_EXPORT': 'C55A11',
           'FICHIER_INCHANGE': 'C55A11', 'AUCUN_FICHIER': 'C00000', 'ETAPE_CRM': '1F4E78', 'NOUVELLE_FICHE_LEAD': GREEN}
for i, v in enumerate(reversed(_jr[-2000:])):
    for c, val in enumerate(v[:8], 1):
        x = J.cell(4 + i, c, val); x.border = box
        x.font = font(b=(v[3] == 'SYNTHESE'), color=_COLORS.get(v[3], '000000') if c in (4, 8) else '000000')
for c, w in enumerate([16, 10, 30, 22, 12, 18, 14, 80], 1):
    J.column_dimensions[L(c)].width = w
J.freeze_panes = 'A4'
J.auto_filter.ref = f'A3:H{3 + max(1, min(len(_jr), 2000))}'

for _ws in wb.worksheets:
    _ws.oddFooter.center.text = FOOT; _ws.oddFooter.center.size = 7; _ws.oddFooter.center.font = None
    _ws.oddHeader.right.text = 'KREID · Pôle Distance'; _ws.oddHeader.right.size = 8; _ws.oddHeader.right.font = None
    _ws.oddFooter.center.color = None; _ws.oddHeader.right.color = None
order = ['Leads', 'Relances', 'Weendeal', 'Finanzen', 'Qwamplify', 'Kreid.fr', 'CRM', 'Bordereaux Budgeo', 'Base leads', 'Journal imports', 'Méthode & sources']
wb._sheets = [wb[n] for n in order]
wb.active = 0
for w in wb.worksheets: w.sheet_view.tabSelected = (w.title == 'Leads')
wb.save(OUT)
print(wb.sheetnames)
