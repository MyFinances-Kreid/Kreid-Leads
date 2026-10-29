"""Ajoute 3 TCD (format éprouvé type EPPlus : cache vide + refreshOnLoad) dans un xlsx existant."""
import zipfile, re, sys
from xml.sax.saxutils import escape

SRC, OUT = sys.argv[1], sys.argv[2]
zin = zipfile.ZipFile(SRC)
files = {n: zin.read(n) for n in zin.namelist()}

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'

FIELDS = ['Fournisseur', "Date d'achat", 'Année', 'Trimestre', 'Mois', 'Nom', 'Prénom', 'Prix lead (€)', 'Accord (1/0)',
       'CA si accord (€)', 'Financé (1/0)', 'CA réel encaissé (€)', 'Résultat lead (€)', 'Lead',
       'Dans CRM (1/0)', 'Simulation (1/0)', 'Dossier reçu (1/0)', 'Dépôt (1/0)', 'Accord CRM (1/0)', 'Refus CRM (1/0)',
       'Étape CRM', 'Alerte suivi', 'Collaborateur', 'CA en attente (€)', 'Téléphone (clé)', 'Ligne CRM', 'Ligne source',
       'N° relance', 'MAF financé (€)']
idx = {f: i for i, f in enumerate(FIELDS)}

# --- formats numériques (ids existants ou ajoutés)
st = files['xl/styles.xml'].decode()
fmts = dict((code, int(i)) for i, code in re.findall(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"/>', st))
nxt = max([int(i) for i in re.findall(r'numFmtId="(\d+)"', st)] + [200]) + 1
def fmt_id(code):
    global st, nxt
    key = escape(code, {'"': '&quot;'})
    if key in fmts: return fmts[key]
    st = st.replace('</numFmts>', f'<numFmt numFmtId="{nxt}" formatCode="{key}"/></numFmts>')
    st = re.sub(r'<numFmts count="(\d+)"', lambda m: f'<numFmts count="{int(m.group(1)) + 1}"', st)
    fmts[key] = nxt; nxt += 1
    return fmts[key]
EUR = fmt_id('#,##0" €"'); EUR2 = fmt_id('#,##0.00" €"'); PCT = fmt_id('0.0%'); INT = 3
files['xl/styles.xml'] = st.encode()

# --- cache
cf = ''.join(f'<cacheField name="{escape(f, {chr(39): "&apos;"})}" numFmtId="0"><sharedItems/></cacheField>' for f in FIELDS)
cache = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
         f'<pivotCacheDefinition xmlns="{MAIN}" xmlns:r="{REL}" r:id="rId1" refreshOnLoad="1" refreshedBy="Kreid" '
         f'refreshedDate="46323" createdVersion="3" refreshedVersion="3" minRefreshableVersion="3" recordCount="0" upgradeOnRefresh="1">'
         f'<cacheSource type="worksheet"><worksheetSource name="Base_TCD"/></cacheSource>'
         f'<cacheFields count="{len(FIELDS)}">{cf}</cacheFields></pivotCacheDefinition>')
records = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheRecords xmlns="{MAIN}" xmlns:r="{REL}" count="0"/>'

def table(name, ref, rows, cols, pages, datas):
    pf = []
    for f in FIELDS:
        ax = 'axisRow' if f in rows else 'axisCol' if f in cols else 'axisPage' if f in pages else None
        isdata = any(d[0] == f for d in datas)
        attrs = (f' axis="{ax}"' if ax else '') + (' dataField="1"' if isdata else '') + ' compact="0" outline="0" showAll="0"'
        pf.append(f'<pivotField{attrs}><items count="1"><item t="default"/></items></pivotField>' if ax else f'<pivotField{attrs}/>')
    x = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
         f'<pivotTableDefinition xmlns="{MAIN}" name="{name}" cacheId="1" dataOnRows="0" applyNumberFormats="0" applyBorderFormats="0" '
         f'applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Valeurs" '
         f'updatedVersion="3" minRefreshableVersion="3" showMemberPropertyTips="0" useAutoFormatting="1" itemPrintTitles="1" '
         f'createdVersion="3" indent="0" compact="0" compactData="0" outline="0" outlineData="0" gridDropZones="0" multipleFieldFilters="0">'
         f'<location ref="{ref}" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"'
         + (f' rowPageCount="{len(pages)}" colPageCount="1"' if pages else '') + '/>'
         f'<pivotFields count="{len(FIELDS)}">{"".join(pf)}</pivotFields>')
    if rows: x += f'<rowFields count="{len(rows)}">' + ''.join(f'<field x="{idx[f]}"/>' for f in rows) + '</rowFields>'
    cfs = [idx[f] for f in cols] + ([-2] if len(datas) > 1 else [])
    if cfs: x += f'<colFields count="{len(cfs)}">' + ''.join(f'<field x="{c}"/>' for c in cfs) + '</colFields>'
    if pages: x += f'<pageFields count="{len(pages)}">' + ''.join(f'<pageField fld="{idx[f]}" hier="-1"/>' for f in pages) + '</pageFields>'
    x += f'<dataFields count="{len(datas)}">'
    for f, fn, lab, nf in datas:
        x += (f'<dataField name="{escape(lab, {chr(39): "&apos;"})}" fld="{idx[f]}"' + (' subtotal="average"' if fn == 'avg' else '')
              + f' baseField="0" baseItem="0" numFmtId="{nf}"/>')
    x += ('</dataFields><pivotTableStyleInfo name="PivotStyleMedium2" showRowHeaders="1" showColHeaders="1" '
          'showRowStripes="1" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>')
    return x

L_ = ('Lead', 'sum', 'Nb leads', INT); C_ = ('Prix lead (€)', 'sum', 'Coût leads (€)', EUR)
CAR = ('CA réel encaissé (€)', 'sum', 'CA réel (€)', EUR); RES = ('Résultat lead (€)', 'sum', 'Résultat (€)', EUR)
D1 = [L_, C_, ('Prix lead (€)', 'avg', 'Prix moyen lead (€)', EUR2), ('Accord CRM (1/0)', 'sum', 'Accords CRM', INT),
      ('Financé (1/0)', 'sum', 'Dossiers financés', INT), ('Financé (1/0)', 'avg', 'Taux de transfo', PCT), CAR, ('MAF financé (€)', 'sum', 'MAF financé', EUR),
      ('CA en attente (€)', 'sum', 'CA en attente (€) ', EUR), RES]
D4 = [L_, ('Dans CRM (1/0)', 'sum', 'Dans CRM', INT), ('Simulation (1/0)', 'sum', 'Simulations', INT),
      ('Dossier reçu (1/0)', 'sum', 'Dossiers reçus', INT), ('Dossier reçu (1/0)', 'avg', 'Taux dossier reçu', PCT),
      ('Dépôt (1/0)', 'sum', 'Dépôts banque', INT), ('Accord CRM (1/0)', 'sum', 'Accords banque', INT),
      ('Accord CRM (1/0)', 'avg', "Taux d'accord", PCT), ('Refus CRM (1/0)', 'sum', 'Refus / sans suite', INT)]
D2 = [L_, C_, ('Dossier reçu (1/0)', 'sum', 'Dossiers reçus', INT), ('Accord CRM (1/0)', 'sum', 'Accords banque', INT), CAR, RES]
D3 = [RES]
tables = [table('TCD_Fournisseur', 'A10', ['Fournisseur'], [], ['Année', 'Trimestre'], D1),
          table('TCD_Entonnoir', 'A22', ['Fournisseur'], [], ['Année', 'Trimestre'], D4),
          table('TCD_Periode', 'A36', ['Mois', 'Fournisseur'], [], ['Année'], D2),
          table('TCD_Matrice', 'J36', ['Mois'], ['Fournisseur'], ['Année'], D3)]

hdr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
rels = lambda items: hdr + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + ''.join(
    f'<Relationship Id="{i}" Type="{RT}{t}" Target="{g}"/>' for i, t, g in items) + '</Relationships>'
files['xl/pivotCache/pivotCacheDefinition1.xml'] = cache.encode()
files['xl/pivotCache/pivotCacheRecords1.xml'] = records.encode()
files['xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels'] = rels([('rId1', 'pivotCacheRecords', 'pivotCacheRecords1.xml')]).encode()
for i, t in enumerate(tables, 1):
    files[f'xl/pivotTables/pivotTable{i}.xml'] = t.encode()
    files[f'xl/pivotTables/_rels/pivotTable{i}.xml.rels'] = rels([('rId1', 'pivotCacheDefinition', '../pivotCache/pivotCacheDefinition1.xml')]).encode()

# feuille Leads = 1re feuille du classeur
wbx = files['xl/workbook.xml'].decode()
first_rid = re.search(r'<sheet [^>]*r:id="([^"]+)"', wbx).group(1)
wbrels = files['xl/_rels/workbook.xml.rels'].decode()
target = re.search(rf'Id="{first_rid}"[^>]*Target="([^"]+)"', wbrels) or re.search(rf'Target="([^"]+)"[^>]*Id="{first_rid}"', wbrels)
sheet_path = 'xl/' + target.group(1)
srels = sheet_path.replace('worksheets/', 'worksheets/_rels/') + '.rels'
assert srels not in files, 'la feuille a déjà des relations'
files[srels] = rels([(f'rIdPT{i}', 'pivotTable', f'../pivotTables/pivotTable{i}.xml') for i in range(1, len(tables) + 1)]).encode()
wbrels = wbrels.replace('</Relationships>', f'<Relationship Id="rIdPC1" Type="{RT}pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/></Relationships>')
files['xl/_rels/workbook.xml.rels'] = wbrels.encode()
dn = '<definedName name="Base_TCD">OFFSET(&apos;Base leads&apos;!$A$3,0,0,&apos;Base leads&apos;!$AE$8+1,29)</definedName>'
wbx = wbx.replace('</definedNames>', dn + '</definedNames>') if '</definedNames>' in wbx else wbx.replace('</sheets>', '</sheets><definedNames>' + dn + '</definedNames>')
pc = '<pivotCaches><pivotCache cacheId="1" r:id="rIdPC1"/></pivotCaches>'
wbx = re.sub(r'(<calcPr[^>]*/>)', r'\1' + pc, wbx, count=1)
assert pc in wbx
files['xl/workbook.xml'] = wbx.encode()

ct = files['[Content_Types].xml'].decode()
ct = re.sub(r'<Override PartName="[^"]*\.rels" ContentType="[^"]*"/>', '', ct)   # overrides parasites LibreOffice
add = '<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>' \
      '<Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>' + \
      ''.join(f'<Override PartName="/xl/pivotTables/pivotTable{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>' for i in range(1, len(tables) + 1))
files['[Content_Types].xml'] = ct.replace('</Types>', add + '</Types>').encode()

with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
    order = ['[Content_Types].xml'] + [n for n in files if n != '[Content_Types].xml']
    for n in order: z.writestr(n, files[n])
print('ok', sheet_path, EUR, EUR2, PCT)
