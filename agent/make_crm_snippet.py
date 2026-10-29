"""Format de sortie (1 ligne par fiche) : ref;collaborateur;téléphones;nom;prénom;réception;simulation;retour dossier;dépôt;accord;refus
Génère le script JavaScript d'extraction CRM (lecture seule) avec la liste des téléphones des leads (fournisseurs + site Kreid.fr).
Lecture page par page, une requête toutes les 0,6 s (pas de rafale vers l'extranet).
Usage : python3 make_crm_snippet.py "Leads data - Rentabilite.xlsx" > crm_snippet.js"""
import sys, re, openpyxl
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
tels = set()
for name, r0, col in (('Weendeal', 3, 7), ('Finanzen', 2, 10), ('Qwamplify', 2, 6), ('Kreid.fr', 2, 4)):
    if name not in wb.sheetnames: continue
    ws = wb[name]
    for r in range(r0, ws.max_row + 1):
        t = re.sub(r'\D', '', str(ws.cell(r, col).value or ''))
        if t.startswith('33'): t = t[2:]
        t = t.lstrip('0')[-9:]
        if len(t) == 9: tels.add(t)
P = ','.join(sorted(tels))
print(r"""(async()=>{
const P=new Set('%s'.split(','));
const rows=[]; let pages=1;
async function page(n){const r=await fetch('/moduleapp/choix_apporteur.asp?PageNo='+n+'&OrderBy=creationdate&OrderDir=desc',{credentials:'include'});
 const html=new TextDecoder('windows-1252').decode(await r.arrayBuffer()); const d=new DOMParser().parseFromString(html,'text/html');
 if(n==1){const m=d.body.textContent.match(/Page 1\/(\d+)/); if(m) pages=+m[1];}
 const t=[...d.querySelectorAll('table')].find(t=>t.textContent.includes('Réception contact')); if(!t) return;
 [...t.querySelectorAll('tr')].slice(1).forEach(tr=>{const td=[...tr.children]; if(td.length<21) return; const c=td.map(x=>x.textContent.trim());
  const keys=[...td[11].childNodes].map(x=>x.textContent.replace(/\D/g,'')).filter(Boolean).map(s=>{s=s.startsWith('33')?s.slice(2):s; return s.replace(/^0+/,'').slice(-9);});
  if(keys.some(k=>P.has(k))) rows.push([c[1],(c[3]||'').split(' ')[0],keys.join(','),c[5],c[6],c[13],c[14].slice(0,10),c[16].slice(0,10),c[17].slice(0,10),c[18].slice(0,10),c[20].slice(0,10)].join(';'));});}
await page(1); for(let p=2;p<=pages;p++){await new Promise(r=>setTimeout(r,600)); await page(p);}
document.body.innerHTML='<pre id="crm"></pre>'; document.getElementById('crm').textContent=rows.join('\n');
return 'OK '+rows.length+' fiches / '+pages+' pages';})()""" % P)
