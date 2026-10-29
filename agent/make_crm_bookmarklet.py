"""Génère « Installer le favori CRM.html » : un favori Chrome à lancer À LA MAIN (Corentin ou Yohan),
depuis sa propre session connectée à l'extranet IOBSP. Lecture seule, une page toutes les 0,6 s,
puis téléchargement de crm_extract_AAAA-MM-JJ.txt (toutes les fiches ; le filtrage sur les leads est fait par build_base.py).
Usage : python3 make_crm_bookmarklet.py "<dossier de sortie>" """
import sys, os, urllib.parse
JS = r"""(async()=>{
if(!location.href.includes('extranet-iobsp.com')){alert('Ouvre d\'abord la liste des contacts de l\'extranet IOBSP (choix_apporteur.asp), puis clique à nouveau sur le favori.');return;}
const rows=[];let pages=1;const T0=document.title;
async function page(n){const r=await fetch('/moduleapp/choix_apporteur.asp?PageNo='+n+'&OrderBy=creationdate&OrderDir=desc',{credentials:'include'});
 if(r.url.includes('login'))throw new Error('session');
 const html=new TextDecoder('windows-1252').decode(await r.arrayBuffer());const d=new DOMParser().parseFromString(html,'text/html');
 if(n==1){const m=d.body.textContent.match(/Page 1\/(\d+)/);if(m)pages=+m[1];}
 const t=[...d.querySelectorAll('table')].find(t=>t.textContent.includes('Réception contact'));if(!t)return;
 [...t.querySelectorAll('tr')].slice(1).forEach(tr=>{const td=[...tr.children];if(td.length<21)return;const c=td.map(x=>x.textContent.trim().replace(/;/g,','));
  const keys=[...td[11].childNodes].map(x=>x.textContent.replace(/\D/g,'')).filter(Boolean).map(s=>{s=s.startsWith('33')?s.slice(2):s;return s.replace(/^0+/,'').slice(-9);}).filter(s=>s.length==9);
  if(keys.length)rows.push([c[1],(c[3]||'').split(' ')[0],keys.join(','),c[5],c[6],c[13],c[14].slice(0,10),c[16].slice(0,10),c[17].slice(0,10),c[18].slice(0,10),c[20].slice(0,10)].join(';'));});}
try{await page(1);for(let p=2;p<=pages;p++){document.title='CRM '+p+'/'+pages;await new Promise(r=>setTimeout(r,600));await page(p);}}
catch(e){document.title=T0;alert('Session extranet expirée : reconnecte-toi puis relance le favori.');return;}
document.title=T0;
const dt=new Date(),f='crm_extract_'+dt.getFullYear()+'-'+String(dt.getMonth()+1).padStart(2,'0')+'-'+String(dt.getDate()).padStart(2,'0')+'.txt';
const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([rows.join('\n')],{type:'text/plain;charset=utf-8'}));a.download=f;document.body.appendChild(a);a.click();a.remove();
alert(rows.length+' fiches lues sur '+pages+' pages.\nEnregistre '+f+' dans OneDrive › Pôle distance › Suivi leads › Imports › CRM.');})()"""
href = 'javascript:' + urllib.parse.quote(' '.join(l.strip() for l in JS.split('\n')), safe="(){}[];,=+-*/<>!&|?:'.@_$\\%#")
html = f"""<!doctype html><html lang="fr"><head><meta charset="utf-8"><title>Favori « Lire le CRM » – Kreid</title>
<style>body{{font-family:Arial,sans-serif;max-width:720px;margin:40px auto;padding:0 16px;color:#1F2A44}}h1{{color:#1F3864}}
a.fav{{display:inline-block;background:#1F3864;color:#fff;padding:12px 20px;border-radius:6px;text-decoration:none;font-weight:bold}}
li{{margin:8px 0}}.n{{color:#666;font-size:13px}}</style></head><body>
<h1>Favori « Lire le CRM »</h1>
<p><a class="fav" href="{href}">📥 Lire le CRM Kreid</a></p>
<ol><li><b>Installation (une fois)</b> : affiche la barre de favoris de Chrome (⌘ Maj B), puis fais glisser le bouton bleu ci-dessus dans la barre.</li>
<li><b>Chaque jour, avant 9h15</b> (et si possible avant 17h45) : ouvre l'extranet IOBSP, connecte-toi, va sur la liste des contacts, clique sur le favori.</li>
<li>Laisse l'onglet ouvert environ 1 minute (le titre de l'onglet affiche « CRM 12/62 »…).</li>
<li>Enregistre le fichier proposé (<code>crm_extract_AAAA-MM-JJ.txt</code>) dans <b>OneDrive › Pôle distance › Suivi leads › Imports › CRM</b>.</li></ol>
<p class="n">Lecture seule : le favori ne modifie rien dans le CRM, ne saisit aucun mot de passe et lit une page toutes les 0,6 s.
Le fichier contient des données personnelles : il reste dans OneDrive (usage interne, RGPD).</p></body></html>"""
out = sys.argv[1] if len(sys.argv) > 1 else '.'
open(os.path.join(out, 'Installer le favori CRM.html'), 'w', encoding='utf-8').write(html)
print('OK', len(href), 'caractères')
