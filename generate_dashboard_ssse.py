"""
Dashboard Qualité — Semoule SSSE
=================================
VERSION CORRIGÉE — Compatible SSO / ADFS entreprise

3 modes d'authentification disponibles (choisis selon ton organisation) :
  MODE 1 : Lien de partage SharePoint public  ← le plus simple, aucune auth
  MODE 2 : Client ID + Secret Azure AD        ← si ton admin peut créer une App
  MODE 3 : Login / mot de passe direct        ← si M365 standard (pas SSO fédéré)

Change la variable AUTH_MODE ci-dessous selon ce qui fonctionne.

Prérequis :
    pip install office365-rest-python-client pandas openpyxl plotly jinja2 requests python-dotenv
"""

import io, os, json, requests
from datetime import datetime

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
from jinja2 import Template

# ============================================================
#  CONFIG — modifie ces valeurs
# ============================================================

# Choix du mode d'authentification : "sharelink" | "appcredential" | "userpassword"
AUTH_MODE = os.getenv("AUTH_MODE", "sharelink")

# --- MODE 1 : Lien de partage (le plus simple) ---
# Ouvre le fichier Excel sur SharePoint → Partager → Copier le lien
# Ajoute &download=1 à la fin
SHARE_LINK = os.getenv("SHARE_LINK",
    "https://roseblanchetn.sharepoint.com/:x:/r/sites/SDAHSESTPA/_layouts/15/Doc.aspx"
    "?sourcedoc=%7B0761FA65-3D84-4B10-B009-8CA5BF050C98%7D"
    "&file=2025%20CQ%20SBOULA%20FOCQT01_02.xlsx&action=default&download=1"
)

# --- MODE 2 : App Registration Azure AD (demande à ton admin IT) ---
TENANT_ID     = os.getenv("TENANT_ID",     "")
CLIENT_ID     = os.getenv("CLIENT_ID",     "")
CLIENT_SECRET = os.getenv("CLIENT_SECRET", "")
SP_SITE_URL   = os.getenv("SP_URL",        "https://roseblanchetn.sharepoint.com/sites/SDAHSESTPA")

# --- MODE 3 : Login / mot de passe (M365 standard uniquement) ---
SP_USER     = os.getenv("SP_USER",     "")
SP_PASSWORD = os.getenv("SP_PASSWORD", "")
SP_URL      = os.getenv("SP_URL",      "https://roseblanchetn.sharepoint.com/sites/SDAHSESTPA")

# Chemin du fichier Excel dans SharePoint (pour l'upload du HTML)
EXCEL_PATH  = os.getenv("EXCEL_PATH",  "Documents Partages/2025 CQ SBOULA FOCQT01_02.xlsx")
SHEET_NAME  = "Semoule SSSE"
OUTPUT_HTML = "dashboard_ssse.html"

# Colonnes du fichier
COL_DATE     = "Date"
COL_LOT      = "N°lot"
COL_ETAPE    = "Etape"
COL_PROBLEME = "Probléme"
COL_NOTIF    = "Notif"
COL_FLUX     = "Flux_Statut"
COL_ECHANT   = "N° de l'échantillon"

MOIS_FR = {
    1:"Janvier",2:"Février",3:"Mars",4:"Avril",5:"Mai",6:"Juin",
    7:"Juillet",8:"Août",9:"Septembre",10:"Octobre",11:"Novembre",12:"Décembre"
}

# ============================================================
#  ETAPE 1 — Lecture du fichier Excel
# ============================================================

def read_excel() -> pd.DataFrame:
    """Lit le fichier Excel selon le mode d'auth choisi."""

    if AUTH_MODE == "sharelink":
        return _read_via_sharelink()
    elif AUTH_MODE == "appcredential":
        return _read_via_app_credential()
    elif AUTH_MODE == "userpassword":
        return _read_via_userpassword()
    else:
        raise ValueError(f"AUTH_MODE inconnu : {AUTH_MODE}")


def _read_via_sharelink() -> pd.DataFrame:
    """
    MODE 1 — Lien de partage SharePoint.
    Le fichier doit avoir un lien de partage actif (interne organisation).
    """
    print(f"  Mode : lien de partage SharePoint")
    # Ajouter le cookie de session si nécessaire
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }

    # Essai 1 : téléchargement direct via l'ID unique du fichier
    direct_url = (
        "https://roseblanchetn.sharepoint.com/sites/SDAHSESTPA"
        "/_layouts/15/download.aspx"
        "?UniqueId=0761FA65-3D84-4B10-B009-8CA5BF050C98"
    )

    resp = requests.get(direct_url, headers=headers, timeout=30)

    # Si redirigé vers login → essayer avec le lien complet
    if resp.status_code != 200 or b"<!DOCTYPE" in resp.content[:100]:
        print("  → Redirection login détectée, essai avec SHARE_LINK...")
        resp = requests.get(SHARE_LINK, headers=headers, timeout=30, allow_redirects=True)

    if resp.status_code != 200 or b"<!DOCTYPE" in resp.content[:100]:
        raise ConnectionError(
            "Impossible de télécharger le fichier via lien public.\n"
            "→ Essaie AUTH_MODE='appcredential' ou demande à ton admin IT\n"
            "  de créer une App Registration Azure AD."
        )

    df = pd.read_excel(io.BytesIO(resp.content), sheet_name=SHEET_NAME, header=0)
    print(f"  {len(df)} lignes lues")
    return df


def _read_via_app_credential() -> pd.DataFrame:
    """
    MODE 2 — Application Azure AD (Client Credentials).
    Nécessite : TENANT_ID, CLIENT_ID, CLIENT_SECRET
    Demande à ton admin IT de créer une App Registration avec
    permission Sites.Read.All sur Microsoft Graph.
    """
    print("  Mode : Azure AD App Credential")

    # 1. Obtenir le token
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    resp = requests.post(token_url, data={
        "grant_type"   : "client_credentials",
        "client_id"    : CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope"        : "https://graph.microsoft.com/.default"
    })
    resp.raise_for_status()
    token = resp.json()["access_token"]

    # 2. Récupérer l'ID du site
    site_resp = requests.get(
        "https://graph.microsoft.com/v1.0/sites/roseblanchetn.sharepoint.com:/sites/SDAHSESTPA",
        headers={"Authorization": f"Bearer {token}"}
    )
    site_resp.raise_for_status()
    site_id = site_resp.json()["id"]

    # 3. Télécharger le fichier
    file_url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}"
        f"/drive/root:/{EXCEL_PATH}:/content"
    )
    file_resp = requests.get(file_url, headers={"Authorization": f"Bearer {token}"})
    file_resp.raise_for_status()

    df = pd.read_excel(io.BytesIO(file_resp.content), sheet_name=SHEET_NAME, header=0)
    print(f"  {len(df)} lignes lues")
    return df


def _read_via_userpassword() -> pd.DataFrame:
    """
    MODE 3 — Login / mot de passe.
    Fonctionne uniquement avec M365 standard (pas SSO/ADFS fédéré).
    """
    from office365.runtime.auth.user_credential import UserCredential
    from office365.sharepoint.client_context import ClientContext

    print("  Mode : Login / mot de passe")
    creds = UserCredential(SP_USER, SP_PASSWORD)
    ctx   = ClientContext(SP_URL).with_credentials(creds)
    web   = ctx.web
    ctx.load(web)
    ctx.execute_query()
    print(f"  Connecté : {web.title}")

    buf         = io.BytesIO()
    server_path = f"{ctx.web.server_relative_url}/{EXCEL_PATH}"
    ctx.web.get_file_by_server_relative_url(server_path)\
        .download(buf).execute_query()
    buf.seek(0)
    df = pd.read_excel(buf, sheet_name=SHEET_NAME, header=0)
    print(f"  {len(df)} lignes lues")
    return df


# ============================================================
#  ETAPE 2 — Nettoyage des données SSSE
# ============================================================

def prepare_data(df: pd.DataFrame):
    df[COL_DATE]  = pd.to_datetime(df[COL_DATE], errors="coerce")
    df[COL_ETAPE] = df[COL_ETAPE].astype(str).str.strip().str.title()
    df["Année"]    = df[COL_DATE].dt.year.astype("Int64").astype(str)
    df["Mois_num"] = df[COL_DATE].dt.month.astype("Int64")
    df["Jour"]     = df[COL_DATE].dt.date.astype(str)

    df_anom = df[df[COL_PROBLEME].notna()].copy()
    df_anom[COL_PROBLEME] = df_anom[COL_PROBLEME].astype(str).str.strip()

    print(f"  Analyses  : {len(df)}")
    print(f"  Anomalies : {len(df_anom)}")
    return df, df_anom


# ============================================================
#  ETAPE 3 — Sérialisation JSON pour le dashboard
# ============================================================

def serialize(df_all, df_anom) -> dict:
    rows = []
    for _, r in df_anom.iterrows():
        rows.append({
            "date"    : str(r["Jour"]),
            "annee"   : str(r["Année"]),
            "mois_num": int(r["Mois_num"]) if pd.notna(r["Mois_num"]) else 0,
            "mois_nom": MOIS_FR.get(int(r["Mois_num"]), "") if pd.notna(r["Mois_num"]) else "",
            "probleme": str(r[COL_PROBLEME]),
            "etape"   : str(r[COL_ETAPE]),
            "lot"     : str(r[COL_LOT])    if pd.notna(r[COL_LOT])    else "",
            "echant"  : str(r[COL_ECHANT]) if pd.notna(r[COL_ECHANT]) else "",
            "notif"   : str(r[COL_NOTIF])  if pd.notna(r[COL_NOTIF])  else "Non",
            "flux"    : str(r[COL_FLUX])   if pd.notna(r[COL_FLUX])   else "",
        })
    return {
        "total_analyses" : len(df_all),
        "total_anomalies": len(df_anom),
        "generated_at"   : datetime.now().strftime("%d/%m/%Y %H:%M"),
        "data"           : rows
    }


# ============================================================
#  ETAPE 4 — Template HTML du dashboard
# ============================================================

DASHBOARD_HTML = """<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Dashboard Qualité — SSSE</title>
<script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Tahoma,sans-serif;background:#0f1923;color:#e8edf2;padding:14px;min-height:100vh}
.header{background:linear-gradient(135deg,#0d2137,#1a4a72);border:1px solid #1e5a8a;border-radius:12px;padding:14px 20px;margin-bottom:14px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px}
.header h1{font-size:16px;font-weight:600;color:#fff}
.header .meta{font-size:11px;color:#7a9ab8;margin-top:3px}
.badge{background:#0d2a42;border:1px solid #2e6a9a;color:#7bc8f0;font-size:11px;padding:4px 12px;border-radius:20px}
.filters{background:#0d1e2e;border:1px solid #1a3a52;border-radius:10px;padding:12px 16px;margin-bottom:14px;display:flex;flex-wrap:wrap;gap:10px;align-items:center}
.f-group{display:flex;align-items:center;gap:6px}
.f-group label{font-size:11px;color:#5a8ab8;white-space:nowrap}
.filters select{background:#0f2840;border:1px solid #1e5a8a;color:#c8dff0;border-radius:6px;padding:5px 10px;font-size:12px;cursor:pointer;outline:none}
.filters select:focus{border-color:#2e8adf}
.btn-reset{background:#1a3a52;border:1px solid #2e6a9a;color:#7bc8f0;border-radius:6px;padding:5px 14px;font-size:12px;cursor:pointer}
.filter-info{font-size:11px;color:#3a6a9a;margin-left:auto}
.kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(155px,1fr));gap:12px;margin-bottom:14px}
.kpi{background:#0d1e2e;border:1px solid #1a3a52;border-radius:10px;padding:14px 18px;position:relative;overflow:hidden}
.kpi::before{content:'';position:absolute;left:0;top:0;bottom:0;width:3px;border-radius:3px 0 0 3px}
.kpi.c-blue::before{background:#2e8adf}.kpi.c-red::before{background:#e24b4a}
.kpi.c-amber::before{background:#f0a030}.kpi.c-green::before{background:#27ae8f}
.kpi.c-purple::before{background:#8e44ad}
.kpi .lbl{font-size:10px;color:#5a7a9a;text-transform:uppercase;letter-spacing:.5px;margin-bottom:5px}
.kpi .val{font-size:26px;font-weight:700;color:#fff;line-height:1}
.kpi .sub{font-size:11px;color:#3a6a9a;margin-top:5px}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px}
.card{background:#0d1e2e;border:1px solid #1a3a52;border-radius:10px;padding:14px;overflow:hidden}
.card.wide{grid-column:1/-1}
.card-title{font-size:11px;color:#5a8ab8;text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px;font-weight:500}
.table-wrap{overflow-x:auto;max-height:360px;overflow-y:auto}
.table-wrap::-webkit-scrollbar{width:4px;height:4px}
.table-wrap::-webkit-scrollbar-thumb{background:#1a4a72;border-radius:2px}
table{width:100%;border-collapse:collapse;font-size:12px}
thead{position:sticky;top:0;z-index:1}
thead th{background:#0a1724;color:#5a8ab8;text-align:left;padding:9px 12px;font-weight:500;border-bottom:1px solid #1a3a52;white-space:nowrap}
tbody tr{border-bottom:1px solid #0d1e2e;transition:background .1s}
tbody tr:hover{background:#0f2840}
tbody td{padding:8px 12px;color:#c8dff0;vertical-align:middle}
.no-data{text-align:center;padding:40px;color:#3a5a7a;font-size:13px}
.pill{display:inline-block;padding:2px 9px;border-radius:10px;font-size:11px;font-weight:500;white-space:nowrap}
.p-red{background:#3a0f0f;color:#f08080;border:1px solid #5a1e1e}
.p-amber{background:#3a2a00;color:#f0c060;border:1px solid #5a4010}
.p-blue{background:#0a1e3a;color:#70b0f0;border:1px solid #1a3a6a}
.p-green{background:#0a2a1a;color:#50c090;border:1px solid #1a4a2a}
.p-purple{background:#1e0a3a;color:#b070f0;border:1px solid #3a1a6a}
.p-gray{background:#1a2a3a;color:#8ab0c0;border:1px solid #2a3a4a}
@media(max-width:600px){.grid{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="header">
  <div>
    <h1>Dashboard Qualité — Semoule SSSE</h1>
    <div class="meta" id="meta-info">Chargement...</div>
  </div>
  <span class="badge">T.P.A · SBOULA</span>
</div>
<div class="filters">
  <div class="f-group"><label>Année</label><select id="f-annee"><option value="">Toutes</option></select></div>
  <div class="f-group"><label>Mois</label><select id="f-mois"><option value="">Tous</option></select></div>
  <div class="f-group"><label>Date</label><select id="f-date"><option value="">Toutes</option></select></div>
  <div class="f-group"><label>Problème</label><select id="f-prob"><option value="">Tous</option></select></div>
  <div class="f-group"><label>Étape</label><select id="f-etape"><option value="">Toutes</option></select></div>
  <button class="btn-reset" onclick="resetFilters()">↺ Réinitialiser</button>
  <span class="filter-info" id="filter-info"></span>
</div>
<div class="kpi-row">
  <div class="kpi c-blue"><div class="lbl">Total analyses</div><div class="val" id="k-total">—</div><div class="sub">Feuille Semoule SSSE</div></div>
  <div class="kpi c-red"><div class="lbl">Anomalies</div><div class="val" id="k-anom">—</div><div class="sub" id="k-anom-sub">—</div></div>
  <div class="kpi c-amber"><div class="lbl">Taux anomalies</div><div class="val" id="k-taux">—</div><div class="sub">sur total analyses</div></div>
  <div class="kpi c-green"><div class="lbl">Notifiées (Oui)</div><div class="val" id="k-notif">—</div><div class="sub" id="k-notif-sub">—</div></div>
  <div class="kpi c-purple"><div class="lbl">Problème dominant</div><div class="val" id="k-top" style="font-size:13px;margin-top:6px;line-height:1.4">—</div></div>
</div>
<div class="grid">
  <div class="card"><div class="card-title">Anomalies par type de problème</div><div id="ch-type"></div></div>
  <div class="card"><div class="card-title">Répartition par étape</div><div id="ch-etape"></div></div>
  <div class="card wide"><div class="card-title">Évolution mensuelle</div><div id="ch-trend"></div></div>
  <div class="card wide">
    <div class="card-title">Détail des anomalies <span id="tbl-count" style="color:#3a6a9a;font-weight:400;margin-left:8px"></span></div>
    <div class="table-wrap"><table>
      <thead><tr><th>Date</th><th>N° Lot</th><th>N° Échantillon</th><th>Étape</th><th>Problème</th><th>Notifié</th><th>Flux</th></tr></thead>
      <tbody id="tbl-body"></tbody>
    </table></div>
  </div>
</div>
<script>
const _DATA=__DATA_JSON__;
const TOTAL_ANALYSES=_DATA.total_analyses;
const GENERATED_AT=_DATA.generated_at;
let ALL_DATA=_DATA.data;
const MOIS_FR=['','Janvier','Février','Mars','Avril','Mai','Juin','Juillet','Août','Septembre','Octobre','Novembre','Décembre'];
const PROB_COLORS={'PIQUAGE BRUN':'#c0392b','PIQUAGE NOIR':'#8e44ad','GRANULOMETRIE':'#e67e22','COULEUR':'#2980b9','TENEUR EN EAU ELEVEE':'#1abc9c','TENEUR EN EAU FAIBLE':'#16a085','MELANGE PRODUITS':'#f39c12','CHARANCONS':'#d35400','RHEOLOGIE':'#7f8c8d'};
const PROB_PILL={'PIQUAGE BRUN':'p-red','PIQUAGE NOIR':'p-purple','GRANULOMETRIE':'p-amber','COULEUR':'p-blue','TENEUR EN EAU ELEVEE':'p-green','TENEUR EN EAU FAIBLE':'p-green','MELANGE PRODUITS':'p-amber','CHARANCONS':'p-red','RHEOLOGIE':'p-gray'};
const LAY={paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{color:'#c8dff0',size:11},margin:{l:10,r:10,t:10,b:10}};
const CFG={displayModeBar:false,responsive:true};
function unique(k){return[...new Set(ALL_DATA.map(r=>r[k]))].filter(Boolean);}
function fillSel(id,vals,lbFn){const s=document.getElementById(id);const p=s.value;while(s.options.length>1)s.remove(1);vals.forEach(v=>{const o=new Option(lbFn?lbFn(v):v,v);s.appendChild(o);});if(p)s.value=p;}
function populateFilters(){
  fillSel('f-annee',unique('annee').sort());
  fillSel('f-mois',unique('mois_num').sort((a,b)=>+a-+b),v=>MOIS_FR[+v]||v);
  fillSel('f-date',unique('date').sort().reverse());
  fillSel('f-prob',unique('probleme').sort());
  fillSel('f-etape',unique('etape').sort());
}
function getFiltered(){
  const a=document.getElementById('f-annee').value;
  const m=document.getElementById('f-mois').value;
  const d=document.getElementById('f-date').value;
  const p=document.getElementById('f-prob').value;
  const e=document.getElementById('f-etape').value;
  return ALL_DATA.filter(r=>(!a||r.annee===a)&&(!m||r.mois_num===+m)&&(!d||r.date===d)&&(!p||r.probleme===p)&&(!e||r.etape===e));
}
function render(){
  const fd=getFiltered();const n=fd.length;
  const isF=n<ALL_DATA.length;
  document.getElementById('filter-info').textContent=isF?`Filtre actif : ${n} / ${ALL_DATA.length} anomalies`:'';
  document.getElementById('k-total').textContent=TOTAL_ANALYSES.toLocaleString('fr');
  document.getElementById('k-anom').textContent=n.toLocaleString('fr');
  document.getElementById('k-anom-sub').textContent=isF?`(total : ${ALL_DATA.length})`:'période complète';
  document.getElementById('k-taux').textContent=TOTAL_ANALYSES>0?(n/TOTAL_ANALYSES*100).toFixed(1)+'%':'—';
  const notifOui=fd.filter(r=>r.notif==='Oui').length;
  document.getElementById('k-notif').textContent=notifOui.toLocaleString('fr');
  document.getElementById('k-notif-sub').textContent=n>0?`${(notifOui/n*100).toFixed(0)}% des anomalies`:'—';
  const cntP={};fd.forEach(r=>cntP[r.probleme]=(cntP[r.probleme]||0)+1);
  const topP=Object.entries(cntP).sort((a,b)=>b[1]-a[1]);
  document.getElementById('k-top').textContent=topP.length?`${topP[0][0]} (${topP[0][1]})`:'—';
  document.getElementById('meta-info').textContent=`Généré le ${GENERATED_AT} · SharePoint · Semoule SSSE`;
  // Barres
  if(!topP.length){Plotly.newPlot('ch-type',[],{...LAY,height:260},CFG);}
  else{const s=topP.slice(0,9).reverse();Plotly.newPlot('ch-type',[{type:'bar',orientation:'h',y:s.map(x=>x[0]),x:s.map(x=>x[1]),marker:{color:s.map(x=>PROB_COLORS[x[0]]||'#2e8adf')},text:s.map(x=>x[1]),textposition:'outside',hovertemplate:'%{y}: %{x}<extra></extra>'}],{...LAY,height:260,margin:{l:190,r:50,t:10,b:20},xaxis:{gridcolor:'#1a3a52'},yaxis:{gridcolor:'rgba(0,0,0,0)'}},CFG);}
  // Camembert
  const cntE={};fd.forEach(r=>cntE[r.etape]=(cntE[r.etape]||0)+1);
  const topE=Object.entries(cntE).sort((a,b)=>b[1]-a[1]);
  if(!topE.length){Plotly.newPlot('ch-etape',[],{...LAY,height:260},CFG);}
  else{Plotly.newPlot('ch-etape',[{type:'pie',labels:topE.map(x=>x[0]),values:topE.map(x=>x[1]),hole:0.42,marker:{colors:['#2e8adf','#27ae8f','#e67e22','#8e44ad','#c0392b','#f39c12','#1abc9c']},textinfo:'label+percent',textfont:{size:10,color:'#c8dff0'},hovertemplate:'%{label}: %{value}<extra></extra>'}],{...LAY,height:260,showlegend:true,legend:{font:{size:10,color:'#7a9ab8'},bgcolor:'rgba(0,0,0,0)'}},CFG);}
  // Tendance
  const cntM={};fd.forEach(r=>{const k=r.annee+'-'+String(r.mois_num).padStart(2,'0');cntM[k]=(cntM[k]||0)+1;});
  const mKeys=Object.keys(cntM).sort();
  const mLabels=mKeys.map(k=>{const[y,m]=k.split('-');return MOIS_FR[+m].slice(0,3)+' '+y;});
  if(!mKeys.length){Plotly.newPlot('ch-trend',[],{...LAY,height:220},CFG);}
  else{Plotly.newPlot('ch-trend',[{type:'scatter',mode:'lines+markers',x:mLabels,y:mKeys.map(k=>cntM[k]),line:{color:'#2e8adf',width:2,shape:'spline'},marker:{color:'#2e8adf',size:7},fill:'tozeroy',fillcolor:'rgba(46,138,223,0.1)',hovertemplate:'%{x}: <b>%{y}</b> anomalie(s)<extra></extra>'}],{...LAY,height:220,margin:{l:40,r:20,t:10,b:70},xaxis:{gridcolor:'#1a3a52',tickangle:-45,tickfont:{size:10}},yaxis:{gridcolor:'#1a3a52',tickfont:{size:10}}},CFG);}
  // Tableau
  const tbody=document.getElementById('tbl-body');
  const display=fd.slice(0,100);
  document.getElementById('tbl-count').textContent=fd.length>100?`(100 sur ${fd.length})`:fd.length>0?`(${fd.length})`:'';;
  if(!display.length){tbody.innerHTML='<tr><td colspan="7" class="no-data">Aucune anomalie</td></tr>';return;}
  tbody.innerHTML=display.map(r=>`<tr>
    <td style="color:#7a9ab8">${r.date}</td><td>${r.lot}</td>
    <td style="font-size:11px;color:#5a8ab8">${r.echant}</td>
    <td><span class="pill p-blue">${r.etape}</span></td>
    <td><span class="pill ${PROB_PILL[r.probleme]||'p-gray'}">${r.probleme}</span></td>
    <td style="color:${r.notif==='Oui'?'#50c090':'#e07070'};font-weight:500">${r.notif}</td>
    <td style="color:#3a6a8a;font-size:11px">${r.flux}</td>
  </tr>`).join('');
}
function resetFilters(){['f-annee','f-mois','f-date','f-prob','f-etape'].forEach(id=>document.getElementById(id).value='');render();}
['f-annee','f-mois','f-date','f-prob','f-etape'].forEach(id=>document.getElementById(id).addEventListener('change',render));
populateFilters();render();
</script>
</body>
</html>"""


def generate_html(payload: dict) -> str:
    data_json = json.dumps(payload, ensure_ascii=False)
    return DASHBOARD_HTML.replace("__DATA_JSON__", data_json)


# ============================================================
#  ETAPE 5 — Upload HTML sur SharePoint
# ============================================================

def upload_html(html: str) -> str:
    """
    Upload dashboard.html sur SharePoint.
    Utilise le même mode d'auth que pour la lecture.
    """
    if AUTH_MODE == "appcredential":
        return _upload_via_graph(html)
    elif AUTH_MODE == "userpassword":
        return _upload_via_office365(html)
    else:
        # Mode sharelink : sauvegarde locale, upload manuel
        local_path = OUTPUT_HTML
        with open(local_path, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"  Sauvegardé localement : {local_path}")
        print("  ⚠ Upload automatique non disponible en mode 'sharelink'.")
        print("  → Dépose manuellement le fichier sur SharePoint.")
        return local_path


def _upload_via_graph(html: str) -> str:
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    resp = requests.post(token_url, data={
        "grant_type":"client_credentials","client_id":CLIENT_ID,
        "client_secret":CLIENT_SECRET,"scope":"https://graph.microsoft.com/.default"
    })
    resp.raise_for_status()
    token = resp.json()["access_token"]
    site_resp = requests.get(
        "https://graph.microsoft.com/v1.0/sites/roseblanchetn.sharepoint.com:/sites/SDAHSESTPA",
        headers={"Authorization":f"Bearer {token}"}
    )
    site_id = site_resp.json()["id"]
    folder = "/".join(EXCEL_PATH.split("/")[:-1])
    upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{folder}/{OUTPUT_HTML}:/content"
    up = requests.put(upload_url, headers={"Authorization":f"Bearer {token}","Content-Type":"text/html"}, data=html.encode("utf-8"))
    up.raise_for_status()
    url = SP_SITE_URL.split("/sites/")[0] + up.json().get("webUrl","")
    return url


def _upload_via_office365(html: str) -> str:
    from office365.runtime.auth.user_credential import UserCredential
    from office365.sharepoint.client_context import ClientContext
    creds  = UserCredential(SP_USER, SP_PASSWORD)
    ctx    = ClientContext(SP_URL).with_credentials(creds)
    folder = "/".join(EXCEL_PATH.split("/")[:-1])
    srv    = f"{ctx.web.server_relative_url}/{folder}"
    up     = ctx.web.get_folder_by_server_relative_url(srv)\
                 .upload_file(OUTPUT_HTML, html.encode("utf-8")).execute_query()
    return SP_URL.split("/sites/")[0] + up.serverRelativeUrl


# ============================================================
#  MAIN
# ============================================================

if __name__ == "__main__":
    print("\n" + "="*50)
    print("  Dashboard Qualité SSSE")
    print(f"  Mode auth : {AUTH_MODE}")
    print("="*50 + "\n")

    print("[1/4] Lecture du fichier Excel SharePoint...")
    df_raw = read_excel()

    print("\n[2/4] Préparation des données...")
    df_all, df_anom = prepare_data(df_raw)

    print("\n[3/4] Génération du dashboard HTML...")
    payload = serialize(df_all, df_anom)
    html    = generate_html(payload)
    print(f"  Taille HTML : {len(html)//1024} Ko")

    print("\n[4/4] Upload sur SharePoint...")
    url = upload_html(html)

    print("\n" + "="*50)
    print("  TERMINÉ")
    print("="*50)
    print(f"  Anomalies  : {payload['total_anomalies']}")
    print(f"  Analyses   : {payload['total_analyses']}")
    print(f"  Généré le  : {payload['generated_at']}")
    print(f"\n  Fichier    : {url}")
    print(f"\n  Power Apps → Web viewer → URL :")
    print(f'  "{url}"')
