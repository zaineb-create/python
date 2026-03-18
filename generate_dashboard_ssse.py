"""
Dashboard Qualité — Semoule SSSE
=================================
Lit le fichier Excel depuis SharePoint (sans Azure AD),
génère un dashboard HTML interactif avec filtres Année / Mois / Date,
et le redépose sur SharePoint pour affichage dans Power Apps.

Prérequis :
    pip install office365-rest-python-client pandas openpyxl plotly jinja2 python-dotenv

Fichier .env à créer dans le même dossier :
    SP_URL      = https://roseblanchetn.sharepoint.com/sites/SDAHSESTPA
    SP_USER     = ton.email@roseblanchetn.com
    SP_PASSWORD = ton_mot_de_passe
    EXCEL_PATH  = Documents Partages/2025 CQ SBOULA FOCQT01_02.xlsx
    SHEET_NAME  = Semoule SSSE
"""

import io
import os
import json
from datetime import datetime
from pathlib import Path

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

from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

# ============================================================
#  CONFIG
# ============================================================

SP_URL      = os.getenv("SP_URL",      "https://roseblanchetn.sharepoint.com/sites/SDAHSESTPA")
SP_USER     = os.getenv("SP_USER",     "ton.email@roseblanchetn.com")
SP_PASSWORD = os.getenv("SP_PASSWORD", "ton_mot_de_passe")
EXCEL_PATH  = os.getenv("EXCEL_PATH",  "Documents Partages/2025 CQ SBOULA FOCQT01_02.xlsx")
SHEET_NAME  = os.getenv("SHEET_NAME",  "Semoule SSSE")
OUTPUT_HTML = "dashboard_ssse.html"

# Colonnes du fichier SSSE (ne pas modifier sauf si la structure change)
COL_DATE    = "Date"
COL_LOT     = "N°lot"
COL_ETAPE   = "Etape"
COL_PROBLEME = "Probléme"
COL_NOTIF   = "Notif"
COL_FLUX    = "Flux_Statut"
COL_ECHANT  = "N° de l'échantillon"

# ============================================================
#  ETAPE 1 — Connexion SharePoint (sans Azure AD)
# ============================================================

def connect_sharepoint() -> ClientContext:
    print(f"  Site : {SP_URL}")
    creds = UserCredential(SP_USER, SP_PASSWORD)
    ctx   = ClientContext(SP_URL).with_credentials(creds)
    web   = ctx.web
    ctx.load(web)
    ctx.execute_query()
    print(f"  Connecté : {web.title}")
    return ctx


# ============================================================
#  ETAPE 2 — Lecture Excel depuis SharePoint
# ============================================================

def read_excel_sharepoint(ctx: ClientContext) -> pd.DataFrame:
    print(f"  Fichier : {EXCEL_PATH}  |  Feuille : {SHEET_NAME}")
    buf = io.BytesIO()
    server_path = f"{ctx.web.server_relative_url}/{EXCEL_PATH}"
    ctx.web.get_file_by_server_relative_url(server_path)\
        .download(buf).execute_query()
    buf.seek(0)
    df = pd.read_excel(buf, sheet_name=SHEET_NAME, header=0)
    print(f"  {len(df)} lignes lues")
    return df


# ============================================================
#  ETAPE 3 — Nettoyage et préparation des données SSSE
# ============================================================

def prepare_data(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Retourne (df_all, df_anomalies).
    Une anomalie = ligne dont la colonne 'Probléme' est renseignée.
    """
    # Nettoyage général
    df[COL_DATE]  = pd.to_datetime(df[COL_DATE], errors="coerce")
    df[COL_ETAPE] = df[COL_ETAPE].astype(str).str.strip()
    df[COL_ETAPE] = df[COL_ETAPE].str.replace(r'\s+', ' ', regex=True)

    # Normaliser "production" → "Production"
    df[COL_ETAPE] = df[COL_ETAPE].str.title()

    # Colonnes temporelles
    df["Année"]     = df[COL_DATE].dt.year.astype("Int64").astype(str)
    df["Mois_num"]  = df[COL_DATE].dt.month.astype("Int64")
    df["Jour"]      = df[COL_DATE].dt.date.astype(str)

    # Filtrer les anomalies (Probléme renseigné)
    df_anom = df[df[COL_PROBLEME].notna()].copy()
    df_anom[COL_PROBLEME] = df_anom[COL_PROBLEME].astype(str).str.strip()

    print(f"  Total analyses  : {len(df)}")
    print(f"  Total anomalies : {len(df_anom)}")
    return df, df_anom


# ============================================================
#  ETAPE 4 — Sérialisation des données pour le dashboard JS
# ============================================================

MOIS_FR = {
    1:"Janvier", 2:"Février", 3:"Mars", 4:"Avril",
    5:"Mai", 6:"Juin", 7:"Juillet", 8:"Août",
    9:"Septembre", 10:"Octobre", 11:"Novembre", 12:"Décembre"
}

def serialize(df_all: pd.DataFrame, df_anom: pd.DataFrame) -> dict:
    rows = []
    for _, r in df_anom.iterrows():
        rows.append({
            "date"     : str(r["Jour"]),
            "annee"    : str(r["Année"]),
            "mois_num" : int(r["Mois_num"]) if pd.notna(r["Mois_num"]) else 0,
            "mois_nom" : MOIS_FR.get(int(r["Mois_num"]), "") if pd.notna(r["Mois_num"]) else "",
            "probleme" : str(r[COL_PROBLEME]),
            "etape"    : str(r[COL_ETAPE]),
            "lot"      : str(r[COL_LOT]) if pd.notna(r[COL_LOT]) else "",
            "echant"   : str(r[COL_ECHANT]) if pd.notna(r[COL_ECHANT]) else "",
            "notif"    : str(r[COL_NOTIF]) if pd.notna(r[COL_NOTIF]) else "Non",
            "flux"     : str(r[COL_FLUX])  if pd.notna(r[COL_FLUX])  else "",
        })
    return {
        "total_analyses" : len(df_all),
        "total_anomalies": len(df_anom),
        "generated_at"   : datetime.now().strftime("%d/%m/%Y %H:%M"),
        "data"           : rows
    }


# ============================================================
#  ETAPE 5 — Template HTML du dashboard
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

/* Header */
.header{background:linear-gradient(135deg,#0d2137 0%,#1a4a72 100%);border:1px solid #1e5a8a;
  border-radius:12px;padding:14px 20px;margin-bottom:14px;
  display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px}
.header h1{font-size:16px;font-weight:600;color:#fff}
.header .meta{font-size:11px;color:#7a9ab8;margin-top:3px}
.badge{background:#0d2a42;border:1px solid #2e6a9a;color:#7bc8f0;
  font-size:11px;padding:4px 12px;border-radius:20px;white-space:nowrap}

/* Filtres */
.filters{background:#0d1e2e;border:1px solid #1a3a52;border-radius:10px;
  padding:12px 16px;margin-bottom:14px;display:flex;flex-wrap:wrap;gap:10px;align-items:center}
.f-group{display:flex;align-items:center;gap:6px}
.f-group label{font-size:11px;color:#5a8ab8;white-space:nowrap}
.filters select{background:#0f2840;border:1px solid #1e5a8a;color:#c8dff0;
  border-radius:6px;padding:5px 10px;font-size:12px;cursor:pointer;outline:none}
.filters select:focus{border-color:#2e8adf}
.btn-reset{background:#1a3a52;border:1px solid #2e6a9a;color:#7bc8f0;
  border-radius:6px;padding:5px 14px;font-size:12px;cursor:pointer;transition:background .15s}
.btn-reset:hover{background:#1e4a6a}
.filter-info{font-size:11px;color:#3a6a9a;margin-left:auto}

/* KPIs */
.kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(155px,1fr));gap:12px;margin-bottom:14px}
.kpi{background:#0d1e2e;border:1px solid #1a3a52;border-radius:10px;
  padding:14px 18px;position:relative;overflow:hidden}
.kpi::before{content:'';position:absolute;left:0;top:0;bottom:0;width:3px;border-radius:3px 0 0 3px}
.kpi.c-blue::before{background:#2e8adf}
.kpi.c-red::before{background:#e24b4a}
.kpi.c-amber::before{background:#f0a030}
.kpi.c-green::before{background:#27ae8f}
.kpi.c-purple::before{background:#8e44ad}
.kpi .lbl{font-size:10px;color:#5a7a9a;text-transform:uppercase;letter-spacing:.5px;margin-bottom:5px}
.kpi .val{font-size:26px;font-weight:700;color:#fff;line-height:1}
.kpi .sub{font-size:11px;color:#3a6a9a;margin-top:5px}

/* Grille graphiques */
.grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px}
.card{background:#0d1e2e;border:1px solid #1a3a52;border-radius:10px;padding:14px;overflow:hidden}
.card.wide{grid-column:1/-1}
.card-title{font-size:11px;color:#5a8ab8;text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px;font-weight:500}

/* Tableau */
.table-wrap{overflow-x:auto;max-height:360px;overflow-y:auto}
.table-wrap::-webkit-scrollbar{width:4px;height:4px}
.table-wrap::-webkit-scrollbar-track{background:#0a1520}
.table-wrap::-webkit-scrollbar-thumb{background:#1a4a72;border-radius:2px}
table{width:100%;border-collapse:collapse;font-size:12px}
thead{position:sticky;top:0;z-index:1}
thead th{background:#0a1724;color:#5a8ab8;text-align:left;
  padding:9px 12px;font-weight:500;border-bottom:1px solid #1a3a52;white-space:nowrap}
tbody tr{border-bottom:1px solid #0d1e2e;transition:background .1s}
tbody tr:hover{background:#0f2840}
tbody td{padding:8px 12px;color:#c8dff0;vertical-align:middle}
.no-data{text-align:center;padding:40px;color:#3a5a7a;font-size:13px}

/* Pills */
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

<!-- HEADER -->
<div class="header">
  <div>
    <h1>Dashboard Qualité — Semoule SSSE</h1>
    <div class="meta" id="meta-info">Chargement...</div>
  </div>
  <span class="badge">T.P.A · SBOULA</span>
</div>

<!-- FILTRES -->
<div class="filters">
  <div class="f-group"><label>Année</label>
    <select id="f-annee"><option value="">Toutes</option></select></div>
  <div class="f-group"><label>Mois</label>
    <select id="f-mois"><option value="">Tous</option></select></div>
  <div class="f-group"><label>Date</label>
    <select id="f-date"><option value="">Toutes</option></select></div>
  <div class="f-group"><label>Problème</label>
    <select id="f-prob"><option value="">Tous</option></select></div>
  <div class="f-group"><label>Étape</label>
    <select id="f-etape"><option value="">Toutes</option></select></div>
  <button class="btn-reset" onclick="resetFilters()">↺ Réinitialiser</button>
  <span class="filter-info" id="filter-info"></span>
</div>

<!-- KPIs -->
<div class="kpi-row">
  <div class="kpi c-blue">
    <div class="lbl">Total analyses</div>
    <div class="val" id="k-total">—</div>
    <div class="sub">Feuille Semoule SSSE</div>
  </div>
  <div class="kpi c-red">
    <div class="lbl">Anomalies détectées</div>
    <div class="val" id="k-anom">—</div>
    <div class="sub" id="k-anom-sub">—</div>
  </div>
  <div class="kpi c-amber">
    <div class="lbl">Taux anomalies</div>
    <div class="val" id="k-taux">—</div>
    <div class="sub">sur total analyses</div>
  </div>
  <div class="kpi c-green">
    <div class="lbl">Notifiées (Oui)</div>
    <div class="val" id="k-notif">—</div>
    <div class="sub" id="k-notif-sub">—</div>
  </div>
  <div class="kpi c-purple">
    <div class="lbl">Problème dominant</div>
    <div class="val" id="k-top" style="font-size:13px;margin-top:6px;line-height:1.4">—</div>
  </div>
</div>

<!-- GRAPHIQUES -->
<div class="grid">
  <div class="card">
    <div class="card-title">Anomalies par type de problème</div>
    <div id="ch-type"></div>
  </div>
  <div class="card">
    <div class="card-title">Répartition par étape de production</div>
    <div id="ch-etape"></div>
  </div>
  <div class="card wide">
    <div class="card-title">Évolution mensuelle des anomalies</div>
    <div id="ch-trend"></div>
  </div>
  <div class="card wide">
    <div class="card-title">Détail des anomalies
      <span id="tbl-count" style="color:#3a6a9a;font-weight:400;margin-left:8px"></span>
    </div>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Date</th>
            <th>N° Lot</th>
            <th>N° Échantillon</th>
            <th>Étape</th>
            <th>Problème</th>
            <th>Notifié</th>
            <th>Flux</th>
          </tr>
        </thead>
        <tbody id="tbl-body"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- DONNEES INJECTEES PAR PYTHON -->
<script>
const _DATA = {{ DATA_JSON }};
const TOTAL_ANALYSES = _DATA.total_analyses;
const GENERATED_AT   = _DATA.generated_at;
let   ALL_DATA       = _DATA.data;

// Palette couleurs par problème
const PROB_COLORS = {
  'PIQUAGE BRUN'       :'#c0392b',
  'PIQUAGE NOIR'       :'#8e44ad',
  'GRANULOMETRIE'      :'#e67e22',
  'COULEUR'            :'#2980b9',
  'TENEUR EN EAU ELEVEE':'#1abc9c',
  'TENEUR EN EAU FAIBLE':'#16a085',
  'MELANGE PRODUITS'   :'#f39c12',
  'CHARANÇONS'         :'#d35400',
  'RHEOLOGIE'          :'#7f8c8d'
};

// Classe pill par problème
const PROB_PILL = {
  'PIQUAGE BRUN':'p-red','PIQUAGE NOIR':'p-purple','GRANULOMETRIE':'p-amber',
  'COULEUR':'p-blue','TENEUR EN EAU ELEVEE':'p-green','TENEUR EN EAU FAIBLE':'p-green',
  'MELANGE PRODUITS':'p-amber','CHARANÇONS':'p-red','RHEOLOGIE':'p-gray'
};

// Noms des mois français
const MOIS_FR = ['','Janvier','Février','Mars','Avril','Mai','Juin',
                 'Juillet','Août','Septembre','Octobre','Novembre','Décembre'];

// ---- Remplir les filtres ----
function unique(key){ return [...new Set(ALL_DATA.map(r=>r[key]))].filter(Boolean); }

function fillSelect(id, values, labelFn){
  const s = document.getElementById(id);
  const prev = s.value;
  while(s.options.length > 1) s.remove(1);
  values.forEach(v => {
    const o = new Option(labelFn ? labelFn(v) : v, v);
    s.appendChild(o);
  });
  if(prev) s.value = prev;
}

function populateFilters(){
  fillSelect('f-annee', unique('annee').sort());
  fillSelect('f-mois',  unique('mois_num').sort((a,b)=>+a-+b), v => MOIS_FR[+v]||v);
  fillSelect('f-date',  unique('date').sort().reverse());
  fillSelect('f-prob',  unique('probleme').sort());
  fillSelect('f-etape', unique('etape').sort());
}

// ---- Filtrage ----
function getFiltered(){
  const a  = document.getElementById('f-annee').value;
  const m  = document.getElementById('f-mois').value;
  const d  = document.getElementById('f-date').value;
  const p  = document.getElementById('f-prob').value;
  const e  = document.getElementById('f-etape').value;
  return ALL_DATA.filter(r =>
    (!a || r.annee    === a)  &&
    (!m || r.mois_num === +m) &&
    (!d || r.date     === d)  &&
    (!p || r.probleme === p)  &&
    (!e || r.etape    === e)
  );
}

// ---- Layout Plotly commun ----
const LAY_BASE = {
  paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
  font:{color:'#c8dff0', size:11},
  margin:{l:10,r:10,t:10,b:10}
};
const CFG = {displayModeBar:false, responsive:true};

function noDataLayout(msg='Aucune donnée'){
  return {...LAY_BASE, height:240,
    annotations:[{text:msg,x:.5,y:.5,xref:'paper',yref:'paper',
      showarrow:false,font:{color:'#3a5a7a',size:13}}]};
}

// ---- Rendu principal ----
function render(){
  const fd = getFiltered();
  const n  = fd.length;

  // Info filtre
  const isFiltered = n < ALL_DATA.length;
  document.getElementById('filter-info').textContent =
    isFiltered ? `Filtre actif : ${n} anomalie(s) affichée(s) sur ${ALL_DATA.length}` : '';

  // KPIs
  document.getElementById('k-total').textContent = TOTAL_ANALYSES.toLocaleString('fr');
  document.getElementById('k-anom').textContent  = n.toLocaleString('fr');
  document.getElementById('k-anom-sub').textContent =
    isFiltered ? `(total : ${ALL_DATA.length})` : 'sur la période complète';

  const taux = TOTAL_ANALYSES > 0 ? (n / TOTAL_ANALYSES * 100).toFixed(1) : '0';
  document.getElementById('k-taux').textContent = taux + '%';

  const notifOui = fd.filter(r => r.notif === 'Oui').length;
  document.getElementById('k-notif').textContent = notifOui.toLocaleString('fr');
  document.getElementById('k-notif-sub').textContent =
    n > 0 ? `${(notifOui/n*100).toFixed(0)}% des anomalies` : '—';

  // Top problème
  const cntP = {};
  fd.forEach(r => cntP[r.probleme] = (cntP[r.probleme]||0) + 1);
  const topP = Object.entries(cntP).sort((a,b) => b[1]-a[1]);
  document.getElementById('k-top').textContent =
    topP.length ? `${topP[0][0]} (${topP[0][1]})` : '—';

  // Méta
  document.getElementById('meta-info').textContent =
    `Généré le ${GENERATED_AT} · Source : SharePoint Excel · Feuille : Semoule SSSE`;

  // ---- Graphique 1 : Barres horizontales par problème ----
  if(!topP.length){
    Plotly.newPlot('ch-type', [], noDataLayout(), CFG);
  } else {
    const sorted = topP.slice(0, 9).reverse();
    Plotly.newPlot('ch-type', [{
      type:'bar', orientation:'h',
      y: sorted.map(x => x[0]),
      x: sorted.map(x => x[1]),
      marker:{ color: sorted.map(x => PROB_COLORS[x[0]] || '#2e8adf') },
      text: sorted.map(x => x[1]),
      textposition:'outside',
      hovertemplate:'%{y} : %{x} cas<extra></extra>'
    }], {
      ...LAY_BASE,
      height:260,
      margin:{l:180,r:50,t:10,b:20},
      xaxis:{gridcolor:'#1a3a52', tickfont:{size:10}},
      yaxis:{gridcolor:'rgba(0,0,0,0)', tickfont:{size:11}}
    }, CFG);
  }

  // ---- Graphique 2 : Camembert par étape ----
  const cntE = {};
  fd.forEach(r => cntE[r.etape] = (cntE[r.etape]||0) + 1);
  const topE = Object.entries(cntE).sort((a,b) => b[1]-a[1]);
  if(!topE.length){
    Plotly.newPlot('ch-etape', [], noDataLayout(), CFG);
  } else {
    Plotly.newPlot('ch-etape', [{
      type:'pie',
      labels: topE.map(x => x[0]),
      values: topE.map(x => x[1]),
      hole: 0.42,
      marker:{ colors:['#2e8adf','#27ae8f','#e67e22','#8e44ad','#c0392b','#f39c12','#1abc9c'] },
      textinfo:'label+percent',
      textfont:{size:10, color:'#c8dff0'},
      hovertemplate:'%{label}: %{value} (%{percent})<extra></extra>'
    }], {
      ...LAY_BASE,
      height:260,
      margin:{l:10,r:10,t:10,b:10},
      showlegend:true,
      legend:{font:{size:10,color:'#7a9ab8'}, bgcolor:'rgba(0,0,0,0)',
              orientation:'v', x:1, y:.5}
    }, CFG);
  }

  // ---- Graphique 3 : Tendance mensuelle ----
  const cntM = {};
  fd.forEach(r => {
    const k = r.annee + '-' + String(r.mois_num).padStart(2,'0');
    cntM[k] = (cntM[k]||0) + 1;
  });
  const mKeys = Object.keys(cntM).sort();
  const mLabels = mKeys.map(k => {
    const [y,m] = k.split('-');
    return MOIS_FR[+m].slice(0,3) + ' ' + y;
  });
  if(!mKeys.length){
    Plotly.newPlot('ch-trend', [], noDataLayout(), CFG);
  } else {
    Plotly.newPlot('ch-trend', [{
      type:'scatter', mode:'lines+markers',
      x: mLabels,
      y: mKeys.map(k => cntM[k]),
      line:{color:'#2e8adf', width:2, shape:'spline'},
      marker:{color:'#2e8adf', size:7, line:{color:'#0f2840',width:2}},
      fill:'tozeroy', fillcolor:'rgba(46,138,223,0.1)',
      hovertemplate:'%{x}: <b>%{y}</b> anomalie(s)<extra></extra>'
    }], {
      ...LAY_BASE,
      height:220,
      margin:{l:40,r:20,t:10,b:70},
      xaxis:{
        gridcolor:'#1a3a52', tickfont:{size:10}, tickangle:-45,
        showgrid:true, zeroline:false
      },
      yaxis:{
        gridcolor:'#1a3a52', tickfont:{size:10},
        showgrid:true, zeroline:false
      }
    }, CFG);
  }

  // ---- Tableau détail (100 premières lignes filtrées) ----
  const tbody = document.getElementById('tbl-body');
  const display = fd.slice(0, 100);
  document.getElementById('tbl-count').textContent =
    fd.length > 100 ? `(100 premières sur ${fd.length})` :
    fd.length > 0   ? `(${fd.length} ligne(s))` : '';

  if(!display.length){
    tbody.innerHTML = '<tr><td colspan="7" class="no-data">Aucune anomalie pour cette sélection</td></tr>';
    return;
  }

  tbody.innerHTML = display.map(r => {
    const pillC  = PROB_PILL[r.probleme] || 'p-gray';
    const notifC = r.notif === 'Oui' ? '#50c090' : '#e07070';
    return `<tr>
      <td style="color:#7a9ab8">${r.date}</td>
      <td>${r.lot}</td>
      <td style="font-size:11px;color:#5a8ab8">${r.echant}</td>
      <td><span class="pill p-blue">${r.etape}</span></td>
      <td><span class="pill ${pillC}">${r.probleme}</span></td>
      <td style="color:${notifC};font-weight:500">${r.notif}</td>
      <td style="color:#3a6a8a;font-size:11px">${r.flux}</td>
    </tr>`;
  }).join('');
}

function resetFilters(){
  ['f-annee','f-mois','f-date','f-prob','f-etape'].forEach(id =>
    document.getElementById(id).value = '');
  render();
}

// Écouter les filtres
['f-annee','f-mois','f-date','f-prob','f-etape'].forEach(id =>
  document.getElementById(id).addEventListener('change', render));

// Init
populateFilters();
render();
</script>
</body>
</html>"""


def generate_html(payload: dict) -> str:
    data_json = json.dumps(payload, ensure_ascii=False)
    return DASHBOARD_HTML.replace("{{ DATA_JSON }}", data_json)


# ============================================================
#  ETAPE 6 — Upload HTML sur SharePoint
# ============================================================

def upload_html(ctx: ClientContext, html: str) -> str:
    folder_path   = "/".join(EXCEL_PATH.split("/")[:-1])
    server_folder = f"{ctx.web.server_relative_url}/{folder_path}"
    folder  = ctx.web.get_folder_by_server_relative_url(server_folder)
    upload  = folder.upload_file(OUTPUT_HTML, html.encode("utf-8")).execute_query()
    rel_url = upload.serverRelativeUrl
    full_url = SP_URL.split("/sites/")[0] + rel_url
    print(f"  URL : {full_url}")
    return full_url


# ============================================================
#  MAIN
# ============================================================

if __name__ == "__main__":
    print("\n" + "="*50)
    print("  Dashboard Qualité SSSE — Génération")
    print("="*50 + "\n")

    print("[1/5] Connexion SharePoint (SSO entreprise)...")
    ctx = connect_sharepoint()

    print("\n[2/5] Lecture fichier Excel...")
    df_raw = read_excel_sharepoint(ctx)

    print("\n[3/5] Préparation des données SSSE...")
    df_all, df_anom = prepare_data(df_raw)

    print("\n[4/5] Sérialisation + génération HTML...")
    payload = serialize(df_all, df_anom)
    html    = generate_html(payload)
    print(f"  Taille HTML : {len(html)//1024} Ko")

    print("\n[5/5] Upload sur SharePoint...")
    url = upload_html(ctx, html)

    print("\n" + "="*50)
    print("  TERMINÉ")
    print("="*50)
    print(f"\n  Anomalies  : {payload['total_anomalies']}")
    print(f"  Analyses   : {payload['total_analyses']}")
    print(f"  Généré le  : {payload['generated_at']}")
    print(f"\n  URL Power Apps (Web viewer) :")
    print(f"  \"{url}\"")
    print()
