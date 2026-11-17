"""
SUIVI EN TEMPS RÉEL – ENQUÊTE SUR L'IDENTIFICATION DES POINTS DE VENTE (INHPC)
Institut National de la Statistique (INS) – Direction des Statistiques Économiques
Financement : projet HISWACA (Banque mondiale)
Application réalisée par OGNIMBA Sadri

Lancer en local :  streamlit run app.py
Paramètres      :  .streamlit/secrets.toml  (voir secrets.toml.exemple)
"""
from __future__ import annotations

import base64
import hmac
import html
import io
import math
import numbers
import os
from datetime import timedelta

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from kobo_suivi import (FUSEAU, POIDS_SCORE, STATUTS, KoboErreur, aplatir, construire_suivi,
                        croise_types_localites, diagnostic, dictionnaire, libelle_court, lire_kobo,
                        performance_agents, performance_localites, tableau_types,
                        validation_par_agent, validation_par_localite)
from rapport import generer_rapport

st.set_page_config(page_title="Suivi collecte INHPC – INS", page_icon="📍", layout="wide")
DOSSIER = os.path.dirname(os.path.abspath(__file__))


# ---------------------------------------------------------------------
# Paramètres (secrets Streamlit, sinon variables d'environnement)
# ---------------------------------------------------------------------
def param(nom, defaut=None):
    try:
        if nom in st.secrets:
            return st.secrets[nom]
    except Exception:
        pass
    return os.environ.get(nom, defaut)


JETON = param("KOBO_TOKEN", "")
SERVEUR = param("KOBO_SERVEUR", "https://kf.kobotoolbox.org")
ASSET_UID = param("ASSET_UID", "abVbPCHqtiYSwKP8YAZ4QP")
MINUTES = max(5, int(param("ACTUALISATION_MINUTES", 15)))
MOT_DE_PASSE = str(param("MOT_DE_PASSE", "") or "")
NB_AGENTS = int(param("NB_AGENTS_PREVUS", 38) or 0)
OBJECTIF_TOTAL = int(param("OBJECTIF_TOTAL", 0) or 0)
OBJECTIFS_LOC = {str(k): int(v) for k, v in
                 dict(param("OBJECTIFS_LOCALITE", None) or param("OBJECTIFS_DEPARTEMENT", {}) or {}).items() if v}
LOC_DEFAUT = str(param("LOCALITE_PAR_DEFAUT", "") or param("DEPARTEMENT_PAR_DEFAUT", "") or "")
COLONNES_FORCEES = dict(param("COLONNES", {}) or {})
SEUILS = {k: float(param(k)) for k in ("DUREE_MIN_MINUTES", "DUREE_MAX_MINUTES", "PRECISION_GPS_MAX_M",
                                       "DISTANCE_DOUBLON_M", "DELAI_VALIDATION_H") if param(k) is not None}

BLEU, BLEU_FONCE, ROUGE, GRIS = "#2a78d6", "#16324f", "#d03b3b", "#7a7974"
COULEURS_STATUT = {"Approuvé": "#0ca30c", "Non approuvé": "#d03b3b", "En attente": "#fab219", "Non vérifié": "#b8b7ae"}
CONFIG_GRAPH = {"displayModeBar": False, "locale": "fr"}


def image_b64(nom):
    chemin = os.path.join(DOSSIER, "assets", nom)
    if not os.path.exists(chemin):
        return ""
    with open(chemin, "rb") as fh:
        return "data:image/png;base64," + base64.b64encode(fh.read()).decode()


LOGO_INS, LOGO_HISWACA = image_b64("logo_ins.png"), image_b64("logo_hiswaca.png")

# ---------------------------------------------------------------------
# Style
# ---------------------------------------------------------------------
st.markdown(f"""
<style>
  .block-container {{ padding-top: 1rem; padding-bottom: 2rem; max-width: 1500px; }}
  .bandeau {{ display: flex; align-items: center; gap: 18px; background: #fff; border: 1px solid rgba(11,11,11,.08);
     border-top: 5px solid {BLEU_FONCE}; border-radius: 12px; padding: 12px 22px; margin-bottom: 14px;
     box-shadow: 0 1px 3px rgba(0,0,0,.05); }}
  .bandeau img.ins {{ height: 78px; }}
  .bandeau img.hiswaca {{ height: 62px; }}
  .bandeau .centre {{ flex: 1; text-align: center; }}
  .bandeau .sur {{ font-size: .78rem; letter-spacing: .08em; text-transform: uppercase; color: {BLEU}; font-weight: 600; }}
  .bandeau h1 {{ font-size: 1.55rem; line-height: 1.25; margin: 2px 0 4px; padding: 0; color: {BLEU_FONCE}; font-weight: 700; }}
  .bandeau .sous {{ font-size: .85rem; color: #52514e; }}
  .kpis {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 12px; margin: 6px 0 4px; }}
  .kpi {{ background: #fff; border: 1px solid rgba(11,11,11,.08); border-left: 4px solid {BLEU};
          border-radius: 10px; padding: 12px 14px; }}
  .kpi.alerte {{ border-left-color: {ROUGE}; }}
  .kpi.ok {{ border-left-color: #0ca30c; }}
  .kpi .lib {{ font-size: .8rem; color: #52514e; line-height: 1.2; }}
  .kpi .val {{ font-size: 1.9rem; font-weight: 700; color: #0b0b0b; line-height: 1.15; margin-top: 4px;
               font-variant-numeric: tabular-nums; white-space: nowrap; }}
  .kpi .sub {{ font-size: .78rem; color: #7a7974; margin-top: 2px; }}
  .titre-bloc {{ font-weight: 650; font-size: 1.02rem; color: {BLEU_FONCE}; margin: 6px 0 2px; }}
  .note-bloc {{ font-size: .8rem; color: #7a7974; margin-bottom: 6px; }}
  table.tb {{ width: 100%; border-collapse: collapse; font-size: .86rem; background: #fff; }}
  table.tb th {{ background: {BLEU_FONCE}; color: #fff; font-weight: 600; padding: 8px 10px; text-align: center;
                 white-space: nowrap; }}
  table.tb th:first-child, table.tb td:first-child {{ text-align: left; }}
  table.tb td {{ padding: 6px 10px; border-bottom: 1px solid #e1e0d9; text-align: center;
                 font-variant-numeric: tabular-nums; }}
  table.tb tr:nth-child(even) td {{ background: #f8f7f4; }}
  table.tb tr.total td {{ background: #e3ecf7 !important; font-weight: 700; border-top: 2px solid {BLEU_FONCE}; }}
  table.tb td.zero {{ color: #b8b7ae; }}
  .defile {{ overflow-x: auto; border: 1px solid #e1e0d9; border-radius: 8px; }}
  .credit {{ background: {BLEU_FONCE}; color: #fff; border-radius: 10px; padding: 12px 14px; font-size: .82rem; }}
  .credit b {{ font-size: .95rem; }}
  .pied {{ text-align: center; color: #7a7974; font-size: .8rem; margin-top: 26px; padding-top: 12px;
           border-top: 1px solid #e1e0d9; }}
  div[data-testid="stTabs"] button p {{ font-size: .95rem; font-weight: 550; }}
</style>""", unsafe_allow_html=True)


def bandeau():
    ins = f'<img class="ins" src="{LOGO_INS}" alt="INS">' if LOGO_INS else ""
    hisw = f'<img class="hiswaca" src="{LOGO_HISWACA}" alt="HISWACA">' if LOGO_HISWACA else ""
    st.markdown(f"""<div class="bandeau">{ins}<div class="centre">
      <div class="sur">Institut National de la Statistique · Direction des Statistiques Économiques</div>
      <h1>Suivi de la collecte – Enquête sur l'identification des points de vente (INHPC)</h1>
      <div class="sous">Données KoboToolbox en temps réel · actualisation automatique toutes les {MINUTES} min</div>
      </div>{hisw}</div>""", unsafe_allow_html=True)


# ---------------------------------------------------------------------
# Accès protégé (facultatif)
# ---------------------------------------------------------------------
def acces_autorise() -> bool:
    if not MOT_DE_PASSE or st.session_state.get("ok_acces"):
        return True
    bandeau()
    _, milieu, _ = st.columns([1, 1, 1])
    with milieu:
        saisie = st.text_input("Code d'accès", type="password")
        if saisie:
            if hmac.compare_digest(saisie, MOT_DE_PASSE):
                st.session_state["ok_acces"] = True
                st.rerun()
            st.error("Code d'accès incorrect.")
    return False


if not acces_autorise():
    st.stop()
if not JETON:
    st.error("Le jeton API Kobo n'est pas configuré : ajoutez KOBO_TOKEN dans les secrets de l'application.")
    st.stop()


# ---------------------------------------------------------------------
# Données (cache partagé entre tous les visiteurs)
# ---------------------------------------------------------------------
@st.cache_data(ttl=MINUTES * 60 - 30, show_spinner="Chargement des données KoboToolbox…")
def charger(serveur, uid, jeton, seuils_t, colonnes_t):
    asset, resultats = lire_kobo(serveur, uid, jeton)
    dico = dictionnaire(asset.get("content", {}) or {})
    brut, types = aplatir(resultats, dico)
    fiches, C = construire_suivi(brut, types, dict(seuils_t), dict(colonnes_t), dico)
    return brut, fiches, C, pd.Timestamp.now(tz=FUSEAU).tz_localize(None), asset.get("name", "")


# ---------------------------------------------------------------------
# Outils d'affichage
# ---------------------------------------------------------------------
def nf(x, dec=0):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "–"
    return f"{x:,.{dec}f}".replace(",", " ").replace(".", ",")


def cartes_kpi(cartes):
    blocs = "".join(
        f'<div class="kpi {c.get("classe", "")}"><div class="lib">{html.escape(c["lib"])}</div>'
        f'<div class="val">{c["val"]}</div><div class="sub">{c.get("sub", "")}</div></div>' for c in cartes)
    st.markdown(f'<div class="kpis">{blocs}</div>', unsafe_allow_html=True)


def titre(txt, note=""):
    st.markdown(f'<div class="titre-bloc">{txt}</div>' + (f'<div class="note-bloc">{note}</div>' if note else ""),
                unsafe_allow_html=True)


def table_html(df: pd.DataFrame, total=False, pct_cols=(), tirets=False):
    entetes = "".join(f"<th>{html.escape(str(c))}</th>" for c in df.columns)
    lignes = []
    for i, (_, r) in enumerate(df.iterrows()):
        est_total = total and i == len(df) - 1
        cells = []
        for c in df.columns:
            v = r[c]
            classe = ""
            if isinstance(v, numbers.Number) and not isinstance(v, bool) and not pd.isna(v):
                if tirets and v == 0:
                    txt, classe = "–", ' class="zero"'
                elif c in pct_cols:
                    txt = nf(v, 1)
                else:
                    txt = nf(v, 0) if float(v).is_integer() else nf(v, 1)
            else:
                txt = html.escape(str(v))
            cells.append(f"<td{classe}>{txt}</td>")
        lignes.append(f'<tr{" class=total" if est_total else ""}>' + "".join(cells) + "</tr>")
    st.markdown(f'<div class="defile"><table class="tb"><thead><tr>{entetes}</tr></thead><tbody>'
                + "".join(lignes) + "</tbody></table></div>", unsafe_allow_html=True)


def mise_en_forme(fig, hauteur=340):
    fig.update_layout(template="plotly_white", height=hauteur, margin=dict(l=10, r=20, t=10, b=10),
                      font=dict(family="Source Sans Pro, Segoe UI, Roboto, Arial, sans-serif", size=13, color="#0b0b0b"),
                      barcornerradius=4, bargap=0.3, separators=", ", showlegend=False,
                      hoverlabel=dict(bgcolor="#0b0b0b", font_color="#fff", font_size=13),
                      paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
    fig.update_xaxes(showgrid=False, linecolor="#c3c2b7", title=None, automargin=True)
    fig.update_yaxes(gridcolor="#e1e0d9", zeroline=False, title=None, automargin=True)
    return fig


def barres_h(serie: pd.Series, libelle="Fiches", couleur=BLEU, suffixe="", maxi=None):
    d = serie.sort_values().reset_index()
    d.columns = ["nom", "n"]
    d["etiquette"] = d["n"].map(lambda v: nf(v, 0 if float(v).is_integer() else 1) + suffixe)
    fig = px.bar(d, x="n", y="nom", orientation="h", text="etiquette", color_discrete_sequence=[couleur])
    fig.update_traces(textposition="outside", cliponaxis=False,
                      hovertemplate="<b>%{y}</b><br>" + libelle + " : %{text}<extra></extra>")
    mise_en_forme(fig, max(220, 34 * len(d) + 40))
    fig.update_xaxes(showgrid=True, gridcolor="#e1e0d9", showticklabels=False, range=[0, (maxi or d["n"].max()) * 1.18])
    return fig


def vers_excel(feuilles: dict) -> bytes:
    tampon = io.BytesIO()
    with pd.ExcelWriter(tampon, engine="openpyxl") as w:
        for nom, d in feuilles.items():
            d.to_excel(w, sheet_name=nom[:31], index=False)
    return tampon.getvalue()


def lib_loc(loc, sup):
    return f"{loc} ({libelle_court(sup, 22)})" if sup else loc


def graphique(fig, cle):
    st.plotly_chart(fig, key=cle, config=CONFIG_GRAPH)


# ---------------------------------------------------------------------
# En-tête et barre latérale
# ---------------------------------------------------------------------
bandeau()
with st.sidebar:
    st.markdown('<div class="credit">Application de suivi INHPC<br>réalisée par <b>OGNIMBA Sadri</b><br>'
                '<span style="opacity:.8">Ingénieur Statisticien Économiste · INS – DSE</span></div>',
                unsafe_allow_html=True)
    st.write("")
    if st.button("🔄 Actualiser maintenant", width="stretch"):
        charger.clear()
        st.rerun()
    st.caption(f"Les données sont relues dans Kobo toutes les {MINUTES} minutes ; la page se met à jour seule.")
    st.divider()
    st.markdown("**Seuils des contrôles qualité**")
    st.caption(f"Durée trop courte < {int(SEUILS.get('DUREE_MIN_MINUTES', 5))} min · "
               f"trop longue > {int(SEUILS.get('DUREE_MAX_MINUTES', 150))} min  \n"
               f"GPS non fiable si précision > {int(SEUILS.get('PRECISION_GPS_MAX_M', 30))} m  \n"
               f"Validation attendue sous {int(SEUILS.get('DELAI_VALIDATION_H', 24))} h")
    st.caption("Une alerte n'est pas forcément une erreur : c'est un point à vérifier avec l'agent.")


# ---------------------------------------------------------------------
# Tableau de bord (fragment relancé automatiquement)
# ---------------------------------------------------------------------
@st.fragment(run_every=timedelta(minutes=MINUTES))
def tableau_de_bord():
    try:
        brut, fiches, C, heure_maj, nom_form = charger(SERVEUR, ASSET_UID, JETON, tuple(sorted(SEUILS.items())),
                                                       tuple(sorted(COLONNES_FORCEES.items())))
    except KoboErreur as e:
        st.error(f"Lecture de Kobo impossible : {e}")
        return
    except Exception as e:  # pragma: no cover
        st.error(f"Erreur inattendue pendant la lecture des données : {e}")
        return
    if fiches.empty:
        st.info("Aucune fiche reçue pour l'instant dans Kobo.")
        return

    auj = pd.Timestamp.now(tz=FUSEAU).date()
    hier = auj - timedelta(days=1)
    ordre = C["ordre_types"]

    # ---------------- Filtres ----------------
    c1, c2, c3, c4 = st.columns([1.4, 1.4, 1.6, 1.1])
    locs = fiches["Localite"].value_counts().index.tolist()
    defaut = [l for l in locs if LOC_DEFAUT and l.lower() == LOC_DEFAUT.lower()]
    f_loc = c1.multiselect("Localité", locs, default=defaut, placeholder="Toutes les localités")
    base = fiches[fiches["Localite"].isin(f_loc)] if f_loc else fiches
    f_agent = c2.multiselect("Agent", sorted(base["Agent"].unique()), placeholder="Tous les agents")
    types_obs = [t for t in ordre if t in set(base["Type_PV"])]
    f_type = c3.multiselect("Type de point de vente", types_obs, format_func=libelle_court, placeholder="Tous les types")
    periode = c4.selectbox("Période", ["Toute la collecte", "Aujourd'hui", "Hier", "7 derniers jours", "Personnalisée"])
    if f_agent:
        base = base[base["Agent"].isin(f_agent)]
    if f_type:
        base = base[base["Type_PV"].isin(f_type)]
    du = au = None
    if periode == "Aujourd'hui":
        du = au = auj
    elif periode == "Hier":
        du = au = hier
    elif periode == "7 derniers jours":
        du, au = auj - timedelta(days=6), auj
    elif periode == "Personnalisée":
        debut = base["Jour"].dropna().min() if base["Jour"].notna().any() else auj
        choix = st.date_input("Du … au …", value=(debut, auj), format="DD/MM/YYYY")
        if isinstance(choix, (tuple, list)) and len(choix) == 2:
            du, au = choix
    t = base[(base["Jour"] >= du) & (base["Jour"] <= au)] if du else base
    perimetre = ", ".join(f_loc) if f_loc else "toutes les localités"
    if f_agent:
        perimetre += f" · {len(f_agent)} agent(s)"
    if du:
        perimetre += f" · du {du:%d/%m} au {au:%d/%m/%Y}"

    # ---------------- Indicateurs clés ----------------
    n = len(t)
    n_auj, n_hier = int((base["Jour"] == auj).sum()), int((base["Jour"] == hier).sum())
    ag_tous = base["Agent"].nunique()
    ag_auj = base.loc[base["Jour"] == auj, "Agent"].nunique()
    prevus = NB_AGENTS if not f_loc and not f_agent and NB_AGENTS else ag_tous
    n_al = int((t["Nb_alertes"] > 0).sum())
    n_app = int((t["Statut"] == "Approuvé").sum())
    n_nv24 = int(t["Hors_delai_validation"].sum())
    ecart = n_auj - n_hier
    cartes = [
        {"lib": "Points de vente identifiés", "val": nf(n),
         "sub": (f"Objectif {nf(OBJECTIF_TOTAL)} · {nf(len(base) / OBJECTIF_TOTAL * 100, 1)} % réalisé"
                 if OBJECTIF_TOTAL and not f_loc else ("Sur la période choisie" if du else "Depuis le début de la collecte"))},
        {"lib": "Fiches reçues aujourd'hui", "val": nf(n_auj),
         "sub": f"Hier : {nf(n_hier)} ({'+' if ecart >= 0 else '−'}{nf(abs(ecart))})"},
        {"lib": "Agents actifs aujourd'hui", "val": f"{ag_auj} / {prevus}",
         "sub": f"{ag_tous} agent(s) ont déjà transmis au moins une fiche"},
        {"lib": "Localités couvertes", "val": nf(t["Localite"].nunique()),
         "sub": f"{nf(t['Superviseur'].replace('', pd.NA).nunique())} superviseur(s)"},
        {"lib": "Fiches avec alerte qualité", "val": nf(n_al), "classe": "alerte" if n_al else "ok",
         "sub": f"{nf(n_al / n * 100 if n else 0, 1)} % des fiches"},
        {"lib": "Fiches approuvées", "val": nf(n_app), "classe": "ok",
         "sub": f"{nf(n_app / n * 100 if n else 0, 1)} % · {nf(n_nv24)} non vérifiée(s) > 24 h"},
    ]
    cartes_kpi(cartes)

    perf = performance_localites(t)
    types_tab = tableau_types(t, ordre)
    valid = validation_par_localite(t)

    i1, i2 = st.columns([3, 1])
    derniere = fiches["Date_soumission"].max()
    i1.caption(f"Données lues dans Kobo à **{heure_maj:%H:%M}** · prochaine lecture vers "
               f"**{heure_maj + timedelta(minutes=MINUTES):%H:%M}** · dernière fiche reçue le "
               f"**{derniere:%d/%m/%Y à %H:%M}**")
    try:
        doc = generer_rapport(t, perf, types_tab, valid, {
            "total": n, "auj": n_auj, "agents_auj": ag_auj, "agents_tous": ag_tous,
            "localites": t["Localite"].nunique(), "alertes": n_al, "pct_alertes": n_al / n * 100 if n else 0,
            "approuvees": n_app, "pct_approuvees": n_app / n * 100 if n else 0},
            perimetre, pd.Timestamp.now(tz=FUSEAU).tz_localize(None), prevus if not f_loc else 0)
        i2.download_button("📄 Télécharger le rapport (Word)", doc, width="stretch", type="primary",
                           file_name=f"Rapport_collecte_INHPC_{auj:%Y%m%d}.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except Exception as e:  # pragma: no cover
        i2.warning(f"Rapport indisponible : {e}")

    onglets = st.tabs(["📈 Vue d'ensemble", "🏆 Localités", "✅ Validation", "👥 Agents", "🗺️ Carte",
                       "⚠️ Qualité", "🛒 Produits", "📋 Données", "🔧 Diagnostic"])

    # ================= 1. Vue d'ensemble =================
    with onglets[0]:
        if t.empty:
            st.info("Aucune fiche pour ces filtres.")
        else:
            titre("Nombre de points de vente identifiés par localité / superviseur",
                  "Une soumission = un point de vente identifié")
            g = t.groupby(["Localite", "Superviseur"]).size().reset_index(name="PV").sort_values("PV", ascending=False)
            g["etiq"] = g.apply(lambda r: f"<b>{r['Localite']}</b>" + (
                f"<br><span style='font-size:11px;color:#52514e'>{libelle_court(r['Superviseur'], 22)}</span>"
                if r["Superviseur"] else ""), axis=1)
            g["obj"] = g["Localite"].map(OBJECTIFS_LOC).fillna(0)
            fig = go.Figure(go.Bar(x=g["etiq"], y=g["PV"], text=g["PV"], marker_color=BLEU, textposition="outside",
                                   cliponaxis=False, customdata=g[["Localite", "Superviseur", "obj"]],
                                   hovertemplate="<b>%{customdata[0]}</b><br>Superviseur : %{customdata[1]}"
                                                 "<br>Points de vente : %{y}<extra></extra>"))
            if OBJECTIFS_LOC and not du:
                fig.add_scatter(x=g["etiq"], y=g["obj"].where(g["obj"] > 0), mode="markers", name="Objectif",
                                marker=dict(symbol="line-ew", size=34, line=dict(width=3, color=BLEU_FONCE)),
                                hovertemplate="Objectif : %{y}<extra></extra>")
            mise_en_forme(fig, 400).update_yaxes(range=[0, max(g["PV"].max(), g["obj"].max()) * 1.15])
            graphique(fig, "loc")

            par_jour = t.groupby("Jour").size()
            plage = pd.date_range(par_jour.index.min(), max(par_jour.index.max(), au or par_jour.index.max()))
            par_jour = par_jour.reindex(plage.date, fill_value=0)
            dj = pd.DataFrame({"Jour": pd.to_datetime(par_jour.index), "Fiches": par_jour.values})
            dj["Cumul"] = dj["Fiches"].cumsum()
            a, b = st.columns(2)
            with a:
                titre("Fiches reçues par jour")
                fig = px.bar(dj, x="Jour", y="Fiches", text="Fiches", color_discrete_sequence=[BLEU])
                fig.update_traces(textposition="outside", cliponaxis=False,
                                  hovertemplate="<b>%{x|%d/%m/%Y}</b><br>%{y} fiches<extra></extra>")
                mise_en_forme(fig, 320).update_xaxes(tickformat="%d/%m", dtick=86400000)
                graphique(fig, "jour")
            with b:
                titre("Cumul des points de vente identifiés")
                fig = px.line(dj, x="Jour", y="Cumul", markers=True, color_discrete_sequence=[BLEU])
                fig.update_traces(line_width=2, marker_size=8,
                                  hovertemplate="<b>%{x|%d/%m/%Y}</b><br>Cumul : %{y}<extra></extra>")
                if OBJECTIF_TOTAL and not f_loc and not du:
                    fig.add_hline(y=OBJECTIF_TOTAL, line_dash="dot", line_color=GRIS,
                                  annotation_text=f"Objectif : {OBJECTIF_TOTAL}", annotation_position="top left")
                mise_en_forme(fig, 320).update_xaxes(tickformat="%d/%m", dtick=86400000)
                graphique(fig, "cumul")

            a, b = st.columns([1.05, 1])
            with a:
                titre("Répartition des points de vente par type", "Ordre du questionnaire")
                table_html(types_tab, total=True, pct_cols=("Pourcentage",))
            with b:
                titre("Points de vente par type")
                s = types_tab.iloc[:-1].set_index("Type de point de vente")["Nbre de points de vente"]
                s = s[s > 0]
                s.index = [libelle_court(i, 40) for i in s.index]
                graphique(barres_h(s, "Points de vente"), "type")

    # ================= 2. Localités =================
    with onglets[1]:
        if perf.empty:
            st.info("Aucune fiche pour ces filtres.")
        else:
            titre("Classement des localités – score de performance /100",
                  "Le classement combine six critères ; le nombre de points de vente seul ne suffit pas à juger.")
            s = perf.set_index(pd.Index([lib_loc(a, b) for a, b in zip(perf["Localité"], perf["Superviseur"])]))["Score /100"]
            graphique(barres_h(s, "Score", suffixe=" /100", maxi=100), "score")
            cols = ["Rang", "Localité", "Superviseur", "PV identifiés", "Agents", "PV / agent / jour", "Productivité",
                    "Complétude", "Fiabilité GPS", "Durée conforme", "Sans doublon", "Validation", "Score /100", "Remarque"]
            tab = perf[cols].copy()
            if OBJECTIFS_LOC:
                tab.insert(4, "Objectif", tab["Localité"].map(OBJECTIFS_LOC))
                tab.insert(5, "Taux de réalisation (%)", (tab["PV identifiés"] / tab["Objectif"] * 100).round(1))
            pc = {k: st.column_config.ProgressColumn(k, min_value=0, max_value=100, format="%.0f %%")
                  for k in POIDS_SCORE}
            pc["Score /100"] = st.column_config.ProgressColumn("Score /100", min_value=0, max_value=100, format="%.1f")
            st.dataframe(tab, width="stretch", hide_index=True, column_config=pc)
            with st.expander("Méthode de calcul du score"):
                st.markdown(
                    "Chaque critère est exprimé en pourcentage (100 = parfait), puis pondéré :\n\n"
                    "| Critère | Mesure | Poids |\n|---|---|---|\n"
                    "| Productivité | PV par agent et par jour travaillé, rapporté au niveau des meilleures localités (3e quartile), plafonné à 100 | 20 % |\n"
                    "| Complétude | % de fiches sans information manquante ou invalide (nom, type, catégories, téléphone…) | 20 % |\n"
                    f"| Fiabilité GPS | % de fiches avec un GPS présent, au Congo et précis à ≤ {int(C['parametres']['PRECISION_GPS_MAX_M'])} m | 15 % |\n"
                    f"| Durée conforme | % d'entretiens entre {int(C['parametres']['DUREE_MIN_MINUTES'])} et {int(C['parametres']['DUREE_MAX_MINUTES'])} min | 15 % |\n"
                    "| Sans doublon | % de fiches non suspectes de doublon | 10 % |\n"
                    "| Validation | % des fiches de plus de 24 h traitées par le superviseur | 20 % |\n\n"
                    "La productivité est mesurée **par agent et par jour de travail** pour ne pas avantager les "
                    "localités qui ont plus d'agents. Une localité de moins de 10 fiches est à interpréter avec prudence.")
            st.write("")
            titre("Répartition des types de point de vente par localité", "Ordre du questionnaire · « – » = aucun")
            ct = croise_types_localites(t, ordre).reset_index()
            table_html(ct, total=True, tirets=True)

    # ================= 3. Validation =================
    with onglets[2]:
        if valid.empty:
            st.info("Aucune fiche pour ces filtres.")
        else:
            titre("Statut de validation des fiches par localité / superviseur",
                  "Cumul de toutes les fiches de la collecte (tous les jours) · délai de validation : 24 h")
            long = valid.melt(id_vars=["Localité", "Superviseur"], value_vars=STATUTS, var_name="Statut", value_name="Fiches")
            long["Lib"] = [lib_loc(a, b) for a, b in zip(long["Localité"], long["Superviseur"])]
            ordre_lib = [lib_loc(a, b) for a, b in zip(valid["Localité"], valid["Superviseur"])]
            fig = px.bar(long, x="Fiches", y="Lib", color="Statut", orientation="h", color_discrete_map=COULEURS_STATUT,
                         category_orders={"Statut": STATUTS, "Lib": ordre_lib}, text="Fiches")
            fig.update_traces(texttemplate="%{text}", textposition="inside", insidetextanchor="middle",
                              marker_line_color="#ffffff", marker_line_width=1,
                              hovertemplate="<b>%{y}</b><br>%{fullData.name} : %{x}<extra></extra>")
            fig.for_each_trace(lambda tr: tr.update(text=[v if v else "" for v in tr.x]))
            mise_en_forme(fig, max(260, 40 * len(valid) + 80)).update_layout(
                showlegend=True, barmode="stack", bargap=0.35, barcornerradius=0,
                legend=dict(orientation="h", y=1.08, x=0, title=None))
            graphique(fig, "valid")
            table_html(valid, pct_cols=("% traitées",))
            st.caption("« Traitées » = approuvées, non approuvées ou mises en attente. "
                       "« Non vérifiées > 24 h » = fiches reçues depuis plus de 24 h et toujours sans statut.")
            with st.expander("Détail par agent (cumul de tous les jours)"):
                st.dataframe(validation_par_agent(t), width="stretch", hide_index=True,
                             column_config={"% traitées": st.column_config.ProgressColumn(
                                 "% traitées", min_value=0, max_value=100, format="%.0f %%")})

    # ================= 4. Agents =================
    with onglets[3]:
        if t.empty:
            st.info("Aucune fiche pour ces filtres.")
        else:
            pa = performance_agents(t, auj, hier)
            inactifs = sorted(set(base["Agent"]) - set(base.loc[base["Jour"] == auj, "Agent"]))
            if inactifs:
                st.warning(f"**{len(inactifs)} agent(s) sans fiche reçue aujourd'hui** : " + ", ".join(inactifs))
            titre("Performance des agents", "Cumul sur la période choisie")
            st.dataframe(pa, width="stretch", hide_index=True, height=min(38 * len(pa) + 40, 620), column_config={
                "Total": st.column_config.ProgressColumn("Total", min_value=0, max_value=int(pa["Total"].max()), format="%d"),
                "% sans alerte": st.column_config.NumberColumn(format="%d %%"),
                "% GPS fiable": st.column_config.NumberColumn(format="%d %%"),
                "% durée conforme": st.column_config.NumberColumn(format="%d %%"),
                "Durée médiane (min)": st.column_config.NumberColumn(format="%d"),
                "Dernière fiche": st.column_config.DatetimeColumn(format="DD/MM HH:mm")})
            titre("Fiches par agent et par jour")
            croise = t.pivot_table(index="Agent", columns="Jour", values="ID", aggfunc="count", fill_value=0)
            croise = croise.loc[croise.sum(axis=1).sort_values(ascending=False).index]
            croise.columns = [pd.Timestamp(c).strftime("%d/%m") for c in croise.columns]
            fig = px.imshow(croise, text_auto=True, aspect="auto",
                            color_continuous_scale=["#f3f2ee", "#cfe0f6", "#7fb0ea", BLEU, "#14457f"])
            fig.update_traces(hovertemplate="<b>%{y}</b><br>%{x} : %{z} fiches<extra></extra>", xgap=2, ygap=2)
            mise_en_forme(fig, max(260, 26 * len(croise) + 60)).update_layout(coloraxis_showscale=False)
            fig.update_xaxes(side="top")
            graphique(fig, "agent_jour")

    # ================= 5. Carte =================
    with onglets[4]:
        geo = t.dropna(subset=["Latitude", "Longitude"]).copy()
        geo = geo[geo["Latitude"].between(-5.2, 3.8) & geo["Longitude"].between(11, 18.8)]
        vue = st.radio("Affichage", ["Points de vente", "Synthèse par localité"], horizontal=True, label_visibility="collapsed")
        st.caption(f"{len(geo)} point(s) de vente géolocalisé(s) au Congo sur {len(t)}"
                   + (" · 🔵 sans alerte · 🔴 avec alerte" if vue == "Points de vente" else " · taille du cercle = nombre de PV"))
        if geo.empty:
            st.info("Aucune coordonnée GPS sur la sélection.")
        else:
            if f_loc or f_agent:
                centre = dict(lat=geo["Latitude"].median(), lon=geo["Longitude"].median())
                etendue = max(geo["Latitude"].max() - geo["Latitude"].min(), geo["Longitude"].max() - geo["Longitude"].min(), 0.02)
                zoom = max(4.5, min(14, 8.6 - 3.32 * math.log10(etendue)))
            else:
                centre, zoom = dict(lat=-0.75, lon=15.25), 4.9      # vue d'ensemble du Congo
            if vue == "Points de vente":
                geo["Statut qualité"] = geo["Nb_alertes"].map(lambda k: "Avec alerte" if k else "Sans alerte")
                geo["Nom"] = geo["Nom_PV"].replace("", "(sans nom)")
                arg = dict(lat="Latitude", lon="Longitude", color="Statut qualité", hover_name="Nom",
                           hover_data={"Localite": True, "Type_PV_court": True, "Agent": True, "Alertes": True,
                                       "Latitude": False, "Longitude": False, "Statut qualité": False},
                           labels={"Localite": "Localité", "Type_PV_court": "Type"},
                           color_discrete_map={"Sans alerte": BLEU, "Avec alerte": ROUGE},
                           category_orders={"Statut qualité": ["Sans alerte", "Avec alerte"]})
                taille = 9
            else:
                agg = geo.groupby("Localite").agg(Latitude=("Latitude", "median"), Longitude=("Longitude", "median"),
                                                  PV=("ID", "count"), Alertes=("Nb_alertes", lambda s: (s > 0).mean() * 100)).reset_index()
                agg["Alertes"] = agg["Alertes"].round(0)
                geo = agg
                arg = dict(lat="Latitude", lon="Longitude", size="PV", size_max=48, hover_name="Localite", text="Localite",
                           hover_data={"PV": True, "Alertes": ":.0f", "Latitude": False, "Longitude": False},
                           labels={"Alertes": "% fiches avec alerte"}, color_discrete_sequence=[BLEU])
                taille = None
            if hasattr(px, "scatter_map"):
                fig = px.scatter_map(geo, map_style="carto-positron", center=centre, zoom=zoom, height=620, **arg)
            else:
                fig = px.scatter_mapbox(geo, mapbox_style="carto-positron", center=centre, zoom=zoom, height=620, **arg)
            if taille:
                fig.update_traces(marker=dict(size=taille, opacity=0.85))
            else:
                fig.update_traces(mode="markers+text", textposition="top center", marker=dict(opacity=0.75),
                                  textfont=dict(size=13, color=BLEU_FONCE))
            fig.update_layout(margin=dict(l=0, r=0, t=0, b=0),
                              legend=dict(title=None, orientation="h", y=1.0, x=0.01, bgcolor="rgba(255,255,255,.85)"))
            st.plotly_chart(fig, key="carte", config={"displayModeBar": False, "scrollZoom": True})

    # ================= 6. Qualité =================
    with onglets[5]:
        al = t[t["Nb_alertes"] > 0]
        if al.empty:
            st.success("✓ Aucune alerte qualité sur la sélection.")
        else:
            types_al = al["Alertes"].str.split(" ; ").explode().str.replace(r"\s*\(.*\)$", "", regex=True)
            a, b = st.columns(2)
            with a:
                titre("Types d'alertes", "Nombre de fiches concernées")
                graphique(barres_h(types_al.value_counts(), "Fiches", couleur=ROUGE), "types_al")
            with b:
                titre("Part des fiches avec alerte par localité")
                pl = (t.groupby("Localite")["Nb_alertes"].apply(lambda s: (s > 0).mean() * 100)).round(1)
                graphique(barres_h(pl, "Fiches avec alerte", couleur=ROUGE, suffixe=" %", maxi=100), "loc_al")
            titre(f"{len(al)} fiche(s) à vérifier", "Les plus récentes en premier")
            cols = ["ID", "Date_soumission", "Localite", "Agent", "Nom_PV", "Quartier", "Telephone",
                    "Duree_min", "Precision_GPS_m", "Alertes"]
            st.dataframe(al[cols], width="stretch", hide_index=True, column_config={
                "Date_soumission": st.column_config.DatetimeColumn("Date", format="DD/MM HH:mm"),
                "Localite": "Localité", "Nom_PV": "Point de vente", "Duree_min": "Durée (min)",
                "Precision_GPS_m": "Précision GPS (m)",
                "Alertes": st.column_config.TextColumn("⚠ Problème(s)", width="large")})
            st.download_button("⬇ Télécharger les fiches à vérifier (Excel)", vers_excel({"Alertes": al[cols]}),
                               file_name=f"alertes_INHPC_{auj:%Y%m%d}.xlsx")

    # ================= 7. Produits =================
    with onglets[6]:
        if t.empty:
            st.info("Aucune fiche pour ces filtres.")
        else:
            cats = t["Categories"].str.split(" ; ").explode().str.strip()
            cats = cats[cats != ""]
            a, b = st.columns(2)
            with a:
                titre("Catégories de produits vendues", "Nombre de points de vente")
                if cats.empty:
                    st.info("Question « catégories de produits » non trouvée ou vide.")
                else:
                    s = cats.value_counts()
                    s.index = [libelle_court(i, 45) for i in s.index]
                    graphique(barres_h(s, "Points de vente"), "cats")
            with b:
                titre("Nombre moyen de produits / rayons cochés par PV", "Selon le type de point de vente")
                moy = t[t["Type_PV"] != ""].groupby("Type_PV_court")["Nb_produits_coches"].mean().round(1)
                if not moy.empty:
                    graphique(barres_h(moy, "Moyenne"), "moyprod")
            titre("Points de vente avec 2 produits cochés ou moins", "À vérifier sur le terrain")
            peu = t[t["Nb_produits_coches"] <= 2][["Date_soumission", "Localite", "Agent", "Nom_PV", "Type_PV_court",
                                                   "Categories", "Nb_produits_coches"]]
            st.dataframe(peu, width="stretch", hide_index=True, column_config={
                "Date_soumission": st.column_config.DatetimeColumn("Date", format="DD/MM HH:mm"),
                "Localite": "Localité", "Nom_PV": "Point de vente", "Type_PV_court": "Type",
                "Nb_produits_coches": "Produits cochés"})

    # ================= 8. Données =================
    with onglets[7]:
        vues = t.drop(columns=["Age_h", "Type_PV"], errors="ignore")
        st.dataframe(vues, width="stretch", hide_index=True, height=480,
                     column_config={"Date_soumission": st.column_config.DatetimeColumn("Date", format="DD/MM/YYYY HH:mm")})
        d1, d2 = st.columns(2)
        d1.download_button("⬇ Base de suivi filtrée (Excel)", vers_excel({"Base_suivi": vues}),
                           file_name=f"suivi_INHPC_{auj:%Y%m%d}.xlsx", width="stretch")
        d2.download_button("⬇ Base Kobo complète (CSV)", brut.to_csv(index=False, sep=";").encode("utf-8-sig"),
                           file_name=f"base_kobo_INHPC_{auj:%Y%m%d}.csv", width="stretch")

    # ================= 9. Diagnostic =================
    with onglets[8]:
        st.markdown("Question du formulaire Kobo utilisée pour chaque information. Si une ligne est fausse, "
                    "corrigez-la dans la section `[COLONNES]` des secrets.")
        st.dataframe(diagnostic(C), width="stretch", hide_index=True)
        st.caption(f"Formulaire : {nom_form}")
        with st.expander("Toutes les colonnes reçues de Kobo"):
            st.write(list(brut.columns))


tableau_de_bord()
st.markdown('<div class="pied">Institut National de la Statistique – Direction des Statistiques Économiques · '
            'Enquête INHPC financée par le projet HISWACA · Application réalisée par <b>OGNIMBA Sadri</b></div>',
            unsafe_allow_html=True)
