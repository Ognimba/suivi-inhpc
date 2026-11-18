"""
SUIVI EN TEMPS RÉEL – ENQUÊTE SUR L'IDENTIFICATION DES POINTS DE VENTE (INHPC)
Application Streamlit connectée à KoboToolbox – INS Congo

Lancer en local :  streamlit run app.py
Paramètres      :  .streamlit/secrets.toml  (voir secrets.toml.exemple)
"""
from __future__ import annotations

import hmac
import io
import os
from datetime import datetime, timedelta

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from kobo_suivi import (FUSEAU, KoboErreur, aplatir, construire_suivi, diagnostic,
                        dictionnaire, lire_kobo)

st.set_page_config(page_title="Suivi collecte INHPC", page_icon="📍", layout="wide")

# ---------------------------------------------------------------------
# Paramètres (secrets Streamlit, sinon variables d'environnement)
# ---------------------------------------------------------------------
def param(nom, defaut=None):
    try:
        if nom in st.secrets:
            return st.secrets[nom]
    except Exception:  # pas de fichier secrets
        pass
    return os.environ.get(nom, defaut)


JETON = param("KOBO_TOKEN", "")
SERVEUR = param("KOBO_SERVEUR", "https://kf.kobotoolbox.org")
ASSET_UID = param("ASSET_UID", "abVbPCHqtiYSwKP8YAZ4QP")
MINUTES = max(5, int(param("ACTUALISATION_MINUTES", 15)))
MOT_DE_PASSE = str(param("MOT_DE_PASSE", "") or "")
SOUS_TITRE = param("SOUS_TITRE", "INS – Direction des Statistiques Économiques")
OBJECTIF_TOTAL = int(param("OBJECTIF_TOTAL", 0) or 0)
OBJECTIF_JOUR_AGENT = int(param("OBJECTIF_JOUR_PAR_AGENT", 0) or 0)
OBJECTIFS_DEPT = {str(k): int(v) for k, v in dict(param("OBJECTIFS_DEPARTEMENT", {}) or {}).items() if v}
DEPT_DEFAUT = param("DEPARTEMENT_PAR_DEFAUT", "") or ""
COLONNES_FORCEES = dict(param("COLONNES", {}) or {})
SEUILS = {k: float(param(k)) for k in ("DUREE_MIN_MINUTES", "DUREE_MAX_MINUTES",
                                         "PRECISION_GPS_MAX_M", "DISTANCE_DOUBLON_M") if param(k) is not None}

BLEU, ROUGE, GRIS = "#2a78d6", "#d03b3b", "#7a7974"

st.markdown("""
<style>
  .block-container { padding-top: 1.2rem; }
  div[data-testid="stMetric"] { background: #fcfcfb; border: 1px solid rgba(11,11,11,.10);
      border-radius: 10px; padding: 12px 14px; }
  div[data-testid="stMetricLabel"] p { font-size: .85rem; }
  .entete { background: #16324f; color: #fff; padding: 14px 20px; border-radius: 10px; margin-bottom: 12px; }
  .entete h1 { font-size: 1.45rem; margin: 0; color: #fff; padding: 0; }
  .entete .sous { font-size: .85rem; opacity: .85; margin-top: 4px; }
</style>""", unsafe_allow_html=True)


# ---------------------------------------------------------------------
# Accès protégé par mot de passe (facultatif)
# ---------------------------------------------------------------------
def acces_autorise() -> bool:
    if not MOT_DE_PASSE or st.session_state.get("ok_acces"):
        return True
    st.markdown('<div class="entete"><h1>Suivi de la collecte – Points de vente (INHPC)</h1></div>',
                unsafe_allow_html=True)
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
    st.error("Le jeton API Kobo n'est pas configuré. Ajoutez KOBO_TOKEN dans les secrets de l'application "
             "(voir le guide README).")
    st.stop()


# ---------------------------------------------------------------------
# Chargement des données (mis en cache, partagé entre tous les visiteurs)
# ---------------------------------------------------------------------
@st.cache_data(ttl=MINUTES * 60 - 30, show_spinner="Chargement des données KoboToolbox…")
def charger(serveur, uid, jeton, seuils_t, colonnes_t):
    asset, resultats = lire_kobo(serveur, uid, jeton)
    dico = dictionnaire(asset.get("content", {}) or {})
    brut, types = aplatir(resultats, dico)
    fiches, C = construire_suivi(brut, types, dict(seuils_t), dict(colonnes_t))
    heure = pd.Timestamp.now(tz=FUSEAU).tz_localize(None)
    return brut, fiches, C, heure, asset.get("name", "")


def aujourdhui():
    return pd.Timestamp.now(tz=FUSEAU).date()


def mise_en_forme(fig, hauteur=330):
    fig.update_layout(template="plotly_white", height=hauteur, margin=dict(l=10, r=10, t=10, b=10),
                      font=dict(family="Segoe UI, Roboto, Helvetica, Arial, sans-serif", size=12, color="#0b0b0b"),
                      barcornerradius=4, bargap=0.25, hoverlabel=dict(bgcolor="#0b0b0b", font_color="#fff"))
    fig.update_xaxes(showgrid=False, linecolor="#c3c2b7", title=None)
    fig.update_yaxes(gridcolor="#e1e0d9", zeroline=False, title=None)
    return fig


def barres_h(serie: pd.Series, libelle="Fiches", hauteur=None, couleur=BLEU):
    d = serie.sort_values().reset_index()
    d.columns = ["nom", "n"]
    fig = px.bar(d, x="n", y="nom", orientation="h", text="n", color_discrete_sequence=[couleur],
                 labels={"n": libelle, "nom": ""})
    fig.update_traces(textposition="outside", cliponaxis=False,
                      hovertemplate="<b>%{y}</b><br>" + libelle + " : %{x}<extra></extra>")
    mise_en_forme(fig, hauteur or max(220, 28 * len(d) + 40))
    fig.update_xaxes(showgrid=True, gridcolor="#e1e0d9", showticklabels=False)
    return fig


def vers_excel(feuilles: dict) -> bytes:
    tampon = io.BytesIO()
    with pd.ExcelWriter(tampon, engine="openpyxl") as w:
        for nom, d in feuilles.items():
            d.to_excel(w, sheet_name=nom[:31], index=False)
    return tampon.getvalue()


# ---------------------------------------------------------------------
# En-tête et barre latérale
# ---------------------------------------------------------------------
st.markdown(f"""<div class="entete"><h1>📍 Suivi de la collecte – Identification des points de vente (INHPC)</h1>
<div class="sous">{SOUS_TITRE} · Données KoboToolbox · actualisation automatique toutes les {MINUTES} min</div></div>""",
            unsafe_allow_html=True)

with st.sidebar:
    st.subheader("Actualisation")
    if st.button("🔄 Actualiser maintenant", width="stretch"):
        charger.clear()
        st.rerun()
    st.caption(f"Les données sont relues dans Kobo toutes les {MINUTES} minutes. "
               "La page se met à jour seule tant qu'elle reste ouverte.")
    st.divider()
    st.caption("Les fiches avec alerte ne sont pas forcément fausses : ce sont des points à vérifier avec l'agent.")


# ---------------------------------------------------------------------
# Tableau de bord (fragment ré-exécuté automatiquement)
# ---------------------------------------------------------------------
@st.fragment(run_every=timedelta(minutes=MINUTES))
def tableau_de_bord():
    try:
        brut, fiches, C, heure_maj, nom_form = charger(SERVEUR, ASSET_UID, JETON,
                                                       tuple(sorted(SEUILS.items())),
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

    auj = aujourdhui()
    hier = auj - timedelta(days=1)

    # ---------------- Filtres ----------------
    c1, c2, c3, c4 = st.columns([1.2, 1.4, 1.4, 1.2])
    depts = sorted([d for d in fiches["Departement"].unique() if d])
    defaut = [d for d in depts if DEPT_DEFAUT and d.lower() == str(DEPT_DEFAUT).lower()]
    f_dept = c1.multiselect("Département", depts, default=defaut, placeholder="Tous") if depts else []
    base = fiches[fiches["Departement"].isin(f_dept)] if f_dept else fiches
    f_agent = c2.multiselect("Agent", sorted(base["Agent"].unique()), placeholder="Tous")
    types_pv = sorted([t for t in base["Type_PV"].unique() if t])
    f_type = c3.multiselect("Type de point de vente", types_pv, placeholder="Tous")
    periode = c4.selectbox("Période", ["Toute la collecte", "Aujourd'hui", "Hier", "7 derniers jours", "Personnalisée"])
    if f_agent:
        base = base[base["Agent"].isin(f_agent)]
    if f_type:
        base = base[base["Type_PV"].isin(f_type)]

    du, au = None, None
    if periode == "Aujourd'hui":
        du = au = auj
    elif periode == "Hier":
        du = au = hier
    elif periode == "7 derniers jours":
        du, au = auj - timedelta(days=6), auj
    elif periode == "Personnalisée":
        jours = base["Jour"].dropna()
        debut = jours.min() if len(jours) else auj
        choix = st.date_input("Du … au …", value=(debut, auj), format="DD/MM/YYYY")
        if isinstance(choix, (tuple, list)) and len(choix) == 2:
            du, au = choix
    t = base
    if du:
        t = t[(t["Jour"] >= du) & (t["Jour"] <= au)]

    # ---------------- Indicateurs clés ----------------
    n_auj = int((base["Jour"] == auj).sum())
    n_hier = int((base["Jour"] == hier).sum())
    agents_tous = base["Agent"].nunique()
    agents_auj = base.loc[base["Jour"] == auj, "Agent"].nunique()
    n_al = int((t["Nb_alertes"] > 0).sum())
    k = st.columns(5)
    k[0].metric("Points de vente identifiés", f"{len(t):,}".replace(",", " "),
                help="Nombre de fiches reçues sur la période et les filtres choisis")
    k[1].metric("Fiches reçues aujourd'hui", n_auj, delta=n_auj - n_hier, delta_color="off",
                help=f"Écart par rapport à hier ({n_hier} fiches)")
    k[2].metric("Agents actifs aujourd'hui", f"{agents_auj} / {agents_tous}")
    k[3].metric("Fiches avec alerte qualité", f"{n_al}  ({(n_al / len(t) * 100 if len(t) else 0):.0f} %)")
    objectif = sum(OBJECTIFS_DEPT.get(d, 0) for d in f_dept) if f_dept and OBJECTIFS_DEPT else OBJECTIF_TOTAL
    if objectif and not f_agent and not f_type:
        taux = len(base) / objectif * 100
        k[4].metric("Taux de réalisation", f"{taux:.1f} %", help=f"{len(base)} fiches sur un objectif de {objectif}")
        k[4].progress(min(1.0, taux / 100))
    else:
        moy = t["Nb_produits_coches"].mean() if len(t) else 0
        k[4].metric("Produits cochés / PV (moyenne)", f"{moy:.1f}".replace(".", ","))

    derniere = fiches["Date_soumission"].max()
    st.caption(f"Données lues dans Kobo à **{heure_maj:%H:%M}** · prochaine lecture vers "
               f"**{(heure_maj + timedelta(minutes=MINUTES)):%H:%M}** · dernière fiche reçue le "
               f"**{derniere:%d/%m/%Y à %H:%M}** · formulaire : {nom_form}")

    onglets = st.tabs(["📈 Avancement", "👥 Agents", "🗺️ Carte", "🛒 Produits", "⚠️ Qualité", "📋 Données", "🔧 Diagnostic"])

    # ---------------- 1. Avancement ----------------
    with onglets[0]:
        if t.empty:
            st.info("Aucune fiche pour ces filtres.")
        else:
            par_jour = t.groupby("Jour").size()
            plage = pd.date_range(par_jour.index.min(), max(par_jour.index.max(), au or par_jour.index.max()))
            par_jour = par_jour.reindex(plage.date, fill_value=0)
            dj = pd.DataFrame({"Jour": pd.to_datetime(par_jour.index), "Fiches": par_jour.values})
            dj["Cumul"] = dj["Fiches"].cumsum()
            g1, g2 = st.columns(2)
            with g1:
                st.markdown("**Fiches reçues par jour**")
                fig = px.bar(dj, x="Jour", y="Fiches", text="Fiches", color_discrete_sequence=[BLEU])
                fig.update_traces(textposition="outside", cliponaxis=False,
                                  hovertemplate="<b>%{x|%d/%m/%Y}</b><br>%{y} fiches<extra></extra>")
                mise_en_forme(fig).update_xaxes(tickformat="%d/%m")
                st.plotly_chart(fig, key="jour")
            with g2:
                st.markdown("**Cumul des fiches**")
                fig = px.line(dj, x="Jour", y="Cumul", markers=True, color_discrete_sequence=[BLEU])
                fig.update_traces(line_width=2, marker_size=8,
                                  hovertemplate="<b>%{x|%d/%m/%Y}</b><br>Cumul : %{y}<extra></extra>")
                if objectif and not f_agent and not f_type and not du:
                    fig.add_hline(y=objectif, line_dash="dot", line_color=GRIS,
                                  annotation_text=f"Objectif : {objectif}", annotation_position="top left")
                mise_en_forme(fig).update_xaxes(tickformat="%d/%m")
                st.plotly_chart(fig, key="cumul")
            g3, g4 = st.columns(2)
            with g3:
                st.markdown("**Fiches par département**")
                st.plotly_chart(barres_h(t["Departement"].replace("", "Non renseigné").value_counts()), key="dept")
            with g4:
                st.markdown("**Fiches par type de point de vente**")
                st.plotly_chart(barres_h(t["Type_PV"].replace("", "Non renseigné").value_counts()), key="type")
            if OBJECTIFS_DEPT:
                st.markdown("**Réalisation par département (toute la collecte)**")
                r = fiches.groupby("Departement").size()
                tab = pd.DataFrame({"Objectif": pd.Series(OBJECTIFS_DEPT)})
                tab["Réalisé"] = r.reindex(tab.index).fillna(0).astype(int)
                tab["Taux (%)"] = (tab["Réalisé"] / tab["Objectif"] * 100).round(1)
                st.dataframe(tab.reset_index(names="Département"), width="stretch", hide_index=True,
                             column_config={"Taux (%)": st.column_config.ProgressColumn(
                                 "Taux (%)", min_value=0, max_value=100, format="%.1f %%")})

    # ---------------- 2. Agents ----------------
    with onglets[1]:
        if t.empty:
            st.info("Aucune fiche pour ces filtres.")
        else:
            g = t.groupby("Agent")
            perf = pd.DataFrame({
                "Département": g["Departement"].agg(lambda s: s.mode().iat[0] if len(s.mode()) else ""),
                "Total": g.size(),
                "Aujourd'hui": g["Jour"].apply(lambda s: int((s == auj).sum())),
                "Hier": g["Jour"].apply(lambda s: int((s == hier).sum())),
                "Jours actifs": g["Jour"].nunique(),
                "Fiches avec alerte": g["Nb_alertes"].apply(lambda s: int((s > 0).sum())),
                "Durée médiane (min)": g["Duree_min"].median(),
                "Dernière fiche": g["Date_soumission"].max(),
            })
            perf["Moyenne / jour"] = (perf["Total"] / perf["Jours actifs"].clip(lower=1)).round(1)
            perf["% alertes"] = (perf["Fiches avec alerte"] / perf["Total"] * 100).round(0)
            if OBJECTIF_JOUR_AGENT:
                perf["Objectif du jour"] = perf["Aujourd'hui"].map(
                    lambda n: f"✓ atteint ({n}/{OBJECTIF_JOUR_AGENT})" if n >= OBJECTIF_JOUR_AGENT
                    else f"✗ {n}/{OBJECTIF_JOUR_AGENT}")
            perf = perf.sort_values("Total", ascending=False).reset_index()
            st.dataframe(perf, width="stretch", hide_index=True, column_config={
                "Total": st.column_config.ProgressColumn("Total", min_value=0, max_value=int(perf["Total"].max()),
                                                         format="%d"),
                "Dernière fiche": st.column_config.DatetimeColumn("Dernière fiche", format="DD/MM HH:mm"),
                "% alertes": st.column_config.NumberColumn("% alertes", format="%d %%"),
                "Durée médiane (min)": st.column_config.NumberColumn(format="%d"),
            })
            st.markdown("**Fiches par agent et par jour**")
            croise = t.pivot_table(index="Agent", columns="Jour", values="ID", aggfunc="count", fill_value=0)
            croise = croise.loc[croise.sum(axis=1).sort_values(ascending=False).index]
            croise.columns = [pd.Timestamp(c).strftime("%d/%m") for c in croise.columns]
            fig = px.imshow(croise, text_auto=True, aspect="auto",
                            color_continuous_scale=["#f3f2ee", "#cfe0f6", "#7fb0ea", BLEU, "#14457f"])
            fig.update_traces(hovertemplate="<b>%{y}</b><br>%{x} : %{z} fiches<extra></extra>", xgap=2, ygap=2)
            mise_en_forme(fig, max(240, 30 * len(croise) + 60)).update_layout(coloraxis_showscale=False)
            fig.update_xaxes(side="top")
            st.plotly_chart(fig, key="agent_jour")

    # ---------------- 3. Carte ----------------
    with onglets[2]:
        geo = t.dropna(subset=["Latitude", "Longitude"]).copy()
        st.caption(f"{len(geo)} point(s) de vente géolocalisé(s) sur {len(t)} · "
                   "🔵 sans alerte · 🔴 avec alerte (cliquez sur un point pour le détail)")
        if geo.empty:
            st.info("Aucune coordonnée GPS sur la sélection.")
        else:
            geo["Statut"] = geo["Nb_alertes"].map(lambda n: "Avec alerte" if n else "Sans alerte")
            geo["Nom"] = geo["Nom_PV"].replace("", "(sans nom)")
            arg = dict(lat="Latitude", lon="Longitude", color="Statut", hover_name="Nom",
                       hover_data={"Type_PV": True, "Agent": True, "Quartier": True, "Alertes": True,
                                   "Latitude": ":.5f", "Longitude": ":.5f", "Statut": False},
                       color_discrete_map={"Sans alerte": BLEU, "Avec alerte": ROUGE},
                       category_orders={"Statut": ["Sans alerte", "Avec alerte"]}, zoom=11, height=560)
            if hasattr(px, "scatter_map"):
                fig = px.scatter_map(geo, map_style="open-street-map", **arg)
            else:
                fig = px.scatter_mapbox(geo, mapbox_style="open-street-map", **arg)
            fig.update_traces(marker=dict(size=11))
            fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), legend=dict(title=None, orientation="h", y=1.02, x=0))
            st.plotly_chart(fig, key="carte")

    # ---------------- 4. Produits ----------------
    with onglets[3]:
        if t.empty:
            st.info("Aucune fiche pour ces filtres.")
        else:
            cats = t["Categories"].str.split(" ; ").explode().str.strip()
            cats = cats[cats != ""]
            p1, p2 = st.columns(2)
            with p1:
                st.markdown("**Catégories de produits vendues (nombre de PV)**")
                if cats.empty:
                    st.info("Question « catégories de produits » non trouvée ou vide.")
                else:
                    st.plotly_chart(barres_h(cats.value_counts(), "Points de vente"), key="cats")
            with p2:
                st.markdown("**Produits / rayons cochés par PV, selon le type**")
                moy = t[t["Type_PV"] != ""].groupby("Type_PV")["Nb_produits_coches"].mean().round(1)
                if moy.empty:
                    st.info("Pas de données.")
                else:
                    st.plotly_chart(barres_h(moy, "Moyenne"), key="moyprod")
            st.markdown("**Points de vente avec peu de produits cochés** (à vérifier sur le terrain)")
            peu = t[t["Nb_produits_coches"] <= 2][["Date_soumission", "Agent", "Departement", "Nom_PV",
                                                   "Type_PV", "Categories", "Nb_produits_coches"]]
            st.dataframe(peu, width="stretch", hide_index=True,
                         column_config={"Date_soumission": st.column_config.DatetimeColumn("Date", format="DD/MM HH:mm")})

    # ---------------- 5. Qualité ----------------
    with onglets[4]:
        al = t[t["Nb_alertes"] > 0]
        if al.empty:
            st.success("✓ Aucune alerte qualité sur la sélection.")
        else:
            types_al = al["Alertes"].str.split(" ; ").explode().str.replace(r"\s*\(.*\)$", "", regex=True)
            q1, q2 = st.columns([1, 1.4])
            with q1:
                st.markdown("**Types d'alertes**")
                st.plotly_chart(barres_h(types_al.value_counts(), "Fiches", couleur=ROUGE), key="types_al")
            with q2:
                st.markdown("**Alertes par agent**")
                st.plotly_chart(barres_h(al["Agent"].value_counts(), "Fiches avec alerte", couleur=ROUGE),
                                key="agent_al")
            st.markdown(f"**{len(al)} fiche(s) à vérifier** – les plus récentes en premier")
            cols = ["ID", "Date_soumission", "Agent", "Departement", "Nom_PV", "Quartier", "Telephone", "Alertes"]
            st.dataframe(al[cols], width="stretch", hide_index=True,
                         column_config={"Date_soumission": st.column_config.DatetimeColumn("Date", format="DD/MM HH:mm"),
                                        "Alertes": st.column_config.TextColumn("⚠ Problème(s)", width="large")})
            st.download_button("⬇ Télécharger les alertes (Excel)", vers_excel({"Alertes": al[cols]}),
                               file_name=f"alertes_INHPC_{auj:%Y%m%d}.xlsx")

    # ---------------- 6. Données ----------------
    with onglets[5]:
        st.dataframe(t, width="stretch", hide_index=True, height=480,
                     column_config={"Date_soumission": st.column_config.DatetimeColumn("Date", format="DD/MM/YYYY HH:mm")})
        d1, d2 = st.columns(2)
        d1.download_button("⬇ Base de suivi filtrée (Excel)", vers_excel({"Base_suivi": t}),
                           file_name=f"suivi_INHPC_{auj:%Y%m%d}.xlsx", width="stretch")
        d2.download_button("⬇ Base Kobo complète (CSV)", brut.to_csv(index=False, sep=";").encode("utf-8-sig"),
                           file_name=f"base_kobo_INHPC_{auj:%Y%m%d}.csv", width="stretch")

    # ---------------- 7. Diagnostic ----------------
    with onglets[6]:
        st.markdown("Questions du formulaire Kobo utilisées pour chaque information. "
                    "Si une ligne est fausse, corrigez-la dans la section `[COLONNES]` des secrets.")
        st.dataframe(diagnostic(C), width="stretch", hide_index=True)
        with st.expander("Toutes les colonnes reçues de Kobo"):
            st.write(list(brut.columns))


tableau_de_bord()
