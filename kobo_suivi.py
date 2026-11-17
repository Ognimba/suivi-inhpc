"""
Lecture des données KoboToolbox et construction de la base de suivi
Enquête sur l'identification des points de vente (INHPC) – INS Congo
Application réalisée par OGNIMBA Sadri

Ce module ne dépend pas de Streamlit : il peut être testé seul.
"""
from __future__ import annotations

import math
import re
import time
import unicodedata

import pandas as pd
import requests

FUSEAU = "Africa/Brazzaville"

STATUTS = ["Approuvé", "Non approuvé", "En attente", "Non vérifié"]
_STATUT_UID = {
    "validation_status_approved": "Approuvé",
    "validation_status_not_approved": "Non approuvé",
    "validation_status_on_hold": "En attente",
}

PARAMETRES_DEFAUT = {
    "DUREE_MIN_MINUTES": 5,       # en dessous : durée trop courte
    "DUREE_MAX_MINUTES": 150,     # au-dessus : durée trop longue
    "PRECISION_GPS_MAX_M": 30,    # au-dessus : GPS non fiable
    "DISTANCE_DOUBLON_M": 50,     # même nom à moins de 50 m : doublon possible
    "DELAI_VALIDATION_H": 24,     # délai accordé aux superviseurs pour valider
}

# Pondération du score de performance des localités (total = 100)
POIDS_SCORE = {
    "Productivité": 20,           # PV par agent et par jour de travail (référence : 3e quartile)
    "Complétude": 20,             # fiches sans information manquante ou invalide
    "Fiabilité GPS": 15,          # GPS présent, au Congo, précision ≤ seuil
    "Durée conforme": 15,         # durée d'entretien entre les deux seuils
    "Sans doublon": 10,           # fiches non suspectes de doublon
    "Validation": 20,             # fiches de plus de 24 h traitées par le superviseur
}


class KoboErreur(Exception):
    """Erreur lisible renvoyée à l'utilisateur."""


# =====================================================================
# 1. Lecture de l'API KoboToolbox (v2)
# =====================================================================
def _appel(url: str, jeton: str, essais: int = 3) -> dict:
    entetes = {"Authorization": f"Token {jeton}", "Accept": "application/json"}
    derniere = None
    for k in range(essais):
        try:
            rep = requests.get(url, headers=entetes, timeout=90)
        except requests.RequestException as e:
            derniere = e
            time.sleep(3 * (k + 1))
            continue
        if rep.status_code in (401, 403):
            raise KoboErreur("Kobo refuse l'accès (code %s) : vérifiez le jeton API." % rep.status_code)
        if rep.status_code == 404:
            raise KoboErreur("Formulaire introuvable (404) : vérifiez ASSET_UID et le serveur.")
        if rep.status_code >= 500 or rep.status_code == 429:
            derniere = KoboErreur("Kobo indisponible (code %s)" % rep.status_code)
            time.sleep(3 * (k + 1))
            continue
        if rep.status_code >= 300:
            raise KoboErreur("Erreur Kobo %s : %s" % (rep.status_code, rep.text[:300]))
        return rep.json()
    raise KoboErreur("Impossible de joindre Kobo après %d essais : %s" % (essais, derniere))


def lire_kobo(serveur: str, asset_uid: str, jeton: str) -> tuple[dict, list]:
    base = serveur.rstrip("/") + "/api/v2/assets/" + asset_uid
    asset = _appel(base + "/?format=json", jeton)
    resultats, url, garde = [], base + "/data/?format=json&limit=1000", 0
    while url and garde < 500:
        page = _appel(url, jeton)
        resultats.extend(page.get("results", []))
        url = page.get("next")
        garde += 1
    return asset, resultats


# =====================================================================
# 2. Libellés des questions et des choix
# =====================================================================
def _sans_html(s) -> str:
    s = "" if s is None else str(s)
    return re.sub(r"\s+", " ", re.sub(r"<[^>]*>", "", s)).strip()


def _premier_libelle(l) -> str:
    if isinstance(l, list):
        for x in l:
            if x:
                return _sans_html(x)
        return ""
    return _sans_html(l)


def dictionnaire(content: dict) -> dict:
    questions, choix = {}, {}
    for c in content.get("choices", []) or []:
        liste, code = c.get("list_name"), c.get("name", c.get("$autovalue"))
        if not liste or code is None:
            continue
        choix.setdefault(liste, {})[str(code)] = _premier_libelle(c.get("label")) or str(code)
    for q in content.get("survey", []) or []:
        nom = q.get("name") or q.get("$autoname")
        if not nom:
            continue
        morceaux = str(q.get("type", "")).split()
        questions[nom] = {
            "libelle": _premier_libelle(q.get("label")),
            "type": morceaux[0] if morceaux else "",
            "liste": q.get("select_from_list_name") or (morceaux[1] if len(morceaux) > 1 else ""),
        }
    return {"questions": questions, "choix": choix}


def aplatir(resultats: list, dico: dict) -> tuple[pd.DataFrame, dict]:
    """Une ligne par soumission, colonnes « Libellé [nom] », codes remplacés par les libellés."""
    types, lignes = {}, []
    for sub in resultats:
        ligne = {}
        for cle, val in sub.items():
            if cle == "_validation_status":
                v = val if isinstance(val, dict) else {}
                ligne["_validation_status"] = _STATUT_UID.get(v.get("uid", ""), "Non vérifié")
                ligne["_validation_par"] = v.get("by_whom", "") or ""
                ligne["_validation_quand"] = v.get("timestamp", "") or ""
                continue
            if cle == "_geolocation":
                if isinstance(val, list) and len(val) == 2:
                    ligne["_geo_lat"], ligne["_geo_lon"] = val
                continue
            if cle.startswith("_") or cle.startswith("meta/") or cle == "formhub/uuid":
                if val is None or not isinstance(val, (dict, list)):
                    ligne[cle] = "" if val is None else str(val)
                continue
            nom = cle.split("/")[-1]
            q = dico["questions"].get(nom, {"libelle": "", "type": "", "liste": ""})
            entete = f"{q['libelle']} [{nom}]" if q["libelle"] else nom
            types[entete] = q["type"]
            if isinstance(val, list):
                ligne[entete] = f"{len(val)} élément(s)"
                continue
            texte = "" if val is None else str(val)
            liste = dico["choix"].get(q["liste"], {})
            if q["type"] in ("select_one", "select_one_from_file"):
                texte = liste.get(texte, texte)
            elif q["type"] == "select_multiple":
                texte = " ; ".join(liste.get(c, c) for c in texte.split() if c)
            ligne[entete] = texte
        lignes.append(ligne)
    df = pd.DataFrame(lignes)
    if "_validation_status" not in df.columns and len(df):
        df["_validation_status"] = "Non vérifié"
    return df.fillna(""), types


# =====================================================================
# 3. Détection automatique des colonnes utiles
# =====================================================================
def norm(s) -> str:
    s = "" if s is None else str(s)
    s = unicodedata.normalize("NFD", s.lower())
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s.replace("’", "'").replace("`", "'")).strip()


def _contient(h: str, w: str) -> bool:
    if re.fullmatch(r"[a-z]+", w):
        return re.search(r"(^|[^a-z0-9])" + w + r"([^a-z0-9]|$)", h) is not None
    return w in h


def col_par_mots(entetes, motifs, exclus=()) -> str | None:
    passes = [lambda h: re.sub(r"\s*\[[^\]]*\]\s*$", "", h), lambda h: h]
    for p in passes:
        for m in motifs:
            for e in entetes:
                h = p(norm(e))
                if not h or h.startswith("_"):
                    continue
                if all(_contient(h, w) for w in m) and not any(_contient(h, w) for w in exclus):
                    return e
    return None


def col_exacte(entetes, nom) -> str | None:
    for e in entetes:
        if norm(e) == nom:
            return e
    return None


def col_suffixe(entetes, suffixe) -> str | None:
    for e in entetes:
        if norm(e).endswith(suffixe):
            return e
    return None


def detecter_colonnes(df: pd.DataFrame, forcer: dict | None = None) -> dict:
    E = list(df.columns)
    C = {
        "id": col_exacte(E, "_id"),
        "soumission": col_exacte(E, "_submission_time"),
        "validation": col_exacte(E, "_validation_status"),
        "debut": col_exacte(E, "start"),
        "fin": col_exacte(E, "end"),
        "localite": col_par_mots(E, [["localite"], ["superviseur"], ["departement"]]),
        "nom_pv": col_par_mots(E, [["nom du point de vente"], ["nom", "point de vente"]],
                               ["responsable", "autre", "preciser", "agent", "enqueteur"]),
        "responsable": col_par_mots(E, [["responsable"]]),
        "quartier": col_par_mots(E, [["quartier"], ["addresse"], ["adresse"]], ["mail", "e-mail"]),
        "telephone": col_par_mots(E, [["numero de telephone"], ["telephone"]], ["autre"]),
        "type_pv": col_par_mots(E, [["type de point de vente"], ["type du point de vente"], ["type", "point de vente"]],
                                ["preciser", "precisez"]),
        "autre_precis": col_par_mots(E, [["preciser", "point de vente"], ["precisez", "point de vente"]]),
        "categories": col_par_mots(E, [["categories de produits"], ["categories", "produits"], ["categorie", "produits"]],
                                   ["/", "sous-groupe", "rayon"]),
        "lat": col_suffixe(E, "_latitude"),
        "lon": col_suffixe(E, "_longitude"),
        "prec": col_suffixe(E, "_precision"),
        "gps": col_par_mots(E, [["gps"], ["coordonnees"]], ["_latitude", "_longitude", "_altitude", "_precision"]),
    }
    agent = col_par_mots(E, [["enqueteur"], ["nom de l'agent"], ["nom", "agent"], ["votre nom"], ["agent"]],
                         ["superviseur", "responsable", "point de vente", "localite"])
    C["agent_source"] = "question" if agent else "compte"
    C["agent"] = agent or col_exacte(E, "_submitted_by") or col_exacte(E, "username") or col_exacte(E, "deviceid")
    for k, v in (forcer or {}).items():
        if v:
            trouve = v if v in E else col_par_mots(E, [[norm(v)]])
            if trouve:
                C[k] = trouve
                if k == "agent":
                    C["agent_source"] = "question"
    return C


# =====================================================================
# 4. Libellés courts et ordre des types de point de vente
# =====================================================================
def libelle_court(s: str, maxi: int = 48) -> str:
    """« 6.Prestataires privés de services (cliniques, …) » -> « 6. Prestataires privés de services »"""
    s = re.sub(r"\s*\([^)]*\)?", "", str(s or "")).replace(")", "").strip().rstrip(".:").strip()
    s = re.sub(r"^(\d+)\s*[.)-]\s*", r"\1. ", s)
    return s if len(s) <= maxi else s[: maxi - 1].rstrip() + "…"


def _numero(s: str) -> float:
    m = re.match(r"\s*(\d+)", str(s))
    return float(m.group(1)) if m else math.inf


def ordre_types(dico: dict | None, C: dict, observes) -> list[str]:
    """Libellés complets des types de PV dans l'ordre du formulaire (y compris ceux encore absents)."""
    ordre = []
    if dico and C.get("type_pv"):
        nom = re.search(r"\[([^\]]+)\]\s*$", C["type_pv"])
        q = dico["questions"].get(nom.group(1) if nom else C["type_pv"], {})
        ordre = list(dico["choix"].get(q.get("liste", ""), {}).values())
    for t in observes:
        if t and t not in ordre:
            ordre.append(t)
    return sorted(ordre, key=lambda t: (_numero(t), ordre.index(t)))


def separer_localite(s: str) -> tuple[str, str]:
    """« Dolisie (OGNIMBA Sadri) » -> (« Dolisie », « OGNIMBA Sadri »)"""
    s = re.sub(r"\s+", " ", str(s or "")).strip()
    m = re.match(r"^(.*?)\s*\((.*)\)\s*$", s)
    if m:
        return m.group(1).strip(" -"), m.group(2).strip()
    return s, ""


# =====================================================================
# 5. Construction de la base de suivi + contrôles qualité
# =====================================================================
def _nombre(v):
    if v is None or v == "":
        return None
    try:
        x = float(str(v).replace(",", "."))
        return None if math.isnan(x) else x
    except ValueError:
        return None


def nom_propre(s) -> str:
    s = re.sub(r"\s+", " ", str(s or "")).strip()
    if not s:
        return ""
    # Nom de famille en majuscules conservé, sinon mise en forme « Jean Mabiala »
    if s.isupper():
        return s
    s = s.lower()
    return re.sub(r"(^|[\s\-'])(\S)", lambda m: m.group(1) + m.group(2).upper(), s)


def distance_m(la1, lo1, la2, lo2) -> float:
    r = math.pi / 180
    dla, dlo = (la2 - la1) * r, (lo2 - lo1) * r
    a = math.sin(dla / 2) ** 2 + math.cos(la1 * r) * math.cos(la2 * r) * math.sin(dlo / 2) ** 2
    return 2 * 6371000 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def construire_suivi(df: pd.DataFrame, types: dict, parametres: dict | None = None,
                     forcer: dict | None = None, dico: dict | None = None) -> tuple[pd.DataFrame, dict]:
    P = {**PARAMETRES_DEFAUT, **(parametres or {})}
    if df.empty:
        return pd.DataFrame(), {"ordre_types": []}
    C = detecter_colonnes(df, forcer)

    def col(k):
        c = C.get(k)
        return df[c].astype(str).str.strip() if c else pd.Series([""] * len(df), index=df.index)

    f = pd.DataFrame(index=df.index)
    f["ID"] = col("id")
    soum = pd.to_datetime(col("soumission").replace("", None), errors="coerce", utc=True)
    f["Date_soumission"] = soum.dt.tz_convert(FUSEAU).dt.tz_localize(None)
    f["Jour"] = f["Date_soumission"].dt.date
    f["Heure"] = f["Date_soumission"].dt.strftime("%H:%M")

    brut_loc = col("localite")
    parts = brut_loc.map(separer_localite)
    f["Localite"] = parts.map(lambda p: p[0]).replace("", "Non renseignée")
    f["Superviseur"] = parts.map(lambda p: p[1])
    f["Agent"] = col("agent").map(nom_propre).replace("", "Non identifié")
    f["Nom_PV"] = col("nom_pv")
    f["Responsable"] = col("responsable")
    f["Quartier"] = col("quartier")
    f["Telephone"] = col("telephone")
    f["Type_PV"] = col("type_pv")
    f["Type_PV_court"] = f["Type_PV"].map(lambda s: libelle_court(s) if s else "Non renseigné")
    f["Autre_precise"] = col("autre_precis")
    f["Categories"] = col("categories")

    # GPS
    lat, lon, prec = col("lat").map(_nombre), col("lon").map(_nombre), col("prec").map(_nombre)
    if C.get("gps"):
        m = col("gps").str.split()
        lat = lat.where(lat.notna(), m.map(lambda p: _nombre(p[0]) if isinstance(p, list) and len(p) > 1 else None))
        lon = lon.where(lon.notna(), m.map(lambda p: _nombre(p[1]) if isinstance(p, list) and len(p) > 1 else None))
        prec = prec.where(prec.notna(), m.map(lambda p: _nombre(p[3]) if isinstance(p, list) and len(p) > 3 else None))
    if "_geo_lat" in df.columns:
        lat = lat.where(lat.notna(), df["_geo_lat"].map(_nombre))
        lon = lon.where(lon.notna(), df["_geo_lon"].map(_nombre))
    f["Latitude"] = pd.to_numeric(lat, errors="coerce")
    f["Longitude"] = pd.to_numeric(lon, errors="coerce")
    f["Precision_GPS_m"] = pd.to_numeric(prec, errors="coerce")

    # Durée de l'entretien (horloge du téléphone : end - start)
    d = pd.to_datetime(col("debut").replace("", None), errors="coerce", utc=True)
    fi = pd.to_datetime(col("fin").replace("", None), errors="coerce", utc=True)
    f["Duree_min"] = ((fi - d).dt.total_seconds() / 60).round()

    multi = [e for e, t in types.items() if t == "select_multiple" and e in df.columns and e != C.get("categories")]
    f["Nb_produits_coches"] = (df[multi].apply(
        lambda r: sum(len([x for x in str(v).split(" ; ") if x.strip()]) for v in r), axis=1) if multi else 0)

    # Statut de validation
    f["Statut"] = df["_validation_status"].replace("", "Non vérifié") if "_validation_status" in df.columns else "Non vérifié"
    f["Valide_par"] = df["_validation_par"] if "_validation_par" in df.columns else ""
    maintenant = pd.Timestamp.now(tz=FUSEAU).tz_localize(None)
    f["Age_h"] = (maintenant - f["Date_soumission"]).dt.total_seconds() / 3600
    f["Hors_delai_validation"] = (f["Statut"] == "Non vérifié") & (f["Age_h"] > P["DELAI_VALIDATION_H"])

    # ---- Contrôles qualité (une colonne par famille, pour le score) ----
    al_comp, al_gps, al_duree = {}, {}, {}
    for i, r in f.iterrows():
        comp, gps, du = [], [], []
        if C["agent_source"] == "question" and r["Agent"] == "Non identifié":
            comp.append("Nom de l'agent manquant")
        if not r["Nom_PV"]:
            comp.append("Nom du PV manquant")
        if C.get("type_pv") and not r["Type_PV"]:
            comp.append("Type de PV manquant")
        if re.search(r"autre", r["Type_PV"], re.I) and C.get("autre_precis") and not r["Autre_precise"]:
            comp.append("Type « autre » non précisé")
        if C.get("categories") and not r["Categories"]:
            comp.append("Aucune catégorie de produits")
        if r["Telephone"]:
            chiffres = re.sub(r"^(00)?242", "", re.sub(r"\D", "", r["Telephone"]))
            if len(chiffres) != 9:
                comp.append("Téléphone invalide")
        if pd.isna(r["Latitude"]) or pd.isna(r["Longitude"]):
            gps.append("GPS manquant")
        elif not (-5.2 <= r["Latitude"] <= 3.8 and 11 <= r["Longitude"] <= 18.8):
            gps.append("GPS hors du Congo")
        elif pd.notna(r["Precision_GPS_m"]) and r["Precision_GPS_m"] > P["PRECISION_GPS_MAX_M"]:
            gps.append(f"GPS non fiable ({round(r['Precision_GPS_m'])} m)")
        if pd.notna(r["Duree_min"]):
            if r["Duree_min"] < P["DUREE_MIN_MINUTES"]:
                du.append(f"Durée trop courte ({int(r['Duree_min'])} min)")
            elif r["Duree_min"] > P["DUREE_MAX_MINUTES"]:
                du.append(f"Durée trop longue ({int(r['Duree_min'])} min)")
        al_comp[i], al_gps[i], al_duree[i] = comp, gps, du

    # Doublons possibles
    partenaires = {i: set() for i in f.index}

    def doublon(i, j):
        partenaires[i].add(j)
        partenaires[j].add(i)

    def a_gps(R):
        return pd.notna(R["Latitude"]) and pd.notna(R["Longitude"])

    cle_nom = f["Nom_PV"].map(norm)
    for _, g in f[cle_nom != ""].groupby(cle_nom[cle_nom != ""]):
        idx = list(g.index)
        for a_ in range(len(idx)):
            for b_ in range(a_ + 1, len(idx)):
                A, B = f.loc[idx[a_]], f.loc[idx[b_]]
                if a_gps(A) and a_gps(B):
                    meme = distance_m(A["Latitude"], A["Longitude"], B["Latitude"], B["Longitude"]) < P["DISTANCE_DOUBLON_M"]
                else:
                    meme = bool(A["Quartier"]) and norm(A["Quartier"]) == norm(B["Quartier"])
                if meme:
                    doublon(idx[a_], idx[b_])
    tel = f["Telephone"].map(lambda t: re.sub(r"\D", "", t)[-9:])
    cle_tel = tel + "|" + f["Responsable"].map(norm)
    ok = (tel.str.len() == 9) & (f["Responsable"] != "")
    for _, g in f[ok].groupby(cle_tel[ok]):
        idx = list(g.index)
        for a_ in range(len(idx)):
            for b_ in range(a_ + 1, len(idx)):
                doublon(idx[a_], idx[b_])

    al_doub = {}
    for i, autres in partenaires.items():
        ids = sorted(str(f.at[j, "ID"]) for j in autres)
        al_doub[i] = ["Doublon possible (ID " + ", ".join(ids[:5]) + (" …" if len(ids) > 5 else "") + ")"] if ids else []

    f["Alerte_completude"] = [bool(al_comp[i]) for i in f.index]
    f["Alerte_gps"] = [bool(al_gps[i]) for i in f.index]
    f["Alerte_duree"] = [bool(al_duree[i]) for i in f.index]
    f["Alerte_doublon"] = [bool(al_doub[i]) for i in f.index]
    f["Alertes"] = [" ; ".join(al_comp[i] + al_gps[i] + al_duree[i] + al_doub[i]) for i in f.index]
    f["Nb_alertes"] = [len(al_comp[i] + al_gps[i] + al_duree[i] + al_doub[i]) for i in f.index]

    f = f.sort_values("Date_soumission", ascending=False, na_position="last").reset_index(drop=True)
    C["ordre_types"] = ordre_types(dico, C, f["Type_PV"].unique())
    C["parametres"] = P
    return f, C


# =====================================================================
# 6. Tableaux d'analyse
# =====================================================================
def tableau_types(f: pd.DataFrame, ordre: list[str]) -> pd.DataFrame:
    """Type de point de vente | Nbre de points de vente | Pourcentage (ordre du formulaire)."""
    n = f["Type_PV"].value_counts()
    lignes = [{"Type de point de vente": libelle_court(t, 80), "Nbre de points de vente": int(n.get(t, 0))} for t in ordre]
    if (f["Type_PV"] == "").any():
        lignes.append({"Type de point de vente": "Non renseigné", "Nbre de points de vente": int((f["Type_PV"] == "").sum())})
    t = pd.DataFrame(lignes, columns=["Type de point de vente", "Nbre de points de vente"])
    total = int(t["Nbre de points de vente"].sum())
    t["Pourcentage"] = (t["Nbre de points de vente"] / total * 100).round(1) if total else 0.0
    t.loc[len(t)] = ["Total", total, 100.0 if total else 0.0]
    return t


def croise_types_localites(f: pd.DataFrame, ordre: list[str]) -> pd.DataFrame:
    """Lignes = types de PV (ordre du formulaire), colonnes = localités, avec totaux."""
    locs = f["Localite"].value_counts().index.tolist()
    lignes = list(ordre) + ([""] if (f["Type_PV"] == "").any() else [])
    ct = pd.crosstab(f["Type_PV"], f["Localite"]).reindex(index=lignes, columns=locs, fill_value=0)
    ct.index = [libelle_court(t, 60) if t else "Non renseigné" for t in ct.index]
    ct["Total"] = ct.sum(axis=1)
    ct.loc["Total"] = ct.sum(axis=0)
    ct.index.name = "Type de point de vente"
    return ct.astype(int)


def validation_par_localite(f: pd.DataFrame) -> pd.DataFrame:
    """Cumul de toutes les fiches de la collecte, par localité / superviseur."""
    g = f.groupby(["Localite", "Superviseur"])
    t = pd.crosstab([f["Localite"], f["Superviseur"]], f["Statut"]).reindex(columns=STATUTS, fill_value=0)
    t.insert(0, "Total fiches", g.size())
    t["Non vérifiées > 24 h"] = g["Hors_delai_validation"].sum().astype(int)
    t["% traitées"] = ((t["Total fiches"] - t["Non vérifié"]) / t["Total fiches"] * 100).round(1)
    return t.sort_values("% traitées", ascending=False).reset_index().rename(
        columns={"Localite": "Localité"})


def validation_par_agent(f: pd.DataFrame) -> pd.DataFrame:
    g = f.groupby(["Localite", "Agent"])
    t = pd.crosstab([f["Localite"], f["Agent"]], f["Statut"]).reindex(columns=STATUTS, fill_value=0)
    t.insert(0, "Total fiches", g.size())
    t["Non vérifiées > 24 h"] = g["Hors_delai_validation"].sum().astype(int)
    t["% traitées"] = ((t["Total fiches"] - t["Non vérifié"]) / t["Total fiches"] * 100).round(1)
    return t.reset_index().rename(columns={"Localite": "Localité"}).sort_values(["Localité", "Agent"])


def _pct(s: pd.Series) -> float:
    return float(s.mean() * 100) if len(s) else float("nan")


def performance_localites(f: pd.DataFrame, poids: dict | None = None) -> pd.DataFrame:
    """Score multicritère /100 par localité. Chaque critère est exprimé en % (100 = parfait),
    la productivité est rapportée à la meilleure localité."""
    W = poids or POIDS_SCORE
    lignes = []
    for (loc, sup), g in f.groupby(["Localite", "Superviseur"]):
        agent_jours = g[["Agent", "Jour"]].drop_duplicates().shape[0]
        anciens = g[g["Age_h"] > 24]
        lignes.append({
            "Localité": loc, "Superviseur": sup,
            "PV identifiés": len(g), "Agents": g["Agent"].nunique(),
            "PV / agent / jour": len(g) / max(agent_jours, 1),
            "Complétude": 100 - _pct(g["Alerte_completude"]),
            "Fiabilité GPS": 100 - _pct(g["Alerte_gps"]),
            "Durée conforme": 100 - _pct(g["Alerte_duree"]),
            "Sans doublon": 100 - _pct(g["Alerte_doublon"]),
            "Validation": _pct(anciens["Statut"] != "Non vérifié") if len(anciens) else 100.0,
            "Durée médiane (min)": g["Duree_min"].median(),
        })
    t = pd.DataFrame(lignes)
    if t.empty:
        return t
    reference = t["PV / agent / jour"].quantile(0.75) or t["PV / agent / jour"].max() or 1
    t["Productivité"] = (t["PV / agent / jour"] / reference * 100).clip(upper=100)
    t["Score /100"] = sum(t[k].fillna(100) * w for k, w in W.items()) / sum(W.values())
    t = t.sort_values("Score /100", ascending=False).reset_index(drop=True)
    t.insert(0, "Rang", range(1, len(t) + 1))
    t["Remarque"] = t["PV identifiés"].map(lambda n: "Effectif faible (< 10 fiches)" if n < 10 else "")
    for c in ["PV / agent / jour"]:
        t[c] = t[c].round(1)
    for c in list(W) + ["Score /100"]:
        t[c] = t[c].round(1)
    return t


def performance_agents(f: pd.DataFrame, auj, hier) -> pd.DataFrame:
    g = f.groupby("Agent")
    t = pd.DataFrame({
        "Localité": g["Localite"].agg(lambda s: s.mode().iat[0] if len(s.mode()) else ""),
        "Total": g.size(),
        "Aujourd'hui": g["Jour"].apply(lambda s: int((s == auj).sum())),
        "Hier": g["Jour"].apply(lambda s: int((s == hier).sum())),
        "Jours actifs": g["Jour"].nunique(),
        "% sans alerte": g["Nb_alertes"].apply(lambda s: round(float((s == 0).mean() * 100), 0)),
        "% GPS fiable": g["Alerte_gps"].apply(lambda s: round(float((~s).mean() * 100), 0)),
        "% durée conforme": g["Alerte_duree"].apply(lambda s: round(float((~s).mean() * 100), 0)),
        "Durée médiane (min)": g["Duree_min"].median(),
        "Approuvées": g["Statut"].apply(lambda s: int((s == "Approuvé").sum())),
        "Dernière fiche": g["Date_soumission"].max(),
    })
    t["Moyenne / jour"] = (t["Total"] / t["Jours actifs"].clip(lower=1)).round(1)
    cols = ["Localité", "Total", "Aujourd'hui", "Hier", "Jours actifs", "Moyenne / jour", "% sans alerte",
            "% GPS fiable", "% durée conforme", "Durée médiane (min)", "Approuvées", "Dernière fiche"]
    return t[cols].sort_values(["Localité", "Total"], ascending=[True, False]).reset_index()


NOMS_COLONNES = {
    "localite": "Localité / Superviseur", "agent": "Agent", "nom_pv": "Nom du PV",
    "responsable": "Responsable", "quartier": "Adresse / quartier", "telephone": "Téléphone",
    "type_pv": "Type de PV", "autre_precis": "Autre type précisé", "categories": "Catégories de produits",
    "gps": "GPS", "lat": "Latitude", "lon": "Longitude", "prec": "Précision GPS",
    "debut": "Début entretien (start)", "fin": "Fin entretien (end)", "soumission": "Date de soumission",
    "validation": "Statut de validation",
}


def diagnostic(C: dict) -> pd.DataFrame:
    lignes = []
    for k, lib in NOMS_COLONNES.items():
        v = C.get(k)
        if not v and k in ("lat", "lon", "prec") and C.get("gps"):
            v = f"lu dans « {C['gps']} »"
        lignes.append({"Information": lib, "Question Kobo retenue": v or "— non trouvée —"})
    if C.get("agent_source") == "compte":
        lignes.append({"Information": "Remarque",
                       "Question Kobo retenue": "Pas de question « nom de l'agent » : l'agent est le compte Kobo."})
    return pd.DataFrame(lignes)
