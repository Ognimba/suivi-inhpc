"""
Rapport Word de fin de collecte – Enquête INHPC (environ 10 pages)
Application réalisée par OGNIMBA Sadri

Structure :
  • Page 1      : page de garde (logos INS et HISWACA), période, synthèse nationale
                  (chiffres clés + classement des départements)
  • 1 page par département, toujours dans le même ordre :
        1. Chiffres clés
        2. Performance (score /100, rang, évolution par jour)
        3. Points d'attention (encadré)
        4. Éléments à revoir (actions concrètes, nom de l'agent)
  • Dernière page : recommandations générales et rappel des règles
Un saut de page sépare chaque partie : le nombre de pages suit le nombre de départements.
"""
from __future__ import annotations

import io
import os

import matplotlib

matplotlib.use("Agg")                     # rendu sans écran (serveur Streamlit)
import matplotlib.pyplot as plt           # noqa: E402
import pandas as pd                       # noqa: E402
from docx import Document                 # noqa: E402
from docx.enum.section import WD_ORIENT   # noqa: E402
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH  # noqa: E402
from docx.oxml import OxmlElement         # noqa: E402
from docx.oxml.ns import qn               # noqa: E402
from docx.shared import Cm, Pt, RGBColor  # noqa: E402

from kobo_suivi import (POIDS_SCORE, PARAMETRES_DEFAUT, STATUT_DEPASSEMENT, STATUT_QUOTA,  # noqa: E402
                        alertes_envoi, controle_quota_produits, fdate, performance_departements)

BLEU = RGBColor(0x16, 0x32, 0x4F)
ACCENT = RGBColor(0x2A, 0x78, 0xD6)
GRIS = RGBColor(0x52, 0x51, 0x4E)
GRIS_CLAIR = RGBColor(0x8A, 0x89, 0x84)
ROUGE = RGBColor(0x9D, 0x22, 0x22)
VERT = RGBColor(0x1E, 0x6B, 0x1E)
BLANC = RGBColor(0xFF, 0xFF, 0xFF)
DOSSIER = os.path.dirname(os.path.abspath(__file__))
LARGEUR_UTILE = 17.8                      # cm (A4 portrait, marges 1,6 cm)


# =====================================================================
# Outils de mise en forme Word
# =====================================================================
def _fond(cell, hexa):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexa)
    tcPr.append(shd)


def _bordures(table, couleur="C3C2B7", aucune=False):
    tblPr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for cote in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{cote}")
        if aucune or cote in ("left", "right", "insideV"):
            e.set(qn("w:val"), "nil")
        else:
            e.set(qn("w:val"), "single")
            e.set(qn("w:sz"), "4")
            e.set(qn("w:color"), couleur)
        b.append(e)
    tblPr.append(b)


def _bord_cellule(cell, cote, couleur, epaisseur=24):
    """Bordure d'un seul côté d'une cellule (ex. barre rouge à gauche de l'encadré)."""
    tcPr = cell._tc.get_or_add_tcPr()
    bords = tcPr.find(qn("w:tcBorders"))
    if bords is None:
        bords = OxmlElement("w:tcBorders")
        tcPr.append(bords)
    e = OxmlElement(f"w:{cote}")
    e.set(qn("w:val"), "single")
    e.set(qn("w:sz"), str(epaisseur))
    e.set(qn("w:color"), couleur)
    bords.append(e)


def _largeurs(table, largeurs_cm):
    """Largeurs fixes respectées par Word et LibreOffice (grille + cellules)."""
    table.autofit = False
    tblPr = table._tbl.tblPr
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tblPr.append(lay)
    for j, w in enumerate(largeurs_cm):
        table.columns[j].width = Cm(w)
        for cell in table.columns[j].cells:
            cell.width = Cm(w)


def _texte(cell, txt, gras=False, couleur=None, taille=8.5, align=None, italique=False):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    if align is not None:
        p.alignment = align
    r = p.add_run(str(txt))
    r.bold, r.italic = gras, italique
    r.font.size = Pt(taille)
    if couleur:
        r.font.color.rgb = couleur
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def _para(conteneur, txt="", taille=9, gras=False, couleur=None, align=None, apres=2, avant=0, italique=False):
    p = conteneur.add_paragraph()
    p.paragraph_format.space_after = Pt(apres)
    p.paragraph_format.space_before = Pt(avant)
    if align is not None:
        p.alignment = align
    if txt:
        r = p.add_run(txt)
        r.bold, r.italic = gras, italique
        r.font.size = Pt(taille)
        if couleur:
            r.font.color.rgb = couleur
    return p


def _run(p, txt, taille=9, gras=False, couleur=None, italique=False):
    r = p.add_run(txt)
    r.bold, r.italic = gras, italique
    r.font.size = Pt(taille)
    if couleur:
        r.font.color.rgb = couleur
    return r


def _titre(doc, txt, saut=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(7)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.page_break_before = saut
    r = p.add_run(txt)
    r.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = BLEU
    pPr = p._p.get_or_add_pPr()
    bd = OxmlElement("w:pBdr")
    bas = OxmlElement("w:bottom")
    for k, v in (("w:val", "single"), ("w:sz", "6"), ("w:space", "1"), ("w:color", "2A78D6")):
        bas.set(qn(k), v)
    bd.append(bas)
    pPr.append(bd)
    return p


def _tableau(conteneur, entetes, lignes, largeurs_cm, derniere_total=False, taille=8, rouges=()):
    """Tableau standard (en-tête bleu foncé, lignes alternées). « rouges » = indices de lignes à surligner."""
    t = conteneur.add_table(rows=1, cols=len(entetes))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    _bordures(t)
    for j, h in enumerate(entetes):
        c = t.rows[0].cells[j]
        _fond(c, "16324F")
        _texte(c, h, gras=True, couleur=BLANC, taille=taille - 0.5,
               align=WD_ALIGN_PARAGRAPH.LEFT if j == 1 or j == 0 and len(entetes) < 4 else WD_ALIGN_PARAGRAPH.CENTER)
    for i, ligne in enumerate(lignes):
        cells = t.add_row().cells
        total = derniere_total and i == len(lignes) - 1
        for j, v in enumerate(ligne):
            al = WD_ALIGN_PARAGRAPH.LEFT if j == 1 or j == 0 and len(entetes) < 4 else WD_ALIGN_PARAGRAPH.CENTER
            _texte(cells[j], v, gras=total, align=al, taille=taille,
                   couleur=ROUGE if i in rouges else None)
            if total:
                _fond(cells[j], "E3ECF7")
            elif i in rouges:
                _fond(cells[j], "FBE3E3")
            elif i % 2:
                _fond(cells[j], "F6F5F1")
    _largeurs(t, largeurs_cm)
    # La ligne d'en-tête est répétée si le tableau passe sur une autre page
    trPr = t.rows[0]._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader")
    th.set(qn("w:val"), "true")
    trPr.append(th)
    return t


def _tuiles(doc, cles, largeur_totale=LARGEUR_UTILE, par_ligne=None):
    """Chiffres clés sous forme de tuiles : libellé (petit) + valeur (gros)."""
    par_ligne = par_ligne or len(cles)
    for debut in range(0, len(cles), par_ligne):
        bloc = cles[debut: debut + par_ligne]
        t = doc.add_table(rows=2, cols=len(bloc))
        _bordures(t, aucune=True)
        for j, (lib, val) in enumerate(bloc):
            for i in (0, 1):
                _fond(t.rows[i].cells[j], "E3ECF7")
                _bord_cellule(t.rows[i].cells[j], "left", "FFFFFF", 18)
            _texte(t.rows[0].cells[j], lib, couleur=GRIS, taille=7.5, align=WD_ALIGN_PARAGRAPH.CENTER)
            _texte(t.rows[1].cells[j], val, gras=True, couleur=BLEU, taille=12.5, align=WD_ALIGN_PARAGRAPH.CENTER)
        _largeurs(t, [largeur_totale / len(bloc)] * len(bloc))
        _para(doc, apres=0).paragraph_format.line_spacing = Pt(4)


def _puce(conteneur, txt=None, morceaux=None, taille=9):
    """Puce simple ; « morceaux » = liste de (texte, gras, couleur) pour mélanger les styles."""
    p = conteneur.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(1.5)
    for m in (morceaux or [(txt, False, None)]):
        _run(p, m[0], taille=taille, gras=m[1], couleur=m[2])
    return p


def _champ(p, code, taille=7.5):
    """Champ Word (numéro de page…)."""
    r = p.add_run()
    r.font.size, r.font.color.rgb = Pt(taille), GRIS
    for typ, texte in (("begin", None), (None, code), ("separate", None), (None, "1"), ("end", None)):
        if typ:
            e = OxmlElement("w:fldChar")
            e.set(qn("w:fldCharType"), typ)
        elif texte == code:
            e = OxmlElement("w:instrText")
            e.set(qn("xml:space"), "preserve")
            e.text = f" {code} "
        else:
            e = OxmlElement("w:t")
            e.text = texte
        r._r.append(e)


def _fmt(x, dec=0):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "–"
    return f"{x:,.{dec}f}".replace(",", " ").replace(".", ",")


def _pct(n, d):
    return n / d * 100 if d else 0.0


def _liste_noms(noms, maxi=4):
    noms = [n for n in noms if n]
    if not noms:
        return ""
    return ", ".join(noms[:maxi]) + (f" et {len(noms) - maxi} autre(s)" if len(noms) > maxi else "")


# =====================================================================
# Graphique : fiches reçues par jour (matplotlib -> image PNG)
# =====================================================================
def _graphique_jours(fd: pd.DataFrame, jours: pd.DatetimeIndex, largeur_cm=8.6, hauteur_cm=4.3) -> io.BytesIO:
    """Barres fines, une seule série : nombre de fiches par jour de collecte (même axe pour tous)."""
    serie = fd.groupby("Jour_collecte").size()
    serie.index = pd.to_datetime(serie.index)
    serie = serie.reindex(jours, fill_value=0)
    fig, ax = plt.subplots(figsize=(largeur_cm / 2.54, hauteur_cm / 2.54), dpi=220)
    ax.bar(range(len(serie)), serie.values, width=0.62, color="#2a78d6", zorder=3)
    for k, v in enumerate(serie.values):
        if v:
            ax.text(k, v, f"{v}", ha="center", va="bottom", fontsize=5.5, color="#3b3a37")
    pas = max(1, len(serie) // 10)
    ax.set_xticks(range(0, len(serie), pas))
    ax.set_xticklabels([d.strftime("%d/%m") for d in serie.index[::pas]], fontsize=5.5, color="#52514e")
    ax.tick_params(axis="y", labelsize=5.5, colors="#52514e", length=0)
    ax.tick_params(axis="x", length=0)
    ax.grid(axis="y", color="#e1e0d9", linewidth=0.5, zorder=0)
    for c in ("top", "right", "left"):
        ax.spines[c].set_visible(False)
    ax.spines["bottom"].set_color("#c3c2b7")
    ax.set_ylim(0, max(serie.max(), 1) * 1.22)
    ax.set_title("Fiches par jour de collecte", fontsize=6.5, color="#16324f", loc="left", pad=3)
    fig.tight_layout(pad=0.3)
    tampon = io.BytesIO()
    fig.savefig(tampon, format="png", dpi=220)
    plt.close(fig)
    tampon.seek(0)
    return tampon


# =====================================================================
# Points d'attention et actions par agent
# =====================================================================
def _points_attention(fd, qd, ed, P) -> list[tuple[str, str, bool]]:
    """Liste (rubrique, constat, problème ?) – toujours les mêmes rubriques, dans le même ordre."""
    n = len(fd)
    res = []
    al = fd[fd["Nb_alertes"] > 0]
    if len(al):
        types = (al["Alertes"].str.split(" ; ").explode().str.replace(r"\s*\(.*\)$", "", regex=True)
                 .value_counts().head(3))
        res.append(("Alertes qualité", f"{len(al)} fiche(s) sur {n} ({_fmt(_pct(len(al), n), 1)} %) – principales : "
                    + ", ".join(f"{k.lower()} ({v})" for k, v in types.items()), True))
    else:
        res.append(("Alertes qualité", "aucune fiche avec alerte", False))

    tar = ed["tardifs"]
    res.append(("Envois tardifs", (f"{len(tar)} fiche(s) envoyée(s) après minuit, retard médian "
                                   f"{_fmt(tar['Retard (h)'].median())} h – {_liste_noms(sorted(tar['Agent'].unique()))}")
                if len(tar) else "toutes les fiches ont été envoyées le jour de la collecte", bool(len(tar))))
    sil = ed["silences"]
    if not ed.get("silence_actif", True):
        res.append(("Agents sans envoi", "contrôle non appliqué (collecte terminée ou suspendue)", False))
    else:
        res.append(("Agents sans envoi", (f"{len(sil)} agent(s) sans envoi depuis plus de {int(P['SILENCE_H'])} h – "
                                          + _liste_noms(sil["Agent"].tolist())) if len(sil)
                    else f"aucun agent silencieux depuis plus de {int(P['SILENCE_H'])} h", bool(len(sil))))
    dep = qd[qd["Statut"] == STATUT_DEPASSEMENT]
    quo = qd[qd["Statut"] == STATUT_QUOTA]
    res.append((f"Règle des {int(P['QUOTA_PV_PRODUIT'])} PV",
                (f"{len(dep)} dépassement(s) (produit coché dans plus de {int(P['QUOTA_PV_PRODUIT'])} PV) – "
                 f"{_liste_noms(sorted(dep['Agent'].unique()))} ; {len(quo)} produit(s) au quota")
                if len(dep) else (f"aucun dépassement ; {len(quo)} produit(s) ont atteint le quota" if len(quo)
                                  else "aucun produit au quota"), bool(len(dep))))
    g = int(fd["Alerte_gps"].sum())
    res.append(("GPS imprécis ou absent", f"{g} fiche(s) ({_fmt(_pct(g, n), 1)} %) – précision attendue ≤ "
                f"{int(P['PRECISION_GPS_MAX_M'])} m" if g else "toutes les coordonnées sont fiables", bool(g)))
    d = int(fd["Alerte_duree"].sum())
    res.append(("Durées anormales", f"{d} entretien(s) hors de la plage {int(P['DUREE_MIN_MINUTES'])}–"
                f"{int(P['DUREE_MAX_MINUTES'])} min" if d else "toutes les durées sont conformes", bool(d)))
    nv = int((fd["Statut"] == "Non vérifié").sum())
    nv24 = int(fd["Hors_delai_validation"].sum())
    na = int((fd["Statut"] == "Non approuvé").sum())
    res.append(("Fiches non validées", (f"{nv} fiche(s) non vérifiée(s), dont {nv24} depuis plus de "
                                        f"{int(P['DELAI_VALIDATION_H'])} h ; {na} non approuvée(s)")
                if nv or na else "toutes les fiches ont été validées", bool(nv24 or na)))
    h = ed["horloge"]
    res.append(("Heure de tablette suspecte", (f"{len(h)} fiche(s) – {_liste_noms(sorted(h['Agent'].unique()))}")
                if len(h) else "aucune anomalie d'horloge", bool(len(h))))
    r = ed["rouvertes"]
    res.append(("Fiches rouvertes", (f"{len(r)} fiche(s) terminée(s) un autre jour – "
                                     f"{_liste_noms(sorted(r['Agent'].unique()))}")
                if len(r) else "aucune fiche modifiée après la collecte", bool(len(r))))
    return res


def _actions_agents(fd, qd, ed, P, maxi=12) -> tuple[list[tuple[str, list[str]]], int]:
    """Actions concrètes par agent, triées par nombre de problèmes. Renvoie (liste, nb d'agents non affichés)."""
    actions = []
    for agent, g in fd.groupby("Agent"):
        a = []
        q = qd[qd["Agent"] == agent]
        dep = q[q["Statut"] == STATUT_DEPASSEMENT]
        if len(dep):
            prods = ", ".join(f"« {p} »" for p in dep["Produit"].head(3)) + (" …" if len(dep) > 3 else "")
            a.append(f"retirer les PV en trop pour {len(dep)} produit(s) cochés dans plus de "
                     f"{int(P['QUOTA_PV_PRODUIT'])} PV ({prods})")
        nq = int((q["Statut"] == STATUT_QUOTA).sum())
        if nq:
            a.append(f"ne plus cocher {nq} produit(s) déjà au quota (liste Excel de l'onglet Produits)")
        sil = ed["silences"]
        s = sil[sil["Agent"] == agent]
        if len(s):
            a.append(f"le contacter : aucun envoi depuis le {fdate(s['Dernier envoi'].iat[0])}")
        nt = int((ed["tardifs"]["Agent"] == agent).sum())
        if nt:
            a.append(f"envoyer chaque soir avant minuit ({nt} fiche(s) envoyée(s) en retard)")
        nh = int((ed["horloge"]["Agent"] == agent).sum())
        if nh:
            a.append(f"régler la date et l'heure automatiques de la tablette ({nh} fiche(s) suspecte(s))")
        nr = int((ed["rouvertes"]["Agent"] == agent).sum())
        if nr:
            a.append(f"justifier {nr} fiche(s) modifiée(s) après la collecte")
        ng = int(g["Alerte_gps"].sum())
        if ng:
            a.append(f"reprendre le GPS de {ng} PV (attendre une précision ≤ {int(P['PRECISION_GPS_MAX_M'])} m)")
        nd = int(g["Alerte_duree"].sum())
        if nd:
            a.append(f"vérifier {nd} durée(s) d'entretien anormale(s)")
        nc = int(g["Alerte_completude"].sum())
        if nc:
            a.append(f"compléter {nc} fiche(s) incomplète(s) (nom, type, téléphone…)")
        nb = int(g["Alerte_doublon"].sum())
        if nb:
            a.append(f"vérifier {nb} doublon(s) possible(s) et faire supprimer les fiches en trop")
        na = int((g["Statut"] == "Non approuvé").sum())
        if na:
            a.append(f"corriger {na} fiche(s) non approuvée(s) par le superviseur")
        if a:
            actions.append((agent, a))
    actions.sort(key=lambda x: (-len(x[1]), x[0]))
    return actions[:maxi], max(0, len(actions) - maxi)


# =====================================================================
# Rapport
# =====================================================================
def generer_rapport(f: pd.DataFrame, maintenant: pd.Timestamp, perimetre: str, nb_agents_prevus: int = 0,
                    quota: pd.DataFrame | None = None, envoi: dict | None = None,
                    parametres: dict | None = None) -> bytes:
    """Rapport Word de fin de collecte.
    f        : fiches filtrées (sortie de construire_suivi, avec les filtres actifs)
    quota    : sortie de controle_quota_produits (recalculée si absente)
    envoi    : sortie de alertes_envoi (recalculée si absente)
    """
    P = {**PARAMETRES_DEFAUT, **(parametres or {})}
    quota = quota if quota is not None else controle_quota_produits(f, int(P["QUOTA_PV_PRODUIT"]))
    envoi = envoi if envoi is not None else alertes_envoi(f, f, maintenant, P)

    doc = Document()
    s = doc.sections[0]
    s.orientation = WD_ORIENT.PORTRAIT
    s.page_width, s.page_height = Cm(21), Cm(29.7)
    s.left_margin = s.right_margin = Cm(1.6)
    s.top_margin, s.bottom_margin = Cm(1.2), Cm(1.3)
    st = doc.styles["Normal"]
    st.font.name = "Calibri"
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    st.font.size = Pt(9.5)
    st.paragraph_format.space_after = Pt(3)

    # Pied de page : crédit + numéro de page
    pied = s.footer.paragraphs[0]
    pied.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(pied, "Rapport de fin de collecte INHPC · INS – DSE · Financement HISWACA · Application réalisée par "
               "OGNIMBA Sadri · Page ", taille=7.5, couleur=GRIS)
    _champ(pied, "PAGE")
    _run(pied, " / ", taille=7.5, couleur=GRIS)
    _champ(pied, "NUMPAGES")

    if f.empty:
        _para(doc, "Aucune fiche pour le périmètre choisi.", taille=11, gras=True, couleur=BLEU)
        tampon = io.BytesIO()
        doc.save(tampon)
        return tampon.getvalue()

    debut, fin = f["Jour_collecte"].min(), f["Jour_collecte"].max()
    jours = pd.date_range(debut, fin)
    perf = performance_departements(f)
    n_dep = len(perf)

    # =================================================================
    # PAGE 1 – Page de garde et synthèse nationale
    # =================================================================
    h = doc.add_table(rows=1, cols=3)
    _bordures(h, aucune=True)
    _largeurs(h, (4.0, 9.8, 4.0))
    for j, nom, haut in ((0, "logo_ins.png", 2.8), (2, "logo_hiswaca.png", 2.3)):
        chemin = os.path.join(DOSSIER, "assets", nom)
        p = h.rows[0].cells[j].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if os.path.exists(chemin):
            p.add_run().add_picture(chemin, height=Cm(haut))
        h.rows[0].cells[j].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    c = h.rows[0].cells[1]
    c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for txt, taille, coul, gras in (("RÉPUBLIQUE DU CONGO", 8.5, GRIS, True),
                                    ("Institut National de la Statistique", 9, GRIS, False),
                                    ("Direction des Statistiques Économiques", 9, GRIS, False)):
        p = c.paragraphs[0] if txt.startswith("RÉP") else c.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        _run(p, txt, taille=taille, couleur=coul, gras=gras)

    _para(doc, apres=10)
    p = _para(doc, align=WD_ALIGN_PARAGRAPH.CENTER, apres=2)
    _run(p, "RAPPORT DE FIN DE COLLECTE", taille=22, gras=True, couleur=BLEU)
    p = _para(doc, align=WD_ALIGN_PARAGRAPH.CENTER, apres=6)
    _run(p, "Enquête sur l'identification des points de vente et des produits élémentaires (INHPC)",
         taille=12, couleur=ACCENT)
    p = _para(doc, align=WD_ALIGN_PARAGRAPH.CENTER, apres=1)
    _run(p, f"Période de collecte : du {debut:%d/%m/%Y} au {fin:%d/%m/%Y} ({len(jours)} jours)",
         taille=10.5, gras=True, couleur=BLEU)
    p = _para(doc, align=WD_ALIGN_PARAGRAPH.CENTER, apres=12)
    _run(p, f"Situation au {maintenant:%d/%m/%Y à %H:%M} · Périmètre : {perimetre}", taille=9, couleur=GRIS)

    _titre(doc, "Synthèse nationale – chiffres clés")
    n = len(f)
    n_pv = f["PV_cle"].nunique()
    n_app = int((f["Statut"] == "Approuvé").sum())
    n_al = int((f["Nb_alertes"] > 0).sum())
    agents = f["Agent"].nunique()
    _tuiles(doc, [
        ("PV identifiés (distincts)", _fmt(n_pv)),
        ("Fiches reçues", _fmt(n)),
        ("Agents actifs", f"{agents}" + (f" / {nb_agents_prevus}" if nb_agents_prevus else "")),
        ("Départements couverts", _fmt(n_dep)),
        ("Localités couvertes", _fmt(f["Localite"].nunique())),
        ("Fiches approuvées", f"{_fmt(_pct(n_app, n), 1)} %"),
        ("Fiches avec alerte", f"{_fmt(_pct(n_al, n), 1)} %"),
        ("Envois tardifs", _fmt(len(envoi["tardifs"]))),
    ], par_ligne=4)

    _titre(doc, "Classement des départements – score de performance /100")
    lignes = []
    for _, r in perf.iterrows():
        fd = f[f["Departement"] == r["Département"]]
        lignes.append([r["Rang"], r["Département"], _fmt(fd["PV_cle"].nunique()), _fmt(len(fd)), r["Agents"],
                       _fmt(r["PV / agent / jour"], 1),
                       _fmt(_pct((fd["Statut"] == "Approuvé").sum(), len(fd)), 1) + " %",
                       _fmt(_pct((fd["Nb_alertes"] > 0).sum(), len(fd)), 1) + " %",
                       _fmt(r["Score /100"], 1)])
    _tableau(doc, ["Rang", "Département", "PV distincts", "Fiches", "Agents", "PV / agent / jour",
                   "Approuvées", "Avec alerte", "Score /100"], lignes,
             [1.1, 4.2, 1.8, 1.5, 1.5, 2.0, 1.9, 1.9, 1.9])
    _para(doc, "Score /100 : productivité 20 %, complétude 20 %, fiabilité GPS 15 %, durée conforme 15 %, "
               "absence de doublon 10 %, validation dans les 24 h 20 % (mêmes critères que l'onglet 🏆 Localités). "
               "« PV distincts » : les fiches en double d'un même point de vente ne comptent qu'une fois.",
          taille=7.5, italique=True, couleur=GRIS, avant=2)

    _titre(doc, "Évolution nationale de la collecte")
    doc.add_picture(_graphique_jours(f, jours, largeur_cm=17.0, hauteur_cm=4.2), width=Cm(17.0))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # =================================================================
    # UNE PAGE PAR DÉPARTEMENT
    # =================================================================
    for _, r in perf.sort_values("Département").iterrows():
        dep = r["Département"]
        fd = f[f["Departement"] == dep]
        agents_dep = set(fd["Agent"])
        qd = quota[quota["Agent"].isin(agents_dep)]
        ed = {k: (v[v["Département"] == dep] if isinstance(v, pd.DataFrame) else v) for k, v in envoi.items()}

        # Bandeau du département
        p = _para(doc, apres=1)
        p.paragraph_format.page_break_before = True
        _run(p, f"DÉPARTEMENT : {dep.upper()}", taille=16, gras=True, couleur=BLEU)
        sup = sorted(x for x in fd["Superviseur"].unique() if x)
        locs = sorted(fd["Localite"].unique())
        p = _para(doc, apres=4)
        _run(p, f"Localité(s) : {_liste_noms(locs, 6)}" + (f" · Superviseur(s) : {', '.join(sup)}" if sup else "")
             + f" · Collecte du {fd['Jour_collecte'].min():%d/%m} au {fd['Jour_collecte'].max():%d/%m/%Y}",
             taille=8.5, couleur=GRIS)

        # 1. Chiffres clés
        _titre(doc, "1. Chiffres clés")
        n_d = len(fd)
        _tuiles(doc, [
            ("PV identifiés", _fmt(fd["PV_cle"].nunique())),
            ("Fiches", _fmt(n_d)),
            ("Agents actifs", _fmt(fd["Agent"].nunique())),
            ("Localités", _fmt(len(locs))),
            ("Fiches approuvées", f"{_fmt(_pct((fd['Statut'] == 'Approuvé').sum(), n_d), 1)} %"),
            ("Fiches avec alerte", f"{_fmt(_pct((fd['Nb_alertes'] > 0).sum(), n_d), 1)} %"),
        ])

        # 2. Performance : critères à gauche, graphique à droite
        _titre(doc, f"2. Performance – score {_fmt(r['Score /100'], 1)} / 100 · rang {r['Rang']} sur {n_dep}")
        cadre = doc.add_table(rows=1, cols=2)
        _bordures(cadre, aucune=True)
        _largeurs(cadre, (8.4, 9.4))
        gauche, droite = cadre.rows[0].cells
        gauche.vertical_alignment = droite.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        crit = [[k, _fmt(r[k], 1) + " %", f"{w} %"] for k, w in POIDS_SCORE.items()]
        crit.append(["Score /100", _fmt(r["Score /100"], 1), "100 %"])
        gauche.paragraphs[0].paragraph_format.space_after = Pt(0)
        _tableau(gauche, ["Critère", "Valeur", "Poids"], crit, [3.6, 2.2, 1.6], derniere_total=True, taille=8)
        _para(gauche, f"PV / agent / jour : {_fmt(r['PV / agent / jour'], 1)} · durée médiane : "
                      f"{_fmt(r['Durée médiane (min)'])} min", taille=7.5, couleur=GRIS, avant=2, apres=0)
        pp = droite.paragraphs[0]
        pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pp.add_run().add_picture(_graphique_jours(fd, jours), width=Cm(9.0))

        # 3. Points d'attention (encadré)
        _titre(doc, "3. Points d'attention")
        boite = doc.add_table(rows=1, cols=1)
        _bordures(boite, aucune=True)
        _largeurs(boite, (LARGEUR_UTILE,))
        cel = boite.rows[0].cells[0]
        _fond(cel, "FDF3EE")
        _bord_cellule(cel, "left", "D03B3B", 30)
        premier = True
        for rubrique, constat, probleme in _points_attention(fd, qd, ed, P):
            p = cel.paragraphs[0] if premier else cel.add_paragraph()
            premier = False
            p.paragraph_format.space_after = Pt(1.5)
            p.paragraph_format.left_indent = Cm(0.15)
            _run(p, ("▲ " if probleme else "✓ "), taille=8.5, gras=True, couleur=ROUGE if probleme else VERT)
            _run(p, f"{rubrique} : ", taille=8.5, gras=True, couleur=BLEU)
            _run(p, constat, taille=8.5, couleur=None if probleme else GRIS)

        # 4. Éléments à revoir
        _titre(doc, "4. Éléments à revoir")
        actions, reste = _actions_agents(fd, qd, ed, P)
        nv = int((fd["Statut"] == "Non vérifié").sum())
        if nv:
            qui = ", ".join(sup) if sup else "Superviseur"
            _puce(doc, morceaux=[(f"{qui} : ", True, BLEU),
                                 (f"valider les {nv} fiche(s) non vérifiée(s) dans KoboToolbox "
                                  f"(dont {int(fd['Hors_delai_validation'].sum())} depuis plus de "
                                  f"{int(P['DELAI_VALIDATION_H'])} h).", False, None)], taille=8.5)
        for agent, a in actions:
            _puce(doc, morceaux=[(f"{agent} : ", True, BLEU), ("; ".join(a) + ".", False, None)], taille=8.5)
        if reste:
            _para(doc, f"… et {reste} autre(s) agent(s) avec des points mineurs (voir l'application).",
                  taille=8, italique=True, couleur=GRIS)
        if not actions and not nv:
            _para(doc, "Aucune action particulière : le travail du département est conforme.", taille=9,
                  couleur=VERT)

    # =================================================================
    # DERNIÈRE PAGE – Recommandations générales
    # =================================================================
    _titre(doc, "Recommandations générales", saut=True)
    reco = []
    tar, hor, rou, sil = envoi["tardifs"], envoi["horloge"], envoi["rouvertes"], envoi["silences"]
    dep_q = quota[quota["Statut"] == STATUT_DEPASSEMENT]
    if len(dep_q) or len(quota):
        reco.append(("Règle des 3 PV par produit",
                     f"Remettre à chaque agent la feuille Excel de sa localité (onglet 🛒 Produits) : "
                     f"{len(quota[quota['Statut'] == STATUT_QUOTA])} produit(s) sont au quota et ne doivent plus "
                     f"être cochés ; {len(dep_q)} dépassement(s) sont à corriger en retirant les PV en trop."))
    if len(tar):
        reco.append(("Envoi le jour même",
                     f"{len(tar)} fiche(s) ont été envoyées après minuit. Rappeler que chaque fiche doit être "
                     "envoyée au serveur le jour de la collecte, avant 00h00, et prévoir un point de connexion "
                     "en fin de journée."))
    if len(sil):
        reco.append(("Suivi quotidien des agents",
                     f"{len(sil)} agent(s) n'ont rien envoyé depuis plus de {int(P['SILENCE_H'])} h : "
                     "les appeler pour connaître la cause (panne, connexion, absence)."))
    if len(hor):
        reco.append(("Réglage des tablettes",
                     f"{len(hor)} fiche(s) présentent une heure de tablette suspecte. Activer « date et heure "
                     "automatiques » et le fuseau Africa/Brazzaville (UTC+1) sur toutes les tablettes."))
    if len(rou):
        reco.append(("Fiches modifiées après la collecte",
                     f"{len(rou)} fiche(s) ont été terminées un autre jour que celui de leur ouverture. "
                     "Demander une justification et éviter de rouvrir les fiches déjà enregistrées."))
    g = int(f["Alerte_gps"].sum())
    if g:
        reco.append(("Qualité du GPS",
                     f"{g} fiche(s) ont un GPS absent ou imprécis : attendre une précision ≤ "
                     f"{int(P['PRECISION_GPS_MAX_M'])} m, à l'extérieur du bâtiment, avant d'enregistrer."))
    nv24 = int(f["Hors_delai_validation"].sum())
    if nv24:
        reco.append(("Validation par les superviseurs",
                     f"{nv24} fiche(s) attendent la validation depuis plus de {int(P['DELAI_VALIDATION_H'])} h : "
                     "valider chaque jour les fiches de la veille."))
    d = int(f["Alerte_duree"].sum())
    if d:
        reco.append(("Durée des entretiens",
                     f"{d} entretien(s) ont une durée anormale : vérifier les fiches trop courtes (saisie "
                     "hors du PV ?) et trop longues (fiche laissée ouverte)."))
    if len(perf) > 1:
        bas = perf.sort_values("Score /100").iloc[0]
        reco.append(("Appui ciblé",
                     f"Le département {bas['Département']} a le score le plus faible "
                     f"({_fmt(bas['Score /100'], 1)}/100) : y organiser en priorité une supervision de terrain."))
    reco.append(("Apurement final",
                  "Avant la clôture : supprimer les doublons confirmés, compléter les fiches incomplètes, "
                  "et archiver la base Kobo complète (onglet 📋 Données)."))
    for k, (titre, texte) in enumerate(reco, start=1):
        p = _para(doc, apres=4)
        _run(p, f"{k}. {titre}. ", taille=9.5, gras=True, couleur=BLEU)
        _run(p, texte, taille=9.5)

    _titre(doc, "Rappel des règles et seuils utilisés")
    regles = [
        ["Règle des 3 PV", f"Un agent ne coche pas un même produit dans plus de {int(P['QUOTA_PV_PRODUIT'])} PV "
                           "distincts (une fiche en double ne compte qu'une fois)."],
        ["Envoi tardif", "Date de réception par le serveur postérieure au jour de collecte (start) ; retard "
                         "compté à partir de minuit."],
        ["Agent sans envoi", f"Aucun formulaire reçu depuis plus de {int(P['SILENCE_H'])} h."],
        ["Heure de tablette", f"Écart de plus de {int(P['ECART_HORLOGE_H'])} h entre « end » et la réception le "
                              "même jour, heure de la tablette en avance, ou date « start » dans le futur."],
        ["Fiche rouverte", f"« start » et « end » sur des jours différents, avec plus de "
                           f"{int(P['REOUVERTURE_H'])} h d'écart."],
        ["Durée", f"Entretien de moins de {int(P['DUREE_MIN_MINUTES'])} min ou de plus de "
                  f"{int(P['DUREE_MAX_MINUTES'])} min."],
        ["GPS", f"Coordonnées absentes, hors du Congo ou précision > {int(P['PRECISION_GPS_MAX_M'])} m."],
        ["Validation", f"Fiche toujours « non vérifiée » plus de {int(P['DELAI_VALIDATION_H'])} h après réception."],
    ]
    _tableau(doc, ["Contrôle", "Définition"], regles, [3.6, 14.2], taille=8.5)
    _para(doc, "Une alerte n'est pas forcément une erreur : c'est un point à vérifier avec l'agent.",
          taille=8, italique=True, couleur=GRIS, avant=4)

    tampon = io.BytesIO()
    doc.save(tampon)
    return tampon.getvalue()
