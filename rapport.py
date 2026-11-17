"""
Rapport Word (2 pages maximum) de la collecte INHPC
Application réalisée par OGNIMBA Sadri
"""
from __future__ import annotations

import io
import os

import pandas as pd
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BLEU = RGBColor(0x16, 0x32, 0x4F)
ACCENT = RGBColor(0x2A, 0x78, 0xD6)
GRIS = RGBColor(0x52, 0x51, 0x4E)
ROUGE = RGBColor(0x9D, 0x22, 0x22)
DOSSIER = os.path.dirname(os.path.abspath(__file__))


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


def _texte(cell, txt, gras=False, couleur=None, taille=8.5, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    if align:
        p.alignment = align
    r = p.add_run(str(txt))
    r.bold = gras
    r.font.size = Pt(taille)
    if couleur:
        r.font.color.rgb = couleur
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def _titre(doc, txt):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
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


def _tableau(doc, entetes, lignes, largeurs_cm, alignements=None, derniere_total=False):
    t = doc.add_table(rows=1, cols=len(entetes))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    _bordures(t)
    for j, h in enumerate(entetes):
        c = t.rows[0].cells[j]
        _fond(c, "16324F")
        _texte(c, h, gras=True, couleur=RGBColor(0xFF, 0xFF, 0xFF), taille=8,
               align=WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER)
    for i, ligne in enumerate(lignes):
        cells = t.add_row().cells
        total = derniere_total and i == len(lignes) - 1
        for j, v in enumerate(ligne):
            al = (alignements[j] if alignements else
                  (WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER))
            _texte(cells[j], v, gras=total, align=al)
            if total:
                _fond(cells[j], "E3ECF7")
            elif i % 2:
                _fond(cells[j], "F6F5F1")
    _largeurs(t, largeurs_cm)
    return t


def _puce(doc, txt):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(txt)
    r.font.size = Pt(9)


def _fmt(x, dec=0):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "–"
    return f"{x:,.{dec}f}".replace(",", " ").replace(".", ",")


def generer_rapport(f: pd.DataFrame, perf: pd.DataFrame, types: pd.DataFrame, valid: pd.DataFrame,
                    kpi: dict, perimetre: str, maintenant: pd.Timestamp, nb_agents_prevus: int = 0) -> bytes:
    doc = Document()
    s = doc.sections[0]
    s.orientation = WD_ORIENT.PORTRAIT
    s.page_width, s.page_height = Cm(21), Cm(29.7)
    s.left_margin = s.right_margin = Cm(1.6)
    s.top_margin, s.bottom_margin = Cm(1.2), Cm(1.2)
    st = doc.styles["Normal"]
    st.font.name = "Calibri"
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    st.font.size = Pt(9.5)
    st.paragraph_format.space_after = Pt(3)

    # ---------- En-tête : logos + titre ----------
    h = doc.add_table(rows=1, cols=3)
    _bordures(h, aucune=True)
    _largeurs(h, (3.2, 11.4, 3.2))
    for j, nom in ((0, "logo_ins.png"), (2, "logo_hiswaca.png")):
        chemin = os.path.join(DOSSIER, "assets", nom)
        if os.path.exists(chemin):
            p = h.rows[0].cells[j].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(chemin, height=Cm(2.0) if j == 0 else Cm(1.7))
    c = h.rows[0].cells[1]
    c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("RAPPORT DE SUIVI DE LA COLLECTE")
    r.bold, r.font.size, r.font.color.rgb = True, Pt(15), BLEU
    for txt, taille, coul in (
        ("Enquête sur l'identification des points de vente et des produits élémentaires (INHPC)", 9.5, GRIS),
        (f"Situation au {maintenant:%d/%m/%Y à %H:%M} · Périmètre : {perimetre}", 9, ACCENT),
    ):
        q = c.add_paragraph()
        q.alignment = WD_ALIGN_PARAGRAPH.CENTER
        q.paragraph_format.space_after = Pt(0)
        rr = q.add_run(txt)
        rr.font.size, rr.font.color.rgb = Pt(taille), coul

    # ---------- 1. Chiffres clés ----------
    _titre(doc, "1. Chiffres clés")
    cles = [
        ("PV identifiés", _fmt(kpi["total"])),
        ("Fiches reçues aujourd'hui", _fmt(kpi["auj"])),
        ("Agents actifs aujourd'hui", f"{kpi['agents_auj']} / {nb_agents_prevus or kpi['agents_tous']}"),
        ("Localités couvertes", _fmt(kpi["localites"])),
        ("Fiches avec alerte", f"{_fmt(kpi['alertes'])} ({_fmt(kpi['pct_alertes'], 1)} %)"),
        ("Fiches approuvées", f"{_fmt(kpi['approuvees'])} ({_fmt(kpi['pct_approuvees'], 1)} %)"),
    ]
    t = doc.add_table(rows=2, cols=len(cles))
    _bordures(t, aucune=True)
    for j, (lib, val) in enumerate(cles):
        _fond(t.rows[0].cells[j], "E3ECF7")
        _fond(t.rows[1].cells[j], "E3ECF7")
        _texte(t.rows[0].cells[j], lib, couleur=GRIS, taille=7.5, align=WD_ALIGN_PARAGRAPH.CENTER)
        _texte(t.rows[1].cells[j], val, gras=True, couleur=BLEU, taille=12, align=WD_ALIGN_PARAGRAPH.CENTER)

    _largeurs(t, [2.96] * len(cles))

    # ---------- 2. Performance par localité ----------
    _titre(doc, "2. Avancement et performance par localité")
    lignes = []
    v = valid.set_index("Localité") if len(valid) else pd.DataFrame()
    for _, r in perf.iterrows():
        nv24 = int(v.loc[r["Localité"], "Non vérifiées > 24 h"]) if r["Localité"] in v.index else 0
        lignes.append([
            r["Rang"], f"{r['Localité']} ({r['Superviseur']})" if r["Superviseur"] else r["Localité"],
            r["PV identifiés"], r["Agents"], _fmt(r["PV / agent / jour"], 1), _fmt(r["Complétude"]) + " %",
            _fmt(r["Fiabilité GPS"]) + " %", _fmt(r["Durée conforme"]) + " %", _fmt(r["Validation"]) + " %",
            nv24, _fmt(r["Score /100"], 1),
        ])
    _tableau(doc, ["Rg", "Localité (superviseur)", "PV", "Agents", "PV/agent/jour", "Complét.", "GPS fiable",
                   "Durée conf.", "Validées", "Non vérif. > 24 h", "Score /100"], lignes,
             [0.7, 4.1, 1.0, 1.25, 1.4, 1.4, 1.3, 1.3, 1.45, 1.5, 1.4])
    n = doc.add_paragraph()
    n.paragraph_format.space_before = Pt(2)
    rr = n.add_run("Score /100 : productivité 20 %, complétude 20 %, fiabilité GPS 15 %, durée conforme 15 %, "
                   "absence de doublon 10 %, validation dans les 24 h 20 %. Une localité de moins de 10 fiches "
                   "est à interpréter avec prudence.")
    rr.italic, rr.font.size, rr.font.color.rgb = True, Pt(7.5), GRIS

    # ---------- 3. Types de points de vente ----------
    _titre(doc, "3. Répartition des points de vente par type")
    lignes = [[r["Type de point de vente"], _fmt(r["Nbre de points de vente"]), _fmt(r["Pourcentage"], 1)]
              for _, r in types.iterrows()]
    _tableau(doc, ["Type de point de vente", "Nbre de points de vente", "Pourcentage (%)"], lignes,
             [11.0, 3.4, 3.4], derniere_total=True)

    # ---------- 4. Points d'attention ----------
    _titre(doc, "4. Points d'attention")
    if len(perf):
        bas = perf.sort_values("Score /100").iloc[0]
        haut = perf.iloc[0]
        _puce(doc, f"Meilleure localité : {haut['Localité']} (score {_fmt(haut['Score /100'], 1)}/100). "
                   f"Localité à appuyer en priorité : {bas['Localité']} (score {_fmt(bas['Score /100'], 1)}/100).")
    nv = int(f["Hors_delai_validation"].sum())
    if nv:
        pire = valid.sort_values("Non vérifiées > 24 h", ascending=False).head(3)
        pire = ", ".join(f"{r['Localité']} ({int(r['Non vérifiées > 24 h'])})" for _, r in pire.iterrows()
                         if r["Non vérifiées > 24 h"] > 0)
        _puce(doc, f"{nv} fiche(s) attendent la validation du superviseur depuis plus de 24 h, surtout à : {pire}.")
    else:
        _puce(doc, "Toutes les fiches de plus de 24 h ont été traitées par les superviseurs.")
    al = f.loc[f["Nb_alertes"] > 0, "Alertes"].str.split(" ; ").explode().str.replace(r"\s*\(.*\)$", "", regex=True)
    if len(al):
        top = al.value_counts().head(3)
        _puce(doc, "Alertes les plus fréquentes : " + ", ".join(f"{k.lower()} ({v})" for k, v in top.items()) + ".")
    if pd.notna(f["Duree_min"].median()):
        _puce(doc, f"Durée médiane d'un entretien : {_fmt(f['Duree_min'].median())} minutes.")
    if nb_agents_prevus:
        _puce(doc, f"{kpi['agents_auj']} agent(s) sur {nb_agents_prevus} ont transmis au moins une fiche aujourd'hui ; "
                   f"{max(nb_agents_prevus - kpi['agents_tous'], 0)} agent(s) n'ont encore transmis aucune fiche.")

    # ---------- Pied de page ----------
    pied = s.footer.paragraphs[0]
    pied.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rr = pied.add_run("Rapport généré automatiquement par l'application de suivi INHPC · INS – DSE · "
                      "Financement HISWACA · Application réalisée par OGNIMBA Sadri")
    rr.font.size, rr.font.color.rgb = Pt(7.5), GRIS

    tampon = io.BytesIO()
    doc.save(tampon)
    return tampon.getvalue()
