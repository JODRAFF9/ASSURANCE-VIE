# -*- coding: utf-8 -*-
"""Génère le questionnaire d'enquête (Word) ayant permis de constituer la base."""
import os

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

HERE = os.path.dirname(os.path.abspath(__file__))

doc = Document()
for section in doc.sections:
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)

style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(11)


def title(text, size=16, center=True, space_after=6):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(size)
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(space_after)
    return p


def h(text, size=13):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(size)
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)


def q(num, text, lines=1):
    p = doc.add_paragraph()
    run = p.add_run(f"Q{num}. ")
    run.bold = True
    p.add_run(text)
    for _ in range(lines):
        doc.add_paragraph("…" * 55)


def case(text):
    doc.add_paragraph("☐  " + text)


# ---------------------------------------------------------------------------
title("ENSAE Pierre Ndiaye — ISE 3 — Méthodes actuarielles / Assurance vie", 11)
title("PROJET I — ENQUÊTE AUPRÈS DES INSTITUTIONS D'ASSURANCE", 15)
title("Questionnaire de recueil des produits d'assurance vie et de collecte des "
      "données de tarification", 12)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Enquêteur : Sié Rachid Traoré — Année académique 2025-2026")
r.italic = True

doc.add_paragraph(
    "Objet : recenser les produits d'assurance vie commercialisés sur la place de "
    "Dakar, identifier le produit le plus commercialisé et collecter les éléments "
    "techniques ainsi que les données de personnel nécessaires à la construction "
    "d'une note technique et d'un simulateur de tarification."
)
doc.add_paragraph(
    "Confidentialité : les informations recueillies sont strictement anonymisées et "
    "utilisées à des fins pédagogiques uniquement."
)

# ---------------------------------------------------------------------------
h("VOLET A — Identification de l'institution", 13)
q(1, "Dénomination de la compagnie :")
q(2, "Statut (SA, mutuelle, filiale de groupe régional…) :")
q(3, "Année d'agrément au Sénégal (branche vie) :")
q(4, "Fonction de la personne interrogée (actuaire, souscripteur, directeur technique…) :")

h("VOLET B — Portefeuille de produits vie", 13)
p = doc.add_paragraph()
r = p.add_run("Q5. ")
r.bold = True
p.add_run("Quels produits d'assurance vie commercialisez-vous ? (cocher)")
for lbl in [
    "Indemnités de Fin de Carrière (IFC) / passif social entreprises",
    "IFC Plus (couverture décès du personnel — indemnités CCNI)",
    "Temporaire décès individuelle",
    "Assurance décès emprunteur (crédit bancaire)",
    "Épargne / capitalisation (capital différé)",
    "Retraite complémentaire par capitalisation",
    "Rente éducation / rente de conjoint",
    "Mixte (épargne + décès)",
    "Autre (préciser) : ………………………………",
]:
    case(lbl)
q(6, "Classez vos trois produits les plus commercialisés par chiffre d'affaires "
     "(primes émises) du dernier exercice :", 3)
q(7, "Part approximative du produit n°1 dans les primes vie de la compagnie (%) :")
q(8, "Quelle clientèle porte ce produit (grandes entreprises, PME, institutions, "
     "particuliers) ?")

h("VOLET C — Caractéristiques techniques du produit le plus commercialisé (IFC)", 13)
q(9, "Quelle convention collective sert de référence au calcul des droits ?")
q(10, "Quelle table de mortalité utilisez-vous ?  (CIMA H / CIMA F / autre)")
q(11, "Quel taux technique (taux d'actualisation) appliquez-vous ? (max 3,5 % — art. 338 code CIMA)")
q(12, "Quel taux d'évolution des salaires retenez-vous par défaut ?")
q(13, "Quel taux de turn-over (rotation du personnel) retenez-vous par défaut ?")
q(14, "Quel taux annuel de licenciement retenez-vous par défaut ?")
q(15, "Quels chargements de gestion appliquez-vous (en % des primes) ?")
q(16, "Quel âge de départ à la retraite retenez-vous (référence IPRES) ?")
q(17, "Quelles modalités de financement proposez-vous à l'entreprise "
      "(prime unique, dette actuarielle + charge normale, amortissement…) ?", 2)
q(18, "Le fonds constitué participe-t-il aux bénéfices financiers et techniques ? "
      "À quelles hauteurs ?")
q(19, "À quelle fréquence recommandez-vous la réactualisation de l'étude actuarielle ?")

h("VOLET D — Fiche de collecte des données du personnel (entreprise souscriptrice)", 13)
doc.add_paragraph(
    "Cette fiche est remise à l'entreprise souhaitant souscrire le contrat IFC. "
    "Les informations ci-dessous sont demandées pour CHAQUE salarié ; elles "
    "constituent la base de données utilisée par le simulateur (feuille « BD »)."
)
tbl = doc.add_table(rows=1, cols=2)
tbl.style = "Table Grid"
hdr = tbl.rows[0].cells
hdr[0].text = "Champ demandé"
hdr[1].text = "Format / précision"
for champ, fmt in [
    ("Matricule", "identifiant anonyme"),
    ("Nom et prénoms", "facultatif (peut être anonymisé)"),
    ("Sexe", "M / F"),
    ("Catégorie professionnelle", "cadre dirigeant, cadre, agent de maîtrise, employé, ouvrier"),
    ("Date de naissance", "JJ/MM/AAAA"),
    ("Date d'embauche", "JJ/MM/AAAA"),
    ("Salaire mensuel brut", "FCFA, bonus compris"),
    ("Nombre de mois de salaire par an", "12 ou 13"),
]:
    row = tbl.add_row().cells
    row[0].text = champ
    row[1].text = fmt
for row in tbl.rows:
    for cell in row.cells:
        for par in cell.paragraphs:
            for run in par.runs:
                run.font.size = Pt(10)

doc.add_paragraph()
q(20, "Effectif total de l'entreprise à la date d'évaluation :")
q(21, "Masse salariale annuelle brute (FCFA) :")
q(22, "Date d'évaluation souhaitée (clôture d'exercice) :")
q(23, "Existe-t-il un fonds IFC déjà constitué ? Si oui, quel montant ?")

doc.add_paragraph()
p = doc.add_paragraph()
r = p.add_run(
    "NB : l'enquête n'ayant pas pu être réalisée sur le terrain, les réponses au "
    "présent questionnaire ont été simulées de manière réaliste (8 compagnies vie "
    "fictives ; base de personnel de 50 salariés générée pour l'entreprise "
    "SENIA S.A.). Le produit ressorti comme le plus commercialisé est le contrat "
    "d'Indemnités de Fin de Carrière (IFC), objet de la note technique."
)
r.italic = True

out = os.path.join(HERE, "Questionnaire_enquete_IFC.docx")
doc.save(out)
print("Questionnaire écrit :", out)
