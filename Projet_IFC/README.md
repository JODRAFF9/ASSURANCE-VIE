# Projet I — Note technique IFC et simulateur Excel

**Énoncé** : faire une enquête auprès des institutions d'assurance afin de recueillir les
produits et faire une note technique du produit le plus commercialisé.

L'enquête n'ayant pas été réalisée sur le terrain, les réponses ont été **simulées** à
partir du questionnaire : le produit ressorti comme le plus commercialisé est le contrat
d'**Indemnités de Fin de Carrière (IFC / IFC+)**, évalué ici pour une entreprise
sénégalaise fictive, **SENIA S.A.** (50 salariés, base simulée).

## Livrables

| Fichier | Description |
|---|---|
| `Note_Technique_IFC_SENIA.Rmd` | Note technique du produit IFC/IFC+ (à tricoter en PDF avec RStudio, même modèle que la note précédente ; les images sont dans `ressource/image/`) |
| `simulateur_IFC_SENIA.xlsx` | Simulateur Excel **sans couleur**, avec formules vivantes ; la base simulée est la feuille `BD` |
| `Questionnaire_enquete_IFC.docx` | Questionnaire d'enquête (volets A–D) ayant permis de recueillir les produits et de constituer la base |

## Scripts de construction (reproductibilité)

- `build_simulateur.py` : génère le classeur Excel (base simulée seed fixe, formules,
  table CIMA H, barème CCNI) et un miroir Python des calculs (`synthese.json`).
- `build_questionnaire.py` : génère le questionnaire Word.

## Méthodologie (source : `Evaluation/doc/Cours/FICHE- IFC-TP.pdf`)

Méthode prospective : droits projetés au terme (salaire final × barème CCNI), pondérés
par la probabilité de survie (table CIMA H) et actualisés au taux technique de 3,5 %
(`v^(n+0,5)`), puis répartis au prorata de l'ancienneté :

- Engagement global = Σ VAP ;
- Dette actuarielle = VAP × ancienneté / ancienneté au terme ;
- Charge normale = VAP / ancienneté au terme ;
- Passif IL = IL × τ_lic × (1 − u^n)/(1 − u), avec u = (1 − τ_rot)/(1 + i) ;
- IFC+ : prime pure = Σ capital décès × qx, prime commerciale = prime pure/(1 − g).

## Résultats (base simulée, 31/12/2025)

| Rubrique | Montant (FCFA) |
|---|---:|
| Masse salariale annuelle | 344 386 800 |
| Engagement global | 230 047 411 |
| Dette actuarielle | 111 103 527 |
| Charge normale | 8 223 109 |
| Passif indemnité de licenciement | 7 831 464 |
| Capitaux garantis IFC+ | 161 949 480 |
| Prime commerciale IFC+ | 1 203 354 |

Les formules du classeur ont été recalculées sous LibreOffice Calc : elles redonnent
exactement les valeurs du miroir Python.
