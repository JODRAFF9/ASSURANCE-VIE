# Projet I : Note technique IFC et simulateur Excel

**Cours d'Assurance Vie, ENSAE Pierre Ndiaye, ISE 3, année académique 2025-2026.**
Groupe 1 : Gado Giovanni Jocelyn, Alagbe Abdou Hamid, Sié Rachid Traoré, Cheikh Sadibou Ngom.
Sous la supervision de Mme Fatou Diop.

**Énoncé** : faire une enquête auprès des institutions d'assurance afin de recueillir les
produits et faire une note technique du produit le plus commercialisé.

Le produit ressorti de l'enquête comme le plus commercialisé est le contrat
d'**Indemnités de Fin de Carrière (IFC / IFC+)**, évalué pour l'entreprise
**SENIA S.A.** (50 salariés, base recueillie via le volet D du questionnaire).

## Livrables

| Fichier | Description |
|---|---|
| `Note_Technique_IFC_SENIA.pdf` | Note technique compilée du produit IFC/IFC+ |
| `Note_Technique_IFC_SENIA.Rmd` | Source de la note (re-tricotable avec RStudio ; images dans `ressource/image/`) |
| `simulateur_IFC_SENIA.xlsx` | Simulateur Excel avec formules vivantes ; la base du personnel est la feuille `BD` |
| `Questionnaire_enquete_IFC.docx` | Questionnaire d'enquête (volets A-D) ayant permis de recueillir les produits et de constituer la base |
| `Presentation_Projet_IFC.pdf` | Présentation Beamer 16:9 du projet (source `.tex`) |
| `Aide_memoire_soutenance.pdf` | Aide-mémoire de soutenance : fil conducteur, lexique, questions probables (source `.md`) |

## Scripts de construction

- `build_simulateur.py` : génère le classeur Excel (base du personnel, formules,
  table CIMA H, barème CCNI) et un contrôle Python des calculs (`synthese.json`).
- `build_questionnaire.py` : génère le questionnaire Word.

## Méthodologie (source : `Evaluation/doc/Cours/FICHE- IFC-TP.pdf`)

Méthode prospective : droits projetés au terme (salaire final x barème CCNI), pondérés
par la probabilité de survie (table CIMA H) et actualisés au taux technique de 3,5 %
(`v^(n+0,5)`), puis répartis au prorata de l'ancienneté :

- Engagement global = somme des VAP ;
- Dette actuarielle = VAP x ancienneté / ancienneté au terme ;
- Charge normale = VAP / ancienneté au terme ;
- Passif IL = IL x t_lic x (1 - u^n)/(1 - u), avec u = (1 - t_rot)/(1 + i) ;
- IFC+ : prime pure = somme des capitaux décès x qx, prime commerciale = prime pure/(1 - g).

## Résultats (au 31/12/2025)

| Rubrique | Montant (FCFA) |
|---|---:|
| Masse salariale annuelle | 344 386 800 |
| Engagement global | 230 047 411 |
| Dette actuarielle | 111 103 527 |
| Charge normale | 8 223 109 |
| Passif indemnité de licenciement | 7 831 464 |
| Capitaux garantis IFC+ | 161 949 480 |
| Prime commerciale IFC+ | 1 203 354 |

Les formules du classeur ont été recalculées sous LibreOffice Calc et redonnent
exactement les valeurs du contrôle Python.
