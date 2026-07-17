---
title: "Aide-mémoire de soutenance, Projet I : Indemnités de Fin de Carrière"
author: "Groupe 1, Cours d'Assurance Vie, ENSAE ISE 3"
date: "Juillet 2026"
lang: fr
geometry: margin=2.2cm
fontsize: 11pt
---

# Le fil conducteur en une minute

Le marché vie sénégalais est porté par les assurances collectives d'entreprise
(63,5 % des primes vie) et croît vite (113,6 milliards FCFA de primes vie en
2024, +14,8 %, source FANAF). Le produit phare de ce segment est le contrat
d'Indemnités de Fin de Carrière : toute entreprise doit, selon la convention
collective (CCNI), verser une indemnité à ses salariés au départ à la retraite,
et les normes comptables l'obligent à provisionner cet engagement. Nous avons
donc retenu l'IFC, bâti sa note technique, et construit un simulateur Excel
appliqué à une entreprise de 50 salariés, alimenté par un questionnaire de
collecte.

# Ce qu'il faut dire, diapo par diapo

1. **Contexte** : citer deux chiffres seulement : 113,6 Mds de primes vie au
   Sénégal en 2024 (+14,8 %) et 63,5 % de collectives dans la vie. Conclure :
   « le produit collectif phare, c'est l'IFC ».
2. **Objectifs** : reprendre l'énoncé (enquête, note technique du produit le
   plus commercialisé) et annoncer les trois livrables.
3. **Démarche** : quatre étapes : revue de marché (FANAF), enquête
   (questionnaire), collecte de la base entreprise, modélisation actuarielle.
   Si on demande les résultats de l'enquête : elle est présentée comme une
   démarche engagée pour consolider le constat FANAF ; ne pas avancer de
   chiffres d'enquête, s'appuyer sur les chiffres FANAF.
4. **Questionnaire** : quatre volets (identification, portefeuille, technique,
   fiche personnel) ; cases à cocher pour le qualitatif, cases numériques pour
   le quantitatif.
5. **Données de marché** : le tableau FANAF ; insister sur la part des
   collectives (63,5 %) et la croissance de la vie.
6. **Méthode (formules)** : dire avec des mots simples : « on projette le
   salaire au départ à la retraite, on applique le barème CCNI, on multiplie
   par la probabilité d'être en vie à 60 ans, on ramène à aujourd'hui avec le
   taux de 3,5 %, et on répartit selon l'ancienneté déjà accomplie ».
7. **Résultats** : engagement global 230,0 M ; dette actuarielle 111,1 M
   (48 % de l'engagement, cohérent avec 10,2 ans d'ancienneté sur 28,1 au
   terme) ; charge normale 8,2 M ; passif IL 7,8 M.
8. **Financement** : trois options ; la 3 (30,4 M par an sur 5 ans, soit
   8,84 % de la masse salariale) est fiscalement avantageuse car la prime est
   entièrement déductible.
9. **IFC+** : couverture décès du personnel ; prime de 1,2 M, soit 0,35 % de
   la masse salariale seulement.
10. **Simulateur** : six feuilles, tout est en formules : changer une
    hypothèse ou coller une autre base recalcule tout.

# Lexique en langage simple

- **Table de mortalité CIMA H** : table réglementaire de la zone CIMA ; l(x)
  donne le nombre de survivants à l'âge x sur 1 000 000 de naissances.
  l(60)/l(x) est donc la probabilité d'atteindre 60 ans quand on a x ans.
- **q(x)** : probabilité de décéder dans l'année à l'âge x, soit
  (l(x) - l(x+1)) / l(x).
- **Taux technique (3,5 %)** : taux d'actualisation garanti ; c'est le maximum
  autorisé par l'article 338 du code CIMA, gage de prudence.
- **Actualisation v^(n+0,5)** : un franc dans n années vaut moins qu'un franc
  aujourd'hui ; le +0,5 place le départ au milieu de l'exercice.
- **Salaire final** : salaire mensuel projeté au départ à la retraite avec
  2 % de progression par an.
- **Barème CCNI** : droit par année de présence : 20 % du salaire mensuel les
  5 premières années, puis 30 %, 40 %, 50 %, 60 % au-delà de 20 ans.
- **VAP (valeur actuelle probable)** : droits au terme multipliés par la
  probabilité de survie et par le facteur d'actualisation.
- **Engagement global** : somme des VAP de tous les salariés.
- **Dette actuarielle** : VAP multipliée par ancienneté acquise / ancienneté au
  terme (méthode prospective, dite des unités de crédit projetées) ; c'est la
  part déjà « gagnée » par les salariés, à inscrire au bilan.
- **Charge normale** : VAP / ancienneté au terme ; le coût d'une année de
  service, versé chaque année dans le fonds.
- **Passif IL** : indemnité de licenciement acquise, multipliée par le taux
  annuel de licenciement (1 %) et capitalisée sur la durée résiduelle avec
  u = (1 - turn-over) / (1 + i).
- **IFC+** : temporaire décès collective à prime annuelle ; capital = indemnité
  de licenciement + indemnité sur demi-salaire ; prime pure = somme des
  capitaux multipliés par q(x) ; prime commerciale = prime pure / (1 - 2 %).
- **Participation aux bénéfices** : le fonds est revalorisé à 3,5 % au minimum
  et reçoit 85 % des bénéfices financiers et 90 % des bénéfices techniques.

# Questions probables et réponses courtes

**Pourquoi 3,5 % ?** C'est le taux technique maximal autorisé par le code CIMA
(article 338) ; prendre moins serait plus prudent mais augmenterait la prime.

**Pourquoi la probabilité l(60)/l(x) ?** L'indemnité n'est due que si le
salarié atteint la retraite : on pondère par sa probabilité de survie lue dans
la table CIMA H.

**Pourquoi diviser par l'ancienneté au terme ?** Pour étaler l'engagement sur
toute la carrière : chaque année de service « achète » une part égale de
l'indemnité finale ; c'est le principe de la méthode prospective.

**Différence entre engagement global et dette actuarielle ?** L'engagement
regarde toute la carrière (passée et future) ; la dette ne retient que la part
correspondant aux années déjà travaillées. Ici 48 % : le rapport moyen des
anciennetés (10,2 / 28,1 ans) est d'environ 36 %, et la proportion monte à
48 % car les salariés les plus anciens, aux droits les plus lourds, pèsent
davantage.

**Pourquoi trois options de financement ?** Pour s'adapter à la trésorerie de
l'entreprise : tout payer d'un coup (option 1), régulariser les droits acquis
puis payer chaque année (option 2), ou lisser la dette sur 5 ans avec
déduction fiscale intégrale (option 3).

**Que couvre l'IFC+ ?** Le décès d'un salarié en activité : l'assureur paie aux
ayants droit l'indemnité conventionnelle, sans limite de nombre de décès dans
l'année ; la prime est recalculée chaque année.

**Pourquoi réactualiser tous les trois ans ?** Les hypothèses (salaires,
effectifs, taux) évoluent ; une prime d'ajustement recale le fonds sur la
dette réelle.

**Et si les salaires augmentent plus vite ?** L'engagement monte : +1 point de
taux de progression le fait passer de 230,0 à 271,4 M FCFA ; d'où l'analyse de
sensibilité.

**D'où vient la base de données ?** De la fiche de collecte du volet D du
questionnaire, remplie par l'entreprise SENIA S.A. (50 salariés).

**Le simulateur est-il réutilisable ?** Oui : il suffit de coller la base d'une
autre entreprise et d'ajuster les hypothèses ; tout est en formules.
