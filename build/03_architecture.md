# 03 — Architecture cible du classeur unique `ORA_ORA_SSR.xlsx` (SSR Polynésie, facturation CPS)

**Objet.** Spécification d'architecture du **classeur Excel unique** qui remplace les 4 fichiers actuels
(`EXCEL_ETP`, `GP_ETP_1`, `EXCEL_POLYVALENT`, `EXCEL_PRESENCE_SRR`), en corrigeant **par construction**
chacune des 10 ruptures diagnostiquées (`build/02_flux_et_diagnostic.md`, R1→R10). Ce document est une
**conception sur le papier** : il ne construit pas le `.xlsx`. Il est destiné à l'agent CONSTRUCTEUR, qui
doit pouvoir l'implémenter sans réinterpréter. Il se conforme au contrat fixé par `CLAUDE.md` (couches,
noms d'onglets, jointure présence↔accord, contrainte openpyxl≠Power Query, 5 cotations, Definition of Done §9)
et le précise. Toute hypothèse non vérifiée est marquée `[à confirmer]` ; toute divergence éventuelle au
CLAUDE.md serait marquée `[arbitrage métier requis]`.

---

## 1. Synthèse exécutive

**Principe cible.** Un seul `.xlsx`, organisé en couches (Référentiels → Registre `DA` → Saisie présences →
Consolidation → Facturation → Pilotage). L'assistante saisit toujours **3 champs** (Patient, Date,
Programmation) ; tout le reste — N° DA, cotation, régime, Nb Je accordés, PEC, groupe, semaine — se calcule.
La chaîne fragile actuelle (recopier-coller inter-fichiers, jointure par nom de patient brut, `IFERROR` qui
masque les erreurs) est remplacée par une **plomberie intégrée intra-classeur** : une source unique par
donnée, une **clé patient normalisée** systématique avant toute jointure, et des gardes d'erreur qui
**signalent en rouge** au lieu de masquer.

**Décision moteur de jointure — formules dynamiques natives robustifiées (recommandé pour v1).**
Le CONSTRUCTEUR bâtit en **openpyxl, qui ne sait pas écrire Power Query** (CLAUDE.md §5). Pour livrer une v1
**100 % scriptable, sans aucune étape manuelle bloquante**, la jointure présence↔accord est réalisée en
**formules dynamiques Excel 365** (`FILTER` + `XLOOKUP`) appuyées sur des **colonnes d'aide de clé
normalisée** (`TRIM` + réduction des espaces multiples + casse homogène — la « règle d'or anti-rupture »
du CLAUDE.md §4). Le calcul est donc **natif et auto-rafraîchi à chaque ouverture/F9**, sans bouton ni
import manuel. Compromis assumé : Power Query offrirait un rafraîchissement encore plus propre et des
volumes plus confortables, mais imposerait une étape manuelle de collage de code M dans l'Éditeur avancé
hors de portée d'openpyxl. **Power Query est donc fourni comme évolution ultérieure documentée** (code M
séparé, `powerquery/*.pq`), pas comme dépendance de livraison v1. Voir §4 pour les formules exactes et §8
pour le plan de bascule.

**Ce que ça corrige.** Source unique (fin du double fichier ETP — R3) ; volet Polyvalent réintégré dans une
table de consolidation unique pilotée par `Parcours` (R1) ; transfert manuel supprimé, tout est formule
intra-classeur (R2) ; clé patient normalisée avant jointure (R7) ; `#REF!` éliminé par reconstruction de la
colonne d'âge sur le référentiel patient (R4) ; doublons `N° DA` et `N° DA+Date` détectés et signalés
(R6, R10) ; présences sans accord signalées en rouge, jamais silencieusement vides (R8) ; clés tronquées/
orphelines repérées (R9) ; `#VALUE!` supprimés (R5). Détail complet en §7.

---

## 2. Carte des onglets du classeur unique

Noms d'onglets **contractuels** (CLAUDE.md §3). « Source d'origine » = fichier(s) parmi les 4 dont la donnée
provient. « Alimentation » = comment l'onglet se remplit dans la cible.

| # | Onglet | Couche | Rôle | Qui écrit | Source d'origine (4 fichiers) | Alimentation cible |
|---|---|---|---|---|---|---|
| 1 | `REF_Patients` | Référentiel | Patients (clé Recherche = Nom+naissance) | Admin | `EXCEL_ETP!TabPatients␣` (2 157 clés) ∪ `DA[Patient]` | CSV seed → table figée |
| 2 | `REF_Cotations` | Référentiel | 5 cotations + parcours + tarif | Admin | `Parametres!Cotation` | CSV seed |
| 3 | `REF_Regimes` | Référentiel | Régimes + équivalences | Admin | `Parametres!Régimes` | CSV seed |
| 4 | `REF_Pathologies` | Référentiel | Pathologies médicales / PMSI | Admin | `Parametres!Pathologies`,`Path`,`PMSI` | CSV seed |
| 5 | `REF_Provenances` | Référentiel | Provenance + code + motif d'entrée UM | Admin | `Parametres!Provenance` | CSV seed |
| 6 | `REF_Prescripteurs` | Référentiel | Prescripteurs (146) | Admin | `Parametres!Prescripteur` | CSV seed |
| 7 | `REF_Statuts` | Référentiel | Statuts DA (12) + drapeau « exclu facturation » | Admin | `Parametres!Statut` | CSV seed |
| 8 | `REF_Mouvements` | Référentiel | Mouvements | Admin | `Parametres!Mouvements` | CSV seed |
| 9 | `REF_Programmation` | Référentiel | Programmation présence (9) + drapeau « compte présent » | Admin | `Parametres!Programmation` | CSV seed |
| 10 | `REF_Groupes` | Référentiel | Groupes (48) | Admin | `EXCEL_ETP!TabGpes` | CSV seed |
| 11 | `REF_Communes` | Référentiel | Communes + code postal | Admin | `Parametres!Commune` | CSV seed |
| 12 | `REF_Motifs` | Référentiel | Motifs hospit / refus CPS / refus OraOra / annulation | Admin | `Parametres!Motifs *` | CSV seed |
| 13 | `REF_Calendrier` | Référentiel | Date → semaine ISO / mois (2 193 dates) | Admin | `TabCalendrier` | CSV seed (ou formules) |
| 14 | `DA` | Registre | Demandes d'accord / séjours, clé pivot **N° DA** | Assistantes | `EXCEL_ETP!DA` (994) ∪ `EXCEL_POLYVALENT!DA` (676) | Import unique → saisie |
| 15 | `Saisie_Presences` | Saisie | Présences ; saisie = Patient + Date + Programmation | Assistantes | `*!Présences` (ETP+Poly) | Saisie + formules |
| 16 | `CONSO_Presences` | Consolidation | Table unique ETP+Poly (colonne `Parcours`) | Formules (auto) | `Saisie_Presences` | Formules / [PQ ultérieur] |
| 17 | `Facturation` | Facturation | Dossiers prêts à facturer + contrôles | Formules | `CONSO_Presences` + `DA` | Formules |
| 18 | `Cockpit` | Pilotage | Compteurs réel + prévisionnel (ETP & Poly) | TCD | `CONSO_Presences` | TCD |
| 19 | `CTRL_Qualite` | Pilotage | Zone d'alertes d'intégrité (compteurs d'anomalies) | Formules | tous | Formules |

> **Décision périmètre — abandons explicites** (§3 du brief) :
> - **Double fichier ETP** : `GP_ETP_1.xlsx` abandonné ; **une seule source DA** (R3). Les `N° DA` étant
>   strictement identiques (994=994, 0 écart de données — R3), on importe depuis **un seul** fichier ETP.
>   `[arbitrage métier requis]` : confirmer que **`GP_ETP_1` (origine « Marion UNG », non réécrit par
>   openpyxl) est le maître** à importer plutôt que `EXCEL_ETP`.
> - **Feuilles brouillon/extraction** abandonnées : `Détails1`, `Détails2`, `Détails JRS`, `Feuil1..Feuil5`,
>   `Tri 1`, `TCD Activité`, `GCC`, `SUIVI PI`, `DATA Presence ETP`, `Visu présences ETP/Poly`,
>   `Chiffres réel/prévisionnels`, `ANS REAL`, `MOIS REAL`, `JRS REAL/PREVI`. Justification : sorties figées,
>   TCD résiduels, doublons de synthèse — tous reconstruits proprement par `Cockpit` (TCD) et `CTRL_Qualite`.
> - **Colonne `Age fixe séjour`** (DA col Y) **abandonnée telle quelle** (porteuse de `#REF!` sur ~quasi
>   toutes les lignes — R4) et **reconstruite** en `Age_sejour` calculée sur `REF_Patients` (§3, §7-R4).
> - **`TabPresencePoly` vide** (R1) abandonnée : remplacée par la table unique `CONSO_Presences` filtrée
>   par `Parcours`.
> - **`Colonne1`** (16e col. de Présences, sans en-tête utile) abandonnée.
>
> **Ce qu'on garde** : registre `DA` (44 colonnes, dont on retient les utiles — §3), présences (3 champs
> saisis), référentiels `Parametres` (éclatés en `REF_*` uniques), `TabCalendrier` (`REF_Calendrier`),
> et le **pilotage par TCD** (reconstruit dans `Cockpit`).

---

## 3. Dictionnaire de données par onglet clé

Conventions : police Arial ; dates au format `JJ/MM/AAAA` ; `N° DA` en **texte** (préserver le zéro et le
format `AA/NNNNNN`) ; tables nommées Excel (ListObjects) pour chaque onglet. Les colonnes **saisies** sont
notées `[S]`, les **calculées** `[C]` (avec formule cible), les **importées/admin** `[A]`.

### 3.1 Référentiels `REF_*` (verrouillés, admin only)

Chaque `REF_*` est une **ListObject** d'une seule source de vérité ; les listes déroulantes pointent ces
tables (fin des cibles INDIRECT multiples — R-nommage). Colonnes :

| Onglet | Table | Colonnes |
|---|---|---|
| `REF_Patients` | `tPatients` | `Recherche` (clé = Nom+naissance), `Nom`, `Prenom`, `DateNaissance`, **`Cle_Norm`** [C] (clé normalisée, voir §4) |
| `REF_Cotations` | `tCotations` | `Cotation` (HJSR, HJST, HJSN, HJSA, HJSM), `Libelle` `[à confirmer]`, `Parcours_defaut` `[à confirmer]`, `Tarif_XPF` `[à confirmer JOPF]` (laissé vide à renseigner) |
| `REF_Regimes` | `tRegimes` | `Regime` (RGS, RNS, RST, SS, Auto-financement), `Equivalence`, `Facturable_CPS` (booléen ; `[à confirmer]` mapping régime→CPS) |
| `REF_Pathologies` | `tPathologies` | `Pathologie`, `Code_PMSI` `[à confirmer : PF sans PMSI — CLAUDE.md §7]` |
| `REF_Provenances` | `tProvenances` | `Provenance`, `Code_provenance`, `Motif_entree_UM` |
| `REF_Prescripteurs` | `tPrescripteurs` | `Prescripteur` |
| `REF_Statuts` | `tStatuts` | `Statut`, **`Exclu_facturation`** [A] (booléen ; VRAI pour `DEP refusée`, `DEP annulée`, `Refus CPS`, `Refus centre` — voir §4 et §6) |
| `REF_Mouvements` | `tMouvements` | `Mouvement` |
| `REF_Programmation` | `tProgrammation` | `Programmation` (9 valeurs), **`Compte_Present`** [A] (booléen ; VRAI pour `Présent`, `Attente CPS / Présent`, `Refus PEC / Présent` `[à confirmer]`) |
| `REF_Groupes` | `tGroupes` | `Groupe` |
| `REF_Communes` | `tCommunes` | `Commune`, `Code_postal` |
| `REF_Motifs` | `tMotifsHospit`, `tMotifsRefusCPS`, `tMotifsRefusOraOra`, `tMotifsAnnul` | une colonne libellé par table |
| `REF_Calendrier` | `tCalendrier` | `Date`, `Annee`, `Mois`, `Semaine_ISO`, `No_Semaine` (ex. `S24-21`) |

> Valeurs vérifiées sur source : Cotation = {HJSA, HJSM, HJSN, HJSR, HJST} ; Régime = {RGS, RNS, RST, SS,
> Auto-financement} ; Statut = {Attente, BEP fait, CPH fait, CPH prévu, DEP annulée, DEP postée, DEP refusée,
> DEP validée, Refus CPS, Refus centre, Auto-financement} ; Programmation = {Présent, Prévu, Absent non
> prévenu, Annulation, Attente CPS / Présent, Indispo a prévenu, Journée gratuite, Refus PEC / Présent, VAD}.

### 3.2 `DA` — registre des demandes d'accord (clé pivot `N° DA`)

Le registre original a **44 colonnes**. On **retient les utiles** et on en **reconstruit** une (`Age_sejour`).
Clé primaire = **`N° DA`** (col V d'origine), format texte `AA/NNNNNN`.

**Colonnes retenues** (saisies `[S]` par l'assistante via listes déroulantes ; calculées `[C]`) :

| Col cible | Type | Statut | Origine (DA 44 col.) | Note |
|---|---|---|---|---|
| `N_DA` | texte | [S] | `N° DA` (V) | **clé pivot** ; unique requise |
| `Patient` | texte | [S] | `Patient` (B) | liste = `REF_Patients[Recherche]` |
| `Cle_Patient_Norm` | texte | [C] | — | clé normalisée (§4), utilisée pour les jointures |
| `Parcours` | texte | [S] | `Parcours` (AP) | liste = {ETP, Polyvalent} |
| `Motif_Hospit` | texte | [S] | `Motif Hospit` (C) | liste `REF_Motifs` |
| `Pathologie_medicale` | texte | [S] | `Pathologie médicale` (E) | liste `REF_Pathologies` |
| `Regime` | texte | [S] | `Régime` (F) | liste `REF_Regimes` |
| `Provenance` | texte | [S] | `Provenance` (G) | liste `REF_Provenances` |
| `Prescripteur` | texte | [S] | `Prescripteur` (H) | liste `REF_Prescripteurs` |
| `Date_demande` | date | [S] | `Date de demande` (I) | |
| `Statut` | texte | [S] | `Statut` (M) | liste `REF_Statuts` ; pilote l'exclusion facturation |
| `Mouvements` | texte | [S] | `Mouvements` (N) | |
| `Nb_Je_demande` | nombre | [S] | `Nb Je demandé` (Q) | |
| `Cotation_demandee` | texte | [S] | `Cotation demandée` (R) | liste `REF_Cotations` |
| `Date_debut_demandee` | date | [S] | `Date DEBUT demandée` (S) | |
| `Date_fin_demandee` | date | [S] | `Date FIN demandée` (T) | |
| `Date_postage` | date | [S] | `Date de postage` (U) | |
| `Numero_sejour` | texte | [S] | `Numéro de séjour` (W) | |
| `PEC` | texte | [S] | `PEC` (X) | |
| `Nb_Je_Accorde` | nombre | [S] | `Nb Je Accordé` (Z) | plafond facturation |
| `Cotation_accordee` | texte | [S] | `Cotation accordé` (AA) | liste `REF_Cotations` |
| **`Date_debut_accorde`** | date | [S] | `Date début accordé` (AB) | **borne basse jointure** |
| **`Date_fin_accorde`** | date | [S] | `Date fin accordé` (AC) | **borne haute jointure** |
| `Date_accord` | date | [S] | `Date d'accord` (AD) | |
| `Groupe` | texte | [S] | `Groupe initial` (AK) | liste `REF_Groupes` |
| `Programme` | texte | [S] | `Programme` (AJ) | |
| `Date_FIN_PEC` | date | [S] | `Date de FIN de PEC` (AH) | |
| `Age_sejour` | nombre | [C] | **reconstruit** (remplace `Age fixe séjour`/Y) | voir formule ci-dessous, R4 |
| `Nb_Je_consommes` | nombre | [C] | reconstruit (`Nbre de J effectués`/AE) | voir formule, alimente contrôle Nb Je |
| `Doublon_N_DA` | booléen | [C] | — | détection R6 |

**Colonnes abandonnées** : `Pathologie PMSI` (D, PMSI absent en PF), `Date Rep Pres` (J), `CHIR` (K),
`Date opération` (L), `Motifs de refus CPS` (O) / `Motif refus OraOra` (P) `[à confirmer : à conserver si
analyse des refus souhaitée]`, `Age fixe séjour` (Y, `#REF!` — reconstruite), `Date de FIN de PEC`
doublonnée, `Commentaire sortie` (AG), `Date envoi notif fin hospit` (AH), `Date envoi CRH` (AI),
`PEC intiale` (AI), `Groupe initial`/`Date début présence`/`Date fin présence`/`Nb Je présence`/
`Cotation Présence` (AL→AQ : **dérivées de la présence, recalculées dans `CONSO_Presences`**, donc retirées
du registre pour respecter « une source par donnée »), `Date sortie admin` / `obs` (AR, divergent ETP↔Poly).

**Formules clés de `DA`** (toutes avec garde **signalante**, pas masquante) :

- `Cle_Patient_Norm` [C] : `=UPPER(TRIM(SUBSTITUTE(SUBSTITUTE([@Patient],CHAR(160)," "),"  "," ")))`
  enveloppé dans une boucle de réduction (voir §4 pour la version robuste anti-espaces multiples).
- `Doublon_N_DA` [C] : `=SI([@N_DA]="";"";SI(NB.SI(tDA[N_DA];[@N_DA])>1;"DOUBLON";""))`
  → alimente une mise en forme conditionnelle rouge (R6).
- `Age_sejour` [C] (reconstruit, **sans `#REF!`**) :
  `=SI(OU([@Patient]="";[@Date_debut_accorde]="");"";`
  `SIERREUR(DATEDIF(RECHERCHEX([@Cle_Patient_Norm];tPatients[Cle_Norm];tPatients[DateNaissance]);[@Date_debut_accorde];"y");"⚠ naiss. introuvable"))`
  → ne masque plus : si la date de naissance est introuvable, la cellule affiche `⚠ naiss. introuvable`
  (signalé), au lieu du `""` silencieux d'origine (R4).
- `Nb_Je_consommes` [C] :
  `=NB.SI.ENS(CONSO_Presences[N_DA];[@N_DA];CONSO_Presences[Est_Present];VRAI)`
  → nombre de présences réelles rattachées à cette DA (alimente le contrôle Nb Je ≤ accordé, §6).

### 3.3 `Saisie_Presences` — feuille de saisie des assistantes

**Ergonomie inchangée** (CLAUDE.md §2) : l'assistante saisit **3 colonnes seulement**. Le reste est calculé
et **affiché en lecture** sur la même ligne (colonnes calculées en aval, fond grisé/verrouillé).

| Col | Nom cible | Statut | Liste déroulante / formule |
|---|---|---|---|
| A | `Patient` | [S] | liste = `REF_Patients[Recherche]` (validation stricte) |
| B | `Date` | [S] | date ; validation date valide |
| C | `Programmation` | [S] | liste = `REF_Programmation[Programmation]` |
| D | `Parcours` | [S] | liste = {ETP, Polyvalent} — **nécessaire** pour discriminer la jointure (R1) |
| E | `Cle_Patient_Norm` | [C] | clé normalisée (§4) |
| F | `Est_Present` | [C] | `=SIERREUR(RECHERCHEX([@Programmation];tProgrammation[Programmation];tProgrammation[Compte_Present]);FAUX)` |
| G | **`N_DA`** | [C] | **résolution jointure** — voir §4 (formule maîtresse) |
| H | `Statut_Resolution` | [C] | `OK` / `HORS_LISTE` / `SANS_DA` / `MULTI_DA` — voir §4 |
| I | `Cotation` | [C] | `=SI([@N_DA]="";"";RECHERCHEX([@N_DA];tDA[N_DA];tDA[Cotation_accordee];"⚠"))` |
| J | `Regime` | [C] | `=SI([@N_DA]="";"";RECHERCHEX([@N_DA];tDA[N_DA];tDA[Regime];"⚠"))` |
| K | `Nb_Je_accordes` | [C] | `=SI([@N_DA]="";"";RECHERCHEX([@N_DA];tDA[N_DA];tDA[Nb_Je_Accorde];"⚠"))` |
| L | `PEC` | [C] | `=SI([@N_DA]="";"";RECHERCHEX([@N_DA];tDA[N_DA];tDA[PEC];""))` |
| M | `Groupe` | [C] | `=SI([@N_DA]="";"";RECHERCHEX([@N_DA];tDA[N_DA];tDA[Groupe];""))` |
| N | `Semaine` | [C] | `=SIERREUR(RECHERCHEX([@Date];tCalendrier[Date];tCalendrier[No_Semaine]);"⚠ date hors calendrier")` |
| O | `Mois` | [C] | `=SI([@Date]="";"";TEXTE([@Date];"mm/aa"))` |
| P | `Doublon_Cle` | [C] | détection (N° DA + Date) en double (R10) — voir §5 |

> La colonne `Parcours` (D) est ajoutée à la saisie par rapport au « Patient+Date+Programmation » du
> CLAUDE.md : elle est **indispensable** pour rattacher la bonne DA quand un patient a fréquenté les deux
> parcours, et pour réparer R1 (Poly distinct d'ETP). C'est une liste déroulante à 2 valeurs, geste minimal.
> `[arbitrage métier requis]` : valider l'ajout de cette 4e colonne de saisie (ou la pré-remplir par défaut
> selon l'assistante/le classeur d'origine).

### 3.4 `CONSO_Presences` — table de consolidation unique (ETP + Polyvalent)

**Une seule table** alimentée par `Saisie_Presences` (réparation R1 : Poly n'est plus une table séparée
vide, c'est une valeur de la colonne `Parcours`). En v1 formules : `CONSO_Presences` est une **image directe**
de `Saisie_Presences` (mêmes colonnes calculées, recopiées par référence structurée ou par
`=Saisie_Presences[...]`), ce qui évite tout recopier-coller (R2). En évolution Power Query, c'est la sortie
du merge (`powerquery/CONSO_Presences.pq`).

Colonnes : `Patient`, `Cle_Patient_Norm`, `Date`, `Programmation`, `Est_Present`, `Parcours`, `N_DA`,
`Cotation`, `Regime`, `Nb_Je_accordes`, `Semaine`, `Mois`, `Groupe`, `PEC`, `Statut_Resolution`.

> **Unification** : ETP et Poly cohabitent par la colonne `Parcours`. Tous les comptages (`Cockpit`,
> `Facturation`) filtrent sur `Parcours` — plus jamais de table Poly vide (R1) ni de COUNTIFS pointant une
> table fantôme.

### 3.5 `Facturation` — dossiers prêts à facturer

Table dérivée (formules / `[PQ ultérieur]`) qui **ne liste que les lignes facturables** (règles §6). Colonnes :

| Col | Nom | Formule / source |
|---|---|---|
| `N_DA` | `CONSO_Presences[N_DA]` (filtré) | |
| `Patient` | `RECHERCHEX(N_DA→tDA[Patient])` | |
| `Parcours` | `CONSO_Presences[Parcours]` | |
| `Cotation` | `tDA[Cotation_accordee]` | |
| `Regime` | `tDA[Regime]` | |
| `Date` | `CONSO_Presences[Date]` | |
| `Tarif_XPF` | `RECHERCHEX(Cotation→tCotations[Tarif_XPF])` | `[à confirmer JOPF]` ; vide tant que tarifs non saisis |
| `Nb_Je_consommes` | `tDA[Nb_Je_consommes]` | |
| `Nb_Je_accordes` | `tDA[Nb_Je_Accorde]` | |
| `Eligible` | formule combinant les 6 conditions §6 (booléen) | |
| `Motif_rejet` | texte expliquant pourquoi non éligible (jamais silencieux) | |

La table **affiche tout** mais marque `Eligible` ; un filtre/vue ne montre que `Eligible=VRAI`. Tant que le
**format CPS n'est pas tranché** (CLAUDE.md §7), `Facturation` produit un **bordereau interne validé**, pas
le fichier CPS final (§9).

### 3.6 `Cockpit` — pilotage (TCD)

TCD construits sur `CONSO_Presences` (réel) et sur `DA` (prévisionnel via `Nb_Je_Accorde`). Axes : `Parcours`,
`Mois`/`Semaine`, `Cotation`, `Regime`. Compteur de présences = somme de `Est_Present`. **Réel ETP et réel
Poly tous deux corrects** (R1 résolu, source unique). Le prévisionnel s'appuie sur `DA` (une seule source —
fin du mélange Poly/ETP de R5-diagnostic « Chiffres prévisionnels »).

### 3.7 `CTRL_Qualite` — zone de contrôle qualité

Une feuille de compteurs d'anomalies (formules `NB.SI`), pour rendre les ruptures **visibles** :
lignes `HORS_LISTE`, `SANS_DA`, `MULTI_DA`, doublons `N_DA`, doublons `N_DA+Date`, présences sans calendrier,
`Nb_Je_consommes > Nb_Je_Accorde`. Chaque compteur > 0 = à traiter. Voir §5.

---

## 4. Mécanisme de jointure détaillé (pièce maîtresse)

### 4.1 Normalisation de la clé patient (règle d'or, CLAUDE.md §4)

Toute clé patient est normalisée **avant** toute jointure : `TRIM` + réduction des espaces multiples internes
à un seul + suppression de l'espace insécable `CHAR(160)` + casse homogène (`UPPER`). Formule cible
(colonne d'aide `Cle_Patient_Norm`, présente dans `REF_Patients`, `DA`, `Saisie_Presences`) :

```
=MAJUSCULE(
   SUPPRESPACE(
     SUBSTITUE(SUBSTITUE(SUBSTITUE([@Patient];CHAR(160);" ");"  ";" ");"  ";" ")
   )
 )
```

`SUPPRESPACE` (`TRIM`) réduit déjà les espaces multiples internes ET de bord en Excel ; les `SUBSTITUE`
neutralisent l'espace insécable et tout résidu. Cette normalisation **neutralise les 66 / 1 123 / 1 229 / 2 694
valeurs à espaces multiples** constatées (R7) et homogénéise la casse. La jointure s'opère **toujours** sur
`Cle_Patient_Norm`, jamais sur `Patient` brut.

### 4.2 Résolution du `N° DA` (Saisie_Presences col G) — formule maîtresse

Règle métier conservée (CLAUDE.md §4) : une présence se rattache à une DA si **Patient normalisé identique**
ET **Date ∈ [`Date_debut_accorde` ; `Date_fin_accorde`]** ET **`Parcours` identique** ET **Statut DA non
exclu** (`Exclu_facturation`=FAUX, c.-à-d. ∉ {DEP refusée, DEP annulée, Refus CPS, Refus centre}).

Réalisée par `FILTER` (dynamique 365) sur les colonnes d'aide normalisées :

```
G (liste des DA candidates, formule de travail, peut vivre en colonne cachée) :
=SI([@Patient]="";"";
  FILTRE(tDA[N_DA];
    (tDA[Cle_Patient_Norm]=[@Cle_Patient_Norm])
   *(tDA[Parcours]=[@Parcours])
   *(tDA[Date_debut_accorde]<=[@Date])
   *(tDA[Date_fin_accorde]>=[@Date])
   *(tDA[Statut_Exclu]=FAUX);
   ""))
```

où `tDA[Statut_Exclu]` est une colonne calculée de `DA` :
`=SIERREUR(RECHERCHEX([@Statut];tStatuts[Statut];tStatuts[Exclu_facturation]);FAUX)`.

### 4.3 Gestion 0 / 1 / N match (jamais de vide silencieux — DoD)

Le `N_DA` final et le `Statut_Resolution` (col H) gèrent explicitement les 3 cas. En posant
`m = NBVAL(FILTRE(...))` (nombre de DA candidates) :

```
N_DA (col G finale) :
=SI([@Cle_Patient_Norm]="";"";
  SI(NB.SI(tPatients[Cle_Norm];[@Cle_Patient_Norm])=0; "⚠ HORS LISTE";
    LET(res; FILTRE(tDA[N_DA]; (… mêmes critères qu'en 4.2 …); "");
        SI(NBVAL(res)=0; "⚠ SANS DA";
           SI(NBVAL(res)>1; "⚠ MULTI DA";
              INDEX(res;1))))))
```

```
Statut_Resolution (col H) :
=SI([@N_DA]="";"";
   SI(GAUCHE([@N_DA];1)="⚠"; STXT([@N_DA];3;20); "OK"))
```

- **Patient hors référentiel** → `⚠ HORS LISTE` (réparation directe de la DoD §9 : nom mal orthographié /
  hors-liste = ligne signalée, jamais vide).
- **0 match DA** → `⚠ SANS DA` (R8 : présence non rattachée rendue **visible** au lieu d'être figée vide).
- **>1 match DA** → `⚠ MULTI DA` (R6/R10 : chevauchement d'accords ou doublon de DA → alerte, on ne prend
  pas silencieusement la 1re comme l'`XLOOKUP` d'origine).
- **1 match** → le `N° DA` est posé, toutes les colonnes I→M se remplissent par `RECHERCHEX`.

Chaque valeur commençant par `⚠` déclenche la **mise en forme conditionnelle rouge** de la ligne (§5).
**Aucun `IFERROR`/`SIERREUR` masquant** sur la chaîne de facturation : les gardes affichent un libellé,
elles ne renvoient jamais `""` silencieux (correction directe de la cause racine R4 du diagnostic).

> Note de robustesse : `Date_debut_accorde`/`Date_fin_accorde` sont des **dates réelles** (vérifié sur
> source : `datetime`), donc la comparaison d'intervalle est numérique et fiable ; les DA d'un même patient
> sont **séquentielles non chevauchantes** (vérifié : ALEXANDRE Nilton — 3 DA bord à bord), ce qui garantit
> qu'un `⚠ MULTI DA` signale une **vraie** anomalie (chevauchement saisi) et non un cas normal.

---

## 5. Contrôles d'intégrité & validations

### 5.1 Listes déroulantes (validation de données)

Toutes les listes pointent une **table `REF_*` unique** (fin des cibles INDIRECT multiples / nommage
incohérent — R-nommage du diagnostic) :

| Cellule | Source liste |
|---|---|
| `Saisie_Presences[Patient]` | `REF_Patients[Recherche]` (validation **stricte** : refus de saisie hors-liste, ou avertissement non bloquant `[arbitrage métier requis]`) |
| `Saisie_Presences[Programmation]` | `REF_Programmation[Programmation]` |
| `Saisie_Presences[Parcours]` | liste statique {ETP, Polyvalent} |
| `DA[Statut]`,`DA[Regime]`,`DA[Cotation_*]`,`DA[Groupe]`,`DA[Provenance]`,`DA[Prescripteur]`,`DA[Pathologie_medicale]`,`DA[Motif_Hospit]` | tables `REF_*` correspondantes |

### 5.2 Mises en forme conditionnelles d'alerte (ligne rouge)

| Déclencheur | Règle | Rupture corrigée |
|---|---|---|
| Clé patient non résolue | `Statut_Resolution="HORS LISTE"` | DoD §9 / R7 |
| Présence sans accord | `Statut_Resolution="SANS DA"` | R8 |
| Accords multiples | `Statut_Resolution="MULTI DA"` | R6 |
| `N° DA` en doublon dans `DA` | `DA[Doublon_N_DA]="DOUBLON"` | R6 |
| Doublon (N° DA + Date) en présence | `NB.SI.ENS(CONSO[N_DA];[@N_DA];CONSO[Date];[@Date])>1` | R10 |
| Date hors calendrier | `Semaine` commence par `⚠` | qualité |
| Nb Je dépassé | `Nb_Je_consommes > Nb_Je_accordes` (sur `Facturation`/`DA`) | DoD §9 |

### 5.3 Feuille `CTRL_Qualite`

Compteurs `NB.SI`/`NB.SI.ENS` de chacun des déclencheurs ci-dessus + total de lignes `Saisie` /
lignes `Eligible`. Objectif : tout compteur > 0 visible d'un coup d'œil ; sert de tableau de bord qualité
avant transmission CPS.

---

## 6. Règles de facturation (DoD §9)

Un dossier (ligne) apparaît comme **`Eligible=VRAI`** dans `Facturation` **si et seulement si TOUTES** ces
conditions sont vraies :

1. `Est_Present = VRAI` (Programmation comptée présente via `REF_Programmation`).
2. `Regime` facturable **CPS** (`REF_Regimes[Facturable_CPS]=VRAI` ; `[à confirmer]` mapping exact, a priori
   RGS/RNS/RST = CPS, SS/Auto-financement exclus).
3. `Cotation` **valide** (∈ `REF_Cotations` = {HJSR, HJST, HJSN, HJSA, HJSM}).
4. `N_DA` **résolu** (`Statut_Resolution="OK"`, donc ni HORS LISTE, ni SANS DA, ni MULTI DA).
5. `Date` **dans la fenêtre d'accord** (`Date_debut_accorde ≤ Date ≤ Date_fin_accorde`) — déjà garanti par
   la jointure §4, re-vérifié en colonne de contrôle.
6. `Nb_Je_consommes ≤ Nb_Je_Accorde` pour la DA (sinon la ligne au-delà du plafond est signalée, non éligible).

`Motif_rejet` documente la 1re condition non satisfaite (jamais de rejet silencieux).

**Tableau `REF_Cotations`** (valeurs vérifiées ; **aucun tarif inventé**) :

| Cotation | Libellé | Tarif (XPF/jour) |
|---|---|---|
| HJSR | `[à confirmer]` | `[à confirmer JOPF]` |
| HJST | `[à confirmer]` | `[à confirmer JOPF]` |
| HJSN | `[à confirmer]` | `[à confirmer JOPF]` |
| HJSA | `[à confirmer]` | `[à confirmer JOPF]` |
| HJSM | `[à confirmer]` | `[à confirmer JOPF]` |

**Statut du format CPS** : non tranché (CLAUDE.md §7 — pas de Carte Vitale, pas de PMSI, tarifs par arrêté
JOPF). Tant que le canal/format GDR-DSI CPS n'est pas confirmé, `Facturation` = **bordereau interne validé**
(liste des dossiers prêts), pas le format de transmission CPS final. Action hors-code : entretien GDR/DSI CPS.

---

## 7. Tableau de traçabilité Rupture → Correction

| Rupture | Gravité | Cause d'origine | Correction par conception |
|---|---|---|---|
| **R1** Volet Poly perdu | 🔴 | `TabPresencePoly` jamais alimentée (collage manuel oublié) | **Table unique `CONSO_Presences`** discriminée par colonne `Parcours` ; plus de table séparée à remplir. Comptages filtrent `Parcours`. |
| **R2** Transfert manuel amont→SRR | 🔴 | Tables SRR sans formule, recopier-coller | **Tout intra-classeur** : `CONSO_Presences` est une dérivation par formule de `Saisie_Presences` ; zéro copier-coller, recalcul natif (ou PQ « Actualiser tout » en évolution). |
| **R3** Double source ETP | 🟠 | 2 copies du même classeur | **Une seule source `DA`** ; `GP_ETP_1` abandonné (données identiques). `[arbitrage : maître = GP_ETP_1]`. |
| **R4** `#REF!` masqué (Age) | 🟠 | `XLOOKUP(#REF!)` + `IFERROR` | Colonne `Age fixe séjour` **abandonnée**, **reconstruite** `Age_sejour` sur `REF_Patients` ; garde **signalante** `⚠ naiss. introuvable`, plus de masquage. |
| **R5** `#VALUE!` SUIVI PI | 🟡 | feuille de suivi non protégée | Feuille `SUIVI PI` **abandonnée** ; suivi reconstruit dans `Cockpit`/`CTRL_Qualite` sans erreur. |
| **R6** Doublons `N° DA` | 🟠 | XLOOKUP prend la 1re occurrence | Colonne `Doublon_N_DA` + MFC rouge ; jointure renvoie `⚠ MULTI DA` au lieu de choisir silencieusement. |
| **R7** Patients clés instables | 🟠 | match nom brut, espaces/casse | **Normalisation systématique** `Cle_Patient_Norm` (TRIM + espaces + `CHAR(160)` + UPPER) avant toute jointure (règle d'or §4). |
| **R8** Présences sans N° DA | 🟠 | jointure échouée figée vide | `Statut_Resolution="SANS DA"` + ligne rouge ; **rendu visible**, jamais vide silencieux. |
| **R9** Orphelins / DA non consommés | 🟡 | clés tronquées, DA Poly non consommées (effet R1) | Validation stricte `REF_Patients` + `CTRL_Qualite` (DA jamais consommées) ; R1 résolu supprime les 675 Poly. |
| **R10** Doublons (N° DA + Date) | 🟡 | collages répétés / double saisie | Colonne `Doublon_Cle` + MFC rouge + compteur `CTRL_Qualite` ; un seul flux de saisie (plus de collage). |

---

## 8. Plan de construction pour le CONSTRUCTEUR

**Ce qui est scriptable openpyxl (build automatique, `build/build_workbook.py`) :**

1. **Extraire les référentiels** des fichiers source vers `data/referentiels/*.csv` (seeds, anonymisés) :
   `patients.csv`, `cotations.csv`, `regimes.csv`, `pathologies.csv`, `provenances.csv`,
   `prescripteurs.csv`, `statuts.csv` (avec colonne `Exclu_facturation`), `mouvements.csv`,
   `programmation.csv` (avec colonne `Compte_Present`), `groupes.csv`, `communes.csv`, `motifs_*.csv`,
   `calendrier.csv`. Source la plus riche = `EXCEL_POLYVALENT!Parametres` (17 tables) ∪ `EXCEL_ETP`
   (`TabPatients␣`, `TabGpes`).
2. **Créer les onglets `REF_*`** depuis les CSV, en **ListObjects** nommées (`tPatients`, `tCotations`, …),
   police Arial, verrouillage (feuilles protégées admin).
3. **Créer `DA`** : structure 44→colonnes retenues (§3.2), colonnes calculées (`Cle_Patient_Norm`,
   `Statut_Exclu`, `Doublon_N_DA`, `Age_sejour`, `Nb_Je_consommes`), 8 listes déroulantes `REF_*`.
4. **Créer `Saisie_Presences`** : 3 (+1 `Parcours`) colonnes de saisie + listes déroulantes + colonnes
   calculées E→P (formules §3.3 / §4), MFC d'alerte.
5. **Créer `CONSO_Presences`** : dérivation par formule de `Saisie_Presences` (v1).
6. **Créer `Facturation`** : colonnes + formule `Eligible`/`Motif_rejet` (§6), tarifs vides.
7. **Créer `CTRL_Qualite`** : compteurs d'anomalies.
8. **Squelette `Cockpit`** (les TCD finaux sont posés à la main — openpyxl ne crée pas de TCD propre).
9. **Recalcul/contrôle** : `python build/build_workbook.py --recalc` (LibreOffice si dispo) → vérifier
   **zéro erreur de formule** (`#REF!`/`#N/A`/`#VALUE!`/`#NAME?`/`#DIV/0!`).

**Ce qui reste manuel (hors openpyxl) :**

- **Import des données réelles** (présences, registre DA) depuis les sources — données patients sensibles,
  **jamais committées** (CLAUDE.md §2 ; RGPD-PF). Le dépôt ne contient que des seeds anonymisés.
- **TCD `Cockpit`** : création/rafraîchissement dans Excel.
- **[Évolution] Power Query** : coller le code M de `powerquery/*.pq` dans l'Éditeur avancé si bascule PQ
  décidée (remplace alors la dérivation par formule de `CONSO_Presences`). Documenté dans
  `docs/04_power_query.md`.

**Livrables de la phase 4 (construction) :**

- `build/build_workbook.py` (générateur) ;
- `data/referentiels/*.csv` (seeds anonymisés, listés ci-dessus) ;
- `output/ORA_ORA_SSR_v0.xlsx` (classeur structure, sans données réelles) ;
- `powerquery/CONSO_Presences.pq` (+ `DA.pq` si besoin) — code M pour l'évolution PQ ;
- `docs/04_power_query.md` (procédure de bascule) ; mise à jour `docs/03_dictionnaire_donnees.md`.

---

## 9. Risques & points ouverts

- **🔴 Dépendance CPS non tranchée** (CLAUDE.md §7) : format/canal de transmission GDR-DSI inconnu. `Facturation`
  reste un **bordereau interne** tant que non résolu. Tarifs JOPF des 5 cotations **non renseignés** (à
  saisir dans `REF_Cotations`, ne pas inventer). Action hors-code prioritaire.
- **`[arbitrage métier requis]` — maître ETP** : confirmer importer **`GP_ETP_1`** (origine préservée) et non
  `EXCEL_ETP`.
- **`[arbitrage métier requis]` — 4e colonne `Parcours` en saisie** : valider le geste, ou définir une valeur
  par défaut. Sans elle, un patient présent sur les deux parcours ne peut être rattaché à la bonne DA.
- **`[arbitrage] — validation Patient stricte vs avertissement** : refuser la saisie hors-liste (sécurise mais
  bloque les nouveaux patients tant que `REF_Patients` n'est pas mis à jour) OU avertir sans bloquer.
- **`[à confirmer]` — règles `Compte_Present`** (quelles Programmation comptent « présent » : `Présent` seul,
  ou aussi `Attente CPS / Présent`, `Refus PEC / Présent` ?) et **`Facturable_CPS`** par régime.
- **`[à confirmer]` — colonnes DA des refus** (`Motifs de refus CPS/OraOra`) : à conserver si l'analyse des
  refus est un besoin métier.
- **Volume Excel 365** : ~26 000 lignes de présence × `FILTER` par ligne est calculable mais lourd ; si la
  latence gêne, basculer `CONSO_Presences` en **Power Query** (évolution prévue) résout aussi la performance.
- **Décalage v1 formules vs PQ** : la v1 exige Excel **365** (fonctions `FILTRE`/`RECHERCHEX`/`LET`
  dynamiques). `[à confirmer]` que toutes les assistantes disposent de 365 ; sinon, PQ devient nécessaire
  dès v1.
