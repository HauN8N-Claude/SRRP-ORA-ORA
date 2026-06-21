# 03 — Architecture cible du classeur unique `ORA_ORA_SSR.xlsx` (v2 — couche Facturation intégrée)

**Objet.** Spécification d'architecture du **classeur Excel unique** qui remplace les 5 fichiers actuels
(`EXCEL_ETP`, `GP_ETP_1`, `EXCEL_POLYVALENT`, `EXCEL_PRESENCE_SRR`, **et le fichier de facturation
`Tableau_suivi_Factures_CPS_2025.xlsm`**), en corrigeant **par construction** chacune des 11 ruptures
diagnostiquées (`build/02_flux_et_diagnostic.md` R1→R10 + `build/02b_facturation_systeme_reel.md` **R11**).
Ce document est une **conception sur le papier** : il ne construit pas le `.xlsx`. Il est destiné à l'agent
CONSTRUCTEUR, qui doit pouvoir l'implémenter sans réinterpréter. Il se conforme au contrat fixé par `CLAUDE.md`
(couches, noms d'onglets, jointure présence↔accord, contrainte openpyxl≠Power Query, 5 cotations, Definition of
Done §9) et le précise. Toute hypothèse non vérifiée est marquée `[à confirmer]` ; tout arbitrage métier ouvert
est marqué `[arbitrage]`.

> **Révision v2.** Cette version **écrase** la v1. Elle intègre trois apports découverts après la v1 :
> 1. le **vrai système de facturation** (fichier `.xlsm`, analysé dans `02b`) — Power Query inter-fichiers
>    `Z:\…`, suivi `TabSuiviApi`, bordereaux, modèle de facture imprimable, tarifs réels `TabTarif` ;
> 2. le **référentiel patient réel** `GP_PATIENTS.xlsx` (feuille `Patients `, ≈1 102 patients, clé `Recherche`) ;
> 3. le **diagnostic corrigé** : **R11 🔴** (déversement = Power Query inter-fichiers, cause n°1 du désalignement)
>    et la **requalification de R2** (transfert automatique par PQ, pas manuel).
>
> **Décisions déjà validées et intégrées (ne pas re-questionner)** :
> - **Moteur = les DEUX** : formules dynamiques natives (`FILTRE`/`RECHERCHEX`/`LET` + clé patient normalisée)
>   pour une v1 100 % scriptable openpyxl, **avec** le code M Power Query fourni en parallèle (documenté) comme
>   évolution.
> - **Colonne `Parcours` ajoutée à la saisie** (4e champ, liste {ETP, Polyvalent}).
> - **Validation patient = avertir sans bloquer** (nom hors-liste → `⚠ HORS LISTE` en rouge, saisie autorisée).
> - **Source ETP maître = `GP_ETP_1.xlsx`** (= « GP ETP.xlsx »).
> - **Objectif réaffirmé** : UN SEUL fichier Excel regroupant TOUTES les fonctions de facturation
>   (référentiels + DA + saisie présences + consolidation + suivi factures + bordereaux + facture imprimable +
>   tarifs).

---

## 1. Synthèse exécutive

**Principe cible.** Un seul `.xlsx`, organisé en couches
(Référentiels → Registre `DA` → Saisie présences → Consolidation → **Facturation complète** → Pilotage).
L'assistante saisit toujours **3 champs métier** (Patient, Date, Programmation) + **1 champ de discrimination**
(`Parcours`) ; tout le reste — N° DA, cotation, régime, Nb Je accordés, PEC, groupe, semaine, **puis montant,
bordereau, facture** — se calcule. La chaîne fragile actuelle (Power Query inter-fichiers `Z:\…`, jointure par
nom de patient brut, `IFERROR`/`SIERREUR` qui masquent les erreurs) est remplacée par une **plomberie intégrée
intra-classeur** : une source unique par donnée, une **clé patient normalisée** systématique avant toute
jointure, et des gardes d'erreur qui **signalent en rouge** au lieu de masquer.

**Le saut de la v2 : la facturation entre dans le fichier.** Le « déversement automatique » que décrit
l'établissement est en réalité un **Power Query inter-fichiers** lisant `Z:\2 PLANNIF & PROGRAMMATION\GESTION DES
PATIENTS\GP ETP.xlsx`, `GP POLYVALENT.xlsx`, `GP Patients.xlsx`, et alimentant un `.xlsm` de facturation (7
requêtes M, suivi `TabSuiviApi`, bordereaux, facture, tarifs). Ce dispositif est **automatique mais fragile** :
positionné/typé **par nom de table et de colonne exact** sur des fichiers externes au format instable (espaces
d'en-tête, dernière colonne divergente ETP↔Poly, chemin réseau `Z:`). C'est la **cause n°1 du désalignement**
(R11). **En ramenant la facturation dans le même classeur que la saisie, le Power Query inter-fichiers disparaît
de la chaîne critique** : les jointures deviennent **intra-classeur** par référence structurée, et la cause R11
s'éteint par construction.

**Décision moteur de jointure — formules dynamiques natives robustifiées (v1) + Power Query documenté (évolution).**
Le CONSTRUCTEUR bâtit en **openpyxl, qui ne sait pas écrire Power Query** (CLAUDE.md §5). Pour livrer une v1
**100 % scriptable, sans étape manuelle bloquante**, toutes les jointures (présence↔accord, présence↔calendrier,
suivi↔DA, suivi↔patient, suivi↔tarif) sont réalisées en **formules dynamiques Excel 365** (`FILTRE` + `RECHERCHEX`
+ `LET`) appuyées sur des **colonnes d'aide de clé normalisée** (la « règle d'or anti-rupture » du CLAUDE.md §4).
Calcul **natif et auto-rafraîchi** à chaque ouverture/F9, sans bouton ni import manuel. **Power Query est fourni
comme évolution ultérieure documentée** (code M `powerquery/*.pq`, transposé en jointures **intra-classeur** —
plus jamais sur `Z:\…`), pas comme dépendance de livraison v1.

**Ce que ça corrige.** Source unique (fin du double fichier ETP — R3) ; volet Polyvalent réintégré par la colonne
`Parcours` (R1) ; transfert PQ inter-fichiers supprimé, tout est intra-classeur (R2, **R11**) ; clé patient
normalisée avant jointure (R7) ; `#REF!` éliminés (R4 dans DA, **et les 443 `#REF!` de la colonne Q parasite de
`GP_PATIENTS`**) ; doublons `N° DA` et `N° DA+Date` détectés et signalés (R6, R10) ; présences sans accord
signalées en rouge (R8) ; `#VALUE!`/`#N/A`/`#REF!` du `.xlsm` (28 erreurs littérales + 5 plages nommées cassées)
supprimés (R5). Détail complet en §8.

---

## 2. Carte des onglets du classeur unique (v2)

Noms d'onglets **contractuels** (CLAUDE.md §3), enrichis de la **couche Facturation**. « Source d'origine » = vrai
mapping des fichiers réels (cf. `02b` §1). « Alimentation » = comment l'onglet se remplit dans la cible.

**Mapping des fichiers réels (rappel `02b`)** : `GP ETP.xlsx` = **`GP_ETP_1`** (maître ETP, origine « Marion UNG ») ;
`GP POLYVALENT.xlsx` = **`EXCEL_POLYVALENT`** ; `GP Patients.xlsx` = **`GP_PATIENTS`** (désormais FOURNI) ;
facturation = **`Tableau_suivi_Factures_CPS_2025.xlsm`** ; `EXCEL_PRESENCE_SRR` = **ancienne consolidation
obsolète** (abandonnée) ; `EXCEL_ETP` = copie/export d'ETP (abandonnée au profit de `GP_ETP_1`).

| # | Onglet | Couche | Rôle | Qui écrit | Source d'origine (réelle) | Alimentation cible |
|---|---|---|---|---|---|---|
| 1 | `REF_Patients` | Référentiel | Patients (clé `Recherche`) + `Cle_Norm` | Admin | **`GP_PATIENTS!Patients `** (≈1 102) | CSV seed → table figée |
| 2 | `REF_Cotations` | Référentiel | 5 cotations + parcours | Admin | `Parametres!Cotation` | CSV seed |
| 3 | `REF_Tarifs` | Référentiel/Factu | **5 forfaits + tarifs réels** | Admin | **`.xlsm!Param`/`TabTarif`** | CSV seed `tarifs.csv` |
| 4 | `REF_Regimes` | Référentiel | Régimes + équivalences + `Facturable_CPS` + `Payeur` | Admin | `Parametres!Régimes` | CSV seed |
| 5 | `REF_Pathologies` | Référentiel | Pathologies médicales / PMSI | Admin | `Parametres!Pathologies`,`Path` | CSV seed |
| 6 | `REF_Provenances` | Référentiel | Provenance + code + motif UM | Admin | `Parametres!Provenance` | CSV seed |
| 7 | `REF_Prescripteurs` | Référentiel | Prescripteurs (146) | Admin | `Parametres!Prescripteur` | CSV seed |
| 8 | `REF_Statuts` | Référentiel | Statuts DA (12) + `Exclu_facturation` | Admin | `Parametres!Statut` | CSV seed |
| 9 | `REF_Mouvements` | Référentiel | Mouvements | Admin | `Parametres!Mouvements` | CSV seed |
| 10 | `REF_Programmation` | Référentiel | Programmation (9) + `Compte_Present` | Admin | `Parametres!Programmation` | CSV seed |
| 11 | `REF_Groupes` | Référentiel | Groupes (48) | Admin | `GP_ETP_1!TabGpes` | CSV seed |
| 12 | `REF_Communes` | Référentiel | Communes + code postal | Admin | **`GP_PATIENTS!Parametres`** (56) | CSV seed |
| 13 | `REF_Motifs` | Référentiel | Motifs hospit / refus CPS / refus OraOra / annulation | Admin | `Parametres!Motifs *` | CSV seed |
| 14 | `REF_CategoriesRefus` | Référentiel/Factu | **Catégories de refus facturation** (`Droits fermés`, `JRS hors DA`, …) | Admin | **`.xlsm!TabSuiviApi[Catégorie de Refus]`** | CSV seed |
| 15 | `REF_Calendrier` | Référentiel | Date → semaine ISO / mois (2 193 dates) | Admin | `TabCalendrier` | CSV seed (ou formules) |
| 16 | `DA` | Registre | Demandes d'accord / séjours, clé pivot **N° DA** | Assistantes | `GP_ETP_1!DA` (994) ∪ `EXCEL_POLYVALENT!DA` (676) | Import unique → saisie |
| 17 | `Saisie_Presences` | Saisie | Présences ; saisie = Patient + Date + Programmation + **Parcours** | Assistantes | `*!Présences` (ETP+Poly) | Saisie + formules |
| 18 | `CONSO_Presences` | Consolidation | Table unique ETP+Poly (colonne `Parcours`) | Formules (auto) | `Saisie_Presences` | Formules / [PQ intra-classeur ultérieur] |
| 19 | `Suivi_Factures` | **Facturation** | Registre de suivi (= `Recap Facturat° futur`/`TabSuiviApi`) : montant, refus, dépôt, paiement | Assistante factu + formules | **`.xlsm!Recap Facturat° futur`** | Formules + saisie |
| 20 | `Bordereaux` | **Facturation** | Totaux par N° de facture / N° DA | Formules | **`.xlsm!Bordereaux (2)`** | Formules `SUMIFS` |
| 21 | `Facture` | **Facturation** | Modèle imprimable (1 facture) | Formules + saisie | **`.xlsm!Facturation`** | Formules `RECHERCHEX` + (option VBA `ConvNumberLetter`) |
| 22 | `Cockpit` | Pilotage | Compteurs réel + prévisionnel (ETP & Poly) | TCD | `CONSO_Presences` + `DA` | TCD |
| 23 | `CTRL_Qualite` | Pilotage | Zone d'alertes d'intégrité (compteurs d'anomalies) | Formules | tous | Formules |

> **Décision périmètre — abandons explicites** :
> - **Double fichier ETP** : `EXCEL_ETP.xlsx` abandonné ; **maître = `GP_ETP_1.xlsx`** (origine « Marion UNG »,
>   non réécrit par openpyxl ; `N° DA` identiques 994=994, 0 écart de données — R3). **Décision validée.**
> - **`EXCEL_PRESENCE_SRR.xlsx`** (ancienne consolidation manuelle, `TabPresencePoly` vide) **abandonné en bloc** :
>   remplacé par `CONSO_Presences` (R1).
> - **Power Query inter-fichiers `Z:\…`** du `.xlsm` (7 requêtes `TabDAETP`, `TabDAPolyvalent`, `TabDA`,
>   `TabPresencesETP`, `TabPresencesPoly`, `TabPrésences`, `TabPatients`) **abandonné** : remplacé par des
>   jointures **intra-classeur** (R11 — voir §6).
> - **Feuilles de cache / brouillon / extraction** abandonnées (ETP/Poly : `Détails*`, `Feuil*`, `Tri 1`,
>   `TCD Activité`, `GCC`, `SUIVI PI`, `DATA Presence ETP`, `Visu présences`, `Chiffres réel/prév.`, `*REAL/PREVI`).
> - **Colonne `Age fixe séjour`** (DA col Y, `#REF!`) abandonnée et reconstruite (§3.2, R4).
> - **Colonne Q parasite de `GP_PATIENTS`** (443 `#REF!`) **abandonnée** : non importée dans `REF_Patients` (§3.1).
> - **5 plages nommées cassées du `.xlsm`** (`Liste_Clients`, `Liste_Prestations`, `Rech_Facture` = `#REF!` ;
>   `Segment_N°_DE_FACTURE`, `Segment_N°_semaine` = `#N/A`) **non reportées** ; remplacées par des tables/segments
>   propres pointant les `REF_*`.
> - **`Colonne1`** (16e col. de Présences) abandonnée.
>
> **Ce qu'on garde** : registre `DA` (colonnes utiles), présences (3+1 champs saisis), référentiels `Parametres`
> (éclatés en `REF_*` uniques), `TabCalendrier`, **tout le moteur de facturation** (suivi, bordereaux, facture,
> tarifs), et le pilotage par TCD (`Cockpit`). La fonction VBA `ConvNumberLetter` (montant en lettres) est
> conservée **en option** sur l'onglet `Facture`.

---

## 3. Dictionnaire de données par onglet clé

Conventions : police Arial ; dates `JJ/MM/AAAA` ; `N° DA` en **texte** (préserver `AA/NNNNNN`) ; tables nommées
Excel (ListObjects). Colonnes **saisies** `[S]`, **calculées** `[C]`, **importées/admin** `[A]`.

### 3.1 `REF_Patients` — référentiel patient réel (`GP_PATIENTS!Patients `)

Source : `GP_PATIENTS.xlsx`, feuille **`Patients `** (espace final ; en-têtes en **ligne 3**, table `Table_3`
A3:N1104, ≈**1 102 patients**, 14 colonnes utiles A→N + **colonne Q parasite à abandonner**). La clé `Recherche`
existe déjà dans la source (col M), reconstruite par formule :

> `Recherche` (source) `= Nom & " " & Prénom & " " & SI(Prénom usuel="";"";"dit " & Prénom usuel & " ") & TEXTE(Date de naissance;"jj-mm-aaaa")`
> (formule réelle relevée dans `GP_PATIENTS` ; idem dans `.xlsm!TabPatients`). **C'est la clé de jointure
> patient de toute la facturation** (`.xlsm` fait `RECHERCHEX(Code recherche; TabPatients[Recherche]; …)`).

| Table | Colonnes (importées `[A]` sauf indication) |
|---|---|
| `tPatients` | `Date_saisie`, `No_administratif`, `DN`, `Nom`, `Prenom`, `Prenom_usuel`, `Date_naissance`, `Sexe`, `Commune`, `Code_postal`, `Adresse`, `N_tel`, `Recherche` (clé), `PEC`, **`Cle_Norm`** [C] |

- **`Cle_Norm`** [C] : clé normalisée (§5) appliquée à `Recherche` — utilisée pour toutes les jointures patient.
- **Nettoyage à l'import** : **NE PAS importer la colonne Q** (parasite, **443 `#REF!`**). Recalculer `Recherche`
  proprement (au lieu de figer le cache). La colonne `Code_postal` (J) est une `ArrayFormula` dans la source →
  importer la **valeur**, pas la formule, ou la dériver de `REF_Communes` via `Commune`.
- **`REF_Communes`** est alimenté par `GP_PATIENTS!Parametres` (feuille `Parametres`, table `Table_1` B3:C59,
  56 communes : `Commune`/`Code_postal`) ; **`Sexe`** {F, M} depuis `Parametres!Table_2` (E3:E5).
- Données patients réelles = **sensibles (RGPD-PF)** : seul un **CSV seed anonymisé** est committé (CLAUDE.md §2) ;
  l'import des vraies données reste hors dépôt.

### 3.2 `DA` — registre des demandes d'accord (clé pivot `N° DA`)

Inchangé v1 dans son principe : on retient les colonnes utiles des 44 d'origine, on reconstruit `Age_sejour`
(remplace `Age fixe séjour`/`#REF!`, R4) et `Nb_Je_consommes`. Clé primaire **`N_DA`** (texte `AA/NNNNNN`).

Colonnes saisies `[S]` (listes `REF_*`) : `N_DA`, `Patient`, `Parcours`, `Motif_Hospit`, `Pathologie_medicale`,
`Regime`, `Provenance`, `Prescripteur`, `Date_demande`, `Statut`, `Mouvements`, `Nb_Je_demande`,
`Cotation_demandee`, `Date_debut_demandee`, `Date_fin_demandee`, `Date_postage`, `Numero_sejour`, `PEC`,
`Nb_Je_Accorde`, `Cotation_accordee`, **`Date_debut_accorde`** (borne basse jointure), **`Date_fin_accorde`**
(borne haute), `Date_accord`, `Groupe`, `Programme`, `Date_FIN_PEC`.
Colonnes calculées `[C]` : `Cle_Patient_Norm`, `Statut_Exclu`, `Age_sejour`, `Nb_Je_consommes`, `Doublon_N_DA`.

**Formules clés `DA`** (gardes **signalantes**, jamais masquantes) :
- `Cle_Patient_Norm` [C] : normalisation §5 sur `[@Patient]`.
- `Statut_Exclu` [C] : `=SIERREUR(RECHERCHEX([@Statut];tStatuts[Statut];tStatuts[Exclu_facturation]);FAUX)`.
- `Doublon_N_DA` [C] : `=SI([@N_DA]="";"";SI(NB.SI(tDA[N_DA];[@N_DA])>1;"DOUBLON";""))` → MFC rouge (R6).
- `Age_sejour` [C] (reconstruit, **sans `#REF!`**) :
  `=SI(OU([@Patient]="";[@Date_debut_accorde]="");"";SIERREUR(DATEDIF(RECHERCHEX([@Cle_Patient_Norm];tPatients[Cle_Norm];tPatients[Date_naissance]);[@Date_debut_accorde];"y");"⚠ naiss. introuvable"))`.
- `Nb_Je_consommes` [C] : `=NB.SI.ENS(CONSO_Presences[N_DA];[@N_DA];CONSO_Presences[Est_Present];VRAI)`.

### 3.3 `Saisie_Presences` — saisie des assistantes (3 champs + `Parcours`)

Saisie `[S]` : A `Patient` (liste `REF_Patients[Recherche]`, **avertir sans bloquer**), B `Date`,
C `Programmation` (liste `REF_Programmation`), **D `Parcours`** (liste {ETP, Polyvalent} — **décision validée**,
indispensable pour rattacher la bonne DA et réparer R1).
Calculées `[C]` : E `Cle_Patient_Norm`, F `Est_Present`, **G `N_DA`** (formule maîtresse §5.2), H `Statut_Resolution`
(`OK`/`HORS_LISTE`/`SANS_DA`/`MULTI_DA`), I `Cotation`, J `Regime`, K `Nb_Je_accordes`, L `PEC`, M `Groupe`,
N `Semaine`, O `Mois`, P `Doublon_Cle` (N° DA + Date). Formules identiques v1 (RECHERCHEX sur `tDA`/`tCalendrier`),
gardes `⚠` non masquantes.

### 3.4 `CONSO_Presences` — table de consolidation unique (ETP + Poly)

Image directe de `Saisie_Presences` par référence structurée (zéro recopier-coller — R2). Colonnes : `Patient`,
`Cle_Patient_Norm`, `Date`, `Programmation`, `Est_Present`, `Parcours`, `N_DA`, `Cotation`, `Regime`,
`Nb_Je_accordes`, `Semaine`, `Mois`, `Groupe`, `PEC`, `Statut_Resolution`. **Unification ETP/Poly par `Parcours`** :
tous les comptages aval filtrent `Parcours` — plus jamais de table Poly vide (R1). C'est l'équivalent intra-classeur
du `TabPrésences` (combine PQ) du `.xlsm` (qui valait 19 240 lignes : 9 024 ETP + 10 217 Poly).

---

## 4. Couche Facturation (gros ajout v2)

Quatre onglets reproduisent **intra-classeur** le moteur de facturation du `.xlsm`, en supprimant le PQ `Z:\…`.
Toutes les jointures aval pointent désormais `tDA`, `tPatients`, `tTarif`, `CONSO_Presences` **du même classeur**.

### 4.1 `REF_Tarifs` (table `tTarif`) — tarifs RÉELS (NON inventés)

Source : `.xlsm!Param`/`TabTarif` (B1:E6). Valeurs relevées telles quelles (`[à vérifier JOPF]`) :

| `PEC` | `Code_PEC` (=Cotation) | `Tarif_XPF` | `Prestation` |
|---|---|---|---|
| Neuro | HJSN | 32000 | HJSN - Forfait Affection Neurologique |
| Amputé | HJSA | 32000 | Forfait Affection Neuro Vasculaire Amputation `[à confirmer libellé exact]` |
| Respi | HJSR | 31000 | Forfait Affection Respiratoire |
| Métabo | HJSM | 30000 | Forfait Affection Métabolique |
| Ortho | HJST | 27000 | Forfait Affection Orthopédique |

> Colonnes `tTarif` : `PEC`, `Code_PEC`, `Tarif_XPF`, `Prestation`. Le `.xlsm` nomme la colonne tarif `Tarifs` et
> la colonne code `Code PEC` ; on conserve ces **noms d'origine en interne** pour que les formules de facture
> reportées restent valides, ou on adapte les formules au renommage `[arbitrage]`. Seed : `data/referentiels/tarifs.csv`.

### 4.2 `Suivi_Factures` (= `Recap Facturat° futur` / `TabSuiviApi`)

Registre de suivi des factures. Source réelle : `.xlsm!Recap Facturat° futur`, table **`TabSuiviApi`** (A11:AL9337,
≈9 335 lignes, 124 258 formules). **Colonnes réelles relevées** (table `tSuiviFactures` cible) :

| Col | Nom réel | Statut | Formule/source cible (intra-classeur) |
|---|---|---|---|
| A | `Mois` | [C] | `=SI(ESTVIDE([@Date_debut]);"";MOIS([@Date_debut]))` (réel : `MONTH(TabSuiviApi[Date début])`) |
| B | `Date_facture` | [S] | date de facture |
| C | `N_semaine` | [S] | liste `REF_Calendrier[No_Semaine]` |
| D | `Ref` | [S/C] | référence interne |
| E | `N_DE_FACTURE` | [S] | numéro de facture (ex. `F0462026`) |
| F | `Code_recherche` | [S] | liste `REF_Patients[Recherche]` → clé patient |
| G | `Type_de_facture` | [S] | type (CPS / SS / Autres — voir §4.5) |
| H | `Cotation` | [S] | liste `REF_Cotations[Cotation]` |
| I | `Type_PEC` | [C] | dérivé PEC |
| J | `DA` (=N° DA) | [S/C] | N° DA (clé de jointure facturation) |
| K | `A_verifier` | [S] | drapeau |
| L | `N_DA_codes_acc` | [S] | N° DA + codes d'accord |
| M | `Regime` | [C] | `=RECHERCHEX([@DA];tDA[N_DA];tDA[Regime];"")` (réel : `XLOOKUP(...,TabDA[N° DA],TabDA[Régime],"")`) |
| N | `DN` | [C] | `=RECHERCHEX([@Code_recherche];tPatients[Recherche];tPatients[DN];"")` |
| O | `Nom` | [C] | `=RECHERCHEX([@Code_recherche];tPatients[Recherche];tPatients[Nom];"")` |
| P | `Prenom` | [C] | `=RECHERCHEX([@Code_recherche];tPatients[Recherche];tPatients[Prenom];"")` |
| Q | `Date_naissance` | [C] | `=RECHERCHEX([@Code_recherche];tPatients[Recherche];tPatients[Date_naissance];"")` |
| R | `Adresse` | [C] | `=RECHERCHEX([@Code_recherche];tPatients[Recherche];tPatients[Adresse];"")` (garde `=0→""`) |
| S | `Commune` | [C] | `=RECHERCHEX([@Code_recherche];tPatients[Recherche];tPatients[Commune];"")` (garde `=0→""`) |
| T | `Telephone` | [C] | `=RECHERCHEX([@Code_recherche];tPatients[Recherche];tPatients[N_tel];"")` |
| U | `Type_de_PEC` | [C] | dérivé |
| V | `Date_debut` | [S] | date début (intervalle d'accord) |
| W | `Date_fin` | [S] | date fin |
| X | `Montant` | [C] | **formule montant — voir §4.6** (réel : `SUBTOTAL(9,…)` en ligne de total) |
| Y/AC | `Categorie_Refus` | [S] | liste `REF_CategoriesRefus` (`Droits fermés`, `JRS hors DA`, …) |
| Z | `Forcer_cellule` | [S] | override manuel |
| AA | `Ratio` | [C] | `=+X/Z` (réel : `=+X10/Z10`) `[à confirmer sémantique]` |
| AB | `Date_depot` | [S] | date de dépôt CPS |
| AD | `Commentaires` | [S] | libre |
| AE | `FAE` | [S] | drapeau |
| AF | `Date_paiement` | [S] | |
| AG | `Controle_paiement` | [C] | contrôle |
| AH | `Paye_NonPaye` | [C] | statut paiement |
| AI | `Total_Paye` | [C] | cumul |
| AJ | `Date_paiement_estime` | [C] | |
| AK | `N_fact_prov` | [S] | n° facture provisoire |
| AL | `Commentaire` | [S] | libre |

> Colonnes de jointure réelles confirmées par les motifs de formules du `.xlsm` :
> `XLOOKUP(TabSuiviApi[DA]; TabDA[N° DA]; TabDA[Nb Je Accordé])`,
> `XLOOKUP(TabSuiviApi[Code recherche]; TabPatients[Recherche]; TabPatients[…])`,
> `XLOOKUP(TabSuiviApi[Cotation]; TabTarif[Code PEC]; TabTarif[PEC])`,
> `COUNTIFS(TabPrésences[semaine]; TabSuiviApi[N° semaine]; TabPrésences[N° DA]; TabSuiviApi[DA])` →
> **nombre de présences de la DA sur la semaine** (base du montant). Toutes reportées sur `tDA`/`tPatients`/
> `tTarif`/`CONSO_Presences` **du même classeur**, avec gardes signalantes au lieu des `""` silencieux d'origine.

### 4.3 `Bordereaux` (= `.xlsm!Bordereaux (2)`)

Totaux par N° de facture / N° DA. Formule réelle relevée :
`=SUMIFS(TabSuiviApi[Montant]; TabSuiviApi[N° DE FACTURE]; <N° facture>)` (et un repli
`SUM(INDIRECT("…:G"&19+n))`). Cible (table `tBordereaux`), en-têtes (réels) : `No_Bordereau`, `N_FACTURE`, `DN`,
`Periode` (« Période du JJ/MM/AAAA au JJ/MM/AAAA »), `Nb_JRS` (« nn JRS »), `Montant`. Formule cible :
`=SOMME.SI.ENS(tSuiviFactures[Montant];tSuiviFactures[N_DE_FACTURE];[@N_FACTURE])`. La liste « N° facture »
s'appuie sur `tSuiviFactures[N_DE_FACTURE]` (le `.xlsm` validait via `INDIRECT("TabSuiviApi[Bordereau]")` — on
remplace l'INDIRECT volatile par une référence structurée directe).

### 4.4 `Facture` (= `.xlsm!Facturation`) — modèle imprimable

Une facture imprimable, sélectionnée par `N° DE FACTURE` (validation = `tSuiviFactures[N_DE_FACTURE]`). Champs
réels relevés : `FACTURE N°`, `ORGANISME PAYEUR`, `Téléphone`, `Nom`, `Prénom`, `Né(e) le`, `Adresse`, `Commune`,
`Date de facture`, `Cotation`, `Nb jours accordés`, `DA`, `Nb de journées`, `Prestation`, `TOTAL`. Formules réelles :
- `=RECHERCHEX(<cotation>;tTarif[Code_PEC];tTarif[Prestation];"")` (libellé prestation),
- `=RECHERCHEX(<cotation>;tTarif[Code_PEC];tTarif[Tarif_XPF];"")` (tarif unitaire),
- `=SI(@="";"";@*@)` (montant = tarif × nb journées),
- `=RECHERCHEX(<N° facture>;tSuiviFactures[N_DE_FACTURE];tSuiviFactures[DN];"")` (en-tête patient),
- **montant en lettres** : `=MAJUSCULE(ConvNumberLetter(@))&" FRANCS CFP"` →
  **fonction VBA `ConvNumberLetter` conservée en option** (CLAUDE.md tolère le VBA anodin ; sinon, formule pure
  `[arbitrage]`). C'est le **seul code VBA réel** du `.xlsm` (Module1), autonome et reproductible à l'identique.

> Le `.xlsm` utilisait un **TCD** alimentant la facture (« ALT+F5 pour actualiser le TCD »). En cible, la facture
> lit directement `tSuiviFactures` par `RECHERCHEX` → **pas de TCD à rafraîchir manuellement**.

### 4.5 Mapping payeur **CPS / SS / Autres**

Le nom du fichier d'origine (« Factures_CPS.SS.Autres ») et la colonne `Type de facture`/`TYPE DE FACTURE` du
suivi confirment **trois canaux payeurs**. On les rattache au **régime** via `REF_Regimes[Payeur]` :

| Régime (source) | Payeur cible `[à confirmer mapping]` |
|---|---|
| RGS, RNS, RST | **CPS** |
| SS | **SS** |
| Auto-financement / autre | **Autres** |

`Suivi_Factures[Type_de_facture]` (G) peut être **dérivé** : `=RECHERCHEX([@Regime];tRegimes[Regime];tRegimes[Payeur];"⚠")`
(remplace la saisie libre actuelle, source d'incohérence). `[à confirmer mapping exact RGS/RNS/RST→CPS]`.

### 4.6 Formules clés de facturation (montant / tarif / bordereau)

- **Nb journées facturables** (par DA et semaine, repris du `.xlsm`) :
  `=NB.SI.ENS(CONSO_Presences[Semaine];[@N_semaine];CONSO_Presences[N_DA];[@DA];CONSO_Presences[Est_Present];VRAI)`.
- **Tarif unitaire** : `=RECHERCHEX([@Cotation];tTarif[Code_PEC];tTarif[Tarif_XPF];"⚠ tarif inconnu")`.
- **Montant ligne** : `Montant = Tarif_unitaire × Nb_journées_facturables`
  → `=SI([@Cotation]="";"";Tarif_unitaire*Nb_journées)`, plafonné par `Nb_Je_Accorde` (cf. règle §7-6).
- **Bordereau** : `=SOMME.SI.ENS(tSuiviFactures[Montant];tSuiviFactures[N_DE_FACTURE];[@N_FACTURE])`.

---

## 5. Mécanisme de jointure détaillé (pièce maîtresse)

### 5.1 Normalisation de la clé patient (règle d'or, CLAUDE.md §4)

Toute clé patient normalisée **avant** toute jointure : `MAJUSCULE` + `SUPPRESPACE` (`TRIM`, réduit espaces
multiples internes et de bord) + suppression de l'espace insécable `CAR(160)`. Colonne d'aide `Cle_Norm` /
`Cle_Patient_Norm` présente dans `REF_Patients`, `DA`, `Saisie_Presences` (et exploitée par `Suivi_Factures` via
`Code_recherche` → `tPatients[Cle_Norm]`) :

```
=MAJUSCULE(SUPPRESPACE(SUBSTITUE(SUBSTITUE(SUBSTITUE([@Patient];CAR(160);" ");"  ";" ");"  ";" ")))
```

Neutralise les 66 / 1 123 / 1 229 / 2 694 valeurs à espaces multiples (R7) et homogénéise la casse. La jointure
s'opère **toujours** sur la clé normalisée, jamais sur le nom brut.

### 5.2 Résolution du `N° DA` (Saisie_Presences col G) — formule maîtresse

Règle métier (CLAUDE.md §4) : présence rattachée à une DA si **Patient normalisé identique** ET
**Date ∈ [`Date_debut_accorde` ; `Date_fin_accorde`]** ET **`Parcours` identique** ET **Statut non exclu**
(`Statut_Exclu`=FAUX). Réalisée par `FILTRE` + `LET` sur clés normalisées, avec gestion explicite 0/1/N :

```
=SI([@Cle_Patient_Norm]="";"";
  SI(NB.SI(tPatients[Cle_Norm];[@Cle_Patient_Norm])=0;"⚠ HORS LISTE";
    LET(res; FILTRE(tDA[N_DA];
                (tDA[Cle_Patient_Norm]=[@Cle_Patient_Norm])
               *(tDA[Parcours]=[@Parcours])
               *(tDA[Date_debut_accorde]<=[@Date])
               *(tDA[Date_fin_accorde]>=[@Date])
               *(tDA[Statut_Exclu]=FAUX); "");
       SI(NBVAL(res)=0;"⚠ SANS DA";
          SI(NBVAL(res)>1;"⚠ MULTI DA";INDEX(res;1))))))
```

- **hors référentiel** → `⚠ HORS LISTE` (DoD §9 : nom mal orthographié = signalé, jamais vide) ;
- **0 match** → `⚠ SANS DA` (R8) ; **>1 match** → `⚠ MULTI DA` (R6/R10) ; **1 match** → `N° DA` posé.
- Toute valeur `⚠` déclenche la **MFC rouge** de ligne. **Aucun `SIERREUR` masquant** sur la chaîne de
  facturation — correction directe de la cause racine R4.

---

## 6. Comment le fichier unique tue R11 (désalignement) — par construction

**R11 (cause n°1 du désalignement)** = le déversement automatique repose sur un **Power Query inter-fichiers**
ultra-sensible : (1) retypage par **nom de colonne littéral** avec pièges typographiques
(`"Date  de FIN de PEC"` deux espaces, `"Date envoi CRH "` espace final, `"No Facture "` espace final) ;
(2) structures ETP↔Poly divergentes (dernière colonne `Date sortie admin` vs `obs`) ; (3) dépendance au **chemin
réseau `Z:\…`** ; (4) exigence de **tables nommées exactes** (`TabDA`, `TabPresences`) dans les fichiers GP ;
(5) clé `N° DA` sans normalisation (2 doublons ETP) ; (6) **28 erreurs littérales** + 5 plages nommées cassées
dans le `.xlsm`.

**Le classeur unique supprime la chaîne entière** :
- **Plus de fichiers externes ni de `Z:\…`** : la saisie (`DA`, `Saisie_Presences`) et la facturation
  (`Suivi_Factures`, `Bordereaux`, `Facture`) **cohabitent**. Les jointures deviennent des **références
  structurées intra-classeur** (`tDA`, `tPatients`, `tTarif`, `CONSO_Presences`) → §3.1, §3.4, §4.2–4.4.
  Disparition de R11.1 (chemin), R11.3 (table nommée externe), R11.4 (Excel.Workbook/File.Contents).
- **Plus de retypage par nom littéral fragile** : les en-têtes sont **figés par le build openpyxl** (noms
  cibles propres, sans espaces parasites) ; aucune assistante ne renomme un en-tête source distant. R11.1
  (typographie) neutralisé.
- **Plus de `Table.Combine` ETP/Poly** : l'unification se fait par la **colonne `Parcours`** d'une table unique
  `CONSO_Presences` (R1/R11.2). La divergence `Date sortie admin` vs `obs` (colonne AR) n'existe plus (colonne
  abandonnée, §2).
- **Clé `N° DA` fiabilisée** : `Doublon_N_DA` + MFC ; jointures renvoient `⚠ MULTI DA` au lieu de choisir
  silencieusement la 1re ligne (R5/R6 du `02b`).
- **Zéro erreur littérale** : les 28 `#REF!`/`#N/A` et 5 plages nommées cassées du `.xlsm` ne sont **pas
  reportées** ; gardes signalantes partout (DoD §9).

> En clair : **le bug de désalignement n'est pas corrigé requête par requête, il est supprimé** — il n'existe
> plus de « inter-fichiers » à désaligner. C'est l'argument central de la cible « un seul fichier ».

---

## 7. Contrôles, validations & règles de facturation

### 7.1 Listes déroulantes (chaque liste → une table `REF_*` unique, fin des INDIRECT multiples)

`Saisie_Presences[Patient]`→`tPatients[Recherche]` (**avertir sans bloquer**) ; `[Programmation]`→`tProgrammation` ;
`[Parcours]`→{ETP, Polyvalent}. `DA[*]`→`REF_*`. **Facturation** : `Suivi_Factures[Cotation]`→`tCotations`,
`[N_semaine]`→`tCalendrier`, `[Code_recherche]`→`tPatients`, `[Categorie_Refus]`→`tCategoriesRefus` ;
`Facture[N° facture]`→`tSuiviFactures[N_DE_FACTURE]` ; `Bordereaux[N_FACTURE]`→`tSuiviFactures[N_DE_FACTURE]`.
(Le `.xlsm` utilisait des `INDIRECT("TabPrésences[…]")`/`INDIRECT("TabSuiviApi[Bordereau]")` volatils → remplacés
par références structurées directes.)

### 7.2 Mises en forme conditionnelles d'alerte (ligne rouge)

`HORS LISTE` (R7/DoD) · `SANS DA` (R8) · `MULTI DA` (R6) · `Doublon_N_DA="DOUBLON"` (R6) · doublon (N° DA+Date)
(R10) · date hors calendrier · `Nb_Je_consommes > Nb_Je_accordes` (DoD) · **`Categorie_Refus` renseignée** (facture
en anomalie : `Droits fermés`, `JRS hors DA`…) · **`Tarif_unitaire="⚠ tarif inconnu"`** (cotation hors `tTarif`).

### 7.3 `CTRL_Qualite`

Compteurs `NB.SI`/`NB.SI.ENS` de chaque déclencheur ci-dessus + total lignes `Saisie`, lignes `Eligible`,
**total factures par catégorie de refus**, **montant total bordereau vs somme suivi** (contrôle de cohérence).

### 7.4 Règles de facturation (DoD §9, enrichies)

Une ligne est **`Eligible=VRAI`** ssi **toutes** ces conditions sont vraies :
1. `Est_Present = VRAI` (via `REF_Programmation[Compte_Present]`).
2. `Regime` facturable **CPS** (`REF_Regimes[Facturable_CPS]=VRAI` ; mapping payeur §4.5 ; `[à confirmer]`).
3. `Cotation` valide ∈ {HJSR, HJST, HJSN, HJSA, HJSM} (présente dans `tTarif`).
4. `N_DA` résolu (`Statut_Resolution="OK"`).
5. `Date ∈ [Date_debut_accorde ; Date_fin_accorde]` (garanti par la jointure §5, re-vérifié).
6. **`Nb_Je_consommes ≤ Nb_Je_Accorde`** (lignes au-delà du plafond signalées, non éligibles).
7. **`Categorie_Refus` vide** (pas de `Droits fermés` / `JRS hors DA` actif).

`Montant = Tarif(Cotation) × Nb_Je_facturables` (§4.6). `Motif_rejet` documente la 1re condition non satisfaite
(jamais de rejet silencieux).

**Statut format CPS** : non tranché (CLAUDE.md §7 — pas de Carte Vitale, pas de PMSI, tarifs par arrêté JOPF).
Tant que le canal GDR-DSI CPS n'est pas confirmé, la sortie est un **bordereau interne validé** + un **modèle de
facture imprimable déjà existant** (`Facture`, repris du `.xlsm`) — l'établissement dispose donc déjà du document
de facturation, seul le **canal de transmission** reste à confirmer. Action hors-code : entretien GDR/DSI CPS.

---

## 8. Traçabilité Rupture → Correction (mise à jour v2 : R11 + requalification R2)

| Rupture | Gravité | Cause d'origine | Correction par conception |
|---|---|---|---|
| **R1** Volet Poly perdu | 🔴 | `TabPresencePoly` jamais alimentée (ancienne conso SRR) | Table unique `CONSO_Presences` par `Parcours` ; Poly n'est plus une table à part. |
| **R2** *requalifiée* — transfert amont→facturation | 🔴 | **Power Query AUTOMATIQUE inter-fichiers** (pas manuel) dépendant de noms/chemins exacts | Tout **intra-classeur** : `CONSO_Presences` = dérivation par formule ; PQ `Z:\…` supprimé. |
| **R3** Double source ETP | 🟠 | 2 copies du même classeur | Une seule source `DA` ; **maître = `GP_ETP_1`** (validé). |
| **R4** `#REF!` masqué (Age + col Q patients) | 🟠 | `XLOOKUP(#REF!)`+IFERROR ; 443 `#REF!` col Q de `GP_PATIENTS` | `Age_sejour` reconstruit signalant ; **col Q non importée** ; `Recherche` recalculée. |
| **R5** Erreurs littérales | 🟡 | 2 `#VALUE!` (SUIVI PI) ; **28 erreurs + 5 plages nommées cassées du `.xlsm`** | Feuilles de cache abandonnées ; plages nommées non reportées ; gardes signalantes. |
| **R6** Doublons `N° DA` | 🟠 | XLOOKUP 1re occurrence | `Doublon_N_DA` + MFC ; jointure → `⚠ MULTI DA`. |
| **R7** Patients clés instables | 🟠 | match nom brut espaces/casse | Normalisation `Cle_Norm`/`Cle_Patient_Norm` avant toute jointure (§5). |
| **R8** Présences sans N° DA | 🟠 | jointure échouée figée vide | `Statut_Resolution="SANS DA"` + ligne rouge. |
| **R9** Orphelins / DA non consommés | 🟡 | clés tronquées ; DA Poly non consommées (effet R1) | Validation `REF_Patients` + `CTRL_Qualite` ; R1 résolu supprime les 675 Poly. |
| **R10** Doublons (N° DA + Date) | 🟡 | collages répétés / double saisie | `Doublon_Cle` + MFC + compteur `CTRL_Qualite`. |
| **R11** *(nouveau)* Désalignement PQ inter-fichiers | 🔴 | **Chemin `Z:\…` + noms tables/colonnes exacts** des fichiers GP (espaces d'en-tête, struct. ETP≠Poly) → rupture totale au moindre écart | **PQ inter-fichiers supprimé** ; jointures intra-classeur par référence structurée ; en-têtes figés par le build (§6). **Disparaît par construction.** |

---

## 9. Plan de construction v2 pour le CONSTRUCTEUR

**Scriptable openpyxl (`build/build_workbook.py`) :**
1. **Extraire les référentiels** vers `data/referentiels/*.csv` (seeds anonymisés) :
   `patients.csv` (depuis **`GP_PATIENTS!Patients `**, colonnes A→N, **SANS la colonne Q `#REF!`** ; clé
   `Recherche` recalculée), `cotations.csv`, **`tarifs.csv`** (5 valeurs réelles §4.1 : HJSN 32000, HJSA 32000,
   HJSR 31000, HJSM 30000, HJST 27000), `regimes.csv` (avec `Facturable_CPS`, `Payeur`), `pathologies.csv`,
   `provenances.csv`, `prescripteurs.csv`, `statuts.csv` (avec `Exclu_facturation`), `mouvements.csv`,
   `programmation.csv` (avec `Compte_Present`), `groupes.csv`, `communes.csv` (depuis **`GP_PATIENTS!Parametres`**),
   `motifs_*.csv`, **`categories_refus.csv`** (depuis `.xlsm!TabSuiviApi[Catégorie de Refus]`), `calendrier.csv`.
2. **Créer les onglets `REF_*`** (ListObjects nommées `tPatients`, `tCotations`, `tTarif`, …), Arial, verrouillage.
3. **Créer `DA`** : colonnes retenues + calculées (`Cle_Patient_Norm`, `Statut_Exclu`, `Doublon_N_DA`,
   `Age_sejour`, `Nb_Je_consommes`), listes `REF_*`.
4. **Créer `Saisie_Presences`** : 3+1 colonnes saisie + listes + colonnes calculées E→P (§5), MFC.
5. **Créer `CONSO_Presences`** : dérivation par formule de `Saisie_Presences`.
6. **Créer la couche Facturation** :
   - `Suivi_Factures` (`tSuiviFactures`, 38 colonnes §4.2, formules `RECHERCHEX`/`COUNTIFS` intra-classeur) ;
   - `Bordereaux` (`tBordereaux`, `SOMME.SI.ENS` §4.3) ;
   - `Facture` (modèle imprimable §4.4, `RECHERCHEX` sur `tTarif`/`tSuiviFactures` ; option VBA
     `ConvNumberLetter` — module1 reporté à l'identique si `.xlsm` cible accepté `[arbitrage]`).
7. **Créer `CTRL_Qualite`** : compteurs d'anomalies + contrôles de cohérence facturation (§7.3).
8. **Squelette `Cockpit`** (TCD posés à la main).
9. **Recalcul/contrôle** : `python build/build_workbook.py --recalc` → **zéro erreur de formule**.

**Manuel (hors openpyxl) :** import des données réelles (sensibles, jamais committées — RGPD-PF) ; TCD `Cockpit` ;
[évolution] Power Query **intra-classeur** (`powerquery/*.pq`, jamais `Z:\…`) si bascule décidée ; option VBA
`ConvNumberLetter` si format `.xlsm` retenu.

**Livrables phase 4 :** `build/build_workbook.py` ; `data/referentiels/*.csv` (dont `tarifs.csv`,
`categories_refus.csv`) ; `output/ORA_ORA_SSR_v0.xlsx` ; `powerquery/*.pq` (intra-classeur) ; `docs/04_power_query.md`.

---

## 10. Multi-utilisateurs (section honnête) `[arbitrage infrastructure ouvert]`

UN fichier unique sur **serveur de fichiers classique** = **un seul rédacteur à la fois** (verrouillage Excel).
Aujourd'hui la saisie est déjà répartie sur **2 fichiers** (`GP ETP` / `GP POLYVALENT`) → 2 personnes peuvent
saisir en parallèle ; un fichier unique **régresse** sur ce point si l'on reste en serveur de fichiers.
Coédition réelle de **4 personnes simultanées** = **SharePoint/OneDrive** requis (cloud → **réserve RGPD données
de santé en PF**, à arbitrer ; alternative = **SharePoint Server on-premise**). La **facturation est de toute
façon mono-utilisateur** par nature (une personne traite le suivi/bordereaux/factures). Modèle d'exploitation
réaliste :
- soit **classeur unique** (saisie + facturation) sur SharePoint sécurisé (coédition, RGPD à valider) ;
- soit **saisie éventuellement répartie** (2 classeurs de saisie comme aujourd'hui) + **1 classeur de facturation
  consolidé** (mais on réintroduit une frontière inter-fichiers — à éviter si possible).

`[arbitrage infrastructure ouvert]` — **non bloquant pour la conception du contenu** : l'architecture des onglets,
jointures et formules ci-dessus est valable quel que soit le choix d'hébergement. Seule la modalité de partage
(mono-poste vs SharePoint) reste à trancher avec la DSI.

---

## 11. Risques & points ouverts

- **🔴 Dépendance CPS non tranchée** (CLAUDE.md §7) : **canal/format de transmission** GDR-DSI inconnu. Le
  **modèle de facture existe déjà** (`Facture`, repris du `.xlsm`) et le **bordereau interne** est produit ; seul
  le canal CPS final reste à confirmer. Action hors-code prioritaire.
- **`[à vérifier JOPF]` — tarifs** : 32000/32000/31000/30000/27000 XPF relevés dans `.xlsm!TabTarif` ; confirmer
  qu'ils correspondent au **dernier arrêté tarifaire** avant mise en facturation. **Ne pas modifier sans source.**
- **`[à confirmer mapping]` — payeur CPS/SS/Autres** : RGS/RNS/RST→CPS, SS→SS, Auto-financement→Autres (§4.5) à
  valider avec le métier ; impacte `REF_Regimes[Payeur]` et `Facturable_CPS`.
- **`[arbitrage infrastructure ouvert]` — multi-utilisateurs** : serveur de fichiers (mono-rédacteur, régression
  vs 2 fichiers actuels) vs SharePoint (coédition, réserve RGPD-PF). Voir §10.
- **`[arbitrage]` — VBA `ConvNumberLetter`** : conserver le module VBA (→ classeur `.xlsm`) pour le montant en
  lettres, ou le réécrire en formule pure (→ classeur `.xlsx` sans macro). Le reste de la facturation est
  **100 % sans VBA**.
- **`[à confirmer]` — `Compte_Present`** (quelles Programmation comptent « présent ») et **`Facturable_CPS`** par
  régime ; **sémantique exacte de `AA = X/Z`** et de la colonne `Z Forcer la cellule` du suivi (override manuel).
- **Volume Excel 365** : ~19 000–26 000 lignes × `FILTRE`/`RECHERCHEX` par ligne est calculable mais lourd ; si
  latence, basculer `CONSO_Presences`/`Suivi_Factures` en **Power Query intra-classeur** (évolution prévue).
- **Excel 365 requis** (`FILTRE`/`RECHERCHEX`/`LET` dynamiques) ; `[à confirmer]` que toutes les assistantes en
  disposent, sinon PQ devient nécessaire dès v1.
- **`[à confirmer]` — colonnes DA des refus** (`Motifs de refus CPS/OraOra`) : à conserver si analyse des refus
  souhaitée (cohérent avec `REF_CategoriesRefus` du suivi).
