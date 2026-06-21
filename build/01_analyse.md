# 01 — Analyse indépendante des 4 classeurs Excel (SSR Polynésie)

**Objet.** Description factuelle, fichier par fichier, de l'existant des 4 classeurs Excel fournis dans `fichiers_source/` (lecture seule). Chaque fichier est analysé isolément : structure des feuilles, champs-clés, formules, validations, tables, TCD, anomalies. Aucune comparaison transversale ni recommandation n'est faite ici — ce n'est pas l'objet de ce livrable. Toute interprétation est marquée `[à confirmer]`.

**Méthode.** Audit JSON pré-généré (`build/_audit/*.json`) comme ossature, complété par des vérifications openpyxl (`read_only`, `data_only=False` pour les formules / `data_only=True` pour les valeurs en cache) et inspection du `.xlsx` dézippé (TCD via `xl/pivotTables`, VBA via `xl/vbaProject.bin`).

**Note de lecture sur les statistiques.** Les colonnes alimentées par formule (XLOOKUP/FILTER/…) renvoient, en lecture `data_only`, la **valeur en cache** au moment de la dernière sauvegarde Excel. Quand ce cache est vide, le décompte « non vide » peut être nul alors que la colonne contient bien des formules. Ce point est signalé là où il survient.

**Aucun des 4 fichiers ne contient de macro VBA** : absence de `xl/vbaProject.bin` vérifiée dans les 4 archives `.xlsx`.

---

## 1. EXCEL_PRESENCE_SRR.xlsx

### Vue d'ensemble
- **10 feuilles**, toutes visibles. Taille fichier ~1,5 Mo.
- Volumétrie max : feuille `TabPresences` ≈ **13 599 lignes** (table `Table_2`, A1:P13600).
- **Aucun TCD** (pas de `xl/pivotTables` dans l'archive), **aucune plage nommée**, **201 formules** au total (faible), **0 erreur littérale**.
- Rôle apparent `[à confirmer]` : classeur de **consolidation/visualisation des présences** (ETP + Polyvalent) avec feuilles de synthèse hebdomadaires et tables de référentiel.

### Tableau des feuilles
| Feuille | État | Dim. | Lignes données | Col. | Formules | Validations | Tables | Rôle apparent `[à confirmer]` |
|---|---|---|---|---|---|---|---|---|
| Visu présences ETP | visible | M177 | 175 | 13 | 6 (+ array) | 2 | 0 | Vue hebdo présences ETP |
| Visu présences Poly | visible | M177 | 175 | 13 | 6 (+ array) | 2 | 0 | Vue hebdo présences Polyvalent |
| Chiffres réel | visible | I25 | 19 | 9 | 67 | 1 | 0 | Synthèse chiffrée réalisée (COUNTIFS) |
| Chiffres prévisionnels | visible | I19 | 15 | 9 | 63 | 1 | 0 | Synthèse prévisionnelle |
| Feuil1 | visible | I19 | 15 | 9 | 59 | 1 | 0 | Doublon/brouillon de « Chiffres » `[à confirmer]` |
| TabPresencePoly | visible | O2 | **2** (≈1 réelle) | 15 | 0 | 0 | Table_1 (A1:O10510) | Table présences Poly — quasi vide |
| TabPresences | visible | P13600 | **13 599** | 16 | 0 | 0 | Table_2 (A1:P13600) | Table présences principale |
| TabPresencesETP | visible | P3898 | 3 898 | 16 | 0 | 0 | Table_3 (A1:P3898) | Table présences ETP |
| TabCalendrier | visible | E2193 | 2 193 | 5 | 0 | 0 | Table_4 (A1:E2193) | Référentiel calendrier (dates → n° semaine) |
| Feuil2 | visible | — | 0 | 0 | 0 | 0 | 0 | Vide |

### Feuilles importantes — en-têtes et clés

**`TabPresences`** (en-têtes ligne 1, types réels) : `Patient`(texte), `Date`(date), `Programmation`(texte, ex. « Présent »), `No Facture`, `N° DA`(texte, ex. `22/004641`), `PEC`(texte, ex. `J1`), `Cotation`(texte), `REGIMES`(texte), `Dates Accord`(texte, ex. `18/02/22 au 17/02/23`), `Nb Je accordés`(float), `semaine`(texte, ex. `S23-01`), `Nb actes`, `Mois`(texte, ex. `01/23`), `Parcours`(texte), `Groupe`, `Colonne1`.

Stats champs-clés `TabPresences` (sur ~13 598 lignes de données) :
| Champ | non vides | distincts | doublons | vides | anomalies |
|---|---|---|---|---|---|
| Patient (A) | **5** | 4 | 1 | 13 594 | colonne quasi vide (renseignée seulement sur les premières lignes) |
| Date (B) | 13 503 | 845 | 12 658 | 96 | — |
| N° DA (E) | 13 106 | 852 | 12 254 | 493 | 493 vides |
| Cotation (G) | 13 105 | **1** (uniquement `HJSM`) | 13 104 | 494 | valeur unique |
| REGIMES (H) | 13 597 | 3 (`CPS`=12 064 / `autre`=841 / `SS`=692) | — | 2 | — |
| Parcours (N) | 13 599 | **1** (`ETP`) | — | 0 | valeur unique |
| Groupe (O) | 13 106 | 19 | — | 493 | **1 195 valeurs avec espace initial** ; 11 506 valeurs à `0` (défaut) |

> `[à confirmer]` : dans `TabPresences`, l'identifiant patient (col A) n'est presque jamais renseigné ; les lignes semblent identifiées par le couple **N° DA + Date**. La colonne `Parcours` vaut systématiquement `ETP` et `Cotation` systématiquement `HJSM` — cette table paraît donc cantonnée à un parcours/cotation unique malgré son nom générique `[à confirmer]`.

**`TabPresencesETP`** : mêmes en-têtes. Stats : `Patient` non vides 3 841 / distincts 355, dont **1 229 valeurs à espaces multiples** ; `N° DA` 3 897 non vides / 385 distincts ; `Cotation` valeur unique (1 distinct).

**`TabPresencePoly`** : table déclarée A1:O10510 mais **2 lignes seulement** (1 en-tête + ~1 donnée). Quasi vide. `[à confirmer]` : table préparée mais non alimentée.

**`TabCalendrier`** : `Date`, `Annee`, `Mois`, `Semaine Iso`, `No Semaine` (ex. `S22-52`) — référentiel de conversion date→semaine, 2 193 lignes.

### Formules caractéristiques
- Feuilles `Chiffres réel`, `Chiffres prévisionnels`, `Feuil1` : massivement des **COUNTIFS croisant deux tables**, ex. :
  `=COUNTIFS(TabPresencePoly!C:C,"Présent",TabPresencePoly!K:K,$A6,TabPresencePoly!M:M,B$5,TabPresencePoly!H:H,"CPS")+COUNTIFS(TabPresencesETP!…,"Présent",…)` ; et des `=SUM(...)`.
- `Visu présences ETP/Poly` : formules de date en cascade `=G5+1` et formules-tableau (ArrayFormula) sur ~170 lignes ; en-tête dynamique `="Semaine du "&TEXT($G$5;…)`.

### Validations de données
Listes déroulantes pointant des colonnes de tables via INDIRECT : `INDIRECT("TabCalendrier[No Semaine]")` (cellules de filtre B1/B2), `INDIRECT("TabPresencesETP[Cotation]")`, `INDIRECT("TabPresencesPoly_2[Cotation]")`.

### TCD
Aucun.

### Points d'attention factuels (ce fichier)
- `TabPresencePoly` quasi vide (≈1 ligne) alors que `TabPresences`≈13 599 et `TabPresencesETP`≈3 898.
- `TabPresences` : col `Patient` renseignée sur **5 lignes** seulement (clé patient absente) ; `Cotation` et `Parcours` mono-valeur.
- `Groupe` : **1 195 valeurs avec espace de tête** et 11 506 cellules à `0` (anomalie de saisie / défaut).
- `TabPresencesETP.Patient` : **1 229 cellules à espaces multiples** (risque de non-correspondance de clé `[à confirmer]`).
- `Feuil1` paraît dupliquer `Chiffres réel` ; `Feuil2` vide.
- 493–494 lignes sans `N° DA`/`Cotation`/`Groupe` dans `TabPresences`.

---

## 2. EXCEL_ETP.xlsx

### Vue d'ensemble
- **23 feuilles**, toutes visibles. Taille ~4,0 Mo.
- Volumétrie max : `Présences` ≈ **13 599 lignes** ; `SUIVI PI` 10 086 ; `DATA Presence ETP` 10 732 ; `Détails JRS` 6 881.
- **34 238 formules** (audit), **2 erreurs littérales** (#VALUE!), **4 TCD**, aucune plage nommée, **pas de VBA**.
- **4 TCD** présents (`xl/pivotTables/pivotTable1..4`), cibles : feuilles **`JRS REAL Année`**, **`JRS REAL Mois`**, **`JRS REAL`**, **`JRS PREVI`**.
- Rôle apparent `[à confirmer]` : classeur **métier ETP complet** = registre des demandes d'accord (`DA`), table de présences calculée (`Présences`) et tableaux de bord (TCD `JRS …`, feuille `GCC`).

### Tableau des feuilles
| Feuille | État | Dim. | Lignes données | Col. | Formules | Validations | Tables | Rôle apparent `[à confirmer]` |
|---|---|---|---|---|---|---|---|---|
| Détails JRS | visible | B6882 | 6 881 | 2 | 0 | 0 | 0 | Liste semaine→présence (alim. cache) |
| Détails1 | visible | P4 | 3 | 16 | 0 | 0 | Table_1 | Extrait (en-têtes type Présences) |
| GCC | visible | C79 | **7** | 3 | 0 | 0 | 0 | Sortie type TCD « présence par patient » (cache) |
| JRS REAL Année / JRS REAL Mois / JRS REAL / JRS PREVI | visible | A1 | 1 | 1 | 0 | 0 | 0 | **Emplacements de TCD** (jours réalisés/prévis.) |
| SUIVI PI | visible | C10089 | 10 086 | 3 | 0 | 0 | 0 | Suivi (Nombre de Date / Nombre de PEC) |
| Détails2 | visible | AQ4 | 3 | 43 | 0 | 0 | Table_2 | Extrait large (mêmes champs que DA) |
| TCD Activité | visible | G24 | 23 | 7 | 0 | 0 | 0 | TCD résiduel (« Date début accordé ») |
| **DA** | visible | AR1010 | **1 008** | 44 | **7 042** | **12** | Table_3 (B3:AR1010) | **Registre demandes d'accord** |
| Feuil2 / Feuil3 / Feuil4 / Feuil5 | visible | K… | 3 484 / 3 003 / 492 / 183 | 11 | 0 | 0 | 0 | Extractions/TCD figés (« Étiquettes de colonnes ») |
| **Présences** | visible | P13599 | **13 599** | 16 | **27 196** | 4 | 0 | **Table présences calculée par formules** |
| Parametres | visible | M227 | 225 | 13 | 0 | 0 | 7 tables | Référentiels (motifs, statuts, prescripteurs…) |
| DATA Presence ETP | visible | P10732 | 10 732 | 16 | 0 | 0 | 0 | Données présences brutes (valeurs) |
| TabGpes | visible | A48 | 48 | 1 | 0 | 0 | Table | Référentiel groupes |
| TabPatients␣ | visible | B1099 | 1 099 | 2 | 0 | 0 | Table | Référentiel patients (Date naiss. + clé Recherche) |
| TabCalendrier | visible | E2193 | 2 193 | 5 | 0 | 0 | Table | Référentiel calendrier |
| Paramètres Fixes | visible | O32 | 20 | 15 | 0 | 0 | 6 tables | Référentiels fixes (régimes, parcours, programmes…) |
| Feuil1 | visible | — | 0 | 0 | 0 | 0 | 0 | Vide |

### Feuille `DA` (registre des demandes d'accord) — en-têtes ligne 3
44 colonnes (B→AR) : `Patient`, `Motif Hospit`, `Pathologie PMSI`, `Pathologie médicale`, `Régime`, `Provenance`, `Prescripteur`, `Date de demande`, `Date Rep Pres`, `CHIR`, `Date opération`, `Statut`, `Mouvements`, `Motifs de refus CPS`, `Motif refus OraOra`, `Nb Je demandé`, `Cotation demandée`, `Date DEBUT demandée`, `Date FIN demandée`, `Date de postage`, **`N° DA`**(V), `Numéro de séjour`, `PEC`, `Age fixe séjour`, `Nb Je Accordé`, `Cotation accordé`, `Date début accordé`, `Date fin accordé`, `Date d'accord`, `Nbre de J effectués`, `Date de FIN de PEC`, `Commentaire sortie`, `Date envoi notif fin hospit`, `Date envoi CRH`, `PEC intiale`, `Programme`, `Groupe initial`, `Date début présence`, `Date fin présence`, `Nb Je présence`, `Cotation Présence`, `Parcours`, `Date sortie admin`.

Stats champs-clés `DA` (≈1 007 lignes) :
| Champ | non vides | distincts | doublons | vides | note |
|---|---|---|---|---|---|
| Patient (B) | 22 (cache) | 20 | 2 | 985 | majorité en formule (cache vide) ; 3 valeurs à espaces multiples |
| **N° DA (V)** | **996** | **994** | **2** | 11 | clé quasi unique ; 2 doublons ; 1 valeur à espace de fin |
| Cotation accordé (AA) | 1 002 (cache) | 1 | — | 5 | cache mono-valeur `HJSM` |
| Régime (F) | 975 | 5 | — | 32 | — |
| Date d'accord (AD) | 738 | 232 | — | 269 | 269 vides |
| Nb Je Accordé (Z) | 951 | 9 | — | 56 | — |

> `N° DA` (col V) est le **candidat clé primaire** du registre DA (994 distincts / 996 non vides). Plusieurs colonnes (`Patient`, `Cotation accordé`, dates dérivées) sont calculées et leur cache est partiel.

**Formules `DA` (exemples réels) :**
- `Age fixe séjour` (col Y) : `=IFERROR(DATEDIF(XLOOKUP(DA!$B4,#REF!,#REF!,"inconnu"),DA!$AN4,"y"),"")` — **contient une référence `#REF!`** présente sur ~**1 006 lignes** (masquée par `IFERROR`, renvoie ""). 
- `Nb Je présence` (col AO) : `=COUNTIFS('Présences'!$C$2:$C$13599,"Présent",'Présences'!$E$2:$E$13599,DA!$V4)` (1 007 occ.).
- Recopies conditionnelles : `=IF(DA!$AB4="",DA!$S4,DA!$AB4)` (date début), `=IF(DA!$Z4>0,DA!$Z4,DA!$Q4)` (Nb Je), `=XLOOKUP(DA!$AP4,'Paramètres Fixes'!…)`.
- Décompte XLOOKUP `DA` : **4 027** ; COUNTIFS **1 007** ; DATEDIF **1 006** ; IFERROR **1 006** ; `#REF!` **1 006**.

**Validations `DA` (12 listes déroulantes)** via INDIRECT sur tables de `Parametres`/`Paramètres Fixes` : `TabPath[[Pathologies ]]`, `TabProv[Provenance]`, `TabRefusOraOra[…]`, `TabGpes[Groupe]`, `TabStatut[Statut]`, `TabMouv[Mouvements]`, `TabReg[Régimes]`, `TableauPrescripteursETP[Prescripteur]`, `TabPatients[Recherche]`, `TabMotifsRefus[…]`, `TabCotations[Cotation]`, `TabMotifHosp[Motifs Hospi]`.

### Feuille `Présences` (table calculée) — en-têtes ligne 1
Mêmes 16 colonnes que `TabPresences` de PRESENCE_SRR. **27 196 cellules de formules** (audit) — en réalité de très nombreuses colonnes sont en formule. Les formules réalisent des **jointures vers `DA` clé `N° DA` (col E)** et vers `TabCalendrier` :
- col E `N° DA` : `=IF($N2="","",IFERROR(_xlfn._xlws.FILTER(DA!$V$4:$V$1010,(DA!$B$4:$B$1010=$A2)*(DA!$AM…<=$B2)*…),""))` → **FILTER** retrouvant le N° DA à partir du patient et de la date.
- col G `Cotation` : `=XLOOKUP($E2,DA!$V$4:$V$1010,DA!$AP$4:$AP$1010,"")`.
- col H `REGIMES`, col I `Dates Accord`, col J `Nb Je accordés`, col O `Groupe` : `XLOOKUP` sur `DA` clé `N° DA`.
- col K `semaine` : `XLOOKUP` sur `TabCalendrier`.
- col M `Mois` : `=TEXT($B2,"mm/aa")` ; col N `Parcours` : `="ETP"`.
- Décompte fonctions (cellules formule) : XLOOKUP ~95 182, FILTER 13 598, COUNTIFS 13 598, IFERROR 13 598, TEXT 27 194 ; **aucun VLOOKUP/RECHERCHEV, aucun INDEX/MATCH**.

> `[à confirmer]` : `Présences` est une feuille **dérivée**, recalculée par formules à partir de `DA` (clé N° DA) et `TabCalendrier`. En lecture des valeurs en cache, `N° DA`, `Cotation`, `Groupe` ressortent vides (cache non matérialisé), tandis que `Patient`, `Date`, `REGIMES`(=`CPS`), `Parcours`(=`ETP`) sont peuplés ; `Patient` présente **2 694 valeurs à espaces multiples**.

### Champs-clés du fichier (synthèse)
- **`N° DA`** : clé primaire du registre `DA` (994 distincts) et clé de jointure dans `Présences`.
- **`Patient`** (clé `Recherche` = Nom + date naissance) : référentiel `TabPatients␣` (1 099 patients) ; nombreuses occurrences à espaces multiples dans `Présences`/`TabPresencesETP`.
- **Date** + **semaine** (`TabCalendrier`) : axe temporel.

### Points d'attention factuels (ce fichier)
- **`#REF!` dans ~1 006 formules** de la colonne `Age fixe séjour` (DA) — masqué par IFERROR mais référence cassée.
- **2 `#VALUE!` littéraux** en `SUIVI PI` (lignes 6760 et 6763, col A).
- `DA.N° DA` : **2 doublons** + 1 valeur à espace de fin ; 11 vides.
- Feuilles « brouillon/extraction » multiples : `Détails1`, `Détails2`, `Feuil2`/`Feuil3` (sorties TCD figées « Étiquettes de colonnes »), `Feuil4`/`Feuil5` (extraits partiels), `Feuil1` vide, `TCD Activité` résiduel.
- `Présences.Patient` : **2 694 valeurs à espaces multiples**.
- Présence simultanée de données **calculées** (`Présences`) et **brutes** (`DATA Presence ETP`, 10 732 lignes) au même format `[à confirmer]`.

---

## 3. GP_ETP_1.xlsx

### Vue d'ensemble
- **23 feuilles** homonymes de EXCEL_ETP, toutes visibles. Taille ~4,0 Mo.
- **34 238 formules** (identique au total ETP), **2 #VALUE!**, **4 TCD** (mêmes cibles `JRS …`), pas de VBA.
- Propriétés fichier distinctes : **creator = « Marion UNG », created = 2023-12-12** (alors que EXCEL_ETP a `creator=openpyxl`, sans date d'origine — réécrit par l'outil d'audit/export `[à confirmer]`).
- Rôle apparent `[à confirmer]` : **même classeur métier ETP** que le précédent.

### Tableau des feuilles
Structure strictement homonyme à EXCEL_ETP (mêmes 23 feuilles, mêmes dimensions générales, mêmes tables/validations). Les volumes par feuille sont identiques à l'unité près à ceux du § 2, **sauf** :
| Feuille | EXCEL_ETP | GP_ETP_1 | Constat |
|---|---|---|---|
| GCC | 7 lignes non vides | **77 lignes non vides** | divergence de **cache** (cf. ci-dessous) |
| Détails JRS | — | — | mêmes 6 881 lignes, mais ~6 738 lignes aux **valeurs en cache différentes** |
| DA (col Patient) | 22 valeurs en cache | **3 valeurs en cache** | divergence de cache de la colonne calculée `Patient` |

### Comparaison cellule-à-cellule avec EXCEL_ETP (formules vs valeurs)
Vérification openpyxl, feuille par feuille :
- **Formules identiques** : `SUIVI PI`, `DATA Presence ETP`, `TabPatients␣`, `Feuil2`, `Présences` (0 ligne de différence en comparant le **texte des formules**, y compris ArrayFormula). Les feuilles `Présences` et `DA` ont la **même structure de formules** dans les deux fichiers.
- **Divergences réelles, sur valeurs en cache (pas sur formules)** :
  - **`GCC`** : ETP = 7 lignes non vides (lignes 7-8 puis « Total général » en 79) ; GP_ETP_1 = **77 lignes non vides** (liste complète patient par patient, lignes 7→78). Même en-tête (« PRESENCE / GESTION ETP », « Mois 12/24 », table croisée patient × « Présent »). → GP_ETP_1 conserve le **résultat détaillé** en cache ; ETP a un cache réduit/replié.
  - **`Détails JRS`** : ~6 738 lignes dont les valeurs en cache diffèrent (ex. col A `S24-21` présent côté ETP, absent côté GP).
  - **`DA`** : 19 lignes diffèrent sur la **valeur en cache** de la colonne `Patient` (col B) ; ETP en matérialise 22, GP_ETP_1 seulement 3.

> `[à confirmer]` : les deux fichiers portent les **mêmes formules** ; les écarts observés tiennent à l'**état de recalcul / cache de valeurs** enregistré (notamment feuilles de sortie `GCC`, `Détails JRS`). Description faite séparément conformément à la consigne.

### Champs-clés, formules, validations
Identiques au § 2 (mêmes en-têtes `DA` 44 col., mêmes 12 validations INDIRECT, même `N° DA` clé primaire ≈ 994 distincts ; même `Présences` à base XLOOKUP/FILTER/COUNTIFS clé `N° DA` ; même `#REF!` sur ~1 006 formules `Age fixe séjour`).

### Points d'attention factuels (ce fichier)
- Mêmes anomalies structurelles que EXCEL_ETP : `#REF!` (~1 006) en `DA.Age fixe séjour`, 2 `#VALUE!` en `SUIVI PI`, doublons `N° DA`, feuilles brouillon/extraction.
- **`GCC` détaillée** (77 lignes en cache) ici, contre repliée (7) dans EXCEL_ETP — état de cache divergent.
- Métadonnées d'origine préservées (creator « Marion UNG », 2023-12-12).

---

## 4. EXCEL_POLYVALENT.xlsx

### Vue d'ensemble
- **12 feuilles**, toutes visibles. Taille ~2,5 Mo.
- Volumétrie max : `Présences` ≈ **12 479 lignes** ; `TabCalendrier` 2 193 ; `DA` 684 ; `Parametres` 147.
- **28 367 formules** (audit), **0 erreur littérale**, **3 TCD** (`xl/pivotTables/1..3`), pas de plage nommée, pas de VBA.
- TCD cibles : feuilles **`JRS REAL`**, **`JRS PREVI`**, **`MOIS REAL`**.
- Rôle apparent `[à confirmer]` : **équivalent « Polyvalent » du classeur ETP** — registre `DA`, `Présences` calculée, référentiels `Parametres`, tableaux de bord (`JRS …`, `Tri 1`).

### Tableau des feuilles
| Feuille | État | Dim. | Lignes données | Col. | Formules | Validations | Tables | Rôle apparent `[à confirmer]` |
|---|---|---|---|---|---|---|---|---|
| Tri 1 | visible | C14 | 12 | 3 | 0 | 0 | 0 | TCD figé (« Mois » / éléments) |
| Détails1 | visible | O5 | 4 | 15 | 0 | 0 | Table_1 (A3:O23) | Extrait type Présences |
| ANS REAL | visible | B4 | 2 | 2 | 0 | 0 | 0 | Mini-synthèse annuelle |
| JRS REAL / JRS PREVI | visible | A1 | 1 | 1 | 0 | 0 | 0 | **Emplacements de TCD** |
| MOIS REAL | visible | — | 0 | 0 | 0 | 0 | 0 | **Emplacement de TCD** (vide hors cache) |
| **Présences** | visible | O12480 | **12 479** | 15 | **24 956** | 3 | 0 | **Table présences calculée par formules** |
| **DA** | visible | AR686 | **684** | 44 | **3 411** | 11 | Table_2 (B3:AR686) | **Registre demandes d'accord** |
| TabCalendrier | visible | E2193 | 2 193 | 5 | 0 | 0 | Table_3 | Référentiel calendrier |
| TabPatients | visible | B3 | **3** | 2 | 0 | 0 | Table_4 | Référentiel patients **quasi vide** |
| Paramètres fixes | visible | — | 0 | 0 | 0 | 0 | 0 | Vide |
| Parametres | visible | AL148 | 147 | 38 | 0 | 0 | **17 tables** | Référentiels (très étoffé) |

### Feuille `DA` — en-têtes ligne 3
44 colonnes B→AR, identiques à celles d'ETP **sauf la dernière** : AR = **`obs`** (contre `Date sortie admin` dans ETP). `N° DA` en col V.

Stats champs-clés `DA` (≈683 lignes) :
| Champ | non vides | distincts | doublons | vides | note |
|---|---|---|---|---|---|
| Patient (B) | 683 | 327 | 356 | 0 | renseigné ; **66 valeurs à espaces multiples** |
| **N° DA (V)** | 676 | **676** | **0** | 7 | clé **strictement unique** sur ce fichier |
| Cotation accordé (AA) | 571 | 4 | — | 112 | — |
| Régime (F) | 682 | 4 | — | 1 | — |
| Date d'accord (AD) | 486 | 227 | — | 197 | 197 vides |
| Nb Je Accordé (Z) | 539 | 16 | — | 144 | — |

> Ici `Patient` est correctement renseigné (683/683) et `N° DA` est **clé unique parfaite** (676 distincts / 0 doublon).

**Formules `DA`** (mêmes motifs qu'ETP) : `=IF(DA!…="",…)` (2 046), `=IF(DA!…>0,…)` (683), `=COUNTIFS('Présences'!…,"Présent",'Présences'!…,DA!…)` (682). Décompte : XLOOKUP **2 726**, COUNTIFS 682, IFERROR 683, DATEDIF 683, **`#REF!` 683** (même formule `Age fixe séjour` à référence cassée, masquée par IFERROR). Aucun VLOOKUP/INDEX-MATCH.

**Validations `DA` (11 listes)** : `TabPath`, `TabProv`, `TabRefusOraOra`, `TabStatut`, `TabReg`, `TabMouv`, `TableauPrescripteursETP`, `TabPatients[Recherche]`, `TabMotifsRefus`, `TabCotations`, `TabMotifHosp` (INDIRECT vers `Parametres`).

### Feuille `Présences` — en-têtes ligne 1
15 colonnes (`Patient`→`Groupe`, sans `Colonne1`). **24 956 formules**. Mêmes mécaniques de jointure qu'ETP :
- col M `Mois` `=TEXT($B…,"mm/aa")` (12 478 occ.) ; col N `Parcours` `="Polyvalent"` (12 478) ; XLOOKUP **74 868**, FILTER **12 478**, COUNTIFS **12 478**, IFERROR **12 478**.
- Validations : `INDIRECT("TabDA[Patient]")` (col A), `INDIRECT("TabProg[Programmation]")` (col C), `INDIRECT("TabMotifAnnul[Motifs annulation]")` (col D).
- En cache : `Patient`(12 479, dont **1 123 à espaces multiples**), `Date`, `REGIMES`(2 distincts), `Parcours`(`Polyvalent`) peuplés ; `N° DA`, `Cotation`, `Groupe` non matérialisés en cache (1 valeur).

### Feuille `Parametres`
Très riche : **17 tables structurées** (`Table_5`…`Table_21`) couvrant `Motifs Hospi`, `Pathologies`, `Provenance`/`Motif d'entrée UM`/`Code de provenance`, `Régimes`/`Equivalence régimes`, `Statut`, `Programmation`, `Thérapeute`, `Path`, `Prescripteur` (146 lignes), `Mouvements`, `Motifs de refus CPS`, `Motifs de refus OraOra`, `Cotation`/`Parcours`, `Programme`/`N° programme`, `Commune`/`Code postal` (56 lignes), `Motifs annulation`.

### Champs-clés du fichier (synthèse)
- **`N° DA`** (DA col V) : clé primaire **unique** (676/676) et clé de jointure de `Présences`.
- **`Patient`** : renseigné dans `DA` (327 distincts) et `Présences` ; anomalies d'espaces (66 + 1 123 occ.).
- **Date / semaine / Mois** via `TabCalendrier`.

### Points d'attention factuels (ce fichier)
- **`#REF!` dans 683 formules** `DA.Age fixe séjour` (masqué par IFERROR), comme dans ETP.
- **`TabPatients` quasi vide** (3 lignes) alors que la validation `DA.Patient` s'appuie dessus `[à confirmer]` ; le référentiel patient effectif paraît être la liste `DA[Patient]`.
- Feuilles vides/figées : `MOIS REAL` (vide hors cache TCD), `Paramètres fixes` (vide), `Tri 1` (TCD figé).
- `Présences.Patient` : **1 123 valeurs à espaces multiples** ; `DA.Patient` : 66.
- En-tête `DA` divergent en col AR (`obs` vs `Date sortie admin` côté ETP).
- `Présences` : valeurs en cache partielles (`N° DA`/`Cotation`/`Groupe` non matérialisés).

---

## Tableau récapitulatif (1 ligne / fichier)

| Fichier | Feuilles | Volumétrie max | TCD / VBA | Champs-clés | Anomalies majeures |
|---|---|---|---|---|---|
| EXCEL_PRESENCE_SRR | 10 | `TabPresences` 13 599 l. | 0 TCD / pas de VBA | N° DA + Date ; (Patient quasi absent) ; Cotation & Parcours mono-valeur | `TabPresencePoly` quasi vide ; Patient renseigné sur 5 l. seult ; Groupe : 1 195 espaces de tête + 11 506 à `0` ; ETP.Patient 1 229 espaces multiples ; ~493 N° DA vides |
| EXCEL_ETP | 23 | `Présences` 13 599 l. | 4 TCD / pas de VBA | **N° DA** (DA, 994 distincts) ; Patient (`Recherche`) ; Date/semaine | `#REF!` sur ~1 006 formules `Age fixe séjour` ; 2 `#VALUE!` (SUIVI PI) ; 2 doublons `N° DA` ; nombreuses feuilles brouillon/extraction ; Présences.Patient 2 694 espaces multiples |
| GP_ETP_1 | 23 | `Présences` 13 599 l. | 4 TCD / pas de VBA | identiques à EXCEL_ETP | mêmes anomalies qu'ETP ; **divergence de cache** : `GCC` 77 l. (vs 7), `Détails JRS` ~6 738 l. de valeurs différentes, `DA.Patient` cache 3 vs 22 ; formules identiques |
| EXCEL_POLYVALENT | 12 | `Présences` 12 479 l. | 3 TCD / pas de VBA | **N° DA** (DA, **676 distincts, 0 doublon**) ; Patient ; Date/semaine | `#REF!` sur 683 formules `Age fixe séjour` ; `TabPatients` quasi vide (3 l.) ; feuilles vides (`MOIS REAL`, `Paramètres fixes`) ; Présences.Patient 1 123 espaces multiples ; en-tête AR `obs` |
