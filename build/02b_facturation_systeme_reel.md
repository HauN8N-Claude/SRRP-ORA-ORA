# 02b — Système réel de facturation (5e fichier `.xlsm`) et chaîne complète

**Objet.** Analyse du fichier `Tableau_de_suivi_Factures_CPS.SS.Autres_2025_2.xlsm` (le « fichier facturation »
fourni en complément), qui **révèle le système de production réel** et la cause exacte du désalignement
GP ↔ facturation. Ce document complète `01_analyse.md` et **corrige** la cartographie de `02_flux_et_diagnostic.md`
(qui, faute de ce fichier, avait conclu à un transfert *manuel* ; le transfert est en réalité **automatique par
Power Query inter-fichiers**). Données vérifiées par extraction du code M (DataMashup), de `connections.xml`,
des tables et des feuilles. `[à confirmer]` pour toute hypothèse.

---

## 1. Le système réel (production)

Le classeur de facturation est alimenté **automatiquement par Power Query** depuis **3 fichiers externes**
situés sur le lecteur réseau `Z:\2 PLANNIF & PROGRAMMATION\GESTION DES PATIENTS\` :

```
   Z:\...\GESTION DES PATIENTS\
   ├── GP ETP.xlsx          (saisie ETP : tables TabDA, TabPresences)
   ├── GP POLYVALENT.xlsx   (saisie Polyvalent : TabDA, TabPresences)
   └── GP Patients.xlsx     (référentiel patients : TabPatients)   ← NON FOURNI
                    │
                    │  Power Query  (Excel.Workbook(File.Contents("Z:\...")))
                    ▼
   Tableau_de_suivi_Factures_CPS.xlsm  (LE fichier facturation)
   ├── Requêtes PQ : TabDAETP, TabDAPolyvalent, TabDA(=combine),
   │                 TabPresencesETP, TabPresencesPoly, TabPrésences(=combine), TabPatients
   ├── Recap Facturat° futur (TabSuiviApi) : suivi des factures + refus
   ├── Bordereaux (2) : totaux par N° DA
   ├── Facturation : modèle de facture imprimable
   ├── Param (TabTarif) : tarifs des 5 forfaits
   └── Macros VBA (vbaProject.bin)
```

**Correspondance avec les fichiers fournis** (`[à confirmer]` sur les égalités exactes) :
- `GP ETP.xlsx` ≈ **`GP_ETP_1.xlsx`** (origine « Marion UNG » préservée) ; `EXCEL_ETP.xlsx` = copie/export.
- `GP POLYVALENT.xlsx` ≈ **`EXCEL_POLYVALENT.xlsx`**.
- `GP Patients.xlsx` = **fichier manquant** (non fourni) — source du référentiel patient.
- `EXCEL_PRESENCE_SRR.xlsx` = **ancienne** consolidation manuelle (sans connexion PQ — d'où R2 « manuel »),
  vraisemblablement antérieure / parallèle à ce dispositif PQ.

> Le « **déversement automatique** » décrit par l'utilisateur = ces **7 requêtes Power Query inter-fichiers**.
> Il est bien automatique (bouton « Actualiser tout »), mais fragile (cf. §3).

---

## 2. Le moteur Power Query (code M extrait)

Chaque requête lit une **table nommée précise** d'un fichier externe, **retype colonne par colonne par nom**,
**sélectionne** un sous-ensemble de colonnes, puis (pour les présences) **filtre `Programmation = "Présent"`** :

| Requête | Source | Table lue | Colonnes conservées | Filtre |
|---|---|---|---|---|
| `TabDAETP` | `GP ETP.xlsx` | `TabDA` | Patient, Motif Hospit, Régime, N° DA, PEC, Nb Je Accordé, Cotation accordé, Parcours | — |
| `TabDAPolyvalent` | `GP POLYVALENT.xlsx` | `TabDA` | idem (8 colonnes) | — |
| `TabDA` | — | `Table.Combine(TabDAETP, TabDAPolyvalent)` | 8 colonnes | — |
| `TabPresencesETP` | `GP ETP.xlsx` | `TabPresences` | Patient, Date, Programmation, N° DA, PEC, Cotation, semaine, Mois | `="Présent"` |
| `TabPresencesPoly` | `GP POLYVALENT.xlsx` | `TabPresences` | idem | `="Présent"` |
| `TabPrésences` | — | `Table.Combine(...)` | 8 colonnes | — |
| `TabPatients` | `GP Patients.xlsx` | `TabPatients` | toutes (14 col.) | — |

**Volumes obtenus** (cache du fichier) : `TabDA` 1 691 lignes (1 008 ETP + 684 Poly = **registre unifié**,
ce que la cible visait), `TabPresencesETP` 9 024, `TabPresencesPoly` **10 217** (≠ la table vide de
`EXCEL_PRESENCE_SRR` — ici le Polyvalent **est** présent, R1 est donc un défaut de l'ancienne consolidation,
pas du dispositif PQ), `TabPrésences` 19 240.

---

## 3. Cause exacte du désalignement GP ↔ facturation

Le PQ est **dépendant des noms de colonnes exacts** et de la **présence de tables nommées précises** dans les
fichiers GP. Points de rupture vérifiés :

1. **Retypage par nom littéral, ultra-sensible** : `Table.TransformColumnTypes` cible des noms exacts dont
   certains portent des **pièges typographiques** : `"Date  de FIN de PEC"` (deux espaces), `"Date envoi CRH "`
   (espace final), `"No Facture "` (espace final). Si une assistante corrige/renomme un en-tête dans le fichier
   GP, l'étape échoue ou la colonne devient vide → **donnée fausse en facturation, silencieusement**.
2. **Structures sources divergentes ETP vs Poly** : dernière colonne `Date sortie admin` (ETP) vs `obs` (Poly) ;
   types déclarés différents (`Pathologie médicale` text/any, `Programme` Int64/any…). Le `Table.Combine`
   n'aligne **par nom** que parce qu'on a réduit aux **8 colonnes communes** avant — toute colonne ajoutée en
   amont casserait l'alignement.
3. **Dépendance au chemin réseau `Z:\…`** : si le lecteur `Z:` n'est pas monté pareil sur chaque poste, ou si
   un fichier GP est déplacé/renommé, **toutes les requêtes tombent** (`DataSource.Error`).
4. **Table nommée requise** : `Source{[Item="TabDA",Kind="Table"]}` exige une table **exactement** nommée
   `TabDA`/`TabPresences` dans le GP ; un simple décalage de structure de saisie rompt le lien.
5. **Clé `N° DA`** : toutes les jointures aval (`Recap`, `TabPrésences`, `Bordereaux`) font
   `XLOOKUP(... ; TabDA[N° DA] ; ...)`. Les **2 doublons N° DA ETP** (R6) et tout **espace parasite** renvoient
   la mauvaise ligne ou `""`. Aucune normalisation de clé.
6. **28 erreurs littérales** dans le classeur (cache) — formules de facturation non gardées.

> En clair : le déversement est automatique **mais positionné/typé par nom exact sur des fichiers externes au
> format instable**. Dès qu'un en-tête, un type, un chemin ou un nom de table bouge côté GP, la facturation
> reçoit des colonnes décalées ou vides — exactement le symptôme décrit (« pas bien alignés, données fausses,
> formules qui buggent »).

---

## 4. Les fonctions de facturation présentes (à regrouper dans le fichier unique)

| Onglet | Table | Rôle | Détail |
|---|---|---|---|
| `Recap Facturat° futur` | `TabSuiviApi` (~9 335 l., 124 258 formules) | **Suivi des factures** | XLOOKUP sur `TabDA[N° DA]` → Nb Je Accordé, Régime, Cotation ; colonne `Catégorie de Refus` (ex. `Droits fermés`, `JRS hors DA`) |
| `Bordereaux (2)` | — (≈217 l.) | **Bordereaux** | `SUMIFS(TabSuiviApi[Montant] ; TabSuiviApi[N° DA] ; …)` : totaux par DA |
| `Facturation` | — | **Facture imprimable** | Modèle : `FACTURE N°`, Organisme payeur, N° Sécurité Sociale, Nom/Prénom/Né(e) le/Adresse ; `XLOOKUP(... ; TabTarif[Code PEC] ; TabTarif[Prestation])` |
| `Param` | `TabTarif` | **Tarifs** | 5 forfaits (voir §5) |
| `TabCalendrier` | — | calendrier | date → semaine |

**Macros VBA** : présentes (`vbaProject.bin`) — rôle **probable `[à confirmer]`** : actualisation des requêtes,
génération/impression des bordereaux et factures, navigation. **Code non encore lu** (nécessite un extracteur
type `oletools/olevba` — à autoriser si l'analyse du VBA est souhaitée).

---

## 5. Tarifs réels (source : `Param`/`TabTarif` du fichier) — NON inventés

| Code (Cotation) | PEC | Prestation | Tarif (XPF) |
|---|---|---|---|
| HJSN | Neuro | Forfait Affection Neurologique | 32 000 |
| HJSA | Amputé | Forfait Affection Neuro Vasculaire Amputation | 32 000 |
| HJSR | Respi | Forfait Affection Respiratoire | 31 000 |
| HJSM | Métabo | Forfait Affection Métabolique | 30 000 |
| HJST | Ortho | Forfait Affection Orthopédique | 27 000 |

> Ce sont les valeurs en service dans le fichier de l'établissement. `[à vérifier au JOPF]` pour confirmer
> qu'elles correspondent au dernier arrêté tarifaire en vigueur avant mise en facturation.

---

## 6. Conséquences pour la cible « un seul fichier »

1. **Regrouper dans un classeur unique** la saisie (DA + présences ETP & Poly), le référentiel patient, et
   **tout le moteur de facturation** (suivi `TabSuiviApi`, bordereaux, modèle de facture, tarifs `TabTarif`).
2. **Supprimer le Power Query inter-fichiers `Z:\…`** : les jointures deviennent **intra-classeur** (références
   structurées par clé) → la cause n°1 du désalignement (§3.1–3.4) **disparaît par construction**.
3. **Normaliser la clé patient** et **fiabiliser `N° DA`** (doublons signalés) → §3.5 traité.
4. **Conserver les fonctions utiles** : suivi des factures + catégories de refus, bordereaux par DA, modèle de
   facture, table des tarifs. **Abandonner** la dépendance `Z:\`, les colonnes pièges (espaces), les feuilles
   de cache redondantes.
5. **Fichier `GP Patients.xlsx` manquant** : à récupérer pour bâtir `REF_Patients` complet (sinon, référentiel
   reconstruit depuis `DA[Patient]`, moins riche).
6. **Tension multi-utilisateurs inchangée** : un classeur unique local ne permet pas 4 saisies simultanées sans
   SharePoint ; à arbitrer (cf. échanges). Le suivi de facturation, lui, est mono-utilisateur par nature.

---

## 7. Mise à jour du diagnostic (`02`)

- **R2 requalifiée** : le transfert amont→facturation n'est **pas manuel** mais **Power Query inter-fichiers**
  dépendant de noms/chemins exacts (fragilité différente, conséquence identique : facturation fausse silencieuse).
- **R1 précisée** : la perte du Polyvalent est un défaut de l'**ancienne** consolidation `EXCEL_PRESENCE_SRR`
  (`TabPresencePoly` vide), **pas** du dispositif PQ (où `TabPresencesPoly` = 10 217 lignes). La cible unifie
  de toute façon par `Parcours`.
- **R6/R7 confirmées critiques** : `N° DA` est l'unique clé des jointures de facturation, sans normalisation.
- **Nouveau R11** 🔴 : dépendance au **chemin réseau `Z:\…` + noms de tables/colonnes exacts** des fichiers GP
  → rupture totale des requêtes au moindre écart (déplacement, renommage, espace d'en-tête).
