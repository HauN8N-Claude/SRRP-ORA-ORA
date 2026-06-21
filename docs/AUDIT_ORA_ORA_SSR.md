# Audit du classeur `ORA_ORA_SSR_v0.xlsx`

**Application** : SSRP Ora Ora — gestion des présences, des demandes d'accord (DA / prise en charge) et de la facturation SSR (Soins de Suite et de Réadaptation) en Polynésie française (montants en **XPF / Francs CFP**, payeur **CPS**).

**Fichier audité** : `ORA_ORA_SSR_v0.xlsx`
**Date de l'audit** : 2026-06-21
**Auteur du fichier (métadonnées)** : `openpyxl` (généré par script — **jamais ouvert ni recalculé dans Excel**)
**Méthode** : audit **statique** (lecture du XML / des formules via openpyxl). Aucune exécution réelle dans Excel n'a été possible dans l'environnement d'analyse.

> ⚠️ **Important** : ce classeur utilise des fonctions **exclusives à Excel 365** (`XLOOKUP`, `FILTER`, `LET`, `LAMBDA`). Il **doit impérativement être ouvert dans Excel 365 (PC ou web)**. Il ne fonctionnera pas correctement dans Excel 2019/2021 hors-365, ni dans LibreOffice/Google Sheets (rendu partiel, erreurs `#NAME?`).

---

## 1. Vue d'ensemble

| Élément | Valeur |
|---|---|
| Nombre d'onglets | 24 (23 visibles + 1 masqué `_Listes`) |
| Tableaux structurés | 20 |
| Plages nommées | 16 (dont la `LAMBDA` `MontantEnLettres`) |
| Cellules contenant une formule | 1 574 |
| Macros VBA | **Aucune** (pas de `vbaProject.bin`) |
| Fonctions Excel 365 | `XLOOKUP` ×764, `LET` ×30, `FILTER` ×30, `LAMBDA` (1 plage nommée) |
| Listes déroulantes (validations) | 21 plages |
| Mises en forme conditionnelles | 7 règles (DA, Saisie, Suivi, CTRL_Qualite) |

### Cartographie des onglets

**Référentiels (saisis manuellement, en lecture par le moteur)**
`REF_Patients`, `REF_Cotations`, `REF_Tarifs`, `REF_Regimes`, `REF_Pathologies`, `REF_Provenances`, `REF_Prescripteurs`, `REF_Statuts`, `REF_Mouvements`, `REF_Programmation`, `REF_Groupes`, `REF_Communes`, `REF_Motifs`, `REF_CategoriesRefus`, `REF_Calendrier`, `_Listes` (masqué : valeurs de Parcours `ETP` / `Polyvalent`).

**Moteur métier**
- `DA` — registre des demandes d'accord / prises en charge (PEC).
- `Saisie_Presences` — saisie quotidienne des présences (résolution automatique de la DA).
- `CONSO_Presences` — recopie/consolidation de `Saisie_Presences` pour les agrégations.
- `Suivi_Factures` — une ligne par facture (calcul journées, tarif, montant, éligibilité).
- `Bordereaux` — regroupement des factures.
- `Facture` — modèle de facture imprimable (montant en lettres).
- `Cockpit` — indicateurs de pilotage (KPI).
- `CTRL_Qualite` — contrôles d'intégrité (doivent tous être à 0).

---

## 2. Synthèse des constats

| # | Sévérité | Constat | Impact |
|---|---|---|---|
| A1 | 🔴 **Critique** | **Toutes les dates des référentiels sont stockées en TEXTE** alors que les colonnes de saisie attendent de vraies dates | Casse en cascade le calcul des semaines, des journées et des montants |
| A2 | 🟠 Élevé | `DA[Age_sejour]` : `DATEDIF` sur une date de naissance au format texte | `#VALUE!` ou âge erroné (dépend de la locale) |
| A3 | 🟡 Moyen | `REF_Tarifs[Tarif_XPF]` stocké en texte (`"32000"`) | Fonctionne par coercition mais fragile ; risque si valeur non numérique |
| A4 | 🔵 Faible | `REF_Calendrier` colonnes `Annee`/`Mois`/`Semaine_Iso` = texte (`"2023.0"`, `"1.0"`) | Colonnes inutilisables pour des TCD ; non utilisées par les formules |
| A5 | ℹ️ Info | Fichier jamais recalculé (généré par openpyxl) | Le premier calcul a lieu à l'ouverture Excel 365 — c'est normal |
| A6 | ℹ️ Info | Capacité limitée à **30 lignes** par table moteur (lignes 2→31) | À étendre avant mise en production |
| A7 | ℹ️ Info | Zone TCD du Cockpit et canal de transmission CPS non finalisés | Marqué « Phase QA » dans le fichier lui-même |

---

## 3. Constats détaillés

### 🔴 A1 — Dates des référentiels stockées en TEXTE (défaut bloquant)

**Faits mesurés :**
- `REF_Calendrier[Date]` : **2192 / 2192** cellules sont du **texte** (`"01/01/2023"`), 0 vraie date.
- `REF_Patients[Date_naissance]` : **7 / 7** en texte.
- `REF_Patients[Date_saisie]` également en texte.
- À l'inverse, les colonnes de **saisie** de dates sont formatées en vraie date `DD/MM/YYYY` et typées numériques :
  `Saisie_Presences[Date]`, `DA[Date_*]`, `Suivi_Factures[Date_facture / Date_debut / Date_fin]`.

**Pourquoi c'est grave :** un `XLOOKUP` / une comparaison entre **un nombre** (vraie date saisie par l'utilisateur) et **une chaîne** (`"01/01/2023"` du calendrier) **ne correspond jamais**.

**Chaîne de rupture (cascade) :**
1. `Saisie_Presences[Semaine]` =
   `XLOOKUP([@Date], tCalendrier[Date], tCalendrier[No_Semaine], "⚠ hors calendrier")`
   → renverra **« ⚠ hors calendrier » sur quasiment toutes les lignes** dès que l'utilisateur saisit une vraie date.
2. `Suivi_Factures[Nb_journees]` =
   `COUNTIFS(CONSO_Presences[N_DA], …, CONSO_Presences[Semaine], [@N_semaine], …)`
   → ne trouve aucune semaine valide → **0 journée**.
3. `Suivi_Factures[Montant]` = `Tarif × Nb_journees` → **0 ou vide**.
4. `Bordereaux[Montant]`, `Cockpit`, `Facture` → **montants faux**.

**Correctif recommandé :** convertir en **vraies dates** toutes les colonnes date des référentiels — en particulier `REF_Calendrier[Date]` et `REF_Patients[Date_naissance]` (et `Date_saisie`). Vérifier ensuite que `Saisie_Presences[Semaine]` retrouve bien les n° de semaine.

---

### 🟠 A2 — `DA[Age_sejour]` : DATEDIF sur date texte

Formule :
```
=IF(OR([@Patient]="",[@Date_debut_accorde]=""),"",
   IFERROR(DATEDIF(XLOOKUP([@Cle_Patient_Norm],tPatients[Cle_Norm],tPatients[Date_naissance],""),
                   [@Date_debut_accorde],"y"),"⚠ naiss. introuvable"))
```
`DATEDIF` attend des **dates sérielles**. La date de naissance étant du texte (cf. A1), le résultat dépend de la locale d'Excel et peut produire `#VALUE!` (capté par `IFERROR` → « ⚠ naiss. introuvable ») ou un âge faux.
**Correctif :** corollaire de A1 — passer `REF_Patients[Date_naissance]` en vraie date.

---

### 🟡 A3 — `REF_Tarifs[Tarif_XPF]` stocké en texte

Valeurs `"32000"`, `"31000"`… stockées en **texte** (type `s`, format General).
Utilisé dans `Suivi_Factures[Montant] = Tarif_unitaire × Nb_journees` et `Facture!B21 = B18 × B19`.
Excel **coerce** automatiquement le texte numérique dans une multiplication, donc le calcul « marche »… **tant que** la valeur reste convertible. C'est fragile (un espace, un caractère parasite → `#VALUE!`) et nuit à l'alignement/format.
**Correctif :** stocker les tarifs en **nombres** et appliquer un format `# ##0 "XPF"`.

---

### 🔵 A4 — Artefacts texte dans `REF_Calendrier`

Colonnes `Annee` (`"2023.0"`), `Mois` (`"1.0"`), `Semaine_Iso` (`"52.0"`) : floats écrits en texte par le script de génération.
Non utilisées par les formules (seules `Date` et `No_Semaine` le sont), **mais** inexploitables pour de futurs TCD.
**Correctif :** régénérer en entiers, ou ignorer si non utilisées.

---

### ℹ️ A5 — Aucun calcul en cache

Toutes les valeurs calculées sont vides (`creator = openpyxl`). C'est **normal** pour un fichier généré par script : Excel recalculera tout à la **première ouverture**. C'est précisément l'étape de test à réaliser dans Excel 365.

### ℹ️ A6 — Capacité 30 lignes

Les tables `DA`, `Saisie_Presences`, `CONSO_Presences`, `Suivi_Factures`, `Bordereaux` sont dimensionnées **lignes 2 à 31** (30 lignes de données). Aucune donnée réelle n'est saisie (seul `REF_Patients` contient 7 patients fictifs de démonstration). Étendre les tables avant mise en production (les formules `@`-structurées se propageront automatiquement).

### ℹ️ A7 — Éléments « Phase QA » non finalisés (notés dans le fichier)

- `Cockpit` : la zone TCD est à poser **manuellement** sur `CONSO_Presences` et `DA`.
- `Facture` : le **canal de transmission CPS n'est pas tranché** ; ce modèle est un bordereau/facture interne, pas le format CPS final.
- `MontantEnLettres` : bornée **0 à 999 999 999 XPF**, sans décimales.

---

## 4. Points positifs (conception solide)

- **Clé patient normalisée cohérente** : `Cle_Norm = UPPER(TRIM(SUBSTITUTE(... CHAR(160) ...)))` appliquée des deux côtés (référentiel et saisies) → les jointures patient sont robustes (espaces insécables, doublons d'espaces gérés).
- **Résolution automatique de la DA** dans `Saisie_Presences[N_DA]` via un `FILTER` (patient + parcours + date dans la fenêtre accordée + DA non exclue), avec gestion explicite des cas `HORS LISTE`, `SANS DA`, `MULTI DA`. Le `FILTER` est **consommé dans un `LET`** (pas de débordement/spill dans la colonne de tableau → pas de `#SPILL!`).
- **Garde-fous `IFERROR`** sur les recherches, valeurs de repli explicites (`"⚠ tarif inconnu"`, `"⚠ hors calendrier"`).
- **Contrôles d'intégrité** centralisés (`CTRL_Qualite`) : doublons, hors-liste, sans DA, multi-DA, tarifs inconnus, cohérence Σ Bordereaux − Σ Suivi.
- **Éligibilité facturation** explicite (`Suivi_Factures[Eligible]` + `Motif_rejet`) : régime CPS facturable, cotation valide, pas de refus, plafond de journées accordées respecté.
- **Validations (listes déroulantes)** sur tous les champs à valeur contrôlée, alimentées par plages nommées.

---

## 5. Plan d'action recommandé (par priorité)

1. **(Critique)** Convertir en vraies dates : `REF_Calendrier[Date]`, `REF_Patients[Date_naissance]`, `REF_Patients[Date_saisie]`. → débloque semaines, journées, montants, âge.
2. **(Moyen)** Convertir `REF_Tarifs[Tarif_XPF]` en nombres + format `# ##0 "XPF"`.
3. **(Faible)** Régénérer `REF_Calendrier[Annee/Mois/Semaine_Iso]` en entiers (ou supprimer si inutiles).
4. **(Test 365)** Ouvrir dans Excel 365, laisser recalculer, saisir un **jeu d'essai** (1 patient → 1 DA → quelques présences → 1 facture) et vérifier `CTRL_Qualite` = 0 partout.
5. **(Prod)** Étendre les tables au-delà de 30 lignes ; poser les TCD du Cockpit ; trancher le format CPS.

---

## 6. Limite de cet audit

Audit **statique** uniquement. Il ne remplace pas un test réel dans Excel 365 : seules l'ouverture et le recalcul dans Excel valident les fonctions 365 (LAMBDA récursive, FILTER, spill) et le rendu. La procédure de test pas-à-pas est décrite dans `SOP_ORA_ORA_SSR.md`.
