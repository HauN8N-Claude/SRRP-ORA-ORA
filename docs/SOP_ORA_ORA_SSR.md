# SOP — Mode opératoire du classeur `ORA_ORA_SSR_v0.xlsx`

**Objet** : procédure standard (Standard Operating Procedure) décrivant le fonctionnement et l'utilisation du classeur de gestion SSR — SSRP Ora Ora.
**Public** : agent de saisie / facturation, référent qualité.
**Pré-requis** : **Excel 365 (PC ou web)** obligatoire. Le classeur utilise `XLOOKUP`, `FILTER`, `LET`, `LAMBDA` qui n'existent pas hors Excel 365.

---

## 0. Principe général

Le classeur fonctionne en **3 couches** :

```
 ┌─────────────────────────────────────────────────────────────┐
 │ 1. RÉFÉRENTIELS (REF_*)  →  données stables, saisies une fois │
 └─────────────────────────────────────────────────────────────┘
                              │  (alimentent par formules)
                              ▼
 ┌─────────────────────────────────────────────────────────────┐
 │ 2. MOTEUR :  DA → Saisie_Presences → CONSO → Suivi_Factures   │
 │              → Bordereaux → Facture                           │
 └─────────────────────────────────────────────────────────────┘
                              │
                              ▼
 ┌─────────────────────────────────────────────────────────────┐
 │ 3. PILOTAGE :  Cockpit (KPI)  +  CTRL_Qualite (contrôles)     │
 └─────────────────────────────────────────────────────────────┘
```

**Règle d'or** : on ne saisit QUE dans les colonnes **sans formule** (colonnes d'entrée). Les colonnes calculées (gris/formules) se remplissent seules — ne pas les écraser.

---

## 1. Mise en route (première ouverture)

1. Ouvrir le fichier dans **Excel 365**.
2. Si Excel propose d'**activer le calcul / mettre à jour**, accepter (le fichier a été généré sans calcul → il recalcule tout à l'ouverture).
3. Vérifier qu'aucune cellule n'affiche `#NAME?` (signe que vous n'êtes pas en Excel 365).
4. Aller sur l'onglet **`CTRL_Qualite`** : c'est le tableau de bord de santé du classeur (voir §8).

---

## 2. Tenir les référentiels (`REF_*`)

À renseigner / maintenir **avant** toute saisie. Onglets clés :

| Onglet | Contenu | Colonnes saisies |
|---|---|---|
| `REF_Patients` | Fichier patients | Date_saisie, No_administratif, DN, Nom, Prénom, Prénom_usuel, Date_naissance, Sexe, Commune, Code_postal, Adresse, N_tel, PEC. Les colonnes `Recherche` et `Cle_Norm` sont calculées. |
| `REF_Tarifs` | Forfaits SSR | PEC, Code_PEC (HJSN/HJSA/HJSR/HJSM/HJST), Tarif_XPF, Prestation |
| `REF_Cotations` | Cotations ↔ PEC | Cotation, PEC |
| `REF_Regimes` | Régimes et payeur | Régime, Équivalence, Facturable_CPS (Vrai/Faux), Payeur |
| `REF_Statuts` | Statuts de DA | Statut, Exclu_facturation (Vrai = exclut la DA de la facturation) |
| `REF_Programmation` | Types de programmation | Programmation, Compte_Present (Vrai = compté présent) |
| `REF_Calendrier` | Dates → n° de semaine | Date, …, No_Semaine (ex. `S24-15`) |
| Autres | Pathologies, Provenances, Prescripteurs, Mouvements, Groupes, Communes, Motifs, Catégories de refus | listes simples |

> ⚠️ **Dates** : saisir les dates comme de **vraies dates** (jamais du texte). Voir l'avertissement §9 (défaut connu sur la version v0).

---

## 3. Enregistrer une demande d'accord — onglet `DA`

Une ligne = une prise en charge accordée pour un patient.

**Colonnes à saisir** (les listes déroulantes guident la saisie) :
`N_DA` (n° unique), `Patient` (liste = champ Recherche du patient), `Parcours` (ETP/Polyvalent), `Motif_Hospit`, `Pathologie_medicale`, `Regime`, `Provenance`, `Prescripteur`, `Date_demande`, `Statut`, `Mouvements`, `Nb_Je_demande`, `Cotation_demandee`, `Date_debut_demandee`, `Date_fin_demandee`, `Date_postage`, `Numero_sejour`, `PEC`, `Nb_Je_Accorde`, `Cotation_accordee`, `Date_debut_accorde`, `Date_fin_accorde`, `Date_accord`, `Groupe`, `Programme`, `Date_FIN_PEC`.

**Colonnes calculées automatiquement** :
- `Cle_Patient_Norm` : clé patient normalisée.
- `Statut_Exclu` : Vrai si le statut exclut la facturation (ex. refus/annulé).
- `Age_sejour` : âge du patient au début de la PEC.
- `Nb_Je_consommes` : journées réellement consommées (comptées depuis les présences).
- `Doublon_N_DA` : « DOUBLON » si le N_DA existe déjà.

---

## 4. Saisir les présences — onglet `Saisie_Presences`

Une ligne = une journée pour un patient.

**Colonnes à saisir** : `Patient` (liste), `Date`, `Programmation` (liste : Prévu, Présent, Annulation, Absent non prévenu, Journée gratuite, Indispo a prévenu, VAD), `Parcours`.

**Le moteur calcule automatiquement** :
- `Est_Present` : Vrai/Faux selon la programmation (seul « Présent » compte comme présent).
- `N_DA` : **résolution automatique** de la demande d'accord par recherche de la DA du patient, sur le bon parcours, dont la fenêtre `[Date_debut_accorde ; Date_fin_accorde]` contient la date, et **non exclue**. Sinon, le moteur affiche un drapeau :
  - `⚠ HORS LISTE` : le patient n'est pas dans `REF_Patients`.
  - `⚠ SANS DA` : aucune DA valide ne couvre cette date.
  - `⚠ MULTI DA` : plusieurs DA correspondent (à arbitrer).
- `Statut_Resolution` : `OK` ou le motif d'anomalie.
- `Cotation`, `Regime`, `Nb_Je_accordes`, `PEC`, `Groupe` : repris de la DA résolue.
- `Semaine` : n° de semaine depuis `REF_Calendrier` (sinon `⚠ hors calendrier`).
- `Mois`, `Doublon_Cle` (« DOUBLON » si même N_DA + même date déjà saisis).

---

## 5. Consolidation — onglet `CONSO_Presences`

**Aucune saisie.** Cet onglet **recopie** `Saisie_Presences` (via `INDEX`) pour servir de base aux agrégations (`COUNTIFS`/`SUMIFS`) du Suivi et du Cockpit. Il se met à jour automatiquement.

---

## 6. Facturer — onglet `Suivi_Factures`

Une ligne = une facture.

**Colonnes à saisir** : `Date_facture`, `N_semaine` (liste), `Ref`, `N_DE_FACTURE` (n° unique de facture, ex. `F0012026`), `Code_recherche` (liste patient), `Cotation` (liste), `DA`, `Date_depot`, et éventuellement `Categorie_Refus`.

**Calculé automatiquement** :
- `Mois`, `Type_de_facture` (payeur depuis le régime), `Type_PEC`, `Regime`.
- Identité patient (`DN`, `Nom`, `Prenom`, `Date_naissance`, `Adresse`, `Commune`, `Telephone`) depuis `REF_Patients`.
- `Nb_journees` : nombre de journées **présentes** pour cette DA **sur la semaine** indiquée.
- `Tarif_unitaire` (depuis `REF_Tarifs`, sinon `⚠ tarif inconnu`).
- `Montant` = `Tarif_unitaire × Nb_journees`.
- `Eligible` (Vrai/Faux) : Vrai si **régime CPS facturable** **ET** cotation valide **ET** pas de catégorie de refus **ET** `Nb_journees ≤ Nb_Je_Accorde`.
- `Motif_rejet` : explicite la cause de non-éligibilité (Régime non CPS / Cotation invalide / Refus / Plafond Nb Je dépassé).

---

## 7. Bordereaux & Facture imprimable

- **`Bordereaux`** : saisir `N_FACTURE` (liste des n° de `Suivi_Factures`). Le n° de bordereau, le `DN` et le `Montant` (somme par facture) se calculent seuls.
- **`Facture`** : en `B3`, choisir le **N° de facture** dans la liste. Toutes les informations (patient, cotation, prestation, tarif, nb de journées, total) se remplissent automatiquement. `Montant en lettres` est généré par la `LAMBDA` `MontantEnLettres` (bornée 0–999 999 999 XPF, sans décimales) suivie de « FRANCS CFP ».

---

## 8. Contrôle qualité — onglet `CTRL_Qualite`

À consulter **systématiquement** après une session de saisie. Tous les indicateurs avec seuil `0` doivent être à **0** :

| Indicateur | Seuil OK | Si ≠ 0 → action |
|---|---|---|
| Présences HORS LISTE | 0 | Créer le patient dans `REF_Patients` |
| Présences SANS DA | 0 | Créer/corriger la DA couvrant la date |
| Présences MULTI DA | 0 | Arbitrer les DA en doublon de période |
| Doublons clé (N° DA + Date) | 0 | Supprimer la présence en double |
| Doublons N° DA (registre DA) | 0 | Rendre le N_DA unique |
| Tarifs inconnus (suivi) | 0 | Corriger la cotation / `REF_Tarifs` |
| Σ Bordereaux − Σ Suivi | 0 | Vérifier l'affectation facture↔bordereau |

Les lignes `info` (catégories de refus, lignes éligibles, total saisies) sont indicatives.

Le **`Cockpit`** affiche les KPI (présences ETP / Polyvalent, total présences, nb de DA, montant total facturé). La zone TCD est à poser manuellement (Phase QA).

---

## 9. ⚠️ Défaut connu sur la version v0 (à corriger avant production)

> **Les dates des référentiels sont stockées en TEXTE** (`REF_Calendrier[Date]`, `REF_Patients[Date_naissance]`, `Date_saisie`), alors que les colonnes de saisie attendent de vraies dates.
> **Conséquence** : `Saisie_Presences[Semaine]` peut afficher « ⚠ hors calendrier » sur toutes les lignes, ce qui met `Nb_journees` à 0 et fausse les montants ; `DA[Age_sejour]` peut renvoyer une erreur.
> **À faire** : convertir ces colonnes en **vraies dates** avant exploitation. Détails et plan d'action complet dans `AUDIT_ORA_ORA_SSR.md` (constat A1).

---

## 10. Procédure de test (recette) dans Excel 365

1. Ouvrir le fichier dans Excel 365 ; laisser recalculer ; vérifier l'absence de `#NAME?`.
2. (Si le correctif §9 n'est pas encore fait) corriger les dates des référentiels.
3. Saisir un **jeu d'essai minimal** :
   - 1 patient dans `REF_Patients` (ou utiliser un patient de démo) ;
   - 1 ligne `DA` (statut non exclu, fenêtre de dates cohérente, cotation valide) ;
   - 2–3 lignes `Saisie_Presences` (programmation « Présent », dates dans la fenêtre de la DA) ;
   - 1 ligne `Suivi_Factures` (même patient, même DA, semaine correspondante).
4. Vérifier :
   - `Saisie_Presences[N_DA]` = `OK` (pas de drapeau) et `[Semaine]` = un n° valide ;
   - `Suivi_Factures[Nb_journees]` > 0, `[Montant]` correct, `[Eligible]` = Vrai ;
   - `Facture` (en sélectionnant le n°) : montant numérique **et** montant en lettres corrects ;
   - `CTRL_Qualite` : tous les seuils `0` sont à **0**.
5. Tester les cas d'erreur volontaires (patient hors liste, date hors DA, doublon) et vérifier que `CTRL_Qualite` les détecte.

---

*Document de référence associé : `AUDIT_ORA_ORA_SSR.md` (audit technique complet et plan d'action).*
