# MÉMOIRE DU PROJET — Classeur Excel unique de facturation CPS (SSR ORA ORA)

> Fichier mémoire / point d'avancement du projet. À lire pour reprendre le travail sans tout relire.
> Dernière mise à jour : 2026-06-21. Branche de dev : `claude/happy-hypatia-qttbs3`.

---

## 1. Mission (rappel)

Remplacer les fichiers Excel séparés de l'hôpital de jour SSR (ORA ORA / Pirae / Tahiti, Polynésie)
par **UN seul classeur Excel** qui fiabilise la chaîne **présence → demande d'accord (DA) → facturation CPS**,
en gardant **Excel comme unique interface** des assistantes. Détail des contraintes : `CLAUDE.md`.

---

## 2. Le système réel découvert (ground truth)

Le dispositif de production réel a été élucidé en analysant **6 fichiers** (les 4 initiaux + le fichier de
facturation `.xlsm` + le référentiel patient). Le « déversement automatique » est un **Power Query inter-fichiers**.

```
Lecteur réseau  Z:\2 PLANNIF & PROGRAMMATION\GESTION DES PATIENTS\
   ├── GP ETP.xlsx          (saisie ETP : tables TabDA, TabPresences)        = GP_ETP_1.xlsx (fourni)
   ├── GP POLYVALENT.xlsx   (saisie Polyvalent : TabDA, TabPresences)         = EXCEL_POLYVALENT.xlsx (fourni)
   └── GP Patients.xlsx     (référentiel patient : TabPatients, clé Recherche)= GP_PATIENTS.xlsx (fourni)
                    │  Power Query (7 requêtes, Excel.Workbook(File.Contents("Z:\...")))
                    ▼
   Tableau_de_suivi_Factures_CPS.SS.Autres_2025_2.xlsm   (LE moteur de facturation)
   ├── Suivi des factures (Recap Facturat° futur / TabSuiviApi) + catégories de refus
   ├── Bordereaux (totaux par N° DA / N° facture)
   ├── Facture (modèle imprimable) + macro VBA anodine (montant en lettres)
   └── Param / TabTarif (tarifs des 5 forfaits)
```

### Correspondance des fichiers fournis
| Fichier fourni | Rôle réel |
|---|---|
| `GP_ETP_1.xlsx` | = « GP ETP.xlsx » — **saisie ETP, source maître ETP** (origine « Marion UNG ») |
| `EXCEL_ETP.xlsx` | copie/export d'ETP — **abandonné** (redondant) |
| `EXCEL_POLYVALENT.xlsx` | = « GP POLYVALENT.xlsx » — saisie Polyvalent |
| `GP_PATIENTS.xlsx` | = « GP Patients.xlsx » — référentiel patient (clé `Recherche`) |
| `Tableau_suivi_Factures_CPS_2025.xlsm` | **moteur de facturation** (Power Query + suivi + bordereaux + facture + tarifs) |
| `EXCEL_PRESENCE_SRR.xlsx` | **ancienne** consolidation manuelle (Polyvalent vide) — **obsolète, abandonnée** |

---

## 3. Diagnostic — 11 ruptures (détail : `build/02_flux_et_diagnostic.md` + `build/02b_...md`)

| # | Gravité | Rupture | Mesure |
|---|---|---|---|
| R1 | 🔴 | Volet Polyvalent perdu (ancienne conso `TabPresencePoly` vide) | 1 ligne vs ~10 216 |
| R2 | 🔴 | Transfert amont→facturation **automatique par Power Query** (requalifié : pas manuel) | — |
| R3 | 🟠 | Double source ETP (EXCEL_ETP vs GP_ETP_1, données identiques) | 994=994, 0 écart |
| R4 | 🟠 | `#REF!` masqué par IFERROR (`Age fixe séjour`) + 443 `#REF!` col Q patients | 2012+1366 / 443 |
| R5 | 🟡 | Erreurs littérales (`#VALUE!` SUIVI PI ; 28 erreurs + 5 plages cassées du `.xlsm`) | — |
| R6 | 🟠 | Doublons `N° DA` (XLOOKUP prend la 1re) | 2 (ETP) ; 0 (Poly) |
| R7 | 🟠 | Clés patient instables (espaces multiples / casse) | 1 123–2 694 valeurs |
| R8 | 🟠 | Présences sans `N° DA` (non facturables) | 493 / 13 599 |
| R9 | 🟡 | Orphelins / DA non consommées | 1 orphelin ; 675 Poly (=effet R1) |
| R10 | 🟡 | Doublons (N° DA + Date) | 25 (+10) couples |
| **R11** | 🔴 | **Désalignement = Power Query inter-fichiers `Z:\…`** (noms tables/colonnes exacts, pièges d'espaces d'en-tête, struct. ETP≠Poly, chemin réseau) — **LE bug décrit par l'utilisateur** | — |

**Cause racine** : pas de chaîne intégrée ; jointures par position/nom exact sur fichiers externes instables,
clés non normalisées, `IFERROR` qui masque les erreurs au lieu de les signaler.

---

## 4. Décisions validées par le client (NE PAS re-questionner)

| Sujet | Décision |
|---|---|
| Moteur de jointure | **Formules dynamiques natives** (FILTER/XLOOKUP/LET + clé patient normalisée) pour la v1 100 % scriptable, **+ Power Query intra-classeur documenté** en évolution |
| Colonne `Parcours` en saisie | **Ajoutée** (4e champ liste {ETP, Polyvalent}) |
| Validation patient | **Avertir sans bloquer** (`⚠ HORS LISTE` en rouge, saisie permise) |
| Source maître ETP | **GP_ETP_1** |
| Format de fichier | **`.xlsx` sans macro** ; montant en lettres = **formule pure (LAMBDA)** |
| Tarifs | Repris du fichier réel (HJSN 32000, HJSA 32000, HJSR 31000, HJSM 30000, HJST 27000 XPF) — `[à vérifier JOPF]` |

### ⚠️ Tension non résolue — multi-utilisateurs
Le client veut **4 personnes en simultané sur le même fichier**, hébergé en **local sur serveur de fichiers**
(pas de Microsoft 365). **Techniquement incompatible** : un `.xlsx` sur partage de fichiers se **verrouille à un
seul rédacteur**. La coédition réelle exige **SharePoint/OneDrive** (cloud → réserve RGPD données de santé, ou
SharePoint Server sur site). **Arbitrage d'infrastructure ouvert** (à trancher avec la DSI) — n'affecte PAS le
contenu du classeur, qui est valable quel que soit l'hébergement. Voir `build/03_architecture.md` §10.

---

## 5. Architecture cible (détail : `build/03_architecture.md` v2)

**Classeur unique, 24 onglets**, jointures **intra-classeur** par clé normalisée (supprime R11 par construction) :
- **Référentiels (15)** : `REF_Patients, REF_Cotations, REF_Tarifs, REF_Regimes, REF_Pathologies, REF_Provenances,
  REF_Prescripteurs, REF_Statuts, REF_Mouvements, REF_Programmation, REF_Groupes, REF_Communes, REF_Motifs,
  REF_CategoriesRefus, REF_Calendrier` (+ feuille cachée `_Listes`).
- **Registre** : `DA` (clé pivot `N° DA`, colonnes calculées `Cle_Patient_Norm`, `Statut_Exclu`, `Age_sejour`
  reconstruite sans #REF!, `Nb_Je_consommes`, `Doublon_N_DA`).
- **Saisie** : `Saisie_Presences` (saisie = Patient + Date + Programmation + Parcours ; reste calculé ; MFC rouge).
- **Consolidation** : `CONSO_Presences` (table unique ETP+Poly par `Parcours` — répare R1).
- **Facturation** : `Suivi_Factures` (suivi + refus + montant), `Bordereaux` (SUMIFS), `Facture` (imprimable).
- **Pilotage** : `Cockpit` (TCD), `CTRL_Qualite` (compteurs d'anomalies).

**Règle d'or** : clé patient normalisée `UPPER(TRIM(SUBSTITUTE(…CHAR(160)…)))` AVANT toute jointure.
**Principe anti-cause-racine** : les gardes **signalent en rouge**, ne masquent jamais.

---

## 6. État de la construction (Phase 4) et QA (Phase 5)

**Construit** (`build/build_workbook.py`, idempotent) → `output/ORA_ORA_SSR_v0.xlsx` :
24 onglets, 20 tables nommées, 21 validations, 5 MFC rouges, 16 plages nommées + LAMBDA `MontantEnLettres`.

**QA (`build/04_qa.md`)** — **9/9 cas synthétiques passent** (logique répliquée en Python, recalcul Excel
impossible dans le sandbox). Défauts corrigés :
- **D1 🔴** table `CONSO_Presences` mal nommée (`tConso`) → 7 formules en `#NAME?` → **corrigé**.
- **D2/D3 🟠** formats date/montant → corrigés.
Vérif statique : 0 référence cassée (1574 formules), 0 nom de fonction français, 0 erreur littérale stockée.

**Limite honnête** : la validation des tableaux dynamiques (FILTER/XLOOKUP/LET) **doit se faire dans Excel 365
réel** (protocole manuel pas-à-pas dans `build/04_qa.md` §7). LibreOffice du sandbox ne charge pas les .xlsx.

---

## 7. Definition of Done — statut

| DoD (CLAUDE.md §9) | Statut |
|---|---|
| Un seul `.xlsx` sans erreur de formule | 🟡 à valider en Excel 365 (0 erreur statique) |
| Saisie présence → DA/cotation/régime/Nb Je auto | 🟡 logique prouvée (Python), à valider Excel |
| Nom hors-liste → signalé rouge, jamais vide | ✅ (formule + MFC + validation warning) |
| `Facturation` ne liste que les dossiers valides | 🟡 logique prouvée, à valider Excel |
| `Cockpit` compteurs réels ETP **et** Poly | 🟡 (R1 corrigé par `Parcours` ; TCD à poser) |
| Référentiels uniques verrouillés, pas de données réelles | ✅ (seeds synthétiques, REF protégées) |
| Procédure CPS documentée OU bordereau interne | ✅/🟡 (bordereau + facture produits ; canal CPS hors-code) |

---

## 8. Points ouverts / actions hors-code

- **🔴 Dépendance CPS** : canal/format de transmission GDR-DSI **non tranché** → sortie = bordereau interne +
  facture imprimable en attendant. Action métier : entretien GDR/DSI CPS.
- **Infrastructure multi-utilisateurs** : serveur de fichiers (mono-rédacteur) vs SharePoint (coédition, RGPD). DSI.
- **`[à confirmer]`** : tarifs vs arrêté JOPF ; mapping payeur RGS/RNS/RST→CPS ; `Compte_Present` (quelles
  programmations comptent présent) ; `Facturable_CPS` par régime.
- **Validation finale en Excel 365** : exécuter le protocole de `build/04_qa.md` §7.
- **Volume** : tables livrées à ~30 lignes-modèles ; éprouver la latence FILTER/XLOOKUP sur ~19 000 lignes
  réelles (sinon basculer `CONSO_Presences`/`Suivi_Factures` en Power Query intra-classeur).

---

## 9. Phases & livrables

| Phase | Agent | Livrable | Statut |
|---|---|---|---|
| 1 | Analyste | `build/01_analyse.md` | ✅ |
| 2 | Cartographe | `build/02_flux_et_diagnostic.md` | ✅ |
| 2b | (suite) | `build/02b_facturation_systeme_reel.md` | ✅ |
| 3 | Architecte | `build/03_architecture.md` (v2) | ✅ validé |
| 4 | Constructeur | `build/build_workbook.py`, `data/referentiels/*.csv`, `output/ORA_ORA_SSR_v0.xlsx`, `powerquery/*.pq` | ✅ |
| 5 | QA | `build/04_qa.md` | ✅ |
| 6 | Guide assistantes | `build/05_guide.md` | ⏳ à faire |

---

## 10. Comment reconstruire / travailler

```bash
python build/build_workbook.py            # régénère output/ORA_ORA_SSR_v0.xlsx depuis les CSV
python build/audit_xlsx.py <fichier.xlsx> # inspecte un fichier source (lecture seule)
```
- `fichiers_source/` = **lecture seule** (jamais écrire). Livrables dans `build/` et `output/`.
- **RGPD** : `.gitignore` exclut `*.xlsx`/`*.xlsm`/`fichiers_source/`/`build/_audit/`/`output/`.
  **Aucune donnée patient réelle ne doit être committée.** Les CSV de référentiels sont non personnels ;
  `data/referentiels/patients.csv` est un échantillon **synthétique**.

## 11. Git / push
Branche : `claude/happy-hypatia-qttbs3`. Le dépôt distant `HauN8N-Claude/SRRP-ORA-ORA` était **vide**.
Push initialement bloqué (proxy 403, puis token fine-grained sans permission **Contents: Read and write**).
→ Un token avec *Contents: write* est requis pour pousser. Penser à **révoquer** tout token exposé en chat.
