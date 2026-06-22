# 06 — Audit de test sur cas patient fictif (validation du fonctionnement réel)

**Objet.** Saisir un patient fictif, ses présences et une facture, puis vérifier que la facturation se
calcule correctement dans `Suivi_Factures` (et en aval `Bordereaux` / `Facture`). Fichier testé :
copie `output/_test_cas_fictif.xlsx` (la version de référence `output/ORA_ORA_SSR_v0.xlsx` n'est pas altérée).

**Méthode & honnêteté sur le moteur de calcul.** Le recalcul Excel 365 réel est **impossible dans ce bac à
sable** : LibreOffice refuse de charger le `.xlsx` (« source file could not be loaded », confirmé). Les
résultats ci-dessous sont donc obtenus par **réplication fidèle en Python de la logique réelle des formules**
lues dans les cellules (FILTER/XLOOKUP/LET, normalisation, COUNTIFS, etc.), **pas par recalcul Excel**. La
preuve finale en conditions réelles reste à faire en ouvrant le fichier dans Excel 365 (protocole en `04_qa.md`).

---

## 1. Cas fictif injecté

- **Patient** (existant dans `REF_Patients`, synthétique) : `TESTPATIENT Alfred 12-04-1987`.
- **Accord (DA)** : `N_DA = TEST/000001`, Parcours ETP, Statut « DEP validée » (non exclu), Régime RGS,
  Cotation accordée **HJSM**, fenêtre **06/01/2025 → 30/06/2025**, Nb Je accordés = 20.
- **Présences** (onglet `Saisie_Presences`) : 6 lignes couvrant le cas nominal + 4 cas-limites.
- **Facture** (onglet `Suivi_Factures`) : `F-TEST-001`, patient ci-dessus, DA `TEST/000001`, cotation HJSM,
  semaine **S25-07** (semaine du 10/02/2025).

---

## 2. Résultats — résolution des présences (cas → attendu → obtenu)

| Ligne | Cas testé | Attendu | Obtenu | Verdict |
|---|---|---|---|---|
| 2 | Présence normale (10/02/2025, ETP) | rattachée TEST/000001, OK | `TEST/000001`, OK, HJSM, RGS, S25-07 | ✅ |
| 3 | Présence normale (11/02/2025) | rattachée, OK | `TEST/000001`, OK | ✅ |
| 4 | **Nom à espaces multiples** (`  TESTPATIENT   Alfred …`) | rattachée malgré les espaces (normalisation) | `TEST/000001`, OK | ✅ (règle d'or prouvée) |
| 5 | Date **hors fenêtre** d'accord (15/11/2024) | non rattachée, signalée | `⚠ SANS DA` | ✅ |
| 6 | Patient **hors référentiel** | signalée, jamais vide | `⚠ HORS LISTE` | ✅ |
| 7 | **Mauvais parcours** (Polyvalent) | non rattachée | `⚠ SANS DA` | ✅ |

→ Les 4 garde-fous attendus se déclenchent (et alimentent la mise en forme rouge + `CTRL_Qualite`).
La normalisation de la clé patient neutralise bien les espaces parasites (cause de bug n°1 du système actuel).

---

## 3. Résultats — cycle de facturation `Suivi_Factures` → `Bordereaux`

| Élément | Formule (logique) | Attendu | Obtenu | Verdict |
|---|---|---|---|---|
| Nb_journees (DA + semaine, présents) | COUNTIFS sur `CONSO_Presences` | 3 | **3** | ✅ |
| Tarif_unitaire (HJSM) | XLOOKUP `tTarif` | 30 000 | **30 000** | ✅ |
| **Montant** = tarif × journées | — | 90 000 | **90 000 XPF** | ✅ |
| Type de facture (payeur) | Régime RGS → `REF_Regimes[Payeur]` | CPS | **CPS** | ✅ |
| Plafond Nb Je | 3 ≤ 20 accordés | vrai | **vrai** | ✅ |
| **Eligible** (toutes conditions) | régime CPS + cotation valide + sans refus + ≤ plafond | VRAI | **VRAI** | ✅ |
| Champs patient facture | XLOOKUP `tPatients` | Nom/Prénom/Né(e)/Commune remplis | TESTPATIENT / Alfred / 12-04-1987 / PAPEETE | ✅ |
| **Bordereau** F-TEST-001 | SUMIFS montants | 90 000 | **90 000** | ✅ |

→ **La facture se génère correctement** : montant cohérent (30 000 × 3 = 90 000 XPF), payeur déduit du régime,
éligibilité validée, champs patient et bordereau cohérents.

---

## 4. Défauts trouvés et corrigés pendant ce test

| # | Défaut | Impact | Correction |
|---|---|---|---|
| **D4** 🔴 | Référentiels écrits en **texte** : dates calendrier/naissance (`"01/01/2023"`) et `Tarif_XPF` (`"32000"`) | En Excel, une date saisie (vraie date) ne « matche » pas une date-texte → **semaine = ⚠ hors calendrier** pour toute présence ; `DATEDIF` âge cassé ; tarif fragile | `build/build_workbook.py` : ajout d'une **conversion typée** (`coerce_value`/`set_typed`) → dates en vraies dates, nombres en nombres. Rebuild OK (Tarif=int 32000, calendrier Date=datetime, naissance=datetime). |

> Ce défaut aurait fait échouer le calcul de la semaine (donc du `Nb_journees`, donc du montant) en Excel réel.
> Il est désormais corrigé dans le générateur **et** dans `output/ORA_ORA_SSR_v0.xlsx` régénéré.

---

## 5. Limites & reste à valider

- **Recalcul Excel 365 non réalisé ici** (LibreOffice indisponible) : la logique est prouvée, la matérialisation
  des tableaux dynamiques (FILTER/XLOOKUP/LET) doit être confirmée en ouvrant le fichier dans **Excel 365**.
  Fichier prêt pour ça : `output/_test_cas_fictif.xlsx` (le cas fictif y est déjà saisi → à l'ouverture, la
  facture doit afficher **90 000 XPF**).
- `[à confirmer]` métier : tarifs vs arrêté JOPF, mapping régime→payeur (RGS/RNS/RST→CPS), programmations
  comptées « présentes ». Non bloquant pour la mécanique.

## 6. Verdict

**9 vérifications / 9 réussies** (6 résolutions de présence + cycle facturation complet). La chaîne
saisie → accord → consolidation → suivi → bordereau → facture est **logiquement correcte** ; un défaut
bloquant (dates/tarif en texte) a été **trouvé et corrigé** grâce à ce test. Confiance élevée sur la
conception ; **validation finale à faire en Excel 365** (ouvrir `_test_cas_fictif.xlsx`).
