# Rapport de validation — patient fictif & génération de facture

**Objet** : valider le fonctionnement réel de la chaîne `REF → DA → Présences → Suivi_Factures → Bordereaux → Facture` à l'aide d'un patient fictif.
**Date** : 2026-06-22
**Fichier de test produit** : `classeur/ORA_ORA_SSR_v0_TEST.xlsx`
**Méthode de calcul** : LibreOffice n'étant pas opérationnel dans l'environnement d'analyse (échec de chargement de tout `.xlsx`), la validation a été réalisée par **simulation fidèle du moteur en Python** (reproduction exacte des formules). Le fichier de test reste à **ouvrir dans Excel 365** pour confirmation visuelle (il recalculera et affichera les mêmes valeurs).

---

## 1. Pré-requis appliqué : correction du défaut A1

La chaîne ne peut pas calculer tant que les dates des référentiels sont stockées en **texte** (cf. `AUDIT_ORA_ORA_SSR.md`, constat A1). Dans la **copie de test**, j'ai donc :

- converti **2192** dates `REF_Calendrier[Date]` texte → **vraies dates** ;
- converti `REF_Patients[Date_naissance]` / `Date_saisie` texte → vraies dates ;
- converti `REF_Tarifs[Tarif_XPF]` texte → **nombres** (constat A3).

> ⚠️ Le classeur original (`ORA_ORA_SSR_v0.xlsx`) **n'a pas été modifié**. Sans cette correction, `Saisie_Presences[Semaine]` renvoie « ⚠ hors calendrier », `Nb_journees` = 0 et le **montant est nul** : la facture ne se génère pas. C'est la démonstration concrète de l'impact du défaut A1.

---

## 2. Jeu de données fictif injecté

| Élément | Valeurs |
|---|---|
| **Patient** (`REF_Patients` L9) | `TESTCLAUDE Jean`, DN `DN0008`, né le 15/03/1980, PAPEETE — clé de recherche `TESTCLAUDE Jean 15-03-1980` |
| **DA** (`DA` L2) | `DA001`, parcours **ETP**, régime **RGS** (CPS facturable), statut **DEP validée** (non exclu), cotation accordée **HJSN**, **10** journées accordées, fenêtre 01/04/2024 → 30/04/2024 |
| **Présences** (`Saisie_Presences` L2-4) | 02, 03 et 04/04/2024, programmation **Présent**, parcours ETP |
| **Facture** (`Suivi_Factures` L2) | n° **F0012026**, semaine **S24-14**, cotation **HJSN**, DA `DA001`, sans refus |
| **Bordereau** (`Bordereaux` L2) | n° facture **F0012026** |

---

## 3. Résultats calculés (validation)

### 3.1 Résolution automatique des présences (`Saisie_Presences`)
| Ligne | Date | Programmation | N_DA résolu | Statut | Présent | Semaine |
|---|---|---|---|---|---|---|
| 2 | 02/04/2024 | Présent | **DA001** | OK | Vrai | S24-14 |
| 3 | 03/04/2024 | Présent | **DA001** | OK | Vrai | S24-14 |
| 4 | 04/04/2024 | Présent | **DA001** | OK | Vrai | S24-14 |

→ Le `FILTER` (patient + parcours + date dans la fenêtre + DA non exclue) résout correctement la DA. Aucun drapeau `HORS LISTE` / `SANS DA` / `MULTI DA`.

### 3.2 Facture générée (`Suivi_Factures` L2)
| Champ | Colonne | Valeur calculée | ✔ |
|---|---|---|---|
| N° facture | E | F0012026 | ✔ |
| Patient | L/M/N | DN0008 / TESTCLAUDE / Jean | ✔ |
| Régime | K | RGS | ✔ |
| Type de facture (payeur) | G | **CPS** | ✔ |
| Cotation | H | HJSN | ✔ |
| Tarif unitaire | V | **32 000 XPF** | ✔ |
| Semaine | C | S24-14 | ✔ |
| **Nb de journées** | U | **3** | ✔ |
| **MONTANT** | W | **96 000 XPF** | ✔ |
| Plafond accordé | — | 3 ≤ 10 | ✔ |
| **Éligible** | Z | **VRAI** | ✔ |
| Motif de rejet | AA | *(vide)* | ✔ |

`Montant = Tarif (32 000) × Nb journées (3) = 96 000 XPF`.

### 3.3 Bordereau & facture imprimable
| Élément | Valeur | ✔ |
|---|---|---|
| `Bordereaux` — N° 1, facture F0012026, DN0008, **Montant** | **96 000 XPF** | ✔ |
| `Facture` B21 — **TOTAL** | **96 000 XPF** | ✔ |
| `Facture` B22 — **Montant en lettres** | **quatre-vingt-seize mille FRANCS CFP** | ✔ |
| Contrôle de cohérence `CTRL_Qualite` Σ Bordereaux − Σ Suivi | 0 | ✔ |

---

## 4. Conclusion

✅ **La facture se génère correctement.** Toute la chaîne fonctionne avec le patient fictif : résolution de la DA, comptage des journées, tarif, montant, éligibilité, report en bordereau et en facture imprimable (montant en lettres inclus).

**Condition impérative** : la correction des **dates des référentiels** (défaut A1) doit être appliquée — c'est elle qui débloque le calcul. Elle est intégrée dans `ORA_ORA_SSR_v0_TEST.xlsx`.

### À faire côté utilisateur
1. Ouvrir **`ORA_ORA_SSR_v0_TEST.xlsx` dans Excel 365** ; laisser recalculer ; confirmer visuellement les valeurs ci-dessus (notamment `Suivi_Factures` L2 et l'onglet `Facture`).
2. Vérifier l'onglet `CTRL_Qualite` : tous les seuils `0` à **0**.
3. Reporter la correction des dates dans le classeur de production.

> Note : `ORA_ORA_SSR_v0_TEST.xlsx` contient les formules mais **pas de valeurs en cache** (généré hors Excel) — Excel 365 les calculera à l'ouverture. Les résultats attendus sont ceux du §3, confirmés par simulation.
