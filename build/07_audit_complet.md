# 07 — Audit complet du classeur (recherche approfondie de bugs)

**Objet.** Test adverse poussé du classeur `output/ORA_ORA_SSR_v0.xlsx` pour débusquer des défauts non vus :
audit structurel statique exhaustif + rejeu de cas-limites + revue d'agencement (facture) + analyse de
montée en charge. Complète `04_qa.md` et `06_audit_test_cas_fictif.md`.

**Limite (rappel honnête).** Recalcul Excel 365 impossible dans le bac à sable (LibreOffice ne charge pas le
fichier). Vérifications par **analyse statique openpyxl** + **réplication fidèle de la logique** des formules.
La matérialisation finale des tableaux dynamiques reste à confirmer dans Excel 365.

---

## 1. Audit structurel statique — RAS

Sur **1 574 formules / 24 onglets / 20 tables** :

| Contrôle | Résultat |
|---|---|
| Références `Table[Colonne]` pointant une table inconnue | **0** |
| Colonnes référencées inexistantes (fautes de frappe/accents) | **0** |
| Noms de fonctions en français (casseraient le fichier) | **0** |
| Erreurs littérales stockées (#REF!/#NOM?/#VALEUR!/#N/A) | **0** |
| Risque `#SPILL!` (FILTER non encapsulé dans une cellule de tableau) | **0** |
| Validations pointant une plage nommée valide | **21 / 21 OK** |
| Plages nommées résolvant correctement | **16 / 16 OK** |
| Formats de date sur colonnes dates (DA, Saisie, Suivi) | **OK** (`DD/MM/YYYY`) |
| Référence circulaire | **Aucune** (DA.Nb_Je_consommes dépend de colonnes de Saisie ≠ sa propre colonne) |

---

## 2. Cas-limites rejoués (logique) — 17/17 OK

| Cas testé | Attendu | Verdict |
|---|---|---|
| Date = borne début d'accord (06/01) | rattaché | ✅ |
| Date = borne fin d'accord (30/06) | rattaché | ✅ |
| Date = fin + 1 jour (01/07) | ⚠ SANS DA | ✅ |
| Statut DA exclu (DEP refusée) | ⚠ SANS DA | ✅ |
| Deux accords chevauchants même patient | ⚠ MULTI DA | ✅ |
| Nom en minuscules + espaces multiples | rattaché (normalisation) | ✅ |
| Programmation « Annulation » | non comptée présente | ✅ |
| Programmation « Présent » / « Attente CPS / Présent » | comptées présentes | ✅ |
| Régime SS → facturable CPS / payeur | Non / SS | ✅ |
| Régime RGS → payeur | CPS | ✅ |
| Cotation invalide (HJSX) | ⚠ tarif inconnu | ✅ |
| Nb journées (6) > Nb Je accordé (5) | non éligible | ✅ |
| Nb journées (5) = plafond (5) | éligible | ✅ |
| Catégorie de refus renseignée | non éligible | ✅ |
| Montant HJST × 4 | 108 000 | ✅ |

---

## 3. Défauts trouvés

### 🟠 F1 — Total de la `Facture` sous-évalué (CORRIGÉ)
- **Problème** : sur l'onglet `Facture`, le TOTAL valait `Tarif unitaire × Nb journées` en ne lisant que la
  **1re ligne** de la facture (XLOOKUP = première occurrence). Pour une facture **multi-lignes**
  (plusieurs semaines), le total était **sous-compté**.
- **Correction appliquée** (`build/build_workbook.py`, onglet `Facture`) :
  - `TOTAL` = `SUMIFS(tSuiviFactures[Montant] ; [N° facture])` → somme de **toutes** les lignes de la facture ;
  - `Nb de journées` = `SUMIFS(tSuiviFactures[Nb_journees] ; [N° facture])`.
  Cohérent avec le `Bordereaux` (qui sommait déjà correctement). Reconstruit et vérifié.

### 🔴 F2 — Capacité limitée à ~30 lignes (À TRAITER avant production) — *limitation, pas erreur de formule*
- **Constat** : les tables de travail (`DA`, `Saisie_Presences`, `CONSO_Presences`, `Suivi_Factures`,
  `Bordereaux`) sont livrées avec **~30 lignes-modèles**. De plus, `CONSO_Presences` recopie `Saisie_Presences`
  **par position** (`INDEX(tPresences ; LIGNE()-1)`) sur ces 30 lignes seulement, et les **validations** /
  **plages nommées** sont bornées (rows 2:31, ou 2:200/600/5000 selon la liste).
- **Impact** : tel quel, au-delà de ~30 présences, les lignes ajoutées **ne sont pas consolidées** par
  `CONSO_Presences` → `Nb_journees`, `Nb_Je_consommes` et les montants **sous-comptent**. Or le volume réel
  est d'environ **19 000 présences**.
- **Ce n'est pas visible en test à petite échelle** (d'où son importance).
- **Remédiation recommandée** (au choix, à décider) :
  1. **Power Query intra-classeur** pour `CONSO_Presences` et le calcul des présences (déjà prévu en évolution,
     `powerquery/`) — la voie robuste pour le volume ;
  2. à défaut, **pré-dimensionner** les tables et étendre validations + plages nommées au volume cible
     (alourdit le fichier, recalcul plus lent) et **remplacer le miroir `INDEX` par une copie pleine hauteur**.
- **Tant que ce point n'est pas traité, le classeur convient pour une démonstration / un pilote, pas pour le
  volume annuel complet.**

### 🟡 F3 — Règle « présent » à confirmer (métier)
`Compte_Present = VRAI` pour `Présent` **et** `Attente CPS / Présent`. À **confirmer** avec le métier que
« Attente CPS / Présent » doit bien être facturée comme présence. (Paramètre modifiable dans `REF_Programmation`.)

### 🟡 F4 — Montant en lettres borné (connu)
La LAMBDA `MontantEnLettres` couvre 0 – 999 999 999 XPF, sans décimales ; raffinements orthographiques mineurs.
Le **montant numérique** reste la valeur de référence.

---

## 4. Points sains confirmés
- Garde-fous présence : `⚠ HORS LISTE / SANS DA / MULTI DA` se déclenchent correctement (+ MFC rouge).
- Normalisation patient (espaces/casse) opérante.
- Tarifs réels, payeurs CPS/SS/Autres, plafond Nb Je, catégories de refus : logique correcte.
- Dates et nombres désormais en **vrais types** (correctif D4 antérieur).
- Aucune donnée patient réelle dans le fichier (7 patients synthétiques).

---

## 5. Verdict

| | |
|---|---|
| Audit statique | ✅ 0 anomalie sur 1 574 formules |
| Cas-limites logiques | ✅ 17/17 |
| Défaut corrigé | **F1** (total facture multi-lignes) |
| **Risque majeur restant** | **F2 — montée en charge (~30 lignes)** : traiter via Power Query ou pré-dimensionnement avant usage réel |
| À confirmer métier | F3 (programmations « présentes »), tarifs JOPF, mapping payeur |
| Validation finale | À faire dans **Excel 365** (recalcul réel) |

**Conclusion** : aucune erreur de formule détectée ; **1 bug de calcul corrigé** (total facture) ; **1 limite
structurelle importante** (volume) à arbitrer avant la mise en production. La conception reste saine.
