# CLAUDE.md — Projet ORA ORA / SSR · Classeur Excel unique de facturation CPS

> Ce fichier est lu automatiquement par Claude Code à l'ouverture du dépôt.
> Il fixe la mission, les contraintes non négociables, l'architecture cible et les conventions de travail.
> **À lire en entier avant toute action.**

---

## 1. Mission

Construire **UN seul classeur Excel** (`ORA_ORA_SSR.xlsx`) qui remplace les 4 fichiers actuels
de l'hôpital de jour SSR (ORA ORA, Pirae, Tahiti) et qui :

1. centralise présences, patients, demandes d'accord (DA) et référentiels ;
2. **fiabilise la chaîne qui mène à la facturation CPS** (le vrai objectif métier :
   raccourcir le délai entre service rendu et facture transmise, pour protéger la trésorerie) ;
3. **conserve Excel comme unique interface des assistantes** — aucune sortie de l'outil.

Le livrable final est un fichier Excel, pas une application web. La logique fragile actuelle
(jointures texte volatiles, copier-coller inter-fichiers, recalcul manuel) est remplacée par
une plomberie robuste (Power Query + tables + référentiels uniques + contrôles de saisie),
**sans changer l'ergonomie de saisie**.

---

## 2. Contraintes NON négociables

- **Excel reste l'interface.** Pas de migration vers un autre outil pour les utilisatrices.
- **La saisie ne change pas d'ergonomie** : l'assistante saisit Patient + Date + Programmation,
  le reste se calcule. Mêmes listes déroulantes qu'aujourd'hui.
- **Zéro erreur de formule** dans le livrable (`#REF!`, `#N/A`, `#VALUE!`, `#NAME?`, `#DIV/0!`).
- **Une seule source de vérité** par donnée (fin des référentiels dupliqués entre fichiers,
  fin du double fichier ETP divergent).
- **Données patients réelles = sensibles (RGPD applicable en PF depuis juin 2019).**
  Ne jamais committer de données patients réelles dans le dépôt. Les fichiers source fournis
  sont déjà anonymisés ; garder cette discipline.
- Ne rien affirmer sur le protocole CPS qui ne soit pas vérifié (voir §7, dépendance critique).

---

## 3. Architecture cible du classeur (couches)

Un seul `.xlsx`, onglets organisés par rôle (préfixes pour l'ordre logique) :

| Couche | Onglets | Rôle | Qui écrit |
|---|---|---|---|
| **Référentiels** | `REF_Patients`, `REF_Cotations`, `REF_Regimes`, `REF_Pathologies`, `REF_Provenances`, `REF_Prescripteurs`, `REF_Statuts`, `REF_Mouvements`, `REF_Programmation`, `REF_Groupes`, `REF_Communes`, `REF_Motifs`, `REF_Calendrier` | Listes maîtres, verrouillées | Admin only |
| **Registre** | `DA` | Demandes d'accord / séjours (clé pivot = N° DA) | Assistantes |
| **Saisie** | `Saisie_Presences` | Présences. Colonnes saisies : Patient, Date, Programmation | Assistantes |
| **Consolidation** | `CONSO_Presences` | Table unique alimentée par Power Query (ETP + Polyvalent) | Power Query (auto) |
| **Facturation** | `Facturation` | Dossiers prêts à facturer + contrôles | Power Query + formules |
| **Pilotage** | `Cockpit` | TCD / compteurs (réel + prévisionnel) | TCD |

Détails complets : `docs/02_architecture_cible.md` et `docs/03_dictionnaire_donnees.md`.

---

## 4. Le cœur technique : la jointure présence ↔ accord

La logique métier actuelle (à conserver, mais fiabiliser) :

> Une ligne de présence est rattachée à une DA si :
> **Patient identique** ET **Date ∈ [Date début accordé ; Date fin accordé]**
> ET **Parcours identique** ET **Statut DA ∉ {DEP refusée, DEP annulée}**.
> La DA fournit alors : N° DA, Cotation, Régime, Nb Je accordés, PEC, Groupe.

Aujourd'hui réalisée par `FILTER`/`XLOOKUP` (Excel 365, volatile, casse au moindre écart de texte).
**Cible : Power Query** (merge sur clé Patient **normalisée** + filtre intervalle de dates),
rafraîchi en un bouton « Actualiser tout ». Voir `docs/04_power_query.md` et `powerquery/*.pq`.

**Règle d'or anti-rupture** : toute clé Patient est normalisée (`Trim` + espaces multiples
réduits à un + casse homogène) AVANT toute jointure. C'est ce qui neutralise les 66 noms à
espaces doubles constatés et tout écart de frappe.

---

## 5. Contrainte technique importante : Power Query ≠ openpyxl

**openpyxl ne sait PAS créer/écrire les requêtes Power Query** (elles vivent dans les
connexions/customXml du classeur). Conséquence sur la méthode de build :

- Le script `build/build_workbook.py` génère **tout ce qui est scriptable** :
  onglets référentiels (depuis `data/referentiels/*.csv`), tables nommées (ListObjects),
  feuille `Saisie_Presences` avec listes déroulantes, `DA`, squelette `CONSO_Presences`
  et `Facturation`, mises en forme conditionnelles d'alerte.
- Les **requêtes Power Query** sont fournies en code M dans `powerquery/*.pq` et doivent
  être collées dans l'**Éditeur avancé** de Power Query (étape manuelle unique dans Excel,
  documentée pas à pas dans `docs/04_power_query.md`).
- Ne pas tenter de simuler Power Query en Python : on perdrait le rafraîchissement natif
  qui est justement l'intérêt.

Alternative possible si l'on veut un build 100 % scriptable (à arbitrer avec le métier) :
jointures en **formules dynamiques natives** (XLOOKUP/FILTER) robustifiées par normalisation
de clé dans une colonne d'aide. Moins idéal que Power Query mais entièrement scriptable.
Voir `docs/04_power_query.md` §Alternative.

---

## 6. Comment travailler dans ce dépôt

```bash
# 1. (Ré)générer le classeur v0 depuis les référentiels
python build/build_workbook.py
#    -> produit output/ORA_ORA_SSR_v0.xlsx

# 2. Recalculer + vérifier zéro erreur (si LibreOffice dispo)
python build/build_workbook.py --recalc

# 3. Inspecter un fichier source pour vérifier une hypothèse
python build/audit_xlsx.py /chemin/fichier.xlsx
```

- **Toujours** régénérer le classeur via le script, jamais à la main : le script est la
  source de vérité de la structure.
- Les `data/referentiels/*.csv` sont les seeds. Mettre à jour un référentiel = éditer le CSV
  puis relancer le build.
- Convention de code : Python concis, pas de valeurs calculées en dur (utiliser des formules
  Excel), polices Arial, voir `docs/03_dictionnaire_donnees.md` pour les types/formats.

---

## 7. Dépendance critique encore ouverte — CPS

Le **format et le canal de transmission des factures à la CPS** (service GDR / DSI) ne sont
**pas** déductibles des fichiers Excel. C'est la dépendance la plus structurante du projet.
Tant qu'elle n'est pas tranchée, l'onglet `Facturation` produit un **bordereau interne validé**
(dossiers prêts), pas le format CPS final. Action à mener hors-code : entretien GDR/DSI CPS.
Voir `docs/05_regles_validation_facturation.md` §Facturation.

Spécificités locales confirmées : pas de Carte Vitale, pas de PMSI ; tarifs CPS fixés par
arrêté publié au JOPF ; 5 cotations en vigueur (HJSR, HJST, HJSN, HJSA, HJSM).

---

## 8. État des lieux (résumé du diagnostic)

Points de rupture vérifiés sur les 4 fichiers source (détail : `docs/01_diagnostic.md`) :

1. 🔴 `TabPresencePoly` quasi vide (1 ligne / 10 510) dans le fichier de consolidation
   → comptage Polyvalent faux.
2. 🔴 Double fichier ETP divergent (`EXCEL_ETP` vs `GP_ETP__1_`, registres DA différents)
   → pas de source unique.
3. 🔴 Jointure facturation = appariement texte exact du nom patient (66 noms à espaces doubles).
4. 🔴 Dépendance au recalcul / copier-coller manuel.
5. 🟠 `Chiffres prévisionnels` : sources mélangées (Poly vs ETP selon la ligne).
6. 🟠 Doublons Patient+Date (801 dans TabPresences, dont 96 à patient vide).
7. 🟠 `#REF!` dans DA (col « Age fixe séjour »), référentiel `Path` = #N/A.
8. 🟠 Nommage incohérent (`TabPatients ` avec espace, validations à cibles multiples).

La cible (§3-4) corrige chacun de ces points par construction.

---

## 9. Definition of Done (v1)

- [ ] Un seul `.xlsx` ouvre sans erreur de formule.
- [ ] Saisie d'une présence (Patient+Date+Programmation) → DA, cotation, régime, Nb Je
      remplis automatiquement après « Actualiser tout ».
- [ ] Saisie d'un nom hors-liste ou mal orthographié → ligne signalée en rouge, jamais
      silencieusement vide.
- [ ] `Facturation` ne liste que : Présent + CPS + cotation valide + N° DA résolu
      + date dans la fenêtre d'accord + Nb Je consommés ≤ Nb Je accordés.
- [ ] `Cockpit` : compteurs réels ETP **et** Polyvalent corrects (plus de table Poly vide).
- [ ] Référentiels uniques, verrouillés ; aucune donnée patient réelle dans le dépôt.
- [ ] Procédure CPS documentée OU bordereau interne validé en attendant le format CPS.
