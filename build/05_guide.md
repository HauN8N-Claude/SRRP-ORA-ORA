# Guide d'utilisation — Classeur unique ORA ORA / SSR (facturation CPS)

*Guide simple pour les assistantes. Pas besoin d'être informaticienne.*

---

## ⭐ Les 3 règles d'or (à lire en premier)

1. **Ouvrir le fichier avec Excel 365** (PC ou navigateur). Avec une vieille version d'Excel, les
   calculs automatiques ne marchent pas.
2. **Ne jamais écrire dans les onglets gris `REF_…`** (ce sont les listes officielles : patients, tarifs,
   régimes…). Ils sont verrouillés exprès. Seul l'administrateur les modifie.
3. **On ne saisit que dans 3 onglets** : `DA` (les accords), `Saisie_Presences` (les présences),
   `Suivi_Factures` (les factures). Le reste se calcule tout seul.

> 💡 **Rouge = attention.** Chaque fois qu'une ligne devient rouge ou affiche un message qui commence par
> « ⚠ », c'est que quelque chose doit être corrigé. **Une case ne reste jamais vide en silence** : c'est
> justement ce qui évitait les erreurs invisibles d'avant.

---

## 🗂️ À quoi servent les onglets

| Onglet | C'est quoi | Qui écrit |
|---|---|---|
| `REF_…` (gris) | Listes officielles : patients, cotations, tarifs, régimes, statuts… | Admin seulement |
| **`DA`** | Les **demandes d'accord** (un accord CPS par patient/séjour) | Vous |
| **`Saisie_Presences`** | Les **présences** au quotidien | Vous |
| `CONSO_Presences` | Regroupe automatiquement les présences | *(automatique, ne pas toucher)* |
| **`Suivi_Factures`** | Le **suivi des factures** (montants, refus, paiements) | Vous |
| `Bordereaux` | Les **bordereaux** (totaux par facture) | *(automatique)* |
| `Facture` | Une **facture imprimable** | *(se remplit toute seule)* |
| `Cockpit` | Les **compteurs** (activité ETP / Polyvalent) | *(automatique)* |
| `CTRL_Qualite` | Le **tableau des anomalies** à vérifier | *(automatique)* |

---

## ✅ Tâche 1 — Enregistrer une demande d'accord (onglet `DA`)

Sur une nouvelle ligne, remplir au minimum :
- **N_DA** (le numéro de l'accord, ex. `25/001234`)
- **Patient** (choisir dans la liste déroulante)
- **Parcours** (ETP ou Polyvalent)
- **Statut** (ex. *DEP validée*)
- **Régime**, **Cotation accordée** (ex. HJSM), **Nb Je accordé**
- **Date début accordé** et **Date fin accordé** (la période couverte par l'accord)

Les colonnes de droite (âge, doublons, etc.) se calculent seules.
👉 Si le **numéro d'accord existe déjà** ailleurs, la colonne « Doublon » l'indique.

---

## ✅ Tâche 2 — Saisir une présence (onglet `Saisie_Presences`)

**C'est le geste de tous les jours. Vous ne remplissez que 4 colonnes :**

| Colonne | Quoi mettre |
|---|---|
| **Patient** | choisir dans la liste |
| **Date** | la date de présence |
| **Programmation** | *Présent*, *Absent*, *Annulation*… (liste) |
| **Parcours** | ETP ou Polyvalent |

➡️ Le reste (N° d'accord, cotation, régime, semaine…) **se remplit automatiquement**.

**Si une ligne devient rouge**, lisez le message :
- **⚠ HORS LISTE** → le nom n'existe pas dans la liste des patients (faute de frappe, ou patient à créer).
- **⚠ SANS DA** → aucune demande d'accord ne couvre ce patient à cette date (vérifier la période de l'accord
  ou le parcours).
- **⚠ MULTI DA** → plusieurs accords correspondent : à vérifier.

> Tant qu'une ligne est rouge, **elle ne partira pas en facturation**. C'est voulu : on facture seulement
> ce qui est propre.

---

## ✅ Tâche 3 — Suivre / générer une facture (onglet `Suivi_Factures`)

Sur une ligne, renseigner :
- **N° DE FACTURE** (ex. `F0012025`)
- **Code recherche** (le patient, depuis la liste)
- **DA** (le numéro d'accord concerné)
- **Cotation** (ex. HJSM)
- **N° semaine** (la semaine facturée)

➡️ Se calculent automatiquement : **nombre de journées**, **tarif**, **MONTANT**, le **payeur** (CPS / SS /
Autres) et l'**éligibilité** (la facture est-elle complète et facturable).

- Si un séjour a été **refusé**, choisir la **catégorie de refus** (ex. *Droits fermés*) : la ligne est suivie
  mais sortie des factures à envoyer.
- L'onglet **`Bordereaux`** additionne les montants par facture, et l'onglet **`Facture`** produit la
  **facture imprimable** : il suffit d'y indiquer le **N° de facture** en haut pour qu'elle se remplisse.

---

## 🔄 Mettre à jour les calculs

Normalement tout se recalcule tout seul. Si un total semble figé : onglet **Données → Actualiser tout**
(ou touche **F9**).

---

## 🔎 Avant d'envoyer à la CPS — le réflexe qualité

Ouvrir l'onglet **`CTRL_Qualite`** : il compte les anomalies (lignes hors liste, sans accord, doublons,
journées au-delà de l'accordé, tarifs inconnus…).
**Objectif : tous les compteurs à 0.** Si un compteur est positif, corriger les lignes rouges concernées
avant de transmettre.

---

## 🚫 À ne pas faire

- Modifier ou renommer les onglets gris `REF_…` ou les **titres de colonnes** (ça casse les calculs).
- Copier-coller des blocs entiers d'un autre fichier (saisir ligne par ligne).
- Travailler à plusieurs **en même temps sur le même fichier** posé sur le serveur : Excel le **verrouille
  pour une seule personne**. (Pour le travail simultané à plusieurs, voir avec l'administrateur :
  cela nécessite un hébergement particulier — point en cours d'arbitrage.)

---

## ❓ En cas de souci

1. Lire le message rouge (il dit quoi corriger).
2. Vérifier l'orthographe du nom (choisir toujours dans la liste).
3. Vérifier que l'**accord (DA)** existe et **couvre la date**.
4. Faire **Actualiser tout** (F9).
5. Si ça persiste, noter l'onglet + le n° de ligne et le signaler à l'administrateur.

---

*Rappel confidentialité : ce classeur contient des données de santé. Ne pas l'envoyer par mail non sécurisé
ni le copier hors du serveur de l'établissement.*
