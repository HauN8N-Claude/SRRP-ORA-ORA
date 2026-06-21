// ============================================================================
// CONSO_Presences — requête Power Query INTRA-CLASSEUR (évolution documentée).
// JAMAIS de chemin Z:\ ni Excel.Workbook(File.Contents(...)) : on lit la table
// du MÊME classeur (Excel.CurrentWorkbook), ce qui supprime la cause R11.
// À coller dans l'Éditeur avancé de Power Query (Données > Obtenir des données >
// Autres sources > Requête vide). Voir docs/04_power_query.md.
// Cette requête est une ALTERNATIVE à la dérivation par formule INDEX de
// l'onglet CONSO_Presences (v1 = formules natives ; PQ = évolution si latence).
// ============================================================================
let
    Source = Excel.CurrentWorkbook(){[Name="tPresences"]}[Content],

    // Normalisation de la clé patient AVANT toute jointure (règle d'or §5.1)
    AjoutCleNorm = Table.AddColumn(Source, "Cle_Norm", each
        Text.Upper(Text.Trim(
            Text.Replace(Text.Replace(Text.Replace(
                [Patient], Character.FromNumber(160), " "),
                "  ", " "), "  ", " "))), type text),

    // On conserve les colonnes utiles à la consolidation (ETP + Polyvalent
    // unifiés par la colonne Parcours déjà présente dans la saisie).
    Selection = Table.SelectColumns(AjoutCleNorm, {
        "Patient", "Cle_Norm", "Date", "Programmation", "Est_Present",
        "Parcours", "N_DA", "Cotation", "Regime", "Nb_Je_accordes",
        "Semaine", "Mois", "Groupe", "PEC", "Statut_Resolution"
    }),

    // On écarte les lignes vides (Patient non saisi).
    Filtre = Table.SelectRows(Selection, each [Patient] <> null and [Patient] <> "")
in
    Filtre
