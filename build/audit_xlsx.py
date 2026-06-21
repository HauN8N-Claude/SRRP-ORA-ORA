#!/usr/bin/env python3
"""audit_xlsx.py — Inspecteur exhaustif d'un classeur .xlsx (lecture seule).

Usage:
    python build/audit_xlsx.py <fichier.xlsx> [--json sortie.json] [--max-formules N]

Produit un dictionnaire structuré (affiché en JSON) décrivant, SANS rien modifier :
  - feuilles : titre, état (visible/caché), dimensions, lignes/colonnes utiles
  - en-têtes : 1re ligne non vide de chaque feuille + types inférés sur un échantillon
  - formules : nombre de cellules à formule par feuille + échantillon de formules distinctes
  - plages nommées (defined names)
  - validations de données (listes déroulantes, etc.)
  - tableaux structurés (ListObjects)
  - cellules fusionnées, mises en forme conditionnelles (compte)
  - erreurs de formule littérales détectées (#REF!, #N/A, ...)

N'écrit jamais dans le fichier source (read_only + keep_vba=False).
"""
import sys
import json
import re
from collections import Counter, defaultdict

import openpyxl
from openpyxl.utils import get_column_letter

ERREURS = ("#REF!", "#N/A", "#VALUE!", "#NAME?", "#DIV/0!", "#NULL!", "#NUM!")


def infer_type(v):
    if v is None:
        return "vide"
    if isinstance(v, bool):
        return "bool"
    if isinstance(v, int):
        return "int"
    if isinstance(v, float):
        return "float"
    # datetime
    import datetime
    if isinstance(v, (datetime.datetime, datetime.date)):
        return "date"
    if isinstance(v, str):
        s = v.strip()
        if s.startswith("="):
            return "formule"
        return "texte"
    return type(v).__name__


def normalize_formula(f):
    """Remplace les références cellule par # pour regrouper les formules par motif."""
    if not isinstance(f, str):
        return str(f)
    g = re.sub(r"\$?[A-Z]{1,3}\$?\d+", "@", f)
    return g


def audit(path, max_formules=12, header_scan=200, type_sample=400):
    # Passe 1 : formules (data_only=False) ; on lit aussi les valeurs textuelles.
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False, keep_links=False)
    result = {
        "fichier": path,
        "proprietes": {},
        "plages_nommees": [],
        "feuilles": [],
        "resume": {},
    }
    # Propriétés
    try:
        p = wb.properties
        result["proprietes"] = {
            "creator": p.creator,
            "lastModifiedBy": p.lastModifiedBy,
            "created": str(p.created) if p.created else None,
            "modified": str(p.modified) if p.modified else None,
            "title": p.title,
        }
    except Exception as e:
        result["proprietes"] = {"erreur": str(e)}

    # Plages nommées (defined names)
    try:
        dn = wb.defined_names
        for name in dn:
            d = dn[name]
            result["plages_nommees"].append({
                "nom": name,
                "ref": str(d.value),
                "cachee": bool(getattr(d, "hidden", False)),
            })
    except Exception as e:
        result["plages_nommees"] = [{"erreur": str(e)}]

    total_formules = 0
    total_erreurs = 0

    for ws in wb.worksheets:
        sheet = {
            "titre": ws.title,
            "etat": ws.sheet_state,  # visible / hidden / veryHidden
            "dimension": None,
            "max_row": None,
            "max_col": None,
            "entetes": [],
            "nb_formules": 0,
            "formules_motifs": [],
            "erreurs_litterales": {},
            "colonnes_types": [],
        }

        # --- Passe unique sur toute la feuille (values_only=True).
        # En read_only + data_only=False, une cellule de formule renvoie sa chaîne "=...".
        header_row_idx = None
        header_values = None
        col_type_counter = defaultdict(Counter)
        col_examples = {}
        formula_patterns = Counter()
        err_counter = Counter()
        max_col = 0
        nb_lignes_non_vides = 0
        derniere_ligne = 0
        data_rows_scanned = 0

        for ridx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            non_vides_idx = [i for i, c in enumerate(row, start=1) if c not in (None, "")]
            if non_vides_idx:
                nb_lignes_non_vides += 1
                derniere_ligne = ridx
                if non_vides_idx[-1] > max_col:
                    max_col = non_vides_idx[-1]
            # détection de l'en-tête
            if header_row_idx is None and ridx <= header_scan:
                non_vides = [row[i - 1] for i in non_vides_idx]
                txt = [c for c in non_vides if isinstance(c, str) and not c.startswith("=")]
                if len(non_vides) >= 2 and len(txt) >= max(1, len(non_vides) // 2):
                    header_row_idx = ridx
                    header_values = row
                    continue
            # lignes de données -> types, exemples, formules, erreurs
            if header_row_idx is not None and ridx > header_row_idx:
                limited = data_rows_scanned < type_sample
                for ci in non_vides_idx:
                    v = row[ci - 1]
                    t = infer_type(v)
                    if t == "formule":
                        sheet["nb_formules"] += 1
                        total_formules += 1
                        formula_patterns[normalize_formula(v)] += 1
                    if limited:
                        col_type_counter[ci][t] += 1
                        if ci not in col_examples:
                            col_examples[ci] = str(v)[:60]
                    if isinstance(v, str):
                        for e in ERREURS:
                            if e in v:
                                err_counter[e] += 1
                                total_erreurs += 1
                data_rows_scanned += 1

        sheet["header_row"] = header_row_idx
        sheet["max_row"] = derniere_ligne
        sheet["max_col"] = max_col
        sheet["nb_lignes_non_vides"] = nb_lignes_non_vides
        sheet["dimension"] = (f"{get_column_letter(max_col)}{derniere_ligne}" if max_col else None)
        if header_values:
            for ci, val in enumerate(header_values, start=1):
                if val not in (None, ""):
                    sheet["entetes"].append({
                        "col": get_column_letter(ci),
                        "nom": str(val).strip(),
                    })

        for ci in sorted(col_type_counter):
            types = col_type_counter[ci]
            sheet["colonnes_types"].append({
                "col": get_column_letter(ci),
                "types": dict(types),
                "exemple": col_examples.get(ci),
            })
        sheet["formules_motifs"] = [
            {"motif": m, "occurrences": n} for m, n in formula_patterns.most_common(max_formules)
        ]
        sheet["erreurs_litterales"] = dict(err_counter)

        # --- Validations de données (non dispo en read_only -> 2e passe ciblée plus bas)
        result["feuilles"].append(sheet)

    wb.close()

    # Passe 2 : validations, tables, fusions, MFC (nécessite read_only=False mais sans data)
    try:
        wb2 = openpyxl.load_workbook(path, read_only=False, data_only=False, keep_links=False)
        by_title = {s["titre"]: s for s in result["feuilles"]}
        for ws in wb2.worksheets:
            s = by_title.get(ws.title)
            if s is None:
                continue
            # Validations
            vals = []
            try:
                for dv in ws.data_validations.dataValidation:
                    vals.append({
                        "type": dv.type,
                        "operator": dv.operator,
                        "formula1": dv.formula1,
                        "ranges": str(dv.sqref),
                        "allow_blank": dv.allow_blank,
                    })
            except Exception as e:
                vals = [{"erreur": str(e)}]
            s["validations"] = vals
            # Tables structurées
            tables = []
            try:
                for tname, tref in (ws.tables.items() if hasattr(ws, "tables") else []):
                    tables.append({"nom": tname, "ref": tref})
            except Exception as e:
                tables = [{"erreur": str(e)}]
            s["tables"] = tables
            # Fusions + MFC
            try:
                s["nb_cellules_fusionnees"] = len(ws.merged_cells.ranges)
            except Exception:
                s["nb_cellules_fusionnees"] = None
            try:
                s["nb_mfc"] = len(ws.conditional_formatting._cf_rules) if ws.conditional_formatting else 0
            except Exception:
                s["nb_mfc"] = None
        wb2.close()
    except Exception as e:
        result["resume"]["erreur_passe2"] = str(e)

    result["resume"]["total_formules"] = total_formules
    result["resume"]["total_erreurs_litterales"] = total_erreurs
    result["resume"]["nb_feuilles"] = len(result["feuilles"])
    return result


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    path = sys.argv[1]
    out_json = None
    max_f = 12
    if "--json" in sys.argv:
        out_json = sys.argv[sys.argv.index("--json") + 1]
    if "--max-formules" in sys.argv:
        max_f = int(sys.argv[sys.argv.index("--max-formules") + 1])
    res = audit(path, max_formules=max_f)
    txt = json.dumps(res, ensure_ascii=False, indent=2, default=str)
    if out_json:
        with open(out_json, "w", encoding="utf-8") as f:
            f.write(txt)
        print(f"Audit écrit -> {out_json}")
        print(f"  feuilles={res['resume']['nb_feuilles']} "
              f"formules={res['resume']['total_formules']} "
              f"erreurs={res['resume']['total_erreurs_litterales']}")
    else:
        print(txt)


if __name__ == "__main__":
    main()
