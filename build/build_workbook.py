#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_workbook.py — Générateur idempotent du classeur unique ORA ORA / SSR.

Produit  : output/ORA_ORA_SSR_v0.xlsx  (23 onglets — spec build/03_architecture.md §2)
Lit      : data/referentiels/*.csv     (seeds non personnels + patients SYNTHÉTIQUES)
Méthode  : openpyxl. Aucune valeur métier calculée en dur : tout passe par des
           formules Excel 365 (FILTER/XLOOKUP/LET) — voir CONTRAINTE openpyxl ci-dessous.

CONTRAINTE openpyxl ≠ Excel : openpyxl écrit la chaîne BRUTE de formule dans le XML.
Excel attend l'anglais, séparateur VIRGULE, et préfixe les fonctions modernes :
  XLOOKUP -> _xlfn.XLOOKUP ; FILTER -> _xlfn._xlws.FILTER ; LET -> _xlfn.LET ;
  TEXTJOIN -> _xlfn.TEXTJOIN ; IFS -> _xlfn.IFS ; LAMBDA -> _xlfn.LAMBDA.
Les helpers ci-dessous encapsulent ces préfixes ; on n'écrit JAMAIS de noms français.

Usage :
  python build/build_workbook.py            # génère le classeur
  python build/build_workbook.py --recalc   # + tente un recalcul LibreOffice headless
"""

import csv
import os
import sys
import subprocess
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

# --------------------------------------------------------------------------- #
# Chemins
# --------------------------------------------------------------------------- #
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REF_DIR = os.path.join(ROOT, "data", "referentiels")
OUT_DIR = os.path.join(ROOT, "output")
OUT_FILE = os.path.join(OUT_DIR, "ORA_ORA_SSR_v0.xlsx")

# --------------------------------------------------------------------------- #
# Style (Arial partout — convention CLAUDE.md)
# --------------------------------------------------------------------------- #
ARIAL = "Arial"
F_BASE = Font(name=ARIAL, size=10)
F_HDR = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=ARIAL, size=12, bold=True)
FILL_HDR = PatternFill("solid", fgColor="1F4E78")        # bleu en-tête
FILL_REF = PatternFill("solid", fgColor="2E75B6")        # bleu référentiel
FILL_CALC = PatternFill("solid", fgColor="DDEBF7")       # colonnes calculées
FILL_RED = PatternFill("solid", fgColor="FFC7CE")        # alerte MFC
FILL_NOTE = PatternFill("solid", fgColor="FFF2CC")       # note
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NB_TEMPLATE_ROWS = 30   # lignes-modèles (formules de colonnes calculées) sans données


# --------------------------------------------------------------------------- #
# Helpers formules : on écrit en anglais + préfixes _xlfn (cf. contrainte).
# --------------------------------------------------------------------------- #
def XLOOKUP(lookup, arr, ret, default='""'):
    return f'_xlfn.XLOOKUP({lookup},{arr},{ret},{default})'

def FILTER(arr, cond, default='""'):
    return f'_xlfn._xlws.FILTER({arr},{cond},{default})'

# --------------------------------------------------------------------------- #
# Lecture CSV
# --------------------------------------------------------------------------- #
def read_csv(name):
    path = os.path.join(REF_DIR, name)
    with open(path, newline="", encoding="utf-8") as f:
        rows = list(csv.reader(f, delimiter=";"))
    return rows[0], rows[1:]

def to_bool(v):
    return True if str(v).strip().upper() in ("VRAI", "TRUE", "1") else False


# --------------------------------------------------------------------------- #
# Pose d'une table (ListObject) + style + Arial + en-têtes
# --------------------------------------------------------------------------- #
def style_header(ws, row, ncol, fill=FILL_HDR):
    for c in range(1, ncol + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = F_HDR
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER

def add_table(ws, name, first_row, ncol, last_row, style="TableStyleMedium2"):
    ref = f"A{first_row}:{get_column_letter(ncol)}{last_row}"
    tbl = Table(displayName=name, ref=ref)
    tbl.tableStyleInfo = TableStyleInfo(
        name=style, showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False)
    ws.add_table(tbl)

def write_headers(ws, headers, row=1):
    for j, h in enumerate(headers, start=1):
        ws.cell(row=row, column=j, value=h)
    style_header(ws, row, len(headers))

def autosize(ws, headers, width=16):
    for j, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(j)].width = max(width, len(str(h)) + 2)


# --------------------------------------------------------------------------- #
# Construction d'un onglet référentiel à partir d'un CSV
# --------------------------------------------------------------------------- #
def build_ref_sheet(wb, sheet, table_name, headers, rows, bool_cols=(), protect=True):
    ws = wb.create_sheet(sheet)
    write_headers(ws, headers)
    for i, r in enumerate(rows, start=2):
        for j, val in enumerate(r, start=1):
            h = headers[j - 1]
            if h in bool_cols:
                ws.cell(row=i, column=j, value=to_bool(val))
            else:
                ws.cell(row=i, column=j, value=val if val != "" else None)
            ws.cell(row=i, column=j).font = F_BASE
    last = max(2, len(rows) + 1)
    add_table(ws, table_name, 1, len(headers), last)
    autosize(ws, headers)
    ws.freeze_panes = "A2"
    if protect:
        ws.protection.sheet = True  # verrouillage admin (référentiels figés)
    return ws


# --------------------------------------------------------------------------- #
# Plages nommées dédiées pour les listes déroulantes
# (openpyxl gère mal =tTable[col] en validation -> on crée des defined names
#  pointant une plage colonne large, puis on référence ces noms.)
# --------------------------------------------------------------------------- #
def add_named_range(wb, name, sheet, col_letter, last_row):
    ref = f"'{sheet}'!${col_letter}$2:${col_letter}${last_row}"
    wb.defined_names.add(DefinedName(name, attr_text=ref))

def list_validation(named_range, warn=False):
    dv = DataValidation(type="list", formula1=f"={named_range}", allow_blank=True)
    if warn:
        dv.errorStyle = "warning"     # avertir SANS bloquer (décision validée)
        dv.error = "Valeur hors liste — saisie autorisée, ligne signalée."
        dv.errorTitle = "Hors référentiel"
    else:
        dv.errorStyle = "stop"
    return dv


# --------------------------------------------------------------------------- #
# Clé patient normalisée (règle d'or §5.1) — formule réutilisée
#   MAJUSCULE(SUPPRESPACE(...)) avec CAR(160) neutralisé + espaces doubles
# --------------------------------------------------------------------------- #
def cle_norm_formula(ref):
    return (f'UPPER(TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE('
            f'{ref},CHAR(160)," "),"  "," "),"  "," ")))')


# =========================================================================== #
#  MAIN BUILD
# =========================================================================== #
def build():
    wb = Workbook()
    wb.remove(wb.active)

    # ---- 1. Référentiels -------------------------------------------------- #
    # patients (SYNTHÉTIQUES) — Cle_Norm calculée par formule
    h, rows = read_csv("patients.csv")
    h_pat = h + ["Cle_Norm"]
    ws = wb.create_sheet("REF_Patients")
    write_headers(ws, h_pat)
    idx_rech = h.index("Recherche") + 1
    for i, r in enumerate(rows, start=2):
        for j, val in enumerate(r, start=1):
            ws.cell(row=i, column=j, value=val if val != "" else None).font = F_BASE
        # Cle_Norm = normalisation de Recherche
        col_norm = len(h_pat)
        rech_ref = f"{get_column_letter(idx_rech)}{i}"
        ws.cell(row=i, column=col_norm,
                value="=" + cle_norm_formula(rech_ref)).font = F_BASE
    last_pat = len(rows) + 1
    add_table(ws, "tPatients", 1, len(h_pat), last_pat)
    autosize(ws, h_pat)
    ws.freeze_panes = "A2"
    ws.protection.sheet = True

    # Référentiels simples depuis CSV
    build_ref_sheet(wb, "REF_Cotations", "tCotations",
                    *read_csv("cotations.csv"))
    build_ref_sheet(wb, "REF_Tarifs", "tTarif",
                    *read_csv("tarifs.csv"))
    h, rows = read_csv("regimes.csv")
    build_ref_sheet(wb, "REF_Regimes", "tRegimes", h, rows,
                    bool_cols=("Facturable_CPS",))
    build_ref_sheet(wb, "REF_Pathologies", "tPathologies",
                    *read_csv("pathologies.csv"))
    build_ref_sheet(wb, "REF_Provenances", "tProvenances",
                    *read_csv("provenances.csv"))
    build_ref_sheet(wb, "REF_Prescripteurs", "tPrescripteurs",
                    *read_csv("prescripteurs.csv"))
    h, rows = read_csv("statuts.csv")
    build_ref_sheet(wb, "REF_Statuts", "tStatuts", h, rows,
                    bool_cols=("Exclu_facturation",))
    build_ref_sheet(wb, "REF_Mouvements", "tMouvements",
                    *read_csv("mouvements.csv"))
    h, rows = read_csv("programmation.csv")
    build_ref_sheet(wb, "REF_Programmation", "tProgrammation", h, rows,
                    bool_cols=("Compte_Present",))
    build_ref_sheet(wb, "REF_Groupes", "tGroupes",
                    *read_csv("groupes.csv"))
    build_ref_sheet(wb, "REF_Communes", "tCommunes",
                    *read_csv("communes.csv"))
    # Motifs : 3 colonnes regroupées sur une feuille
    ws_m = wb.create_sheet("REF_Motifs")
    mh, mrows = read_csv("motifs_hospit.csv")
    ch, crows = read_csv("motifs_refus_cps.csv")
    oh, orows = read_csv("motifs_refus_oraora.csv")
    motif_headers = ["Motif_Hospit", "Motif_refus_CPS", "Motif_refus_OraOra"]
    write_headers(ws_m, motif_headers)
    nmax = max(len(mrows), len(crows), len(orows))
    for i in range(nmax):
        if i < len(mrows): ws_m.cell(row=i+2, column=1, value=mrows[i][0]).font = F_BASE
        if i < len(crows): ws_m.cell(row=i+2, column=2, value=crows[i][0]).font = F_BASE
        if i < len(orows): ws_m.cell(row=i+2, column=3, value=orows[i][0]).font = F_BASE
    add_table(ws_m, "tMotifs", 1, 3, nmax + 1)
    autosize(ws_m, motif_headers, 30)
    ws_m.protection.sheet = True
    build_ref_sheet(wb, "REF_CategoriesRefus", "tCategoriesRefus",
                    *read_csv("categories_refus.csv"))
    h, rows = read_csv("calendrier.csv")
    build_ref_sheet(wb, "REF_Calendrier", "tCalendrier", h, rows)

    # ---- Plages nommées pour validations --------------------------------- #
    # (étendues large pour absorber l'import réel ultérieur)
    add_named_range(wb, "nrPatients", "REF_Patients",
                    get_column_letter(idx_rech), 5000)
    add_named_range(wb, "nrProgrammation", "REF_Programmation", "A", 200)
    add_named_range(wb, "nrCotations", "REF_Cotations", "A", 200)
    add_named_range(wb, "nrRegimes", "REF_Regimes", "A", 200)
    add_named_range(wb, "nrStatuts", "REF_Statuts", "A", 200)
    add_named_range(wb, "nrMouvements", "REF_Mouvements", "A", 200)
    add_named_range(wb, "nrProvenances", "REF_Provenances", "A", 200)
    add_named_range(wb, "nrPrescripteurs", "REF_Prescripteurs", "A", 600)
    add_named_range(wb, "nrPathologies", "REF_Pathologies", "A", 200)
    add_named_range(wb, "nrGroupes", "REF_Groupes", "A", 200)
    add_named_range(wb, "nrCategoriesRefus", "REF_CategoriesRefus", "A", 200)
    add_named_range(wb, "nrSemaine", "REF_Calendrier", "E", 5000)
    add_named_range(wb, "nrMotifHospit", "REF_Motifs", "A", 200)
    add_named_range(wb, "nrParcours", "_Listes", "A", 4)  # rempli plus bas

    # Feuille technique cachée pour la liste {ETP, Polyvalent}
    ws_l = wb.create_sheet("_Listes")
    ws_l["A1"] = "Parcours"
    ws_l["A2"] = "ETP"
    ws_l["A3"] = "Polyvalent"
    for r in (1, 2, 3):
        ws_l[f"A{r}"].font = F_BASE
    ws_l.sheet_state = "hidden"
    add_named_range(wb, "nrParcours", "_Listes", "A", 3)

    # ---- 2. DA ------------------------------------------------------------ #
    build_DA(wb)

    # ---- 3. Saisie_Presences --------------------------------------------- #
    build_saisie(wb)

    # ---- 4. CONSO_Presences ---------------------------------------------- #
    build_conso(wb)

    # ---- 5. Couche Facturation ------------------------------------------- #
    build_suivi(wb)
    build_bordereaux(wb)
    build_facture(wb)

    # ---- 6. Pilotage ----------------------------------------------------- #
    build_cockpit(wb)
    build_ctrl(wb)

    os.makedirs(OUT_DIR, exist_ok=True)
    wb.save(OUT_FILE)
    return wb


# --------------------------------------------------------------------------- #
# DA — registre des demandes d'accord (clé pivot N° DA en texte)
# --------------------------------------------------------------------------- #
def build_DA(wb):
    ws = wb.create_sheet("DA")
    headers = [
        "N_DA", "Patient", "Parcours", "Motif_Hospit", "Pathologie_medicale",
        "Regime", "Provenance", "Prescripteur", "Date_demande", "Statut",
        "Mouvements", "Nb_Je_demande", "Cotation_demandee", "Date_debut_demandee",
        "Date_fin_demandee", "Date_postage", "Numero_sejour", "PEC",
        "Nb_Je_Accorde", "Cotation_accordee", "Date_debut_accorde",
        "Date_fin_accorde", "Date_accord", "Groupe", "Programme", "Date_FIN_PEC",
        # calculées :
        "Cle_Patient_Norm", "Statut_Exclu", "Age_sejour", "Nb_Je_consommes",
        "Doublon_N_DA",
    ]
    write_headers(ws, headers)
    col = {h: i + 1 for i, h in enumerate(headers)}

    for r in range(2, 2 + NB_TEMPLATE_ROWS):
        pat = f"B{r}"
        cle = f"{get_column_letter(col['Cle_Patient_Norm'])}{r}"
        # Cle_Patient_Norm
        ws.cell(row=r, column=col["Cle_Patient_Norm"],
                value="=" + cle_norm_formula(pat))
        # Statut_Exclu
        ws.cell(row=r, column=col["Statut_Exclu"], value=(
            f'=IFERROR({XLOOKUP("[@Statut]", "tStatuts[Statut]", "tStatuts[Exclu_facturation]")},FALSE)'
        ))
        # Age_sejour (reconstruit, garde signalante, sans #REF!)
        ws.cell(row=r, column=col["Age_sejour"], value=(
            '=IF(OR([@Patient]="",[@Date_debut_accorde]=""),"",'
            'IFERROR(DATEDIF(' +
            XLOOKUP("[@Cle_Patient_Norm]", "tPatients[Cle_Norm]", "tPatients[Date_naissance]") +
            ',[@Date_debut_accorde],"y"),"⚠ naiss. introuvable"))'
        ))
        # Nb_Je_consommes
        ws.cell(row=r, column=col["Nb_Je_consommes"], value=(
            '=COUNTIFS(CONSO_Presences[N_DA],[@N_DA],'
            'CONSO_Presences[Est_Present],TRUE)'
        ))
        # Doublon_N_DA
        ws.cell(row=r, column=col["Doublon_N_DA"], value=(
            '=IF([@N_DA]="","",IF(COUNTIF(tDA[N_DA],[@N_DA])>1,"DOUBLON",""))'
        ))
    last = 1 + NB_TEMPLATE_ROWS
    add_table(ws, "tDA", 1, len(headers), last)
    autosize(ws, headers, 15)
    ws.freeze_panes = "B2"
    fmt_text_column(ws, col["N_DA"], last)        # N° DA en TEXTE
    fmt_date_columns(ws, [col[c] for c in headers if c.startswith("Date_")], last)

    # validations
    add_dv(ws, "nrPatients", f"B2:B{last}", warn=True)
    add_dv(ws, "nrParcours", f"C2:C{last}")
    add_dv(ws, "nrMotifHospit", f"D2:D{last}")
    add_dv(ws, "nrPathologies", f"E2:E{last}")
    add_dv(ws, "nrRegimes", f"F2:F{last}")
    add_dv(ws, "nrProvenances", f"G2:G{last}")
    add_dv(ws, "nrPrescripteurs", f"H2:H{last}")
    add_dv(ws, "nrStatuts", f"J2:J{last}")
    add_dv(ws, "nrMouvements", f"K2:K{last}")
    add_dv(ws, "nrCotations", f"M2:M{last}")
    add_dv(ws, "nrCotations", f"T2:T{last}")
    add_dv(ws, "nrGroupes", f"X2:X{last}")

    # MFC : doublon N° DA + naissance introuvable
    c_doub = get_column_letter(col["Doublon_N_DA"])
    ws.conditional_formatting.add(
        f"A2:{get_column_letter(len(headers))}{last}",
        FormulaRule(formula=[f'${c_doub}2="DOUBLON"'], fill=FILL_RED))


# --------------------------------------------------------------------------- #
# Saisie_Presences — 3 champs métier + Parcours, reste calculé
# --------------------------------------------------------------------------- #
def build_saisie(wb):
    ws = wb.create_sheet("Saisie_Presences")
    headers = [
        "Patient", "Date", "Programmation", "Parcours",       # A B C D (saisis)
        "Cle_Patient_Norm", "Est_Present", "N_DA", "Statut_Resolution",  # E F G H
        "Cotation", "Regime", "Nb_Je_accordes", "PEC", "Groupe",         # I J K L M
        "Semaine", "Mois", "Doublon_Cle",                                 # N O P
    ]
    write_headers(ws, headers)
    col = {h: get_column_letter(i + 1) for i, h in enumerate(headers)}

    for r in range(2, 2 + NB_TEMPLATE_ROWS):
        # E Cle_Patient_Norm
        ws[f"{col['Cle_Patient_Norm']}{r}"] = "=" + cle_norm_formula(f"A{r}")
        # F Est_Present
        ws[f"{col['Est_Present']}{r}"] = (
            '=IF([@Programmation]="","",IFERROR(' +
            XLOOKUP("[@Programmation]", "tProgrammation[Programmation]",
                    "tProgrammation[Compte_Present]", "FALSE") + ',FALSE))'
        )
        # G N_DA — formule maîtresse (FILTER + LET, gestion 0/1/N)
        ws[f"{col['N_DA']}{r}"] = (
            '=IF([@Cle_Patient_Norm]="","",'
            'IF(COUNTIF(tPatients[Cle_Norm],[@Cle_Patient_Norm])=0,"⚠ HORS LISTE",'
            '_xlfn.LET(res,' +
            FILTER("tDA[N_DA]",
                   "(tDA[Cle_Patient_Norm]=[@Cle_Patient_Norm])"
                   "*(tDA[Parcours]=[@Parcours])"
                   "*(tDA[Date_debut_accorde]<=[@Date])"
                   "*(tDA[Date_fin_accorde]>=[@Date])"
                   "*(tDA[Statut_Exclu]=FALSE)") +
            ',IF(COUNTA(res)=0,"⚠ SANS DA",'
            'IF(COUNTA(res)>1,"⚠ MULTI DA",INDEX(res,1))))))'
        )
        # H Statut_Resolution dérivé de G
        ws[f"{col['Statut_Resolution']}{r}"] = (
            f'=IF([@N_DA]="","",IF(LEFT([@N_DA],1)="⚠",'
            f'SUBSTITUTE([@N_DA],"⚠ ",""),"OK"))'
        )
        # I Cotation (depuis DA si résolu)
        ws[f"{col['Cotation']}{r}"] = (
            '=IF([@Statut_Resolution]="OK",' +
            XLOOKUP("[@N_DA]", "tDA[N_DA]", "tDA[Cotation_accordee]") + ',"")'
        )
        # J Regime
        ws[f"{col['Regime']}{r}"] = (
            '=IF([@Statut_Resolution]="OK",' +
            XLOOKUP("[@N_DA]", "tDA[N_DA]", "tDA[Regime]") + ',"")'
        )
        # K Nb_Je_accordes
        ws[f"{col['Nb_Je_accordes']}{r}"] = (
            '=IF([@Statut_Resolution]="OK",' +
            XLOOKUP("[@N_DA]", "tDA[N_DA]", "tDA[Nb_Je_Accorde]", "0") + ',"")'
        )
        # L PEC
        ws[f"{col['PEC']}{r}"] = (
            '=IF([@Statut_Resolution]="OK",' +
            XLOOKUP("[@N_DA]", "tDA[N_DA]", "tDA[PEC]") + ',"")'
        )
        # M Groupe
        ws[f"{col['Groupe']}{r}"] = (
            '=IF([@Statut_Resolution]="OK",' +
            XLOOKUP("[@N_DA]", "tDA[N_DA]", "tDA[Groupe]") + ',"")'
        )
        # N Semaine (via calendrier)
        ws[f"{col['Semaine']}{r}"] = (
            '=IF([@Date]="","",' +
            XLOOKUP("[@Date]", "tCalendrier[Date]", "tCalendrier[No_Semaine]",
                    '"⚠ hors calendrier"') + ')'
        )
        # O Mois
        ws[f"{col['Mois']}{r}"] = '=IF([@Date]="","",TEXT([@Date],"mm/aa"))'
        # P Doublon_Cle (N° DA + Date)
        ws[f"{col['Doublon_Cle']}{r}"] = (
            '=IF(OR([@N_DA]="",[@Date]=""),"",'
            'IF(COUNTIFS([N_DA],[@N_DA],[Date],[@Date])>1,"DOUBLON",""))'
        )

    last = 1 + NB_TEMPLATE_ROWS
    add_table(ws, "tPresences", 1, len(headers), last)
    autosize(ws, headers, 14)
    ws.freeze_panes = "A2"
    fmt_text_column(ws, 7, last)          # N_DA texte
    fmt_date_columns(ws, [2], last)       # Date

    # validations
    add_dv(ws, "nrPatients", f"A2:A{last}", warn=True)   # avertir sans bloquer
    add_dv(ws, "nrProgrammation", f"C2:C{last}")
    add_dv(ws, "nrParcours", f"D2:D{last}")

    # MFC : ligne rouge si anomalie de résolution ou doublon
    c_stat = col["Statut_Resolution"]
    c_doub = col["Doublon_Cle"]
    rng = f"A2:{get_column_letter(len(headers))}{last}"
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'OR(${c_stat}2="HORS LISTE",${c_stat}2="SANS DA",${c_stat}2="MULTI DA")'],
        fill=FILL_RED))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'${c_doub}2="DOUBLON"'], fill=FILL_RED))


# --------------------------------------------------------------------------- #
# CONSO_Presences — image de Saisie_Presences par référence structurée
# --------------------------------------------------------------------------- #
def build_conso(wb):
    ws = wb.create_sheet("CONSO_Presences")
    headers = [
        "Patient", "Cle_Patient_Norm", "Date", "Programmation", "Est_Present",
        "Parcours", "N_DA", "Cotation", "Regime", "Nb_Je_accordes",
        "Semaine", "Mois", "Groupe", "PEC", "Statut_Resolution",
    ]
    write_headers(ws, headers)
    # mapping vers colonnes de tPresences
    src = {
        "Patient": "Patient", "Cle_Patient_Norm": "Cle_Patient_Norm",
        "Date": "Date", "Programmation": "Programmation", "Est_Present": "Est_Present",
        "Parcours": "Parcours", "N_DA": "N_DA", "Cotation": "Cotation",
        "Regime": "Regime", "Nb_Je_accordes": "Nb_Je_accordes",
        "Semaine": "Semaine", "Mois": "Mois", "Groupe": "Groupe",
        "PEC": "PEC", "Statut_Resolution": "Statut_Resolution",
    }
    for r in range(2, 2 + NB_TEMPLATE_ROWS):
        for j, h in enumerate(headers, start=1):
            # Image ligne à ligne de tPresences par INDEX (référence structurée).
            ws.cell(row=r, column=j,
                    value=f'=IF(ROW()-1<=ROWS(tPresences),'
                          f'INDEX(tPresences[{src[h]}],ROW()-1),"")')
    last = 1 + NB_TEMPLATE_ROWS
    add_table(ws, "tConso", 1, len(headers), last)
    autosize(ws, headers, 14)
    ws.freeze_panes = "A2"
    fmt_text_column(ws, 7, last)
    fmt_date_columns(ws, [3], last)
    # note explicative
    note(ws, len(headers) + 2,
         "CONSO_Presences = image directe de Saisie_Presences (INDEX par référence "
         "structurée). Unification ETP+Polyvalent par la colonne Parcours.")


# --------------------------------------------------------------------------- #
# Suivi_Factures — registre de suivi (TabSuiviApi cible), jointures intra-classeur
# --------------------------------------------------------------------------- #
def build_suivi(wb):
    ws = wb.create_sheet("Suivi_Factures")
    headers = [
        "Mois", "Date_facture", "N_semaine", "Ref", "N_DE_FACTURE",
        "Code_recherche", "Type_de_facture", "Cotation", "Type_PEC", "DA",
        "Regime", "DN", "Nom", "Prenom", "Date_naissance", "Adresse",
        "Commune", "Telephone", "Date_debut", "Date_fin", "Nb_journees",
        "Tarif_unitaire", "Montant", "Categorie_Refus", "Date_depot",
        "Eligible", "Motif_rejet",
    ]
    write_headers(ws, headers)
    col = {h: get_column_letter(i + 1) for i, h in enumerate(headers)}

    for r in range(2, 2 + NB_TEMPLATE_ROWS):
        ws[f"{col['Mois']}{r}"] = '=IF([@Date_debut]="","",MONTH([@Date_debut]))'
        # Type_de_facture dérivé du régime
        ws[f"{col['Type_de_facture']}{r}"] = (
            '=IF([@Regime]="","",' +
            XLOOKUP("[@Regime]", "tRegimes[Regime]", "tRegimes[Payeur]", '"⚠"') + ')'
        )
        # Regime depuis DA
        ws[f"{col['Regime']}{r}"] = "=" + XLOOKUP("[@DA]", "tDA[N_DA]", "tDA[Regime]")
        # Patient depuis tPatients via Code_recherche
        for field, tcol in (("DN", "DN"), ("Nom", "Nom"), ("Prenom", "Prenom"),
                            ("Date_naissance", "Date_naissance"),
                            ("Adresse", "Adresse"), ("Commune", "Commune"),
                            ("Telephone", "N_tel")):
            ws[f"{col[field]}{r}"] = "=" + XLOOKUP(
                "[@Code_recherche]", "tPatients[Recherche]", f"tPatients[{tcol}]")
        # Type_PEC depuis cotation
        ws[f"{col['Type_PEC']}{r}"] = "=" + XLOOKUP(
            "[@Cotation]", "tTarif[Code_PEC]", "tTarif[PEC]")
        # Nb_journees = présences de la DA sur la semaine (compte présent)
        ws[f"{col['Nb_journees']}{r}"] = (
            '=IF(OR([@DA]="",[@N_semaine]=""),"",'
            'COUNTIFS(CONSO_Presences[N_DA],[@DA],'
            'CONSO_Presences[Semaine],[@N_semaine],'
            'CONSO_Presences[Est_Present],TRUE))'
        )
        # Tarif_unitaire
        ws[f"{col['Tarif_unitaire']}{r}"] = (
            '=IF([@Cotation]="","",' +
            XLOOKUP("[@Cotation]", "tTarif[Code_PEC]", "tTarif[Tarif_XPF]",
                    '"⚠ tarif inconnu"') + ')'
        )
        # Montant = tarif * nb journées
        ws[f"{col['Montant']}{r}"] = (
            '=IF(OR([@Cotation]="",[@Nb_journees]="",'
            '[@Tarif_unitaire]="⚠ tarif inconnu"),"",'
            '[@Tarif_unitaire]*[@Nb_journees])'
        )
        # Eligible (règles §7.4)
        ws[f"{col['Eligible']}{r}"] = (
            '=IF([@DA]="","",'
            'AND(' +
            XLOOKUP("[@Regime]", "tRegimes[Regime]", "tRegimes[Facturable_CPS]", "FALSE") +
            ',COUNTIF(tTarif[Code_PEC],[@Cotation])>0'
            ',[@Categorie_Refus]=""'
            ',[@Nb_journees]<=' + XLOOKUP("[@DA]", "tDA[N_DA]", "tDA[Nb_Je_Accorde]", "0") +
            '))'
        )
        # Motif_rejet : 1re condition non satisfaite
        ws[f"{col['Motif_rejet']}{r}"] = (
            '=IF([@DA]="","",'
            'IF(' + XLOOKUP("[@Regime]", "tRegimes[Regime]", "tRegimes[Facturable_CPS]", "FALSE") +
            '=FALSE,"Régime non CPS",'
            'IF(COUNTIF(tTarif[Code_PEC],[@Cotation])=0,"Cotation invalide",'
            'IF([@Categorie_Refus]<>"","Refus: "&[@Categorie_Refus],'
            'IF([@Nb_journees]>' + XLOOKUP("[@DA]", "tDA[N_DA]", "tDA[Nb_Je_Accorde]", "0") +
            ',"Plafond Nb Je dépassé","")))))'
        )

    last = 1 + NB_TEMPLATE_ROWS
    add_table(ws, "tSuiviFactures", 1, len(headers), last)
    autosize(ws, headers, 13)
    ws.freeze_panes = "A2"
    fmt_text_column(ws, headers.index("N_DE_FACTURE") + 1, last)
    fmt_text_column(ws, headers.index("DA") + 1, last)
    fmt_date_columns(ws, [headers.index(c) + 1 for c in
                          ("Date_facture", "Date_debut", "Date_fin", "Date_depot")], last)

    # validations
    add_dv(ws, "nrSemaine", f"C2:C{last}")
    add_dv(ws, "nrPatients", f"F2:F{last}", warn=True)
    add_dv(ws, "nrCotations", f"H2:H{last}")
    add_dv(ws, "nrCategoriesRefus", f"X2:X{last}")

    # MFC
    c_ref = col["Categorie_Refus"]
    c_tar = col["Tarif_unitaire"]
    rng = f"A2:{get_column_letter(len(headers))}{last}"
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'${c_ref}2<>""'], fill=FILL_RED))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'${c_tar}2="⚠ tarif inconnu"'], fill=FILL_RED))


# --------------------------------------------------------------------------- #
# Bordereaux — totaux par N° de facture (SUMIFS)
# --------------------------------------------------------------------------- #
def build_bordereaux(wb):
    ws = wb.create_sheet("Bordereaux")
    headers = ["No_Bordereau", "N_FACTURE", "DN", "Periode", "Nb_JRS", "Montant"]
    write_headers(ws, headers)
    col = {h: get_column_letter(i + 1) for i, h in enumerate(headers)}
    for r in range(2, 2 + NB_TEMPLATE_ROWS):
        ws[f"{col['No_Bordereau']}{r}"] = (
            f'=IF([@N_FACTURE]="","",ROW()-1)')
        ws[f"{col['DN']}{r}"] = (
            '=IF([@N_FACTURE]="","",' +
            XLOOKUP("[@N_FACTURE]", "tSuiviFactures[N_DE_FACTURE]",
                    "tSuiviFactures[DN]") + ')'
        )
        ws[f"{col['Montant']}{r}"] = (
            '=IF([@N_FACTURE]="","",'
            'SUMIFS(tSuiviFactures[Montant],'
            'tSuiviFactures[N_DE_FACTURE],[@N_FACTURE]))'
        )
    last = 1 + NB_TEMPLATE_ROWS
    add_table(ws, "tBordereaux", 1, len(headers), last)
    autosize(ws, headers, 18)
    ws.freeze_panes = "A2"
    add_dv(ws, "nrFactures", f"B2:B{last}")
    # plage nommée pour la liste des n° de facture
    add_named_range_simple(wb, "nrFactures", "Suivi_Factures", "E", 2, last)


# --------------------------------------------------------------------------- #
# Facture — modèle imprimable (1 facture) + montant en lettres (LAMBDA)
# --------------------------------------------------------------------------- #
def build_facture(wb):
    ws = wb.create_sheet("Facture")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "SSRP Ora Ora — Facture (modèle imprimable)"
    ws["A1"].font = F_TITLE

    def lbl(cell, text):
        ws[cell] = text
        ws[cell].font = Font(name=ARIAL, size=10, bold=True)
    def val(cell, formula):
        ws[cell] = formula
        ws[cell].font = F_BASE

    # Sélecteur de facture
    lbl("A3", "FACTURE N°")
    ws["B3"] = "F0012026"
    ws["B3"].font = F_BASE
    ws["B3"].fill = FILL_NOTE
    ws["B3"].number_format = "@"
    add_dv(ws, "nrFactures", "B3")

    # En-tête patient (XLOOKUP sur le suivi)
    def lookup_suivi(retcol):
        return "=" + XLOOKUP("$B$3", "tSuiviFactures[N_DE_FACTURE]",
                             f"tSuiviFactures[{retcol}]")
    lbl("A5", "Nom :");       val("B5", lookup_suivi("Nom"))
    lbl("A6", "Prénom :");    val("B6", lookup_suivi("Prenom"))
    lbl("A7", "Né(e) le :");  val("B7", lookup_suivi("Date_naissance"))
    lbl("A8", "Adresse :");   val("B8", lookup_suivi("Adresse"))
    lbl("A9", "Commune :");   val("B9", lookup_suivi("Commune"))
    lbl("A10", "Téléphone :"); val("B10", lookup_suivi("Telephone"))
    lbl("A11", "ORGANISME PAYEUR :"); val("B11", lookup_suivi("Type_de_facture"))
    lbl("A12", "Date de facture :");  val("B12", lookup_suivi("Date_facture"))

    # Ligne de prestation
    lbl("A14", "Cotation :"); val("B14", lookup_suivi("Cotation"))
    # B15 = N° DA brut (clé de jointure) ; B15bis affiche "DA xxx"
    lbl("A15", "DA :")
    val("B15", "=" + XLOOKUP("$B$3", "tSuiviFactures[N_DE_FACTURE]", "tSuiviFactures[DA]"))
    ws["B15"].number_format = "@"
    lbl("A16", "Nb jours accordés :")
    val("B16", "=" + XLOOKUP("B15", "tDA[N_DA]", "tDA[Nb_Je_Accorde]", "0"))
    lbl("A17", "Prestation :")
    val("B17", "=" + XLOOKUP("B14", "tTarif[Code_PEC]", "tTarif[Prestation]"))
    lbl("A18", "Tarif unitaire :")
    val("B18", "=" + XLOOKUP("B14", "tTarif[Code_PEC]", "tTarif[Tarif_XPF]", "0"))
    lbl("A19", "Nb de journées :")
    val("B19", "=" + XLOOKUP("$B$3", "tSuiviFactures[N_DE_FACTURE]",
                             "tSuiviFactures[Nb_journees]", "0"))
    lbl("A21", "TOTAL :")
    val("B21", '=IF(OR(B18="",B19=""),"",B18*B19)')
    ws["B21"].font = Font(name=ARIAL, size=11, bold=True)
    lbl("A22", "Montant en lettres :")
    val("B22", '=IF(B21="","",MontantEnLettres(B21)&" FRANCS CFP")')

    note(ws, 24,
         "[à affiner en QA] Le montant en lettres utilise la LAMBDA 'MontantEnLettres' "
         "(defined name). Bornée 0–999 999 999 XPF, sans décimales (XPF entier). "
         "Le montant NUMÉRIQUE (B21) reste la valeur de référence de la facture.")
    note(ws, 25,
         "Canal de transmission CPS NON tranché (CLAUDE.md §7) : ce modèle = "
         "bordereau/facture interne validé, pas le format CPS final.")

    for r in range(1, 26):
        ws.row_dimensions[r].height = 16
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 42

    # LAMBDA montant en lettres (français, bornée). Definie au niveau classeur.
    add_montant_en_lettres_lambda(wb)


# --------------------------------------------------------------------------- #
# Cockpit — squelette TCD + compteurs de synthèse
# --------------------------------------------------------------------------- #
def build_cockpit(wb):
    ws = wb.create_sheet("Cockpit")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Cockpit — pilotage (réel + prévisionnel)"
    ws["A1"].font = F_TITLE
    note(ws, 3, "Zone réservée aux Tableaux Croisés Dynamiques (TCD) — à poser "
                "manuellement sur CONSO_Presences et DA (Phase QA).")
    ws["A5"] = "Synthèse rapide (formules) :"
    ws["A5"].font = Font(name=ARIAL, bold=True)
    rows = [
        ("Présences ETP (comptées présent)",
         '=COUNTIFS(CONSO_Presences[Parcours],"ETP",CONSO_Presences[Est_Present],TRUE)'),
        ("Présences Polyvalent (comptées présent)",
         '=COUNTIFS(CONSO_Presences[Parcours],"Polyvalent",CONSO_Presences[Est_Present],TRUE)'),
        ("Total présences comptées présent",
         '=COUNTIF(CONSO_Presences[Est_Present],TRUE)'),
        ("Nb DA enregistrées",
         '=COUNTA(tDA[N_DA])'),
        ("Montant total facturé (suivi)",
         '=SUM(tSuiviFactures[Montant])'),
    ]
    for i, (label, f) in enumerate(rows, start=6):
        ws[f"A{i}"] = label
        ws[f"A{i}"].font = F_BASE
        ws[f"B{i}"] = f
        ws[f"B{i}"].font = F_BASE
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 18


# --------------------------------------------------------------------------- #
# CTRL_Qualite — compteurs d'anomalies + cohérence
# --------------------------------------------------------------------------- #
def build_ctrl(wb):
    ws = wb.create_sheet("CTRL_Qualite")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "CTRL_Qualite — contrôles d'intégrité"
    ws["A1"].font = F_TITLE
    ws["A3"] = "Indicateur"
    ws["B3"] = "Valeur"
    ws["C3"] = "Seuil OK"
    for c in ("A3", "B3", "C3"):
        ws[c].font = F_HDR
        ws[c].fill = FILL_HDR
    checks = [
        ("Présences HORS LISTE",
         '=COUNTIF(tPresences[Statut_Resolution],"HORS LISTE")', "= 0"),
        ("Présences SANS DA",
         '=COUNTIF(tPresences[Statut_Resolution],"SANS DA")', "= 0"),
        ("Présences MULTI DA",
         '=COUNTIF(tPresences[Statut_Resolution],"MULTI DA")', "= 0"),
        ("Doublons clé (N° DA + Date)",
         '=COUNTIF(tPresences[Doublon_Cle],"DOUBLON")', "= 0"),
        ("Doublons N° DA (registre DA)",
         '=COUNTIF(tDA[Doublon_N_DA],"DOUBLON")', "= 0"),
        ("Tarifs inconnus (suivi)",
         '=COUNTIF(tSuiviFactures[Tarif_unitaire],"⚠ tarif inconnu")', "= 0"),
        ("Factures avec catégorie de refus",
         '=COUNTIF(tSuiviFactures[Categorie_Refus],"<>")', "info"),
        ("Lignes suivi éligibles",
         '=COUNTIF(tSuiviFactures[Eligible],TRUE)', "info"),
        ("Cohérence : Σ Bordereaux − Σ Suivi",
         '=SUM(tBordereaux[Montant])-SUM(tSuiviFactures[Montant])', "= 0"),
        ("Total lignes de saisie",
         '=COUNTA(tPresences[Patient])', "info"),
    ]
    for i, (label, f, seuil) in enumerate(checks, start=4):
        ws[f"A{i}"] = label; ws[f"A{i}"].font = F_BASE
        ws[f"B{i}"] = f;      ws[f"B{i}"].font = F_BASE
        ws[f"C{i}"] = seuil;  ws[f"C{i}"].font = F_BASE
    # MFC : valeur > 0 sur les contrôles "= 0"
    ws.conditional_formatting.add(
        "B4:B9", FormulaRule(formula=['B4>0'], fill=FILL_RED))
    ws.conditional_formatting.add(
        "B12:B12", FormulaRule(formula=['B12<>0'], fill=FILL_RED))
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 10


# --------------------------------------------------------------------------- #
# LAMBDA "MontantEnLettres" — version française bornée (defined name)
# --------------------------------------------------------------------------- #
def add_montant_en_lettres_lambda(wb):
    """
    Convertit un entier 0..999 999 999 en lettres françaises (XPF entier).
    Construite par composition de LAMBDA imbriquées via LET. Bornée et signalée
    [à affiner en QA] : gère unités/dizaines/centaines/milliers/millions, avec
    règles 'quatre-vingt(s)', 'cent(s)', 'et un'. Pas de centimes (XPF entier).
    """
    # Tables internes encodées en CHOOSE pour rester 100% formule.
    u = ('CHOOSE(n+1,"zéro","un","deux","trois","quatre","cinq","six","sept",'
         '"huit","neuf","dix","onze","douze","treize","quatorze","quinze","seize",'
         '"dix-sept","dix-huit","dix-neuf")')
    # dizaines de base (20..90 par pas) gérées dans la lambda cent2
    formula = (
        '=_xlfn.LAMBDA(montant,_xlfn.LET('
        # --- helper "deuxchiffres" 0..99 ---
        'dz,_xlfn.LAMBDA(n,_xlfn.LET('
        '  d,INT(n/10),uni,MOD(n,10),'
        '  IF(n<20,' + u + ','
        '  IF(d=2,IF(uni=0,"vingt","vingt-"&' + u.replace('n+1', 'uni+1') + '),'
        '  IF(d=3,IF(uni=0,"trente","trente-"&' + u.replace('n+1', 'uni+1') + '),'
        '  IF(d=4,IF(uni=0,"quarante","quarante-"&' + u.replace('n+1', 'uni+1') + '),'
        '  IF(d=5,IF(uni=0,"cinquante","cinquante-"&' + u.replace('n+1', 'uni+1') + '),'
        '  IF(d=6,IF(uni=0,"soixante","soixante-"&' + u.replace('n+1', 'uni+1') + '),'
        '  IF(d=7,IF(uni=0,"soixante-dix","soixante-"&' + u.replace('n+1', '(uni+10)+1') + '),'
        '  IF(d=8,IF(uni=0,"quatre-vingts","quatre-vingt-"&' + u.replace('n+1', 'uni+1') + '),'
        '  IF(uni=0,"quatre-vingt-dix","quatre-vingt-"&' + u.replace('n+1', '(uni+10)+1') + ')'
        '  )))))))))),'
        # --- helper "troischiffres" 0..999 ---
        'ct,_xlfn.LAMBDA(n,_xlfn.LET('
        '  c,INT(n/100),reste,MOD(n,100),'
        '  IF(n=0,"",'
        '  IF(c=0,dz(reste),'
        '  IF(c=1,IF(reste=0,"cent","cent "&dz(reste)),'
        '  ' + u.replace('n+1', 'c+1') + '&IF(reste=0," cents"," cent "&dz(reste)))))'
        ')),'
        # --- assemblage millions / milliers / unités ---
        'mil,INT(montant/1000000),'
        'milel,INT(MOD(montant,1000000)/1000),'
        'uni,MOD(montant,1000),'
        'pMil,IF(mil=0,"",IF(mil=1,"un million ",ct(mil)&" millions ")),'
        'pMille,IF(milel=0,"",IF(milel=1,"mille ",ct(milel)&" mille ")),'
        'pUni,IF(uni=0,IF(montant=0,"zéro",""),ct(uni)),'
        'TRIM(pMil&pMille&pUni)'
        '))'
    )
    wb.defined_names.add(DefinedName("MontantEnLettres", attr_text=formula))


# --------------------------------------------------------------------------- #
# Utilitaires divers
# --------------------------------------------------------------------------- #
def add_dv(ws, named_range, cell_range, warn=False):
    dv = list_validation(named_range, warn=warn)
    ws.add_data_validation(dv)
    dv.add(cell_range)

def add_named_range_simple(wb, name, sheet, col_letter, first, last):
    if name in wb.defined_names:
        del wb.defined_names[name]
    ref = f"'{sheet}'!${col_letter}${first}:${col_letter}${last}"
    wb.defined_names.add(DefinedName(name, attr_text=ref))

def fmt_text_column(ws, col_idx, last):
    for r in range(2, last + 1):
        ws.cell(row=r, column=col_idx).number_format = "@"

def fmt_date_columns(ws, col_indices, last):
    for c in col_indices:
        for r in range(2, last + 1):
            ws.cell(row=r, column=c).number_format = "DD/MM/YYYY"

def note(ws, row, text):
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(name=ARIAL, size=9, italic=True, color="7F6000")
    cell.fill = FILL_NOTE
    cell.alignment = Alignment(wrap_text=False)


# --------------------------------------------------------------------------- #
# Recalcul LibreOffice (option) — détection grossière d'erreurs de formule
# --------------------------------------------------------------------------- #
def recalc():
    soffice = None
    for cand in ("soffice", "libreoffice"):
        if subprocess.run(["which", cand], capture_output=True).returncode == 0:
            soffice = cand
            break
    if not soffice:
        print("LibreOffice introuvable — recalcul ignoré.")
        return
    print(f"Recalcul via {soffice} (headless)...")
    outdir = os.path.join(OUT_DIR, "_recalc")
    os.makedirs(outdir, exist_ok=True)
    r = subprocess.run(
        [soffice, "--headless", "-env:UserInstallation=file:///tmp/ora_lo",
         "--convert-to", "xlsx", "--outdir", outdir, OUT_FILE],
        capture_output=True, timeout=180, text=True)
    produced = os.path.join(outdir, os.path.basename(OUT_FILE))
    if os.path.exists(produced):
        print(f"Recalcul OK -> {produced}")
    else:
        print("ATTENTION : LibreOffice n'a pas pu charger le classeur dans cet "
              "environnement (limitation sandbox connue). Sortie soffice :",
              (r.stderr or r.stdout or "").strip()[:200])
    print("NB : LibreOffice ne simule pas parfaitement FILTER/XLOOKUP/LET "
          "dynamiques — validation finale obligatoire dans Excel 365 (Phase QA).")


# --------------------------------------------------------------------------- #
if __name__ == "__main__":
    wb = build()
    print(f"OK -> {OUT_FILE}")
    print(f"Onglets ({len(wb.sheetnames)}): {', '.join(wb.sheetnames)}")
    if "--recalc" in sys.argv:
        recalc()
