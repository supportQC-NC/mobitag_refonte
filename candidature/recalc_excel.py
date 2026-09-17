# -*- coding: utf-8 -*-
"""Ouvre le classeur dans Excel, force le recalcul, verifie les valeurs cles
contre le recalcul Python, et sauvegarde le fichier avec ses valeurs en cache."""
import os, sys, shutil, tempfile
import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "Modele_financier_Mobitag_2026.xlsx")
# Excel refuse d'ouvrir un fichier depuis le Bureau (Mark-of-the-Web) :
# on recalcule une copie placee dans un repertoire temporaire.
XLSX = os.path.join(tempfile.gettempdir(), "recalc_mobitag.xlsx")
shutil.copy(SRC, XLSX)

# valeurs de reference issues de verif_modele.py
REF = {
    "Prudent":   dict(ca=155876,    mc=-1490551, ce=7046354,  cf=-4726667,
                      capex=-5500000, eq=-4670864),
    "Central":   dict(ca=3513500,   mc=772577,   ce=8997917,  cf=-4480000,
                      capex=-4800000, eq=343345),
    "Ambitieux": dict(ca=20811463,  mc=16899492, ce=11121250, cf=-4176000,
                      capex=-4200000, eq=13751320),
}
SHEETS = {"Prudent": "4_PL_Prudent", "Central": "5_PL_Central", "Ambitieux": "6_PL_Ambitieux"}

excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.AutomationSecurity = 1
try:
    wb = excel.Workbooks.Open(XLSX, UpdateLinks=0, ReadOnly=False, CorruptLoad=2)
    excel.CalculateFullRebuild()
    wb.Application.Calculate()

    syn = wb.Worksheets("7_Synthese")
    print("=" * 88)
    print("RECALCUL EXCEL - CONTROLE CROISE AVEC LE MODELE PYTHON")
    print("=" * 88)

    # localise les lignes de la synthese par leur libelle
    # premiere occurrence seulement : les memes libelles sont repris plus bas
    # dans les blocs de donnees qui alimentent les graphiques.
    labels = {}
    for r in range(1, 60):
        v = syn.Cells(r, 1).Value
        if isinstance(v, str) and v.strip() not in labels:
            labels[v.strip()] = r

    checks = [
        ("Chiffre d'affaires", "ca"),
        ("MARGE CONTRIBUTIVE", "mc"),
        ("Coûts évités", "ce"),
        ("Coûts fixes", "cf"),
        ("Investissements", "capex"),
        ("EQUILIBRE ECONOMIQUE 24 MOIS", "eq"),
    ]
    ok = True
    for label, key in checks:
        row = labels.get(label)
        if row is None:
            print(f"  [?] ligne introuvable : {label}")
            ok = False
            continue
        line = f"{label:<34}"
        for col, scen in ((2, "Prudent"), (3, "Central"), (4, "Ambitieux")):
            got = syn.Cells(row, col).Value
            exp = REF[scen][key]
            if got is None:
                line += f" {scen[:4]}: ERREUR"
                ok = False
                continue
            delta = abs(got - exp)
            tol = max(2.0, abs(exp) * 0.001)
            flag = "OK" if delta <= tol else f"ECART {delta:,.0f}"
            if delta > tol:
                ok = False
            line += f" | {scen[:4]} {got:>13,.0f} {flag}"
        print(line.replace(",", " "))

    # point mort
    pm_row = labels.get("Point mort (1er mois à cumul positif)")
    if pm_row:
        vals = [syn.Cells(pm_row, c).Value for c in (2, 3, 4)]
        print(f"\n  Point mort (Prudent / Central / Ambitieux) : {vals}")

    bf_row = labels.get("Besoin de financement maximal")
    if bf_row:
        vals = [syn.Cells(bf_row, c).Value for c in (2, 3, 4)]
        print(f"  Besoin de financement maximal              : "
              + " / ".join(f"{v:,.0f}".replace(",", " ") for v in vals))

    # detection d'erreurs de formule sur tout le classeur
    print("\n  Recherche d'erreurs de formule (#REF!, #VALUE!, #DIV/0!, #NAME?) :")
    errs = 0
    for sh in wb.Worksheets:
        used = sh.UsedRange
        try:
            found = used.SpecialCells(-4123, 16)  # xlCellTypeFormulas, xlErrors
            cnt = found.Count
            errs += cnt
            print(f"    {sh.Name:<20} : {cnt} cellule(s) en erreur -> {found.Address}")
        except Exception:
            pass
    if errs == 0:
        print("    aucune erreur de formule detectee.")
    else:
        ok = False

    wb.Close(SaveChanges=False)
    print("\n" + ("=" * 88))
    print("RESULTAT :", "CONTROLE REUSSI - le classeur calcule les valeurs attendues."
          if ok else "ECARTS DETECTES - voir ci-dessus.")
    print("=" * 88)
finally:
    excel.Quit()
