# -*- coding: utf-8 -*-
"""
Challenge Mobitag 2026 - Modele financier 24 mois, 3 scenarios.
Genere un classeur XLSX a formules vivantes.

Perimetre : compte de resultat du SERVICE Mobitag Nouvelle Generation (vue OPT-NC).
Aucune ligne de revenu ne repose sur la vente de volume SMS au particulier :
le SMS est deja illimite dans les forfaits M et les recharges packagees Helia.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference

# ---------------------------------------------------------------- style -----
NAVY, BLUE, LBLUE = "0B2B46", "1F6FB2", "DCE9F5"
GREY, LGREEN, LAMBER, WHITE = "F2F4F7", "DDF0E5", "FBEEDB", "FFFFFF"

F_TITLE = Font(name="Calibri", size=16, bold=True, color=NAVY)
F_SUB   = Font(name="Calibri", size=11, bold=True, color=BLUE)
F_H1    = Font(name="Calibri", size=10, bold=True, color=WHITE)
F_H2    = Font(name="Calibri", size=10, bold=True, color=NAVY)
F_BOLD  = Font(name="Calibri", size=10, bold=True)
F_BASE  = Font(name="Calibri", size=10)
F_SMALL = Font(name="Calibri", size=8, italic=True, color="666666")
F_TOT   = Font(name="Calibri", size=10, bold=True, color=NAVY)

FILL_H1  = PatternFill("solid", fgColor=NAVY)
FILL_H2  = PatternFill("solid", fgColor=LBLUE)
FILL_IN  = PatternFill("solid", fgColor="FFF6D6")
FILL_TOT = PatternFill("solid", fgColor=GREY)
FILL_GRN = PatternFill("solid", fgColor=LGREEN)
FILL_AMB = PatternFill("solid", fgColor=LAMBER)

THIN = Side(style="thin", color="BFC7D1")
B_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FMT_XPF = '#,##0\\ "F"'
FMT_NUM = '#,##0'
FMT_DEC = '#,##0.0'
FMT_PCT = '0.0%'
FMT_PCT2 = '0.00%'

MONTHS = 24
C0 = 3
def CL(m): return get_column_letter(C0 + m - 1)      # M1 -> C ... M24 -> Z
LASTC = CL(MONTHS)

HYP = "'1_Hypotheses'"
CEV = "'2_Couts_evites'"
SCEN = [("Prudent", "C", "4_PL_Prudent"),
        ("Central", "D", "5_PL_Central"),
        ("Ambitieux", "E", "6_PL_Ambitieux")]

wb = Workbook()

def band(ws, row, last_col, label, fill=FILL_H2, font=F_H2):
    for c in range(1, last_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.border = B_ALL
    cell = ws.cell(row=row, column=1, value=label)
    cell.font = font


# ===========================================================================
# 0 - LISEZ-MOI
# ===========================================================================
ws = wb.active
ws.title = "0_Lisez-moi"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 36
ws.column_dimensions["C"].width = 98

r = 2
ws.cell(row=r, column=2, value="CHALLENGE MOBITAG 2026 - MODELE FINANCIER 24 MOIS").font = F_TITLE
r += 1
ws.cell(row=r, column=2, value="Candidat : [RAISON SOCIALE]  -  Proposition « Mobitag Nouvelle Génération »").font = F_SUB
r += 1
ws.cell(row=r, column=2, value="Montants en francs CFP (XPF), hors taxes. Parité de référence : 1 EUR = 119,3317 XPF (parité fixe).").font = F_SMALL
r += 2

README = [
    ("OBJET", None),
    ("", "Réponse à la section 6 du dossier candidats : projection 24 mois en F CFP distinguant chiffre d'affaires, "
         "coûts variables, coûts fixes, investissements, économies générées, fiscalité, reversements, résultat "
         "économique et besoin de financement, en trois scénarios."),
    ("", ""),
    ("PERIMETRE DU COMPTE DE RESULTAT", None),
    ("", "Le P&L modélisé est celui du SERVICE Mobitag Nouvelle Génération dans son ensemble, en vue OPT-NC. "
         "Ce choix est imposé par la formule d'équilibre du dossier, qui intègre les coûts évités : "
         "ces coûts sont ceux de l'OPT-NC. La répartition de la valeur entre l'OPT-NC, le candidat et les "
         "partenaires fait l'objet de l'onglet 9."),
    ("", ""),
    ("PRINCIPE DIRECTEUR : AUCUNE REDONDANCE AVEC LE CATALOGUE HELIA", None),
    ("", "Le SMS est déjà illimité pour la quasi-totalité des clients Helia : les quatre Forfaits M (2 Go à 100 Go, "
         "1 000 à 10 000 F TTC/mois) incluent tous les SMS illimités, les Recharges Packagées également, et une option "
         "Pack SMS existe dès 318 F TTC. Toute monetisation fondée sur la vente de volume, de quota, de crédit ou de "
         "palier d'envoi au particulier serait donc redondante par construction."),
    ("", "AUCUNE ligne de ce modèle ne repose sur la vente de SMS au particulier. Le socle d'envoi reste gratuit, "
         "sans plafond payant et sans option d'extension de volume. Voir la cartographie de non-redondance "
         "dans la note de candidature."),
    ("", ""),
    ("LES CINQ SOURCES DE SOUTENABILITE RETENUES", None),
    ("1. TCO legacy évité  (pilier)",
        "Arrêt du socle technique historique et de sa TMA. Contributeur principal à l'équilibre : démontrable, "
        "immédiat, et sans aucun risque de cannibalisation puisqu'il s'agit de charges internes et non d'une offre. "
        "Détail poste par poste dans l'onglet 2."),
    ("2. Coûts d'abus évités",
        "Le dispositif de confiance sur le canal 1000 (vérification d'origine, signalement en un clic, maîtrise par le "
        "destinataire) réduit mécaniquement le traitement manuel des réclamations et des abus côté OPT-NC. "
        "Ce poste porte son propre taux d'évitement dans l'onglet 2."),
    ("3. Apport d'affaires B2B vers Helia PRO",
        "Mise en œuvre monetisée de l'exigence F10 : les usages professionnels détectés sont orientés vers "
        "Helia PRO. Anti-cannibalisation par construction, puisque le service alimente l'offre professionnelle au lieu "
        "de la doubler. Valorisé en marge attribuée, avec un taux d'attribution volontairement prudent."),
    ("4. Apport d'affaires B2C - activation de lignes",
        "Mobitag est le seul point de contact de l'OPT-NC avec les personnes sans ligne active : téléphone perdu, "
        "volé, cassé, hors forfait, poste partagé d'une mairie ou d'une médiathèque. Cette audience est "
        "structurellement inaccessible au reste du catalogue Helia. Orientation clairement identifiée, jamais "
        "dissimulée, jamais insérée dans un message personnel."),
    ("5. Sponsoring consenti",
        "Partenariat identifié et accepté par l'utilisateur, hors du contenu des messages. Non redondant : l'OPT-NC "
        "ne commercialise pas d'espace publicitaire. Inventaire étroit, donc volontairement NUL en scénario prudent : "
        "le modèle ne dépend jamais de cette ligne."),
    ("", ""),
    ("FORMULES IMPOSEES PAR LE DOSSIER (section 6)", None),
    ("Marge contributive", "= chiffre d'affaires − coûts variables − reversements éventuels"),
    ("Equilibre économique 24 mois",
        "= marge contributive + coûts évités − coûts fixes − investissements − fiscalité applicable"),
    ("", "Ces deux formules sont appliquées littéralement dans les onglets 4 à 6, aux lignes "
         "« Marge contributive » et « Equilibre économique cumulé »."),
    ("", ""),
    ("MODE D'EMPLOI", None),
    ("", "Seules les cellules sur fond JAUNE sont saisissables. Toutes les autres sont calculées. Modifier une "
         "hypothèse dans l'onglet 1 met à jour les trois scénarios, le point mort, la synthèse et la sensibilité."),
    ("", "Chaque hypothèse porte une source et un niveau de confiance. Les hypothèses marquées "
         "« A VALIDER OPT » sont celles dont la confirmation est demandée à l'OPT-NC : volumétrie réelle du service, "
         "coût complet actuel, et part des envois émanant d'utilisateurs sans ligne Helia active."),
    ("", ""),
    ("SOMMAIRE", None),
    ("1_Hypotheses", "Hypothèses paramétrables en trois scénarios, avec source et niveau de confiance."),
    ("2_Couts_evites", "Coût complet annuel actuel du service, poste par poste, et taux d'évitement retenu."),
    ("3_Unit_economics", "Coûts unitaires, coût de revient d'un SMS gratuit, valeur d'un apport, ratios de contrôle."),
    ("4_PL_Prudent", "Compte de résultat mensuel sur 24 mois - scénario prudent."),
    ("5_PL_Central", "Compte de résultat mensuel sur 24 mois - scénario central."),
    ("6_PL_Ambitieux", "Compte de résultat mensuel sur 24 mois - scénario ambitieux."),
    ("7_Synthese", "Comparaison des trois scénarios : point mort, résultat cumulé, besoin de financement, graphiques."),
    ("8_Sensibilite", "Sensibilité du résultat 24 mois aux sept variables les plus structurantes."),
    ("9_Partage_valeur", "Répartition de la valeur entre OPT-NC, candidat et partenaires, par phase."),
]
for label, txt in README:
    if txt is None:
        ws.cell(row=r, column=2, value=label).font = F_H2
        for c in (2, 3):
            ws.cell(row=r, column=c).fill = FILL_H2
            ws.cell(row=r, column=c).border = B_ALL
        r += 1
        continue
    ws.cell(row=r, column=2, value=label).font = F_BOLD
    ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="top")
    cc = ws.cell(row=r, column=3, value=txt)
    cc.font = F_BASE
    cc.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = (13.5 * (1 + len(txt) // 108)) if txt else 6
    r += 1


# ===========================================================================
# 1 - HYPOTHESES
# ===========================================================================
hs = wb.create_sheet("1_Hypotheses")
hs.sheet_view.showGridLines = False
hs.column_dimensions["A"].width = 52
hs.column_dimensions["B"].width = 9
for col in ("C", "D", "E"):
    hs.column_dimensions[col].width = 14
hs.column_dimensions["F"].width = 13
hs.column_dimensions["G"].width = 62
hs.freeze_panes = "C6"

hs.cell(row=1, column=1, value="HYPOTHESES - 3 SCENARIOS").font = F_TITLE
hs.cell(row=2, column=1, value="Cellules jaunes = saisie. Montants en XPF HT.").font = F_SMALL
hs.cell(row=3, column=1,
        value="Aucune hypothèse de vente de volume SMS au particulier : le SMS est déjà illimité dans les "
              "Forfaits M et les Recharges Packagées Helia.").font = F_SMALL

HDR = 5
for c, txt in enumerate(["Hypothèse", "Unité", "Prudent", "Central", "Ambitieux",
                         "Confiance", "Source / justification"], start=1):
    cell = hs.cell(row=HDR, column=c, value=txt)
    cell.fill, cell.font, cell.border = FILL_H1, F_H1, B_ALL
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
hs.row_dimensions[HDR].height = 26

H = {}   # cle -> numero de ligne
row = HDR + 1

def hsec(title):
    global row
    band(hs, row, 7, title)
    row += 1

def hyp(key, label, unit, p, c, a, conf, src, fmt=FMT_NUM):
    global row
    hs.cell(row=row, column=1, value=label).font = F_BASE
    hs.cell(row=row, column=1).alignment = Alignment(wrap_text=True, vertical="center")
    hs.cell(row=row, column=2, value=unit).font = F_SMALL
    for col, val in (("C", p), ("D", c), ("E", a)):
        cell = hs[f"{col}{row}"]
        cell.value = val
        cell.fill = FILL_IN
        cell.font = F_BASE
        cell.border = B_ALL
        cell.number_format = fmt
        cell.alignment = Alignment(horizontal="right")
    cf = hs.cell(row=row, column=6, value=conf)
    cf.font = F_SMALL
    cf.alignment = Alignment(horizontal="center")
    if conf.startswith("A VALIDER"):
        cf.fill = FILL_AMB
    elif conf == "Elevée":
        cf.fill = FILL_GRN
    s = hs.cell(row=row, column=7, value=src)
    s.font = F_SMALL
    s.alignment = Alignment(wrap_text=True, vertical="center")
    H[key] = row
    row += 1

hsec("A. AUDIENCE ET VOLUMETRIE DU SERVICE")
hyp("mau0", "Utilisateurs actifs mensuels (MAU) au mois 1", "pers.", 5000, 8000, 12000,
    "A VALIDER OPT", "Volumétrie actuelle du service non publiée. Fourchette déduite de la population "
    "de Nouvelle-Calédonie (~268 000 hab.) et du parc mobile Mobilis/Liberté. Question posée à l'OPT-NC.")
hyp("growth", "Croissance mensuelle des MAU après refonte", "%/mois", 0.015, 0.025, 0.035,
    "Moyenne", "Effet refonte UX + accessibilité mobile et bas débit (F02, F08). Prudent = simple maintien "
    "de la tendance, ambitieux = relance par la notoriété du canal 1000.", FMT_PCT)
hyp("maucap", "Plafond de saturation des MAU", "pers.", 10000, 18000, 22000,
    "Moyenne", "Borne haute réaliste au regard de la taille du marché calédonien. Empêche toute projection "
    "de croissance non bornée.")
hyp("smspm", "SMS gratuits envoyés par MAU et par mois", "SMS", 3.0, 3.5, 4.0,
    "A VALIDER OPT", "Usage occasionnel par nature. Sert au calcul du coût d'acheminement, jamais à "
    "une facturation.", FMT_DEC)

hsec("B. APPORT D'AFFAIRES B2B VERS HELIA PRO  (exigence F10 monetisée)")
hyp("prodet", "Part des MAU détectés en usage professionnel", "%/mois", 0.004, 0.007, 0.010,
    "Moyenne", "Détection par faisceau d'indices (récurrence, signature commerciale, gabarit de message). "
    "Le dossier impose déjà cette distinction en F10.", FMT_PCT2)
hyp("proclic", "Taux de prise de contact vers Helia PRO", "%", 0.25, 0.35, 0.45,
    "Moyenne", "Orientation contextuelle au moment précis du besoin non satisfait : taux supérieur à une "
    "sollicitation publicitaire classique.", FMT_PCT)
hyp("proconv", "Taux de conversion en client Helia PRO", "%", 0.04, 0.06, 0.09,
    "A VALIDER OPT", "Taux de transformation d'un lead entrant qualifié. A confronter au tunnel commercial "
    "réel d'Helia PRO.", FMT_PCT)
hyp("promarge", "Marge annuelle moyenne d'un client Helia PRO", "XPF/an", 90000, 130000, 190000,
    "A VALIDER OPT", "MARGE et non chiffre d'affaires. Ordre de grandeur déduit du positionnement de la "
    "plateforme d'envoi en masse (60 000 F HT/mois, 10 000 SMS inclus) et des petits comptes.", FMT_XPF)
hyp("proattr", "Taux d'attribution de la marge à Mobitag NG", "%", 0.30, 0.40, 0.45,
    "Moyenne", "Attribution volontairement partielle : l'effort commercial Helia PRO conserve la part "
    "majoritaire ou paritaire de la valeur. Evite toute surestimation.", FMT_PCT)
hyp("prochurn", "Attrition mensuelle des comptes PRO apportés", "%/mois", 0.020, 0.013, 0.009,
    "Moyenne", "Equivaut à une rétention annuelle de 78 % / 85 % / 90 %.", FMT_PCT2)
hyp("prostart", "Mois de démarrage de l'orientation PRO", "mois", 10, 8, 7,
    "Elevée", "Démarrage après bascule du service refondu et mise en place de la convention d'apport.")

hsec("C. APPORT D'AFFAIRES B2C - ACTIVATION DE LIGNES  (audience sans ligne active)")
hyp("nolinepct", "Part des MAU sans ligne Helia active ou hors forfait", "%", 0.12, 0.18, 0.25,
    "A VALIDER OPT", "C'est l'hypothèse la plus structurante du modèle et la question prioritaire posée à "
    "l'OPT-NC. Usages visés : mobile perdu, volé, cassé, déchargé, hors forfait, poste partagé "
    "(mairie, médiathèque), personne sans mobile.", FMT_PCT)
hyp("b2cclic", "Taux de prise de contact vers une offre Helia identifiée", "%/mois", 0.08, 0.12, 0.15,
    "Moyenne", "Encart clairement identifié, jamais inséré dans un message personnel, jamais dissimulé. "
    "Conforme à la section 15 du dossier.", FMT_PCT)
hyp("b2cconv", "Taux d'activation ou de souscription", "%", 0.04, 0.06, 0.07,
    "A VALIDER OPT", "Population en situation de besoin réel et immédiat de rétablir sa ligne.", FMT_PCT)
hyp("b2cmarge", "Marge annuelle moyenne d'une ligne grand public", "XPF/an", 18000, 26000, 36000,
    "A VALIDER OPT", "MARGE et non chiffre d'affaires. Ordre de grandeur déduit des Forfaits M "
    "(1 000 à 10 000 F TTC/mois) et du prépayé, pondéré vers le bas du catalogue.", FMT_XPF)
hyp("b2cattr", "Taux d'attribution de la marge à Mobitag NG", "%", 0.25, 0.35, 0.40,
    "Moyenne", "Attribution minoritaire : la souscription reste portée par le réseau de distribution Helia.", FMT_PCT)
hyp("b2cchurn", "Attrition mensuelle des lignes apportées", "%/mois", 0.024, 0.018, 0.013,
    "Moyenne", "Equivaut à une rétention annuelle de 75 % / 80 % / 85 %.", FMT_PCT2)
hyp("b2cstart", "Mois de démarrage de l'orientation B2C", "mois", 10, 8, 7,
    "Elevée", "Post-bascule, après validation du parcours par le DPO et la direction marketing.")

hsec("D. SPONSORING CONSENTI  (ligne d'appoint, nulle en scénario prudent)")
hyp("sponsor", "Revenu mensuel en régime établi", "XPF/mois", 0, 120000, 280000,
    "Moyenne", "Inventaire étroit sur un marché de 268 000 habitants. Nul en prudent : la soutenabilité du "
    "modèle ne dépend jamais de cette ligne.", FMT_XPF)
hyp("sponstart", "Mois de démarrage du sponsoring", "mois", 25, 14, 11,
    "Moyenne", "25 = jamais activé sur l'horizon en scénario prudent.")
hyp("sponramp", "Durée de montée en charge", "mois", 6, 6, 5,
    "Moyenne", "Montée linéaire jusqu'au régime établi.")
hyp("sponrev", "Reversement régie / partenaire", "% du CA", 0.15, 0.15, 0.15,
    "Moyenne", "Commission d'une régie locale. Entre dans les « reversements éventuels » de la formule "
    "de marge contributive.", FMT_PCT)

hsec("E. COUTS VARIABLES")
hyp("costsms", "Coût interne d'acheminement d'un SMS", "XPF/SMS", 2.5, 2.0, 1.5,
    "A VALIDER OPT", "Coût interne réseau, hors tarif commercial. Décroît avec la mutualisation de la "
    "passerelle.", FMT_DEC)
hyp("costmau", "Coût d'infrastructure variable par MAU", "XPF/mois", 4, 3, 2,
    "Elevée", "Calculé sur tarifs publics de calcul, stockage et sortie réseau pour une charge "
    "conteneurisée sans état.", FMT_DEC)

hsec("F. COUTS FIXES DU SERVICE NOUVELLE GENERATION")
hyp("etp", "Run technique et produit", "ETP", 0.28, 0.25, 0.20,
    "Moyenne", "Exploitation d'une stack conteneurisée sans état, avec CI/CD, déploiements sans interruption et supervision automatisée. Ce run ne démarre qu'à la bascule : avant, la charge de construction est portée par l'investissement.", FMT_DEC)
hyp("etpcost", "Coût annuel chargé d'un ETP", "XPF/an", 8500000, 8000000, 7500000,
    "Elevée", "Coût chargé d'un profil technique confirmé en Nouvelle-Calédonie.", FMT_XPF)
hyp("cloud", "Hébergement cloud", "XPF/mois", 70000, 55000, 45000,
    "Elevée", "Socle conteneurisé, base de données managée, sauvegardes, observabilité, environnements "
    "de test.", FMT_XPF)
hyp("support", "Support et modération résiduels", "XPF/mois", 30000, 22000, 16000,
    "Moyenne", "Résiduel après automatisation de l'anti-abus et du signalement (F05, F11).", FMT_XPF)
hyp("comm", "Communication et notoriété", "XPF/mois", 25000, 40000, 60000,
    "Moyenne", "Poste nouveau, absent du service historique. Soutient la croissance des MAU.", FMT_XPF)

hyp("devenv", "Environnements de dev, test et pré-production (avant bascule)", "XPF/mois",
    20000, 25000, 30000,
    "Elevée", "Seul coût fixe courant avant la bascule. La charge de construction est portée par l'investissement, jamais par le run : le modèle ne compte pas deux fois la même équipe.", FMT_XPF)

hsec("G. INVESTISSEMENTS")
hyp("capex", "Build, industrialisation et reprise", "XPF", 5500000, 4800000, 4200000,
    "Elevée", "Conception, développement, sécurité, accessibilité, reprise et recette. Chiffré à partir "
    "de la charge estimée de l'équipe.", FMT_XPF)
hyp("capexm", "Etalement de l'investissement", "mois", 9, 7, 6,
    "Elevée", "Etalement linéaire à compter du mois 1.")

hsec("H. FISCALITE ET BASCULE")
hyp("is", "Impôt sur les sociétés (Nouvelle-Calédonie)", "%", 0.30, 0.30, 0.30,
    "Elevée", "Taux normal de l'IS en Nouvelle-Calédonie. Report des déficits antérieurs appliqué dans "
    "le calcul mensuel.", FMT_PCT)
hyp("tgc", "TGC applicable aux services (mémo)", "%", 0.11, 0.11, 0.11,
    "Elevée", "Taux normal de la Taxe Générale sur la Consommation. Neutre au compte de résultat "
    "(collectée puis reversée) : présentée pour mémoire.", FMT_PCT)
hyp("switch", "Mois de bascule vers le service refondu", "mois", 11, 10, 9,
    "Elevée", "Fin du sprint POC, pilote, puis bascule. Déclenche la réalisation des coûts évités.")
hyp("switchramp", "Durée de réalisation des économies", "mois", 4, 3, 2,
    "Moyenne", "Montée linéaire : période de double run entre l'ancien et le nouveau socle.")

HROW_END = row


# ===========================================================================
# 2 - COUTS EVITES
# ===========================================================================
cs = wb.create_sheet("2_Couts_evites")
cs.sheet_view.showGridLines = False
cs.column_dimensions["A"].width = 46
cs.column_dimensions["B"].width = 16
for col in "CDE":
    cs.column_dimensions[col].width = 13
cs.column_dimensions["F"].width = 13
cs.column_dimensions["G"].width = 62

cs.cell(row=1, column=1, value="COUT COMPLET ACTUEL DU SERVICE ET TAUX D'EVITEMENT").font = F_TITLE
cs.cell(row=2, column=1,
        value="Le coût évité est le contributeur principal à l'équilibre. Il porte sur des charges internes "
              "OPT-NC : aucun risque de cannibalisation d'une offre.").font = F_SMALL
cs.cell(row=3, column=1,
        value="Chaque poste porte son propre taux d'évitement. Cellules jaunes = saisie.").font = F_SMALL

CH = 5
for c, txt in enumerate(["Poste de coût du service historique", "Coût annuel actuel (XPF)",
                         "Evité Prudent", "Evité Central", "Evité Ambitieux",
                         "Confiance", "Justification du taux d'évitement"], start=1):
    cell = cs.cell(row=CH, column=c, value=txt)
    cell.fill, cell.font, cell.border = FILL_H1, F_H1, B_ALL
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
cs.row_dimensions[CH].height = 30

CE_POSTS = [
    ("Hébergement et infrastructure on-premise allouée", 1800000, 0.75, 0.85, 0.95,
     "A VALIDER OPT",
     "Bascule vers un socle conteneurisé mutualisé. Quote-part serveurs, stockage, sauvegarde et "
     "supervision du service historique."),
    ("TMA et maintenance corrective", 3600000, 0.72, 0.80, 0.90,
     "A VALIDER OPT",
     "Poste le plus lourd du legacy. Remplaçé par une stack open source maintenue en continu, "
     "avec tests automatisés et déploiement reproductible."),
    ("Exploitation interne (0,25 ETP)", 2000000, 0.60, 0.75, 0.85,
     "A VALIDER OPT",
     "Supervision, redémarrages, gestion des incidents et des évolutions mineures. Automatisée par "
     "l'observabilité et les déploiements sans interruption."),
    ("Support N1/N2 : abus, réclamations, signalements", 900000, 0.60, 0.75, 0.90,
     "Moyenne",
     "Directement réduit par le dispositif de confiance sur le canal 1000 : vérification d'origine "
     "d'un Mobitag, signalement en un clic, maîtrise par le destinataire. C'est la contribution "
     "économique de la ligne F11."),
    ("Licences et composants propriétaires", 450000, 0.85, 0.95, 1.00,
     "Moyenne",
     "Supprimées par l'exigence open source du challenge. Evitement quasi total par construction."),
    ("Acheminement des SMS du service historique", 700000, 1.00, 1.00, 1.00,
     "A VALIDER OPT",
     "Charge reprise A L'IDENTIQUE par le service refondu, où elle figure en coût variable. "
     "Neutre au différentiel : présentée des deux côtés pour que le comparatif soit auditable et "
     "qu'aucun coût du socle gratuit ne soit dissimulé."),
]

r = CH + 1
CE_FIRST = r
for label, cost, p, c, a, conf, src in CE_POSTS:
    cs.cell(row=r, column=1, value=label).font = F_BASE
    cs.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="center")
    cell = cs.cell(row=r, column=2, value=cost)
    cell.fill, cell.font, cell.border, cell.number_format = FILL_IN, F_BASE, B_ALL, FMT_XPF
    for col, val in (("C", p), ("D", c), ("E", a)):
        cc = cs[f"{col}{r}"]
        cc.value, cc.fill, cc.font, cc.border, cc.number_format = val, FILL_IN, F_BASE, B_ALL, FMT_PCT
    cf = cs.cell(row=r, column=6, value=conf)
    cf.font, cf.alignment = F_SMALL, Alignment(horizontal="center")
    if conf.startswith("A VALIDER"):
        cf.fill = FILL_AMB
    s = cs.cell(row=r, column=7, value=src)
    s.font, s.alignment = F_SMALL, Alignment(wrap_text=True, vertical="center")
    cs.row_dimensions[r].height = 30
    r += 1
CE_LAST = r - 1

# totaux
cs.cell(row=r, column=1, value="TOTAL COUT COMPLET ANNUEL ACTUEL").font = F_TOT
cs.cell(row=r, column=2, value=f"=SUM(B{CE_FIRST}:B{CE_LAST})").font = F_TOT
cs.cell(row=r, column=2).number_format = FMT_XPF
for c in range(1, 8):
    cs.cell(row=r, column=c).fill = FILL_TOT
    cs.cell(row=r, column=c).border = B_ALL
CE_TOTAL = r
r += 1

cs.cell(row=r, column=1, value="ECONOMIE ANNUELLE RETENUE (régime établi)").font = F_TOT
for col in "CDE":
    cell = cs[f"{col}{r}"]
    cell.value = f"=SUMPRODUCT($B${CE_FIRST}:$B${CE_LAST},{col}{CE_FIRST}:{col}{CE_LAST})"
    cell.font, cell.number_format = F_TOT, FMT_XPF
for c in range(1, 8):
    cs.cell(row=r, column=c).fill = FILL_GRN
    cs.cell(row=r, column=c).border = B_ALL
CE_SAVE_Y = r
r += 1

cs.cell(row=r, column=1, value="Economie mensuelle en régime établi").font = F_BOLD
for col in "CDE":
    cell = cs[f"{col}{r}"]
    cell.value = f"={col}{CE_SAVE_Y}/12"
    cell.font, cell.number_format = F_BOLD, FMT_XPF
CE_SAVE_M = r
r += 1

cs.cell(row=r, column=1, value="Taux d'évitement global").font = F_BASE
for col in "CDE":
    cell = cs[f"{col}{r}"]
    cell.value = f"={col}{CE_SAVE_Y}/$B${CE_TOTAL}"
    cell.font, cell.number_format = F_BASE, FMT_PCT
r += 2

for note in [
    "Absence de double comptage : les coûts évités ci-dessus correspondent à l'ARRET du service historique.",
    "Les coûts d'exploitation du service refondu (run, cloud, support, communication) sont portés "
    "séparément en coûts fixes dans les onglets 4 à 6. Le modèle ne compte donc jamais deux fois la même charge.",
    "La réalisation de ces économies débute au mois de bascule (hypothèse 'switch') et monte linéairement "
    "sur la durée de double run (hypothèse 'switchramp').",
]:
    cs.cell(row=r, column=1, value=note).font = F_SMALL
    cs.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    cs.row_dimensions[r].height = 24
    cs.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="center")
    r += 1


# ===========================================================================
# 4/5/6 - COMPTES DE RESULTAT
# ===========================================================================
PL_ROWS = {}

def build_pl(name, col):
    """Construit un onglet P&L 24 mois pour la colonne de scenario donnee."""
    ps = wb.create_sheet(name)
    ps.sheet_view.showGridLines = False
    ps.column_dimensions["A"].width = 46
    ps.column_dimensions["B"].width = 11
    for m in range(1, MONTHS + 1):
        ps.column_dimensions[CL(m)].width = 11
    ps.freeze_panes = "C6"

    scen_name = name.split("_")[-1]
    ps.cell(row=1, column=1, value=f"COMPTE DE RESULTAT MENSUEL 24 MOIS - SCENARIO {scen_name.upper()}").font = F_TITLE
    ps.cell(row=2, column=1, value="Montants en XPF HT. Vue service Mobitag Nouvelle Génération (périmètre OPT-NC).").font = F_SMALL
    ps.cell(row=3, column=1,
            value="Aucune ligne de revenu issue de la vente de volume SMS au particulier.").font = F_SMALL

    HDRR = 5
    ps.cell(row=HDRR, column=1, value="Poste").fill = FILL_H1
    ps.cell(row=HDRR, column=1).font = F_H1
    ps.cell(row=HDRR, column=2, value="Total 24M").fill = FILL_H1
    ps.cell(row=HDRR, column=2).font = F_H1
    ps.cell(row=HDRR, column=2).alignment = Alignment(horizontal="center")
    for m in range(1, MONTHS + 1):
        cell = ps.cell(row=HDRR, column=C0 + m - 1, value=f"M{m}")
        cell.fill, cell.font, cell.border = FILL_H1, F_H1, B_ALL
        cell.alignment = Alignment(horizontal="center")
    ps.cell(row=HDRR, column=1).border = B_ALL
    ps.cell(row=HDRR, column=2).border = B_ALL

    R = {}
    rr = HDRR + 1
    LAST_COL = C0 + MONTHS - 1

    def sec(title):
        nonlocal rr
        band(ps, rr, LAST_COL, title)
        rr += 1

    def line(key, label, formula_fn, fmt=FMT_XPF, total=False, tot_kind="sum",
             bold=False, fill=None):
        nonlocal rr
        R[key] = rr          # enregistre avant : certaines lignes s'auto-referencent (M-1)
        lc = ps.cell(row=rr, column=1, value=label)
        lc.font = F_TOT if (total or bold) else F_BASE
        lc.alignment = Alignment(wrap_text=False, vertical="center")
        for m in range(1, MONTHS + 1):
            cell = ps.cell(row=rr, column=C0 + m - 1, value=formula_fn(m))
            cell.number_format = fmt
            cell.font = F_TOT if (total or bold) else F_BASE
            cell.border = B_ALL
            if fill:
                cell.fill = fill
        # colonne total
        tc = ps.cell(row=rr, column=2)
        if tot_kind == "sum":
            tc.value = f"=SUM(C{rr}:{LASTC}{rr})"
        elif tot_kind == "last":
            tc.value = f"={LASTC}{rr}"
        elif tot_kind == "avg":
            tc.value = f"=AVERAGE(C{rr}:{LASTC}{rr})"
        else:
            tc.value = None
        tc.number_format = fmt
        tc.font = F_TOT
        tc.border = B_ALL
        tc.fill = fill or FILL_TOT
        if total:
            for c in range(1, LAST_COL + 1):
                if not ps.cell(row=rr, column=c).fill or ps.cell(row=rr, column=c).fill.fgColor.rgb in (None, "00000000"):
                    ps.cell(row=rr, column=c).fill = FILL_TOT
        rr += 1

    def h(key):
        return f"{HYP}!${col}${H[key]}"

    # ---- VOLUMETRIE
    sec("VOLUMETRIE")
    line("mau", "Utilisateurs actifs mensuels (MAU)",
         lambda m: f"={h('mau0')}" if m == 1
         else f"=MIN({h('maucap')},{CL(m-1)}{R['mau']}*(1+{h('growth')}))",
         FMT_NUM, tot_kind="last")
    line("sms", "SMS gratuits envoyés",
         lambda m: f"={CL(m)}{R['mau']}*{h('smspm')}", FMT_NUM)
    line("detpro", "Usages professionnels détectés",
         lambda m: f"={CL(m)}{R['mau']}*{h('prodet')}", FMT_DEC)
    line("detb2c", "Utilisateurs sans ligne active détectés",
         lambda m: f"={CL(m)}{R['mau']}*{h('nolinepct')}", FMT_NUM)

    # ---- APPORTS (stocks)
    sec("APPORTS D'AFFAIRES - CONSTITUTION DES STOCKS")
    line("convpro", "Nouveaux clients Helia PRO apportés",
         lambda m: f"=IF({m}>={h('prostart')},{CL(m)}{R['detpro']}*{h('proclic')}*{h('proconv')},0)",
         FMT_DEC)
    line("stockpro", "Stock de clients Helia PRO apportés",
         lambda m: f"={CL(m)}{R['convpro']}" if m == 1
         else f"={CL(m-1)}{R['stockpro']}*(1-{h('prochurn')})+{CL(m)}{R['convpro']}",
         FMT_DEC, tot_kind="last")
    line("convb2c", "Nouvelles lignes grand public activées",
         lambda m: f"=IF({m}>={h('b2cstart')},{CL(m)}{R['detb2c']}*{h('b2cclic')}*{h('b2cconv')},0)",
         FMT_DEC)
    line("stockb2c", "Stock de lignes grand public apportées",
         lambda m: f"={CL(m)}{R['convb2c']}" if m == 1
         else f"={CL(m-1)}{R['stockb2c']}*(1-{h('b2cchurn')})+{CL(m)}{R['convb2c']}",
         FMT_DEC, tot_kind="last")

    # ---- CA
    sec("CHIFFRE D'AFFAIRES / VALEUR CREEE")
    line("capro", "Apport d'affaires B2B - Helia PRO",
         lambda m: f"={CL(m)}{R['stockpro']}*{h('promarge')}/12*{h('proattr')}")
    line("cab2c", "Apport d'affaires B2C - activation de lignes",
         lambda m: f"={CL(m)}{R['stockb2c']}*{h('b2cmarge')}/12*{h('b2cattr')}")
    line("caspon", "Sponsoring consenti",
         lambda m: f"=IF({m}<{h('sponstart')},0,{h('sponsor')}*MIN(1,({m}-{h('sponstart')}+1)/{h('sponramp')}))")
    line("ca", "CHIFFRE D'AFFAIRES TOTAL",
         lambda m: f"=SUM({CL(m)}{R['capro']}:{CL(m)}{R['caspon']})",
         total=True)

    # ---- COUTS VARIABLES
    sec("COUTS VARIABLES ET REVERSEMENTS")
    line("cvsms", "Acheminement des SMS gratuits",
         lambda m: f"=-{CL(m)}{R['sms']}*{h('costsms')}")
    line("cvinfra", "Infrastructure variable",
         lambda m: f"=-{CL(m)}{R['mau']}*{h('costmau')}")
    line("cvrev", "Reversement régie sponsoring",
         lambda m: f"=-{CL(m)}{R['caspon']}*{h('sponrev')}")
    line("cv", "TOTAL COUTS VARIABLES ET REVERSEMENTS",
         lambda m: f"=SUM({CL(m)}{R['cvsms']}:{CL(m)}{R['cvrev']})",
         total=True)
    line("mc", "MARGE CONTRIBUTIVE",
         lambda m: f"={CL(m)}{R['ca']}+{CL(m)}{R['cv']}",
         total=True, fill=FILL_GRN)

    # ---- COUTS EVITES
    sec("COUTS EVITES  (arrêt du service historique)")
    line("ceramp", "Taux de réalisation des économies",
         lambda m: f"=IF({m}<{h('switch')},0,MIN(1,({m}-{h('switch')}+1)/{h('switchramp')}))",
         FMT_PCT, tot_kind="last")
    for i, (lbl, _, _, _, _, _, _) in enumerate(CE_POSTS):
        srow = CE_FIRST + i
        line(f"ce{i}", f"  {lbl}",
             lambda m, srow=srow: f"={CEV}!$B${srow}*{CEV}!${col}${srow}/12*{CL(m)}{R['ceramp']}")
    line("ce", "TOTAL COUTS EVITES",
         lambda m: f"=SUM({CL(m)}{R['ce0']}:{CL(m)}{R['ce' + str(len(CE_POSTS)-1)]})",
         total=True, fill=FILL_GRN)

    # ---- COUTS FIXES
    sec("COUTS FIXES DU SERVICE NOUVELLE GENERATION")
    line("cfetp", "Run technique et produit (à compter de la bascule)",
         lambda m: f"=IF({m}>={h('switch')},-{h('etp')}*{h('etpcost')}/12,0)")
    line("cfcloud", "Hébergement (production après bascule, environnements avant)",
         lambda m: f"=IF({m}>={h('switch')},-{h('cloud')},-{h('devenv')})")
    line("cfsup", "Support et modération",
         lambda m: f"=IF({m}>={h('switch')},-{h('support')},0)")
    line("cfcom", "Communication et notoriété",
         lambda m: f"=IF({m}>={h('switch')},-{h('comm')},0)")
    line("cf", "TOTAL COUTS FIXES",
         lambda m: f"=SUM({CL(m)}{R['cfetp']}:{CL(m)}{R['cfcom']})",
         total=True)

    # ---- INVESTISSEMENTS
    sec("INVESTISSEMENTS")
    line("capex", "Build, industrialisation et reprise",
         lambda m: f"=IF({m}<={h('capexm')},-{h('capex')}/{h('capexm')},0)",
         total=True)

    # ---- RESULTAT
    sec("RESULTAT ECONOMIQUE")
    line("rai", "Résultat avant impôt",
         lambda m: f"={CL(m)}{R['mc']}+{CL(m)}{R['ce']}+{CL(m)}{R['cf']}+{CL(m)}{R['capex']}",
         total=True)
    line("raicum", "Résultat avant impôt cumulé",
         lambda m: f"={CL(m)}{R['rai']}" if m == 1
         else f"={CL(m-1)}{R['raicum']}+{CL(m)}{R['rai']}",
         tot_kind="last")
    line("is", "Impôt sur les sociétés (report des déficits appliqué)",
         lambda m: (f"=-IF(AND({CL(m)}{R['rai']}>0,{CL(m)}{R['raicum']}>0),"
                    f"{h('is')}*MIN({CL(m)}{R['rai']},{CL(m)}{R['raicum']}),0)"))
    line("rn", "RESULTAT NET MENSUEL",
         lambda m: f"={CL(m)}{R['rai']}+{CL(m)}{R['is']}",
         total=True, fill=FILL_GRN)
    line("rncum", "EQUILIBRE ECONOMIQUE CUMULE",
         lambda m: f"={CL(m)}{R['rn']}" if m == 1
         else f"={CL(m-1)}{R['rncum']}+{CL(m)}{R['rn']}",
         total=True, tot_kind="last", fill=FILL_GRN)
    line("bfin", "Besoin de financement (point bas du cumul)",
         lambda m: f"=MIN(0,MIN($C{R['rncum']}:{CL(m)}{R['rncum']}))",
         tot_kind="last", fill=FILL_AMB)

    # memo TGC
    rr += 1
    ps.cell(row=rr, column=1, value="Mémo TGC (neutre au résultat)").font = F_SMALL
    for m in range(1, MONTHS + 1):
        cell = ps.cell(row=rr, column=C0 + m - 1,
                       value=f"={CL(m)}{R['ca']}*{h('tgc')}")
        cell.number_format, cell.font = FMT_XPF, F_SMALL
    rr += 2
    for note in [
        "Marge contributive = chiffre d'affaires − coûts variables − reversements  (formule du dossier, section 6).",
        "Equilibre économique = marge contributive + coûts évités − coûts fixes − investissements − fiscalité  (idem).",
        "Les charges sont saisies en négatif : toutes les lignes s'additionnent, ce qui rend chaque total auditable.",
        "La TGC est collectée puis reversée : elle n'affecte pas le résultat et figure pour mémoire.",
    ]:
        ps.cell(row=rr, column=1, value=note).font = F_SMALL
        rr += 1

    PL_ROWS[name] = R
    return R


for scen_name, scen_col, sheet_name in SCEN:
    build_pl(sheet_name, scen_col)


# ===========================================================================
# 3 - UNIT ECONOMICS
# ===========================================================================
us = wb.create_sheet("3_Unit_economics")
us.sheet_view.showGridLines = False
us.column_dimensions["A"].width = 52
us.column_dimensions["B"].width = 12
for c in "CDE":
    us.column_dimensions[c].width = 15
us.column_dimensions["F"].width = 70

us.cell(row=1, column=1, value="UNIT ECONOMICS ET RATIOS DE CONTROLE").font = F_TITLE
us.cell(row=2, column=1, value="Valeurs calculées au mois 24. Montants en XPF HT.").font = F_SMALL

UH = 4
for c, t in enumerate(["Indicateur", "Unité", "Prudent", "Central", "Ambitieux", "Lecture"], start=1):
    cell = us.cell(row=UH, column=c, value=t)
    cell.fill, cell.font, cell.border = FILL_H1, F_H1, B_ALL
    cell.alignment = Alignment(horizontal="center", wrap_text=True)

r = UH + 1
def uline(label, unit, fns, fmt, note, bold=False):
    global r
    us.cell(row=r, column=1, value=label).font = F_BOLD if bold else F_BASE
    us.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="center")
    us.cell(row=r, column=2, value=unit).font = F_SMALL
    for col, (scen_name, scen_col, sheet) in zip("CDE", SCEN):
        cell = us[f"{col}{r}"]
        v = fns(sheet, scen_col)
        cell.value = v if str(v).startswith("=") else "=" + str(v)
        cell.number_format, cell.border = fmt, B_ALL
        cell.font = F_TOT if bold else F_BASE
    n = us.cell(row=r, column=6, value=note)
    n.font, n.alignment = F_SMALL, Alignment(wrap_text=True, vertical="center")
    us.row_dimensions[r].height = 28
    r += 1

def pr(sheet, key):
    return f"'{sheet}'!{LASTC}{PL_ROWS[sheet][key]}"
def prsum(sheet, key):
    return f"'{sheet}'!B{PL_ROWS[sheet][key]}"

band(us, r, 6, "COUT DE REVIENT DU SOCLE GRATUIT")
r += 1
uline("Coût variable d'un SMS gratuit envoyé", "XPF/SMS",
      lambda s, c: f"={HYP}!${c}${H['costsms']}", FMT_DEC,
      "Coût marginal d'acheminement. Reste à la charge du service : le socle gratuit n'est jamais facturé.")
uline("Coût complet mensuel du socle gratuit par MAU", "XPF/MAU",
      lambda s, c: f"=-({pr(s,'cv')}+{pr(s,'cf')})/{pr(s,'mau')}", FMT_DEC,
      "Coûts variables et fixes rapportés à l'utilisateur actif au mois 24. Mesure la sobriété du service refondu.")
uline("Coût complet annuel du socle gratuit", "XPF/an",
      lambda s, c: f"=-({pr(s,'cv')}+{pr(s,'cf')})*12", FMT_XPF,
      "A comparer au coût complet actuel du service, onglet 2.", bold=True)

band(us, r, 6, "VALEUR DES APPORTS D'AFFAIRES")
r += 1
uline("Valeur annuelle attribuée d'un client Helia PRO apporté", "XPF/an",
      lambda s, c: f"={HYP}!${c}${H['promarge']}*{HYP}!${c}${H['proattr']}", FMT_XPF,
      "Marge annuelle moyenne multipliée par le taux d'attribution. Attribution volontairement partielle.")
uline("Durée de vie moyenne d'un compte PRO apporté", "mois",
      lambda s, c: f"=1/{HYP}!${c}${H['prochurn']}", FMT_DEC,
      "Inverse de l'attrition mensuelle.")
uline("Valeur vie attribuée d'un client Helia PRO", "XPF",
      lambda s, c: (f"={HYP}!${c}${H['promarge']}/12*{HYP}!${c}${H['proattr']}"
                    f"/{HYP}!${c}${H['prochurn']}"), FMT_XPF,
      "LTV attribuée au service Mobitag NG.", bold=True)
uline("Valeur vie attribuée d'une ligne grand public", "XPF",
      lambda s, c: (f"={HYP}!${c}${H['b2cmarge']}/12*{HYP}!${c}${H['b2cattr']}"
                    f"/{HYP}!${c}${H['b2cchurn']}"), FMT_XPF,
      "Idem côté activation de lignes.", bold=True)
uline("Coût d'acquisition d'un apport (CAC proxy)", "XPF",
      lambda s, c: (f"=-({HYP}!${c}${H['comm']}*24)"
                    f"/({prsum(s,'convpro')}+{prsum(s,'convb2c')})"), FMT_XPF,
      "Budget communication sur 24 mois rapporté au nombre total d'apports générés.")
uline("Ratio LTV / CAC (apports B2C)", "x",
      lambda s, c: (f"=({HYP}!${c}${H['b2cmarge']}/12*{HYP}!${c}${H['b2cattr']}"
                    f"/{HYP}!${c}${H['b2cchurn']})"
                    f"/(-({HYP}!${c}${H['comm']}*24)"
                    f"/({prsum(s,'convpro')}+{prsum(s,'convb2c')}))"), FMT_DEC,
      "Au-dessus de 3, l'acquisition est saine. Sert de garde-fou au budget communication.", bold=True)

band(us, r, 6, "STRUCTURE ECONOMIQUE AU MOIS 24")
r += 1
uline("Part des coûts évités dans l'équilibre", "%",
      lambda s, c: f"={pr(s,'ce')}/({pr(s,'ce')}+{pr(s,'mc')})", FMT_PCT,
      "Mesure la dépendance du modèle au levier économies. Volontairement majoritaire : c'est le levier "
      "le plus certain et le seul totalement non redondant.", bold=True)
uline("Part des apports d'affaires dans l'équilibre", "%",
      lambda s, c: f"={pr(s,'mc')}/({pr(s,'ce')}+{pr(s,'mc')})", FMT_PCT,
      "Contribution des lignes B2B et B2C, sponsoring inclus.")
uline("Taux de couverture des coûts fixes", "%",
      lambda s, c: f"=({pr(s,'mc')}+{pr(s,'ce')})/-{pr(s,'cf')}", FMT_PCT,
      "Au-dessus de 100 %, le service s'autofinance hors investissement.", bold=True)
uline("Marge contributive annualisée", "XPF/an",
      lambda s, c: f"={pr(s,'mc')}*12", FMT_XPF,
      "Marge contributive du mois 24 annualisée.")


# ===========================================================================
# 7 - SYNTHESE
# ===========================================================================
ss = wb.create_sheet("7_Synthese")
ss.sheet_view.showGridLines = False
ss.column_dimensions["A"].width = 54
for c in "BCDE":
    ss.column_dimensions[c].width = 17
ss.column_dimensions["F"].width = 66

ss.cell(row=1, column=1, value="SYNTHESE DES TROIS SCENARIOS").font = F_TITLE
ss.cell(row=2, column=1, value="Montants en XPF HT, cumuls sur 24 mois sauf mention contraire.").font = F_SMALL

SH = 4
for c, t in enumerate(["Indicateur", "Prudent", "Central", "Ambitieux", "", "Commentaire"], start=1):
    cell = ss.cell(row=SH, column=c, value=t)
    cell.fill, cell.font, cell.border = FILL_H1, F_H1, B_ALL
    cell.alignment = Alignment(horizontal="center")

r = SH + 1
def sline(label, fn, fmt, note, bold=False, fill=None):
    global r
    ss.cell(row=r, column=1, value=label).font = F_TOT if bold else F_BASE
    for col, (sn, sc, sheet) in zip("BCD", SCEN):
        cell = ss[f"{col}{r}"]
        v = fn(sheet, sc)
        cell.value = v if str(v).startswith("=") else "=" + str(v)
        cell.number_format, cell.border = fmt, B_ALL
        cell.font = F_TOT if bold else F_BASE
        if fill:
            cell.fill = fill
    n = ss.cell(row=r, column=6, value=note)
    n.font, n.alignment = F_SMALL, Alignment(wrap_text=True, vertical="center")
    ss.row_dimensions[r].height = 26
    r += 1

band(ss, r, 6, "ACTIVITE AU MOIS 24")
r += 1
sline("Utilisateurs actifs mensuels", lambda s, c: pr(s, "mau"), FMT_NUM,
      "Audience du service refondu au terme de l'horizon.")
sline("SMS gratuits envoyés sur 24 mois", lambda s, c: prsum(s, "sms"), FMT_NUM,
      "Volume cumulé du socle gratuit. Jamais facturé, jamais plafonné par une offre payante.")
sline("Clients Helia PRO apportés (stock M24)", lambda s, c: pr(s, "stockpro"), FMT_DEC,
      "Comptes professionnels orientés vers Helia PRO et encore actifs.")
sline("Lignes grand public activées (stock M24)", lambda s, c: pr(s, "stockb2c"), FMT_DEC,
      "Lignes réactivées ou souscrites depuis l'audience sans ligne active.")

band(ss, r, 6, "COMPTE DE RESULTAT CUMULE 24 MOIS")
r += 1
sline("Chiffre d'affaires", lambda s, c: prsum(s, "ca"), FMT_XPF,
      "Apports d'affaires B2B et B2C, sponsoring. Aucune vente de SMS au particulier.")
sline("Coûts variables et reversements", lambda s, c: prsum(s, "cv"), FMT_XPF,
      "Acheminement des SMS gratuits, infrastructure variable, reversement régie.")
sline("MARGE CONTRIBUTIVE", lambda s, c: prsum(s, "mc"), FMT_XPF,
      "= CA − coûts variables − reversements.", bold=True, fill=FILL_GRN)
sline("Coûts évités", lambda s, c: prsum(s, "ce"), FMT_XPF,
      "Arrêt du socle historique. Pilier du modèle.", bold=True, fill=FILL_GRN)
sline("Coûts fixes", lambda s, c: prsum(s, "cf"), FMT_XPF,
      "Run, cloud, support, communication du service refondu.")
sline("Investissements", lambda s, c: prsum(s, "capex"), FMT_XPF,
      "Build, industrialisation et reprise, étalés.")
sline("Fiscalité (IS)", lambda s, c: prsum(s, "is"), FMT_XPF,
      "IS Nouvelle-Calédonie 30 %, report des déficits appliqué.")
sline("EQUILIBRE ECONOMIQUE 24 MOIS", lambda s, c: pr(s, "rncum"), FMT_XPF,
      "= marge contributive + coûts évités − coûts fixes − investissements − fiscalité.",
      bold=True, fill=FILL_GRN)

band(ss, r, 6, "INDICATEURS DE PILOTAGE")
r += 1
PTMORT = r
sline("Point mort (1er mois à cumul positif)",
      lambda s, c: (f"=IFERROR(MATCH(TRUE,INDEX('{s}'!C{PL_ROWS[s]['rncum']}:{LASTC}{PL_ROWS[s]['rncum']}>0,0),0),"
                    f"\"> 24 mois\")"), '0',
      "Mois à partir duquel l'équilibre cumulé devient positif. « > 24 mois » signale une trajectoire "
      "d'équilibre non encore atteinte sur l'horizon.", bold=True)
sline("Besoin de financement maximal", lambda s, c: pr(s, "bfin"), FMT_XPF,
      "Point bas du cumul. Correspond au financement à mobiliser avant l'autofinancement.",
      bold=True, fill=FILL_AMB)
sline("Résultat net du mois 24 (régime)", lambda s, c: pr(s, "rn"), FMT_XPF,
      "Résultat mensuel une fois l'investissement achevé et les économies réalisées.")
sline("Résultat annualisé en régime établi", lambda s, c: f"={pr(s,'rn')}*12", FMT_XPF,
      "Extrapolation du mois 24 sur douze mois. Sert de base à la discussion pilote / industrialisation.",
      bold=True)

r += 1
ss.cell(row=r, column=1,
        value="LECTURE : le modèle est conçu pour que la soutenabilité repose d'abord sur un levier certain "
              "(les coûts évités) et non sur une hypothèse de marché. Même en scénario prudent, sans aucun "
              "sponsoring et avec les taux de conversion les plus bas, le service réduit le coût complet "
              "supporté par l'OPT-NC.").font = F_BOLD
ss.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
ss.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="center")
ss.row_dimensions[r].height = 42
r += 2

# --- donnees pour graphiques : cumul par scenario
GR = r
ss.cell(row=GR, column=1, value="Données graphiques - équilibre cumulé (XPF)").font = F_H2
GR += 1
ss.cell(row=GR, column=1, value="Mois").font = F_BOLD
for i, (sn, sc, sheet) in enumerate(SCEN):
    ss.cell(row=GR, column=2 + i, value=sn).font = F_BOLD
for m in range(1, MONTHS + 1):
    ss.cell(row=GR + m, column=1, value=f"M{m}").font = F_SMALL
    for i, (sn, sc, sheet) in enumerate(SCEN):
        cell = ss.cell(row=GR + m, column=2 + i,
                       value=f"='{sheet}'!{CL(m)}{PL_ROWS[sheet]['rncum']}")
        cell.number_format = FMT_XPF
        cell.font = F_SMALL

chart = LineChart()
chart.title = "Equilibre économique cumulé sur 24 mois (XPF)"
chart.style = 2
chart.height, chart.width = 9, 22
chart.y_axis.title = "XPF cumulés"
chart.x_axis.title = "Mois"
data = Reference(ss, min_col=2, max_col=4, min_row=GR, max_row=GR + MONTHS)
cats = Reference(ss, min_col=1, min_row=GR + 1, max_row=GR + MONTHS)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
ss.add_chart(chart, f"H{SH}")

# --- graphique structure de l'equilibre (scenario central)
GR2 = GR + MONTHS + 2
ss.cell(row=GR2, column=1, value="Données graphiques - structure de l'équilibre (cumul 24M, scénario central)").font = F_H2
labels = [("Marge contributive", "mc"), ("Coûts évités", "ce"),
          ("Coûts fixes", "cf"), ("Investissements", "capex"), ("Fiscalité", "is")]
ss.cell(row=GR2 + 1, column=1, value="Poste").font = F_BOLD
ss.cell(row=GR2 + 1, column=2, value="Central").font = F_BOLD
for i, (lbl, key) in enumerate(labels):
    ss.cell(row=GR2 + 2 + i, column=1, value=lbl).font = F_SMALL
    cell = ss.cell(row=GR2 + 2 + i, column=2, value=f"='5_PL_Central'!B{PL_ROWS['5_PL_Central'][key]}")
    cell.number_format = FMT_XPF
    cell.font = F_SMALL

bar = BarChart()
bar.type = "col"
bar.title = "Structure de l'équilibre 24 mois - scénario central (XPF)"
bar.height, bar.width = 8, 22
bar.y_axis.title = "XPF cumulés"
bdata = Reference(ss, min_col=2, min_row=GR2 + 1, max_row=GR2 + 1 + len(labels))
bcats = Reference(ss, min_col=1, min_row=GR2 + 2, max_row=GR2 + 1 + len(labels))
bar.add_data(bdata, titles_from_data=True)
bar.set_categories(bcats)
bar.legend = None
ss.add_chart(bar, f"H{SH + 20}")


# ===========================================================================
# 8 - SENSIBILITE
# ===========================================================================
sv = wb.create_sheet("8_Sensibilite")
sv.sheet_view.showGridLines = False
sv.column_dimensions["A"].width = 52
for c in "BCDEF":
    sv.column_dimensions[c].width = 18
sv.column_dimensions["G"].width = 60

sv.cell(row=1, column=1, value="ANALYSE DE SENSIBILITE - SCENARIO CENTRAL").font = F_TITLE
sv.cell(row=2, column=1,
        value="Impact sur l'équilibre économique cumulé à 24 mois d'une variation isolée de chaque variable. "
              "Les valeurs sont recalculées analytiquement à partir des lignes du P&L central.").font = F_SMALL
sv.cell(row=3, column=1,
        value="Mode d'emploi : pour un test grandeur nature, modifier la valeur dans l'onglet 1 et lire "
              "l'onglet 7. Ce tableau donne l'ordre de grandeur et le classement des risques.").font = F_SMALL

VH = 5
for c, t in enumerate(["Variable testée", "Valeur centrale", "Variation testée",
                       "Impact sur l'équilibre 24M", "Criticité", "Maîtrise du risque"], start=1):
    cell = sv.cell(row=VH, column=c, value=t)
    cell.fill, cell.font, cell.border = FILL_H1, F_H1, B_ALL
    cell.alignment = Alignment(horizontal="center", wrap_text=True)
sv.row_dimensions[VH].height = 28

PLC = "'5_PL_Central'"
RC = PL_ROWS["5_PL_Central"]

SENS = [
    ("Taux d'évitement global des coûts (onglet 2)",
     f"={CEV}!$D${CE_SAVE_Y}/{CEV}!$B${CE_TOTAL}", FMT_PCT, "− 10 points",
     f"={PLC}!B{RC['ce']}*(-0.10/({CEV}!$D${CE_SAVE_Y}/{CEV}!$B${CE_TOTAL}))",
     "CRITIQUE",
     "Variable la plus structurante. Maîtrisée par un audit contradictoire du coût complet avec l'OPT-NC "
     "dès le sprint POC, et par un taux déjà prudent poste par poste."),
    ("Part des MAU sans ligne active",
     f"={HYP}!$D${H['nolinepct']}", FMT_PCT, "− 1/3",
     f"=-{PLC}!B{RC['cab2c']}/3", "ELEVEE",
     "Hypothèse prioritaire à valider auprès de l'OPT-NC. Question posée avant la clôture. "
     "Un tiers d'erreur reste absorbable par le pilier coûts évités."),
    ("Nombre d'utilisateurs actifs mensuels au mois 1",
     f"={HYP}!$D${H['mau0']}", FMT_NUM, "− 30 %",
     f"=({PLC}!B{RC['capro']}+{PLC}!B{RC['cab2c']})*-0.30-{PLC}!B{RC['cv']}*-0.30",
     "ELEVEE",
     "Volumétrie non publiée. Impact atténué : une baisse des MAU réduit aussi les coûts variables."),
    ("Taux de conversion en client Helia PRO",
     f"={HYP}!$D${H['proconv']}", FMT_PCT, "− 50 %",
     f"=-{PLC}!B{RC['capro']}/2", "MOYENNE",
     "Ligne volontairement minoritaire dans l'équilibre. Le modèle tient sans elle."),
    ("Marge annuelle d'une ligne grand public",
     f"={HYP}!$D${H['b2cmarge']}", FMT_XPF, "− 25 %",
     f"=-{PLC}!B{RC['cab2c']}/4", "MOYENNE",
     "A confronter aux données internes OPT-NC. Attribution déjà minoritaire (35 %)."),
    ("Revenu de sponsoring",
     f"={HYP}!$D${H['sponsor']}", FMT_XPF, "− 100 %",
     f"=-{PLC}!B{RC['caspon']}*(1-{HYP}!$D${H['sponrev']})", "FAIBLE",
     "Ligne nulle par construction en scénario prudent : sa disparition complète ne remet pas "
     "l'équilibre en cause."),
    ("Investissement de build et d'industrialisation",
     f"={HYP}!$D${H['capex']}", FMT_XPF, "+ 30 %",
     f"={PLC}!B{RC['capex']}*0.30", "MOYENNE",
     "Dépassement classique sur un projet de refonte. Maîtrisé par un périmètre POC resserré et "
     "une industrialisation par lots."),
    ("Décalage du mois de bascule",
     f"={HYP}!$D${H['switch']}", '0', "+ 3 mois",
     f"=-{CEV}!$D${CE_SAVE_M}*3", "ELEVEE",
     "Chaque mois de retard coûte un mois d'économies. Maîtrisé par une bascule progressive "
     "et une reprise sans interruption de service."),
]

r = VH + 1
for label, val_f, val_fmt, var, imp_f, crit, mit in SENS:
    sv.cell(row=r, column=1, value=label).font = F_BASE
    sv.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="center")
    c2 = sv.cell(row=r, column=2, value=val_f)
    c2.number_format, c2.font, c2.border = val_fmt, F_BASE, B_ALL
    c3 = sv.cell(row=r, column=3, value=var)
    c3.font, c3.alignment, c3.border = F_BASE, Alignment(horizontal="center"), B_ALL
    c4 = sv.cell(row=r, column=4, value=imp_f)
    c4.number_format, c4.font, c4.border = FMT_XPF, F_TOT, B_ALL
    c5 = sv.cell(row=r, column=5, value=crit)
    c5.font, c5.alignment, c5.border = F_BOLD, Alignment(horizontal="center"), B_ALL
    c5.fill = {"CRITIQUE": PatternFill("solid", fgColor="F6D5D5"),
               "ELEVEE": FILL_AMB, "MOYENNE": PatternFill("solid", fgColor="FFF6D6"),
               "FAIBLE": FILL_GRN}[crit]
    c6 = sv.cell(row=r, column=6, value=mit)
    c6.font, c6.alignment = F_SMALL, Alignment(wrap_text=True, vertical="center")
    sv.row_dimensions[r].height = 34
    r += 1

r += 1
sv.cell(row=r, column=1, value="TEST DE ROBUSTESSE COMBINE").font = F_H2
band(sv, r, 6, "TEST DE ROBUSTESSE COMBINE")
r += 1
sv.cell(row=r, column=1,
        value="Scénario de stress : toutes les lignes de revenus à zéro (aucun apport, aucun sponsoring), "
              "seul le levier coûts évités subsiste, au taux du scénario prudent.").font = F_BASE
sv.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
sv.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="center")
sv.row_dimensions[r].height = 30
r += 1
sv.cell(row=r, column=1, value="Equilibre 24 mois en stress total").font = F_TOT
cell = sv.cell(row=r, column=2,
               value=f"='4_PL_Prudent'!B{PL_ROWS['4_PL_Prudent']['ce']}"
                     f"+'4_PL_Prudent'!B{PL_ROWS['4_PL_Prudent']['cf']}"
                     f"+'4_PL_Prudent'!B{PL_ROWS['4_PL_Prudent']['capex']}"
                     f"+'4_PL_Prudent'!B{PL_ROWS['4_PL_Prudent']['cvsms']}"
                     f"+'4_PL_Prudent'!B{PL_ROWS['4_PL_Prudent']['cvinfra']}")
cell.number_format, cell.font, cell.fill, cell.border = FMT_XPF, F_TOT, FILL_AMB, B_ALL
sv.cell(row=r, column=6,
        value="Mesure le plancher du modèle : ce que vaut la proposition si aucune hypothèse "
              "commerciale ne se réalise.").font = F_SMALL
sv.cell(row=r, column=6).alignment = Alignment(wrap_text=True, vertical="center")


# ===========================================================================
# 9 - PARTAGE DE VALEUR
# ===========================================================================
pv = wb.create_sheet("9_Partage_valeur")
pv.sheet_view.showGridLines = False
pv.column_dimensions["A"].width = 26
pv.column_dimensions["B"].width = 30
pv.column_dimensions["C"].width = 30
pv.column_dimensions["D"].width = 30
pv.column_dimensions["E"].width = 46

pv.cell(row=1, column=1, value="REPARTITION DE LA VALEUR - OPT-NC / CANDIDAT / PARTENAIRES").font = F_TITLE
pv.cell(row=2, column=1,
        value="Présenté sans préjuger des conditions d'une future contractualisation, conformément "
              "à la section 6 du dossier candidats.").font = F_SMALL

PH = 4
for c, t in enumerate(["Phase", "OPT-NC", "Candidat", "Partenaires", "Principe retenu"], start=1):
    cell = pv.cell(row=PH, column=c, value=t)
    cell.fill, cell.font, cell.border = FILL_H1, F_H1, B_ALL
    cell.alignment = Alignment(horizontal="center")

PHASES = [
    ("Sprint POC\n(oct. - nov. 2026)",
     "Met à disposition un environnement de test encadré et deux points de mentorat. "
     "Aucune dépense d'industrialisation.",
     "Finance intégralement la réalisation du POC. Conserve la propriété de ses développements.",
     "Sans objet.",
     "La dotation du challenge rémunère l'effort de POC. Aucun transfert de propriété intellectuelle."),
    ("Pilote 90 jours\n(2027, sous réserve)",
     "Finance le pilote au forfait. Bénéficie dès la bascule des coûts évités et de la totalité "
     "de la marge des clients apportés.",
     "Forfait de mise en œuvre et de run. Engagement de résultat sur la disponibilité et la sécurité.",
     "Régie sponsoring rémunérée à la commission (15 % du CA sponsoring).",
     "Modèle au forfait : l'OPT-NC conserve la maîtrise du service public et la totalité du "
     "bénéfice des économies."),
    ("Industrialisation\n(contrat distinct)",
     "Propriétaire du service et de la relation utilisateur. Perçoit les coûts évités et la marge "
     "des clients Helia PRO et grand public apportés.",
     "Redevance de run annuelle, plus un intéressement plafonné sur la valeur d'apport effectivement "
     "attribuée et constatée.",
     "Commission de régie sur le sponsoring uniquement.",
     "L'intéressement aligne les intérêts sans transférer de risque à l'OPT-NC : il n'est dû "
     "que sur une valeur réellement constatée."),
]

r = PH + 1
for phase, opt, cand, part, principe in PHASES:
    for c, txt in enumerate([phase, opt, cand, part, principe], start=1):
        cell = pv.cell(row=r, column=c, value=txt)
        cell.font = F_BOLD if c == 1 else F_BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.border = B_ALL
    pv.row_dimensions[r].height = 72
    r += 1

r += 1
band(pv, r, 5, "PRINCIPES DE GOUVERNANCE DE LA VALEUR")
r += 1
for txt in [
    "L'OPT-NC conserve la totalité du bénéfice des coûts évités : le candidat n'en prélève aucune part. "
    "C'est le levier principal du modèle, et il revient intégralement à l'établissement.",
    "L'intéressement du candidat porte exclusivement sur la valeur d'apport constatée, mesurée par un "
    "dispositif d'attribution contradictoire et plafonnée. Aucun intéressement sur le socle gratuit.",
    "Aucune rémunération ne dépend du volume de SMS envoyés : le modèle n'incite à aucun moment "
    "à augmenter artificiellement le trafic.",
    "Le sponsoring est le seul poste faisant intervenir un tiers rémunéré. Il est plafonné, identifié "
    "et révocable à tout moment par l'OPT-NC.",
    "Aucune donnée personnelle n'entre dans le partage de valeur : ni revente, ni profilage, ni "
    "monetisation d'audience au sens publicitaire classique.",
]:
    cell = pv.cell(row=r, column=1, value="•  " + txt)
    cell.font = F_BASE
    pv.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    cell.alignment = Alignment(wrap_text=True, vertical="center")
    pv.row_dimensions[r].height = 30
    r += 1


# ---------------------------------------------------------------- save -----
wb.move_sheet("3_Unit_economics", offset=-3)
out = "Modele_financier_Mobitag_2026.xlsx"
wb.save(out)
print("Classeur genere :", out)
print("Onglets :", wb.sheetnames)
