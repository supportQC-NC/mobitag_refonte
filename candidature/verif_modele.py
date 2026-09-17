# -*- coding: utf-8 -*-
"""Recalcul independant du modele financier, pour verifier les resultats
et alimenter la note de candidature avec des chiffres exacts."""

M = 24

HYP = {
    #                       prudent   central   ambitieux
    "mau0":      (5000,    8000,    12000),
    "growth":    (0.015,   0.025,   0.035),
    "maucap":    (10000,   18000,   22000),
    "smspm":     (3.0,     3.5,     4.0),

    "prodet":    (0.004,   0.007,   0.010),
    "proclic":   (0.25,    0.35,    0.45),
    "proconv":   (0.04,    0.06,    0.09),
    "promarge":  (90000,   130000,  190000),
    "proattr":   (0.30,    0.40,    0.45),
    "prochurn":  (0.020,   0.013,   0.009),
    "prostart":  (10,      8,       7),

    "nolinepct": (0.12,    0.18,    0.25),
    "b2cclic":   (0.08,    0.12,    0.15),
    "b2cconv":   (0.04,    0.06,    0.07),
    "b2cmarge":  (18000,   26000,   36000),
    "b2cattr":   (0.25,    0.35,    0.40),
    "b2cchurn":  (0.024,   0.018,   0.013),
    "b2cstart":  (10,      8,       7),

    "sponsor":   (0,       120000,  280000),
    "sponstart": (25,      14,      11),
    "sponramp":  (6,       6,       5),
    "sponrev":   (0.15,    0.15,    0.15),

    "costsms":   (2.5,     2.0,     1.5),
    "costmau":   (4,       3,       2),

    "etp":       (0.28,    0.25,    0.20),
    "etpcost":   (8500000, 8000000, 7500000),
    "cloud":     (70000,   55000,   45000),
    "support":   (30000,   22000,   16000),
    "comm":      (25000,   40000,   60000),

    "capex":     (5500000, 4800000, 4200000),
    "devenv":    (20000,   25000,   30000),
    "capexm":    (9,       7,       6),

    "is":        (0.30,    0.30,    0.30),
    "switch":    (11,      10,      9),
    "switchramp":(4,       3,       2),
}

# poste, cout annuel, taux evitement (p, c, a)
CE_POSTS = [
    ("Hebergement / infra on-premise",        1800000, 0.75, 0.85, 0.95),
    ("TMA et maintenance corrective",         3600000, 0.72, 0.80, 0.90),
    ("Exploitation interne (0,25 ETP)",       2000000, 0.60, 0.75, 0.85),
    ("Support N1/N2 abus et reclamations",     900000, 0.60, 0.75, 0.90),
    ("Licences et composants proprietaires",   450000, 0.85, 0.95, 1.00),
    ("Acheminement des SMS du service",         700000, 1.00, 1.00, 1.00),
]

SCEN = ["Prudent", "Central", "Ambitieux"]


def run(i):
    h = {k: v[i] for k, v in HYP.items()}
    ce_total = sum(p[1] for p in CE_POSTS)
    ce_year = sum(p[1] * p[2 + i] for p in CE_POSTS)

    mau = stockpro = stockb2c = 0.0
    cum_rai = 0.0
    cum_rn = 0.0
    low = 0.0
    breakeven = None
    T = {k: 0.0 for k in ("ca", "capro", "cab2c", "caspon", "cv", "cvsms", "cvinfra",
                          "cvrev", "mc", "ce", "cf", "capex", "is", "rn", "sms",
                          "convpro", "convb2c")}
    last = {}

    for m in range(1, M + 1):
        mau = h["mau0"] if m == 1 else min(h["maucap"], mau * (1 + h["growth"]))
        sms = mau * h["smspm"]
        detpro = mau * h["prodet"]
        detb2c = mau * h["nolinepct"]

        convpro = detpro * h["proclic"] * h["proconv"] if m >= h["prostart"] else 0.0
        stockpro = stockpro * (1 - h["prochurn"]) + convpro
        convb2c = detb2c * h["b2cclic"] * h["b2cconv"] if m >= h["b2cstart"] else 0.0
        stockb2c = stockb2c * (1 - h["b2cchurn"]) + convb2c

        capro = stockpro * h["promarge"] / 12 * h["proattr"]
        cab2c = stockb2c * h["b2cmarge"] / 12 * h["b2cattr"]
        caspon = (0.0 if m < h["sponstart"]
                  else h["sponsor"] * min(1, (m - h["sponstart"] + 1) / h["sponramp"]))
        ca = capro + cab2c + caspon

        cvsms = -sms * h["costsms"]
        cvinfra = -mau * h["costmau"]
        cvrev = -caspon * h["sponrev"]
        cv = cvsms + cvinfra + cvrev
        mc = ca + cv

        ramp = 0.0 if m < h["switch"] else min(1, (m - h["switch"] + 1) / h["switchramp"])
        ce = ce_year / 12 * ramp

        if m >= h["switch"]:
            cf = -(h["etp"] * h["etpcost"] / 12 + h["cloud"] + h["support"] + h["comm"])
        else:
            cf = -h["devenv"]
        capex = -h["capex"] / h["capexm"] if m <= h["capexm"] else 0.0

        rai = mc + ce + cf + capex
        cum_rai += rai
        tax = -(h["is"] * min(rai, cum_rai)) if (rai > 0 and cum_rai > 0) else 0.0
        rn = rai + tax
        cum_rn += rn
        low = min(low, cum_rn)
        if breakeven is None and cum_rn > 0:
            breakeven = m

        for k, v in (("ca", ca), ("capro", capro), ("cab2c", cab2c), ("caspon", caspon),
                     ("cv", cv), ("cvsms", cvsms), ("cvinfra", cvinfra), ("cvrev", cvrev),
                     ("mc", mc), ("ce", ce), ("cf", cf), ("capex", capex),
                     ("is", tax), ("rn", rn), ("sms", sms),
                     ("convpro", convpro), ("convb2c", convb2c)):
            T[k] += v
        last = dict(mau=mau, stockpro=stockpro, stockb2c=stockb2c, rn=rn,
                    mc=mc, ce=ce, cf=cf, ca=ca)

    return dict(h=h, T=T, last=last, cum_rn=cum_rn, low=low, breakeven=breakeven,
                ce_total=ce_total, ce_year=ce_year)


def f(x):
    return f"{x:,.0f}".replace(",", " ")


res = [run(i) for i in range(3)]

print("=" * 96)
print("VERIFICATION DU MODELE FINANCIER - MOBITAG NG")
print("=" * 96)

print("\n--- COUTS EVITES (onglet 2) ---")
print(f"Cout complet annuel actuel du service : {f(res[0]['ce_total'])} XPF/an")
for i, s in enumerate(SCEN):
    r = res[i]
    print(f"  {s:10s} : economie retenue {f(r['ce_year'])} XPF/an "
          f"({r['ce_year']/r['ce_total']*100:.0f} % du cout complet) "
          f"= {f(r['ce_year']/12)} XPF/mois")

print("\n--- ACTIVITE AU MOIS 24 ---")
for i, s in enumerate(SCEN):
    l = res[i]["last"]
    print(f"  {s:10s} : MAU {f(l['mau'])} | clients Helia PRO apportes {l['stockpro']:.0f} "
          f"| lignes grand public {l['stockb2c']:.0f}")

print("\n--- COMPTE DE RESULTAT CUMULE 24 MOIS (XPF) ---")
hdr = f"{'Poste':<42}" + "".join(f"{s:>17}" for s in SCEN)
print(hdr)
print("-" * 96)
rows = [
    ("Apport d'affaires B2B (Helia PRO)", "capro"),
    ("Apport d'affaires B2C (lignes)", "cab2c"),
    ("Sponsoring consenti", "caspon"),
    ("CHIFFRE D'AFFAIRES", "ca"),
    ("Couts variables et reversements", "cv"),
    ("MARGE CONTRIBUTIVE", "mc"),
    ("Couts evites", "ce"),
    ("Couts fixes", "cf"),
    ("Investissements", "capex"),
    ("Fiscalite (IS)", "is"),
]
for label, key in rows:
    print(f"{label:<42}" + "".join(f"{f(res[i]['T'][key]):>17}" for i in range(3)))
print("-" * 96)
print(f"{'EQUILIBRE ECONOMIQUE 24 MOIS':<42}" + "".join(f"{f(res[i]['cum_rn']):>17}" for i in range(3)))

print("\n--- INDICATEURS DE PILOTAGE ---")
for i, s in enumerate(SCEN):
    r = res[i]
    if r["breakeven"]:
        be = f"M{r['breakeven']}"
    else:
        reg = r["last"]["rn"]
        be = f"~M{24 + int(-r['cum_rn']/reg) + 1}" if reg > 0 else "non atteint"
    print(f"  {s:10s} : point mort {be:>10s} | besoin de financement max {f(r['low']):>12} XPF"
          f" | resultat M24 annualise {f(r['last']['rn']*12):>12} XPF/an")

print("\n--- STRUCTURE DE L'EQUILIBRE AU MOIS 24 ---")
for i, s in enumerate(SCEN):
    l = res[i]["last"]
    tot = l["ce"] + l["mc"]
    print(f"  {s:10s} : couts evites {l['ce']/tot*100:5.1f} % | apports {l['mc']/tot*100:5.1f} % "
          f"| couverture des couts fixes {(l['mc']+l['ce'])/-l['cf']*100:5.0f} %")

print("\n--- TEST DE STRESS : toutes lignes de revenus a zero, scenario prudent ---")
r0 = res[0]
stress = r0["T"]["ce"] + r0["T"]["cf"] + r0["T"]["capex"] + r0["T"]["cvsms"] + r0["T"]["cvinfra"]
print(f"  Equilibre 24 mois en stress total : {f(stress)} XPF")

print("\n--- COUT DE REVIENT DU SOCLE GRATUIT (mois 24, annualise) ---")
for i, s in enumerate(SCEN):
    l = res[i]["last"]
    cout_an = -(res[i]["T"]["cv"] / 24 + l["cf"]) * 12
    print(f"  {s:10s} : {f(cout_an)} XPF/an  (vs {f(res[i]['ce_total'])} XPF/an aujourd'hui)"
          f"  -> {f(res[i]['ce_total'] - cout_an)} XPF/an d'ecart brut")
print()
