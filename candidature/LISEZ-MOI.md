# Dossier de candidature — Challenge Mobitag 2026

**Proposition : « Mobitag Nouvelle Génération — Le canal qui reste »**
Date limite de dépôt : **30 septembre 2026, 23h (heure de Nouméa)** → `concoursopt@lafrenchtech.nc`

---

## Ce qui est livré

| Fichier | Rôle | État |
|---|---|---|
| `Note_de_candidature.pdf` | Note de candidature — **8 pages exactement** (limite du dossier) | ✅ Prêt, placeholders à compléter |
| `Note_de_candidature.html` | Source éditable de la note | ✅ |
| `Modele_financier_Mobitag_2026.xlsx` | Modèle financier 24 mois, 3 scénarios, 10 onglets | ✅ Formules vérifiées dans Excel |
| `build_modele_financier.py` | Générateur du classeur | Régénère le XLSX à l'identique |
| `verif_modele.py` | Recalcul indépendant du modèle en Python pur | Contrôle croisé |
| `recalc_excel.py` | Ouvre le classeur dans Excel, recalcule, compare à `verif_modele.py` | Contrôle qualité |

**Contrôle effectué** : Excel recalcule exactement les valeurs du modèle Python indépendant sur les six
agrégats clés et les trois scénarios, sans aucune erreur de formule dans le classeur.

### Régénérer les livrables

```bash
python build_modele_financier.py   # régénère le XLSX
python verif_modele.py             # recalcule et affiche tous les chiffres clés
python recalc_excel.py             # contrôle croisé Excel ↔ Python

# régénérer le PDF depuis le HTML
"C:/Program Files/Google/Chrome/Application/chrome.exe" --headless --disable-gpu \
  --no-pdf-header-footer --print-to-pdf="Note_de_candidature.pdf" \
  "file:///C:/Users/Support/Desktop/mobitag/candidature/Note_de_candidature.html"
```

> Excel refuse d'ouvrir un fichier posé sur le Bureau (Mark-of-the-Web) : `recalc_excel.py` recalcule
> donc une copie temporaire. C'est normal, pas un défaut du classeur.

---

## Le positionnement, en une page

**Le recadrage qui structure tout** : le SMS est déjà illimité dans les quatre Forfaits M et les
Recharges Packagées, et une option Pack SMS existe dès 318 F TTC. **Toute monétisation du volume au
particulier est redondante par construction** — un jury qui connaît le catalogue le verrait
immédiatement. Le freemium d'envoi est donc écarté, et aucune ligne de revenu du modèle ne repose sur
le SMS vendu au particulier.

**Les trois espaces blancs réels, et ce qu'on en fait :**

1. **L'usage sans ligne mobile** — tout le catalogue Helia suppose une SIM active. Mobitag est le seul
   canal quand ce n'est pas le cas (mobile perdu, volé, cassé, hors forfait, poste partagé de mairie ou
   de médiathèque). C'est la légitimité du service, et c'est inaccessible à Helia PRO.
2. **Le destinataire et le canal 1000** — tout ce que vend l'OPT-NC est côté émission. Personne ne
   s'occupe de celui qui reçoit : d'où vient ce message, est-il légitime, puis-je signaler. Territoire
   vierge → c'est le F11.
3. **Le coût interne du service** — ce sont des charges, pas une offre : aucune redondance possible.

**Deux garde-fous inscrits dans le produit** (et non dans les intentions), qui rendent la
non-redondance structurelle et vérifiable dans le POC :
- **un seul destinataire par envoi, jamais de liste** ;
- **aucune API d'envoi exposée à des tiers**.

---

## Les chiffres à connaître par cœur

| | Prudent | **Central** | Ambitieux |
|---|---:|---:|---:|
| Chiffre d'affaires 24M | 155 876 | **3 513 500** | 20 811 463 |
| Marge contributive | −1 490 551 | **772 577** | 16 899 492 |
| Coûts évités | 7 046 354 | **8 997 917** | 11 121 250 |
| Coûts fixes | −4 726 667 | **−4 480 000** | −4 176 000 |
| Investissements | −5 500 000 | **−4 800 000** | −4 200 000 |
| **Équilibre 24 mois** | **−4 670 864** | **+343 345** | **+13 751 320** |
| Point mort | ≈ M51 | **M24** | M14 |
| Besoin de financement max | −6 665 768 | **−5 897 720** | −5 067 791 |
| Couverture des coûts d'exploitation à M24 | 155 % | **314 %** | 1 099 % |

**Les trois phrases qui portent le dossier :**
- Le coût complet du service passe de **9,45 M F/an à 4,8 M F/an** — divisé par deux.
- **Le particulier ne paie rien, dans les trois scénarios.**
- Même en scénario prudent, sans aucun sponsoring et avec une audience trois fois plus faible
  qu'espéré, **le service couvre 155 % de ses coûts d'exploitation dès M24**.

---

## Ce qu'il reste à faire

### 1. Compléter les placeholders (surlignés en jaune dans le PDF)
- [ ] Raison sociale, RIDET, forme juridique, date de création, effectif, siège, site web
- [ ] Contact référent + disponibilité pendant le sprint POC (10 oct. → 5 nov.)
- [ ] **Section 9 — l'équipe** : membres, rôles, compétences, partenaires. C'est 10 points au barème.
- [ ] **Section 5.6 — éléments de validation** : entretiens, tests, lettres d'intérêt. Le dossier les
      valorise explicitement, et c'est ce qui sépare un modèle crédible d'un modèle théorique.
- [ ] Déclaration de conflits d'intérêts (une relation existante n'exclut pas, dès lors qu'elle est déclarée)

### 2. Deux questions à envoyer à `concoursopt@lafrenchtech.nc` — **cette semaine**
Les réponses valident ou invalident le positionnement de base, et les hypothèses correspondantes
sont marquées « À VALIDER OPT » dans le classeur :
- [ ] **Quelle part des envois actuels provient d'utilisateurs sans ligne Helia active ?**
      (hypothèse la plus structurante du modèle : 12 % / 18 % / 25 %)
- [ ] **Mobitag est-il accessible depuis l'étranger ?**
- [ ] Bonus utile : volumétrie réelle du service, et éléments de coût complet actuel

### 3. Les autres pièces du dossier
- [ ] **Vidéo de présentation — 3 minutes max**, lien privé accessible au jury
- [ ] **Déclaration sur l'honneur signée** — reproduire l'Annexe B du dossier candidats, signature du
      représentant légal + cachet
- [ ] Annexes facultatives (5 pages max, hors limite des 8 pages)

### 4. Ajuster le modèle si besoin
Toutes les hypothèses sont saisissables dans l'onglet `1_Hypotheses` (cellules jaunes). Modifier une
valeur met à jour les trois scénarios, le point mort et la synthèse. Les postes de coût du service
historique se règlent dans `2_Couts_evites`.

Si vous modifiez le XLSX à la main, pensez à répercuter dans `verif_modele.py` avant de relancer
`recalc_excel.py`, sinon le contrôle croisé signalera des écarts qui n'en sont pas.

---

## Où chaque exigence du dossier est traitée

| Exigence | Où |
|---|---|
| Proposition de valeur | Note § 1 et § 3 |
| Positionnement écosystème + **cartographie des exclusions** | Note § 4 (le tableau attendu comme preuve) |
| Modèle économique 24 mois, 3 scénarios | Note § 5 + classeur complet |
| Marge contributive / équilibre (formules imposées) | Classeur, onglets 4 à 6, lignes éponymes |
| Sensibilité + test de robustesse | Note § 5.5 + onglet `8_Sensibilite` |
| Partage de valeur OPT / candidat / partenaires | Onglet `9_Partage_valeur` |
| Confiance, sécurité, données, anti-abus | Note § 6 (seuil éliminatoire : 8/15) |
| Mise en œuvre phase 2 | Note § 7 |
| Feuille de route, TCO, passage à l'échelle | Note § 8 |
| Équipe | Note § 9 — **à compléter** |
| Déclarations | Note § 10 + Annexe B signée |
| Capacités F01 à F11 | Couvertes par le périmètre décrit en note § 3, récapitulées § 7 |

---

## Calendrier

| Jalon | Date |
|---|---|
| ~~Petit-déjeuner d'information @ STATION N~~ | ~~16 sept. 2026~~ (passé — demander le support) |
| **Clôture des candidatures** | **30 sept. 2026, 23h** |
| Annonce des 3 finalistes | 10 oct. 2026 |
| Sprint POC | 10 oct. → début nov. 2026 |
| Demo Day | 5 nov. 2026 |

**Barème phase 1** — 100 points, seuil recommandé 65, avec deux minima éliminatoires :
**12/25 sur le modèle économique** et **8/15 sur la sécurité et les données**.
