# 🎯 INSIGHTER AI — Générateur de Présentations

**Génère des présentations PowerPoint professionnelles en 60 secondes grâce à l'IA.**

---

## 📋 Fonctionnalités

✅ **Recherche Données Financières**
- Rentre le nom d'une entreprise → récupère ses chiffres des 3 dernières années
- Calcule automatiquement : ROE, ROA, EPS
- Affiche : compte de résultat, cash flow, bilan, notation crédit

✅ **Génération Présentation**
- Convertit les données en présentation PowerPoint pro
- 10 slides avec design premium (navy/teal/amber)
- Contient : résumé exécutif, KPIs, recommandations, prochaines étapes

✅ **Interface Web**
- Accès simple via navigateur
- 2 onglets : Recherche → Génération
- Auto-intégration des données

---

## 🚀 Installation & Lancement

### Prérequis
- **Python 3.9+** (vérifier : `python3 --version`)
- **Clé API Anthropic** (gratuite sur https://console.anthropic.com)

### Étapes

#### 1️⃣ Extraire le ZIP
```bash
unzip insighter_ai_v1.zip
cd insighter_ai
```

#### 2️⃣ Installer les dépendances
```bash
pip3 install anthropic flask python-pptx
```

#### 3️⃣ Définir la clé API
```bash
export ANTHROPIC_API_KEY='ta-clé-api-ici'
```

#### 4️⃣ Lancer le serveur
```bash
python3 app.py
```

Vous verrez :
```
=======================================================
  INSIGHTER AI — Serveur Web
=======================================================
  Accès : http://localhost:5002
=======================================================
```

#### 5️⃣ Ouvrir dans le navigateur
Allez à : **http://localhost:5002**

---

## 📖 Guide d'Utilisation

### Onglet 1 : Recherche Données

1. Rentre le **nom d'une entreprise** (ex: "TotalEnergies", "Apple", "CMA CGM")
2. Clique **"Chercher les données"**
3. Attends 30 secondes pour le résultat
4. Clique **"Utiliser pour présentation"** → passe à l'onglet 2

### Onglet 2 : Génération Présentation

1. Les données sont **déjà pré-remplies** (client + chiffres)
2. Modifie si besoin (type de présentation, langue)
3. Clique **"Générer la présentation"**
4. Attends ~45 secondes
5. Clique **"Télécharger"** → PowerPoint téléchargé ✅

---

## 📁 Structure des Fichiers

```
insighter_ai/
├── app.py                      # Serveur Flask
├── generate_presentation.py    # Moteur de génération PPTX
├── financial_data.py           # Recherche données financières
├── templates/
│   └── index.html              # Interface web
└── exports/                    # Dossier des PPTX générés
```

---

## 🎨 Design de la Présentation

Palette Patrick Kilzi (professionnelle) :
- **Navy** (#1A3C5E) — titres, accents
- **Teal** (#0D9488) — séparateurs, highlights
- **Amber** (#F59E0B) — barres d'accent
- **Blanc/Gris** — backgrounds, texte

Slides générées automatiquement :
1. Titre (dark background)
2. Résumé exécutif (4 points clés)
3. KPIs (4 cartes avec chiffres)
4. Sections contenu (2-4 slides)
5. Recommandations (4 points)
6. Prochaines étapes (4 étapes)
7. Clôture (remerciements)

---

## 🔍 Fonctionnalités Avancées

### Sources de Données
L'IA cherche dans :
1. APIs financières officielles (si disponibles)
2. Web search pour les rapports publics
3. Données officielles de l'entreprise

### Devises
- **Toutes les devises** converties en **USD** automatiquement
- 3 dernières années d'historique
- Calculs de ratios inclus

### Ratios Calculés
- **ROE** (Return on Equity) — rentabilité pour actionnaires
- **ROA** (Return on Assets) — efficacité des actifs
- **EPS** (Earnings Per Share) — bénéfice par action

---

## ⚠️ Notes Importantes

1. **Clé API** : Garde-la **secrète**, ne la partage pas publiquement
2. **Entreprises publiques** : Les meilleures résultats pour les sociétés cotées
3. **Temps de traitement** : 30-60 secondes selon l'IA
4. **Fichiers générés** : Sauvegardés dans `exports/` du serveur

---

## 🐛 Troubleshooting

### Erreur : "ANTHROPIC_API_KEY not found"
```bash
export ANTHROPIC_API_KEY='ta-clé'
python3 app.py
```

### Erreur : "Port 5002 already in use"
```bash
# Utilise un autre port, modifie app.py ligne dernière :
app.run(debug=False, port=5003)  # Remplace 5002 par 5003
```

### "Entreprise non trouvée"
- Essaie le nom officiel complet (pas le diminutif)
- Entreprises privées : les données sont moins précises

### Les chiffres semblent incorrects
- Vérifie la source (rapport annuel officiel)
- Les données publiques peuvent avoir des variations selon la source

---

## 📞 Support

Pour des questions ou améliorations :
1. Vérifie que Python 3.9+ est installé
2. Vérifie que la clé API est valide
3. Redémarre le serveur

---

## 📄 Licence

**** — Patrick Kilzi  
Fondateur : Data, IA, Informatique

Utilisation : Usage personnel et commercial ✅

---

**Version** : 1.0  
**Date** : Juin 2024  
**Framework** : Flask + Claude API + python-pptx
