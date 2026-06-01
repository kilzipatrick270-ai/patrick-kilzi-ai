#!/bin/bash

echo "=================================================="
echo "  🎯 INSIGHTER AI — Démarrage du serveur"
echo "=================================================="
echo ""

# Vérifier Python
if ! command -v python3 &> /dev/null; then
    echo "❌ Python3 n'est pas installé"
    echo "   Installation : brew install python3 (Mac) ou apt-get install python3 (Linux)"
    exit 1
fi

echo "✅ Python $(python3 --version | cut -d' ' -f2) détecté"
echo ""

# Vérifier clé API
if [ -z "$ANTHROPIC_API_KEY" ]; then
    echo "⚠️  ANTHROPIC_API_KEY non définie"
    echo ""
    echo "   Définis ta clé API :"
    echo "   export ANTHROPIC_API_KEY='sk-ant-...'"
    echo ""
    read -p "   Veux-tu continuer sans clé ? (non recommandé) [y/N]: " -n 1 -r
    echo
    if [[ ! $REPLY =~ ^[Yy]$ ]]; then
        exit 1
    fi
else
    echo "✅ ANTHROPIC_API_KEY détectée"
fi

echo ""
echo "Vérification des dépendances..."

# Vérifier dependencies
pip3 show anthropic flask python-pptx &> /dev/null
if [ $? -ne 0 ]; then
    echo "📦 Installation des dépendances..."
    pip3 install anthropic flask python-pptx
fi

echo "✅ Toutes les dépendances sont OK"
echo ""
echo "=================================================="
echo "  🚀 Démarrage du serveur..."
echo "=================================================="
echo ""
echo "   Accès : http://localhost:5002"
echo "   Appuie sur CTRL+C pour arrêter"
echo ""
echo "=================================================="
echo ""

python3 app.py
