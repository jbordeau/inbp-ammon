#!/bin/bash

# On se place dans le dossier où se trouve le script
cd "$(dirname "$0")"

clear
echo "=========================================="
echo "   🥐 INSTALLATEUR EXTRACTEUR INBP"
echo "=========================================="
echo ""

# 1. Vérification de Python 3
if ! command -v python3 &> /dev/null; then
    echo "❌ Erreur : Python 3 n'est pas installé sur ce Mac."
    echo "Veuillez l'installer depuis https://www.python.org/downloads/"
    exit 1
fi

# 2. Création de l'environnement virtuel
echo "📦 Création de l'environnement de travail (.venv)..."
python3 -m venv .venv
if [ $? -ne 0 ]; then
    echo "❌ Erreur lors de la création du venv."
    exit 1
fi

# 3. Installation des dépendances
echo "usr Installation des modules nécessaires (Mistral, Excel, Pandas)..."
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
if [ $? -ne 0 ]; then
    echo "❌ Erreur lors de l'installation des dépendances."
    exit 1
fi

# 4. Préparation du fichier .env
if [ ! -f .env ]; then
    echo "📝 Création du fichier de configuration .env..."
    echo "MISTRAL_API_KEY=" > .env
    echo "✅ Fichier .env créé."
else
    echo "ℹ️  Fichier .env déjà existant."
fi

# 5. Création des dossiers nécessaires
mkdir -p input output existants
echo "✅ Dossiers de travail vérifiés."

echo ""
echo "=========================================="
echo "🎉 INSTALLATION TERMINÉE AVEC SUCCÈS !"
echo "=========================================="
echo ""
echo "PROCHAINES ÉTAPES :"
echo "1. Ouvrez le fichier '.env' et collez votre clé MISTRAL_API_KEY."
echo "2. Placez vos exports Ammon dans le dossier 'existants'."
echo "3. Configurez votre application Automator."
echo ""
read -p "Appuyez sur [Entrée] pour quitter..."
