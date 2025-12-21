# Guide de création de l'application Automator

Ce guide vous explique pas à pas comment créer une application Mac qui permettra d'extraire les données des PDF d'inscription en glissant-déposant simplement les fichiers.

## 🎯 Objectif

Créer une application Mac qui :
- Accepte les PDF en glisser-déposer
- Extrait automatiquement les données
- Génère les fichiers Excel pour Ammon
- Affiche une notification quand c'est terminé

## 📝 Étapes de création

### 1. Ouvrir Automator

1. Appuyez sur `Cmd + Espace` pour ouvrir Spotlight
2. Tapez "Automator" et appuyez sur Entrée
3. Ou : `Applications` > `Automator`

### 2. Créer une nouvelle application

1. Dans la fenêtre qui s'ouvre, cliquez sur "Nouvelle Application"
2. Si la fenêtre n'apparaît pas :
   - Menu `Fichier` > `Nouveau`
   - Choisissez "Application"
   - Cliquez sur "Choisir"

### 3. Ajouter l'action Shell Script

1. Dans la colonne de gauche, recherchez "shell" dans la barre de recherche
2. Double-cliquez sur "Exécuter un script Shell"
3. L'action apparaît dans la zone de droite

### 4. Configurer l'action

Dans l'action "Exécuter un script Shell" :

**Paramètres en haut :**
- Shell : `/bin/bash`
- Passer en entrée : `comme arguments`

**Code du script :**

```bash
#!/bin/bash

# Configuration
SCRIPT_PATH="$HOME/Documents/INBP/extract_inscription.py"
TEMPLATE_PATH="$HOME/Documents/INBP/Import_Entreprises.xlsx"
OUTPUT_DIR="$HOME/Desktop/Imports_Ammon"
TEMP_DIR="$HOME/Desktop/.temp_inscriptions_$$"

# Créer les dossiers
mkdir -p "$OUTPUT_DIR"
mkdir -p "$TEMP_DIR"

# Copier tous les PDFs reçus dans le dossier temporaire
for pdf_file in "$@"
do
    if [[ "$pdf_file" == *.pdf ]]; then
        filename=$(basename "$pdf_file")
        cp "$pdf_file" "$TEMP_DIR/"
        echo "📄 Ajouté: $filename"
    fi
done

# Compter les PDFs
pdf_count=$(ls "$TEMP_DIR"/*.pdf 2>/dev/null | wc -l)

if [ $pdf_count -eq 0 ]; then
    osascript -e "display notification \"Aucun PDF à traiter\" with title \"Extracteur INBP\" sound name \"Basso\""
    rm -rf "$TEMP_DIR"
    exit 1
fi

# Message de début
if [ $pdf_count -eq 1 ]; then
    message="Traitement de 1 PDF..."
else
    message="Traitement de $pdf_count PDFs..."
fi

echo "🔄 $message"
osascript -e "display notification \"$message\" with title \"Extracteur INBP\""

# Traiter tous les PDFs en une seule fois (génère UN SEUL fichier Excel)
if [ -f "$TEMPLATE_PATH" ]; then
    python3 "$SCRIPT_PATH" "$TEMP_DIR" \
        --template "$TEMPLATE_PATH" \
        --output "$OUTPUT_DIR" 2>&1
else
    python3 "$SCRIPT_PATH" "$TEMP_DIR" \
        --output "$OUTPUT_DIR" 2>&1
fi

# Vérifier le résultat
if [ $? -eq 0 ]; then
    if [ $pdf_count -eq 1 ]; then
        success_message="1 entreprise extraite"
    else
        success_message="$pdf_count entreprises dans 1 fichier Excel"
    fi
    
    osascript -e "display notification \"$success_message\" with title \"Extracteur INBP\" sound name \"Glass\""
    
    # Ouvrir le dossier de sortie
    open "$OUTPUT_DIR"
    
    echo "✅ $success_message"
else
    osascript -e "display notification \"Erreur lors du traitement\" with title \"Extracteur INBP\" sound name \"Basso\""
    echo "❌ Erreur lors du traitement"
    rm -rf "$TEMP_DIR"
    exit 1
fi

# Nettoyer le dossier temporaire
rm -rf "$TEMP_DIR"

echo "✨ Terminé ! Fichier disponible dans : $OUTPUT_DIR"
```

### 5. Tester l'application

Avant de sauvegarder :

1. Cliquez sur le bouton "Exécuter" (▶️) en haut à droite
2. Une fenêtre s'ouvre pour sélectionner un PDF
3. Choisissez un bulletin d'inscription PDF
4. Vérifiez que le fichier Excel est créé sur le Bureau dans `Imports_Ammon/`

### 6. Sauvegarder l'application

1. Menu `Fichier` > `Enregistrer` (ou `Cmd + S`)
2. Nom du fichier : **"Extraire Inscription INBP"**
3. Emplacement : **Bureau** (ou Applications)
4. Format du fichier : **Application**
5. Cliquez sur "Enregistrer"

## 🎨 Personnalisation de l'icône (optionnel)

Pour donner une icône personnalisée à votre application :

1. Trouvez une image au format PNG (512x512px recommandé)
2. Ouvrez l'image avec Aperçu
3. `Cmd + A` pour tout sélectionner, `Cmd + C` pour copier
4. Clic droit sur l'application Automator > "Lire les informations"
5. Cliquez sur la petite icône en haut à gauche
6. `Cmd + V` pour coller la nouvelle icône

## 🚀 Utilisation de l'application

### Méthode 1 : Glisser-déposer (RECOMMANDÉ pour batch)

**Pour traiter plusieurs PDFs en une seule fois :**
1. Sélectionnez plusieurs PDF (Cmd+clic ou Cmd+A pour tout sélectionner)
2. Glissez-les TOUS ENSEMBLE sur l'icône de l'application
3. Attendez la notification (quelques secondes)
4. **UN SEUL fichier Excel** est créé avec toutes les entreprises
5. Le dossier `Imports_Ammon` s'ouvre automatiquement

**Exemple :**
- Vous avez 5 bulletins d'inscription (PDF)
- Vous les sélectionnez tous
- Vous les glissez sur l'app
- Résultat : **1 fichier Excel avec 5 entreprises**

### Méthode 2 : Double-clic
1. Double-cliquez sur l'application
2. Une fenêtre s'ouvre pour sélectionner les PDF
3. Vous pouvez en sélectionner plusieurs (Cmd+clic)
4. Cliquez sur "Choisir"
5. Un seul fichier Excel est généré

### 🎯 Avantages du mode batch

✅ **Gain de temps** : Un seul import dans Ammon au lieu de plusieurs
✅ **Organisation** : Toutes les inscriptions d'une session dans un seul fichier
✅ **Moins d'erreurs** : Manipulation unique du fichier Excel
✅ **Traçabilité** : Vue d'ensemble de toutes les inscriptions en un coup d'œil

## 📍 Emplacement des fichiers

```
Bureau/
└── Imports_Ammon/
    ├── Import_Entreprise_20251221_120000.xlsx
    ├── Import_Entreprise_20251221_120100.xlsx
    └── ...

Documents/
└── INBP/
    ├── extract_inscription.py
    ├── Import_Entreprises.xlsx (template, optionnel)
    └── README.md
```

## 🔧 Dépannage

### L'application ne se lance pas

**Problème** : Message "L'application ne peut pas être ouverte"

**Solution** :
1. Clic droit sur l'application > "Ouvrir"
2. Ou : `Préférences Système` > `Sécurité et confidentialité` > Autoriser

### Erreur "python3: command not found"

**Solution** :
1. Ouvrez Terminal
2. Tapez : `python3 --version`
3. Si erreur, installez Python depuis https://www.python.org/downloads/

### Erreur "No module named 'pypdf'"

**Solution** :
1. Ouvrez Terminal
2. Tapez : `pip3 install pypdf openpyxl --break-system-packages`

### Rien ne se passe après glisser-déposer

**Vérifications** :
1. Le script Python est-il dans `~/Documents/INBP/` ?
2. Les dépendances sont-elles installées ?
3. Consultez la Console (Applications > Utilitaires > Console) pour voir les erreurs

### Les données ne sont pas extraites correctement

**Causes possibles** :
- Le PDF n'est pas au format standard INBP
- Le PDF est protégé ou crypté
- Les champs du formulaire ne sont pas remplis

**Solution** :
- Vérifiez que le PDF s'ouvre normalement
- Contactez le support si le problème persiste

## 💡 Astuces

### Ajouter l'application au Dock

Glissez l'application vers le Dock pour un accès rapide.

### Créer un raccourci clavier

1. `Préférences Système` > `Clavier` > `Raccourcis`
2. `Services` > Trouvez votre application
3. Ajoutez un raccourci

### Workflow quotidien optimisé

**Scénario** : Vous recevez 10 inscriptions par email chaque matin

1. Enregistrez tous les PDFs dans un dossier temporaire (ex: Bureau/PDFs_du_jour)
2. Sélectionnez tous les PDFs (Cmd+A)
3. Glissez-les sur l'application
4. **Résultat** : 1 fichier Excel avec 10 entreprises
5. Importez ce fichier unique dans Ammon

### Traitement par session de formation

**Organisation recommandée :**
```
Bureau/
└── Inscriptions_Stage_Galettes_Oct2026/
    ├── bulletin_dupont.pdf
    ├── bulletin_martin.pdf
    ├── bulletin_bernard.pdf
    └── ...
```

Glissez tous les PDFs du dossier → **1 fichier Excel par session**

### Mode fichier unique vs mode batch

Le script détecte automatiquement :
- **1 PDF** → Génère `Import_Entreprise_AAAAMMJJ_HHMMSS.xlsx`
- **Plusieurs PDFs** → Génère `Import_Entreprises_BATCH_AAAAMMJJ_HHMMSS.xlsx`

### Logs détaillés

Pour voir les logs en temps réel et déboguer :
1. Ne pas double-cliquer sur l'application
2. Clic droit > "Ouvrir avec" > "Utilitaire de script"
3. Les logs s'affichent dans une fenêtre
4. Utile pour comprendre pourquoi une extraction échoue

## 📚 Ressources

- [Documentation Automator (Apple)](https://support.apple.com/fr-fr/guide/automator/)
- [Guide Python pour macOS](https://docs.python.org/fr/3/using/mac.html)

## ✅ Checklist finale

Avant de considérer l'installation terminée :

- [ ] Python 3 est installé
- [ ] Les modules pypdf et openpyxl sont installés
- [ ] Le script est dans ~/Documents/INBP/
- [ ] L'application Automator est créée et sauvegardée
- [ ] Un test avec un PDF réel a réussi
- [ ] Le dossier Imports_Ammon est créé sur le Bureau
- [ ] Les notifications fonctionnent

## 🎉 Félicitations !

Votre extracteur automatique est maintenant opérationnel. Il suffit de glisser-déposer les PDF d'inscription pour générer automatiquement les fichiers Excel d'import Ammon !
