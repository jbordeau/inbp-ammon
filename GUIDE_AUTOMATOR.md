# 🤖 Guide Automator pour Mac

Ce guide explique comment transformer le script en une application "Glisser-Déposer".

## 🛠 Création de l'application

1. **Ouvrir Automator** (via Spotlight : `Cmd + Espace`).
2. **Nouveau Document** > **Application**.
3. Chercher l'action **"Exécuter un script Shell"** et la glisser à droite.
4. Configurer l'action :
   - Shell : `/bin/bash`
   - Passer en entrée : **"comme arguments"**
5. Coller le code suivant (en vérifiant que `INBP_PATH` correspond à votre dossier) :

```bash
#!/bin/bash

# --- CONFIGURATION ---
INBP_PATH="$HOME/Documents/INBP"
PYTHON_EXE="$INBP_PATH/.venv/bin/python3"
SCRIPT_PATH="$INBP_PATH/main.py"
OUTPUT_DIR="$HOME/Desktop/Imports_Ammon"
FILES_INPUT="$@"
LOG_FILE="/tmp/inbp_last_run.log"
# ---------------------

mkdir -p "$OUTPUT_DIR"

# Notification de début
osascript -e "display notification \"Analyse des PDFs en cours (Mistral AI)...\" with title \"INBP Automator\""

# Exécution du script (on passe tous les fichiers glissés en argument)
# On utilise le python du venv directement
"$PYTHON_EXE" "$SCRIPT_PATH" -i "$FILES_INPUT" -o "$OUTPUT_DIR"

if [ $? -eq 0 ]; then
    osascript -e "display notification \"Extraction terminée ! Fichiers dispos sur le bureau.\" with title \"INBP Automator\" sound name \"Glass\""
    open "$OUTPUT_DIR"
else
    osascript -e "display notification \"Erreur lors de l'extraction. Vérifiez la connexion ou la clé API.\" with title \"INBP Automator\" sound name \"Basso\""
fi
open -a "Console" "$LOG_FILE"
```
6. **Enregistrer** l'application sur votre Bureau sous le nom **"Extraire Inscription INBP"**.

## 🚀 Comment l'utiliser ?
- **Unitaire** : Glissez un PDF sur l'icône de l'application.
- **Batch** : Glissez un **dossier** contenant plusieurs PDF sur l'icône.
- **Résultat** : Le dossier `Imports_Ammon` sur votre bureau s'ouvrira avec les fichiers Excel prêts pour Ammon Campus.

## ⚠️ Dépannage
- Si rien ne se passe, vérifiez que votre fichier `.env` contient bien la clé Mistral.
- Vérifiez que le chemin `INBP_PATH` dans le script Automator est le bon.