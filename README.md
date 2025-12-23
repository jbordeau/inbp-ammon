# Extracteur de bulletins d'inscription INBP

Outil d'extraction automatique des données des bulletins d'inscription PDF pour import dans Ammon Campus.

## 📋 Prérequis

### Sur Mac
Python 3 est préinstallé sur macOS. Vérifiez avec :
```bash
python3 --version
```

### Installation des dépendances

```bash
pip3 install pypdf openpyxl --break-system-packages
```

Ou si vous préférez utiliser un environnement virtuel :
```bash
python3 -m venv venv
source venv/bin/activate
pip install pypdf openpyxl
```

## 🚀 Utilisation

### Utilisation avec un seul PDF

```bash
python3 ammon_inscription.py chemin/vers/bulletin.pdf
```

Le fichier Excel sera généré dans le répertoire courant avec 1 entreprise.

### Utilisation avec plusieurs PDFs (BATCH)

```bash
# Pointer vers un dossier contenant plusieurs PDFs
python3 ammon_inscription.py chemin/vers/dossier_pdfs/
```

**🎯 Un SEUL fichier Excel sera généré** contenant toutes les entreprises extraites des PDFs du dossier.

### Avec options

```bash
# Spécifier un dossier de sortie
python3 ammon_inscription.py dossier_pdfs/ --output ~/Downloads

# Utiliser le template Ammon (recommandé)
python3 ammon_inscription.py dossier_pdfs/ --template Import_Entreprises.xlsx --output ~/Downloads
```

## 📋 Exemples concrets

### Exemple 1 : Un seul PDF
```bash
python3 ammon_inscription.py "Bulletin_Dupont.pdf"
# Génère : Import_Entreprise_20241221_143000.xlsx (1 entreprise)
```

### Exemple 2 : Plusieurs PDFs dans un dossier
```bash
python3 ammon_inscription.py "Inscriptions_Janvier/"
# Dossier contient : bulletin_01.pdf, bulletin_02.pdf, bulletin_03.pdf
# Génère : Import_Entreprises_BATCH_20241221_143000.xlsx (3 entreprises)
```

### Exemple 3 : Workflow quotidien
```bash
# 1. Créer un dossier pour les PDFs du jour
mkdir ~/Bureau/Inscriptions_$(date +%Y%m%d)

# 2. Y déposer tous les PDFs reçus par email

# 3. Lancer l'extraction
python3 ammon_inscription.py ~/Bureau/Inscriptions_$(date +%Y%m%d) \
    --template ~/Documents/INBP/Import_Entreprises.xlsx \
    --output ~/Bureau/Imports_Ammon

# 4. Un seul fichier Excel avec toutes les entreprises est créé
```

## 📦 Ce qui est extrait

Le script extrait automatiquement :

### Données entreprise
- ✅ Nom de l'entreprise
- ✅ Adresse complète
- ✅ Code postal
- ✅ Ville
- ✅ Pays
- ✅ SIRET
- ✅ Code NAF(A)
- ✅ Téléphone
- ✅ Email

### Données stagiaire
- ✅ Nom
- ✅ Prénom
- ✅ Date de naissance

## 📊 Format de sortie

Le script génère un fichier Excel compatible avec l'import Ammon Campus :
- `Import_Entreprise_AAAAMMJJ_HHMMSS.xlsx`

Le fichier contient :
- Une référence externe unique (INBP_SIRET_TIMESTAMP)
- Toutes les données au format attendu par Ammon
- Type d'entreprise : "E" (Entreprise)
- Est un siège : Oui (-1)
- Adresse de type : "Siège"

## ⚙️ Automatisation avec Automator (Mac)

### Création de l'application Automator

1. **Ouvrir Automator**
   - Applications > Automator

2. **Créer une nouvelle application**
   - Fichier > Nouveau
   - Choisir "Application"

3. **Ajouter l'action "Exécuter un script Shell"**
   - Rechercher "Exécuter un script Shell" dans la bibliothèque
   - Glisser-déposer dans le workflow

4. **Configurer le script**
   - Shell : `/bin/bash`
   - Passer en entrée : `comme arguments`
   - Coller le code suivant :

```bash
#!/bin/bash

# Configuration
SCRIPT_PATH="$HOME/Documents/INBP/ammon_inscription.py"
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
        cp "$pdf_file" "$TEMP_DIR/"
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

osascript -e "display notification \"$message\" with title \"Extracteur INBP\""

# Traiter tous les PDFs en une seule fois
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
        success_message="$pdf_count entreprises extraites"
    fi
    
    osascript -e "display notification \"$success_message avec succès\" with title \"Extracteur INBP\" sound name \"Glass\""
    open "$OUTPUT_DIR"
else
    osascript -e "display notification \"Erreur lors du traitement\" with title \"Extracteur INBP\" sound name \"Basso\""
fi

# Nettoyer
rm -rf "$TEMP_DIR"

echo "✨ Terminé !"
```

5. **Sauvegarder l'application**
   - Fichier > Enregistrer
   - Nom : "Extraire Inscription INBP"
   - Emplacement : Bureau ou Applications

### Utilisation de l'application Automator

#### Mode 1 : Glisser-déposer plusieurs PDFs
1. Sélectionnez un ou plusieurs PDF (Cmd+clic pour sélection multiple)
2. Glissez-les TOUS ENSEMBLE sur l'icône de l'application
3. **Un seul fichier Excel** est généré avec toutes les entreprises
4. Le dossier `Imports_Ammon` s'ouvre automatiquement

#### Mode 2 : Double-clic
1. Double-cliquez sur l'application
2. Une fenêtre s'ouvre pour sélectionner les PDF
3. Vous pouvez en sélectionner plusieurs (Cmd+clic)
4. Cliquez sur "Choisir"

### 🎯 Avantages du mode batch

- ✅ **Un seul fichier Excel** pour tous les PDFs du jour
- ✅ Import unique dans Ammon (gain de temps)
- ✅ Moins d'erreurs de manipulation
- ✅ Traçabilité : toutes les inscriptions d'une session ensemble

## 📁 Structure des fichiers

```
INBP/
├── ammon_inscription.py          # Script principal
├── Import_Entreprises.xlsx         # Template Ammon (optionnel)
└── README.md                        # Ce fichier
```

## 🔍 Dépannage

### Le script ne trouve pas certaines données

Le script analyse le texte extrait du PDF. Si certaines données ne sont pas trouvées :
1. Vérifiez que le PDF est bien rempli
2. Vérifiez que le format correspond au bulletin INBP standard
3. Consultez les logs affichés dans le terminal

### Erreur "Module not found"

```bash
pip3 install pypdf openpyxl --break-system-packages
```

### Problèmes d'encodage

Le script gère automatiquement les accents et caractères spéciaux français.

## 📝 Notes importantes

- **Mode batch** : Lorsque vous passez un dossier, UN SEUL fichier Excel est généré avec toutes les entreprises
- **Mode unitaire** : Lorsque vous passez un seul PDF, un fichier Excel avec une seule entreprise est créé
- Les données de stagiaire sont également extraites pour référence
- Le SIRET est utilisé pour générer une référence externe unique
- Pour les auto-entrepreneurs, le nom de l'entreprise = Prénom + Nom
- Chaque entreprise dans le fichier batch a une référence unique (suffixe _001, _002, etc.)

## 🔄 Prochaines étapes

1. ✅ Import des entreprises (fichier unique ou batch)
2. ⏳ Import des stagiaires (à venir)
3. ⏳ Vérification des doublons SIRET avant génération (à venir)
4. ⏳ Option pour fusionner avec un fichier Excel existant (à venir)

## 📧 Support

Pour toute question ou problème :
- Vérifiez d'abord ce README
- Consultez les messages d'erreur dans le terminal
- Testez avec le PDF d'exemple fourni
