# 🥐 Extracteur d'inscriptions INBP

Outil d'intelligence artificielle (Mistral OCR) pour extraire automatiquement les données des bulletins d'inscription PDF vers des fichiers d'import Ammon Campus (Entreprises et Stagiaires).

## 🚀 Installation Rapide

1. **Cloner le projet** dans `~/Documents/INBP/`
2. **Configurer l'environnement** :
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```
3. **Clé API** : Créer un fichier `.env` à la racine :
   ```text
   MISTRAL_API_KEY=votre_cle_ici
   ```

## 📖 Utilisation

### Mode Manuel (Terminal)
```bash
# Pour un seul fichier ou un dossier complet (BATCH)
python3 main.py -i ./input/mon_bulletin.pdf
```

### Mode Automatique (Mac)
Si vous avez configuré l'application Automator (voir [Guide Automator](GUIDE_AUTOMATOR.md)) :
1. Sélectionnez vos PDF.
2. Glissez-les sur l'icône **"Extraire Inscription INBP"**.
3. Récupérez vos fichiers Excel dans `~/Desktop/Imports_Ammon`.

## 📊 Fichiers Générés
À chaque extraction, le script génère **deux** fichiers Excel dans le dossier `output/` :
1. `Import_Entreprises_...xlsx` : Pour créer les fiches sociétés.
2. `Import_Stagiaires_...xlsx` : Pour créer les fiches personnes et les lier aux entreprises.

## ⚙️ Structure du Projet
- : Point d'entrée du script. `main.py`
- : Logique métier (Nettoyage SIRET, calcul Ref_Ext). `models.py`
- : Connexion à l'IA Mistral. `inscription_extractor.py`
- `ammon_generator_*.py` : Logique de création des fichiers Excel.

💡 _Besoin d'aide pour l'automatisation ? Consultez le [GUIDE_AUTOMATOR.md](GUIDE_AUTOMATOR.md)._
