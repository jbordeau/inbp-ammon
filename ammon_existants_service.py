from typing import Any

import pandas as pd
import re
from pathlib import Path

class ExistantsService:
    def __init__(self, folder_path="./existants"):
        self.folder_path = Path(folder_path)
        self.siret_to_ref = {}  # Map: {SIRET: RefExt}
        self._load_existants()

    def _load_existants(self):
        """Charge le fichier \d+-VIE_ENTREPRISE.xls en mémoire"""
        if not self.folder_path.exists():
            print(f"⚠️ Dossier {self.folder_path} non trouvé. Aucune vérification de doublons possible.")
            return

        # Recherche du fichier par regex
        pattern = re.compile(r"\d+-VIE_ENTREPRISE\.xls$")
        target_file = None
        for f in self.folder_path.glob("*.xls"):
            if pattern.match(f.name):
                target_file = f
                break

        if not target_file:
            print("⚠️ Aucun fichier d'export entreprises trouvé dans /existants (format attendu: XXXXX-VIE_ENTREPRISE.xls)")
            return

        print(f"🔍 Chargement des entreprises existantes depuis {target_file.name}...")
        try:
            # Lecture du XLS via pandas
            df = pd.read_excel(target_file)
            
            # On nettoie et on mappe (SIRET -> RefExt)
            # Adaptes les noms de colonnes ici si elles diffèrent dans l'export Ammon
            if 'SOC_cSIRET' in df.columns and 'cRefExt' in df.columns:
                # On enlève les espaces des SIRET et on convertit en string
                df['SOC_cSIRET'] = df['SOC_cSIRET'].astype(str).str.replace(r'\s+', '', regex=True)
                # Création du dictionnaire
                self.siret_to_ref = pd.Series(df.cRefExt.values, index=df.SOC_cSIRET).to_dict()
                print(f"✅ {len(self.siret_to_ref)} entreprises chargées en mémoire.")
            else:
                print(f"❌ Colonnes 'SOC_cSIRET' ou 'cRefExt' manquantes dans le fichier {target_file.name}")
        except Exception as e:
            print(f"❌ Erreur lors de la lecture du fichier existants: {e}")

    def get_existing_ref(self, siret: str) -> str | None:
        """Retourne la RefExt si le SIRET existe, sinon None"""
        if not siret:
            return None
        clean_siret = str(siret).replace(' ', '')
        return self.siret_to_ref.get(clean_siret)
