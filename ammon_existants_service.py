from typing import Any

import pandas as pd
import re
from pathlib import Path

class ExistantsService:
    def __init__(self, folder_path="./existants"):
        self.folder_path = Path(folder_path)
        self.siret_to_ref = {}      # Map: {SIRET: RefExt}
        self.nom_prenom_to_ref = {} # Map: {NOM+PRENOM: RefExt}
        self._load_existants()

    def _load_existants(self):
        """Charge les fichiers d'export Ammon en mémoire"""
        if not self.folder_path.exists():
            print(f"⚠️ Dossier {self.folder_path} non trouvé.")
            return

        # 1. Chargement Entreprises
        self._load_file(r"\d+-VIE_ENTREPRISE\.xls$", self._process_entreprises)

        # 2. Chargement Personnes
        self._load_file(r"\d+-VIE_PERSONNE\.xls$", self._process_personnes)

    def _load_file(self, regex_pattern, process_func):
        pattern = re.compile(regex_pattern)
        for f in self.folder_path.glob("*.xls"):
            if pattern.match(f.name):
                try:
                    df = pd.read_excel(f)
                    process_func(df, f.name)
                except Exception as e:
                    print(f"❌ Erreur lecture {f.name}: {e}")
                return

    def _process_entreprises(self, df, filename):
        if 'SOC_cSIRET' in df.columns and 'cRefExt' in df.columns:
            df['SOC_cSIRET'] = df['SOC_cSIRET'].astype(str).str.replace(r'\s+', '', regex=True)
            self.siret_to_ref = pd.Series(df.cRefExt.values, index=df.SOC_cSIRET).to_dict()
            print(f"✅ {len(self.siret_to_ref)} entreprises chargées ({filename})")

    def _process_personnes(self, df, filename):
        if all(col in df.columns for col in ['PER_cNom', 'PER_cPrenom', 'cRefExt']):
            # Création d'une clé normalisée : NOM+PRENOM en majuscules, sans espaces
            df['key'] = (df['PER_cNom'].astype(str) + df['PER_cPrenom'].astype(str)).str.replace(r'\s+', '', regex=True).str.upper()
            self.nom_prenom_to_ref = pd.Series(df.cRefExt.values, index=df.key).to_dict()
            print(f"✅ {len(self.nom_prenom_to_ref)} personnes chargées ({filename})")

    def get_existing_entreprise_ref(self, siret: str) -> str | None:
        if not siret: return None
        return self.siret_to_ref.get(str(siret).replace(' ', ''))

    def get_existing_personne_ref(self, nom: str, prenom: str) -> str | None:
        if not nom or not prenom: return None
        # Même normalisation qu'au chargement
        key = (str(nom) + str(prenom)).replace(' ', '').upper()
        return self.nom_prenom_to_ref.get(key)