from dataclasses import dataclass, field
from datetime import datetime


@dataclass
class Entreprise:
    nom: str
    adresse: str
    code_postal: str
    ville: str
    pays: str
    siret: str
    code_nafa: str
    telephone: str
    email: str
    date_entree: str
    # field(init=False) indique que ce n'est pas un argument du constructeur
    ref_ext: str = field(init=False)

    def __post_init__(self):
        # 1. Nettoyage du SIRET (on enlève les espaces)
        if self.siret:
            self.siret = self.siret.replace(' ', '')

        # 2. Génération de la référence externe
        date_str = datetime.now().strftime('%Y%m%d')
        self.ref_ext = f"INBP_{self.siret}_{date_str}" if self.siret else f"INBP_{date_str}"

    @property
    def is_valid(self) -> bool:
        """Vérifie si les données vitales sont présentes"""
        return bool(self.nom and self.siret)

    def display_summary(self):
        """Affiche un résumé de l'état de l'entreprise dans la console"""
        if not self.nom:
            print("   ⚠️  Nom de l'entreprise non trouvé")
        if not self.siret:
            print("   ⚠️  SIRET non trouvé")

        if self.is_valid:
            print(f"   ✅ {self.nom} - SIRET: {self.siret}")


@dataclass
class Stagiaire:
    civilite: str
    nom: str
    prenom: str
    adresse: str
    code_postal: str
    ville: str
    pays: str
    portable: str
    email: str
    date_naissance: str


@dataclass
class Inscription:
    entreprise: Entreprise
    stagiaire: Stagiaire
