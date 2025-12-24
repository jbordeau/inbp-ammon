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
        self.ref_ext = f"ENTRE_{self.siret}_{date_str}" if self.siret else f"ENTRE_{date_str}"

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
    ref_ext: str = field(init=False)

    def __post_init__(self):
        # 1. Nettoyage du portable (enlève espaces et points)
        if self.portable:
            self.portable = str(self.portable).replace(' ', '').replace('.', '')

        # 2. Génération de la référence externe unique
        timestamp = datetime.now().strftime('%Y%m%d')
        nom_clean = self.nom.replace(' ', '')[:5].upper()
        self.ref_ext = f"PERS_{nom_clean}_{timestamp}"

    @property
    def is_valid(self) -> bool:
        """Vérifie si les données vitales sont présentes"""
        return bool(self.nom and self.prenom)

    @property
    def sexe(self) -> str:
        """Déduit le sexe à partir de la civilité"""
        c = str(self.civilite).upper()
        if 'MME' in c or 'MLLE' in c:
            return 'F'
        return 'M'

    @property
    def civilite_ammon(self) -> str:
        """Normalise la civilité pour Ammon (M. ou MME)"""
        c = str(self.civilite).upper()
        if 'MME' in c or 'MLLE' in c:
            return 'MME'
        return 'M.'

@dataclass
class Inscription:
    entreprise: Entreprise
    stagiaire: Stagiaire
