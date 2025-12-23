from dataclasses import dataclass

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
