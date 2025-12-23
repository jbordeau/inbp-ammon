from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

from ammon_code_pays import PaysCode


class EntrepriseExcelGenerator:
    """Générateur de fichier Excel pour l'import dans Ammon Campus"""

    def __init__(self, pays_code: PaysCode =None):
        self.pays_codes = pays_code

    def create_entreprises_excel(self, data_list, output_path):
        """Crée un fichier Excel d'import avec plusieurs entreprises"""

        # Créer un nouveau workbook avec les en-têtes
        wb = Workbook()
        ws = wb.active
        ws.title = "Entreprise"

        # En-têtes selon le template Ammon
        headers = [
            'cRefExt', 'iDesactive', 'SOC_cRaisonSociale', 'SOC_cType',
            'SOC_iEstSiege', 'SOC_cCateg', 'SOC_cSIRET', 'SOC_cNACE',
            'ADR_IESTADRCOURRIER', 'ADR_cAdresseNature', 'ADR_cAdresse1',
            'ADR_cAdresse2', 'ADR_cAdresse3', 'ADR_cAdresse4', 'ADR_cCodePostal',
            'ADR_cVille', 'ADR_cPays', 'ADR_cSiteWeb', 'ADR_cTel', 'ADR_cEmail',
            'LIE_cCode', 'LIE_cLibelle', 'LIE_cRefext', 'org_cAgrementAnimateur'
        ]
        ws.append(headers)

        print(f"📊 Génération du fichier Excel avec {len(data_list)} entreprise(s)...\n")

        # Ajouter une ligne pour chaque entreprise
        for i, inscription in enumerate(data_list, 1):
            ent = inscription.entreprise  # Accès direct à l'objet entreprise

            # Préparer les données
            siret = ent.siret
            if siret:
                siret = siret.replace(' ', '')

            # Générer une référence externe unique
            date_str = datetime.now().strftime('%Y%m%d')  # Format YYYYMMDD
            ref_ext = f"INBP_{siret}_{date_str}" if siret else f"INBP_{date_str}"

            # Construire la ligne de données
            row_data = [
                ref_ext,
                0,
                ent.nom,
                'E',
                -1,
                'SGE',
                siret,
                ent.code_nafa,
                -1,
                'PR',
                ent.adresse,
                '',
                '',
                '',
                ent.code_postal,
                ent.ville,
                self.pays_codes.get_pays_code(ent.pays),
                '',
                ent.telephone,
                ent.email,
                '',                                         # LIE_cCode
                '',                                         # LIE_cLibelle
                '',                                         # LIE_cRefext
                ''                                          # org_cAgrementAnimateur
            ]

            ws.append(row_data)

            # Afficher un résumé de chaque ligne ajoutée
            entreprise_nom = ent.nom if ent.nom else 'N/A'
            siret_display = siret if siret else 'N/A'
            print(f"   {i}. {entreprise_nom} (SIRET: {siret_display})")

        # Sauvegarder le fichier
        output_file = Path(output_path)
        wb.save(output_file)

        print(f"\n💾 Fichier Excel créé: {output_file}")
        print(f"   📈 {len(data_list)} entreprise(s) dans le fichier")

        return output_file
