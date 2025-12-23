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
        for i, data in enumerate(data_list, 1):
            # Préparer les données
            siret = data.get('N° de SIRET', '')
            if siret:
                siret = siret.replace(' ', '')  # Nettoyer le SIRET

            # Générer une référence externe unique
            date_str = datetime.now().strftime('%Y%m%d')  # Format YYYYMMDD
            ref_ext = f"INBP_{siret}_{date_str}" if siret else f"INBP_{date_str}"

            # Construire la ligne de données
            row_data = [
                ref_ext,                                    # cRefExt - VERSION COURTE
                0,                                          # iDesactive
                data.get('nom de l\'entreprise', ''),            # SOC_cRaisonSociale
                'E',                                        # SOC_cType
                -1,                                         # SOC_iEstSiege
                'SGE',                                      # ← CHANGEMENT 1: était ''
                siret,                                      # SOC_cSIRET
                data.get('Code NAFA', ''),                 # SOC_cNACE
                -1,                                         # ADR_IESTADRCOURRIER
                'PR',                                       # ← CHANGEMENT 2: était 'Siège'
                data.get('adresse de l\'entreprise', ''),        # ADR_cAdresse1
                '',                                         # ADR_cAdresse2
                '',                                         # ADR_cAdresse3
                '',                                         # ADR_cAdresse4
                data.get('Code postal', ''),               # ADR_cCodePostal
                data.get('Ville', ''),                     # ADR_cVille
                self.pays_codes.get_pays_code(data.get('Pays')),      # ← CHANGEMENT 3: utiliser get_pays_code()
                '',                                         # ADR_cSiteWeb
                data.get('Tél', ''),                 # ADR_cTel
                data.get('Email', ''),                     # ADR_cEmail
                '',                                         # LIE_cCode
                '',                                         # LIE_cLibelle
                '',                                         # LIE_cRefext
                ''                                          # org_cAgrementAnimateur
            ]

            ws.append(row_data)

            # Afficher un résumé de chaque ligne ajoutée
            entreprise_nom = data.get('nom_entreprise', 'N/A')
            siret_display = siret if siret else 'N/A'
            print(f"   {i}. {entreprise_nom} (SIRET: {siret_display})")

        # Sauvegarder le fichier
        output_file = Path(output_path)
        wb.save(output_file)

        print(f"\n💾 Fichier Excel créé: {output_file}")
        print(f"   📈 {len(data_list)} entreprise(s) dans le fichier")

        return output_file
