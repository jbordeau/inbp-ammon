from datetime import datetime
from pathlib import Path
from openpyxl import Workbook
from ammon_code_pays import PaysCode

class PersonneExcelGenerator:
    """Générateur de fichier Excel pour l'import des stagiaires dans Ammon Campus"""

    def __init__(self, pays_code: PaysCode = None):
        self.pays_codes = pays_code

    def create_personnes_excel(self, data_list, output_path):
        """Crée un fichier Excel d'import pour les stagiaires"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Personnes"

        # En-têtes basés sur Import_Personnes.csv
        headers = [
            'cRefExt', 'SOC_cRefExt', 'SOC_cRefExtService', 'PER_CCODESTRUCTURE_RATTACHEMENT',
            'PER_cNumeroSS', 'PER_cCivilite', 'PER_cSexe', 'PER_cNom', 'PER_cPrenom',
            'PER_cNomJeun', 'PER_xDateNaiss', 'PER_cCommuneNaiss', 'PER_cDeptNaiss',
            'PER_cPaysNaiss', 'PER_cSitFam', 'PER_cNationalite', 'PER_cNIvForm',
            'PER_cCateg', 'iDesactive', 'PER_BNPAI_MAIL', 'PER_BBL_COMM_MAIL',
            'psl_cTel', 'psl_cTelPort', 'psl_cEmail', 'ADR_cAdresseNature',
            'ADR_cAdresse1', 'ADR_cAdresse2', 'ADR_cAdresse3', 'ADR_cAdresse4',
            'ADR_cCodePostal', 'ADR_cVille', 'ADR_cPays', 'ADR_cSiteWeb',
            'ADR_CTEL', 'ADR_CTELPORT', 'ADR_CEMAIL'
        ]
        ws.append(headers)

        print(f"👤 Génération du fichier Excel avec {len(data_list)} stagiaire(s)...\n")

        for i, inscription in enumerate(data_list, 1):
            stg = inscription.stagiaire
            ent = inscription.entreprise

            row_data = [
                stg.ref_ext,                    # cRefExt
                ent.ref_ext,                    # SOC_cRefExt (Lien avec l'entreprise)
                '',                             # SOC_cRefExtService
                'CLIENT',                       # PER_CCODESTRUCTURE_RATTACHEMENT
                '',                             # PER_cNumeroSS
                stg.civilite_ammon,             # PER_cCivilite
                stg.sexe,                       # PER_cSexe
                stg.nom.upper(),                # PER_cNom
                stg.prenom,                     # PER_cPrenom
                '',                             # PER_cNomJeun
                stg.date_naissance,             # PER_xDateNaiss
                '',                             # PER_cCommuneNaiss
                '',                             # PER_cDeptNaiss
                self.pays_codes.get_pays_code(stg.pays), # PER_cPaysNaiss
                '',                             # PER_cSitFam
                self.pays_codes.get_pays_code(stg.pays), # PER_cNationalite
                '',                             # PER_cNIvForm
                'INT,STA',                      # PER_cCateg
                0,                              # iDesactive
                0,                              # PER_BNPAI_MAIL
                0,                              # PER_BBL_COMM_MAIL
                '',                             # psl_cTel
                stg.portable,                   # psl_cTelPort (Pro)
                stg.email,                      # psl_cEmail (Pro)
                'PERS',                         # ADR_cAdresseNature (Personnelle)
                stg.adresse,                    # ADR_cAdresse1
                '',                             # ADR_cAdresse2
                '',                             # ADR_cAdresse3
                '',                             # ADR_cAdresse4
                stg.code_postal,                # ADR_cCodePostal
                stg.ville,                      # ADR_cVille
                self.pays_codes.get_pays_code(stg.pays), # ADR_cPays
                '',                             # ADR_cSiteWeb
                '',                             # ADR_CTEL
                stg.portable,                   # ADR_CTELPORT
                stg.email                       # ADR_CEMAIL
            ]
            ws.append(row_data)
            print(f"   {i}. {stg.prenom} {stg.nom.upper()}")

        output_file = Path(output_path)
        wb.save(output_file)
        print(f"\n💾 Fichier Excel Stagiaires créé: {output_file}")
        return output_file
