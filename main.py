#!/usr/bin/env python3
"""
Script d'extraction des données d'inscription depuis un PDF
et génération du fichier Excel d'import pour Ammon Campus

Usage: python3 main.py <fichier_pdf> [--output <dossier_sortie>]
"""

import sys
import argparse
from pathlib import Path
from datetime import datetime
from mistralai import Mistral
from dotenv import load_dotenv
import os

from ammon_code_pays import PaysCode
from ammon_generator_entreprise import EntrepriseExcelGenerator
from ammon_generator_personne import PersonneExcelGenerator
from inscription_extractor import InscriptionExtractor
from ammon_existants_service import ExistantsService


def main():
    parser = argparse.ArgumentParser(
        description='Extrait les données d\'inscription depuis des PDFs et génère un fichier Excel pour Ammon Campus'
    )
    parser.add_argument('--pdf_input', '-i', default='./input', help='Chemin vers un fichier PDF ou un dossier contenant des PDFs')
    parser.add_argument('--output', '-o', default='./output', help='Dossier de sortie pour le fichier Excel')
    parser.add_argument('--template', '-t', default='./Template_Import_Entreprises.xlsx', help='Chemin vers le template Excel Ammon')

    args = parser.parse_args()

    # Déterminer s'il s'agit d'un fichier ou d'un dossier
    input_path = Path(args.pdf_input)

    if not input_path.exists():
        print(f"❌ Erreur: Le chemin {input_path} n'existe pas")
        sys.exit(1)

    # Collecter les fichiers PDF à traiter
    pdf_files = []
    if input_path.is_file():
        if input_path.suffix.lower() == '.pdf':
            pdf_files = [input_path]
        else:
            print(f"❌ Erreur: {input_path} n'est pas un fichier PDF")
            sys.exit(1)
    elif input_path.is_dir():
        pdf_files = sorted(input_path.glob('*.pdf'))
        if not pdf_files:
            print(f"❌ Erreur: Aucun fichier PDF trouvé dans {input_path}")
            sys.exit(1)

    print(f"📁 {len(pdf_files)} fichier(s) PDF à traiter\n")

    load_dotenv()
    api_key = os.getenv('MISTRAL_API_KEY')
    client = Mistral(api_key=api_key)

    existants = ExistantsService(folder_path="./existants")

    # Extraire les données de tous les PDFs
    all_data = []
    for pdf_file in pdf_files:
        print(f"📄 Traitement: {pdf_file.name}")
        try:
            extractor = InscriptionExtractor(pdf_file,client=client)
            inscription = extractor.extract()
            ent = inscription.entreprise
            ent.display_summary()

            # 2. Vérification des doublons
            existing_ref = existants.get_existing_ref(ent.siret)
            if existing_ref:
                print(f"   🚫 L'entreprise existe déjà dans Ammon (Ref: {existing_ref}). Ignorée.")
                continue

            all_data.append(inscription)
        except Exception as e:
            print(f"   ❌ Erreur lors du traitement: {e}")
            continue

    if not all_data:
        print("\n❌ Aucune donnée n'a pu être extraite")
        sys.exit(1)

    print(f"\n✅ {len(all_data)} inscription(s) extraite(s) avec succès\n")

    # Génération du fichier Excel avec toutes les données
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Nom du fichier de sortie
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    ent_output = output_dir / f"Import_Entreprise_{timestamp}.xlsx"
    personne_output = output_dir / f"Import_Stagiaires_{timestamp}.xlsx"

    pays_code = PaysCode(template_path=args.template)

    # Génération Entreprises
    generator = EntrepriseExcelGenerator(pays_code=pays_code)
    generator.create_entreprises_excel(all_data, ent_output)

    # Génération Stagiaires
    stg_gen = PersonneExcelGenerator(pays_code=pays_code)
    stg_gen.create_personnes_excel(all_data, personne_output)


print("✨ Traitement terminé avec succès!")


if __name__ == '__main__':
    main()
