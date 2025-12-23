import base64
import json
from pathlib import Path
from mistralai import Mistral

class InscriptionExtractor:
    """Extracteur de données depuis le bulletin d'inscription PDF"""

    def __init__(self, pdf_path, client:Mistral):
        self.pdf_path = Path(pdf_path)
        self.data = {}
        self.client = client

    def encode_file(self):
        with open(self.pdf_path, 'rb') as file:
            return base64.b64encode(file.read()).decode('utf-8')

    def call_mistral(self):
        base64_file = self.encode_file()
        return self.client.ocr.process(
        model="mistral-ocr-latest",
        pages=[0],
        document={
            "type": "document_url",
            "document_url": f"data:application/pdf;base64,{base64_file}"
        },
        include_image_base64=False,
        extract_footer=False,
        extract_header=False,
        document_annotation_format={
            "type": "json_schema",
            "json_schema": {
                "name": "response_schema",
                "schema": {
                    "type": "object",
                    "title": "StructuredData",
                    "required": [
                        "Civilité",
                        "Nom du stagiaire",
                        "Prénom du stagiaire",
                        "Adresse du stagiaire",
                        "Code postal du stagiaire",
                        "Ville du stagiaire",
                        "Pays du stagiare",
                        "Portable du stagiaire",
                        "Email du stagiaire",
                        "Date de naissance",
                        "nom de l'entreprise",
                        "adresse de l'entreprise",
                        "Code postal",
                        "Ville",
                        "Pays",
                        "Date d'entrée dans l'entreprise",
                        "Tél",
                        "N° de SIRET",
                        "Code NAFA",
                        "Email"
                    ],
                    "properties": {
                        "Civilité": {
                            "type": "string"
                        },
                        "Nom du stagiare": {
                            "type": "string"
                        },
                        "Prénom du stagiaire": {
                            "type": "string"
                        },
                        "Afresse du stagiaire": {
                            "type": "string"
                        },
                        "Code postal du stagiaire": {
                            "type": "string"
                        },
                        "Ville du stagiaire": {
                            "type": "string"
                        },
                        "Pays du stagiaire": {
                            "type": "string"
                        },
                        "Portable du stagiaire": {
                            "type": "string"
                        },
                        "Email du stagiaire": {
                            "type": "string"
                        },
                        "Date de naissance": {
                            "type": "string"
                        },
                        "nom de l'entreprise": {
                            "type": "string"
                        },
                        "adresse de l'entreprise": {
                            "type": "string"
                        },
                        "Code postal": {
                            "type": "string"
                        },
                        "Ville": {
                            "type": "string",
                        },
                        "Pays": {
                            "type": "string",
                        },
                        "Date d'entrée dans l'entreprise": {
                            "type": "string",
                        },
                        "Tél": {
                            "type": "string",
                        },
                        "N° de SIRET": {
                            "type": "string",
                        },
                        "Code NAFA": {
                            "type": "string",
                        },
                        "Email": {
                            "type": "string",
                        },
                    }
                }
            }
        }
    )

    def extract(self):
        """Méthode principale d'extraction"""
        print(f"📄 Extraction des données de: {self.pdf_path.name}")
        ocr_response = self.call_mistral()
        print(ocr_response)
        self.data = json.loads(ocr_response.document_annotation)
        return self.data
