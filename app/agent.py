"""Agent principal pour charger, analyser et reformater un PowerPoint."""

import json
import os

from app.ppt_loader import PPTXLoader
from app.ppt_builder import PPTBuilder


class PPTFormattingAgent:
    """Agent pour mettre en forme une présentation PowerPoint."""

    def __init__(self, input_pptx: str, output_pptx: str, style_config: dict = None):
        """Initialise l'agent.
        
        Args:
            input_pptx: Chemin du fichier PowerPoint d'entrée
            output_pptx: Chemin du fichier PowerPoint de sortie
            style_config: Configuration de style personnalisée (optionnel)
        """
        self.input_pptx = input_pptx
        self.output_pptx = output_pptx
        self.style_config = style_config or {
            "title_font_size": 44,
            "subtitle_font_size": 28,
            "body_font_size": 18,
            "primary_color": "0A5F9A",
            "accent_color": "E86A33",
            "background_color": "F7F9FB",
            "font_family": "Calibri",
        }

    def run(self) -> bool:
        """Lance le processus complet de reformatage.
        
        Returns:
            True si succès, False sinon
        """
        try:
            print("\n" + "=" * 60)
            print("🚀 AGENT DE FORMATAGE POWERPOINT")
            print("=" * 60 + "\n")

            print(f"📂 Chargement du fichier : {self.input_pptx}")
            loader = PPTXLoader(self.input_pptx)
            print("✓ Fichier chargé avec succès")

            print("\n🔍 Analyse du contenu...")
            summary = loader.get_summary()
            print(f"✓ {summary['total_slides']} slides détectées")

            print("\n🛠️  Transformation en format structuré...")
            structured_data = loader.export_to_dict()
            print("✓ Contenu structuré")

            print("\n🎨 Génération de la présentation reformattée...")
            builder = PPTBuilder(self.style_config)
            builder.build_from_config(structured_data["slides"])
            print("✓ Présentation construite")

            print("\n💾 Sauvegarde...")
            builder.save(self.output_pptx)

            print("\n" + "=" * 60)
            print("✅ SUCCÈS !")
            print("=" * 60)
            print(f"Fichier de sortie : {self.output_pptx}")
            print(f"Nombre de slides : {summary['total_slides']}")
            print("=" * 60 + "\n")
            return True

        except FileNotFoundError as e:
            print(f"\n❌ Erreur : {e}")
            return False
        except Exception as e:
            print(f"\n❌ Erreur inattendue : {e}")
            import traceback
            traceback.print_exc()
            return False

    def save_structure_to_json(self, json_output: str = "output/extracted_structure.json"):
        """Exporte la structure extraite en JSON.
        
        Args:
            json_output: Chemin du fichier JSON de sortie
        """
        try:
            loader = PPTXLoader(self.input_pptx)
            structured_data = loader.export_to_dict()

            os.makedirs(os.path.dirname(json_output) or ".", exist_ok=True)
            with open(json_output, "w", encoding="utf-8") as f:
                json.dump(structured_data, f, ensure_ascii=False, indent=2)

            print(f"✓ Structure JSON exportée : {json_output}")
        except Exception as e:
            print(f"❌ Erreur lors de l'export JSON : {e}")
