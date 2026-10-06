"""CLI entry point for PPT formatting agent."""

import argparse
import json
import os

from app.ppt_loader import PPTXLoader


def main() -> None:
    parser = argparse.ArgumentParser(description="Parse an existing PowerPoint and export its structure.")
    parser.add_argument("pptx_path", nargs="?", default="202608 - COPIL Métiers - SI AM  - VF.pptx", help="Path to the input .pptx file")
    parser.add_argument("--output", default="output/extracted_structure.json", help="Path to output JSON summary")
    args = parser.parse_args()

    if not os.path.exists(args.pptx_path):
        raise FileNotFoundError(f"Le fichier PPTX n'existe pas : {args.pptx_path}")

    output_path = PPTXLoader.save_json(args.pptx_path, args.output)
    print(f"Fichier JSON exporté : {output_path}")


if __name__ == "__main__":
    main()
