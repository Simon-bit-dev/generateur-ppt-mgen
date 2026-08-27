from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .errors import MgenGeneratorError
from .portable_service import generate_powerpoint


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Génère un PowerPoint MGEN depuis un YAML Conférence ou Atelier validé.")
    parser.add_argument("--yaml", required=True, type=Path, help="YAML final Conférence ou Atelier")
    parser.add_argument("--output", required=True, type=Path, help="PowerPoint de sortie")
    parser.add_argument("--report", type=Path, help="Rapport Markdown de contrôle")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        result = generate_powerpoint(args.yaml, args.output, args.report)
        print(f"Profil détecté: {result.summary.support_type}")
        print(f"PowerPoint généré: {result.output_path}")
        print(f"Rapport généré: {result.report_path}")
        return 0
    except MgenGeneratorError as exc:
        print(f"ERREUR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
