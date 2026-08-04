from __future__ import annotations

import os
import re
import sys
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .atelier import audit_atelier_template, build_atelier_specs, validate_atelier_yaml
from .conference import build_conference_specs
from .engine import create_presentation, inspect_generated
from .errors import MgenGeneratorError, ValidationError
from .profiles import detect_support_type
from .validation import audit_template as audit_conference_template
from .validation import load_yaml, validate_semantic_yaml


@dataclass(frozen=True)
class YamlSummary:
    support_type: str
    support_label: str
    title: str
    subtitle: str
    date: str
    architecture: str
    expected_slides: int


@dataclass(frozen=True)
class GenerationResult:
    summary: YamlSummary
    output_path: Path
    report_path: Path
    template_sha256: str
    generated_slides: int


def application_root() -> Path:
    """Return resources root in source mode and in a PyInstaller bundle."""
    if hasattr(sys, "_MEIPASS"):
        return Path(getattr(sys, "_MEIPASS"))
    return Path(__file__).resolve().parent.parent


def resource_path(*parts: str) -> Path:
    return application_root().joinpath(*parts)


def _atelier_title(data: dict[str, Any]) -> str:
    footer = str(data.get("VARIABLES_GLOBALES", {}).get("GLOBAL_FOOTER_LEFT", "")).strip()
    if ":" in footer:
        return footer.split(":", 1)[1].strip() or footer
    return footer or "Atelier MGEN"


def inspect_yaml(yaml_path: Path) -> YamlSummary:
    data = load_yaml(yaml_path)
    try:
        support_type = detect_support_type(data)
    except ValueError as exc:
        raise ValidationError(str(exc)) from exc

    if support_type == "conference":
        counts = validate_semantic_yaml(data)
        presentation = data["presentation"]
        label = "Conférence" if presentation["format"] == "CONFÉRENCE" else (
            "Webinaire" if presentation["format"] == "WEBINAIRE" else "Format à préciser"
        )
        return YamlSummary(
            support_type="conference",
            support_label=label,
            title=str(presentation["title"]),
            subtitle=str(presentation.get("subtitle", "")),
            date=str(presentation["date"]),
            architecture=str(presentation["architecture_variant"]),
            expected_slides=counts["total_slides"],
        )

    schema_path = resource_path("resources", "schemas", "schema_atelier_mgen.yaml")
    schema = load_yaml(schema_path)
    counts = validate_atelier_yaml(data, schema)
    variables = data["VARIABLES_GLOBALES"]
    return YamlSummary(
        support_type="atelier",
        support_label="Atelier",
        title=_atelier_title(data),
        subtitle="",
        date=str(variables["GLOBAL_FOOTER_RIGHT"]),
        architecture="ATELIER_STANDARD_33",
        expected_slides=counts["total_slides"],
    )


def _safe_stem(text: str) -> str:
    normalized = unicodedata.normalize("NFKD", text)
    ascii_text = normalized.encode("ascii", "ignore").decode("ascii")
    ascii_text = re.sub(r"[^A-Za-z0-9]+", "_", ascii_text).strip("_")
    return ascii_text[:100] or "Support_MGEN"


def default_output_name(summary: YamlSummary) -> str:
    return f"{_safe_stem(summary.title)}_MGEN_FINAL.pptx"


def _write_portable_report(
    path: Path,
    *,
    summary: YamlSummary,
    yaml_path: Path,
    template_path: Path,
    output_path: Path,
    template_audit: dict[str, Any],
    counts: dict[str, int],
    generation: dict[str, Any],
) -> None:
    conform = (
        generation["slide_count"] == generation["expected_slide_count"]
        and generation["layout_sequence_matches"]
    )
    lines = [
        "# Rapport de génération PowerPoint MGEN",
        "",
        f"- **Statut technique :** {'CONFORME' if conform else 'NON CONFORME'}",
        "- **Contrôle visuel :** à effectuer dans PowerPoint avant diffusion",
        f"- **Profil :** {summary.support_label}",
        f"- **Titre :** {summary.title}",
        f"- **YAML :** `{yaml_path.name}`",
        f"- **Template intégrée :** `{template_path.name}`",
        f"- **PowerPoint généré :** `{output_path.name}`",
        "",
        "## Contrôles automatiques",
        "",
        f"- Empreinte SHA-256 de la template : `{template_audit['sha256']}`",
        f"- Masques : {template_audit['slide_master_count']}",
        f"- Dispositions : {template_audit['layout_count']}",
        f"- Slides ordinaires dans la template : {template_audit['ordinary_slide_count']}",
        f"- Slides attendues : {counts['total_slides']}",
        f"- Slides générées : {generation['slide_count']}",
        f"- Séquence des dispositions conforme : {'oui' if generation['layout_sequence_matches'] else 'non'}",
        "",
        "## Contrôle humain avant diffusion",
        "",
        "Ouvrir le fichier dans PowerPoint et vérifier rapidement :",
        "- l'absence de débordement de texte ;",
        "- la lisibilité des puces et numérotations ;",
        "- les retours à la ligne ;",
        "- la présence des éléments institutionnels fixes ;",
        "- le nom du support et la date.",
        "",
        "## Séquence générée",
        "",
    ]
    for index, (layout, role) in enumerate(zip(generation["layout_sequence"], generation["roles"]), start=1):
        lines.append(f"{index}. `{layout}` — {role}")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def generate_powerpoint(yaml_path: Path, output_path: Path) -> GenerationResult:
    yaml_path = yaml_path.resolve()
    output_path = output_path.resolve()
    data = load_yaml(yaml_path)
    try:
        support_type = detect_support_type(data)
    except ValueError as exc:
        raise ValidationError(str(exc)) from exc

    summary = inspect_yaml(yaml_path)
    if support_type == "conference":
        template_path = resource_path("resources", "templates", "conference_mgen.pptx")
        counts = validate_semantic_yaml(data)
        template_audit = audit_conference_template(template_path, data)
        specs = build_conference_specs(data)
    else:
        template_path = resource_path("resources", "templates", "atelier_mgen.pptx")
        schema_path = resource_path("resources", "schemas", "schema_atelier_mgen.yaml")
        schema = load_yaml(schema_path)
        counts = validate_atelier_yaml(data, schema)
        template_audit = audit_atelier_template(template_path, data, schema)
        specs = build_atelier_specs(data)

    if len(specs) != counts["total_slides"]:
        raise MgenGeneratorError(
            f"Le moteur a construit {len(specs)} slides au lieu de {counts['total_slides']}."
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs = create_presentation(template_path, specs, output_path)
    generation = inspect_generated(prs, specs)
    if generation["slide_count"] != generation["expected_slide_count"]:
        raise MgenGeneratorError("Le nombre de slides générées est incohérent.")
    if not generation["layout_sequence_matches"]:
        raise MgenGeneratorError("La séquence des dispositions générées est incohérente.")

    report_path = output_path.with_name(f"{output_path.stem}_RAPPORT.md")
    _write_portable_report(
        report_path,
        summary=summary,
        yaml_path=yaml_path,
        template_path=template_path,
        output_path=output_path,
        template_audit=template_audit,
        counts=counts,
        generation=generation,
    )
    return GenerationResult(
        summary=summary,
        output_path=output_path,
        report_path=report_path,
        template_sha256=str(template_audit["sha256"]),
        generated_slides=generation["slide_count"],
    )


def open_path(path: Path) -> None:
    path = path.resolve()
    if sys.platform.startswith("win"):
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        import subprocess
        subprocess.Popen(["open", str(path)])
    else:
        import subprocess
        subprocess.Popen(["xdg-open", str(path)])
