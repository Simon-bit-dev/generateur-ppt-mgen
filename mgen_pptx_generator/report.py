from __future__ import annotations

from pathlib import Path
from typing import Any


def write_markdown_report(
    path: Path,
    *,
    support_type: str,
    yaml_path: Path,
    template_path: Path,
    output_path: Path,
    template_audit: dict[str, Any],
    validation_counts: dict[str, int],
    generation: dict[str, Any],
    reference_comparison: dict[str, Any] | None = None,
    render_status: str = "NON EXÉCUTÉ",
    overflow_status: str = "NON EXÉCUTÉ",
    warnings: list[str] | None = None,
) -> None:
    warnings = warnings or []
    ok = (
        generation["slide_count"] == generation["expected_slide_count"]
        and generation["layout_sequence_matches"]
        and render_status == "OK"
        and overflow_status == "OK"
    )
    label = "Conférence / webinaire" if support_type == "conference" else "Atelier"
    lines = [
        "# Rapport de génération PowerPoint MGEN",
        "",
        f"- **Statut global :** {'CONFORME' if ok else 'À CONTRÔLER'}",
        f"- **Profil :** {label}",
        f"- **YAML :** `{yaml_path.name}`",
        f"- **Template :** `{template_path.name}`",
        f"- **PowerPoint généré :** `{output_path.name}`",
        "",
        "## Audit de la template",
        "",
        f"- SHA-256 : `{template_audit['sha256']}`",
        f"- Masques : {template_audit['slide_master_count']}",
        f"- Dispositions : {template_audit['layout_count']}",
        f"- Slides ordinaires dans la template : {template_audit['ordinary_slide_count']}",
        "",
        "## Validation du YAML",
        "",
    ]
    if support_type == "conference":
        lines.extend([
            f"- Repères théoriques : {validation_counts['theoretical_slides']}",
            f"- Conseils pratiques : {validation_counts['practical_slides']}",
            f"- Slides Passer à l’action : {validation_counts['action_slides']}",
        ])
    else:
        lines.extend([
            f"- Apports théoriques : {validation_counts['theoretical_slides']}",
            f"- Questions Vrai ou Faux : {validation_counts['quiz_questions']}",
            f"- Corrections : {validation_counts['quiz_corrections']}",
            f"- Exercices pratiques : {validation_counts['exercise_slides']}",
        ])
    lines.extend([
        f"- Nombre total attendu : {validation_counts['total_slides']}",
        "",
        "## Génération",
        "",
        f"- Slides générées : {generation['slide_count']}",
        f"- Séquence de dispositions conforme : {'oui' if generation['layout_sequence_matches'] else 'non'}",
        f"- Rendu LibreOffice : {render_status}",
        f"- Contrôle hors-canevas : {overflow_status}",
    ])
    if reference_comparison is not None:
        lines.extend([
            "",
            "## Comparaison au PowerPoint témoin",
            "",
            f"- Nombre de slides témoin : {reference_comparison['reference_slide_count']}",
            f"- Même nombre de slides : {'oui' if reference_comparison['same_slide_count'] else 'non'}",
            f"- Même séquence de dispositions : {'oui' if reference_comparison['same_layout_sequence'] else 'non'}",
            "",
            "La comparaison porte sur la structure technique et la séquence des dispositions.",
        ])
    if warnings:
        lines.extend(["", "## Avertissements", ""] + [f"- {warning}" for warning in warnings])
    lines.extend(["", "## Séquence générée", ""])
    for index, (layout, role) in enumerate(zip(generation["layout_sequence"], generation["roles"]), start=1):
        lines.append(f"{index}. `{layout}` — {role}")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
