from __future__ import annotations

import hashlib
import re
from pathlib import Path
from typing import Any

from pptx import Presentation

from .errors import ValidationError
from .models import PlaceholderValue as PH
from .models import SlideSpec

ATELIER_ROOTS = {"VARIABLES_GLOBALES", "SLIDES", "ANOMALIES_TECHNIQUES"}
SCHEMA_ROOTS = {"VARIABLES_GLOBALES", "SLIDES"}
_VARIABLE_PATTERN = re.compile(r"\{\{([A-Z0-9_]+)\}\}")
_UNFILLED_MARKERS = ("__", "À RENSEIGNER", "À RENSEIGNER OU LAISSER VIDE")

EXPECTED_LAYOUTS = [
    "couverture",
    "Bienvenue!",
    "intervenants",
    "Accueil et objectifs",
    "sommaire",
    "liste objectifs",
    "brise glace",
    "transition",
    "apports théoriques",
    "VraiFaux transition",
    "VraiFaux?",
    "Faux !",
    "Vrai!",
    "exercices pratiques",
    "synthèse",
    "pour aller plus loin",
    "Vivoptim",
    "4e de couverture",
]

EXPECTED_PLACEHOLDERS: dict[str, set[int]] = {
    "couverture": {23, 24, 25, 26},
    "Bienvenue!": {21, 22},
    "intervenants": {2, 21, 25, 26},
    "Accueil et objectifs": {24, 25},
    "sommaire": {2, 25, 26},
    "liste objectifs": {2, 24, 25, 26},
    "brise glace": {2, 24, 25, 26, 27},
    "transition": {21, 23, 24},
    "apports théoriques": {2, 25, 26, 27, 28},
    "VraiFaux transition": {21},
    "VraiFaux?": {2, 21, 25, 26},
    "Faux !": {2, 25, 26, 27, 28},
    "Vrai!": {2, 21, 25, 26, 27},
    "exercices pratiques": {2, 25, 26, 27, 28},
    "synthèse": {2, 25, 26, 27},
    "pour aller plus loin": {2, 25, 26, 27},
    "Vivoptim": {2, 25, 26},
    "4e de couverture": {2, 26},
}

CORRECTION_RULES = {
    "MGEN-12": ("Faux !", {"PH25", "PH26", "PH27", "PH28"}),
    "MGEN-13": ("Vrai!", {"PH21", "PH25", "PH26", "PH27"}),
}


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise ValidationError(message)


def _contains_unfilled(value: Any) -> bool:
    if isinstance(value, str):
        return any(marker in value for marker in _UNFILLED_MARKERS)
    if isinstance(value, list):
        return any(_contains_unfilled(item) for item in value)
    if isinstance(value, dict):
        return any(_contains_unfilled(item) for item in value.values())
    return False


def _parse_ph_key(key: str) -> int:
    match = re.fullmatch(r"PH(\d+)", key)
    if not match:
        raise ValidationError(f"Clé de placeholder invalide: {key!r}.")
    return int(match.group(1))


def _resolve_variables(value: str, variables: dict[str, str]) -> str:
    def replace(match: re.Match[str]) -> str:
        name = match.group(1)
        if name not in variables:
            raise ValidationError(f"Variable globale inconnue: {name}.")
        return variables[name]

    resolved = _VARIABLE_PATTERN.sub(replace, value)
    if "{{" in resolved or "}}" in resolved:
        raise ValidationError(f"Variable globale mal formée ou non résolue dans: {value!r}.")
    return resolved


def _normalize_roots(data: dict[str, Any]) -> set[str]:
    return set(data.keys())


def validate_atelier_yaml(data: dict[str, Any], schema: dict[str, Any]) -> dict[str, int]:
    _require(_normalize_roots(data) == ATELIER_ROOTS, "Le YAML Atelier doit conserver exactement les trois clés racines attendues.")
    _require(set(schema.keys()).issuperset(SCHEMA_ROOTS), "Le schéma Atelier doit contenir VARIABLES_GLOBALES et SLIDES.")
    _require(not _contains_unfilled(data), "Le YAML Atelier contient encore un marqueur de remplissage.")

    variables = data["VARIABLES_GLOBALES"]
    _require(set(variables) == {"GLOBAL_FOOTER_LEFT", "GLOBAL_FOOTER_RIGHT"}, "Les deux variables globales Atelier doivent être conservées exactement.")
    _require(all(isinstance(value, str) and value.strip() for value in variables.values()), "Les variables globales Atelier doivent être des chaînes non vides.")

    anomalies = data["ANOMALIES_TECHNIQUES"]
    _require(isinstance(anomalies, dict) and anomalies.get("statut_global") == "CONFORME", "Le YAML Atelier doit avoir un statut technique CONFORME.")
    _require(anomalies.get("anomalies") == [], "Le YAML Atelier contient des anomalies techniques déclarées.")

    slides = data["SLIDES"]
    schema_slides = schema["SLIDES"]
    _require(isinstance(slides, list) and len(slides) == 33, "Le profil Atelier standard doit contenir exactement 33 slides.")
    _require(isinstance(schema_slides, list) and len(schema_slides) == 33, "Le schéma Atelier fourni doit décrire exactement 33 slides.")

    layout_counts: dict[str, int] = {}
    for index, (slide, schema_slide) in enumerate(zip(slides, schema_slides), start=1):
        expected_id = f"S{index:02d}"
        _require(slide.get("slide_id") == expected_id, f"Slide {index}: slide_id attendu {expected_id}.")
        _require(slide.get("numero") == index, f"Slide {expected_id}: numero attendu {index}.")
        _require(slide.get("slide_id") == schema_slide.get("slide_id"), f"Slide {expected_id}: divergence avec le schéma.")
        _require(slide.get("numero") == schema_slide.get("numero"), f"Slide {expected_id}: numero divergent du schéma.")
        _require(slide.get("partie") == schema_slide.get("partie"), f"Slide {expected_id}: partie divergente du schéma.")
        _require(slide.get("statut") in {"variable", "fixe — conserver intégralement"}, f"Slide {expected_id}: statut invalide.")

        code = slide.get("layout_code")
        name = slide.get("layout_name_exact")
        schema_code = schema_slide.get("layout_code")
        schema_name = schema_slide.get("layout_name_exact")
        placeholders = slide.get("placeholders")
        _require(isinstance(placeholders, dict), f"Slide {expected_id}: placeholders doit être un objet.")

        if isinstance(schema_code, str) and schema_code.startswith("__MGEN_12"):
            _require(code in CORRECTION_RULES, f"Slide {expected_id}: correction attendue en MGEN-12 ou MGEN-13.")
            expected_name, expected_keys = CORRECTION_RULES[code]
            _require(name == expected_name, f"Slide {expected_id}: nom de correction incompatible avec {code}.")
            _require(set(placeholders) == expected_keys, f"Slide {expected_id}: placeholders de correction incompatibles avec {code}.")
        else:
            _require(code == schema_code, f"Slide {expected_id}: layout_code divergent du schéma.")
            _require(name == schema_name, f"Slide {expected_id}: layout_name_exact divergent du schéma.")
            _require(set(placeholders) == set(schema_slide.get("placeholders", {})), f"Slide {expected_id}: clés de placeholders divergentes du schéma.")

        _require(name in EXPECTED_LAYOUTS, f"Slide {expected_id}: disposition Atelier inconnue {name!r}.")
        allowed_idxs = EXPECTED_PLACEHOLDERS[name]
        for key, value in placeholders.items():
            idx = _parse_ph_key(key)
            _require(idx in allowed_idxs, f"Slide {expected_id}: {key} n’existe pas sur la disposition {name}.")
            _require(isinstance(value, str), f"Slide {expected_id}: {key} doit contenir une chaîne.")
            _resolve_variables(value, variables)

        layout_counts[name] = layout_counts.get(name, 0) + 1

    _require(layout_counts.get("apports théoriques") == 8, "L’atelier standard doit contenir exactement 8 slides d’apports théoriques.")
    _require(layout_counts.get("VraiFaux?") == 3, "L’atelier standard doit contenir exactement 3 questions Vrai ou Faux.")
    _require(layout_counts.get("Faux !", 0) + layout_counts.get("Vrai!", 0) == 3, "L’atelier standard doit contenir exactement 3 corrections.")
    _require(layout_counts.get("exercices pratiques") == 4, "L’atelier standard doit contenir exactement 4 exercices pratiques.")
    _require(slides[-3]["layout_name_exact"] == "pour aller plus loin", "La slide Pour aller plus loin doit précéder Vivoptim.")
    _require(slides[-2]["layout_name_exact"] == "Vivoptim", "Vivoptim doit être l’avant-dernière séquence institutionnelle.")
    _require(slides[-1]["layout_name_exact"] == "4e de couverture", "La quatrième de couverture doit être la dernière slide.")

    return {
        "total_slides": 33,
        "theoretical_slides": 8,
        "quiz_questions": 3,
        "quiz_corrections": 3,
        "exercise_slides": 4,
    }


def audit_atelier_template(path: Path, data: dict[str, Any], schema: dict[str, Any]) -> dict[str, Any]:
    del data, schema  # The YAML family has no embedded template fingerprint in this version.
    actual_sha = sha256_file(path)
    prs = Presentation(str(path))
    _require(len(prs.slide_masters) == 1, "La template Atelier doit contenir exactement un masque.")
    _require(len(prs.slide_layouts) == 18, "La template Atelier doit contenir exactement 18 dispositions.")
    _require(len(prs.slides) == 0, "La template Atelier doit être dépourvue de slides ordinaires.")

    names = [layout.name for layout in prs.slide_layouts]
    _require(names == EXPECTED_LAYOUTS, "Les noms ou l’ordre des 18 dispositions Atelier diffèrent du registre attendu.")
    _require(len(set(names)) == len(names), "Les noms de dispositions Atelier doivent être uniques.")

    for layout in prs.slide_layouts:
        actual = {shape.placeholder_format.idx for shape in layout.shapes if shape.is_placeholder}
        expected = EXPECTED_PLACEHOLDERS[layout.name]
        _require(actual == expected, f"Disposition {layout.name}: placeholders attendus {sorted(expected)}, obtenus {sorted(actual)}.")

    objectives_layout = prs.slide_layouts[names.index("liste objectifs")]
    fixed_texts = [shape.text for shape in objectives_layout.shapes if hasattr(shape, "text") and shape.text]
    _require("Les objectifs de l’atelier" in fixed_texts, "Le titre fixe de la disposition Objectifs Atelier est non conforme.")

    return {
        "sha256": actual_sha,
        "slide_master_count": len(prs.slide_masters),
        "layout_count": len(prs.slide_layouts),
        "ordinary_slide_count": len(prs.slides),
        "layout_names": names,
    }


def build_atelier_specs(data: dict[str, Any]) -> list[SlideSpec]:
    variables = data["VARIABLES_GLOBALES"]
    specs: list[SlideSpec] = []
    for slide in data["SLIDES"]:
        placeholders: list[PH] = []
        for key, raw_value in slide["placeholders"].items():
            idx = _parse_ph_key(key)
            value = _resolve_variables(raw_value, variables)
            placeholders.append(PH(idx, value, style="plain", required=True))
        specs.append(
            SlideSpec(
                slide["layout_name_exact"],
                f"{slide['slide_id']} — {slide['partie']}",
                tuple(placeholders),
                clone_slide_number=True,
            )
        )
    return specs
