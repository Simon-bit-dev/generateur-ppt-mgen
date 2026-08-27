from __future__ import annotations

from pathlib import Path

import pytest
import yaml
from pptx import Presentation

from mgen_pptx_generator.atelier import (
    audit_atelier_template,
    validate_atelier_yaml,
)
from mgen_pptx_generator.errors import ValidationError
from mgen_pptx_generator.portable_service import generate_powerpoint, inspect_yaml
from mgen_pptx_generator.validation import (
    audit_template,
    load_yaml,
    validate_semantic_yaml,
)

ROOT = Path(__file__).resolve().parents[1]
CONFERENCE_YAML = ROOT / "examples" / "exemple_conference.yaml"
ATELIER_YAML = ROOT / "examples" / "exemple_atelier.yaml"
CONFERENCE_TEMPLATE = ROOT / "resources" / "templates" / "conference_mgen.pptx"
ATELIER_TEMPLATE = ROOT / "resources" / "templates" / "atelier_mgen.pptx"
ATELIER_SCHEMA = ROOT / "resources" / "schemas" / "schema_atelier_mgen.yaml"


def test_conference_example_validation() -> None:
    counts = validate_semantic_yaml(load_yaml(CONFERENCE_YAML))

    assert counts["total_slides"] == 32


def test_atelier_example_validation() -> None:
    counts = validate_atelier_yaml(
        load_yaml(ATELIER_YAML),
        load_yaml(ATELIER_SCHEMA),
    )

    assert counts["total_slides"] == 33


def test_conference_template_audit() -> None:
    audit = audit_template(CONFERENCE_TEMPLATE, load_yaml(CONFERENCE_YAML))

    assert audit["layout_count"] == 22
    assert audit["ordinary_slide_count"] == 0


def test_atelier_template_audit() -> None:
    audit = audit_atelier_template(
        ATELIER_TEMPLATE,
        load_yaml(ATELIER_YAML),
        load_yaml(ATELIER_SCHEMA),
    )

    assert audit["layout_count"] == 18
    assert audit["ordinary_slide_count"] == 0


@pytest.mark.parametrize(
    ("yaml_path", "filename", "expected_slides"),
    [
        (CONFERENCE_YAML, "conference.pptx", 32),
        (ATELIER_YAML, "atelier.pptx", 33),
    ],
)
def test_complete_generation(
    tmp_path: Path,
    yaml_path: Path,
    filename: str,
    expected_slides: int,
) -> None:
    output_path = tmp_path / filename

    result = generate_powerpoint(yaml_path, output_path)

    assert result.generated_slides == expected_slides
    assert result.output_path == output_path.resolve()
    assert result.output_path.exists()
    assert result.report_path.exists()
    assert len(Presentation(str(result.output_path)).slides) == expected_slides


def test_missing_nested_required_field_is_a_clear_validation_error(
    tmp_path: Path,
) -> None:
    data = load_yaml(CONFERENCE_YAML)
    del data["sections"]["conclusion"]["synthesis"]
    yaml_path = tmp_path / "conference_sans_synthese.yaml"
    yaml_path.write_text(
        yaml.safe_dump(data, allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )

    with pytest.raises(
        ValidationError,
        match=r"sections\.conclusion\.synthesis",
    ):
        inspect_yaml(yaml_path)


def test_missing_required_atelier_field_is_a_clear_validation_error() -> None:
    data = load_yaml(ATELIER_YAML)
    del data["SLIDES"][0]["layout_name_exact"]

    with pytest.raises(
        ValidationError,
        match=r"SLIDES\[0\]\.layout_name_exact",
    ):
        validate_atelier_yaml(data, load_yaml(ATELIER_SCHEMA))


def test_missing_required_root_is_reported_by_the_service(tmp_path: Path) -> None:
    data = load_yaml(CONFERENCE_YAML)
    del data["sections"]
    yaml_path = tmp_path / "conference_sans_sections.yaml"
    yaml_path.write_text(
        yaml.safe_dump(data, allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )

    with pytest.raises(ValidationError, match=r"sections"):
        inspect_yaml(yaml_path)
