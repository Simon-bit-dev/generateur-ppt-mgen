from __future__ import annotations

import hashlib
import tempfile
from pathlib import Path

import streamlit as st

from mgen_pptx_generator.errors import MgenGeneratorError
from mgen_pptx_generator.portable_service import (
    default_output_name,
    generate_powerpoint,
    inspect_yaml,
)


st.set_page_config(
    page_title="Générateur PowerPoint MGEN",
    page_icon="📊",
    layout="centered",
)


# ---------------------------------------------------------------------------
# Fonctions utilitaires
# ---------------------------------------------------------------------------

def _digest(data: bytes) -> str:
    """Calcule l’empreinte du YAML afin de détecter tout changement."""
    return hashlib.sha256(data).hexdigest()


def _clean_pasted_yaml(text: str) -> str:
    """
    Nettoie uniquement l’enveloppe Markdown d’un YAML collé.

    La fonction retire les éventuelles balises :
    ```yaml
    ...
    ```

    Elle ne tente pas de reconstruire ou de réindenter automatiquement
    un contenu transformé en liste à puces.
    """
    cleaned = text.strip()

    if not cleaned:
        return ""

    lines = cleaned.splitlines()

    if lines and lines[0].strip().lower() in {
        "```yaml",
        "```yml",
        "```",
    }:
        lines = lines[1:]

    if lines and lines[-1].strip() == "```":
        lines = lines[:-1]

    return "\n".join(lines).strip()


def _detect_pasted_yaml_problem(text: str) -> str | None:
    """
    Détecte les altérations fréquentes provoquées par un copier-coller
    depuis un traitement de texte ou une interface qui transforme
    le YAML en liste à puces.
    """
    if not text:
        return None

    lines = text.splitlines()

    bullet_lines: list[int] = []
    tab_lines: list[int] = []

    for line_number, line in enumerate(lines, start=1):
        stripped = line.lstrip()

        if stripped.startswith(("•", "‣", "◦", "▪", "●")):
            bullet_lines.append(line_number)

        if "\t" in line:
            tab_lines.append(line_number)

    problems: list[str] = []

    if bullet_lines:
        displayed = ", ".join(str(number) for number in bullet_lines[:8])

        if len(bullet_lines) > 8:
            displayed += ", …"

        problems.append(
            "des puces typographiques ont été détectées "
            f"aux lignes {displayed}"
        )

    if tab_lines:
        displayed = ", ".join(str(number) for number in tab_lines[:8])

        if len(tab_lines) > 8:
            displayed += ", …"

        problems.append(
            "des tabulations ont été détectées "
            f"aux lignes {displayed}"
        )

    if not problems:
        return None

    return (
        "Le contenu collé n’est pas du YAML brut : "
        + " et ".join(problems)
        + ". Copiez directement le contenu d’un bloc de code YAML, "
        "en conservant ses espaces d’indentation."
    )


def _safe_filename(filename: str, fallback: str) -> str:
    """Évite qu’un nom de fichier contienne un chemin ou soit vide."""
    safe_name = Path(filename).name.strip()
    return safe_name or fallback


# ---------------------------------------------------------------------------
# Style
# ---------------------------------------------------------------------------

st.markdown(
    """
    <style>
      .block-container {
          max-width: 860px;
          padding-top: 2.4rem;
          padding-bottom: 3rem;
      }

      .main-title {
          font-size: 2.15rem;
          font-weight: 750;
          margin-bottom: .2rem;
      }

      .main-subtitle {
          color: #5d6470;
          margin-bottom: 1.5rem;
      }

      .status-ok {
          padding: .85rem 1rem;
          border-radius: .55rem;
          background: #eef8f1;
          border: 1px solid #b8dfc3;
      }

      .status-error {
          padding: .85rem 1rem;
          border-radius: .55rem;
          background: #fff1f1;
          border: 1px solid #e5b3b3;
      }

      .small-note {
          font-size: .88rem;
          color: #68707d;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="main-title">Générateur PowerPoint MGEN</div>',
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="main-subtitle">
        Collez ou téléversez le YAML définitif.
        Le type de support, la template et le schéma sont sélectionnés automatiquement.
    </div>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Entrée du YAML
# ---------------------------------------------------------------------------

input_mode = st.radio(
    "Mode d’entrée",
    options=[
        "Coller le YAML",
        "Téléverser un fichier",
    ],
    horizontal=True,
)

raw: bytes | None = None
source_name = "yaml_colle.yaml"
input_error: str | None = None

if input_mode == "Coller le YAML":
    pasted_yaml = st.text_area(
        "Code YAML définitif",
        height=500,
        placeholder=(
            "Collez ici l’intégralité du YAML définitif.\n\n"
            "Exemple :\n"
            "VARIABLES_GLOBALES:\n"
            "  GLOBAL_FOOTER_LEFT: \"...\"\n\n"
            "SLIDES:\n"
            "  - slide_id: \"S01\"\n"
            "    numero: 1"
        ),
        help=(
            "Copiez directement le contenu du bloc de code YAML. "
            "Les espaces d’indentation doivent être conservés."
        ),
    )

    cleaned_yaml = _clean_pasted_yaml(pasted_yaml)

    if cleaned_yaml:
        input_error = _detect_pasted_yaml_problem(cleaned_yaml)

        if input_error is None:
            raw = cleaned_yaml.encode("utf-8")

else:
    uploaded = st.file_uploader(
        "Fichier YAML définitif",
        type=["yaml", "yml"],
        accept_multiple_files=False,
        help="Formats acceptés : .yaml et .yml",
    )

    if uploaded is not None:
        raw = uploaded.getvalue()
        source_name = _safe_filename(
            uploaded.name,
            fallback="support_mgen.yaml",
        )


# ---------------------------------------------------------------------------
# Erreur de copier-coller
# ---------------------------------------------------------------------------

if input_error is not None:
    st.markdown(
        """
        <div class="status-error">
            <strong>Le contenu collé a perdu sa structure YAML.</strong>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.error(input_error)

    st.markdown(
        """
        Le YAML doit notamment respecter cette structure :

        ```yaml
        VARIABLES_GLOBALES:
          GLOBAL_FOOTER_LEFT: "Atelier de prévention"
          GLOBAL_FOOTER_RIGHT: "[Date à compléter]"

        SLIDES:
          - slide_id: "S01"
            numero: 1
            partie: "Ouverture institutionnelle"
            layout_code: "MGEN-01"
            placeholders:
              PH23: "{{GLOBAL_FOOTER_LEFT}}"
        ```
        """
    )


# ---------------------------------------------------------------------------
# Aucun YAML fourni
# ---------------------------------------------------------------------------

elif raw is None:
    st.info(
        "Collez un YAML définitif ou téléversez un fichier "
        "Atelier, Conférence ou Webinaire pour commencer."
    )

    st.markdown("### Exemples")

    col1, col2 = st.columns(2)

    conference_example = Path("examples/exemple_conference.yaml")
    atelier_example = Path("examples/exemple_atelier.yaml")

    if conference_example.exists():
        col1.download_button(
            "Télécharger l’exemple Conférence",
            data=conference_example.read_bytes(),
            file_name=conference_example.name,
            mime="application/yaml",
            use_container_width=True,
        )

    if atelier_example.exists():
        col2.download_button(
            "Télécharger l’exemple Atelier",
            data=atelier_example.read_bytes(),
            file_name=atelier_example.name,
            mime="application/yaml",
            use_container_width=True,
        )


# ---------------------------------------------------------------------------
# Inspection et génération
# ---------------------------------------------------------------------------

else:
    current_digest = _digest(raw)

    # Réinitialise les fichiers générés lorsque le YAML change.
    if st.session_state.get("yaml_digest") != current_digest:
        st.session_state["yaml_digest"] = current_digest
        st.session_state.pop("generated_pptx", None)
        st.session_state.pop("generated_report", None)
        st.session_state.pop("generated_name", None)
        st.session_state.pop("report_name", None)
        st.session_state.pop("generated_slides", None)

    # Contrôle préalable du YAML.
    try:
        with tempfile.TemporaryDirectory(prefix="mgen_inspect_") as tmp:
            yaml_path = Path(tmp) / source_name
            yaml_path.write_bytes(raw)

            summary = inspect_yaml(yaml_path)

    except Exception as exc:
        st.markdown(
            """
            <div class="status-error">
                <strong>YAML non conforme.</strong>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.code(str(exc), language=None)
        st.stop()

    st.markdown(
        """
        <div class="status-ok">
            <strong>YAML conforme et prêt à être généré.</strong>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.write("")

    c1, c2 = st.columns(2)

    c1.metric(
        "Type détecté",
        summary.support_label,
    )

    c2.metric(
        "Nombre de slides",
        summary.expected_slides,
    )

    st.write(f"**Titre :** {summary.title}")

    if summary.subtitle:
        st.write(f"**Sous-titre :** {summary.subtitle}")

    st.write(f"**Architecture :** {summary.architecture}")
    st.write(f"**Date :** {summary.date}")

    suggested_name = default_output_name(summary)

    requested_output_name = st.text_input(
        "Nom du PowerPoint",
        value=suggested_name,
        key=f"output_name_{current_digest}",
    )

    output_name = _safe_filename(
        requested_output_name,
        fallback=suggested_name,
    )

    if not output_name.lower().endswith(".pptx"):
        output_name += ".pptx"

    if st.button(
        "Générer le PowerPoint",
        type="primary",
        use_container_width=True,
    ):
        try:
            with st.spinner("Génération et contrôles techniques en cours…"):
                with tempfile.TemporaryDirectory(
                    prefix="mgen_generate_"
                ) as tmp:
                    tmpdir = Path(tmp)

                    yaml_path = tmpdir / source_name
                    output_path = tmpdir / output_name

                    yaml_path.write_bytes(raw)

                    result = generate_powerpoint(
                        yaml_path,
                        output_path,
                    )

                    st.session_state["generated_pptx"] = (
                        result.output_path.read_bytes()
                    )

                    st.session_state["generated_report"] = (
                        result.report_path.read_bytes()
                    )

                    st.session_state["generated_name"] = (
                        result.output_path.name
                    )

                    st.session_state["report_name"] = (
                        result.report_path.name
                    )

                    st.session_state["generated_slides"] = (
                        result.generated_slides
                    )

            st.success(
                "PowerPoint généré avec succès : "
                f"{st.session_state['generated_slides']} slides."
            )

        except MgenGeneratorError as exc:
            st.error(f"Génération impossible : {exc}")

        except Exception as exc:
            st.error(f"Erreur inattendue : {exc}")

    if st.session_state.get("generated_pptx"):
        st.markdown("### Fichiers générés")

        d1, d2 = st.columns(2)

        d1.download_button(
            "Télécharger le PowerPoint",
            data=st.session_state["generated_pptx"],
            file_name=st.session_state["generated_name"],
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "presentationml.presentation"
            ),
            type="primary",
            use_container_width=True,
        )

        d2.download_button(
            "Télécharger le rapport",
            data=st.session_state["generated_report"],
            file_name=st.session_state["report_name"],
            mime="text/markdown",
            use_container_width=True,
        )

        st.warning(
            "Une vérification visuelle rapide dans PowerPoint "
            "reste nécessaire avant diffusion."
        )


# ---------------------------------------------------------------------------
# Pied de page
# ---------------------------------------------------------------------------

st.divider()

st.markdown(
    """
    <div class="small-note">
        Les fichiers sont traités temporairement pendant la session.
        Aucun compte IA n’est utilisé pour la génération du PowerPoint.
    </div>
    """,
    unsafe_allow_html=True,
)
