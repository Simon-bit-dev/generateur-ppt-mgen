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
    Nettoie uniquement l’enveloppe technique d’un YAML collé.

    La fonction :
    - retire un éventuel BOM Unicode ;
    - normalise les fins de ligne ;
    - retire les éventuelles balises Markdown ```yaml ... ``` ;
    - conserve strictement le contenu et l’indentation du YAML.

    Elle ne remplace pas les puces, ne réindente pas le document et ne tente
    pas de corriger sa structure. La validation réelle est confiée au moteur.
    """
    if not text:
        return ""

    cleaned = (
        text.replace("\ufeff", "")
        .replace("\r\n", "\n")
        .replace("\r", "\n")
        .strip()
    )

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


def _safe_filename(filename: str, fallback: str) -> str:
    """Évite qu’un nom de fichier contienne un chemin ou soit vide."""
    safe_name = Path(filename).name.strip()
    return safe_name or fallback


def _ensure_pptx_filename(filename: str, fallback: str) -> str:
    """Retourne un nom de fichier sûr se terminant exactement par .pptx."""
    safe_name = _safe_filename(filename, fallback)

    if safe_name.lower().endswith(".pptx"):
        return safe_name

    stem = Path(safe_name).stem.strip()
    if not stem:
        stem = Path(fallback).stem or "Support_MGEN"

    return f"{stem}.pptx"


def _reset_generated_files() -> None:
    """Supprime de la session les résultats liés à un ancien YAML."""
    for key in (
        "generated_pptx",
        "generated_report",
        "generated_name",
        "report_name",
        "generated_slides",
    ):
        st.session_state.pop(key, None)


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
            "Le contenu est transmis au véritable parseur YAML et au schéma "
            "métier du générateur. Les puces présentes dans les blocs de texte "
            "sont conservées."
        ),
    )

    cleaned_yaml = _clean_pasted_yaml(pasted_yaml)

    if cleaned_yaml:
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
# Aucun YAML fourni
# ---------------------------------------------------------------------------

if raw is None:
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
        _reset_generated_files()

    # Contrôle préalable par le parseur YAML et le schéma métier du moteur.
    try:
        with tempfile.TemporaryDirectory(prefix="mgen_inspect_") as tmp:
            yaml_path = Path(tmp) / source_name
            yaml_path.write_bytes(raw)
            summary = inspect_yaml(yaml_path)

    except MgenGeneratorError as exc:
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

    except Exception as exc:
        st.markdown(
            """
            <div class="status-error">
                <strong>Erreur inattendue pendant l’analyse du YAML.</strong>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.code(f"{type(exc).__name__}: {exc}", language=None)
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

    output_name = _ensure_pptx_filename(
        requested_output_name,
        fallback=suggested_name,
    )

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
            _reset_generated_files()
            st.error(f"Génération impossible : {exc}")

        except Exception as exc:
            _reset_generated_files()
            st.error(
                "Erreur inattendue pendant la génération : "
                f"{type(exc).__name__}: {exc}"
            )

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
