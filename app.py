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

st.markdown(
    """
    <style>
      .block-container {max-width: 860px; padding-top: 2.4rem; padding-bottom: 3rem;}
      .main-title {font-size: 2.15rem; font-weight: 750; margin-bottom: .2rem;}
      .main-subtitle {color: #5d6470; margin-bottom: 1.5rem;}
      .status-ok {padding: .85rem 1rem; border-radius: .55rem; background: #eef8f1; border: 1px solid #b8dfc3;}
      .status-error {padding: .85rem 1rem; border-radius: .55rem; background: #fff1f1; border: 1px solid #e5b3b3;}
      .small-note {font-size: .88rem; color: #68707d;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="main-title">Générateur PowerPoint MGEN</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="main-subtitle">Déposez le YAML définitif. Le type de support, la template et le schéma sont sélectionnés automatiquement.</div>',
    unsafe_allow_html=True,
)

uploaded = st.file_uploader(
    "Fichier YAML définitif",
    type=["yaml", "yml"],
    accept_multiple_files=False,
    help="Formats acceptés : .yaml et .yml",
)


def _digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


if uploaded is None:
    st.info("Sélectionnez un YAML Atelier, Conférence ou Webinaire pour commencer.")
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
else:
    raw = uploaded.getvalue()
    current_digest = _digest(raw)
    if st.session_state.get("uploaded_digest") != current_digest:
        st.session_state["uploaded_digest"] = current_digest
        st.session_state.pop("generated_pptx", None)
        st.session_state.pop("generated_report", None)
        st.session_state.pop("generated_name", None)
        st.session_state.pop("report_name", None)

    try:
        with tempfile.TemporaryDirectory(prefix="mgen_inspect_") as tmp:
            yaml_path = Path(tmp) / uploaded.name
            yaml_path.write_bytes(raw)
            summary = inspect_yaml(yaml_path)
    except Exception as exc:
        st.markdown(
            f'<div class="status-error"><strong>YAML non conforme.</strong><br>{exc}</div>',
            unsafe_allow_html=True,
        )
        st.stop()

    st.markdown('<div class="status-ok"><strong>YAML conforme et prêt à être généré.</strong></div>', unsafe_allow_html=True)
    st.write("")
    c1, c2 = st.columns(2)
    c1.metric("Type détecté", summary.support_label)
    c2.metric("Nombre de slides", summary.expected_slides)
    st.write(f"**Titre :** {summary.title}")
    if summary.subtitle:
        st.write(f"**Sous-titre :** {summary.subtitle}")
    st.write(f"**Architecture :** {summary.architecture}")
    st.write(f"**Date :** {summary.date}")

    suggested_name = default_output_name(summary)
    output_name = st.text_input("Nom du PowerPoint", value=suggested_name)
    if not output_name.lower().endswith(".pptx"):
        output_name += ".pptx"

    if st.button("Générer le PowerPoint", type="primary", use_container_width=True):
        try:
            with st.spinner("Génération et contrôles techniques en cours…"):
                with tempfile.TemporaryDirectory(prefix="mgen_generate_") as tmp:
                    tmpdir = Path(tmp)
                    yaml_path = tmpdir / uploaded.name
                    output_path = tmpdir / output_name
                    yaml_path.write_bytes(raw)
                    result = generate_powerpoint(yaml_path, output_path)
                    st.session_state["generated_pptx"] = result.output_path.read_bytes()
                    st.session_state["generated_report"] = result.report_path.read_bytes()
                    st.session_state["generated_name"] = result.output_path.name
                    st.session_state["report_name"] = result.report_path.name
                    st.session_state["generated_slides"] = result.generated_slides
            st.success(
                f"PowerPoint généré avec succès : {st.session_state['generated_slides']} slides."
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
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
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
        st.warning("Une vérification visuelle rapide dans PowerPoint reste nécessaire avant diffusion.")

st.divider()
st.markdown(
    '<div class="small-note">Les fichiers sont traités temporairement pendant la session. Aucun compte IA n’est utilisé pour la génération du PowerPoint.</div>',
    unsafe_allow_html=True,
)
