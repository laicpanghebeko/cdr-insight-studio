"""Streamlit entry point for CDR Insight Studio."""

from __future__ import annotations

import os

import streamlit as st

from cdr_insight.demo import build_demo_pdf
from cdr_insight.service import (
    ProcessingError,
    ProcessingOptions,
    Upload,
    process_uploads,
)

st.set_page_config(
    page_title="CDR Insight Studio",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
      .block-container {max-width: 1180px; padding-top: 2.2rem;}
      [data-testid="stMetricValue"] {font-size: 1.65rem;}
      .hero-kicker {color: #0f766e; font-size: .78rem; font-weight: 750;
                    letter-spacing: .14em; text-transform: uppercase;}
      .hero-title {font-size: clamp(2rem, 5vw, 3.5rem); font-weight: 760;
                   letter-spacing: -.04em; line-height: 1.02; margin: .3rem 0 .8rem;}
      .hero-copy {color: #475569; font-size: 1.05rem; max-width: 760px;}
    </style>
    <div class="hero-kicker">Telecom record intelligence</div>
    <div class="hero-title">CDR Insight Studio</div>
    <div class="hero-copy">
      Analyse locale de listings téléphoniques et IMEI, détection automatique
      des formats et génération de rapports Word structurés.
    </div>
    """,
    unsafe_allow_html=True,
)

st.info(
    "Les PDF sont traités dans un répertoire temporaire de l'instance puis supprimés. "
    "Aucune base de données n'est requise. Utilisez uniquement des données que vous êtes "
    "autorisé à analyser."
)


def _serpapi_key() -> str:
    key = os.getenv("SERPAPI_API_KEY", "").strip()
    if key:
        return key
    try:
        return str(st.secrets.get("SERPAPI_API_KEY", "")).strip()
    except (FileNotFoundError, KeyError):
        return ""


with st.sidebar:
    st.header("Paramètres")
    listing_label = st.selectbox(
        "Type de listing",
        ["Détection automatique", "Numéro de téléphone", "IMEI"],
        help="Le moteur vérifie toujours la structure réelle des PDF.",
    )
    listing_type = {
        "Détection automatique": "auto",
        "Numéro de téléphone": "phone",
        "IMEI": "imei",
    }[listing_label]
    output_label = st.radio(
        "Sortie",
        ["Un rapport combiné", "Un fichier par PDF"],
    )
    output_mode = "combined" if output_label == "Un rapport combiné" else "separate"
    top_n = st.slider("Top correspondants", 1, 20, 5)

    with st.expander("Analyse horaire", expanded=True):
        start_hour = st.number_input("Début de plage", 0, 23, 22)
        end_hour = st.number_input("Fin de plage", 0, 23, 7)
        add_extra_present = st.checkbox(
            "Ajouter un correspondant présent dans la zone retenue",
            value=True,
        )

    theft_date = ""
    if listing_type in {"auto", "imei"}:
        theft_date = st.text_input(
            "Date de référence IMEI (optionnelle)",
            placeholder="JJ/MM/AAAA",
        )

    api_key = _serpapi_key()
    with st.expander("Enrichissement web optionnel"):
        st.caption(
            "Si activé, des numéros et noms peuvent être envoyés à SerpAPI. "
            "Vérifiez votre base légale et votre politique de confidentialité."
        )
        enable_web = st.checkbox(
            "Activer la recherche web",
            value=False,
            disabled=not bool(api_key),
        )
        if not api_key:
            st.caption("Ajoutez SERPAPI_API_KEY aux secrets Streamlit pour activer cette option.")
        max_web_results = st.slider("Liens par correspondant", 1, 5, 3)

st.subheader("Documents")
uploaded_files = st.file_uploader(
    "Déposez jusqu'à 10 listings PDF",
    type=["pdf"],
    accept_multiple_files=True,
    help="25 Mo maximum par fichier, 75 Mo par lot.",
)

demo_col, process_col = st.columns([1, 2])
with demo_col:
    st.download_button(
        "Télécharger le PDF de démonstration",
        data=build_demo_pdf(),
        file_name="listing_demo_synthetique.pdf",
        mime="application/pdf",
        use_container_width=True,
    )
with process_col:
    process_clicked = st.button(
        "Analyser et générer le rapport",
        type="primary",
        disabled=not uploaded_files,
        use_container_width=True,
    )

if process_clicked:
    uploads = [
        Upload(filename=item.name, content=item.getvalue()) for item in uploaded_files
    ]
    options = ProcessingOptions(
        listing_type=listing_type,
        output_mode=output_mode,
        start_hour=int(start_hour),
        end_hour=int(end_hour),
        top_n=top_n,
        add_extra_present=add_extra_present,
        theft_date=theft_date,
        enable_internet_search=enable_web,
        serpapi_api_key=api_key,
        max_web_results=max_web_results,
    )
    try:
        with st.spinner("Lecture, classification et génération du rapport…"):
            result = process_uploads(uploads, options)
    except ProcessingError as exc:
        st.error(str(exc))
    except Exception:
        st.exception(RuntimeError("Une erreur inattendue est survenue pendant l'analyse."))
    else:
        st.success("Rapport généré. Les fichiers temporaires ont été supprimés.")
        metric_a, metric_b, metric_c = st.columns(3)
        metric_a.metric("Type détecté", "IMEI" if result.listing_type == "imei" else "Téléphone")
        metric_b.metric("PDF traités", len(result.detected_files))
        metric_c.metric("Rapports", len(result.summaries))

        st.subheader("Détection")
        st.dataframe(
            [
                {
                    "Fichier": item.filename,
                    "Type": item.listing_type,
                    "Pages": item.pages,
                    "Identifiant": item.identifier or "—",
                }
                for item in result.detected_files
            ],
            use_container_width=True,
            hide_index=True,
        )
        st.subheader("Synthèse")
        st.dataframe(result.summaries, use_container_width=True, hide_index=True)
        st.download_button(
            "Télécharger le résultat",
            data=result.content,
            file_name=result.filename,
            mime=result.mime_type,
            type="primary",
            use_container_width=True,
        )

st.divider()
st.caption(
    "CDR Insight Studio — moteur autonome, sans dépendance à une plateforme CRM ou téléphonique."
)
