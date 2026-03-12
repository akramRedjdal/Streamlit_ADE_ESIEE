#!/usr/bin/env python3
"""
ade_app.py - Interface Streamlit pour analyser les heures d'enseignement ADE.

Usage: streamlit run ade_app.py
Dépendances: streamlit, pandas, openpyxl (+ ade_heures.py dans le même dossier)
"""

import io
import os
import tempfile
from collections import defaultdict

import pandas as pd
import streamlit as st

from ade_heures import (
    HETD_COEFFICIENTS,
    MODALITY_COLORS,
    MODALITY_ORDER,
    generate_excel,
    hetd,
    parse_ics,
    process_events,
)

# ---------------------------------------------------------------------------
# Config page
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="ADE Heures",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.title("Analyse des heures d'enseignement — ADE")
st.caption("Importez un fichier `.ics` exporté depuis ADE pour obtenir le détail et le récapitulatif de vos heures.")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def records_to_df(records):
    """Convert list of record dicts to a clean DataFrame."""
    rows = []
    for r in records:
        rows.append({
            "Nom":        r["nom"],
            "Date":       r["dtstart"].strftime("%d/%m/%Y"),
            "Début":      r["dtstart"].strftime("%H:%M"),
            "Fin":        r["dtend"].strftime("%H:%M"),
            "Durée (h)":  round(r["duration_h"], 2),
            "HETD (h)":   round(hetd(r["duration_h"], r["modality"]), 2),
            "Lieu":       r["location"],
            "Modalité":   r["modality"],
            # keep raw dtstart for sorting
            "_dtstart":   r["dtstart"],
        })
    return pd.DataFrame(rows)


def build_modality_summary(df):
    """Build summary table grouped by modality."""
    grp = (
        df.groupby("Modalité", sort=False)
        .agg(
            Séances=("Nom", "count"),
            Heures=("Durée (h)", "sum"),
            HETD=("HETD (h)", "sum"),
        )
        .reset_index()
    )
    grp["Coeff HETD"] = grp["Modalité"].map(HETD_COEFFICIENTS).fillna(0)
    grp["Heures"] = grp["Heures"].round(2)
    grp["HETD"]   = grp["HETD"].round(2)
    # Apply fixed order
    order = {m: i for i, m in enumerate(MODALITY_ORDER)}
    grp["_order"] = grp["Modalité"].map(order).fillna(999)
    grp = grp.sort_values("_order").drop(columns="_order").reset_index(drop=True)
    # Add TOTAL row
    total = pd.DataFrame([{
        "Modalité":   "TOTAL",
        "Coeff HETD": "",
        "Séances":    int(grp["Séances"].sum()),
        "Heures":     round(grp["Heures"].sum(), 2),
        "HETD":       round(grp["HETD"].sum(), 2),
    }])
    return pd.concat([grp, total], ignore_index=True)[
        ["Modalité", "Coeff HETD", "Séances", "Heures", "HETD"]
    ]


def build_course_summary(df):
    """Build summary table grouped by course name."""
    grp = (
        df.groupby("Nom", sort=False)
        .agg(
            Séances=("Nom", "count"),
            Heures=("Durée (h)", "sum"),
            HETD=("HETD (h)", "sum"),
            Modalités=("Modalité", lambda s: ", ".join(sorted(s.unique()))),
        )
        .reset_index()
        .sort_values("Heures", ascending=False)
        .reset_index(drop=True)
    )
    grp["Heures"] = grp["Heures"].round(2)
    grp["HETD"]   = grp["HETD"].round(2)
    return grp


def make_excel_bytes(records):
    """Generate Excel file in memory and return bytes."""
    buf = io.BytesIO()
    generate_excel(records, buf)
    buf.seek(0)
    return buf.read()


# Modality → hex color for pandas Styler (strip FF alpha prefix)
def _hex(rrggbbaa):
    return "#" + rrggbbaa[2:]

MOD_CSS = {m: _hex(c) for m, c in MODALITY_COLORS.items()}


def style_modality(df_display):
    """Apply row background based on Modalité column."""
    def row_style(row):
        color = MOD_CSS.get(row.get("Modalité", ""), "#ffffff")
        return [f"background-color: {color}" for _ in row]
    return df_display.style.apply(row_style, axis=1)


# ---------------------------------------------------------------------------
# File upload
# ---------------------------------------------------------------------------

uploaded = st.file_uploader(
    "Choisissez un fichier ADE (.ics)",
    type=["ics"],
    help="Exportez votre emploi du temps depuis ADE au format iCalendar (.ics)",
)

if uploaded is None:
    st.info("Importez un fichier `.ics` pour commencer.")
    st.stop()

# Parse — save to temp file so parse_ics can open it normally
with tempfile.NamedTemporaryFile(suffix=".ics", delete=False) as tmp:
    tmp.write(uploaded.read())
    tmp_path = tmp.name

try:
    raw_events = parse_ics(tmp_path)
    records    = process_events(raw_events)
finally:
    os.unlink(tmp_path)

if not records:
    st.error("Aucun événement valide trouvé dans ce fichier.")
    st.stop()

df = records_to_df(records)

# ---------------------------------------------------------------------------
# Top metrics
# ---------------------------------------------------------------------------

st.markdown("---")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Séances totales",  len(records))
c2.metric("Heures totales",   f"{df['Durée (h)'].sum():.1f} h")
c3.metric("HETD totales",     f"{df['HETD (h)'].sum():.1f}")
c4.metric("Cours distincts",  df["Nom"].nunique())

st.markdown("---")

# ---------------------------------------------------------------------------
# Section 1 — Récapitulatif par modalité
# ---------------------------------------------------------------------------

st.header("📊 Récapitulatif par modalité")

summary_df = build_modality_summary(df)

col_table, col_chart = st.columns([1, 1], gap="large")

with col_table:
    st.dataframe(
        style_modality(summary_df.iloc[:-1]),  # sans la ligne TOTAL pour le style
        use_container_width=True,
        hide_index=True,
    )
    # Ligne TOTAL en gras sous le tableau
    total_row = summary_df.iloc[-1]
    st.markdown(
        f"**TOTAL — {int(total_row['Séances'])} séances | "
        f"{total_row['Heures']:.2f} h | "
        f"{total_row['HETD']:.2f} HETD**"
    )

with col_chart:
    chart_data = summary_df.iloc[:-1].set_index("Modalité")[["Heures"]]
    st.bar_chart(chart_data, use_container_width=True)

# ---------------------------------------------------------------------------
# Section 2 — Téléchargement Excel
# ---------------------------------------------------------------------------

st.markdown("---")
st.header("⬇️ Télécharger le fichier Excel")

excel_bytes = make_excel_bytes(records)
default_name = uploaded.name.replace(".ics", "_heures.xlsx")

st.download_button(
    label="📥 Télécharger le fichier Excel complet",
    data=excel_bytes,
    file_name=default_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

# ---------------------------------------------------------------------------
# Section 3 — Exploration
# ---------------------------------------------------------------------------

st.markdown("---")
st.header("🔍 Explorer les séances")

tab_mod, tab_cours = st.tabs(["Par modalité", "Par nom de cours"])

# ---- Tab : Par modalité ----
with tab_mod:
    all_mods = [m for m in MODALITY_ORDER if m in df["Modalité"].values]
    # Add any unexpected modality at the end
    for m in df["Modalité"].unique():
        if m not in all_mods:
            all_mods.append(m)

    selected_mods = st.multiselect(
        "Filtrer par modalité",
        options=all_mods,
        default=all_mods,
        key="filter_mod",
    )

    df_mod = df[df["Modalité"].isin(selected_mods)].drop(columns="_dtstart").sort_values("Date")

    st.caption(f"{len(df_mod)} séance(s) — {df_mod['Durée (h)'].sum():.2f} h — {df_mod['HETD (h)'].sum():.2f} HETD")
    st.dataframe(
        style_modality(df_mod.reset_index(drop=True)),
        use_container_width=True,
        hide_index=True,
    )

# ---- Tab : Par nom de cours ----
with tab_cours:
    course_summary = build_course_summary(df)

    st.subheader("Résumé par cours")
    st.dataframe(course_summary, use_container_width=True, hide_index=True)

    st.subheader("Détail d'un cours")
    course_names = course_summary["Nom"].tolist()
    selected_course = st.selectbox("Sélectionner un cours", options=course_names, key="select_course")

    if selected_course:
        df_course = (
            df[df["Nom"] == selected_course]
            .drop(columns="_dtstart")
            .sort_values("Date")
            .reset_index(drop=True)
        )
        n = len(df_course)
        h = df_course["Durée (h)"].sum()
        e = df_course["HETD (h)"].sum()
        mods = ", ".join(sorted(df_course["Modalité"].unique()))

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Séances",  n)
        c2.metric("Heures",   f"{h:.2f} h")
        c3.metric("HETD",     f"{e:.2f}")
        c4.metric("Modalités", mods)

        st.dataframe(
            style_modality(df_course),
            use_container_width=True,
            hide_index=True,
        )
