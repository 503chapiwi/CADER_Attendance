"""Generador de reportes MAGA Totonicapán.

Convierte las asistencias del formulario Kobo "Planillas MAGA Toto" en:
  1. Listado de beneficiarios (registro mensual por intervención)
  2. Informe mensual de talleres o eventos de capacitación
"""

from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

import kobo_client as kc
from reports import MESES, SECCION_INFORME, filtrar_envios, generar_informe, generar_listado, indexar_participantes, indexar_promotores, registro_nuevos_miembros

st.set_page_config(page_title="Reportes MAGA Totonicapán", page_icon="🌽", layout="centered")


def get_token():
    try:
        token = st.secrets.get("KOBO_TOKEN", None)
    except Exception:  # no existe secrets.toml
        token = None
    if not token:
        token_file = Path(__file__).parent / "kobo_token.txt"
        if token_file.exists():
            token = token_file.read_text().strip()
    return token


@st.cache_data(ttl=600, show_spinner=False)
def cargar_datos_kobo(token):
    asset = kc.fetch_asset(token)
    choice_maps, field_lists = kc.build_label_maps(asset)
    submissions = kc.fetch_submissions(token)
    participantes = kc.fetch_participantes(token)
    promotores = kc.fetch_promotores(token)
    return choice_maps, field_lists, submissions, participantes, promotores


st.title("🌽 Reportes MAGA Totonicapán")
st.markdown(
    "Esta herramienta descarga las asistencias del formulario de Kobo y genera "
    "el **Listado de Beneficiarios** y el **Informe Mensual** listos para entregar."
)

token = get_token()
if not token:
    token = st.text_input("Token de Kobo", type="password", help="Pida el token al administrador del sistema.")
    if not token:
        st.stop()

try:
    with st.spinner("Descargando datos de Kobo..."):
        choice_maps, field_lists, submissions, participantes_df, promotores_df = cargar_datos_kobo(token)
except Exception as e:
    st.error(f"No se pudo conectar con Kobo. Revise el internet y el token.\n\nDetalle: {e}")
    st.stop()

municipios = choice_maps.get("municipio", {})
municipio_labels = {code: label for code, label in sorted(municipios.items(), key=lambda x: x[1]) if kc._slug(code) in SECCION_INFORME}

st.header("1. Elija el municipio y el mes")
col1, col2, col3 = st.columns([2, 1, 1])
with col1:
    municipio_code = st.selectbox("Municipio", options=list(municipio_labels), format_func=lambda c: municipio_labels[c])
with col2:
    hoy = date.today()
    mes = st.selectbox("Mes", options=list(range(1, 13)), index=hoy.month - 1, format_func=lambda m: MESES[m - 1])
with col3:
    anio = st.number_input("Año", min_value=2024, max_value=2050, value=hoy.year)

st.header("2. Elija los reportes")
hacer_listado = st.checkbox("Listado de Beneficiarios", value=True)
hacer_informe = st.checkbox("Informe Mensual de Capacitaciones", value=True)

with st.expander("Opciones avanzadas"):
    st.markdown(
        "La base de participantes CADER se descarga automáticamente de Kobo "
        f"(**{kc.PARTICIPANTES_FILENAME}**, {len(participantes_df)} personas). "
        "Si tiene una versión más reciente, súbala aquí:"
    )
    archivo_participantes = st.file_uploader("Base de participantes (CSV o Excel)", type=["csv", "xlsx"])
    if st.button("🔄 Actualizar datos de Kobo"):
        st.cache_data.clear()
        st.rerun()

if st.button("Generar reportes", type="primary", disabled=not (hacer_listado or hacer_informe)):
    if archivo_participantes is not None:
        if archivo_participantes.name.lower().endswith(".csv"):
            participantes_df = pd.read_csv(archivo_participantes, dtype=str)
        else:
            participantes_df = pd.read_excel(archivo_participantes, dtype=str)

    envios = filtrar_envios(submissions, municipio_code, mes, anio)
    nombre_mes = MESES[mes - 1]
    etiqueta_muni = municipio_labels[municipio_code]

    if not envios:
        st.warning(f"No se encontraron envíos de **{etiqueta_muni}** en **{nombre_mes} {anio}**. Revise el municipio, el mes y que las boletas ya estén enviadas en Kobo.")
        st.stop()

    st.success(f"Se encontraron **{len(envios)}** boletas de **{etiqueta_muni}** en **{nombre_mes} {anio}**.")
    sufijo = f"{etiqueta_muni.replace(' ', '_')}_{nombre_mes}_{anio}"

    if hacer_listado:
        with st.spinner("Generando listado de beneficiarios..."):
            participantes_idx = indexar_participantes(participantes_df)
            nuevos_idx = registro_nuevos_miembros(submissions)
            listado_bytes, resumen = generar_listado(envios, participantes_idx, nuevos_idx, choice_maps, field_lists, mes, anio)
        st.subheader("📋 Listado de Beneficiarios")
        detalle = " · ".join(f"{hoja}: {n}" for hoja, n in resumen["por_hoja"].items())
        st.markdown(f"**{resumen['beneficiarios']}** beneficiarios en **{resumen['eventos']}** eventos ({detalle})")
        if resumen["sin_datos"]:
            st.warning(
                f"⚠️ {len(resumen['sin_datos'])} CUI(s) no aparecen en la base de participantes. "
                "Salen en el listado solo con su CUI; complete sus datos a mano o actualice la base."
            )
            st.dataframe(pd.DataFrame(resumen["sin_datos"]), hide_index=True)
        st.download_button(
            "⬇️ Descargar Listado de Beneficiarios",
            data=listado_bytes,
            file_name=f"Listado_Beneficiarios_{sufijo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if hacer_informe:
        with st.spinner("Generando informe mensual..."):
            promotores_idx = indexar_promotores(promotores_df)
            informe_bytes, resumen = generar_informe(envios, promotores_idx, choice_maps, field_lists, municipio_code, mes, anio)
        st.subheader("📊 Informe Mensual")
        if resumen["eventos"] == 0:
            st.warning("Ninguna boleta de este mes tiene la sección de informe llenada (pregunta 'informe' = Sí). El archivo saldrá con la sección vacía.")
        else:
            st.markdown(f"**{resumen['eventos']}** eventos en la sección de **{etiqueta_muni}**")
        for adv in resumen["advertencias"]:
            st.warning("⚠️ " + adv)
        st.download_button(
            "⬇️ Descargar Informe Mensual",
            data=informe_bytes,
            file_name=f"Informe_Mensual_{sufijo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
