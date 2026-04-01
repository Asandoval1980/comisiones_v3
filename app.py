from __future__ import annotations

import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from liquidador import CommissionProcessor, FIXED_RULES, OUTPUT_GROUP_ORDER

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "assets" / "Liquidacion_Plantilla.xlsx"

st.set_page_config(page_title="Liquidador de comisiones v3", page_icon="💼", layout="wide")

st.title("💼 Liquidador de comisiones v3")
st.caption("Exporta el archivo final con la misma base de hojas, filas de encabezado y columnas del formato ejemplo 'Liquidación de Comisiones'.")

with st.sidebar:
    st.header("1) Carga de archivos")
    cobros = st.file_uploader("Cobros por vendedor", type=["xls", "xlsx", "html"])
    analisis = st.file_uploader("Análisis_Vendedores", type=["xls", "xlsx", "html"])
    bnt_clpg = st.file_uploader("BNT - CLPG", type=["xlsx"])
    bnt_dgtl = st.file_uploader("BNT - DGTL", type=["xlsx"])
    procesar = st.button("Procesar liquidación", type="primary", use_container_width=True)

col1, col2 = st.columns([1.2, 0.8])
with col1:
    st.subheader("Alcance")
    st.markdown(
        """
- Reglas de comisión fijas dentro del código.
- No solicita el archivo **Tabla de Comisiones**.
- Exporta en Excel usando la misma estructura base del archivo **Liquidación de Comisiones**.
- Mantiene únicamente las hojas del formato original.
- Barranquilla sigue excluido del cálculo automático.
- Las validaciones se muestran en pantalla, no como hojas extra del Excel final.
        """
    )
    with st.expander("Reglas fijas cargadas"):
        st.json(FIXED_RULES, expanded=False)

with col2:
    st.subheader("2) Penalizaciones opcionales")
    penalty_default = pd.DataFrame(columns=["group", "user", "vendor", "penalty_count", "penalty_discount", "penalty_note"], data=[["INBOUND AM", "", "", 0, 0, ""]])
    penalties_df = st.data_editor(
        penalty_default,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "group": st.column_config.SelectboxColumn("Grupo", options=OUTPUT_GROUP_ORDER, required=False),
            "user": st.column_config.TextColumn("User"),
            "vendor": st.column_config.TextColumn("Agente comercial"),
            "penalty_count": st.column_config.NumberColumn("Penalidades", min_value=0, step=1),
            "penalty_discount": st.column_config.NumberColumn("Descuento USD", min_value=0.0, step=1.0, format="%.2f"),
            "penalty_note": st.column_config.TextColumn("Observación"),
        },
    )

required_ok = all([cobros, analisis, bnt_clpg, bnt_dgtl])
if not procesar:
    st.info("Carga los archivos y luego pulsa Procesar liquidación.")

if procesar:
    if not required_ok:
        st.error("Debes cargar: Cobros por vendedor, Análisis_Vendedores, BNT - CLPG y BNT - DGTL.")
    else:
        with tempfile.TemporaryDirectory() as td:
            td_path = Path(td)
            files = {
                "cobros": (cobros, td_path / cobros.name),
                "analisis": (analisis, td_path / analisis.name),
                "bnt_clpg": (bnt_clpg, td_path / bnt_clpg.name),
                "bnt_dgtl": (bnt_dgtl, td_path / bnt_dgtl.name),
            }
            for uploaded, path in files.values():
                path.write_bytes(uploaded.getbuffer())

            penalties_clean = penalties_df.copy().fillna({"group": "", "user": "", "vendor": "", "penalty_count": 0, "penalty_discount": 0, "penalty_note": ""})
            penalties_clean = penalties_clean[(penalties_clean["group"].astype(str).str.strip() != "") | (penalties_clean["user"].astype(str).str.strip() != "") | (penalties_clean["vendor"].astype(str).str.strip() != "") | (pd.to_numeric(penalties_clean["penalty_count"], errors="coerce").fillna(0) != 0) | (pd.to_numeric(penalties_clean["penalty_discount"], errors="coerce").fillna(0) != 0)]

            with st.spinner("Procesando liquidación..."):
                processor = CommissionProcessor(
                    cobros_path=files["cobros"][1],
                    analisis_path=files["analisis"][1],
                    bnt_clpg_path=files["bnt_clpg"][1],
                    bnt_dgtl_path=files["bnt_dgtl"][1],
                    penalties_df=penalties_clean,
                )
                result = processor.process()
                output_path = td_path / "Liquidacion_Comisiones_V3.xlsx"
                processor.export_to_template(TEMPLATE_PATH, output_path, result)

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Fecha inicio", result.date_start or "-")
            m2.metric("Fecha fin", result.date_end or "-")
            m3.metric("Total liquidado", f"${result.total_table['TOTAL'].fillna(0).sum():,.0f}")
            m4.metric("Penalizaciones USD", f"${result.penalties_table['penalty_discount'].sum() if not result.penalties_table.empty else 0:,.2f}")

            tabs = st.tabs(["Resumen total", "Validaciones", "Penalizaciones"] + OUTPUT_GROUP_ORDER)
            with tabs[0]:
                st.dataframe(result.total_table, use_container_width=True)
            with tabs[1]:
                st.dataframe(result.validations, use_container_width=True)
            with tabs[2]:
                st.dataframe(result.penalties_table if not result.penalties_table.empty else pd.DataFrame(columns=["Sin penalizaciones registradas"]), use_container_width=True)
            for idx, group in enumerate(OUTPUT_GROUP_ORDER, start=3):
                with tabs[idx]:
                    st.dataframe(result.detail_tables[group], use_container_width=True)

            st.download_button(
                "📥 Descargar liquidación en Excel",
                data=output_path.read_bytes(),
                file_name=output_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
