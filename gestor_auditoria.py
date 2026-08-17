import io
import json
import re
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# -----------------------------------------------------------------------------
# CONFIGURACIÓN INICIAL DE STREAMLIT Y ESTILOS CSS
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Gestor y Auditor de Comprobantes",
    page_icon="📊",
    layout="wide",
)

st.markdown(
    """
    <style>
    /* Ajuste de tamaño para los valores de métricas */
    [data-testid="stMetricValue"] {
        font-size: 15pt !important;
        font-weight: 700 !important;
        line-height: 1.3 !important;
        color: #1F2937 !important;
    }
    
    [data-testid="stMetricLabel"] {
        font-size: 11pt !important;
        font-weight: 600 !important;
        color: #4B5563 !important;
        margin-bottom: 2px !important;
    }

    /* Barra de herramientas en tablas Streamlit estándar */
    [data-testid="stElementToolbar"] {
        opacity: 1 !important;
        visibility: visible !important;
        display: flex !important;
        background-color: #F8FAFC !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 6px !important;
        padding: 2px 6px !important;
        box-shadow: 0px 1px 3px rgba(0,0,0,0.08) !important;
    }

    [data-testid="stElementToolbar"] button {
        color: #334155 !important;
        font-size: 12px !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# CONSTANTES Y FUNCIONES AUXILIARES
# -----------------------------------------------------------------------------

TIPOS_COMPROBANTES_ESTANDAR = [
    "Factura A",
    "Factura B",
    "Nota de Crédito A",
    "Nota de Crédito B",
]


def normalizar_pdv(val) -> str:
    """Garantiza que cualquier formato de PDV se estandarice a 4 dígitos."""
    if pd.isna(val) or str(val).strip() == "":
        return "0000"
    val_clean = str(val).split(".")[0].strip()
    try:
        return str(int(val_clean)).zfill(4)
    except ValueError:
        return val_clean.zfill(4)


def extraer_metadatos_hoja(df_raw):
    """Extrae Establecimiento (celda C2) y Período (celda E2) de la hoja Excel."""
    est = "No Especificado"
    periodo = "No Especificado"

    try:
        if df_raw.shape[0] >= 2 and df_raw.shape[1] >= 3:
            val_c2 = df_raw.iloc[1, 2]
            if pd.notna(val_c2) and str(val_c2).strip() != "":
                est = str(val_c2).strip()
    except Exception:
        pass

    try:
        if df_raw.shape[0] >= 2 and df_raw.shape[1] >= 5:
            val_e2 = df_raw.iloc[1, 4]
            if pd.notna(val_e2) and str(val_e2).strip() != "":
                periodo = str(val_e2).strip()
    except Exception:
        pass

    if periodo == "No Especificado":
        for col in df_raw.columns:
            serie_texto = df_raw[col].dropna().astype(str)
            for texto in serie_texto:
                if "Periodo:" in texto or "Período:" in texto:
                    fechas = re.findall(r"\b\d{1,2}/\d{1,2}/\d{4}\b", texto)
                    if len(fechas) >= 2:
                        periodo = f"Desde el {fechas[0]} al {fechas[1]}"
                        break

    return est, periodo


def procesar_multiples_reportes(archivos_subidos):
    """Lee y unifica múltiples reportes Excel procesando TODAS las pestañas."""
    todas_facturas = []
    patron_factura = re.compile(r"^\d+-\d+$")

    for archivo in archivos_subidos:
        try:
            hojas_excel = pd.read_excel(archivo, header=None, sheet_name=None)

            for nombre_hoja, df_raw in hojas_excel.items():
                est_hoja, periodo_hoja = extraer_metadatos_hoja(df_raw)
                tipo_comprobante_actual = "Factura A"

                for index, row in df_raw.iterrows():
                    val_col0 = str(row[0]).strip() if pd.notna(row[0]) else ""

                    if not val_col0:
                        continue

                    if patron_factura.match(val_col0):
                        pdv_raw, num_factura_str = val_col0.split("-")
                        pdv_norm = normalizar_pdv(pdv_raw)

                        todas_facturas.append({
                            "N°Factura": (
                                f"{pdv_norm}-{str(int(num_factura_str)).zfill(8)}"
                            ),
                            "TipoComprobante": tipo_comprobante_actual,
                            "Fecha": (
                                str(row[1]).strip() if pd.notna(row[1]) else ""
                            ),
                            "Cliente": (
                                str(row[2]).strip() if pd.notna(row[2]) else ""
                            ),
                            "Cuit": (
                                str(row[3]).strip() if pd.notna(row[3]) else ""
                            ),
                            "Número de Comanda": (
                                str(row[4]).strip() if pd.notna(row[4]) else ""
                            ),
                            "Total": float(row[5]) if pd.notna(row[5]) else 0.0,
                            "punto_venta": pdv_norm,
                            "numero_factura": int(num_factura_str),
                            "fecha_emision_dt": pd.to_datetime(
                                row[1], errors="coerce", dayfirst=True
                            ),
                            "archivo_origen": f"{archivo.name} ({nombre_hoja})",
                            "fila_excel": index + 1,
                            "establecimiento": est_hoja,
                            "periodo": periodo_hoja,
                        })
                    else:
                        val_upper = val_col0.upper()
                        if any(
                            k in val_upper
                            for k in [
                                "FACTURA",
                                "NOTA",
                                "CREDITO",
                                "CRÉDITO",
                                "NC",
                                "FAC",
                            ]
                        ):
                            if "NOTA" in val_upper or "NC" in val_upper:
                                if " B" in val_upper or val_upper.endswith(" B"):
                                    tipo_comprobante_actual = (
                                        "Nota de Crédito B"
                                    )
                                else:
                                    tipo_comprobante_actual = (
                                        "Nota de Crédito A"
                                    )
                            elif (
                                "FACTURA B" in val_upper
                                or "FAC B" in val_upper
                                or "FB" in val_upper
                            ):
                                tipo_comprobante_actual = "Factura B"
                            elif (
                                "FACTURA A" in val_upper
                                or "FAC A" in val_upper
                                or "FA" in val_upper
                            ):
                                tipo_comprobante_actual = "Factura A"

        except Exception as e:
            st.error(f"Error procesando el archivo '{archivo.name}': {e}")

    if not todas_facturas:
        return pd.DataFrame()

    df_resultado = pd.DataFrame(todas_facturas)
    df_resultado.sort_values(
        by=["punto_venta", "TipoComprobante", "numero_factura"], inplace=True
    )
    df_resultado.reset_index(drop=True, inplace=True)

    return df_resultado


def auditar_correlatividad(df_facturas):
    """Audita correlatividad agrupando estrictamente por PDV y Tipo de Comprobante."""
    anomalias = []
    COLUMNAS_ANOMALIAS = [
        "Tipo de Anomalía",
        "PDV",
        "Comprobante",
        "N° de Factura",
        "Detalle de la Observación",
    ]

    if df_facturas.empty:
        return pd.DataFrame(columns=COLUMNAS_ANOMALIAS), set()

    indices_anomalos = set()
    grupos = df_facturas.groupby(["punto_venta", "TipoComprobante"])

    for (pdv, tipo_comp), grupo in grupos:
        grupo = grupo.sort_values("numero_factura").reset_index()

        for i in range(len(grupo)):
            row_actual = grupo.iloc[i]
            idx_orig = row_actual["index"]

            # 1. Comprobantes Duplicados
            if i > 0:
                row_prev = grupo.iloc[i - 1]
                if row_actual["numero_factura"] == row_prev["numero_factura"]:
                    indices_anomalos.add(idx_orig)
                    indices_anomalos.add(row_prev["index"])
                    anomalias.append({
                        "Tipo de Anomalía": "Comprobante Duplicado",
                        "PDV": pdv,
                        "Comprobante": tipo_comp,
                        "N° de Factura": row_actual["N°Factura"],
                        "Detalle de la Observación": (
                            f"Número duplicado en {tipo_comp} con fila"
                            f" {row_prev['fila_excel']} ({row_prev['archivo_origen']})"
                        ),
                    })

                # 2. Salto de Numeración
                salto = (
                    row_actual["numero_factura"] - row_prev["numero_factura"]
                )
                if salto > 1:
                    indices_anomalos.add(idx_orig)
                    indices_anomalos.add(row_prev["index"])
                    faltantes = salto - 1
                    anomalias.append({
                        "Tipo de Anomalía": "Salto de Numeración",
                        "PDV": pdv,
                        "Comprobante": tipo_comp,
                        "N° de Factura": row_actual["N°Factura"],
                        "Detalle de la Observación": (
                            f"Faltan {faltantes} {tipo_comp}(s) entre"
                            f" {row_prev['N°Factura']} y {row_actual['N°Factura']}"
                        ),
                    })

                # 3. Inconsistencia de Fecha Cronológica
                if (
                    pd.notna(row_actual["fecha_emision_dt"])
                    and pd.notna(row_prev["fecha_emision_dt"])
                    and row_actual["fecha_emision_dt"]
                    < row_prev["fecha_emision_dt"]
                ):
                    indices_anomalos.add(idx_orig)
                    indices_anomalos.add(row_prev["index"])
                    anomalias.append({
                        "Tipo de Anomalía": "Inconsistencia de Fecha",
                        "PDV": pdv,
                        "Comprobante": tipo_comp,
                        "N° de Factura": row_actual["N°Factura"],
                        "Detalle de la Observación": (
                            f"Fecha ({row_actual['Fecha']}) anterior a la del"
                            f" comprobante previo ({row_prev['Fecha']})"
                        ),
                    })

    df_anomalias = (
        pd.DataFrame(anomalias)
        if anomalias
        else pd.DataFrame(columns=COLUMNAS_ANOMALIAS)
    )
    return df_anomalias, indices_anomalos


def generar_resumen_cierre(df_facturas):
    """Genera el registro de cierre completo por cada Punto de Venta."""
    if df_facturas.empty:
        return pd.DataFrame()

    pdvs = sorted(df_facturas["punto_venta"].unique())
    filas_resumen = []

    for pdv in pdvs:
        df_pdv = df_facturas[df_facturas["punto_venta"] == pdv]

        for tipo in TIPOS_COMPROBANTES_ESTANDAR:
            df_tipo = df_pdv[df_pdv["TipoComprobante"] == tipo]

            if not df_tipo.empty:
                df_tipo_sorted = df_tipo.sort_values("numero_factura")
                ultimo_row = df_tipo_sorted.iloc[-1]

                ultimo_num = int(ultimo_row["numero_factura"])
                proximo_num = ultimo_num + 1
                proximo_factura_str = f"{pdv}-{str(proximo_num).zfill(8)}"

                filas_resumen.append({
                    "PDV": pdv,
                    "Tipo Comprobante": tipo,
                    "Último N° Emitido": ultimo_row["N°Factura"],
                    "Próximo N° Esperado": proximo_factura_str,
                    "Fecha Última Emisión": ultimo_row["Fecha"],
                    "Cant. Comprobantes": len(df_tipo),
                    "Monto Total ($)": df_tipo["Total"].sum(),
                    "Estado": "🟢 Con Movimiento",
                })
            else:
                filas_resumen.append({
                    "PDV": pdv,
                    "Tipo Comprobante": tipo,
                    "Último N° Emitido": "Sin emisiones en el período",
                    "Próximo N° Esperado": "A determinar",
                    "Fecha Última Emisión": "-",
                    "Cant. Comprobantes": 0,
                    "Monto Total ($)": 0.0,
                    "Estado": "⚪ Sin Movimiento",
                })

    return pd.DataFrame(filas_resumen)


def exportar_excel_bytes(df, nombre_hoja="Reporte"):
    """Convierte un DataFrame a un archivo Excel (.xls) binario para descargar."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=nombre_hoja)
    return output.getvalue()


# -----------------------------------------------------------------------------
# COMPONENTE DE TABLA DESPLEGABLE (HTML + JS)
# -----------------------------------------------------------------------------

def renderizar_tabla_interactiva(df_base, pdv_key):
    """Renderea una tabla HTML/JS interactiva con botones +/- y despliegue de anomalías."""
    if df_base.empty:
        st.info("No hay registros para mostrar.")
        return

    filas_html = []
    for idx, row in df_base.iterrows():
        es_anomalo = bool(row.get("es_anomalo", False))
        obs = str(row.get("observacion", ""))
        nro_fact = str(row.get("N°Factura", ""))
        tipo = str(row.get("TipoComprobante", ""))
        fecha = str(row.get("Fecha", ""))
        cliente = str(row.get("Cliente", ""))
        cuit = str(row.get("Cuit", ""))
        comanda = str(row.get("Número de Comanda", ""))
        total_val = row.get("Total", 0.0)
        
        try:
            total_str = f"$ {float(total_val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except (ValueError, TypeError):
            total_str = "$ 0,00"

        row_id = f"r_{pdv_key}_{idx}"
        detail_id = f"d_{pdv_key}_{idx}"

        btn_symbol = "-" if es_anomalo else "+"
        btn_class = "btn-toggle expanded" if es_anomalo else "btn-toggle"
        tr_class = "row-anomaly" if es_anomalo else ""
        detail_display_class = "open" if es_anomalo else ""

        badge_html = ' <span class="badge-anomaly">⚠️ ANOMALÍA DETECTADA</span>' if es_anomalo else ''

        row_html = f"""
        <tr class="{tr_class}">
            <td style="text-align: center; width: 45px;">
                <button id="btn_{row_id}" class="{btn_class}" onclick="toggleRow('{row_id}', '{detail_id}')">{btn_symbol}</button>
            </td>
            <td><strong>{nro_fact}</strong>{badge_html}</td>
            <td>{tipo}</td>
            <td>{fecha}</td>
            <td>{cliente}</td>
            <td>{cuit}</td>
            <td>{comanda}</td>
            <td style="text-align: right; font-weight: 600;">{total_str}</td>
        </tr>
        """

        if es_anomalo and obs:
            obs_clean = obs.replace("⚠️ ", "")
            row_html += f"""
            <tr id="{detail_id}" class="detail-row {detail_display_class}">
                <td colspan="8" style="padding: 0; border: none;">
                    <div class="detail-box">
                        <div class="detail-title">REPORTE DETALLADO DE ANOMALÍA</div>
                        <div>{obs_clean}</div>
                    </div>
                </td>
            </tr>
            """
        else:
            row_html += f"""
            <tr id="{detail_id}" class="detail-row">
                <td colspan="8" style="padding: 0; border: none;">
                    <div class="detail-box-normal">
                        <div><strong>Detalle del Registro:</strong> Comprobante verificado correctamente. Sin observaciones de correlatividad.</div>
                    </div>
                </td>
            </tr>
            """

        filas_html.append(row_html)

    filas_joined = "\n".join(filas_html)

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        * {{ box-sizing: border-box; }}
        body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; margin: 0; padding: 2px; background: transparent; }}
        .table-container {{ border: 1px solid #d1d5db; border-radius: 8px; overflow: hidden; background: #ffffff; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }}
        table {{ width: 100%; border-collapse: collapse; font-size: 13px; color: #1f2937; }}
        th {{ background-color: #f3f4f6; text-align: left; padding: 10px 12px; font-weight: 700; color: #374151; border-bottom: 1px solid #e5e7eb; font-size: 12px; text-transform: uppercase; letter-spacing: 0.03em; }}
        td {{ padding: 10px 12px; border-bottom: 1px solid #f3f4f6; vertical-align: middle; }}
        .row-anomaly {{ background-color: #fde8e8 !important; color: #991b1b; }}
        .btn-toggle {{ width: 24px; height: 24px; border-radius: 4px; border: none; background-color: #2563eb; color: white; font-weight: bold; cursor: pointer; display: inline-flex; align-items: center; justify-content: center; font-size: 15px; line-height: 1; transition: all 0.15s ease; }}
        .btn-toggle.expanded {{ background-color: #dc2626; }}
        .btn-toggle:hover {{ opacity: 0.85; transform: scale(1.05); }}
        .badge-anomaly {{ background-color: #dc2626; color: white; padding: 2px 6px; border-radius: 4px; font-weight: bold; font-size: 10px; margin-left: 6px; display: inline-block; text-transform: uppercase; }}
        .detail-row {{ display: none; }}
        .detail-row.open {{ display: table-row; }}
        .detail-box {{ background-color: #fecaca; border: 1px solid #f87171; border-radius: 6px; padding: 12px 16px; margin: 8px 12px 12px 12px; color: #7f1d1d; font-size: 13px; line-height: 1.4; }}
        .detail-box-normal {{ background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 10px 14px; margin: 8px 12px 12px 12px; color: #475569; font-size: 12px; }}
        .detail-title {{ font-weight: 800; margin-bottom: 4px; text-transform: uppercase; font-size: 11px; color: #991b1b; letter-spacing: 0.05em; }}
    </style>
    </head>
    <body>
        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th style="text-align: center;">Status</th>
                        <th>N° Factura</th>
                        <th>Tipo</th>
                        <th>Fecha</th>
                        <th>Cliente</th>
                        <th>CUIT</th>
                        <th>Comanda</th>
                        <th style="text-align: right;">Total</th>
                    </tr>
                </thead>
                <tbody>
                    {filas_joined}
                </tbody>
            </table>
        </div>

        <script>
            function toggleRow(rowId, detailId) {{
                var detail = document.getElementById(detailId);
                var btn = document.getElementById('btn_' + rowId);
                if (detail.classList.contains('open')) {{
                    detail.classList.remove('open');
                    btn.classList.remove('expanded');
                    btn.textContent = '+';
                }} else {{
                    detail.classList.add('open');
                    btn.classList.add('expanded');
                    btn.textContent = '-';
                }}
            }}
        </script>
    </body>
    </html>
    """

    altura_calculada = max(300, min(1500, len(df_base) * 55 + 120))
    components.html(html_content, height=altura_calculada, scrolling=True)


# -----------------------------------------------------------------------------
# INTERFAZ PRINCIPAL Y CONTROL DE ESTADO
# -----------------------------------------------------------------------------

if "collapsed_invoices" not in st.session_state:
    st.session_state.collapsed_invoices = set()
if "expanded_invoices" not in st.session_state:
    st.session_state.expanded_invoices = set()


def set_vista_pdv(pdv, vista):
    st.session_state[f"vista_{pdv}"] = vista


st.title("📊 Gestor y Auditor de Comprobantes")

archivos_subidos = st.sidebar.file_uploader(
    "Cargar reportes de facturación (.xls / .xlsx)",
    type=["xls", "xlsx"],
    accept_multiple_files=True,
)

if not archivos_subidos:
    st.info(
        "👈 Por favor, carga uno o varios archivos Excel en el menú lateral"
        " para iniciar la auditoría."
    )
else:
    df_original = procesar_multiples_reportes(archivos_subidos)

    if df_original.empty:
        st.warning(
            "No se encontraron comprobantes con el formato esperado en los"
            " archivos cargados."
        )
    else:
        df_anomalias, indices_anomalos = auditar_correlatividad(df_original)

        df_original["es_anomalo"] = df_original.index.isin(indices_anomalos)
        mapa_obs = {}
        if not df_anomalias.empty:
            for _, r in df_anomalias.iterrows():
                mapa_obs[r["N° de Factura"]] = (
                    f"⚠️ [{r['Tipo de Anomalía']}] {r['Detalle de la Observación']}"
                )
        df_original["observacion"] = df_original["N°Factura"].map(
            lambda x: mapa_obs.get(x, "")
        )

        pdvs_detectados = sorted(df_original["punto_venta"].unique().tolist())
        nombres_pestañas = (
            ["🌐 General (Todos)", "📌 Cierre / Últimos Emitidos"]
            + [f"📍 PDV {pdv}" for pdv in pdvs_detectados]
        )
        pestañas = st.tabs(nombres_pestañas)

        def renderizar_tab(pdv_actual, df_orig_tab, df_anom_tab):
            est_list = [
                x
                for x in df_orig_tab["establecimiento"].unique()
                if x != "No Especificado"
            ]
            est_txt = " / ".join(est_list) if est_list else "No Especificado"

            per_list = [
                x
                for x in df_orig_tab["periodo"].unique()
                if x != "No Especificado"
            ]
            per_txt = " / ".join(per_list) if per_list else "No Especificado"

            col_a, col_b = st.columns(2)
            col_a.metric("🏢 ESTABLECIMIENTO", est_txt)
            col_b.metric("📅 PERÍODO", per_txt)

            sin_fallas = (
                df_anom_tab.empty
                if not df_anom_tab.empty
                else not df_orig_tab["es_anomalo"].any()
            )
            if sin_fallas:
                st.success(
                    "✅ No se encontraron anomalías de correlatividad para"
                    " estos registros."
                )

            key_vista = f"vista_{pdv_actual}"
            if key_vista not in st.session_state:
                st.session_state[key_vista] = "completo"
            vista_actual = st.session_state[key_vista]

            c1, c2, c3 = st.columns(3)
            c1.button(
                "📋 Ver listado completo",
                key=f"btn_comp_{pdv_actual}",
                on_click=set_vista_pdv,
                args=(pdv_actual, "completo"),
                use_container_width=True,
            )
            c2.button(
                "🚨 Ver solo anomalías",
                key=f"btn_anom_{pdv_actual}",
                on_click=set_vista_pdv,
                args=(pdv_actual, "anomalias"),
                use_container_width=True,
            )
            c3.button(
                "🔄 Vista Intercalada",
                key=f"btn_inter_{pdv_actual}",
                on_click=set_vista_pdv,
                args=(pdv_actual, "intercalada"),
                use_container_width=True,
            )

            if vista_actual == "anomalias":
                df_base = df_orig_tab[df_orig_tab["es_anomalo"]].copy()
            else:
                df_base = df_orig_tab.copy()

            if df_base.empty:
                if vista_actual == "anomalias":
                    st.success(
                        "✅ No se encontraron anomalías para este punto de"
                        " venta."
                    )
                else:
                    st.info("No hay registros para mostrar en esta vista.")
                return

            cols_mostrar = [
                "N°Factura",
                "TipoComprobante",
                "Fecha",
                "Cliente",
                "Cuit",
                "Número de Comanda",
                "Total",
            ]

            # Botón de Descarga directa en Excel (.xls)
            excel_bytes = exportar_excel_bytes(
                df_base[cols_mostrar],
                nombre_hoja=f"PDV_{pdv_actual}",
            )
            st.download_button(
                label="📥 Descargar Tabla en Excel (.xls)",
                data=excel_bytes,
                file_name=f"Reporte_Comprobantes_{pdv_actual}.xls",
                mime="application/vnd.ms-excel",
                key=f"dl_btn_{pdv_actual}_{vista_actual}",
            )

            # Renderizado de la tabla con filas desplegables y botones (+)/(-)
            renderizar_tabla_interactiva(df_base, pdv_actual)

        # Pestaña General
        with pestañas[0]:
            renderizar_tab("TODOS", df_original, df_anomalias)

        # Pestaña de Registro de Cierre
        with pestañas[1]:
            st.subheader(
                "📌 Registro de Cierre y Próximo N° Esperado por PDV"
            )
            st.markdown(
                "Este cuadro registra el **último comprobante emitido** de cada"
                " tipo y determina el **próximo número a emitir** para tener"
                " en cuenta al cargar el siguiente reporte."
            )

            df_cierre = generar_resumen_cierre(df_original)

            bytes_cierre = exportar_excel_bytes(
                df_cierre, nombre_hoja="Cierre_Facturacion"
            )
            st.download_button(
                label="📥 Descargar Cierre en Excel (.xls)",
                data=bytes_cierre,
                file_name="Registro_Cierre_Facturacion.xls",
                mime="application/vnd.ms-excel",
                key="dl_btn_cierre",
            )

            st.dataframe(
                df_cierre,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Monto Total ($)": st.column_config.NumberColumn(
                        "Monto Total ($)", format="$ %.2f"
                    )
                },
            )

        # Pestañas individuales por PDV
        for i, pdv_actual in enumerate(pdvs_detectados):
            with pestañas[i + 2]:
                df_orig_tab = df_original[
                    df_original["punto_venta"] == pdv_actual
                ]
                df_anom_tab = (
                    df_anomalias[df_anomalias["PDV"] == pdv_actual]
                    if not df_anomalias.empty
                    else pd.DataFrame()
                )
                renderizar_tab(pdv_actual, df_orig_tab, df_anom_tab)