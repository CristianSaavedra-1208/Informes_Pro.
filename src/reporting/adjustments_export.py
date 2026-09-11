import io
from datetime import datetime
from collections import defaultdict
import pandas as pd
from sqlalchemy import or_

from src.models.database import SessionLocal
from src.models.consolidacion import ConsolidationGroup, ConsolidationJournalEntry
from src.core.consolidacion_engine import resolver_montos_asiento
from src.core.excel_utils import format_periodo

def get_adjustments_dataset(
    grupo_id: int,
    filter_mode: str = "ALL", # "SINGLE_VOUCHER", "SINGLE_MONTH", "RANGE", "ALL"
    periodo: str = None,
    periodo_inicio: str = None,
    periodo_fin: str = None,
    voucher_codigo: str = None,
    db = None
):
    """
    Extrae y resuelve los montos de ajustes contables de consolidación sin mutar la base de datos.
    Retorna un diccionario con metadatos del grupo, lista de líneas detalladas y resúmenes.
    """
    is_local_db = False
    if db is None:
        db = SessionLocal()
        is_local_db = True

    try:
        grupo = db.query(ConsolidationGroup).filter_by(id=grupo_id).first()
        grupo_nombre = grupo.nombre_grupo if grupo else f"Grupo #{grupo_id}"
        empresa_matriz = grupo.empresa_matriz if grupo else "Matriz"
        empresa_filial = grupo.empresa_filial if grupo else "Filial"

        query = db.query(ConsolidationJournalEntry).filter(ConsolidationJournalEntry.grupo_id == grupo_id)

        # Aplicar filtros según modalidad
        filter_mode_clean = str(filter_mode).upper().strip()
        filtro_desc = "Todos los períodos históricos"

        if filter_mode_clean == "SINGLE_VOUCHER":
            if voucher_codigo:
                query = query.filter(ConsolidationJournalEntry.asiento_codigo == voucher_codigo)
                filtro_desc = f"Comprobante Folio {voucher_codigo}"
            elif periodo:
                query = query.filter(ConsolidationJournalEntry.periodo == periodo)
                filtro_desc = f"Comprobante en {format_periodo(periodo)}"
        elif filter_mode_clean == "SINGLE_MONTH":
            if periodo:
                query = query.filter(ConsolidationJournalEntry.periodo == periodo)
                filtro_desc = f"Mes {format_periodo(periodo)}"
        elif filter_mode_clean == "RANGE":
            p_ini = periodo_inicio or "1900-01"
            p_fin = periodo_fin or "2099-12"
            query = query.filter(
                ConsolidationJournalEntry.periodo >= p_ini,
                ConsolidationJournalEntry.periodo <= p_fin
            )
            filtro_desc = f"Rango desde {format_periodo(p_ini)} hasta {format_periodo(p_fin)}"
        else: # ALL
            filtro_desc = "Historial Completo (Todos los Períodos)"

        raw_entries = query.order_by(
            ConsolidationJournalEntry.periodo.asc(),
            ConsolidationJournalEntry.asiento_codigo.asc(),
            ConsolidationJournalEntry.num_linea.asc(),
            ConsolidationJournalEntry.id.asc()
        ).all()

        if filter_mode_clean == "SINGLE_VOUCHER" and raw_entries:
            first_e = raw_entries[0]
            code_show = getattr(first_e, 'asiento_codigo', None) or voucher_codigo or f"AST-{first_e.periodo}"
            filtro_desc = f"Comprobante [{code_show}] ({format_periodo(first_e.periodo)}) {first_e.columna_ajuste} — {first_e.glosa}"

        if not raw_entries:
            return {
                "grupo_id": grupo_id,
                "grupo_nombre": grupo_nombre,
                "empresa_matriz": empresa_matriz,
                "empresa_filial": empresa_filial,
                "filtro_desc": filtro_desc,
                "filter_mode": filter_mode_clean,
                "periodo": periodo,
                "periodo_inicio": periodo_inicio,
                "periodo_fin": periodo_fin,
                "voucher_codigo": voucher_codigo,
                "lineas": [],
                "resumen_comprobantes": [],
                "resumen_tipos": [],
                "total_debe": 0.0,
                "total_haber": 0.0,
                "cuadrado": True
            }

        # Agrupar entradas por comprobante para resolver eliminaciones dinámicas adecuadamente
        # Clave del comprobante: (periodo, columna_ajuste, glosa, asiento_codigo)
        vouchers_dict = defaultdict(list)
        for e in raw_entries:
            v_key = (
                e.periodo,
                e.columna_ajuste,
                e.glosa,
                getattr(e, 'asiento_codigo', None) or f"AST-{e.periodo}-{e.id}"
            )
            vouchers_dict[v_key].append(e)

        lineas_detalladas = []
        resumen_comprobantes = []
        resumen_tipos_dict = defaultdict(lambda: {"comprobantes": set(), "lineas": 0, "debe": 0.0, "haber": 0.0})

        total_debe_general = 0.0
        total_haber_general = 0.0

        for (p_val, col_val, glosa_val, codigo_val), entries_list in vouchers_dict.items():
            # Resolver montos reales para el comprobante
            try:
                resolved = resolver_montos_asiento(grupo_id, p_val, entries_list, db=db, columna_destino=col_val)
            except Exception:
                resolved = [{
                    "debe_calculado": e.debe or 0.0,
                    "haber_calculado": e.haber or 0.0,
                    "saldo_base": 0.0,
                    "elimina_saldo_total": e.elimina_saldo_total
                } for e in entries_list]

            comp_total_debe = 0.0
            comp_total_haber = 0.0
            comp_fecha = None
            comp_usuario = None
            comp_recurrente = False

            for idx, e in enumerate(entries_list):
                res_line = resolved[idx] if idx < len(resolved) else {}
                debe_val = float(res_line.get("debe_calculado", e.debe or 0.0))
                haber_val = float(res_line.get("haber_calculado", e.haber or 0.0))
                saldo_base = float(res_line.get("saldo_base", 0.0))

                comp_total_debe += debe_val
                comp_total_haber += haber_val
                comp_fecha = e.created_at or e.fecha or comp_fecha
                comp_usuario = e.created_by or comp_usuario
                if e.es_recurrente:
                    comp_recurrente = True

                tipo_calc = "Eliminación Dinámica (100%)" if e.elimina_saldo_total else "Monto Fijo"
                fecha_str = comp_fecha.strftime("%d-%m-%Y %H:%M") if isinstance(comp_fecha, datetime) else str(comp_fecha or "")

                lineas_detalladas.append({
                    "Código Folio": codigo_val,
                    "Período": p_val,
                    "Mes / Año": format_periodo(p_val),
                    "Fecha Registro": fecha_str,
                    "Tipo de Ajuste": col_val,
                    "Glosa / Explicación": glosa_val,
                    "Línea #": getattr(e, "num_linea", idx + 1) or (idx + 1),
                    "Rubro EEFF Afectado": e.linea_item,
                    "Nota / Desglose": getattr(e, "linea_nota", None) or "Sin Detalle",
                    "Tipo Cálculo": tipo_calc,
                    "Saldo Base": saldo_base,
                    "Debe ($)": debe_val,
                    "Haber ($)": haber_val,
                    "Recurrente": "Sí" if e.es_recurrente else "No",
                    "Registrado Por": getattr(e, "created_by", "Sistema") or "Sistema",
                    "Modificado Por": getattr(e, "updated_by", None) or ""
                })

            diff_comp = round(comp_total_debe - comp_total_haber, 2)
            resumen_comprobantes.append({
                "Código Folio": codigo_val,
                "Período": p_val,
                "Mes / Año": format_periodo(p_val),
                "Tipo de Ajuste": col_val,
                "Glosa / Justificación": glosa_val,
                "N° Líneas": len(entries_list),
                "Total Debe ($)": comp_total_debe,
                "Total Haber ($)": comp_total_haber,
                "Diferencia": diff_comp,
                "Estado": "✓ Cuadrado" if abs(diff_comp) < 0.01 else "⚠️ Descuadrado",
                "Recurrente": "Sí" if comp_recurrente else "No",
                "Registrado Por": comp_usuario or "Sistema"
            })

            res_tipo = resumen_tipos_dict[col_val]
            res_tipo["comprobantes"].add(codigo_val)
            res_tipo["lineas"] += len(entries_list)
            res_tipo["debe"] += comp_total_debe
            res_tipo["haber"] += comp_total_haber

            total_debe_general += comp_total_debe
            total_haber_general += comp_total_haber

        resumen_tipos = []
        for col_name, data in sorted(resumen_tipos_dict.items()):
            diff_t = round(data["debe"] - data["haber"], 2)
            resumen_tipos.append({
                "Tipo de Ajuste / Columna": col_name,
                "N° Comprobantes": len(data["comprobantes"]),
                "N° Líneas": data["lineas"],
                "Total Debe ($)": data["debe"],
                "Total Haber ($)": data["haber"],
                "Diferencia": diff_t,
                "Estado": "✓ Cuadrado" if abs(diff_t) < 0.01 else "⚠️ Descuadrado"
            })

        diff_general = round(total_debe_general - total_haber_general, 2)

        return {
            "grupo_id": grupo_id,
            "grupo_nombre": grupo_nombre,
            "empresa_matriz": empresa_matriz,
            "empresa_filial": empresa_filial,
            "filtro_desc": filtro_desc,
            "filter_mode": filter_mode_clean,
            "periodo": periodo,
            "periodo_inicio": periodo_inicio,
            "periodo_fin": periodo_fin,
            "voucher_codigo": voucher_codigo,
            "lineas": lineas_detalladas,
            "resumen_comprobantes": resumen_comprobantes,
            "resumen_tipos": resumen_tipos,
            "total_debe": total_debe_general,
            "total_haber": total_haber_general,
            "diferencia": diff_general,
            "cuadrado": abs(diff_general) < 0.01
        }

    finally:
        if is_local_db:
            db.close()


def generate_adjustments_excel(
    grupo_id: int,
    filter_mode: str = "ALL",
    periodo: str = None,
    periodo_inicio: str = None,
    periodo_fin: str = None,
    voucher_codigo: str = None,
    db = None
) -> bytes:
    """
    Genera un archivo Excel (.xlsx) con formato contable corporativo y metadatos de auditoría.
    Retorna los bytes del archivo listos para descarga o almacenamiento.
    """
    dataset = get_adjustments_dataset(
        grupo_id=grupo_id,
        filter_mode=filter_mode,
        periodo=periodo,
        periodo_inicio=periodo_inicio,
        periodo_fin=periodo_fin,
        voucher_codigo=voucher_codigo,
        db=db
    )

    output = io.BytesIO()

    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    # Paleta de estilos corporativos
    HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    SUBHEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    TOTAL_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    WHITE_HEADER_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="1F4E78")
    SUBTITLE_FONT = Font(name="Calibri", size=10, italic=True, color="595959")
    BOLD_FONT = Font(name="Calibri", size=10, bold=True)
    NORMAL_FONT = Font(name="Calibri", size=10)
    SUCCESS_FONT = Font(name="Calibri", size=10, bold=True, color="385723")
    ALERT_FONT = Font(name="Calibri", size=10, bold=True, color="C00000")

    THIN_BORDER = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    TOTAL_BORDER = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='double', color='000000')
    )

    NUMBER_FMT = r'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1. PESTAÑA PRINCIPAL: LIBRO DIARIO DE AJUSTES (DETALLE)
        df_lineas = pd.DataFrame(dataset["lineas"])
        if df_lineas.empty:
            df_lineas = pd.DataFrame(columns=[
                "Código Folio", "Período", "Mes / Año", "Fecha Registro", "Tipo de Ajuste",
                "Glosa / Explicación", "Línea #", "Rubro EEFF Afectado", "Nota / Desglose",
                "Tipo Cálculo", "Saldo Base", "Debe ($)", "Haber ($)", "Recurrente", "Registrado Por"
            ])

        # Escribir con inicio en fila 6 para dejar espacio al banner de auditoría
        df_lineas.to_excel(writer, sheet_name="Libro de Ajustes", index=False, startrow=5)
        ws_det = writer.sheets["Libro de Ajustes"]

        # Encabezado corporativo
        ws_det["A1"] = "INFORMES PRO - MÓDULO DE CONSOLIDACIÓN FINANCIERA"
        ws_det["A2"] = "LIBRO DIARIO DE AJUSTES Y ELIMINACIONES DE CONSOLIDACIÓN"
        ws_det["A3"] = f"Grupo: {dataset['grupo_nombre']} | Matriz: {dataset['empresa_matriz']} | Filial: {dataset['empresa_filial']}"
        ws_det["A4"] = f"Alcance: {dataset['filtro_desc']} | Fecha de Emisión: {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}"

        ws_det["A1"].font = Font(name="Calibri", size=9, bold=True, color="595959")
        ws_det["A2"].font = TITLE_FONT
        ws_det["A3"].font = BOLD_FONT
        ws_det["A4"].font = SUBTITLE_FONT

        # Estilo de encabezado de tabla (Fila 6)
        header_row = 6
        ws_det.row_dimensions[header_row].height = 24
        for col_idx in range(1, len(df_lineas.columns) + 1):
            cell = ws_det.cell(row=header_row, column=col_idx)
            cell.fill = HEADER_FILL
            cell.font = WHITE_HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = THIN_BORDER

        # Formatear filas de datos
        start_data = 7
        end_data = start_data + len(df_lineas) - 1

        for r_idx in range(start_data, end_data + 1):
            ws_det.row_dimensions[r_idx].height = 19
            for c_idx in range(1, len(df_lineas.columns) + 1):
                cell = ws_det.cell(row=r_idx, column=c_idx)
                cell.font = NORMAL_FONT
                cell.border = THIN_BORDER

                col_name = df_lineas.columns[c_idx - 1]
                if col_name in ["Debe ($)", "Haber ($)", "Saldo Base"]:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = NUMBER_FMT
                elif col_name in ["Línea #"]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_name in ["Código Folio", "Período", "Recurrente"]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

        # Fila de Totales Generales
        if len(df_lineas) > 0:
            tot_row = end_data + 1
            ws_det.row_dimensions[tot_row].height = 22
            ws_det.cell(row=tot_row, column=1, value="TOTAL GENERAL").font = BOLD_FONT
            
            # Buscar columnas de Debe y Haber
            debe_col_idx = list(df_lineas.columns).index("Debe ($)") + 1
            haber_col_idx = list(df_lineas.columns).index("Haber ($)") + 1

            debe_letter = get_column_letter(debe_col_idx)
            haber_letter = get_column_letter(haber_col_idx)

            cell_tot_debe = ws_det.cell(row=tot_row, column=debe_col_idx)
            cell_tot_debe.value = f"=SUM({debe_letter}{start_data}:{debe_letter}{end_data})"
            cell_tot_debe.number_format = NUMBER_FMT
            cell_tot_debe.font = BOLD_FONT
            cell_tot_debe.alignment = Alignment(horizontal="right", vertical="center")

            cell_tot_haber = ws_det.cell(row=tot_row, column=haber_col_idx)
            cell_tot_haber.value = f"=SUM({haber_letter}{start_data}:{haber_letter}{end_data})"
            cell_tot_haber.number_format = NUMBER_FMT
            cell_tot_haber.font = BOLD_FONT
            cell_tot_haber.alignment = Alignment(horizontal="right", vertical="center")

            for c_idx in range(1, len(df_lineas.columns) + 1):
                cell = ws_det.cell(row=tot_row, column=c_idx)
                cell.fill = TOTAL_FILL
                cell.border = TOTAL_BORDER

        # Autoajustar ancho de columnas en hoja de detalle
        for col in ws_det.columns:
            col_letter = col[0].column_letter
            max_len = 0
            for cell in col:
                if cell.row < 6:
                    continue
                val_str = str(cell.value or "")
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws_det.column_dimensions[col_letter].width = max(max_len + 4, 12)

        # 2. PESTAÑA DE RESUMEN DE COMPROBANTES (si hay datos)
        if dataset["resumen_comprobantes"]:
            df_res_comp = pd.DataFrame(dataset["resumen_comprobantes"])
            df_res_comp.to_excel(writer, sheet_name="Resumen Comprobantes", index=False, startrow=4)
            ws_comp = writer.sheets["Resumen Comprobantes"]

            ws_comp["A1"] = "INFORMES PRO - RESUMEN DE COMPROBANTES DE AJUSTE"
            ws_comp["A2"] = f"Grupo: {dataset['grupo_nombre']} | Alcance: {dataset['filtro_desc']}"
            ws_comp["A3"] = f"Total Comprobantes: {len(df_res_comp)} | Emitido el: {datetime.now().strftime('%d-%m-%Y %H:%M')}"

            ws_comp["A1"].font = TITLE_FONT
            ws_comp["A2"].font = BOLD_FONT
            ws_comp["A3"].font = SUBTITLE_FONT

            comp_header_row = 5
            ws_comp.row_dimensions[comp_header_row].height = 22
            for col_idx in range(1, len(df_res_comp.columns) + 1):
                cell = ws_comp.cell(row=comp_header_row, column=col_idx)
                cell.fill = HEADER_FILL
                cell.font = WHITE_HEADER_FONT
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = THIN_BORDER

            comp_start = 6
            comp_end = comp_start + len(df_res_comp) - 1
            for r_idx in range(comp_start, comp_end + 1):
                ws_comp.row_dimensions[r_idx].height = 19
                for c_idx in range(1, len(df_res_comp.columns) + 1):
                    cell = ws_comp.cell(row=r_idx, column=c_idx)
                    cell.font = NORMAL_FONT
                    cell.border = THIN_BORDER
                    c_name = df_res_comp.columns[c_idx - 1]
                    if "Total" in c_name or "Diferencia" in c_name:
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                        cell.number_format = NUMBER_FMT
                    elif c_name in ["Código Folio", "Período", "N° Líneas", "Recurrente"]:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    elif c_name == "Estado":
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.font = SUCCESS_FONT if "Cuadrado" in str(cell.value) else ALERT_FONT
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")

            # Fila de totales en resumen
            tot_comp_row = comp_end + 1
            ws_comp.row_dimensions[tot_comp_row].height = 22
            ws_comp.cell(row=tot_comp_row, column=1, value="TOTAL GENERAL").font = BOLD_FONT

            debe_c_idx = list(df_res_comp.columns).index("Total Debe ($)") + 1
            haber_c_idx = list(df_res_comp.columns).index("Total Haber ($)") + 1
            debe_l = get_column_letter(debe_c_idx)
            haber_l = get_column_letter(haber_c_idx)

            c_debe = ws_comp.cell(row=tot_comp_row, column=debe_c_idx)
            c_debe.value = f"=SUM({debe_l}{comp_start}:{debe_l}{comp_end})"
            c_debe.number_format = NUMBER_FMT
            c_debe.font = BOLD_FONT
            c_debe.alignment = Alignment(horizontal="right", vertical="center")

            c_haber = ws_comp.cell(row=tot_comp_row, column=haber_c_idx)
            c_haber.value = f"=SUM({haber_l}{comp_start}:{haber_l}{comp_end})"
            c_haber.number_format = NUMBER_FMT
            c_haber.font = BOLD_FONT
            c_haber.alignment = Alignment(horizontal="right", vertical="center")

            for c_idx in range(1, len(df_res_comp.columns) + 1):
                c = ws_comp.cell(row=tot_comp_row, column=c_idx)
                c.fill = TOTAL_FILL
                c.border = TOTAL_BORDER

            for col in ws_comp.columns:
                col_letter = col[0].column_letter
                max_len = 0
                for cell in col:
                    if cell.row < 5:
                        continue
                    val_str = str(cell.value or "")
                    if len(val_str) > max_len:
                        max_len = len(val_str)
                ws_comp.column_dimensions[col_letter].width = max(max_len + 4, 14)

        # 3. PESTAÑA DE RESUMEN POR TIPO DE AJUSTE (si hay más de 1 tipo o modo rango/all)
        if dataset["resumen_tipos"] and len(dataset["resumen_tipos"]) > 1:
            df_res_tipos = pd.DataFrame(dataset["resumen_tipos"])
            df_res_tipos.to_excel(writer, sheet_name="Resumen por Tipo", index=False, startrow=4)
            ws_tipos = writer.sheets["Resumen por Tipo"]

            ws_tipos["A1"] = "INFORMES PRO - RESUMEN POR CATEGORÍA DE AJUSTE"
            ws_tipos["A2"] = f"Grupo: {dataset['grupo_nombre']} | Alcance: {dataset['filtro_desc']}"
            ws_tipos["A3"] = f"Total Categorías: {len(df_res_tipos)} | Emitido el: {datetime.now().strftime('%d-%m-%Y %H:%M')}"

            ws_tipos["A1"].font = TITLE_FONT
            ws_tipos["A2"].font = BOLD_FONT
            ws_tipos["A3"].font = SUBTITLE_FONT

            tipo_header_row = 5
            ws_tipos.row_dimensions[tipo_header_row].height = 22
            for col_idx in range(1, len(df_res_tipos.columns) + 1):
                cell = ws_tipos.cell(row=tipo_header_row, column=col_idx)
                cell.fill = HEADER_FILL
                cell.font = WHITE_HEADER_FONT
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = THIN_BORDER

            tipo_start = 6
            tipo_end = tipo_start + len(df_res_tipos) - 1
            for r_idx in range(tipo_start, tipo_end + 1):
                ws_tipos.row_dimensions[r_idx].height = 19
                for c_idx in range(1, len(df_res_tipos.columns) + 1):
                    cell = ws_tipos.cell(row=r_idx, column=c_idx)
                    cell.font = NORMAL_FONT
                    cell.border = THIN_BORDER
                    t_name = df_res_tipos.columns[c_idx - 1]
                    if "Total" in t_name or "Diferencia" in t_name:
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                        cell.number_format = NUMBER_FMT
                    elif "N°" in t_name:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    elif t_name == "Estado":
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.font = SUCCESS_FONT if "Cuadrado" in str(cell.value) else ALERT_FONT
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")

            # Totales en resumen de tipos
            tot_t_row = tipo_end + 1
            ws_tipos.row_dimensions[tot_t_row].height = 22
            ws_tipos.cell(row=tot_t_row, column=1, value="TOTAL GENERAL").font = BOLD_FONT

            t_debe_idx = list(df_res_tipos.columns).index("Total Debe ($)") + 1
            t_haber_idx = list(df_res_tipos.columns).index("Total Haber ($)") + 1
            t_debe_l = get_column_letter(t_debe_idx)
            t_haber_l = get_column_letter(t_haber_idx)

            ct_debe = ws_tipos.cell(row=tot_t_row, column=t_debe_idx)
            ct_debe.value = f"=SUM({t_debe_l}{tipo_start}:{t_debe_l}{tipo_end})"
            ct_debe.number_format = NUMBER_FMT
            ct_debe.font = BOLD_FONT
            ct_debe.alignment = Alignment(horizontal="right", vertical="center")

            ct_haber = ws_tipos.cell(row=tot_t_row, column=t_haber_idx)
            ct_haber.value = f"=SUM({t_haber_l}{tipo_start}:{t_haber_l}{tipo_end})"
            ct_haber.number_format = NUMBER_FMT
            ct_haber.font = BOLD_FONT
            ct_haber.alignment = Alignment(horizontal="right", vertical="center")

            for c_idx in range(1, len(df_res_tipos.columns) + 1):
                c = ws_tipos.cell(row=tot_t_row, column=c_idx)
                c.fill = TOTAL_FILL
                c.border = TOTAL_BORDER

            for col in ws_tipos.columns:
                col_letter = col[0].column_letter
                max_len = 0
                for cell in col:
                    if cell.row < 5:
                        continue
                    val_str = str(cell.value or "")
                    if len(val_str) > max_len:
                        max_len = len(val_str)
                ws_tipos.column_dimensions[col_letter].width = max(max_len + 4, 16)

    output.seek(0)
    return output.getvalue()
