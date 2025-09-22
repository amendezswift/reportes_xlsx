from odoo import models, fields

import io
import xlsxwriter
import base64


class SaleReportWizard(models.TransientModel):
    _name = "sale.report.wizard"

    fecha_inicial = fields.Date(string="Fecha inicial", required=True)
    fecha_final = fields.Date(string="Fecha final", required=True)

    def _get_tipo_documeto(self, factura):
        if not factura.numero_dte:
            return "INVALIDO"

        if factura.debit_origin_id:
            return "NDEB"

        if factura.move_type == "out_refund":
            return "NCRE"

        if factura.tipo_factura == "fact":
            return "FACT"

        if factura.tipo_factura == "fact_cambiaria":
            return "FCAM"

    def generar_reporte(self):
        documentos_cliente = self.env["account.move"].search(
            [
                ("invoice_date", ">=", self.fecha_inicial),
                ("invoice_date", "<=", self.fecha_final),
                ("state", "=", "posted"),
                ("move_type", "in", ["out_invoice", "out_refund"]),
            ]
        )

        empresa = self.env.company.name
        nit_empresa = self.env.company.vat
        divisa_empresa = self.env.company.currency_id.name

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {"in_memory": True})
        sheet = workbook.add_worksheet("Libro de ventas")

        # FORMATOS

        txt = workbook.add_format({"align": "center"})
        titulos = workbook.add_format(
            {
                "align": "center",
                "valign": "center",
                "top": 1,
                "right": 1,
                "bottom": 1,
                "left": 1,
            }
        )
        bold = workbook.add_format(
            {
                "bold": True,
                "align": "center",
                "valign": "center",
                "bg_color": "#C2C2C2",
                "top": 1,
                "right": 1,
                "bottom": 1,
                "left": 1,
            }
        )
        totales = workbook.add_format(
            {
                "bold": True,
                "align": "right",
                "bg_color": "#8FAF8F",
                "top": 1,
                "right": 1,
                "bottom": 1,
                "left": 1,
            }
        )
        totales_monto = workbook.add_format(
            {
                "bold": True,
                "align": "right",
                "bg_color": "#8FAF8F",
                "top": 1,
                "right": 1,
                "bottom": 1,
                "left": 1,
                "num_format": "Q#,##0.00",
            }
        )
        currency_format = workbook.add_format(
            {"num_format": "Q#,##0.00", "align": "end"}
        )

        # ENCABEZADOS

        sheet.set_column("A:XFD", 18)
        sheet.write("B2", "Reporte", bold)
        sheet.write("B3", "Empresa", bold)
        sheet.write("B4", "NIT", bold)
        sheet.write("B5", "Periodo", bold)
        sheet.write("B6", "Moneda", bold)
        sheet.merge_range("C2:D2", "Libro de ventas", titulos)
        sheet.merge_range("C3:D3", f"{empresa}", titulos)
        sheet.merge_range("C4:D4", f"{nit_empresa}", titulos)
        sheet.merge_range(
            "C5:D5",
            f"{self.fecha_inicial.strftime('%d/%m/%Y')} - {self.fecha_final.strftime('%d/%m/%Y')}",
            titulos,
        )
        sheet.merge_range("C6:D6", f"{divisa_empresa}", titulos)
        sheet.write("B9", "Fecha de documento", bold)
        sheet.write("C9", "Serie", bold)
        sheet.write("D9", "Número DTE", bold)
        sheet.write("E9", "Tipo de documento", bold)
        sheet.write("F9", "Tipo de identificación", bold)
        sheet.write("G9", "Número de identificación", bold)
        sheet.write("H9", "Proveedor", bold)
        sheet.merge_range("I8:J8", "Gravables", bold)
        sheet.write("I9", "Bienes", bold)
        sheet.write("J9", "Servicios", bold)
        sheet.merge_range("K8:L8", "Exentos", bold)
        sheet.write("K9", "Bienes", bold)
        sheet.write("L9", "Servicios", bold)
        sheet.write("M9", "Exportaciones", bold)

        # POSICIONES INICIALES

        row = 9
        column_actual = 13
        impuestos = []
        impuestos_posiciones = {}
        impuestos_totales = {}

        # CREACIÓN DE DICCIONARIO CON EL TOTAL DE LOS IMPUESTOS, UNO CON LAS POSICIONES Y UNA LISTA CON LOS IMPUESTOS EXISTENTES

        for factura in documentos_cliente:
            for group_name, tax_groups in factura.tax_totals[
                "groups_by_subtotal"
            ].items():
                for tax_group in tax_groups:
                    if tax_group["tax_group_name"] not in impuestos:
                        sheet.write(
                            8,
                            column_actual,
                            f"{tax_group['tax_group_name']}",
                            bold,
                        )
                        impuestos.append(tax_group["tax_group_name"])
                        impuestos_posiciones[tax_group["tax_group_name"]] = (
                            column_actual
                        )
                        column_actual += 1

                    nombre_impuesto = tax_group["tax_group_name"]
                    monto_impuesto = factura.currency_id._convert(
                        tax_group["tax_group_amount"],
                        factura.company_id.currency_id,
                        factura.company_id,
                        factura.date,
                    )

                    if factura.move_type == "out_refund":
                        monto_impuesto *= -1

                    impuestos_totales[nombre_impuesto] = (
                        impuestos_totales.get(nombre_impuesto, 0) + monto_impuesto
                    )

        # INICIALIZACIÓN DE VARIALES DESTINADAS PARA TOTALES

        column_total = column_actual
        sheet.write(8, column_total, "Total", bold)
        total_bienes_gravables = 0
        total_servicios_gravables = 0
        total_bienes_exentos = 0
        total_servicios_exentos = 0
        total_notas_credito = 0
        total_exportaciones = 0
        total_totales = 0

        # RECOPILACIÓN DE DATOS DE LA FACTURA

        for factura in documentos_cliente:
            sheet.write(row, 1, factura.invoice_date.strftime("%d/%m/%Y"), txt)
            sheet.write(row, 2, factura.serie, txt)
            sheet.write(row, 3, factura.numero_dte, txt)
            tipo_documento = self._get_tipo_documeto(factura)
            sheet.write(row, 4, tipo_documento, txt)
            nit = factura.partner_id.vat

            if not nit:
                sheet.write(row, 5, "DPI/Pasaporte", txt)
                dpi_pasaporte = factura.partner_id.cui
                sheet.write(row, 6, dpi_pasaporte, txt)
            else:
                sheet.write(row, 5, "NIT", txt)
                sheet.write(row, 6, nit, txt)

            sheet.write(row, 7, factura.partner_id.name, txt)

            bienes_gravables = 0
            servicios_gravables = 0
            bienes_exentos = 0
            servicios_exentos = 0

            impuestos_factura = []

            # OBTENCIÓN DE IMPUESTOS PARA CADA FACTURA

            for group_name, tax_groups in factura.tax_totals[
                "groups_by_subtotal"
            ].items():
                for tax_group in tax_groups:
                    column_index = impuestos_posiciones[tax_group["tax_group_name"]]
                    impuestos_factura.append(tax_group["tax_group_name"])

                    monto_gtq = 0.0

                    divisa_extranjera = (
                        factura.currency_id != factura.company_id.currency_id
                    )
                    if divisa_extranjera:
                        monto_gtq = factura.currency_id._convert(
                            tax_group["tax_group_amount"],
                            factura.company_id.currency_id,
                            factura.company_id,
                            factura.date,
                        )

                    if not divisa_extranjera:
                        monto_gtq = tax_group["tax_group_amount"]

                    if factura.move_type == "out_refund":
                        monto_gtq *= -1

                    sheet.write(
                        row,
                        column_index,
                        monto_gtq,
                        currency_format,
                    )

            # IMPRESIÓN DE 0 SI EL IMPUESTO NO EXISTE PARA LA FACTURA

            for impuesto in impuestos:
                if impuesto not in impuestos_factura:
                    column_index = impuestos_posiciones[impuesto]
                    sheet.write(row, column_index, 0, currency_format)

            # SEPARACIÓN DE FACTURAS DE EXPORTACIÓN
            if (
                factura.partner_id.country_id.code != "GT"
                or factura.fiscal_position_id.name == "Exportación"
            ):
                monto_exportacion = abs(factura.amount_total_signed)
                if factura.move_type == "out_refund":
                    monto_exportacion *= -1

                sheet.write(row, 12, monto_exportacion, currency_format)
                total_exportaciones += monto_exportacion
            else:
                sheet.write(row, 12, 0, currency_format)
                for linea in factura.invoice_line_ids:
                    if factura.currency_id != factura.company_id.currency_id:
                        monto_gtq = factura.currency_id._convert(
                            linea.price_subtotal,
                            factura.company_id.currency_id,
                            factura.company_id,
                            factura.date,
                            round=False,
                        )
                    else:
                        monto_gtq = linea.price_subtotal

                    # SEPARACIÓN DE BIENES Y SERVICIOS GRAVABLES O EXENTOS
                    productos = ["consu", "product"]
                    if not any(
                        impuesto.name == "IVA 12%" for impuesto in linea.tax_ids
                    ):
                        if linea.product_id.detailed_type in productos:
                            bienes_exentos += monto_gtq
                        elif linea.product_id.detailed_type == "service":
                            servicios_exentos += monto_gtq
                    else:
                        if linea.product_id.detailed_type in productos:
                            bienes_gravables += monto_gtq
                        elif linea.product_id.detailed_type == "service":
                            servicios_gravables += monto_gtq

            total_fila = abs(factura.amount_total_signed)
            if factura.move_type == "out_refund":
                bienes_gravables *= -1
                bienes_exentos *= -1
                servicios_gravables *= -1
                servicios_exentos *= -1
                total_fila *= -1
                total_notas_credito += abs(factura.amount_total_signed)

            sheet.write(row, 8, bienes_gravables, currency_format)
            sheet.write(row, 9, servicios_gravables, currency_format)
            sheet.write(row, 10, bienes_exentos, currency_format)
            sheet.write(row, 11, servicios_exentos, currency_format)
            sheet.write(row, column_total, total_fila, currency_format)

            # CALCULO DE LOS TOTALES
            total_bienes_gravables += bienes_gravables
            total_servicios_gravables += servicios_gravables
            total_bienes_exentos += bienes_exentos
            total_servicios_exentos += servicios_exentos
            total_totales += total_fila

            row += 1

        # IMPRESIONES FINALES
        sheet.write(row, 8, total_bienes_gravables, totales_monto)
        sheet.write(row, 9, total_servicios_gravables, totales_monto)
        sheet.write(row, 10, total_bienes_exentos, totales_monto)
        sheet.write(row, 11, total_servicios_exentos, totales_monto)
        sheet.write(row, 12, total_exportaciones, totales_monto)

        for impuesto_name, total_impuesto in impuestos_totales.items():
            column_index = impuestos_posiciones[impuesto_name]
            sheet.write(row, column_index, total_impuesto, totales_monto)

        sheet.write(row, column_total, total_totales, totales_monto)

        row += 1
        sheet.merge_range(f"B{row}:H{row}", "TOTALES", totales)

        # TABLA DE RESUMEN
        row_totales = row + 3
        sheet.merge_range(f"B{row_totales}:C{row_totales}", "TOTALES", bold)
        sheet.write(row_totales, 1, "Categoría", bold)
        sheet.write(row_totales, 2, "Monto", bold)

        row_totales += 1

        categorias_totales = [
            ("Bienes gravables", total_bienes_gravables),
            ("Servicios gravables", total_servicios_gravables),
            ("Bienes exentos", total_bienes_exentos),
            ("Servicios exentos", total_servicios_exentos),
            ("Exportaciones", total_exportaciones),
            ("Notas de crédito", total_notas_credito),
        ]

        for categoria, total in categorias_totales:
            sheet.write(row_totales, 1, categoria, bold)
            sheet.write(row_totales, 2, total, totales_monto)
            row_totales += 1

        for impuesto_name, total_impuesto in impuestos_totales.items():
            sheet.write(row_totales, 1, impuesto_name, bold)
            sheet.write(row_totales, 2, total_impuesto, totales_monto)
            row_totales += 1

        sheet.write(row_totales, 1, "Total", bold)
        sheet.write(row_totales, 2, total_totales, totales_monto)

        workbook.close()

        output.seek(0)
        xlsx_data = output.read()
        output.close()

        attachment = self.env["ir.attachment"].create(
            {
                "name": f"Libro_de_ventas {self.fecha_inicial.strftime('%d/%m/%Y')} - {self.fecha_final.strftime('%d/%m/%Y')}.xlsx",
                "datas": base64.b64encode(xlsx_data),
                "type": "binary",
                "res_model": "sale.report.wizard",
                "res_id": self.id,
                "mimetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            }
        )

        return {
            "type": "ir.actions.act_url",
            "url": "/web/content/%s?download=true" % attachment.id,
            "target": "new",
        }
