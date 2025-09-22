from odoo import models, fields
from odoo.exceptions import UserError

import io
import xlsxwriter
import base64


class SaleReportWizard(models.TransientModel):
    _name = "purchase.report.wizard"

    fecha_inicial = fields.Date(string="Fecha inicial", required=True)
    fecha_final = fields.Date(string="Fecha final", required=True)

    def _get_tipo_documeto(self, factura):
        if factura.tipo_factura:
            lista_selecciones = self.env["account.move"]._fields["tipo_factura"].selection
            diccionario_selecciones = dict(lista_selecciones)
            return diccionario_selecciones.get(factura.tipo_factura)

        if factura.factura_especial:
            return "Factura_especial"

        if factura.move_type == "in_invoice" and not factura.factura_especial:
            return "Factura"

        if factura.move_type == "in_refund":
            return "Nota de crédito"

    def contar_documentos(self):
        facturas = self.env["account.move"].search(
            [
                ("invoice_date", ">=", self.fecha_inicial),
                ("invoice_date", "<=", self.fecha_final),
                ("move_type", "in", ["in_invoice", "in_refund"]),
                ("state", "in", ["posted", "cancel"]),
            ]
        )

        informacion_facturas = {}

        for factura in facturas:
            tipo = factura.tipo_factura

        return facturas

    def _cerrar_libro(self, workbook, output):
        workbook.close()

        output.seek(0)
        xlsx_data = output.read()
        output.close()

        attachment = self.env["ir.attachment"].create(
            {
                "name": f"Libro_de_compras {self.fecha_inicial.strftime('%d/%m/%Y')} - {self.fecha_final.strftime('%d/%m/%Y')}.xlsx",
                "datas": base64.b64encode(xlsx_data),
                "type": "binary",
                "res_model": "purchase.report.wizard",
                "res_id": self.id,
                "mimetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            }
        )

        return {
            "type": "ir.actions.act_url",
            "url": "/web/content/%s?download=true" % attachment.id,
            "target": "new",
        }

    def generar_reporte(self):
        empresa = self.env.company.name
        nit_empresa = self.env.company.vat
        divisa_empresa = self.env.company.currency_id.name

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {"in_memory": True})
        sheet = workbook.add_worksheet("Libro de compras")

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
        currency_format = workbook.add_format({"num_format": "Q#,##0.00", "align": "end"})

        # ENCABEZADOS

        sheet.set_column("A:XFD", 18)
        sheet.write("B2", "Reporte", bold)
        sheet.write("B3", "Empresa", bold)
        sheet.write("B4", "NIT", bold)
        sheet.write("B5", "Periodo", bold)
        sheet.write("B6", "Moneda", bold)
        sheet.merge_range("C2:D2", "Libro de compras", titulos)
        sheet.merge_range("C3:D3", f"{empresa}", titulos)
        sheet.merge_range("C4:D4", f"{nit_empresa}", titulos)
        sheet.merge_range(
            "C5:D5",
            f"{self.fecha_inicial.strftime('%d/%m/%Y')} - {self.fecha_final.strftime('%d/%m/%Y')}",
            titulos,
        )
        sheet.merge_range("C6:D6", f"{divisa_empresa}", titulos)
        sheet.write("B9", "Fecha de factura", bold)
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
        sheet.write("M9", "Otros", bold)
        sheet.write("N9", "IMP", bold)

        documentos = self.env["account.move"].search(
            [
                ("invoice_date", ">=", self.fecha_inicial),
                ("invoice_date", "<=", self.fecha_final),
                ("state", "=", "posted"),
                "|",
                ("move_type", "in", ["in_invoice", "in_refund"]),
                ("factura_especial", "=", True),
            ]
        )

        # DECLARACIÓN DE POSICIONES INICIALES
        row = 9
        column_actual = 14
        impuestos = []
        impuestos_posiciones = {}
        impuestos_totales = {}
        combus = 0

        for factura in documentos:
            if factura.tipo_factura == "recibo":
                continue

            for linea in factura.invoice_line_ids:
                if any(tax.tax_group_id.name == "IDP" for tax in linea.tax_ids):
                    combus += linea.price_subtotal

            if factura.tipo_factura == "poliza":
                for linea in factura.invoice_line_ids:
                    nombre_impuesto = linea.product_id.name
                    monto_impuesto = factura.currency_id._convert(
                        linea.price_subtotal,
                        factura.company_id.currency_id,
                        factura.company_id,
                        factura.invoice_date,
                    )

                    for etiqueta in linea.product_id.product_tag_ids:
                        if (
                            etiqueta.name in ["IVA", "DAI"]
                            and nombre_impuesto not in impuestos
                        ):
                            sheet.write(8, column_actual, nombre_impuesto, bold)
                            impuestos.append(nombre_impuesto)
                            impuestos_posiciones[nombre_impuesto] = column_actual
                            column_actual += 1

                        impuestos_totales[nombre_impuesto] = (
                            impuestos_totales.get(nombre_impuesto, 0) + monto_impuesto
                        )

            for group_name, tax_groups in factura.tax_totals["groups_by_subtotal"].items():
                for tax_group in tax_groups:
                    if tax_group["tax_group_name"] not in impuestos:
                        sheet.write(
                            8,
                            column_actual,
                            f"{tax_group['tax_group_name']}",
                            bold,
                        )
                        impuestos.append(tax_group["tax_group_name"])
                        impuestos_posiciones[tax_group["tax_group_name"]] = column_actual
                        column_actual += 1

                    nombre_impuesto = tax_group["tax_group_name"]
                    monto_impuesto = factura.currency_id._convert(
                        tax_group["tax_group_amount"],
                        factura.company_id.currency_id,
                        factura.company_id,
                        factura.invoice_date,
                    )

                    if factura.move_type == "in_refund":
                        monto_impuesto *= -1

                    impuestos_totales[nombre_impuesto] = (
                        impuestos_totales.get(nombre_impuesto, 0) + monto_impuesto
                    )

        column_total = column_actual
        sheet.write(8, column_total, "Total", bold)
        total_bienes_gravables = 0
        total_servicios_gravables = 0
        total_bienes_exentos = 0
        total_servicios_exentos = 0
        total_notas_credito = 0
        total_otros = 0
        total_importaciones = 0
        total_totales = 0

        # RECOPILACIÓN DE DATOS DE LA FACTURA

        for factura in documentos:
            if factura.tipo_factura == "recibo":
                continue

            serie = factura.serie_proveedor
            numero_dte = factura.dte_proveedor

            if factura.factura_especial:
                serie = factura.serie
                numero_dte = factura.numero_dte

            sheet.write(row, 1, factura.invoice_date.strftime("%d/%m/%Y"), txt)
            sheet.write(row, 2, serie, txt)
            sheet.write(row, 3, numero_dte, txt)
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
            monto_importacion = 0
            otros = 0

            impuestos_factura = []

            # OBTENCIÓN DE IMPUESTOS PARA CADA FACTURA

            divisa_extranjera = factura.currency_id != factura.company_id.currency_id

            if factura.tipo_factura == "poliza":
                productos = {}
                for linea in factura.invoice_line_ids:
                    for etiqueta in linea.product_id.product_tag_ids:
                        if etiqueta.name in ["IVA", "DAI"]:
                            nombre_producto = linea.product_id.name
                            monto_producto = factura.currency_id._convert(
                                linea.price_subtotal,
                                factura.company_id.currency_id,
                                factura.company_id,
                                factura.invoice_date,
                            )
                            productos[nombre_producto] = (
                                productos.get(nombre_producto, 0) + monto_producto
                            )

                for producto, precio in productos.items():
                    column_index = impuestos_posiciones[producto]
                    if producto not in impuestos_factura:
                        impuestos_factura.append(producto)

                    if producto == "IVA Importaciones":
                        monto_importacion = round((precio / 0.12), 2)
                        total_importaciones += monto_importacion
                        sheet.write(row, 13, monto_importacion, currency_format)

                    sheet.write(row, column_index, precio, currency_format)

            for group_name, tax_groups in factura.tax_totals["groups_by_subtotal"].items():
                for tax_group in tax_groups:
                    column_index = impuestos_posiciones[tax_group["tax_group_name"]]
                    impuestos_factura.append(tax_group["tax_group_name"])

                    monto_gtq = 0.0

                    if divisa_extranjera:
                        monto_gtq = factura.currency_id._convert(
                            tax_group["tax_group_amount"],
                            factura.company_id.currency_id,
                            factura.company_id,
                            factura.invoice_date,
                        )

                    if not divisa_extranjera:
                        monto_gtq = tax_group["tax_group_amount"]

                    if factura.move_type == "in_refund":
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

            # SEPARACIÓN DE FACTURAS DE IMPORTACION
            if factura.partner_id.country_id.code != "GT" and factura.tipo_factura != "poliza":
                monto_importacion = abs(factura.amount_total_signed)

                if factura.move_type == "in_refund":
                    monto_importacion *= -1

                sheet.write(row, 13, monto_importacion, currency_format)
                total_importaciones += monto_importacion
            else:
                if factura.tipo_factura != "poliza":
                    sheet.write(row, 13, 0, currency_format)

                for linea in factura.invoice_line_ids:
                    if factura.currency_id != factura.company_id.currency_id:
                        monto_gtq = factura.currency_id._convert(
                            linea.price_subtotal,
                            factura.company_id.currency_id,
                            factura.company_id,
                            factura.invoice_date,
                        )
                    else:
                        monto_gtq = linea.price_subtotal

                    # SEPARACIÓN DE BIENES Y SERVICIOS GRAVABLES O EXENTOS
                    productos = ["consu", "product"]
                    if factura.tipo_factura != "poliza":
                        if not any(impuesto.name == "IVA 12%" for impuesto in linea.tax_ids):
                            if linea.product_id.detailed_type in productos:
                                bienes_exentos += monto_gtq
                            elif linea.product_id.detailed_type == "service":
                                servicios_exentos += monto_gtq
                            elif not linea.product_id:
                                otros += monto_gtq
                        else:
                            if linea.product_id.detailed_type in productos:
                                bienes_gravables += monto_gtq
                            elif linea.product_id.detailed_type == "service":
                                servicios_gravables += monto_gtq
                            elif not linea.product_id:
                                otros += monto_gtq

            total_fila = abs(factura.amount_total_signed)

            if factura.tipo_factura == "poliza":
                total_fila += monto_importacion

            if factura.move_type == "in_refund":
                bienes_gravables *= -1
                bienes_exentos *= -1
                servicios_gravables *= -1
                servicios_exentos *= -1
                otros *= -1
                total_fila *= -1
                total_notas_credito += abs(factura.amount_total_signed)

            sheet.write(row, 8, bienes_gravables, currency_format)
            sheet.write(row, 9, servicios_gravables, currency_format)
            sheet.write(row, 10, bienes_exentos, currency_format)
            sheet.write(row, 11, servicios_exentos, currency_format)
            sheet.write(row, 12, otros, currency_format)
            sheet.write(row, column_total, total_fila, currency_format)

            # CALCULO DE LOS TOTALES
            total_bienes_gravables += bienes_gravables
            total_servicios_gravables += servicios_gravables
            total_bienes_exentos += bienes_exentos
            total_servicios_exentos += servicios_exentos
            total_otros += otros
            total_totales += total_fila

            row += 1

        # IMPRESIONES FINALES
        sheet.write(row, 8, total_bienes_gravables, totales_monto)
        sheet.write(row, 9, total_servicios_gravables, totales_monto)
        sheet.write(row, 10, total_bienes_exentos, totales_monto)
        sheet.write(row, 11, total_servicios_exentos, totales_monto)
        sheet.write(row, 12, total_otros, totales_monto)
        sheet.write(row, 13, total_importaciones, totales_monto)

        for impuesto_name, total_impuesto in impuestos_totales.items():
            column_index = impuestos_posiciones[impuesto_name]
            sheet.write(row, column_index, total_impuesto, totales_monto)

        sheet.write(row, column_total, total_totales, totales_monto)

        row += 1
        sheet.merge_range(f"B{row}:H{row}", "TOTALES", totales)

        # TABLA DE RESUMEN
        # ENCABEZADOS
        row_totales = row + 3
        # row_documentos_totales = row + 3
        # bienes_gravables_combusless = total_bienes_gravables - combus
        # sheet.write(row_totales, 1, "Categorías", bold)
        # sheet.write(row_totales, 2, "Base gravable", bold)
        # sheet.write(row_totales, 3, "Base excenta", bold)
        # sheet.write(row_totales, 4, "IVA", bold)
        # row_totales += 1

        # CATEGORIAS = [
        #     (
        #         "Bienes",
        #         bienes_gravables_combusless,
        #         total_bienes_exentos,
        #         (bienes_gravables_combusless * 0.12),
        #     ),
        #     (
        #         "Servicios",
        #         total_servicios_gravables,
        #         total_servicios_exentos,
        #         (total_servicios_gravables * 0.12),
        #     ),
        #     ("Combustible", combus, 0, (combus * 0.12)),
        #     ("Otros", total_otros, 0, (total_otros * 0.12)),
        #     ("IMP", total_importaciones, 0, 0),
        #     ("Notas de crédito", total_notas_credito, 0, (total_notas_credito * 0.12)),
        # ]

        # # TABLA CON IVA
        # for categoria, total_gravable, total_excento, iva in CATEGORIAS:
        #     sheet.write(row_totales, 1, categoria, bold)
        #     sheet.write(row_totales, 2, total_gravable, currency_format)
        #     sheet.write(row_totales, 3, total_excento, currency_format)
        #     sheet.write(row_totales, 4, iva, currency_format)
        #     row_totales += 1

        # res = self.contar_documentos()
        # raise UserError(len(res))

        facturas_encontradas = {}

        # facturas_conteo = self.env["account.move"].search(
        #     [
        #         ("invoice_date", ">=", self.fecha_inicial),
        #         ("invoice_date", "<=", self.fecha_final),
        #         ("move_type", "in", ["in_invoice", "in_refund"]),
        #         ("state", "in", ["posted", "cancel"]),
        #     ]
        # )

        # for factura_encontrada in facturas_conteo:
        #     tipo = factura_encontrada.tipo_factura or "Sin tipo de documento"
        #     estado = factura.state

        #     if tipo not in facturas_encontradas:
        #         facturas_encontradas[tipo] = {
        #             "Tipo de documento": tipo,
        #             "Vigentes": 0,
        #             "Canceladas": 0,
        #         }

        #     if estado == "posted":
        #         facturas_encontradas[tipo]["Vigentes"] += 1
        #     elif estado == "cancel":
        #         facturas_encontradas[tipo]["Canceladas"] += 1

        # resultados = [
        #     (tipo, datos["Vigentes"], datos["Canceladas"])
        #     for tipo, datos in facturas_encontradas.items()
        # ]

        # sheet.write(row_documentos_totales, 1, "Tipo de documento", bold)
        # sheet.write(row_documentos_totales, 2, "Vigentes", bold)
        # sheet.write(row_documentos_totales, 3, "Cancelados", bold)
        # row_documentos_totales += 1

        # for tipo_, vigente_, cancelada_ in resultados:
        #     sheet.write(row_documentos_totales, 1, tipo_, txt)
        #     sheet.write(row_documentos_totales, 2, vigente_, txt)
        #     sheet.write(row_documentos_totales, 3, cancelada_, txt)
        #     row_documentos_totales += 1

        categorias_totales = [
            ("Bienes gravables", total_bienes_gravables - combus),
            ("Servicios gravables", total_servicios_gravables),
            ("Bienes exentos", total_bienes_exentos),
            ("Servicios exentos", total_servicios_exentos),
            ("Combustible", combus),
            ("Otros", total_otros),
            ("IMP", total_importaciones),
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

        return self._cerrar_libro(workbook, output)
