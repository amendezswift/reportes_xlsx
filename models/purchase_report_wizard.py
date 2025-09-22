import base64
import io

import xlsxwriter

from odoo import fields, models


class PurchaseReportWizard(models.TransientModel):
    _name = "purchase.report.wizard"
    _description = "Purchase Report Wizard"

    fecha_inicial = fields.Date(string="Fecha inicial", required=True)
    fecha_final = fields.Date(string="Fecha final", required=True)

    # -------------------------------------------------------------------------
    # Helpers
    # -------------------------------------------------------------------------
    def _convert_to_company_currency(self, factura, amount):
        if not amount:
            return 0.0
        company = factura.company_id
        currency = factura.currency_id
        if currency == company.currency_id:
            return amount
        conversion_date = factura.invoice_date or factura.date or fields.Date.context_today(self)
        return currency._convert(amount, company.currency_id, company, conversion_date)

    def _get_tipo_documento(self, factura):
        if factura.tipo_factura:
            selection = dict(self.env["account.move"]._fields["tipo_factura"].selection)
            return selection.get(factura.tipo_factura)
        if factura.factura_especial:
            return "Factura especial"
        if factura.move_type == "in_invoice":
            return "Factura"
        if factura.move_type == "in_refund":
            return "Nota de crédito"
        return ""

    def _close_workbook(self, workbook, output, filename):
        workbook.close()
        output.seek(0)
        xlsx_data = output.read()
        output.close()
        attachment = self.env["ir.attachment"].create(
            {
                "name": filename,
                "datas": base64.b64encode(xlsx_data).decode(),
                "type": "binary",
                "res_model": "purchase.report.wizard",
                "res_id": self.id,
                "mimetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            },
        )
        return {
            "type": "ir.actions.act_url",
            "url": f"/web/content/{attachment.id}?download=true",
            "target": "new",
        }

    # -------------------------------------------------------------------------
    # Business
    # -------------------------------------------------------------------------
    def generar_reporte(self):
        self.ensure_one()

        company = self.env.company
        invoices = self.env["account.move"].search(
            [
                ("invoice_date", ">=", self.fecha_inicial),
                ("invoice_date", "<=", self.fecha_final),
                ("state", "=", "posted"),
                ("company_id", "=", company.id),
                "|",
                ("move_type", "in", ["in_invoice", "in_refund"]),
                ("factura_especial", "=", True),
            ],
        )

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {"in_memory": True})
        sheet = workbook.add_worksheet("Libro de compras")

        txt = workbook.add_format({"align": "center"})
        titles_format = workbook.add_format(
            {
                "align": "center",
                "valign": "center",
                "top": 1,
                "right": 1,
                "bottom": 1,
                "left": 1,
            },
        )
        header_format = workbook.add_format(
            {
                "bold": True,
                "align": "center",
                "valign": "center",
                "bg_color": "#C2C2C2",
                "top": 1,
                "right": 1,
                "bottom": 1,
                "left": 1,
            },
        )
        totals_label_format = workbook.add_format(
            {
                "bold": True,
                "align": "right",
                "bg_color": "#8FAF8F",
                "top": 1,
                "right": 1,
                "bottom": 1,
                "left": 1,
            },
        )
        totals_amount_format = workbook.add_format(
            {
                "bold": True,
                "align": "right",
                "bg_color": "#8FAF8F",
                "top": 1,
                "right": 1,
                "bottom": 1,
                "left": 1,
                "num_format": "Q#,##0.00",
            },
        )
        currency_format = workbook.add_format({"num_format": "Q#,##0.00", "align": "end"})

        # Encabezado
        sheet.set_column("A:XFD", 18)
        sheet.write("B2", "Reporte", header_format)
        sheet.write("B3", "Empresa", header_format)
        sheet.write("B4", "NIT", header_format)
        sheet.write("B5", "Periodo", header_format)
        sheet.write("B6", "Moneda", header_format)
        sheet.merge_range("C2:D2", "Libro de compras", titles_format)
        sheet.merge_range("C3:D3", company.display_name, titles_format)
        sheet.merge_range("C4:D4", company.vat or "", titles_format)
        period = f"{self.fecha_inicial.strftime('%d/%m/%Y')} - {self.fecha_final.strftime('%d/%m/%Y')}"
        sheet.merge_range("C5:D5", period, titles_format)
        sheet.merge_range("C6:D6", company.currency_id.display_name, titles_format)

        sheet.write("B9", "Fecha de factura", header_format)
        sheet.write("C9", "Serie", header_format)
        sheet.write("D9", "Número DTE", header_format)
        sheet.write("E9", "Tipo de documento", header_format)
        sheet.write("F9", "Tipo de identificación", header_format)
        sheet.write("G9", "Número de identificación", header_format)
        sheet.write("H9", "Proveedor", header_format)
        sheet.merge_range("I8:J8", "Gravables", header_format)
        sheet.write("I9", "Bienes", header_format)
        sheet.write("J9", "Servicios", header_format)
        sheet.merge_range("K8:L8", "Exentos", header_format)
        sheet.write("K9", "Bienes", header_format)
        sheet.write("L9", "Servicios", header_format)
        sheet.write("M9", "Otros", header_format)
        sheet.write("N9", "IMP", header_format)

        row = 9
        current_column = 14
        taxes = []
        tax_positions = {}
        tax_totals = {}
        combustible = 0

        for factura in invoices:
            if factura.tipo_factura == "recibo":
                continue

            for line in factura.invoice_line_ids:
                if any(tax.tax_group_id.name == "IDP" for tax in line.tax_ids):
                    combustible += self._convert_to_company_currency(
                        factura, line.price_subtotal,
                    )

            if factura.tipo_factura == "poliza":
                for line in factura.invoice_line_ids:
                    tags = {tag.name for tag in line.product_id.product_tag_ids}
                    if not tags.intersection({"IVA", "DAI"}):
                        continue
                    tax_name = line.product_id.name
                    amount = self._convert_to_company_currency(factura, line.price_subtotal)
                    if tax_name not in taxes:
                        sheet.write(8, current_column, tax_name, header_format)
                        taxes.append(tax_name)
                        tax_positions[tax_name] = current_column
                        current_column += 1
                    tax_totals[tax_name] = tax_totals.get(tax_name, 0) + amount

            if factura.tax_totals:
                for group in factura.tax_totals.get("groups_by_subtotal", {}).values():
                    for tax_group in group:
                        name = tax_group["tax_group_name"]
                        if name not in taxes:
                            sheet.write(8, current_column, name, header_format)
                            taxes.append(name)
                            tax_positions[name] = current_column
                            current_column += 1
                        amount = self._convert_to_company_currency(
                            factura, tax_group["tax_group_amount"],
                        )
                        if factura.move_type == "in_refund":
                            amount *= -1
                        tax_totals[name] = tax_totals.get(name, 0) + amount

        total_column = current_column
        sheet.write(8, total_column, "Total", header_format)

        total_bienes_gravables = 0
        total_servicios_gravables = 0
        total_bienes_exentos = 0
        total_servicios_exentos = 0
        total_notas_credito = 0
        total_otros = 0
        total_importaciones = 0
        grand_total = 0

        for factura in invoices:
            if factura.tipo_factura == "recibo":
                continue

            serie = factura.serie_proveedor or factura.serie
            numero_dte = factura.dte_proveedor or factura.numero_dte

            date_str = factura.invoice_date.strftime("%d/%m/%Y") if factura.invoice_date else ""
            sheet.write(row, 1, date_str, txt)
            sheet.write(row, 2, serie or "", txt)
            sheet.write(row, 3, numero_dte or "", txt)
            sheet.write(row, 4, self._get_tipo_documento(factura), txt)

            nit = factura.partner_id.vat
            if nit:
                sheet.write(row, 5, "NIT", txt)
                sheet.write(row, 6, nit, txt)
            else:
                sheet.write(row, 5, "DPI/Pasaporte", txt)
                sheet.write(row, 6, getattr(factura.partner_id, "cui", ""), txt)

            sheet.write(row, 7, factura.partner_id.name or "", txt)

            bienes_gravables = 0
            servicios_gravables = 0
            bienes_exentos = 0
            servicios_exentos = 0
            otros = 0
            monto_importacion = 0

            taxes_for_invoice = set()

            if factura.tipo_factura == "poliza":
                products_amount = {}
                for line in factura.invoice_line_ids:
                    for tag in line.product_id.product_tag_ids:
                        if tag.name in {"IVA", "DAI"}:
                            name = line.product_id.name
                            amount = self._convert_to_company_currency(
                                factura, line.price_subtotal,
                            )
                            products_amount[name] = products_amount.get(name, 0) + amount
                for name, amount in products_amount.items():
                    column_index = tax_positions[name]
                    taxes_for_invoice.add(name)
                    if name == "IVA Importaciones":
                        monto_importacion = round(amount / 0.12, 2)
                        total_importaciones += monto_importacion
                        sheet.write(row, 13, monto_importacion, currency_format)
                    sheet.write(row, column_index, amount, currency_format)

            if factura.tax_totals:
                for group in factura.tax_totals.get("groups_by_subtotal", {}).values():
                    for tax_group in group:
                        name = tax_group["tax_group_name"]
                        column_index = tax_positions[name]
                        taxes_for_invoice.add(name)
                        amount = self._convert_to_company_currency(
                            factura, tax_group["tax_group_amount"],
                        )
                        if factura.move_type == "in_refund":
                            amount *= -1
                        sheet.write(row, column_index, amount, currency_format)

            for name in taxes:
                if name not in taxes_for_invoice:
                    column_index = tax_positions[name]
                    sheet.write(row, column_index, 0, currency_format)

            partner_is_foreign = factura.partner_id.country_id.code != "GT"
            if partner_is_foreign and factura.tipo_factura != "poliza":
                monto_importacion = abs(factura.amount_total_signed)
                if factura.move_type == "in_refund":
                    monto_importacion *= -1
                sheet.write(row, 13, monto_importacion, currency_format)
                total_importaciones += monto_importacion
            elif factura.tipo_factura != "poliza":
                sheet.write(row, 13, 0, currency_format)
                for line in factura.invoice_line_ids:
                    amount = self._convert_to_company_currency(factura, line.price_subtotal)
                    is_product = line.product_id.detailed_type in {"consu", "product"}
                    has_iva = any(tax.name == "IVA 12%" for tax in line.tax_ids)
                    if not has_iva:
                        if is_product:
                            bienes_exentos += amount
                        elif line.product_id.detailed_type == "service":
                            servicios_exentos += amount
                        else:
                            otros += amount
                    else:
                        if is_product:
                            bienes_gravables += amount
                        elif line.product_id.detailed_type == "service":
                            servicios_gravables += amount
                        else:
                            otros += amount

            total_row = abs(factura.amount_total_signed)
            if factura.tipo_factura == "poliza":
                total_row += monto_importacion
            if factura.move_type == "in_refund":
                bienes_gravables *= -1
                servicios_gravables *= -1
                bienes_exentos *= -1
                servicios_exentos *= -1
                otros *= -1
                total_row *= -1
                total_notas_credito += abs(factura.amount_total_signed)

            sheet.write(row, 8, bienes_gravables, currency_format)
            sheet.write(row, 9, servicios_gravables, currency_format)
            sheet.write(row, 10, bienes_exentos, currency_format)
            sheet.write(row, 11, servicios_exentos, currency_format)
            sheet.write(row, 12, otros, currency_format)
            sheet.write(row, total_column, total_row, currency_format)

            total_bienes_gravables += bienes_gravables
            total_servicios_gravables += servicios_gravables
            total_bienes_exentos += bienes_exentos
            total_servicios_exentos += servicios_exentos
            total_otros += otros
            grand_total += total_row

            row += 1

        sheet.write(row, 8, total_bienes_gravables, totals_amount_format)
        sheet.write(row, 9, total_servicios_gravables, totals_amount_format)
        sheet.write(row, 10, total_bienes_exentos, totals_amount_format)
        sheet.write(row, 11, total_servicios_exentos, totals_amount_format)
        sheet.write(row, 12, total_otros, totals_amount_format)
        sheet.write(row, 13, total_importaciones, totals_amount_format)

        for name, amount in tax_totals.items():
            column_index = tax_positions[name]
            sheet.write(row, column_index, amount, totals_amount_format)

        sheet.write(row, total_column, grand_total, totals_amount_format)

        row += 1
        sheet.merge_range(f"B{row}:H{row}", "TOTALES", totals_label_format)

        row_totales = row + 3
        categories = [
            ("Bienes gravables", total_bienes_gravables - combustible),
            ("Servicios gravables", total_servicios_gravables),
            ("Bienes exentos", total_bienes_exentos),
            ("Servicios exentos", total_servicios_exentos),
            ("Combustible", combustible),
            ("Otros", total_otros),
            ("IMP", total_importaciones),
            ("Notas de crédito", total_notas_credito),
        ]
        for label, amount in categories:
            sheet.write(row_totales, 1, label, header_format)
            sheet.write(row_totales, 2, amount, totals_amount_format)
            row_totales += 1

        for name, amount in tax_totals.items():
            sheet.write(row_totales, 1, name, header_format)
            sheet.write(row_totales, 2, amount, totals_amount_format)
            row_totales += 1

        sheet.write(row_totales, 1, "Total", header_format)
        sheet.write(row_totales, 2, grand_total, totals_amount_format)

        filename = (
            f"Libro_de_compras {self.fecha_inicial.strftime('%d/%m/%Y')}"
            f" - {self.fecha_final.strftime('%d/%m/%Y')}.xlsx"
        )
        return self._close_workbook(workbook, output, filename)