import base64
import io
import json

import xlsxwriter

from odoo import api, fields, models


class KardexReportWizard(models.TransientModel):
    _name = "kardex.report.wizard"
    _description = "Kardex Report Wizard"

    fecha_inicial = fields.Date(string="Fecha inicial", required=True)
    fecha_final = fields.Date(string="Fecha final", required=True)
    empresa = fields.Many2one(
        comodel_name="res.company",
        string="Empresa",
        default=lambda self: self.env.company,
        domain=lambda self: [
            ("id", "in", self.env.user.company_ids.ids),
        ],
        required=True,
    )
    almacen = fields.Many2one(
        comodel_name="stock.warehouse",
        string="Almacén",
        domain="[('company_id', '=', empresa)]",
    )
    ubicacion = fields.Many2one(
        comodel_name="stock.location",
        string="Ubicación",
    )
    ubicaciones_domain = fields.Char(
        string="Dominio para ubicaciones",
        compute="_compute_ubicaciones_domain",
    )

    @api.depends("almacen")
    def _compute_ubicaciones_domain(self):
        for wizard in self:
            if wizard.almacen:
                wizard.ubicaciones_domain = json.dumps(
                    [("warehouse_id", "=", wizard.almacen.id)],
                )
            else:
                wizard.ubicaciones_domain = "[]"

    def generar_reporte(self):
        self.ensure_one()

        start_date = fields.Datetime.to_datetime(self.fecha_inicial)
        if start_date:
            start_date = fields.Datetime.begin_of(start_date, "day")
        end_date = fields.Datetime.to_datetime(self.fecha_final)
        if end_date:
            end_date = fields.Datetime.end_of(end_date, "day")

        domain = [
            ("date", ">=", start_date),
            ("date", "<=", end_date),
            ("state", "=", "done"),
            ("company_id", "=", self.empresa.id),
        ]
        if self.almacen:
            domain.append(("picking_type_id.warehouse_id", "=", self.almacen.id))
        if self.ubicacion:
            domain.extend(
                [
                    "|",
                    ("location_id", "child_of", self.ubicacion.id),
                    ("location_dest_id", "child_of", self.ubicacion.id),
                ],
            )

        moves = self.env["stock.move"].search(domain, order="date, id")

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {"in_memory": True})
        sheet = workbook.add_worksheet("Kardex")
        txt = workbook.add_format({"align": "center"})
        bold = workbook.add_format({"bold": True, "align": "center"})

        sheet.set_column("B:G", 18)

        sheet.write("B2", "Fecha", bold)
        sheet.write("C2", "Producto", bold)
        sheet.write("D2", "Referencia", bold)
        sheet.write("E2", "Cantidad", bold)
        sheet.write("F2", "Ubicación origen", bold)
        sheet.write("G2", "Ubicación destino", bold)

        row = 2
        for move in moves:
            timestamp = fields.Datetime.context_timestamp(self, move.date) if move.date else False
            sheet.write(row, 1, timestamp.strftime("%d/%m/%Y") if timestamp else "", txt)
            sheet.write(row, 2, move.product_id.display_name, txt)
            sheet.write(row, 3, move.reference or "", txt)
            sheet.write(row, 4, move.quantity_done, txt)
            sheet.write(row, 5, move.location_id.complete_name, txt)
            sheet.write(row, 6, move.location_dest_id.complete_name, txt)
            row += 1

        workbook.close()

        output.seek(0)
        xlsx_data = output.read()
        output.close()

        attachment = self.env["ir.attachment"].create(
            {
                "name": f"Kardex_{self.fecha_inicial}_{self.fecha_final}.xlsx",
                "datas": base64.b64encode(xlsx_data).decode(),
                "type": "binary",
                "res_model": "kardex.report.wizard",
                "res_id": self.id,
                "mimetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            },
        )

        return {
            "type": "ir.actions.act_url",
            "url": f"/web/content/{attachment.id}?download=true",
            "target": "new",
        }