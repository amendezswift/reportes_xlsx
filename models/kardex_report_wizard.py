from odoo import models, fields, api
from odoo.exceptions import UserError

import io
import xlsxwriter
import base64


class KardexReportWizard(models.TransientModel):
    _name = "kardex.report.wizard"

    fecha_inicial = fields.Date(string="Fecha inicial", required=True)
    fecha_final = fields.Date(string="Fecha final", required=True)
    empresa = fields.Many2one(
        comodel_name="res.company",
        string="Empresa",
        default=lambda self: self.env.company,
        domain=lambda self: [("id", "in", self.env.user.company_ids.ids)],
        required=True,
    )
    almacen = fields.Many2one(comodel_name="stock.warehouse", string="Almacén")
    ubicacion = fields.Many2one(
        comodel_name="stock.location",
        string="Ubicación",
    )
    ubicaciones_domain = fields.Char(
        string="Dominio para ubicaciones", compute="_compute_ubicaciones_domain"
    )

    @api.depends("almacen")
    def _compute_ubicaciones_domain(self):
        for wizard in self:
            domain = f"[('warehouse_id', '=', {wizard.almacen.id})]"
            self.ubicaciones_domain = domain

    def generar_reporte(self):
        movimientos = self.env["stock.move"].search(
            [
                ("date", ">=", self.fecha_inicial),
                ("date", "<=", self.fecha_final),
                ("state", "=", "done"),
            ]
        )

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
        for move in movimientos:
            fecha = fields.Datetime.context_timestamp(self, move.date)
            sheet.write(row, 1, fecha.strftime("%d/%m/%Y"), txt)
            sheet.write(row, 2, move.product_id.name, txt)
            sheet.write(row, 3, move.reference, txt)
            sheet.write(row, 4, move.product_uom_qty, txt)
            sheet.write(row, 5, move.location_id.name, txt)
            sheet.write(row, 6, move.location_dest_id.name, txt)
            row += 1

        workbook.close()

        output.seek(0)
        xlsx_data = output.read()
        output.close()

        attachment = self.env["ir.attachment"].create(
            {
                "name": f"Kardex_{self.fecha_inicial} - {self.fecha_final}.xlsx",
                "datas": base64.b64encode(xlsx_data),
                "type": "binary",
                "res_model": "kardex.report.wizard",
                "res_id": self.id,
                "mimetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            }
        )

        return {
            "type": "ir.actions.act_url",
            "url": "/web/content/%s?download=true" % attachment.id,
            "target": "new",
        }
