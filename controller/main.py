import json
from contextlib import suppress

from odoo import http
from odoo.http import content_disposition, request
from odoo.http import serialize_exception as _serialize_exception
from odoo.tools import html_escape


class XLSXReportController(http.Controller):
    @http.route("/xlsx_reports", type="http", auth="user", methods=["POST"], csrf=False)
    def get_report_xlsx(self, model, options, output_format, **kw):
        """Return a dynamically generated XLSX report."""
        uid = request.session.uid
        context = dict(request.env.context)
        if "context" in kw:
            with suppress(json.JSONDecodeError):
                context.update(json.loads(kw["context"]))
        report_obj = request.env[model].with_user(uid).with_context(context)
        report_options = {}
        if options:
            with suppress(json.JSONDecodeError):
                report_options = json.loads(options)
        token = kw.get("token", "dummy-because-api-expects-one")
        try:
            if output_format == "xlsx":
                filename = report_options.get("output_name") or report_options.get("report_name")
                if not filename:
                    filename = report_obj._description or "report"
                if not filename.lower().endswith(".xlsx"):
                    filename = f"{filename}.xlsx"
                response = request.make_response(
                    None,
                    headers=[
                        (
                            "Content-Type",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        ),
                        ("Content-Disposition", content_disposition(filename)),
                    ],
                )
                report_obj.get_xlsx_report(report_options, response)
                response.set_cookie("fileToken", token)
                return response
        except Exception as exc:  # noqa: BLE001
            se = _serialize_exception(exc)
            error = {"code": 200, "message": "Odoo Server Error", "data": se}
            return request.make_response(html_escape(json.dumps(error)))
        return request.not_found()