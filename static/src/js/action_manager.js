/** @odoo-module **/

import { _t } from "@web/core/l10n/translation";
import { download } from "@web/core/network/download";
import { registry } from "@web/core/registry";

registry.category("ir.actions.report handlers").add(
    "xlsx_report_handler",
    async (action, options = {}, env) => {
        if (action.report_type !== "xlsx") {
            return false;
        }
        const { action: actionService, notification, ui } = env.services;
        const data = { ...(action.data || {}) };
        if (!data.output_format) {
            data.output_format = "xlsx";
        }
        if (action.context && !data.context) {
            data.context = JSON.stringify(action.context);
        }
        ui.block();
        try {
            await download({
                url: "/xlsx_reports",
                data,
            });
        } catch (error) {
            notification.add(error.message || _t("An error occurred while generating the XLSX report."), {
                type: "danger",
            });
            throw error;
        } finally {
            ui.unblock();
        }
        const onClose = options.onClose;
        if (action.close_on_report_download) {
            return actionService.doAction(
                { type: "ir.actions.act_window_close" },
                { onClose }
            );
        }
        if (onClose) {
            onClose();
        }
        return true;
    }
);