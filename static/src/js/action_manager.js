/** @odoo-module **/
import { registry } from "@web/core/registry";
import { BlockUI } from "@web/core/ui/block_ui";
import { download } from "@web/core/network/download";

registry.category("ir.actions.report handlers").add("qwerty_xlsx", async function (action) {
    if (action.report_type === 'xlsx') {
        BlockUI;
	
   await download({
	
           url: '/xlsx_reports',
	
           data: action.data,
	
           complete: () => unblockUI,
	
           error: (error) => self.call('crash_manager', 'rpc_error', error),
	
           });    }
});

/** @odoo-module **/

// import {download} from "@web/core/network/download";
// import {registry} from "@web/core/registry";

// registry
//     .category("ir.actions.report handlers")
//     .add("xlsx_handler", async function (action, options, env) {
//         if (action.report_type === "xlsx") {
//             const type = action.report_type;
//             let url = `/report/${type}/${action.report_name}`;
//             const actionContext = action.context || {};
//             if (action.data && JSON.stringify(action.data) !== "{}") {
//                 const action_options = encodeURIComponent(JSON.stringify(action.data));
//                 const context = encodeURIComponent(JSON.stringify(actionContext));
//                 url += `?options=${action_options}&context=${context}`;
//             } else {
//                 if (actionContext.active_ids) {
//                     url += `/${actionContext.active_ids.join(",")}`;
//                 }
//                 if (type === "xlsx") {
//                     const context = encodeURIComponent(
//                         JSON.stringify(env.services.user.context)
//                     );
//                     url += `?context=${context}`;
//                 }
//             }
//             env.services.ui.block();
//             try {
//                 await download({
//                     url: "/report/download",
//                     data: {
//                         data: JSON.stringify([url, action.report_type]),
//                         context: JSON.stringify(env.services.user.context),
//                     },
//                 });
//             } finally {
//                 env.services.ui.unblock();
//             }
//             const onClose = options.onClose;
//             if (action.close_on_report_download) {
//                 return env.services.action.doAction(
//                     {type: "ir.actions.act_window_close"},
//                     {onClose}
//                 );
//             } else if (onClose) {
//                 onClose();
//             }
//             return Promise.resolve(true);
//         }
//         return Promise.resolve(false);
//     });
