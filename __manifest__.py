{
    "name": "Excel Reporting - Swift Solutions",
    "summary": "Custom purchase, sales, and stock Excel reports.",
    "version": "19.0.1.0.0",
    "author": "Swift Solutions",
    "website": "",
    "category": "Reporting",
    "license": "LGPL-3",
    "depends": ["base", "web", "account", "stock"],
    "data": [
        "security/ir.model.access.csv",
        "views/wizard_libro_ventas.xml",
        "views/action_libro_ventas.xml",
        "views/wizard_libro_compras.xml",
        "views/action_libro_compras.xml",
        "views/wizard_reporte_kardex.xml",
        "views/action_reporte_kardex.xml",
        "views/menu_reportes.xml",
    ],
    "assets": {
        "web.assets_backend": [
            "reportes_xlsx/static/src/js/action_manager.js",
        ],
    },
    "installable": True,
    "application": False,
}