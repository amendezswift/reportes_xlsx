{
    "name": "Reportería en Excel - Swift Solutions",
    "version": "17.0.0.0.1",
    "depends": ["base", "web", "sale"],
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
        ]
    },
    "installable": True,
    "application": False,
}
