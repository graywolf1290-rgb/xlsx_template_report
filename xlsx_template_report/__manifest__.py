# Part of Excel Template Reports. License LGPL-3, see LICENSE file. Copyright (c) 2026 Gray Bob.
{
    "name": "Excel Template Reports",
    "summary": "Design reports in your own Excel file: Odoo fills in the data (free). "
               "Scheduling and Email/Telegram delivery: Excel Template Reports Pro",
    "version": "19.0.1.1.0",
    "category": "Productivity",
    "author": "Gray Bob",
    "support": "gray.wolf1290@gmail.com",
    # Free entry module: LGPL-3 so anyone can install it; the paid features live in the Pro addon (OPL-1).
    "license": "LGPL-3",
    "depends": ["base", "mail"],
    "external_dependencies": {"python": ["openpyxl"]},
    "data": [
        "security/security.xml",
        "security/ir.model.access.csv", "security/ir_rules.xml",
        "views/xlsx_report_views.xml",
        "data/sample_report.xml",
    ],
    "images": ["static/description/banner.png"],
    "application": True,
    "installable": True,
}
