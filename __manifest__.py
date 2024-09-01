# -*- coding: utf-8 -*-

{
    'name': 'Kardex Report XLSX',
    'version': '1.0.1',
    "category": "Reporting",
    'author': 'Xetechs GT',
    'depends': [
        'stock', 
        'stock_account',
        'report_xlsx'
        ],
    'license': 'AGPL-3',
    'data': [
        'security/ir.model.access.csv',
        'wizard/kardex_report_wizard_view.xml',
        'report/ir_action_report.xml',
        'view/menu_view.xml',
        'view/stock_picking_views.xml'
    ]
}
