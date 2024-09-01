# -*- coding: utf-8 -*-

from odoo import fields, models

class KardexReportWizard(models.TransientModel):
    _name='kardex.report.wizard'
    _description = "Kardex Report"

    start_date = fields.Date(string="Start Date")
    end_date = fields.Date(string="End Date")
    product_id = fields.Many2one('product.product', string="Product")
    warehouse = fields.Many2one('stock.warehouse', string="Warehouse")

    def print_kardex_report_xls(self):
        return self.env.ref('denteco_02_bodega_kardex_report.bodega_kardex_report').report_action(self, data={})
