from odoo import models, fields

class StockPicking(models.Model):
    _inherit = 'stock.picking'
    _description = 'Stock Picking'
    
    def button_validate(self):
        res = super(StockPicking, self).button_validate()
        self.user_id = self.env.user.id 
        return res     
               
    def action_cancel(self):
        res = super(StockPicking, self).action_cancel()
        self.user_id = self.env.user.id
        return res    

class StockPickingType(models.Model):
    _inherit = 'stock.picking.type'
    _description = 'Stock Picking Type'
    
    alias = fields.Char('Alias')