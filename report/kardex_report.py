import logging
from datetime import datetime
import pytz
from lxml import html
from odoo import models, _

class KardexReportXlsx(models.AbstractModel):
    _name = 'report.denteco_02_bodega_kardex_report.kardex_report'
    _inherit = 'report.report_xlsx.abstract'

    def _server_to_user_datetime(self, date, local_time):        
        user_datetime = pytz.utc.localize(date).astimezone(local_time)
        return user_datetime.replace(tzinfo=None)
    
    def _user_to_server_date(self, date, user_tz):
        server_datetime = pytz.timezone(user_tz).localize(date).astimezone(pytz.utc)
        return server_datetime.replace(tzinfo=None) 

    def generate_xlsx_report(self, workbook, data, wizard):
        wizard.ensure_one()
        sheet = workbook.add_worksheet()

        # Tiemezone parameters
        user_tz = self.env.context.get('tz') or self.env.user.tz
        local_time = pytz.timezone(user_tz)

        # Sheets format parameters
        datetime_style = workbook.add_format({'align': 'left', 'num_format': 'dd-mm-yyyy hh:mm:ss', 'border': True})
        head = workbook.add_format({'align': 'center', 'bold': True, 'bg_color': 'black','font_color':'white', 'border': True})
        format_merge = workbook.add_format({'align': 'right', 'bold': True, 'bg_color': 'black','font_color':'white', 'border': True})
        string_format = workbook.add_format({'align': 'center','bg_color': 'black','font_color':'white', 'border': True})
        value_format = workbook.add_format({ 'text_wrap': True,'align': 'left', 'border': True})
        my_decimal_format = workbook.add_format({'num_format': '#,##0.00', 'border': True, 'align': 'right'})
        my_format = workbook.add_format({'num_format': '0.00', 'border': True, 'align': 'right'})

        start_date = False
        end_date = False
        internal_moves_partner = "DENTECO. S.A."

        if wizard.start_date:
            start_user_date = datetime(wizard.start_date.year, wizard.start_date.month, wizard.start_date.day)
            start_date = self._user_to_server_date(start_user_date, user_tz)

        if wizard.end_date:
            end_user_date = datetime(wizard.end_date.year, wizard.end_date.month, wizard.end_date.day, 23, 59)
            end_date = self._user_to_server_date(end_user_date, user_tz)

        domain_stock_move_ids = []
        if wizard.product_id:
            domain_stock_move_ids.append(('product_id', '=', wizard.product_id.id))
            domain_stock_move_ids.append(('state', '=', 'done'))
            if wizard.start_date:
                domain_stock_move_ids.append(('date', '>=', start_date))
            if wizard.end_date:
                domain_stock_move_ids.append(('date', '<=', end_date))
            stock_move_ids = self.env['stock.move'].search(domain_stock_move_ids, order="date asc")
            inventory_at_date = False
            if wizard.end_date:
                inventory_at_date = end_date
            onhand_qty = 1
            onhand_qty = wizard.product_id.with_context(location=wizard.warehouse.view_location_id.id, to_date=inventory_at_date, compute_child=False).qty_available
            report_date = _("KARDEX REPORT") + " DEL " + str(start_date or '') + " AL " + str(end_date or '')
            product_name = wizard.product_id.name + "  EXISTENCIA : " + str(onhand_qty)
            brand_name = _("Brand: ") + (wizard.product_id.brand_id.name or 'Sin marca')
            brand_code = _("Product Code : ") + (wizard.product_id.default_code or 'Sin código Marca')
            sheet.merge_range('B1:N1',report_date ,head)
            sheet.merge_range('B2:N2',product_name ,head)
            sheet.merge_range('B3:E3',_(""),format_merge)
            sheet.merge_range('F3:K3', brand_name, format_merge)
            sheet.merge_range('L3:N3', brand_code, format_merge)
            sheet.write(3,1,_("Effective Date"),string_format)
            sheet.write(3,2,_("Reference"),string_format)
            sheet.write(3,3,_("Document"),string_format)
            sheet.write(3,4,_("Responsible"),string_format)
            sheet.write(3,5,_("Type Of Operation"),string_format)
            sheet.write(3,6,_("Location Of Origin"),string_format)
            sheet.write(3,7,_("Location Of Destination"),string_format)
            sheet.write(3,8,_("Receive from"),string_format)
            sheet.write(3,9,_("Delivery to"),string_format)
            sheet.write(3,10,_("Inputs"),string_format)
            sheet.write(3,11,_("Outputs"),string_format)
            sheet.write(3,12,_("Balance"),string_format)
            sheet.write(3,13,_("Note"),string_format)
            row = 4
            qty_available = 0
            total_cost = 0
            balance = 0
            flag_ids = []
            for move in stock_move_ids:
                warehouse_id = wizard.warehouse
                if not warehouse_id in [move.location_id.warehouse_id, move.location_dest_id.warehouse_id] or move.location_id.warehouse_id == move.location_dest_id.warehouse_id:
                    continue
                picking_id = move.picking_id
                if picking_id:
                    picking_type = picking_id.picking_type_id.alias or picking_id.picking_type_id.name
                    sheet.write(row,1,self._server_to_user_datetime(picking_id.date_done, local_time), datetime_style)
                    sheet.write(row,2,_(picking_id.name) or None,value_format)
                    sheet.write(row,3,picking_id.origin or _(picking_id.name),value_format)
                    sheet.write(row,4,picking_id.user_id.name or move.write_uid.name, my_format)
                    sheet.write(row,5,picking_type or None, my_decimal_format)
                    sheet.write(row,6,picking_id.location_id.complete_name or None, my_decimal_format)
                    sheet.write(row,7,picking_id.location_dest_id.complete_name or None, my_format)
                    
                    if picking_id.picking_type_id.code == 'incoming':
                        sheet.write(row,8,picking_id.partner_id.name, my_decimal_format)
                        sheet.write(row,10,abs(move.quantity_done), my_format)
                    else:
                        sheet.write(row,8,None, my_decimal_format)
                        sheet.write(row,10, None, my_format)

                    if picking_id.picking_type_id.code == 'outgoing':
                        sheet.write(row,9,picking_id.partner_id.name, my_decimal_format)
                        sheet.write(row,11,abs(move.quantity_done), my_decimal_format)
                    else:
                        sheet.write(row,9,None, my_decimal_format)
                        sheet.write(row,11, None, my_decimal_format)

                    if picking_id.picking_type_id.code == 'internal':
                        if picking_id.location_dest_id.warehouse_id.id == warehouse_id.id:
                            sheet.write(row,8, picking_id.partner_id.name, my_decimal_format)
                            sheet.write(row,10,abs(move.quantity_done), my_format)
                        else:
                            sheet.write(row,8,None, my_decimal_format)
                            sheet.write(row,10, None, my_format)
                        if picking_id.location_id.warehouse_id.id == warehouse_id.id:
                            sheet.write(row,9,picking_id.partner_id.name, my_decimal_format)
                            sheet.write(row,11,abs(move.quantity_done), my_decimal_format)
                        else:
                            sheet.write(row,9, None, my_decimal_format)
                            sheet.write(row,11, None, my_decimal_format)

                    qty_available = wizard.product_id.with_context(location=warehouse_id.view_location_id.id, to_date=move.date, compute_child=False).qty_available
                    sheet.write(row,12,abs(qty_available), my_decimal_format)

                    notes = [picking_id.note, picking_id.sale_id.note if picking_id.sale_id else None]
                    text = " ".join(" ".join(html.fromstring(note).xpath("//text()")) for note in notes if note)

                    sheet.write(row, 13, text, my_decimal_format)
                    row += 1

                else:
                    reference = move.reference
                    origin = move.origin
                    picking_type = move.picking_type_id.alias or move.picking_type_id.name
                    if "Quantity Updated" in move.reference or "Quantity Confirmed" in move.reference or "Cantidad" in move.reference:
                        reference = "Cantidad Actualizada"
                        origin = "Ajuste de Inventario - %s"%(move.id)
                    if move.quantity_done == 0.0:
                        continue
                    sheet.write(row,1,move.date, datetime_style)
                    sheet.write(row,2,_(reference), value_format)
                    sheet.write(row,3,origin or reference, value_format)
                    sheet.write(row,4,move.write_uid.name, my_format)
                    sheet.write(row,5,picking_type or "Ajuste de Inventario", my_decimal_format)
                    sheet.write(row,6,move.location_id.complete_name, my_decimal_format)
                    sheet.write(row,7,move.location_dest_id.complete_name, my_format)
                    

                    if move.location_id.usage not in ["internal", "transit"] and move.location_dest_id.usage in ["internal", "transit"]:
                        sheet.write(row,8,internal_moves_partner, my_decimal_format)
                        sheet.write(row,10,abs(move.quantity_done), my_format)
                    else:
                        sheet.write(row,8,None, my_decimal_format)
                        sheet.write(row,10, None, my_format)
                    if move.location_id.usage in ["internal", "transit"] and move.location_dest_id.usage not in ["internal", "transit"]:
                        sheet.write(row,9,internal_moves_partner, my_decimal_format)
                        sheet.write(row,11,abs(move.quantity_done), my_decimal_format)
                    else:
                        sheet.write(row,9,None, my_decimal_format)
                        sheet.write(row,11, None, my_decimal_format)

                    qty_available = wizard.product_id.with_context(location=warehouse_id.view_location_id.id, to_date=move.date, compute_child=False).qty_available
                    sheet.write(row,12,abs(qty_available), my_decimal_format)
                    sheet.write(row,13,None, my_decimal_format)
                    row += 1