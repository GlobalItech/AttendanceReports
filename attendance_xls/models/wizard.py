from odoo import models, fields, api


class StockReport(models.TransientModel):
#     _name = "wizard.stock.history"
    _name = "wizard.attendance.reporthistory"
    _description = "Attendance Report"
 
#     warehouse = fields.Many2many('stock.warehouse', 'wh_wiz_rel', 'wh', 'wiz', string='Warehouse', required=True)
#     category = fields.Many2many('product.category', 'categ_wiz_rel', 'categ', 'wiz', string='Warehouse')

    date_daily = fields.Date('Date')
    date_from = fields.Date('Date From')
    date_to = fields.Date('Date To')
    
    report_type =  fields.Selection([
            ('dr','Daily Report'),
            ('mr','Monthly Report')
        ])
    report_daily =  fields.Selection([
            ('ar','Daily Absent Report'),
            ('dar','Daily Attendance Report'),
            ('dhlr','Daily Half Leave Report'),
            ('dlar','Daily Late Arrival Report'),
            ('delr','Daily Early Left Report'),
            ('dltr','Daily Leaves Types Report'),
        ],string="Select Daily Report")
    
    report_month =  fields.Selection([
            ('mr','Monthly Absent Report'),
            ('mar','Monthly Attendance Report'),
            ('mlar','Monthly Late Arrival Report'),
            ('melr','Monthly Early Left Report'),
            ('mhlr','Monthly Half Leaves Report'),
            ('mltwr','Monthly Leaves Type wise Report'),
            ('mewr','Monthly Employee wise Report'),
        ],string="Select Monthly Report")

    filterby1 =  fields.Selection([
            ('emp','Employee'),
            ('dep','Department'),
        ],string="Filter By")
    employee_id = fields.Many2one('hr.employee', string="Employee")
    department_id = fields.Many2one('hr.department',string="Department")    
    @api.multi
    def export_xls(self):
        context = self._context
        datas = {'ids': context.get('active_ids', [])}
        datas['model'] = 'hr.attendance'
        datas['form'] = self.read()[0]
        for field in datas['form'].keys():
            if isinstance(datas['form'][field], tuple):
                datas['form'][field] = datas['form'][field][0]
        if context.get('xls_export'):
            return {'type': 'ir.actions.report.xml',
                    'report_name': 'attendance_xls.attendance_report_xls.xlsx',
                    'datas': datas,
                    'name': 'Attendance Report'
                    }
