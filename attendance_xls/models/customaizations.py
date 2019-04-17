from odoo import api, fields, models,exceptions,_
import datetime
from dateutil.relativedelta import relativedelta
from datetime import date
from datetime import datetime, timedelta
from odoo.fields import Boolean
import calendar

from odoo.exceptions import UserError, AccessError, ValidationError


from lxml import etree
from odoo.osv.orm import setup_modifiers
from odoo.tools.safe_eval import safe_eval
from odoo.tools.float_utils import float_compare
from odoo.tools import float_is_zero
import json




class EmployeeCode(models.Model):
    _inherit = 'hr.employee'

    employe_code =fields.Char('Employee Code')
    
        
class attpolicymodified1(models.Model):
    _name = 'policy.form'

    time_mint=fields.Float("Relaxation of Time" )
    short_leave_policy = fields.Float(" Lates Allow in a Month")
    
#     Levae policy fields        
#             
     
    short_leaves = fields.Float("Short Leave Time",required=True)
    half_leaves=fields.Float("Half Leave Time" ,required=True)
    num_of_leaves_per_month = fields.Float("Short Leaves Allow in a Month")
    leavedaycut = fields.Float("Full Day Leave Time" )
    
            


    
    
    
    
 
#     
#     @api.multi
#     def unlink(self):
#         for leave in self:
#             if leave.state not in ('draft'):
#                 raise UserError(_('You cannot delete an leave policy form which is not draft or cancelled. You should ask the technical staff.'))
#             
#         return super(Leavepolicy, self).unlink()
#     
#     
#     
#     @api.model
#     def fields_view_get(self, view_id=None, view_type='form', toolbar=False, submenu=False):
#         result = super(Leavepolicy, self).fields_view_get(view_id, view_type, toolbar=toolbar, submenu=submenu)
#         if view_type=="form":
#             doc = etree.XML(result['arch'])
#             for node in doc.iter(tag="field"):
#                 if 'readonly' in node.attrib.get("modifiers",''):
#                     attrs = node.attrib.get("attrs",'')
#                     if 'readonly' in attrs:
#                         attrs_dict = safe_eval(node.get('attrs'))
#                         r_list = attrs_dict.get('readonly',)
#                         if type(r_list)==list:
#                             r_list.insert(0,('state','=','approve'))
#                             if len(r_list)>1:
#                                 r_list.insert(0,'|')
#                         attrs_dict.update({'readonly':r_list})
#                         node.set('attrs', str(attrs_dict))
#                         setup_modifiers(node, result['fields'][node.get("name")])
#                         continue
#                     else:
#                         continue
#                 node.set('attrs', "{'readonly':[('state','=','approve')]}")
#                 setup_modifiers(node, result['fields'][node.get("name")])
#                  
#             result['arch'] = etree.tostring(doc)
#         return result
#      
# 
#      
                  
    
# class Leavepolicy(models.Model):
#     _name='leave.policy.form'
#     _description="Leave  Policy Application"
#     _inherit = ['mail.thread', 'ir.needaction_mixin']
#     
#     name1 = fields.Char(string='Late Policy name' ,required=True )    
#     name = fields.Char(related="name1")
#     