import datetime
from odoo.addons.report_xlsx.report.report_xlsx import ReportXlsx
from odoo.tools import DEFAULT_SERVER_DATETIME_FORMAT
from collections import OrderedDict
from datetime import datetime
from datetime import date, timedelta
import re
import calendar
from odoo.exceptions import UserError, AccessError, ValidationError
from odoo import exceptions, _

class StockReportXls(ReportXlsx):

    TIME_FORMAT = re.compile("(24:00|2[0-3]:[0-5][0-9]|[0-1][0-9]:[0-5][0-9])")
    def get_warehouse(self, data):
        if data.get('form', False) and data['form'].get('display_name', False):
            l1 = []
            l2 = []
            model_name = data['form'].get('display_name', False).split(',')
            obj = self.env[model_name[0]].search([('id', 'in', [model_name[1]])])
            if(data['form']['report_daily']):
                name = dict(obj._fields['report_daily'].selection)[data['form']['report_daily']] 
            elif(data['form']['report_month']):
                name = dict(obj._fields['report_month'].selection)[data['form']['report_month']] 
#             for j in obj:
#                 l1.append(j.name)
#                 l2.append(j.id)
        return name

    def get_category(self, data):
        if data.get('form', False) and data['form'].get('category', False):
            l2 = []
            obj = self.env['product.category'].search([('id', 'in', data['form']['category'])])
            for j in obj:
                l2.append(j.id)
            return l2
        return ''
    #def func(self,r):
        
          
 # monthly late arrival
    def find_check_in_category(self,check_in,employee_id):
        policy = self.env['policy.form'].search([])
        emp_check_in_conver = str(datetime.strptime(check_in, '%Y-%m-%d %H:%M:%S')+timedelta(hours=5))
        check_late = self.env['policy.form'].search([])
        emp_check_in ='{0:02.0f}:{1:02.0f}'.format(*divmod(check_late.half_leaves * 60, 60))
        emp_checkout_fake = str(datetime.strptime(check_in, '%Y-%m-%d %H:%M:%S')+timedelta(minutes=5))
        my_date = datetime.strptime(check_in, '%Y-%m-%d %H:%M:%S').date()
        time_set_from = '{0:02.0f}:{1:02.0f}'.format(*divmod(employee_id.calendar_id.attendance_ids.search([('name','=',calendar.day_name[my_date.weekday()])]).hour_from * 60, 60)) #self.employee_id.calendar_id.attendance_ids.search([('name','=',calendar.day_name[my_date.weekday()])])#'{0:02.0f}:{1:02.0f}'.format(*divmod(policy.time_sets_from * 60, 60))
        late = '{0:02.0f}:{1:02.0f}'.format(*divmod(policy.time_mint * 60, 60))
        
        splitcheck = emp_check_in_conver.split(' ')
        set_time_val = datetime.strptime(time_set_from, '%H:%M')
        split_late = late.split(':')
        time_buffer = set_time_val +timedelta(minutes=int(split_late[1]))
        res_time_buffer = str(time_buffer).split(' ')

        if(splitcheck[1] > res_time_buffer[1] and splitcheck[1] < emp_check_in):
            return splitcheck
        elif(splitcheck[1] > res_time_buffer[1] and splitcheck[1] > emp_check_in):
            return splitcheck
        
    def early_left(self,check_outemp,employee_id):
        policy = self.env['policy.form'].search([])
        short_late_policy = self.env['policy.form'].search([])
        my_date = datetime.strptime(check_outemp, '%Y-%m-%d %H:%M:%S').date()
        #test
        pt=self.env['resource.calendar'].search([])
        s=calendar.day_name[my_date.weekday()]
        chkoutday=""
        off_timeof_chkoutday=0.0
        for e in pt:
            for h in e.attendance_ids:
                p=dict(h.fields_get(allfields=['dayofweek'])['dayofweek']['selection'])[h.dayofweek]
                if s==p:
                    off_timeof_chkoutday=h.hour_to
                    chkoutday=s
        #test
        check_out1 ='{0:02.0f}:{1:02.0f}'.format(*divmod(off_timeof_chkoutday*60,60))# '{0:02.0f}:{1:02.0f}'.format(*divmod(employee_id.calendar_id.attendance_ids.search([('name','=',calendar.day_name[my_date.weekday()])]).hour_to * 60, 60))#'{0:02.0f}:{1:02.0f}'.format(*divmod(policy.time_sets_to * 60, 60))
        check_out = str(datetime.strptime(check_out1, '%H:%M')) #+timedelta(hours=12)).split(' ')[1]
        short_leaves = '{0:02.0f}:{1:02.0f}'.format(*divmod(short_late_policy.short_leaves * 60, 60))
        full_day_leave_cut ='{0:02.0f}:{1:02.0f}'.format(*divmod(short_late_policy.leavedaycut * 60, 60))
        count_shortLeaves = 0
        
        emp_check_out = str(datetime.strptime(check_outemp, '%Y-%m-%d %H:%M:%S')+timedelta(hours=5))
        
        if((emp_check_out.split(' ')[1] >= short_leaves) and (emp_check_out.split(' ')[1] < check_out.split(' ')[1])):
            return emp_check_out
    
    def half_leave(self,check_outemp,employee_id):
        policy = self.env['policy.form'].search([])
        short_late_policy = self.env['policy.form'].search([])
        my_date = datetime.strptime(check_outemp, '%Y-%m-%d %H:%M:%S').date()
        check_out1 = '{0:02.0f}:{1:02.0f}'.format(*divmod(employee_id.calendar_id.attendance_ids.search([('name','=',calendar.day_name[my_date.weekday()])]).hour_to * 60, 60))#'{0:02.0f}:{1:02.0f}'.format(*divmod(policy.time_sets_to * 60, 60))
        check_out = str(datetime.strptime(check_out1, '%H:%M')) #+timedelta(hours=12)).split(' ')[1]
        short_leaves = '{0:02.0f}:{1:02.0f}'.format(*divmod(short_late_policy.short_leaves * 60, 60))
        full_day_leave_cut ='{0:02.0f}:{1:02.0f}'.format(*divmod(short_late_policy.leavedaycut * 60, 60))
        count_shortLeaves = 0
        
        emp_check_out = str(datetime.strptime(check_outemp, '%Y-%m-%d %H:%M:%S')+timedelta(hours=5))
        
        if(emp_check_out.split(' ')[1] <= short_leaves ):
            return emp_check_out
    
    def count_days(self,date1,date2):
        d1 = date1.split(' ')[0].split('-')
        d2 = date2.split(' ')[0].split('-')
        
        d0 = date(int(d1[0]), int(d1[1]), int(d1[2]))
        d1 = date(int(d2[0]), int(d2[1]), int(d2[2]))
        delta = d1 - d0
        return delta.days+1
    def get_dayName_from_date(self,date_To_day):
        rprtdate=datetime.strptime(date_To_day, '%Y-%m-%d')
         
        rprtday=rprtdate.strftime("%A")
        print(rprtday)
        return rprtday
        
    def get_formatted_time (self, hour):
        hou,sec = divmod(hour * 60*60, 3600) 
        mn, sec = divmod(sec, 60) 
        work_time = "{:02.0f}:{:02.0f}:{:02.0f}".format(hou, mn, sec)
        return datetime.strptime(work_time,'%H:%M:%S')
    
    
    
    def calc_time_difference(self,check_InOut_time,dt_work_time):
        rprt_id=self.env.context.get('active_id', False)
        rprt=self.env['wizard.attendance.reporthistory'].search([('id','=', rprt_id)])
        chek_inout_time=check_InOut_time
        chek_inout_time_todatetime=datetime.strptime(chek_inout_time,'%H:%M:%S')
           
        if(((rprt.report_type == 'mr') and (rprt.report_month == 'mlar')) or ((rprt.report_type == 'dr') and(rprt.report_daily == 'dlar'))):
            differ=chek_inout_time_todatetime-dt_work_time
            late_arrival_time=str(timedelta(seconds=differ.seconds))
            return late_arrival_time
        
        else:
                if(((rprt.report_type == 'mr') and(rprt.report_month == 'melr')) or ((rprt.report_type == 'dr') and(rprt.report_daily == 'delr'))):
                    differ=dt_work_time-chek_inout_time_todatetime
                    early_left_time=str(timedelta(seconds=differ.seconds))
                    return early_left_time
                
                
                
            
    
    def get_lines(self, data, reporttype):
        lines = []
        club_lines = []
        attendance = self.env[data['model']].search([])
        categ = self.get_category(data)
        leave_types = self.env['hr.holidays.status'].search([])
        
        #code
        if(data['form']['date_daily']):
            rprt=self.env['wizard.attendance.reporthistory'].search([('id','=',data['form']['id'])])
        #if rprt.report_type == 'dr':
            #rprtdate=datetime.strptime(rprt.date_daily, '%Y-%m-%d')
         
            #rprtday=rprtdate.strftime("%A")
           # print(rprtday)
            rprtday=self.get_dayName_from_date(rprt.date_daily)
            pt= self.env['resource.calendar'].search([])
        #hr_fro=self.env['resource.calendar.attendance'].search([])
#        strt_time=[]
#        off_time=[]
            h_frm=0.0
            h_to=0.0   
            for hr in pt:
                for e in hr.attendance_ids:
                #getting the value of selection field instead of index 
                    dayname=dict(e.fields_get(allfields=['dayofweek'])['dayofweek']['selection'])[e.dayofweek]
                
                
#                 st=e.hour_from
#                 oft=e.hour_to
#                 print(dayname,st,oft)
#                 dayName.append(dayname)
#                 strt_time.append(st)
#                 off_time.append(oft)
                    if rprtday == dayname:
                        h_frm =  e.hour_from 
                        h_to=  e.hour_to 
                        break
#                    strt_time.append(h_frm)
#            h_frm         off_time.append(h_to)   
#             hou,sec = divmod(h_frm * 60*60, 3600) 
#             mn, sec = divmod(sec, 60) 
#             work_from = "{:02.0f}:{:02.0f}:{:02.0f}".format(hou, mn, sec)
#         
#             h,s=divmod(h_to * 60*60, 3600) 
#             m, s = divmod(s, 60)
#             work_to = "{:02.0f}:{:02.0f}:{:02.0f}".format(h, m, s)
#         
#             dt_work_from=datetime.strptime(work_from,'%H:%M:%S')
#             dt_work_to=datetime.strptime(work_to,'%H:%M:%S')       
              
                dt_work_from = self.get_formatted_time(h_frm)
                dt_work_to = self.get_formatted_time(h_to)
 
        #code   

        
        sr =1
        day=0
        res = {}
        emp_outscope = 0
        if(data['form']['date_from']):
            stringdate_to_date = datetime.strptime(data['form']['date_from'], '%Y-%m-%d')
            to_datetime = datetime.strftime(stringdate_to_date, DEFAULT_SERVER_DATETIME_FORMAT)
            end_day_time = datetime.strptime(data['form']['date_to'], '%Y-%m-%d')
            ends = datetime.strftime(end_day_time, DEFAULT_SERVER_DATETIME_FORMAT)
            
            
            #code
            all_dates=[]
            rprt=self.env['wizard.attendance.reporthistory'].search([('id','=',data['form']['id'])]) 
            rprt_date_from = rprt.date_from
#            rprt_date_from_todatetime= datetime.strptime(rprt_date_from,'%Y')
             
            rprt_date_to= rprt.date_to
             
            while(stringdate_to_date <= end_day_time):
                all_dates.append(stringdate_to_date)
                stringdate_to_date =stringdate_to_date + timedelta(days =1)
              
            pt= self.env['resource.calendar'].search([])
             
            h_frm=[]
            h_to=[]
            for i in range(len(all_dates)):
                rprtday=all_dates[i].strftime("%A")
                print(rprtday) 
                for hr in pt:
                    for e in hr.attendance_ids:
                     
                         
                        dayname=dict(e.fields_get(allfields=['dayofweek'])['dayofweek']['selection'])[e.dayofweek]
                         
                        if rprtday == dayname:
                            h_frm.extend([[rprtday,e.hour_from]]) 
                            h_to.extend([[rprtday,e.hour_to]]) 
                         
            #all_dates[i]=all_dates[i]+timedelta(days=1)  
#             
#             dt_wrk_from=[]
#             dt_wrk_to=[]
#             for i in range(len(h_frm)):
#                 for j in range(len(h_frm[i])):
#                     
#                     hou,sec=divmod(h_frm[j][1] * 60*60, 3600) 
#                     mn, sec = divmod(sec, 60) 
#                     wrk_from ="{:02.0f}:{:02.0f}:{:02.0f}".format(hou, mn, sec)
#                  
#                     t1= datetime.strptime(wrk_from,'%H:%M:%S')
#                     dt_wrk_from.append(t1)
#                  
#             for k in range(len(h_to)): 
#                 for m in range(len(h_to[k])):   
#                     h,s=divmod(h_to[m][1] * 60*60, 3600) 
#                     m, s = divmod(s, 60)
#                     wrk_to="{:02.0f}:{:02.0f}:{:02.0f}".format(h, m, s)
#                  
#                     t2= datetime.strptime(wrk_to,'%H:%M:%S')
#                     dt_wrk_to.append(t2)   
                       
            #code  
            
            
            
            
            
            
            
            
            
        
        if(data['form']['employee_id']):
            emp_outscope = self.env['hr.employee'].search([('id','=',data['form']['employee_id'])])
        elif(data['form']['department_id']):
            emp_outscope = self.env['hr.employee'].search([('department_id','=',data['form']['department_id'])])
        else:
            emp_outscope = self.env['hr.employee'].search([])
        if(data['form']['report_month'] == 'mr'):
            
            #original code
            start = int(data['form']['date_from'].split('-')[2])-1
            enddate = int(data['form']['date_to'].split('-')[2])-1
            employee = self.env['hr.employee'].search([('attendance_ids','=',False)])
            #testint
            d1 = int(data['form']['date_from'].split('-')[2])
            d2 = int(data['form']['date_to'].split('-')[2])+1 
            date1=d2-d1
            
            em=self.env['wizard.attendance.reporthistory'].search([('id','=',data['form']['id'])])
            if (em.filterby1=='emp'):
                prlist=[]
                filteredemp=em.employee_id
                if filteredemp.attendance_ids:
                    for i in filteredemp.attendance_ids:
                        attd=i.check_in.split(' ')[0]
                        attd_todate=datetime.strptime(attd,"%Y-%m-%d")
                        prlist.append(attd_todate)
                for start in range(date1): 
                    stringdate_to_date = datetime.strptime(data['form']['date_from'], '%Y-%m-%d') + timedelta(start)
                    if stringdate_to_date not in prlist:
                        to_datetime = datetime.strftime(stringdate_to_date, DEFAULT_SERVER_DATETIME_FORMAT)
                        end_day_time = datetime.strptime(data['form']['date_to'], '%Y-%m-%d') + timedelta(start)
                        ends = datetime.strftime(end_day_time, DEFAULT_SERVER_DATETIME_FORMAT)
                        
                        emp_abesnt = self.env['hr.attendance'].search([('employee_id.id','=',filteredemp.id),('check_in','>=',to_datetime),('check_out','<=',ends)])
                        
                        vals = {
    #                                 'sr':sr,
                                    'empcode':filteredemp.employe_code,
                                    'name': filteredemp.name,
                                    #'status': filteredemp.attendance_ids or 'Absent',
                                    'status':'Absent',
                                    'date':str(stringdate_to_date).split(' ')[0],
                                    'emp_id':filteredemp.id,
                                }
                        lines.append(vals)
                       
            else:    
            #testing
                #for fetching data of all employees
                employee=emp_outscope
                
                for e in employee:
                    prlist=[]
                    if e.attendance_ids:
                        for i in e.attendance_ids:
                            attd=i.check_in.split(' ')[0]
                            attd_todate=datetime.strptime(attd,"%Y-%m-%d")
                            prlist.append(attd_todate)
                        for start in range(date1): 
                            stringdate_to_date = datetime.strptime(data['form']['date_from'], '%Y-%m-%d') + timedelta(start)
                            if stringdate_to_date not in prlist:
                                to_datetime = datetime.strftime(stringdate_to_date, DEFAULT_SERVER_DATETIME_FORMAT)
                                end_day_time = datetime.strptime(data['form']['date_to'], '%Y-%m-%d') + timedelta(start)
                                ends = datetime.strftime(end_day_time, DEFAULT_SERVER_DATETIME_FORMAT)
                        
                                emp_abesnt = self.env['hr.attendance'].search([('employee_id.id','=',e.id),('check_in','>=',to_datetime),('check_out','<=',ends)])
                        
                                vals = {
    #                                 'sr':sr,
                                    'empcode':e.employe_code,
                                    'name': e.name,
                                    #'status': filteredemp.attendance_ids or 'Absent',
                                    'status':'Absent',
                                    'date':str(stringdate_to_date).split(' ')[0],
                                    'emp_id':e.id,
                                     }
                                lines.append(vals) 
                        del prlist[:]           
                    #for start in range(enddate):
                    else:
                        for start in range(date1): 
                            stringdate_to_date = datetime.strptime(data['form']['date_from'], '%Y-%m-%d') + timedelta(start)
                            to_datetime = datetime.strftime(stringdate_to_date, DEFAULT_SERVER_DATETIME_FORMAT)
                            end_day_time = datetime.strptime(data['form']['date_to'], '%Y-%m-%d') + timedelta(start)
                            ends = datetime.strftime(end_day_time, DEFAULT_SERVER_DATETIME_FORMAT) 
                            #emp_abesnt = self.env['hr.attendance'].search([('employee_id.id','=',e.id),('check_in','>=',to_datetime),('check_out','<=',to_datetime)])
            #                 self.env['hr.holidays'].search([('date_from','>=',to_datetime),('date_to','<=',ends),('state','=','validate'),('employee_id','=',e.id)])
                            emp_abesnt = self.env['hr.attendance'].search([('employee_id.id','=',e.id),('check_in','>=',to_datetime),('check_out','<=',ends)])
                    #original
        #              #just testing
        #             
        #            
        #             #just testing
                            vals = {
        #                                 'sr':sr,
                                        'empcode': e.employe_code,
                                        'name': e.name,
                                        'status': e.attendance_ids or 'Absent',
                                        'date':str(stringdate_to_date).split(' ')[0],
                                        'emp_id':e.id,
                                    }
                            
        #                     sr =sr+1
                            lines.append(vals)
                        
        elif(data['form']['report_month'] == 'mar'):
            day +=1
            employee = emp_outscope#self.env['hr.employee'].search([])
            stringdate_to_date = datetime.strptime(data['form']['date_from'], '%Y-%m-%d')
            to_datetime = datetime.strftime(stringdate_to_date, DEFAULT_SERVER_DATETIME_FORMAT)
            end_day_time = datetime.strptime(data['form']['date_to'], '%Y-%m-%d')
            ends = datetime.strftime(end_day_time, DEFAULT_SERVER_DATETIME_FORMAT)
            dey = ((datetime.strptime(ends, DEFAULT_SERVER_DATETIME_FORMAT)))+ timedelta(1)
            
            
            for e in employee:
                #emp_abesnt = self.env['hr.attendance'].search([('employee_id.id','=',e.id),('check_in','>=',to_datetime),('check_in','<=',ends)])
                emp_abesnt = self.env['hr.attendance'].search([('employee_id.id','=',e.id),('check_in','>=',to_datetime),('check_in','<=',str(dey))])
 
                for a in emp_abesnt:    
                    if(a.check_in):
                        check_inn = datetime.strptime(a.check_in, '%Y-%m-%d %H:%M:%S').strftime('%d-%m-%Y %H:%M:%S')
                        if(a.check_out): 
                            check_outt = datetime.strptime(a.check_out, '%Y-%m-%d %H:%M:%S').strftime('%d-%m-%Y %H:%M:%S')
                            vals = {
                                    #'sr':sr,
                                    'empcode': e.employe_code,
                                    'name': e.name,
                                    'status': str(datetime.strptime(check_inn, '%d-%m-%Y %H:%M:%S')+timedelta(hours=5))+' '+str(datetime.strptime(check_outt, '%d-%m-%Y %H:%M:%S')+timedelta(hours=5)) or 'Absent',#str(datetime.strptime(a.check_in, '%Y-%m-%d %H:%M:%S').strftime('%d-%m-%Y %H:%M:%S')+timedelta(hours=5))+' '+str(datetime.strptime(a.check_out, '%Y-%m-%d %H:%M:%S').strftime('%d-%m-%Y %H:%M:%S')+timedelta(hours=5)) or 'Absent',
                                    'emp_id':e.id,
                                }
                        else:
                            vals = {
                                    
                                    #'sr':sr,
                                    'empcode': e.employe_code,
                                    'name': e.name,
                                    'status': str(datetime.strptime(check_inn, '%d-%m-%Y %H:%M:%S')+timedelta(hours=5)) or 'Absent',
                                    'emp_id':e.id,
                                }
                                
                        sr =sr+1
                        lines.append(vals)
                           
        
        elif(data['form']['report_month'] == 'mlar'):
#             i=0
            
            employee = emp_outscope#self.env['hr.employee'].search([])
            dey = ((datetime.strptime(ends, DEFAULT_SERVER_DATETIME_FORMAT)))+ timedelta(1)
            for e in employee:
                emp_abesnt = self.env['hr.attendance'].search([('employee_id.id','=',e.id),('check_in','>=',to_datetime),('check_in','<=',str(dey))])
                
                for a in emp_abesnt:    
                    res1 = self.find_check_in_category(a.check_in,e)
#                     types = a.reason1.split(' ')
#                     if('Late' in types):
                    
                    if(res1):
                        
                        #code
                        g=[]
                        rsd= datetime.strptime(res1[0],'%Y-%m-%d')
                        c= all_dates.index(rsd)
                        rday=all_dates[c].strftime("%A")
                        for i in h_frm:
                            if(rday==i[0]):
                                
                                g.append(i)
                        
                        #hou,sec=divmod(g[0][1] * 60*60, 3600) 
                        #mn, sec = divmod(sec, 60) 
                        #work_from = "{:02.0f}:{:02.0f}:{:02.0f}".format(hou, mn, sec)
                        #dt_work_from=datetime.strptime(work_from,'%H:%M:%S')    
                        dt_work_from=self.get_formatted_time(g[0][1]) 
#                         
#                        
                                                  
#                         chek_in_time=res1[1]
#                         chek_in_time_todatetime=datetime.strptime(chek_in_time,'%H:%M:%S')
#                         differ=chek_in_time_todatetime-dt_work_from
#                         late_arrival_time=str(timedelta(seconds=differ.seconds))
                        late_arrival_time=self.calc_time_difference(res1[1], dt_work_from)
                        vals = {
                                #'sr':sr,
                                'empcode': e.employe_code,
                                'name': e.name,
                                'status': datetime.strptime(res1[0], '%Y-%m-%d').strftime('%d-%m-%Y') +' '+ res1[1] or 'Absent' ,#a.date_from+' '+types[3] or 'Absent',
                                'emp_id':e.id,
                                'late' : late_arrival_time
                            }
                        
                        sr =sr+1
                        lines.append(vals)
        elif(data['form']['report_month'] == 'melr'):
            #just testing
            stringdate_to_date = datetime.strptime(data['form']['date_from'], '%Y-%m-%d')
            end_day_time = datetime.strptime(data['form']['date_to'], '%Y-%m-%d')
            dey = ((datetime.strptime(ends, DEFAULT_SERVER_DATETIME_FORMAT)))+ timedelta(1)
            #just testing
            employee =emp_outscope #self.env['hr.employee'].search([('check_in','>=',to_datetime),('check_out','<=',ends)])#self.env['leave.form'].search([('name.id','=',e.id),('date_from','>=',data['form']['date_from']),('date_to','<=',data['form']['date_to'])])
             
            
            #just testing
            #employee =self.env['hr.attendance'].search([])
            #just testing
            for e in employee:
                
                emp_abesnt = self.env['hr.attendance'].search([('employee_id.id','=',e.id),('check_in','>=',to_datetime),('check_out','<=',str(dey))])#self.env['leave.form'].search([('name.id','=',e.id),('date_from','>=',data['form']['date_from']),('date_to','<=',data['form']['date_to'])])
                #emp_abesnt = self.env['hr.attendance'].search([('check_in','>=',to_datetime),('check_out','<=',ends)])#self.env['leave.form'].search([('name.id','=',e.id),('date_from','>=',data['form']['date_from']),('date_to','<=',data['form']['date_to'])])
    
                for a in emp_abesnt:    
                    res1 = self.early_left(a.check_out,e)
#                        types = a.reason1.split(' ')
#                        if('Short' in types):
                    if(res1):
                        #code
                        r=res1.split(" ")
                        #i=datetime.strptime(res1, "%Y-%m-%d %H:%M:%S") 
                        #chek_out_time=str(i.time())
                        #chek_out_time_todatetime=datetime.strptime(chek_out_time,'%H:%M:%S')
                        #differ=dt_work_to-chek_out_time_todatetime
                        #early_left_time=str(timedelta(seconds=differ.seconds))  
                        g=[]
                        rsd= datetime.strptime(r[0],'%Y-%m-%d')
                        c= all_dates.index(rsd)
                        rday=all_dates[c].strftime("%A")
                        for i in h_to:
                            if(rday==i[0]): 
                                g.append(i)
                        
                        dt_work_to=self.get_formatted_time(g[0][1])     
#                         hou,sec=divmod(g[0][1] * 60*60, 3600) 
#                         mn, sec = divmod(sec, 60) 
#                         work_to = "{:02.0f}:{:02.0f}:{:02.0f}".format(hou, mn, sec)
#                         dt_work_to=datetime.strptime(work_to,'%H:%M:%S')    
# #                           
# #                            
# #                           
#                                                      
#                         chek_out_time=r[1]
#                         chek_out_time_todatetime=datetime.strptime(chek_out_time,'%H:%M:%S')
#                         differ=dt_work_to-chek_out_time_todatetime
#                         early_left_time=str(timedelta(seconds=differ.seconds))
                        early_left_time=self.calc_time_difference(r[1], dt_work_to)
                        #code
                        vals = {
                                   #'sr':sr,
                                   'empcode':e.employe_code,
                                   'name': e.name,
                                   'status': res1 or 'Absent',
                                   'emp_id':e.id,
                                   'earlyleft':early_left_time
                               }
                        sr =sr+1
                        lines.append(vals)
        elif(data['form']['report_month'] == 'mltwr'):
            employee = emp_outscope#self.env['hr.employee'].search([])
            stringdate_to_date = datetime.strptime(data['form']['date_from'], '%Y-%m-%d')
            to_datetime = datetime.strftime(stringdate_to_date, DEFAULT_SERVER_DATETIME_FORMAT)
            end_day_time = datetime.strptime(data['form']['date_to'], '%Y-%m-%d')
            ends = datetime.strftime(end_day_time, DEFAULT_SERVER_DATETIME_FORMAT)

            for e in employee:
                print(e,e.name)
#                 print(e.name)
                holidays = self.env['hr.holidays'].search([('employee_id.id','=',e.id),('type','=','remove'),('date_from','>=',to_datetime),('date_to','<=',ends)])#self.env['hr.holidays'].search([('employee_id.id','=',e.id),('type','=','remove')])
                for h in holidays:
                    if(h.holiday_status_id):
                        r = self.count_days(h.date_from,h.date_to)
                        vals = {
                                'sr':sr,
                                'empcode': e.employe_code,
                                'name': e.name,
                                'holiday_name':h.holiday_status_id.name,
                                'holiday_id':h.holiday_status_id.id,
                                'emp_id':e.id,
                                'datefrom':h.date_from.split(' ')[0],
                                'dateto':h.date_to.split(' ')[0],
                                'count':r
                            }
                        sr =sr+1
                        lines.append(vals)
        
        elif(data['form']['report_month'] == 'mhlr'):
            employee = emp_outscope# self.env['hr.employee'].search([])
            dey = ((datetime.strptime(ends, DEFAULT_SERVER_DATETIME_FORMAT)))+ timedelta(1)
            for e in employee:
                emp_abesnt = self.env['hr.attendance'].search([('employee_id.id','=',e.id),('check_out','>=',to_datetime),('check_out','<=',str(dey))])#self.env['leave.form'].search([('name.id','=',e.id),('date_from','>=',data['form']['date_from']),('date_to','<=',data['form']['date_to'])])
                if(e.id == 175):
                        pass
                for a in emp_abesnt: 
                     
                    res1 = self.half_leave(a.check_out,e)
#                     types = a.reason1.split(' ')
#                     if('Half' in types):
                    if(res1):
                        vals = {
                                #'sr':sr,
                                'empcode': e.employe_code,
                                'name': e.name,
                                'status': res1 or 'Absent',
                                'emp_id':e.id,
                            }
                        sr =sr+1
                        lines.append(vals)
        
        elif(data['form']['report_month'] == 'mewr'):
            employee = emp_outscope#self.env['hr.employee'].search([])
            stringdate_to_date = datetime.strptime(data['form']['date_from'], '%Y-%m-%d')
            to_datetime = datetime.strftime(stringdate_to_date, DEFAULT_SERVER_DATETIME_FORMAT)
            end_day_time = datetime.strptime(data['form']['date_to'], '%Y-%m-%d')
            ends = datetime.strftime(end_day_time, DEFAULT_SERVER_DATETIME_FORMAT)
#             dic = DefaultListOrderedDict()
            for e in employee:
                print(e,e.name)
#                 print(e.name)
                holidays = self.env['hr.holidays'].search([('employee_id.id','=',e.id),('type','=','remove'),('date_from','>=',to_datetime),('date_to','<=',ends)])#self.env['hr.holidays'].search([('employee_id.id','=',e.id),('type','=','remove')])
                for h in holidays:
                    if(h.holiday_status_id):
                        datime_from=datetime.strptime(h.date_from.split(' ')[0],'%Y-%m-%d').strftime('%d-%m-%Y'),h.date_from.split(' ')[1]
                        vals = {
                                'sr':sr,
                                'empcode': e.employe_code,
                                'name': e.name,
                                'holiday_name':h.holiday_status_id.name,
                                'holiday_id':h.holiday_status_id.id,
                                'date_from':datetime.strptime(h.date_from.split(' ')[0],'%Y-%m-%d').strftime('%d-%m-%Y'),
                                'date_to':datetime.strptime(h.date_from.split(' ')[0],'%Y-%m-%d').strftime('%d-%m-%Y'),
                                'emp_id':e.id,
                                'rep_type':'mewr',
                            }
                        sr =sr+1
                        lines.append(vals)
        
        
        """                    
         daily arrival report
        """
        if(data['form']['report_daily'] == 'ar'):
            employee = emp_outscope.search([('attendance_ids','=',False)])#emp_outscope#self.env['hr.employee'].search([('attendance_ids','=',False)])
            stringdate_to_date = datetime.strptime(data['form']['date_daily'], '%Y-%m-%d')
            to_datetime = datetime.strftime(stringdate_to_date, DEFAULT_SERVER_DATETIME_FORMAT)
            
            #testing
            for e in emp_outscope:
                
                if  e.attendance_ids:
                    prlist=[]
                    for i in e.attendance_ids:
                        date_of_employee=i[0].check_in.split(' ')[0]
                        date_of_employee_to_date_time=datetime.strptime(date_of_employee, '%Y-%m-%d')
                        prlist.append(date_of_employee_to_date_time)
                    if stringdate_to_date not in prlist:
                        vals={
                                  'sr':sr,
                                'empcode': e.employe_code,
                                'name': e.name,
                                'status':  'Absent',
                                }
                        sr =sr+1
                        lines.append(vals)
                        prlist[:]    
                else: 
                    vals={
                           'sr':sr,
                            'empcode': e.employe_code,
                            'name': e.name,
                            'status':  'Absent',
                                }
                    sr =sr+1
                    lines.append(vals)           
            #testing
#             for e in employee:
# #                 emp_abesnt = self.env['hr.attendance'].search([('employee_id.id','=',e.id),('check_in','>=',to_datetime),('check_out','<=',to_datetime)])
#                 vals = {
#                         'sr':sr,
#                         'empcode': e.employe_code,
#                         'name': e.name,
#                         'status': e.attendance_ids or 'Absent',
#                         
#                     }
#                 sr =sr+1
#                 lines.append(vals)
        elif(data['form']['report_daily'] == 'dar'):
            day +=1
            employee = emp_outscope#self.env['hr.employee'].search([])
            stringdate_to_date = datetime.strptime(data['form']['date_daily'], '%Y-%m-%d')
            to_datetime = datetime.strftime(stringdate_to_date, DEFAULT_SERVER_DATETIME_FORMAT)
#             next_date =  ((datetime.strptime(stringdate_to_date, DEFAULT_SERVER_DATETIME_FORMAT)))+ timedelta(day)
            end_day_time = stringdate_to_date + timedelta(hours=24)
            ends = datetime.strftime(end_day_time, DEFAULT_SERVER_DATETIME_FORMAT)
            for e in employee:
#                 emp_abesnt = self.env['hr.attendance'].search([('employee_id.id','=',e.id),('check_in','>=',to_datetime)])
                emp_abesnt = self.env['hr.attendance'].search([('employee_id.id','=',e.id),('check_in','>=',to_datetime),('check_in','<=',ends)])
                for eb in emp_abesnt:
                    if(eb.check_in):
                        check_inn = datetime.strptime(eb.check_in, '%Y-%m-%d %H:%M:%S').strftime('%d-%m-%Y %H:%M:%S')
                        if(eb.check_out):  
                            check_outt = datetime.strptime(eb.check_out, '%Y-%m-%d %H:%M:%S').strftime('%d-%m-%Y %H:%M:%S')
                            vals = {
                                    'sr':sr,
                                    'empcode': e.employe_code,
                                    'name': e.name,
                                    'status': str(datetime.strptime(check_inn, '%d-%m-%Y %H:%M:%S')+timedelta(hours=5))+' '+str(datetime.strptime(check_outt, '%d-%m-%Y %H:%M:%S')+timedelta(hours=5)) or 'Absent',
                                    
                                }
                        else:
                            vals = {
                                    'sr':sr,
                                    'empcode': e.employe_code,
                                    'name': e.name,
                                    'status': str(datetime.strptime(check_inn, '%d-%m-%Y %H:%M:%S')+timedelta(hours=5)) or 'Absent',
                                    
                                }
                        sr =sr+1
                        lines.append(vals)
        
        elif(data['form']['report_daily'] == 'dlar'):
            employee = emp_outscope#self.env['hr.employee'].search([])
            stringdate_to_date = datetime.strptime(data['form']['date_daily'], '%Y-%m-%d')
            to_datetime = datetime.strftime(stringdate_to_date, DEFAULT_SERVER_DATETIME_FORMAT)
#             next_date =  ((datetime.strptime(stringdate_to_date, DEFAULT_SERVER_DATETIME_FORMAT)))+ timedelta(day)
            end_day_time = stringdate_to_date + timedelta(hours=24)
            ends = datetime.strftime(end_day_time, DEFAULT_SERVER_DATETIME_FORMAT)
            
            for e in employee:
#                 emp_abesnt = self.env['leave.form'].search([('name.id','=',e.id),('date_from','>=',data['form']['date_daily']),('date_to','<=',data['form']['date_daily'])])
                emp_abesnt = self.env['hr.attendance'].search([('employee_id.id','=',e.id),('check_in','>=',to_datetime),('check_in','<=',ends)])
#                 eatt = self.env['hr.attendance'].search([('employee_id.id','=',e.id),('check_in','>=',to_datetime),('check_out','<=',ends)])
#                 emp_abesnt = self.env['hr.attendance'].search([('employee_id.id','=',e.id),('check_in','>=',to_datetime),('check_in','<=',ends)])
                
                for eb in emp_abesnt:
#                     types = eb.reason1.split(' ')
                    #test
                    s='13:00:00'
                    st=datetime.strptime(s,'%H:%M:%S')
                    #test 
                    res1 = self.find_check_in_category(eb.check_in,e)
#                     if('Half' in types):
                    if(res1):
                        late_arrival_time=self.calc_time_difference(res1[1], dt_work_from)
                    #code
#                         chek_in_time=res1[1]
#                         chek_in_time_todatetime=datetime.strptime(chek_in_time,'%H:%M:%S')
#                         differ=chek_in_time_todatetime-dt_work_from 
#                         late_arrival_time=str(timedelta(seconds=differ.seconds))
                            
                                  
                    #code                        
                        vals = {
                                'sr':sr,
                                'empcode': e.employe_code,
                                'name': e.name,
                                'status': datetime.strptime(res1[0], '%Y-%m-%d').strftime('%d-%m-%Y') +' '+ res1[1] or 'Absent' ,#eatt.check_in+' '+eatt.check_out or 'Absent',
                                'late': late_arrival_time
#                                     'status': types[3] or 'Absent',
                                
                            }
                        sr =sr+1
                        lines.append(vals)
#                     elif('Late' in types):
#                         vals = {
#                                 'sr':sr,
#                                 'empcode': e.employe_code,
#                                 'name': e.name,
#                                 'status': eatt.check_in+' '+eatt.check_out or 'Absent',
# #                                     'status': types[3] or 'Absent',
#                                 
#                             }
#                         sr =sr+1
#                         lines.append(vals)
        elif(data['form']['report_daily'] == 'dhlr'):
            employee = emp_outscope#self.env['hr.employee'].search([])
            stringdate_to_date = datetime.strptime(data['form']['date_daily'], '%Y-%m-%d')
            to_datetime = datetime.strftime(stringdate_to_date, DEFAULT_SERVER_DATETIME_FORMAT)
#             next_date =  ((datetime.strptime(stringdate_to_date, DEFAULT_SERVER_DATETIME_FORMAT)))+ timedelta(day)
            end_day_time = stringdate_to_date + timedelta(hours=24)
            ends = datetime.strftime(end_day_time, DEFAULT_SERVER_DATETIME_FORMAT)
            
            for e in employee:
#                 emp_abesnt = self.env['leave.form'].search([('name.id','=',e.id),('date_from','>=',data['form']['date_daily']),('date_to','<=',data['form']['date_daily'])])
                emp_abesnt = self.env['hr.attendance'].search([('employee_id.id','=',e.id),('check_out','>=',to_datetime),('check_out','<=',ends)])
#                 eatt = self.env['hr.attendance'].search([('employee_id.id','=',e.id),('check_in','>=',to_datetime),('check_out','<=',ends)])
#                 emp_abesnt = self.env['hr.attendance'].search([('employee_id.id','=',e.id),('check_in','>=',to_datetime),('check_in','<=',ends)])
                
                for eb in emp_abesnt:
#                     types = eb.reason1.split(' ')
                    res1 = self.half_leave(eb.check_out,e)
#                     if('Half' in types):chek_out_time_todatetime=datetime.strptime(chek_out_time,'%H:%M:%S')
                    if(res1):
                        vals = {
                                'sr':sr,
                                'empcode': e.employe_code,
                                'name': e.name,
                                'status': res1 or 'Absent' ,#eatt.check_in+' '+eatt.check_out or 'Absent',
#                                     'status': types[3] or 'Absent',
                                
                            }
                        sr =sr+1
                        lines.append(vals)
                        
        elif(data['form']['report_daily'] == 'delr'):
            employee = emp_outscope#self.env['hr.employee'].search([])
            stringdate_to_date = datetime.strptime(data['form']['date_daily'], '%Y-%m-%d')
            to_datetime = datetime.strftime(stringdate_to_date, DEFAULT_SERVER_DATETIME_FORMAT)
            end_day_time = stringdate_to_date + timedelta(hours=24)
            ends = datetime.strftime(end_day_time, DEFAULT_SERVER_DATETIME_FORMAT)
            for e in employee:
#                 emp_abesnt = self.env['leave.form'].search([('name.id','=',e.id),('date_from','>=',data['form']['date_daily']),('date_to','<=',data['form']['date_daily'])])
                eatt = self.env['hr.attendance'].search([('employee_id.id','=',e.id),('check_in','>=',to_datetime),('check_out','<=',ends)])
                for a in eatt:
                    res1 = self.early_left(a.check_out,e)
#                     if('Short' in types):
                    if(res1):
                        #code
                        i=datetime.strptime(res1, "%Y-%m-%d %H:%M:%S") 
                        chek_out_time=str(i.time())
                        chek_out_time_todatetime=datetime.strptime(chek_out_time,'%H:%M:%S')
                        differ=dt_work_to-chek_out_time_todatetime
                        early_left_time=str(timedelta(seconds=differ.seconds))        
                            #code 
                        vals = {
                                'sr':sr,
                                'empcode': e.employe_code,
                                'name': e.name,
                                'status': res1 or 'Absent',
                                'earlyleft':early_left_time
#                                 'status': types[3] or 'Absent',
                                
                            }
                        sr =sr+1
                        lines.append(vals)
        elif(data['form']['report_daily'] == 'dltr'):
            employee = emp_outscope#self.env['hr.employee'].search([])
            stringdate_to_date = datetime.strptime(data['form']['date_daily'], '%Y-%m-%d')
            to_datetime = datetime.strftime(stringdate_to_date, DEFAULT_SERVER_DATETIME_FORMAT)
#             next_date =  ((datetime.strptime(stringdate_to_date, DEFAULT_SERVER_DATETIME_FORMAT)))+ timedelta(day)
            end_day_time = stringdate_to_date + timedelta(hours=24)
            ends = datetime.strftime(end_day_time, DEFAULT_SERVER_DATETIME_FORMAT)
#             dic = DefaultListOrderedDict()
            for e in employee:
                print(e,e.name)
#                 print(e.name)
                holidays = self.env['hr.holidays'].search([('employee_id.id','=',e.id),('type','=','remove'),('date_from','>=',to_datetime),('date_from','<=',ends)])#self.env['hr.holidays'].search([('employee_id.id','=',e.id),('type','=','remove')])
                for h in holidays:
                    if(h.holiday_status_id):
                        vals = {
                                'sr':sr,
                                'empcode': e.employe_code,
                                'name': e.name,
                                'holiday_name':h.holiday_status_id.name,
                                'holiday_id':h.holiday_status_id.id,
                                
                                
                            }
                        sr =sr+1
                        lines.append(vals)
                
        
        for item in lines:
            if('rep_type' in item):
                res.setdefault(item['emp_id'], []).append(item)  
            elif('holiday_id' in item):
                res.setdefault(item['holiday_id'], []).append(item)
            
            elif('emp_id' in item):
                res.setdefault(item['emp_id'], []).append(item)        
        pass            

        if(res):
            return res
        
        elif(lines):
            return lines

    def generate_xlsx_report(self, workbook, data, lines):
        get_warehouse = self.get_warehouse(data)
        count = len(get_warehouse[0]) * 11 + 6
        sheet = workbook.add_worksheet('Stock Info')
        format1 = workbook.add_format({'font_size': 14, 'bottom': True, 'right': True, 'left': True, 'top': True, 'align': 'vcenter', 'bold': True})
        format11 = workbook.add_format({'font_size': 12, 'align': 'center', 'right': True, 'left': True, 'bottom': True, 'top': True, 'bold': True})
        format21 = workbook.add_format({'font_size': 10, 'align': 'center', 'right': True, 'left': True,'bottom': True, 'top': True, 'bold': True})
        format3 = workbook.add_format({'bottom': True, 'top': True, 'font_size': 12})
        font_size_8 = workbook.add_format({'bottom': True, 'top': True, 'right': True, 'left': True, 'font_size': 8})
        format81 = workbook.add_format({'font_size': 8, 'align': 'center', 'right': True, 'left': True,'bottom': True, 'top': True, 'bold': True})
        red_mark = workbook.add_format({'bottom': True, 'top': True, 'right': True, 'left': True, 'font_size': 8,
                                        'bg_color': 'red'})
        font_size_9 = workbook.add_format({'bottom': True, 'top': True, 'right': True, 'left': True, 'font_size': 8})
        format9 = workbook.add_format({'bottom': True, 'top': True, 'right': True, 'left': True, 'font_size': 8, 'bold': True})
        justify = workbook.add_format({'bottom': True, 'top': True, 'right': True, 'left': True, 'font_size': 12})
        format3.set_align('center')
        font_size_8.set_align('center')
        font_size_9.set_align('left')
        justify.set_align('justify')
        format1.set_align('center')
        red_mark.set_align('center')
        format9.set_align('center')
        
#         sheet.merge_range('A3:G3', 'Report Date: ' + str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M %p")), format1)
        
#         sheet.merge_range(2, 7, 2, count, 'Warehouses', format1)
        if(data['form']['date_daily'] is not False):
            if((data['form']['report_daily'] != 'dhlr') and (data['form']['report_daily'] != 'ar')):
                sheet.merge_range('A3:J3', get_warehouse, format1)
                sheet.merge_range('A4:B4', 'Date', format21)
                sheet.merge_range('C4:D4', datetime.strptime(str(data['form']['date_daily']),"%Y-%m-%d").strftime("%d-%m-%Y"), format21)
                
            elif(data['form']['report_daily'] == 'dhlr'):
            
                sheet.merge_range('A3:G3', get_warehouse, format1)
                sheet.merge_range('A4:B4', 'Date', format21)
                #sheet.merge_range('C4:D4', data['form']['date_daily'], format21)
                sheet.merge_range('C4:D4', datetime.strptime(str(data['form']['date_daily']),"%Y-%m-%d").strftime("%d-%m-%Y"), format21)
                
                
            elif(data['form']['report_daily'] == 'ar'): 
                sheet.merge_range('A3:G3', get_warehouse, format1)  
                sheet.merge_range('A4:B4', 'Date', format21)
                #sheet.merge_range('C4:D4', data['form']['date_daily'], format21)
                sheet.merge_range('C4:D4', datetime.strptime(str(data['form']['date_daily']),"%Y-%m-%d").strftime("%d-%m-%Y"), format21)
                
                
#             else:
#                 if(data['form']['report_daily'] != 'ar'):
#                     sheet.merge_range('A3:H3', get_warehouse, format1)
#                     sheet.merge_range('A4:B4', 'Date', format21)
#                 #sheet.merge_range('C4:D4', data['form']['date_daily'], format21)
#                     sheet.merge_range('C4:D4', datetime.strptime(str(data['form']['date_daily']),"%Y-%m-%d").strftime("%d-%m-%Y"), format21)
#          
                    
                    
        elif(data['form']['date_from'] is not False):
            #code
            if(data['form']['report_month'] != 'mhlr') and (data['form']['report_month'] != 'mr'):# and (data['form']['report_month'] != 'mewr') :
                sheet.merge_range('A3:J3', get_warehouse, format1)
                sheet.merge_range('A4:C4', 'Date From: ', format21)
                sheet.merge_range(3,3,3,4,  datetime.strptime(str(data['form']['date_from']),"%Y-%m-%d").strftime("%d-%m-%Y"), format21)
                sheet.merge_range('F4:H4', 'Date To', format21)
                sheet.merge_range('I4:J4', datetime.strptime(str(data['form']['date_to']),"%Y-%m-%d").strftime("%d-%m-%Y"), format21)
                
                
                
#             elif(data['form']['report_month'] == 'mewr'):
#                 sheet.merge_range('A3:J3', get_warehouse, format1)
#                 sheet.merge_range('A4:B4', 'Date From: ', format21)
#                 sheet.merge_range(3,2,3,4,  datetime.strptime(str(data['form']['date_from']),"%Y-%m-%d").strftime("%d-%m-%Y"), format21)
#                 sheet.merge_range('F4:G4', 'Date To', format21)
#                 sheet.merge_range('H4:J4', datetime.strptime(str(data['form']['date_to']),"%Y-%m-%d").strftime("%d-%m-%Y"), format21)    
#     
            else:
                sheet.merge_range('A3:H3', get_warehouse, format1)
                sheet.merge_range('A4:B4', 'Date From: ', format21)
                sheet.merge_range(3,2,3,3,  datetime.strptime(str(data['form']['date_from']),"%Y-%m-%d").strftime("%d-%m-%Y"), format21)
                sheet.merge_range('E4:F4', 'Date To', format21)
                sheet.merge_range('G4:H4', datetime.strptime(str(data['form']['date_to']),"%Y-%m-%d").strftime("%d-%m-%Y"), format21)    
            #code 
        w_col_no = 6
        w_col_no1 = 7
#         for i in get_warehouse[0]:
#             w_col_no = w_col_no + 11
#             sheet.merge_range(3, w_col_no1, 3, w_col_no, i, format11)
#             w_col_no1 = w_col_no1 + 11
        if((data['form']['report_month'] != 'mewr') and (data['form']['report_month'] != 'mhlr') and( data['form']['report_month'] != 'mltwr') and (data['form']['report_daily'] != 'ar')and (data['form']['report_daily'] != 'dhlr') and(data['form']['report_month'] != 'mr')and(data['form']['report_month'] != 'melr') and(data['form']['report_month'] != 'mar')  and(data['form']['report_month'] != 'mlar')):
            sheet.write(4, 0, 'Sr. #', format21)
            sheet.write(4, 1, 'Emp.Code', format21)
            sheet.merge_range(4, 2, 4, 5, 'Emp. Name', format21)
        #code added if,else    
        elif((data['form']['report_month'] == 'mhlr') or (data['form']['report_month'] == 'mr') or(data['form']['report_month'] == 'melr') or (data['form']['report_month'] == 'mar') or (data['form']['report_month'] == 'mlar')) : 
            sheet.write(5, 0, 'Sr. #', format21)
            sheet.write(5, 1, 'Emp.Code', format21)
            sheet.merge_range(5, 2, 5, 3, 'Emp. Name', format21)
            sheet.merge_range(5, 4, 5, 5, 'Date', format21)
        
        elif(data['form']['report_daily'] == 'ar'):
            sheet.write(5, 0, 'Sr. #', format21)
            sheet.write(5, 1, 'Emp.Code', format21)
            sheet.merge_range(5, 2, 5, 4, 'Emp. Name', format21)
                
        else:
            if(data['form']['report_daily'] == 'dhlr'):
                sheet.write(5, 0, 'Sr. #', format21)
                sheet.write(5, 1, 'Emp.Code', format21)
                sheet.merge_range(5, 2, 5, 4, 'Emp. Name', format21)    
        #code      
        if(data['form']['report_daily'] == 'ar'): #or data['form']['report_month'] == 'mr'):
            sheet.merge_range(5, 5,5,6, 'Status', format21)
            
        elif(data['form']['report_month']== 'mr'):
            sheet.merge_range(5,6,5,7, 'Status' ,format21)
                
        elif( data['form']['report_month'] == 'mltwr'):
            sheet.write(6, 0, 'Sr. #', format21)
            sheet.write(6, 1, 'Emp.Code', format21)
            sheet.merge_range(6, 2, 6, 3, 'Emp. Name', format21)
            sheet.merge_range(6, 4,6,5, 'Date From', format21)
            sheet.merge_range(6, 6,6,7, 'Date To', format21)
            sheet.merge_range(6,8,6,9, 'Count', format21)
         
        elif( data['form']['report_month'] == 'mewr'):
            sheet.write(5, 0, 'Sr. #', format21)
            sheet.write(5, 1, 'Emp.Code', format21)
            sheet.merge_range(5, 2, 5, 3, 'Emp. Name', format21)
            sheet.merge_range(5, 4,5,5, 'Date From', format21)
            sheet.merge_range(5, 6,5,7, 'Date To', format21)
            sheet.merge_range(5,8,5,9, 'Leave Type', format21)    
#         elif(data['form']['report_daily'] in [ 'dar','dlar'] or data['form']['report_month'] in [ 'mar','mlar']):
#             sheet.merge_range(4, 6,4,7, 'Sign In (Time)', format21)
#         elif(data['form']['report_daily'] in [ 'delr'] or data['form']['report_month'] in [ 'melr','mhlr']):
#             sheet.merge_range(4, 6,4,7, 'Sign Out (Time)', format21)

        elif(data['form']['report_daily'] in [ 'dar'] ):#or data['form']['report_month'] in [ 'mar']):
            sheet.merge_range(4, 6,4,7, 'Sign In (Time)', format21)
            sheet.merge_range(4, 8,4,9, 'Sign Out (Time)', format21)
        elif(data['form']['report_month'] in [ 'mar'] ):#or data['form']['report_month'] in [ 'mar']):
            sheet.merge_range(5, 6,5,7, 'Sign In (Time)', format21)
            sheet.merge_range(5, 8,5,9, 'Sign Out (Time)', format21)
            
        elif(data['form']['report_daily'] in [ 'delr'] ):#or data['form']['report_month'] in [ 'melr']):
            sheet.merge_range(4, 6,4,7, 'Sign Out (Time)', format21)
            sheet.merge_range(4, 8,4,9, 'Early left (Time)', format21)
        
        elif(data['form']['report_month'] in [ 'melr'] ):
            sheet.merge_range(5, 6,5,7, 'Sign Out (Time)', format21)
            sheet.merge_range(5, 8,5,9, 'Early left (Time)', format21)    
        #code    
        #elif(data['form']['report_daily'] in ['dhlr'] or data['form']['report_month'] in [ 'mhlr']):
         #   sheet.merge_range(4, 6,4,7, 'Sign Out (Time)', format21)
        elif(data['form']['report_daily'] in [ 'dhlr']):
            sheet.merge_range(5, 5,5,6, 'Sign Out (Time)', format21)  
            
            
        elif(data['form']['report_month'] in [ 'mhlr']):
            sheet.merge_range(5, 6,5,7, 'Sign Out (Time)', format21) 
        #code    
        elif(data['form']['report_daily'] in [ 'dlar']):
            sheet.merge_range(4, 6,4,7, 'Sign In (Time)', format21)
            sheet.merge_range(4, 8,4,9, 'Late (Time)', format21)
        elif(data['form']['report_month'] in [ 'mlar']):
            sheet.merge_range(5, 6,5,7, 'Sign In (Time)', format21)
            sheet.merge_range(5, 8,5,9, 'Late (Time)', format21)    
            
        prod_row = 5
        prod_col = 0
        get_line = self.get_lines(data, get_warehouse)
        if(get_line):
            if(data['form']['report_daily']):
                if(data['form']['report_daily'] not in [ 'dltr']):
                    for each in get_line:
                        if(data['form']['report_daily'] not in [ 'dhlr'] and  data['form']['report_daily'] not in [ 'ar']):
                            sheet.write(prod_row, prod_col, each['sr'], font_size_8)
                            sheet.write(prod_row, prod_col + 1,each['empcode'], font_size_8)
                            sheet.merge_range(prod_row, prod_col + 2, prod_row, prod_col + 5, each['name'], font_size_8)
                            
                        elif(data['form']['report_daily'] in [ 'ar']):
                            sheet.write(prod_row+1, prod_col, each['sr'], font_size_8)
                            sheet.write(prod_row+1, prod_col + 1,each['empcode'], font_size_8)
                            sheet.merge_range(prod_row+1, prod_col + 2, prod_row+1, prod_col + 4, each['name'], font_size_9)   
                            
                        elif(data['form']['report_daily'] in [ 'dhlr']): 
                            sheet.write(prod_row+1, prod_col, each['sr'], font_size_8)
                            sheet.write(prod_row+1, prod_col + 1,each['empcode'], font_size_8)
                            sheet.merge_range(prod_row+1, prod_col + 2, prod_row+1, prod_col + 4, each['name'], font_size_8)      
                        else:
                            sheet.write(prod_row+1, prod_col, each['sr'], font_size_8)
                            sheet.write(prod_row+1, prod_col + 1,each['empcode'], font_size_8)
                            sheet.merge_range(prod_row+1, prod_col + 2, prod_row+1, prod_col + 5, each['name'], font_size_8)    
                        if(data['form']['report_daily'] in [ 'dar','dlar','delr','dhlr'] or data['form']['report_month'] in [ 'mar','mlar','melr','mhlr']):
                            coll = 6
                            coll2 = 7
                            splt = each['status'].split(' ')
                            date_track = None
                            for ir in splt:
                                if(bool(self.TIME_FORMAT.match(ir))):
                                    if(data['form']['report_daily'] not in ['dhlr'] and  data['form']['report_daily'] not in [ 'ar']):
                                        sheet.merge_range(prod_row, prod_col + coll,prod_row,coll2, ir, font_size_8)
                                    
                                    
                                    elif(data['form']['report_daily'] in ['dhlr']):   
                                        sheet.merge_range(prod_row+1, prod_col + coll-1,prod_row+1,coll, ir, font_size_8)
                                    else:
                                        sheet.merge_range(prod_row+1, prod_col + coll,prod_row+1,coll2, ir, font_size_8)

                                    coll = coll + 2
                                    coll2 = coll2 + 2
                                    
                            #print late time        
                            if(data['form']['report_daily'] in ['dlar']):
                                sheet.merge_range(prod_row, prod_col + 8, prod_row, prod_col + 9, each['late'], font_size_8)
                            
                                #print early left time    
                            else:
                                if(data['form']['report_daily'] in ['delr']):
                                    sheet.merge_range(prod_row, prod_col + 8, prod_row, prod_col + 9, each['earlyleft'], font_size_8)
                                              
    #                         sheet.merge_range(prod_row, prod_col + 6,prod_row,7, each['status'].split(' ')[1], font_size_8)
    #                         sheet.merge_range(prod_row, prod_col + 8,prod_row,9, each['status'].split(' ')[3], font_size_8)
                        else: 
                            if(data['form']['report_daily'] in [ 'ar']):
                                sheet.merge_range(prod_row+1, prod_col + 5,prod_row+1,6, each['status'], font_size_8)    
                            else:  
                                sheet.merge_range(prod_row, prod_col + 6,prod_row,7, each['status'], font_size_8)
                        prod_row = prod_row + 1
                else:
                    for e in get_line.iterkeys():
                        count_line = 0
                        sr = 1
                        for each in get_line[e]:
                            if(count_line == 0):
                                sheet.merge_range(prod_row, prod_col,prod_row, prod_col+1, each['holiday_name'], font_size_8)
                                prod_row = prod_row + 1
                                count_line = 1
                            sheet.write(prod_row, prod_col,sr, font_size_8)
                            sheet.write(prod_row, prod_col + 1,each['empcode'], font_size_8)
                            sheet.merge_range(prod_row, prod_col + 2, prod_row, prod_col + 5, each['name'], font_size_8)
        #                     sheet.merge_range(prod_row, prod_col + 6,prod_row,7, each['status'], font_size_8)
                            prod_row = prod_row + 1
                            sr = sr + 1
            elif(data['form']['report_month']):
                if(data['form']['report_month'] not in [ 'mltwr']):
                    sr = 1

                    for e in get_line.iterkeys():
                        count_line = 0
                        #sr = 1
                        for each in get_line[e]:
                            if('rep_type' not in each):
                                if(data['form']['report_month'] == 'mr'):
                                    if(count_line == 0):
                                        #sheet.write(prod_row, prod_col, each['sr'], font_size_8)
                                        sheet.write(prod_row+1, prod_col, sr, font_size_8)
                                        #sheet.write(prod_row, prod_col + 1,each['empcode'], font_size_8)
                                        #sheet.merge_range(prod_row, prod_col + 2, prod_row, prod_col + 5, each['name'], font_size_8)
                                        sheet.write(prod_row+1, prod_col + 1,each['empcode'], font_size_8)
                                        sheet.merge_range(prod_row+1, prod_col + 2, prod_row+1, prod_col + 3, each['name'], font_size_8)

                                        prod_row = prod_row + 1
                                        count_line = 1
                                        sr=sr+1
                                    s=datetime.strptime(each['date'],"%Y-%m-%d").strftime("%d-%m-%Y")    
                                    sheet.merge_range(prod_row+1, prod_col + 4,prod_row+1,5,s, font_size_8)
                                    sheet.merge_range(prod_row+1, prod_col + 6,prod_row+1,7, each['status'], font_size_8)
                                    prod_row = prod_row + 1
                                    
                                else: 
                                    #put data here   
                                    if(count_line == 0):
                                        if((data['form']['report_month'] != 'mhlr') and (data['form']['report_month'] != 'melr') and (data['form']['report_month'] != 'mlar') and (data['form']['report_month'] != 'mar')):
                                            #sheet.write(prod_row, prod_col, each['sr'], font_size_8)
                                            sheet.write(prod_row, prod_col, each['sr'], font_size_8)
                                            sheet.write(prod_row, prod_col + 1,each['empcode'], font_size_8)
                                            sheet.merge_range(prod_row, prod_col + 2, prod_row, prod_col + 5, each['name'], font_size_8)
                                            prod_row = prod_row + 1
                                            
                                            count_line = 1
#                                         elif(data['form']['report_month'] == 'melr'):
#                                             sheet.write(prod_row, prod_col,sr, font_size_8)
#                                             sheet.write(prod_row, prod_col + 1,each['empcode'], font_size_8)
#                                             sheet.merge_range(prod_row, prod_col + 2, prod_row, prod_col + 5, each['name'], font_size_8)
#                                             prod_row = prod_row + 1
#                                             sr=sr+1
#                                             count_line = 1    
                                            
                                        else:    
                                            
                                            sheet.write(prod_row+1, prod_col, sr, font_size_8)
                                            sheet.write(prod_row+1, prod_col + 1,each['empcode'], font_size_8)
                                            sheet.merge_range(prod_row+1, prod_col + 2, prod_row+1, prod_col + 3, each['name'], font_size_8)
                                            prod_row = prod_row + 1
                                            sr=sr+1
                                            count_line = 1
    #                                     datetime.strptime(data['form']['date_from'], '%Y-%m-%d') + timedelta(start)
                                    splt = each['status'].split(' ')
                                    date_track = None
                                    coll = 4
                                    coll2 = 5
                                    for ir in splt:
    #                                     type(time.strptime(ir, '%H:%M:%S')) is (time.struct_time)
    #                                     dd = datetime.strptime(ir, '%Y-%m-%d').strftime('%d-%m-%Y')
    #                                     ds = datetime.strptime(dd, '%d-%m-%Y')
                                        
                                        if(ir != date_track):
                                            #code
                                            if((data['form']['report_month'] == 'mar') or (data['form']['report_month']== 'melr')):
                                                if splt[0]==ir:
                                                    dmy=datetime.strptime(ir,"%Y-%m-%d").strftime("%d-%m-%Y")
                                                    sheet.merge_range(prod_row+1, prod_col + coll,prod_row+1,coll2, dmy, font_size_8)
                                                else:
                                                    sheet.merge_range(prod_row+1, prod_col + coll,prod_row+1,coll2, ir, font_size_8)    
                                            #code
                                            elif(data['form']['report_month'] == 'mhlr') :
                                                #change date format
                                                if splt[0]==ir:
                                                    dmy=datetime.strptime(ir,"%Y-%m-%d").strftime("%d-%m-%Y")
                                                    sheet.merge_range(prod_row+1, prod_col + coll,prod_row+1,coll2, dmy, font_size_8)
                                                else:
                                                    sheet.merge_range(prod_row+1, prod_col + coll,prod_row+1,coll2, ir, font_size_8)
                                                    
                                            elif(data['form']['report_month'] == 'mlar'):
                                                sheet.merge_range(prod_row+1, prod_col + coll,prod_row+1,coll2, ir, font_size_8)
                                     
                                                #sheet.merge_range(prod_row, prod_col + coll,prod_row,coll2, ir, font_size_8)
                                            else:
                                                sheet.merge_range(prod_row, prod_col + coll,prod_row,coll2, ir, font_size_8)
 
                                            if(bool(self.TIME_FORMAT.match(ir)) == False):
                                                date_track = ir
                                            coll = coll + 2
                                            coll2 = coll2 + 2
                                            if(data['form']['report_month'] in [ 'mlar']):
                                                sheet.merge_range(prod_row+1, prod_col + coll, prod_row+1, prod_col + coll2, each['late'], font_size_8)
                                            else:
                                                if(data['form']['report_month'] in [ 'melr']): 
                                                    sheet.merge_range(prod_row+1, prod_col + coll, prod_row+1, prod_col + coll2, each['earlyleft'], font_size_8)   
    #                                     sheet.merge_range(prod_row, prod_col + 6,prod_row,7, each['status'].split(' ')[1], font_size_8)
    #                                 
    #                                     sheet.merge_range(prod_row, prod_col + 8,prod_row,9, each['status'].split(' ')[3], font_size_8)
                                    prod_row = prod_row + 1
                                   
                            else:
                                #======================put data here===============================
                                if(count_line == 0):
                                    sheet.write(prod_row+1, prod_col,sr, format81)
                                    sheet.write(prod_row+1, prod_col+1, each['empcode'], format81)
                                    sheet.merge_range(prod_row+1, prod_col + 2, prod_row+1, prod_col + 3, each['name'], format81)
                                    
                                    prod_row = prod_row + 1
                                    count_line = 1
                                    sr=sr+1
                                sheet.merge_range(prod_row+1, prod_col + 4, prod_row+1, prod_col + 5, each['date_from'], font_size_8)
                                sheet.merge_range(prod_row+1, prod_col + 6, prod_row+1, prod_col + 7, each['date_to'], font_size_8)    
    
                                #sheet.write(prod_row, prod_col, each['sr'], font_size_8)
                                #sheet.write(prod_row, prod_col + 1,each['empcode'], font_size_8)
                                #sheet.merge_range(prod_row, prod_col + 2, prod_row, prod_col + 5, each['name'], font_size_8)
                                sheet.merge_range(prod_row+1, prod_col + 8,prod_row+1,9, each['holiday_name'], font_size_8)
                                prod_row = prod_row + 1
                                
                else:
                    for e in get_line.iterkeys():
                        count_line = 0
                        sr = 1
                        for each in get_line[e]:
                            if(count_line == 0):
                                sheet.merge_range(prod_row, prod_col,prod_row, prod_col+1, each['holiday_name'], format9)
                                #test
                                sheet.write(prod_row+1, prod_col , 'Sr. #', format21)
                                sheet.write(prod_row+1, prod_col+1 , 'Emp.Code', format21)
                                 
                                sheet.merge_range(prod_row+1, prod_col+2, prod_row+1, prod_col+3, 'Emp. Name', format21)
                                sheet.merge_range(prod_row+1, prod_col+4,prod_row+1,prod_col+5, 'Date From', format21)
                                sheet.merge_range(prod_row+1, prod_col+6,prod_row+1,prod_col+7, 'Date To', format21)
                                sheet.merge_range(prod_row+1,prod_col+8,prod_row+1,prod_col+9, 'Count', format21)
                                #test
                                #prod_row = prod_row + 1
                                prod_row = prod_row + 2
                                count_line = 1
                            sheet.write(prod_row, prod_col,sr, font_size_8)
                            sheet.write(prod_row, prod_col + 1,each['empcode'], font_size_8)
                            sheet.merge_range(prod_row, prod_col + 2, prod_row, prod_col + 3, each['name'], font_size_8)
                            sheet.merge_range(prod_row, prod_col + 4, prod_row, prod_col + 5, datetime.strptime( each['datefrom'],"%Y-%m-%d").strftime("%d-%m-%Y")
                                              , font_size_8)
                            sheet.merge_range(prod_row, prod_col + 6, prod_row, prod_col + 7,  datetime.strptime( each['dateto'],"%Y-%m-%d").strftime("%d-%m-%Y")
                                             , font_size_8)
                            sheet.merge_range(prod_row, prod_col + 8, prod_row, prod_col + 9, each['count'], font_size_8)
    #                         sheet.merge_range(prod_row, prod_col + 6,prod_row,7, each['status'], font_size_8)
                            prod_row = prod_row + 1
                            sr = sr + 1
        else:
            raise exceptions.ValidationError(_("No Data was Found "))
#         for i in get_warehouse[1]:
#             get_line = self.get_lines(data, i)
#             for each in get_line:
#                 sheet.write(prod_row, prod_col, each['sku'], font_size_8)
#                 sheet.merge_range(prod_row, prod_col + 1, prod_row, prod_col + 3, each['name'], font_size_8)
#                 sheet.merge_range(prod_row, prod_col + 4, prod_row, prod_col + 5, each['category'], font_size_8)
#                 sheet.write(prod_row, prod_col + 6, each['cost_price'], font_size_8)
#                 prod_row = prod_row + 1
#             break
#         prod_row = 5
#         prod_col = 7
#         for i in get_warehouse[1]:
#             get_line = self.get_lines(data, i)
#             for each in get_line:
#                 if each['available'] < 0:
#                     sheet.write(prod_row, prod_col, each['available'], red_mark)
#                 else:
#                     sheet.write(prod_row, prod_col, each['available'], font_size_8)
#                 if each['virtual'] < 0:
#                     sheet.write(prod_row, prod_col + 1, each['virtual'], red_mark)
#                 else:
#                     sheet.write(prod_row, prod_col + 1, each['virtual'], font_size_8)
#                 if each['incoming'] < 0:
#                     sheet.write(prod_row, prod_col + 2, each['incoming'], red_mark)
#                 else:
#                     sheet.write(prod_row, prod_col + 2, each['incoming'], font_size_8)
#                 if each['outgoing'] < 0:
#                     sheet.write(prod_row, prod_col + 3, each['outgoing'], red_mark)
#                 else:
#                     sheet.write(prod_row, prod_col + 3, each['outgoing'], font_size_8)
#                 if each['net_on_hand'] < 0:
#                     sheet.merge_range(prod_row, prod_col + 4, prod_row, prod_col + 5, each['net_on_hand'], red_mark)
#                 else:
#                     sheet.merge_range(prod_row, prod_col + 4, prod_row, prod_col + 5, each['net_on_hand'], font_size_8)
#                 if each['sale_value'] < 0:
#                     sheet.merge_range(prod_row, prod_col + 6, prod_row, prod_col + 7, each['sale_value'], red_mark)
#                 else:
#                     sheet.merge_range(prod_row, prod_col + 6, prod_row, prod_col + 7, each['sale_value'], font_size_8)
#                 if each['purchase_value'] < 0:
#                     sheet.merge_range(prod_row, prod_col + 8, prod_row, prod_col + 9, each['purchase_value'], red_mark)
#                 else:
#                     sheet.merge_range(prod_row, prod_col + 8, prod_row, prod_col + 9, each['purchase_value'], font_size_8)
#                 if each['total_value'] < 0:
#                     sheet.write(prod_row, prod_col + 10, each['total_value'], red_mark)
#                 else:
#                     sheet.write(prod_row, prod_col + 10, each['total_value'], font_size_8)
#                 prod_row = prod_row + 1
#             prod_row = 5
#             prod_col = prod_col + 11

StockReportXls('report.attendance_xls.attendance_report_xls.xlsx', 'hr.attendance')
