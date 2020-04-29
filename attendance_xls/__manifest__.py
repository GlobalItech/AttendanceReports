
{
    'name': 'HR (Daily/Monthly) Reports in Excel',
    'summary': "Print Attendance Report in Excel from Attendances",
    'author':  'itech resources',
    'company': 'itech resources',
    'license': 'LGPL-3',
    'depends': [
                'base',
                'hr',
                'hr_attendance',
                'report_xlsx'
                ],
    'data': [
            'views/wizard_view.xml',
            'views/customizations.xml',
          
            ],
    'images': ['static/description/banner.gif'], 
    'installable': True,
    'auto_install': False,
    'price':'80.0',
    'currency': 'EUR',
}
