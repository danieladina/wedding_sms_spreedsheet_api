import gspread
from oauth2client.service_account import ServiceAccountCredentials





sheet_body={
    'properties':{
        'title':"החתונה של פלוני ואלמונית",
        'timeZone': 'Jerusalem',
        'autoRecalc':'ON_CHANGE'
    },
    'sheets':[
        {
            'properties':{
                'title':'sheet1',
                "rightToLeft": True
            }
        },




    ]
}





