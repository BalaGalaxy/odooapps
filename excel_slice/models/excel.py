from odoo import models, fields, api, _
from odoo.exceptions import ValidationError
import tempfile
import base64
import pandas as pd
import zipfile
import os, shutil

class HospitalExcel(models.Model):
    _name = 'excel.slice'

    select_file = fields.Binary(string='Selected File', attachment=True)
    file_name = fields.Char("File Name")
    split_into = fields.Integer(string='Split into')



    @api.depends('split_into')
    def split_excel(self):
        file_path = tempfile.gettempdir() +'/select_file.xlsx'        
        f = open(file_path, 'wb+')
        f.write(base64.decodebytes(self.select_file))
        f.close()
        i = 0
        df = pd.read_excel(file_path)
        size = self.split_into
        list_of_dfs = [df.loc[i:i + size - 1, :] for i in range(0, len(df), size)]

        file_path  = os.path.join(os.path.dirname(os.path.abspath(__file__)))                    
        new_path = file_path.replace('models','static')


        i = 1
        for df in list_of_dfs:
            writer = pd.ExcelWriter(new_path+'/templates/doc/output_' + str(i) + '.xlsx')
            df.to_excel(writer, index=False)
            # save the excel
            writer.save()
            print('DataFrame is written successfully to Excel File.')
            i = i + 1
            handle = zipfile.ZipFile(new_path+'/all_xlsx.zip', 'w')
            os.chdir(new_path+'/templates/doc')
            for x in os.listdir():
                if x.endswith('.xlsx'):
                    handle.write(x, compress_type=zipfile.ZIP_DEFLATED)
            handle.close()
        target = new_path+'/templates/doc/'
        for x in os.listdir(target):
            if x.endswith('.xlsx'):
                os.unlink(target + x)
        return {
            'type': 'ir.actions.act_url',
            'url': str('/excel_slice/static/all_xlsx.zip'),
            'target': 'new',
        }
        
    @api.constrains('select_file')
    def _check_file(self):
        if str(self.file_name.split(".")[1]) != 'xlsx':
            raise ValidationError("Cannot upload file different from .xlsx file")
