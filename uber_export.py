import os
import re
import json
import string
import itertools				
from collections import OrderedDict
import csv
import xlwt
from xlwt import Workbook

class Uber_report:
    """Read uber weekly csv, and creates
    excel file as named Compiled.xls output"""
    
    def __init__(
        self,
        save_dir="."
        ):
        self.save_dir = save_dir
        self.count = 0
        self.fieldnames = set()
        
    @staticmethod
    def count_file(dir_d):
        """Counts files in given destination
        if destination doesnt exist, creates directory,
        returns int(file_count) and list(file_list) """
        if os.path.exists(dir_d) == False:
            os.makedirs(dir_d)
        file_list = []
        file_count = len(next(os.walk(dir_d))[2])
        for root, dirs, files in os.walk(dir_d):
            file_list.append(files)
        return file_count, file_list[0]

    
    def read_file(self, dir_c, fname):
        """Reads given file from given destination
        if destination doesnt exist, creates directory"""
        result_dict = {}
        if os.path.exists(dir_c) == False:
            os.makedirs(dir_c)
        with open(os.path.join(dir_c, fname), 'r') as filer:
            output = csv.DictReader(filer)
            fieldnames = output.fieldnames
            for row in output:
                row_id = 'id%s' % self.count
                result_dict.update({row_id: {}})
                for field in fieldnames:
                    self.fieldnames.add(field)
                    result_dict[row_id][field] = row[field]
                self.count += 1
        return result_dict
    
    @property
    def report_data(self):
        file_cnt, files = self.count_file(self.save_dir)
        rep_data = {}
        for filen in files:
            if filen.startswith('statement'):
                fdata = self.read_file(self.save_dir, filen)
                rep_data.update(dict(fdata))
        return rep_data
    
    def write_xls_report(self):
        data = self.report_data
        wb = Workbook()
        style = xlwt.easyxf('font: bold 1, color red;') 
        unwanted_columns = ['Email', 'Phone number', 'Type']
        # add_sheet is used to create sheet. 
        sheet1 = wb.add_sheet('All Trips')
        iterate = 0
        for nr, field in enumerate(sorted(self.fieldnames)):
            if field not in unwanted_columns:
                sheet1.write(0, iterate, field, style)
                iterate += 1
                column_values = self.fill_column(data, field)
                max_width = 12
                for y, value in enumerate(column_values):
                    first_col = sheet1.col(iterate - 1)
                    if len(value) > max_width:
                        max_width = len(value)
                    first_col.width = 256 * max_width
                    sheet1.write(y + 1, iterate - 1, value)
        file_nr = self.count_xls()
        wb.save('Uber_report_%s.xls' % file_nr)
        
#    def sort_rows(self, data_tuple):
#        trip_date = data_tuple[1].get('Date/Time')
#        return trip_date

    def fill_column(self, data, key):
        res = []
        for row in data:
            res.append(data[row][key])
        return res
    
    def count_xls(self):
        file_number = 0
        count, all_files = self.count_file(self.save_dir)
        for filen in all_files:
            if filen.startswith('Uber_report_'):
                file_number += 1
        return file_number

if __name__ == "__main__":
    report = Uber_report()
    report.write_xls_report()
#    print(json.dumps(report.report_data, indent=4))

