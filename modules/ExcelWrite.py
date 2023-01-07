import io
import sys
from os import path
import xlsxwriter

class ExcelWrite(object):
    def __init__(self):
        self.OutFile = ""
        self.workbook = None
        
    def open(self, OutFile):
        self.OutFile = OutFile
        self.workbook = xlsxwriter.Workbook(OutFile)
        

    def write(self, InTable, Sheetname, SheetHeader):
        
        sheet1 = self.workbook.add_worksheet(Sheetname)
    
        formatter = self.workbook.add_format()
        formatter.set_border(1)
        formatter.font_name = "Calibri"    
    
        title_formatter = self.workbook.add_format()
        title_formatter.set_border(1)
        title_formatter.set_bg_color('#cccccc')
        title_formatter.set_align('center')
        title_formatter.set_bold()
    
        ave_formatter = self.workbook.add_format()
        ave_formatter.set_border(1)
        ave_formatter.set_num_format('0.00')
    
        col_width_dict = dict()  # create a dictionary var
    
        row_number = int(0)
        column_num = int(0)
        
        # write header
        if SheetHeader is not None:
            head = SheetHeader.split(",")
            for itm in head:
                sheet1.write_string(row_number, column_num, itm, title_formatter)
                column_num += 1
            #sheet1.write(row_number, column_num, head, title_formatter)
    
        for i in range(50):  # fill it up with 0s first so Python doesn't complain
            col_width_dict[i] = 0
                                
        # write xlsx record
        row_number = int(1)
        column_num = int(0)
    
        for recitem in InTable:
    
            column_num = int(0)
            
            for itm in recitem:
                if itm is not None:                
                    try:
                        sheet1.write(row_number, column_num, str(itm), formatter)
    
                    except Exception as e:
                        print("cell[" + str(column_num) + "][" + str(row_number) + "] input error reason " + e)
                        print("cell[" + str(column_num) + "][" + str(row_number) + "] value=[" + str(itm) + "]")
                        exit(5)
    
                column_num += 1
    
            row_number += 1
            
    def close(self):
        self.workbook.close()
        print(" Convert [" + self.OutFile + "] completed. Filesize = " + str(path.getsize(self.OutFile)) + " bytes")
