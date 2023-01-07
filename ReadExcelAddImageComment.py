###########################################################################################################################
#
#	Programmed by Jame Leung
#	Date: 2023/01/01
#
#
#
# 1. This program will only search for the numeric worksheet name (i.e. 0,1,2, and so on) and extract the data below into the image
# 2. Field to be extracted = Item Number, brand, description, finsih (the column A,B,C,D)
# 3. The image to be placed is refer to the column AY. So need to ensure the filename is the same as the AY
# 4. currently the image source supports JPG only
#

from PIL import Image
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont
import numpy as np
from modules.Configctrl import Configctrl
from modules.timectrl import timectrl
import xlrd
import os
import textwrap
import pprint
import re

def drawrect(drawcontext, xy, outline=None, width=0):
    (x1, y1), (x2, y2) = xy
    points = (x1, y1), (x2, y1), (x2, y2), (x1, y2), (x1, y1)
    drawcontext.line(points, fill=outline, width=width)
    
class GenImage:

    def __init__(self, image_path, ItmName, Brand,Description, Finish,ItemCode,ProjectName,Sales,image_outpath,image_outfilename,x_coord,y_coord,DrawFrame,FrameWidth):
        self.image_path=image_path
        self.ItmName = ItmName
        self.Brand = Brand
        self.Description= Description
        self.Finish=Finish
        self.ItemCode=ItemCode
        self.ProjectName=ProjectName
        self.Sales=Sales
        self.image_outpath=image_outpath
        self.image_outfilename=image_outfilename
        #orig=160,390
        self.x_coord=x_coord
        self.y_coord=y_coord
        self.DrawFrame=DrawFrame
        self.FrameWidth=FrameWidth
            
    def createimage(self):
        
        try:
            image1 = Image.open(self.image_path  + "\\" + self.ItemCode + ".jpg")

        except Exception as e:
            print(e)

        myFont = ImageFont.truetype('C:/Windows/Fonts/arial.ttf', 12)
        myFont1 = ImageFont.truetype('C:/Windows/Fonts/arialbd.ttf', 14)

        I1 = ImageDraw.Draw(image1)
        I1.text((self.x_coord, self.y_coord),self.ItmName, font=myFont, fill=(0,0,0))
        
        I1.text((self.x_coord + 30, self.y_coord), self.Brand, font=myFont, fill=(0,0,0))

        aa=[]
        if len(self.Description)<=45: 
            aa.append(self.Description)
        else: 
            aa.extend(textwrap.wrap(self.Description,45))
        
        NoOfLine=len(aa)
        
        lastline=self.y_coord

        I1.text(( self.x_coord + 90, self.y_coord), aa[0], font=myFont, fill=(0,0,0))

        if NoOfLine >1 :
            I1.text((self.x_coord + 90, self.y_coord + 20), aa[1], font=myFont, fill=(0,0,0))
            lastline=self.y_coord + 20
        
        if NoOfLine> 2:
            I1.text((self.x_coord + 90, self.y_coord + 40), aa[2], font=myFont, fill=(0,0,0))
            lastline=self.y_coord + 40

        if NoOfLine>3:
            I1.text((self.x_coord + 90, self.y_coord + 60), aa[3], font=myFont, fill=(0,0,0))
            lastline=self.y_coord + 60

        if NoOfLine>4:
            I1.text((self.x_coord + 90, self.y_coord + 80), aa[4], font=myFont, fill=(0,0,0))
            lastline=self.y_coord + 80

        I1.text((self.x_coord + 420, self.y_coord), self.Finish, font=myFont, fill=(0,0,0))

        #project name and sales person name
        I1.text((self.x_coord, lastline + 40), self.ProjectName, font=myFont1, fill=(0,0,0))
        I1.text((self.x_coord, lastline + 60), self.Sales, font=myFont1, fill=(0,0,0))

        ok_wrd="YyTt"
        
        if (self.DrawFrame[0] in ok_wrd):
            drawrect(I1, [(self.x_coord-6, self.y_coord-6), (self.x_coord + 420 + len(self.Finish)* 6 + 10, lastline + 25)], outline="black", width=self.FrameWidth)

        image1.save(self.image_outpath  + "\Result-" + self.image_outfilename +".jpg", quality=100, subsampling=0)

        

cfg = Configctrl()
cfginfo = cfg.LoadCfg("readimage.ini")
print("#########################################################################\n#")
print("# This program is written by Jame Leung\n# from Idemia France\n# Date:3rd Jan 2023\n#\n# It is to extract the excel data into image\n#\n#")
print("# 1. This program will only search for the numeric worksheet name (i.e. 0,1,2, and so on)")
print("# 2. Field to be extracted = Item Number, brand, description, finsih (the column A,B,C,D)")
print("# 3. The image to be placed is refer to the column AY. So need to ensure the filename is the same as the AY")
print("# 4. currently the image source supports JPG only\n#")
print("#########################################################################\n\n")
#debug mesg ctrl
ok_wrd="YyTt"
ShowDebugMsg=False

if (cfginfo["General"]["Debug"][0] in ok_wrd):
    ShowDebugMsg=True

if ShowDebugMsg!=False:
    print("[" + timectrl.getTimeStamp() + "] Read Config file readimage.ini")

# Change the current working directory
os.chdir(cfginfo["General"]["RunPath"])
if ShowDebugMsg!=False:
    print("Current working directory: {0}".format(os.getcwd()) + "\n")
    
print("[" + timectrl.getTimeStamp() + "] Program start")

#read excel file
if ShowDebugMsg!=False:
    print("[" + timectrl.getTimeStamp() + "]Open Excel File=" + cfginfo["General"]["RunPath"] + "\\" + cfginfo["General"]["ExcelFileName"])

 # Define variable to load the dataframe
wb = xlrd.open_workbook( cfginfo["General"]["RunPath"] + "\\" + cfginfo["General"]["ExcelFileName"])

if ShowDebugMsg!=False:
    print("Total " + str(len(wb.sheet_names())) + " worksheets in the excel file. Start Extract data...\n")

StartCap= cfginfo["General"]["KeywordStartCapture"]
StopCap= cfginfo["General"]["KeywordStopCapture"]

imageprocessed=0
for n in range(len(wb.sheet_names())):
    ws = wb.sheets()[n]
    if wb.sheet_names()[n].isdigit():
        if ShowDebugMsg!=False:
            print("### Process Worksheet [" + wb.sheet_names()[n] + "]")

        data6=""
        data7=""

        #  Get Project Name

        for i in range(ws.nrows):
            if str(ws.cell_value(rowx=i, colx=0))[0:7]=="PROJECT": 
                data6= ws.cell_value(rowx=i, colx=0) + ws.cell_value(rowx=i, colx=1) + ws.cell_value(rowx=i, colx=2)
                break

        for i in range(ws.nrows):
            if str(ws.cell_value(rowx=i, colx=0))[0:4]=="Sale": 
                data7= ws.cell_value(rowx=i, colx=0)
                break

        # Iterate the loop to read the cell values
        StartCapture=False
    
        for i in range(ws.nrows):

            if str(ws.cell_value(rowx=i, colx=0))[0:4]==StartCap:
                StartCapture=True
                if ShowDebugMsg!=False:
                    print("[" + timectrl.getTimeStamp() + "] Row number=" + str(i) +" found keyword [" + StartCap + "], Start Capture Item Data.")
                    
                continue
            if str(ws.cell_value(rowx=i, colx=0))[0:4]==StopCap:
                StartCapture=False
                if ShowDebugMsg!=False:
                    print("[" + timectrl.getTimeStamp() + "] Row number=" + str(i) +" found keyword [" + StopCap + "], Stop Capture Item Data.")
                break
            if StartCapture==True:
                if len(ws.cell_value(rowx=i, colx=0))>0:
                    data1=ws.cell_value(rowx=i, colx=0)
                    data2=ws.cell_value(rowx=i, colx=1)
                    data3=ws.cell_value(rowx=i, colx=2)
                    data3=re.sub(' +', ' ', data3)
                    data3 =repr(data3)
                    data3 = data3.replace('\n', ' ')
                    data4=ws.cell_value(rowx=i, colx=3)
                    data5=str(ws.cell_value(rowx=i, colx=50))
                    if data5.find(".")!=-1:
                        #get rid of the float digit
                        data5=data5[:data5.index('.')]
                    data8=wb.sheet_names()[n] + "-" + data5

                    if ShowDebugMsg!=False:
                        print("data1=" + data1) 
                        print("data2=" + data2)
                        print("data3=" + data3)
                        print("data4=" + data4)
                        print("data5=" + data5)
                        print("data6=" + data6)
                        print("data7=" + data7)
                        print("data8=" + data8)

                    a = GenImage(image_path=cfginfo["General"]["ImageSourcePath"], ItmName=data1,Brand=data2, Description=data3, 
                    Finish=data4, ItemCode=data5,ProjectName=data6, Sales=data7,
                    image_outpath=cfginfo["General"]["ImageResultPath"],image_outfilename=data8,
                    x_coord=int(cfginfo["General"]["Left"]),
                    y_coord=int(cfginfo["General"]["Top"]),
                    DrawFrame=cfginfo["General"]["DrawFrame"],
                    FrameWidth=int(cfginfo["General"]["FrameWidth"]))

                    a.createimage()
                    imageprocessed+=1

print("[" + timectrl.getTimeStamp() + "] Process Complete. " + str(imageprocessed)+" Images Processed." )

