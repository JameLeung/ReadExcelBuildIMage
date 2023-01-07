# ReadExcelBuildIMage
To build the excel file and build the content into the image file inside

	Programmed by Jame Leung
	Date: 2023/01/01

	from Idemia Hong Kong

Coding implemented = python3

Platform to used = Any platform (example is used the Windows platform)

	1. This program will only search for the numeric worksheet name (i.e. 0,1,2, and so on) and extract the data below into the image
	2. Field to be extracted = Item Number, brand, description, finsih (the column A,B,C,D)
	3. The image to be placed is refer to the column AY. So need to ensure the filename is the same as the AY
	4. currently the image source supports JPG only

You may refer to the excel file "quotation2.xls" to understand how to work

How to make it work
----------------------------------------------------------

a. All coding is done using Python. Some common library is added for easier debugging and configure using INI.
b. It require to download the library below before access.

1. you may need to install those library before proceed

pip install pillow

pip install numpy

pip install xlrd

pip install xlsxwriter

2. open command prompt (command:cmd) and create directory 

md c:\temp1
md c:\temp1\ImageSource
md c:\temp1\Result

3. Copy all the images from sample directory from this project folder "\sample"
into c:\temp\ImageSource


4. In c:\temp1, run the program "python ReadExcelBuildImage"

Any other simple way to run?
----------------------------------------------------------

Alterantively, if you do not want to use Python to run. run the executable I compiled using auto-py-to-exe

1. Copy the project folder ReadExcelBuildIMage.exe into c:\temp1
2. Copy the ini file into c:\temp as well
3. Then run the command c:\temp1\ReadExcelBuildIMage.exe

----------------------------------------------------------

INI explain 

[General]

#Run path = This is the directory path putting the excel file inside

RunPath=c:\temp1

#Excel file name. It supports XLSX and XLS format only

ExcelFileName=quotation2.xls

#The soure image file directory. Please ensure filename has to match with the column AY of the quotation. Case sensitive

ImageSourcePath=c:\temp1\ImageSource

#The output directory to get the image

ImageResultPath=c:\temp1\Result

#to control if the descrption has to draw frame (May choose True,False,Yes,No to control, case insensitive)

DrawFrame=yes

#the thickness of the frame

FrameWidth=1

#X location of the frame to be printed
#(Default: Left=160)

Left=160

#Y location of the frame to be printed
#(Default: Top=800)

Top=800

#to contol in the worksheet when the data has to be captured

KeywordStartCapture=ITEM

#to contol in the worksheet when the data stopped to capture 

KeywordStopCapture=Term

#to control if the debug messages has to be displayed (May choose True,False,Yes,No to control, case insensitive)

Debug=yes

----------------------------------------------------------
