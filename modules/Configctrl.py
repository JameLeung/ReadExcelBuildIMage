import platform
import sys
from os import path,getcwd
import csv
import configparser 
import io

class Configctrl():
    def __init__(self):
        self.configfile = []
        self.__pathstroke = ""

    def GetPathStroke(self):
        if len(self.__pathstroke) == 0:            
            if getattr(sys, "getwindowsversion", None) is not None:
                self.__pathstroke  = "\\"
            else:
                self.__pathstroke = "/"
            
        return self.__pathstroke

    def LoadCfg(self, arg):  
            
        configfilepath = getcwd() + self.GetPathStroke() + arg       
        if path.exists(configfilepath) is not True:
            print ("Config file not exist [" + configfilepath + "]")
            return 0
        
        fileext= arg.lower()
        
        #cfg file, INI format
        config = configparser.ConfigParser()
        
        try:
            config.read(configfilepath, encoding='utf8')
            #config.read(configfilepath, encoding='big5')
        except Exception as e:
            print ("Unable to open [" + configfilepath + "]")
            print (e.message)
            
            return 0
        
        config.sections()

        #print ("CFG file [" + configfilepath + "] loaded")
            
        return config
        
    def LoadProxy(self, arg):  
            
        configfilepath = getcwd() + self.GetPathStroke() + arg
        if path.exists(configfilepath) is not True:
            print ("Config file not exist [" + configfilepath + "]")
            return 0
        
        fileext= arg.lower()
        
        #proxy file, csv format
        with open(configfilepath) as csvfile:
            readCSV = csv.reader(csvfile, delimiter=',')
            ipaddrs= []
            for row in readCSV:
                ipaddr = row[2] + "://" + row[0] + ":" + row[1]
                ipaddrs.append([row[2],ipaddr,row[3],row[4],row[5],row[6],row[7]])
                        
        print ("Proxy file [" + configfilepath + "] loaded")                    

        
        return ipaddrs

    def LoadDate(self, arg):  
            
        datefilepath = getcwd() + self.GetPathStroke() + arg
        
        if path.exists(datefilepath) is not True:
            print ("Date file not exist [" + configfilepath + "]")
            return 0
        
        fileext= arg.lower()
        
        #txt file      
        dateval=[]
        with open(datefilepath, "r") as f:
            while True:
                row = f.readline()
                if not row:break
                dateval.append(row.strip('\n'))
        print ("Date file [" + datefilepath + "] loaded")                    
        
        return dateval
