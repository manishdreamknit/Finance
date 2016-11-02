#---------------------------------------------------------------
# This function file contains all the common functions required for testing DT web services
#---------------------------------------------------------------
# Adding required libraries
import urllib2
import base64
import string, os, sys
import collections
from xlrd import open_workbook,cellname,colname,cellnameabs
from xlwt import Workbook, easyxf
from xlutils.copy import copy
from xml.dom.minidom import parse, parseString
from os import path
import xml.parsers.expat
import subprocess, signal
import re
import glob
import csv, xlwt
from subprocess import Popen
import	shutil
import	time
import  datetime
from time import gmtime, strftime
import urllib
import random

import json
import requests
#---------------------------------------------------------------
# Global class definition
#---------------------------------------------------------------
# Class for handling exceptions




class ErrorOccured(Exception):
    # Any error occured
    def __init__(self, value):      # built in method to initialize
        self.value = value
    def __str__(self):              # built in method. This will print the message passed.
        return self.value


#----------------------------------------------------------
# Description:      Find the length of given list
# Input Parameters: Directory Path
# Return Values:    Array of files
# Author:           Naveed Hasan
#---------------------------------------------------------- 
def GetLengthOfList(listArray):
    length = len(listArray)
    return    length

#----------------------------------------------------------
# Description:      Return Current Time
# Author:           Naveed Hasan
#---------------------------------------------------------- 
def GetCurrentTime():
    current_time = datetime.datetime.now().time()
    #current_time = str(datetime.now())					#strftime("%Y-%m-%d %H:%M:%S", gmtime())
    current_time = str(current_time)
    return current_time
#----------------------------------------------------------
# Description:      Find the names of file in a given directory
# Input Parameters: Directory Path
# Return Values:    Array of files
# Author:           Naveed Hasan
#---------------------------------------------------------- 
def RetainFileNamesInDir(DirRef,Extention):
    #listFiles = glob.glob(DirRef + "/*.txt")
    listFiles = glob.glob(DirRef + "/*." + Extention)
    return    listFiles
	#glob.glob("/home/adam/*.txt

def DeleteFilesInDir(DirRef,Extention):
    listFiles = glob.glob(DirRef + "/*." + Extention)
    for item in listFiles:
           if path.isfile(item):
			   try:
				  os.remove(item)
			   except OSError:
				  pass
    listFiles = glob.glob(DirRef + "/*./*")
    return listFiles
#----------------------------------------------------------
# Description:      Find Starting Position Of Mutiple Nodes\
# Input Parameters: N/A
# Return Values:    Array of positions
# Author:           Naveed Hasan
#---------------------------------------------------------- 
def GetNodePositions(StringToSearchIn,StringToSearchFor):
    starts = [match.start() for match in re.finditer(re.escape(StringToSearchFor), StringToSearchIn)]
    return    starts

#----------------------------------------------------------
# Description:      Function to retain a Sub string based Start & End Pos
#					From a given string.
# Input Parameters: String to trim
# Return Values:    String Type
# Author:           Naveed Hasan
#---------------------------------------------------------- 
def GetSubStringFromAStringByPositions(StringToPurseFrom,StartPos,EndPos):
    return			StringToPurseFrom[int(StartPos):int(EndPos)]
#----------------------------------------------------------
# Description:      Function to return the Path of Working Dir
# Input Parameters: N/A
# Return Values:    String Type
# Author:           Naveed Hasan
#---------------------------------------------------------- 
def getWorkingDirectory(__file__):
    #return os.path.dirname(os.path.realpath(sys.argv[0])) 
    exec_filepath = os.path.realpath(__file__)
    exec_dirpath = exec_filepath[0:len(exec_filepath)-len(os.path.basename(__file__))]
    return    exec_dirpath
#----------------------------------------------------------
# Description:      Function to trim x number of characters
#					From a given string.
# Input Parameters: String to trim
# Return Values:    String Type
# Author:           Naveed Hasan
#---------------------------------------------------------- 
def trimFromRight(StringToTrim,NumOfChacToTrim):
    #return			StringToTrim[:-1]
    return			StringToTrim[:-int(NumOfChacToTrim)]
	
#----------------------------------------------------------
# Description:      Function to trim x number of characters
#					From a given string.
# Input Parameters: String to trim
# Return Values:    String Type
# Author:           Naveed Hasan
#---------------------------------------------------------- 
def trimFromLeft(StringToTrim,NumOfChacToTrim):
    return StringToTrim[NumOfChacToTrim:]
	
#----------------------------------------------------------
# Description:      Function to perform string replace
#					For a given string either all occurent or 
#                   First given number of occurences
# Input Parameters: String in which given string to be replaced,
#					String that need to be replaced & one with replaced with
# Return Values:    String Type
# Author:           Naveed Hasan
#---------------------------------------------------------- 
def StringReplaceWith(OrinalString,StringToBeReplaced,StringToBeReplacedBy,NumberOfOccureneToReplace):      
    
    x = int(NumberOfOccureneToReplace)
    if x > 0:
        rVal = OrinalString.replace(StringToBeReplaced, StringToBeReplacedBy, x);
    else:
        rVal = OrinalString.replace(StringToBeReplaced, StringToBeReplacedBy); 
    #print "Modified String", rVal		
    return rVal 
#----------------------------------------------------------
# Description:      Function to retain a Sub string based on their position
#					From a given string.
# Input Parameters: String to trim
# Return Values:    String Type
# Author:           Naveed Hasan
#---------------------------------------------------------- 
def GetSubStringFromAString(StringToPurseFrom,StartPos,EndPos):
    return			StringToPurseFrom[int(StartPos):-int(EndPos)]
#----------------------------------------------------------
# Description:      Function to Find a string in string
#					From a given string.
# Input Parameters: String to Search & String where to search for
# Return Values:    -1 if not found otherwise the position
# Author:           Naveed Hasan
#---------------------------------------------------------- 
def FindStringInString(StringToSearchIn,StringToSearchToSearchFor):
    #return			StringToTrim[:-1]
    return			StringToSearchIn.find(StringToSearchToSearchFor)
	
#----------------------------------------------------------
# Description:      Function to return Random int between given range
# Input Parameters: N/A
# Return Values:    Int Type
# Author:           Naveed Hasan
#---------------------------------------------------------- 
def KillProcessByName(ProcessName):
    #for proc in psutil.process_iter():
        #if proc.name == ProcessName:
                                    #proc.kill()
    #os.system('taskkill /f /im' + ProcessName)
    #os.system("killall -9" + ProcessName)
    #os.system("ps -C" + ProcessName + "-o pid=|xargs kill -9")
    os.system('taskkill /f /im iexplore.exe')
    os.system('taskkill /f /im firefox.exe')
    os.system('taskkill /f /im chrome.exe')
    #os.system('taskkill /f /im chrome.exe*32')
    # p = subprocess.Popen(['pgrep', '-l' , 'firefox'], stdout=subprocess.PIPE)
    # out, err = p.communicate()
    # for line in out.splitlines():        
        # line = bytes.decode(line)
        # pid = int(line.split(None, 1)[0])
        # os.kill(pid, signal.SIGKILL)
		
def KillProcessFirefox():
    os.system('taskkill /f /im firefox.exe')
    #os.system('taskkill /f /im' +' '+ 'firefox.exe')
def KillProcessChrome():
    os.system('taskkill /f /im chrome.exe')
    #os.system('taskkill /f /im' +' '+ 'chrome.exe')
def KillProcessIE():
    os.system('taskkill /f /im iexplore.exe')
    #os.system('taskkill /f /im' +' '+ 'iexplore.exe')
  

#----------------------------------------------------------
# Description:      Function to validate response received for PD_Get Vehicles web service. Response code received is 200.
# Input Parameters: FileName - Path of excel file
# Return Values:    book - opened workbook
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Open_Excel(FileName):
    try:
        book = open_workbook(FileName,formatting_info=True)         # Open requested excel file
    except IOError, e:
        print e
        raise ErrorOccured("Problem in opening file: " + FileName)  # handle IOError
    except:
        raise ErrorOccured("Unknown error occured.")                # handle any other type of error     

    #print "in open_excel"
    return book

def row_count(FileName,  sheetname):
	book = Open_Excel(FileName)
	RowCount = book.sheet_by_name(sheetname).nrows
	return  RowCount

#----------------------------------------------------------
# Description:      Function to return number of rows or columns of excel file
# Input Parameters: Filename - Path of excel file
# 		    RowsorColumns - "Rows" if no. of rows required and "Columns" if no. of columns required
# Return Values:    Number of rows or columns as per request
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Get_NumberOfRowsColumns(FileName,RowsorColumns):
    # open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)
    # return no of rows or columns as per request
    if RowsorColumns == "ROWS":
        rVal = sheet.nrows                                      # Return number of rows
    else:
        rVal = sheet.ncols                                      # Return number of columns
    return rVal                                                 # Return number of rows/columns

#----------------------------------------------------------
# Description:      Function to reset value in required column to some value
# Input Parameters: Filename - Path of excel file
# 		    TotalRows - total number of rows
#                   ColumnName - column name
#                   ReqValue - value to be set in column
# Return Values:    None
# Author:           Manisha Gadekar
#----------------------------------------------------------
def ResetResultColumn(FileName, TotalRows, ColumnName, ReqValue):
    # open excel file
    book = Open_Excel(FileName)

    # Calling a function to create list of headers from excel file
    HeaderList = []
    HeaderList = Generate_HeaderList(FileName)
    
    Result_col = Get_ColumnNumber(HeaderList, ColumnName)   # Calling a function to get column no of "Send Request Result" column
    
    workbook = copy(book)                                   # using copy utility to write to excel file
    w_sheet = workbook.get_sheet(0)                         # locating exact worksheet

    TotalRows = int(TotalRows)

    for i in range (1,TotalRows):                           # row 0 contains titles
        w_sheet.write(i,Result_col,ReqValue)                # using write method to wrote result in cell
    
    workbook.save(FileName)                                 # save changes to excel file

#----------------------------------------------------------
# Description:      Function to return value in required column of excel file 
# Input Parameters: Filename - Path of excel file
# 		    RowNo - Row no in excel file
#                   ColumnName - name of column
# Return Values:    Value in required column of excel file
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Get_ExcelValue(FileName,RowNo,ColumnName):
    # open excel file
    book = Open_Excel(FileName)

    # Calling a function to create list of headers from excel file
    HeaderList = []
    HeaderList = Generate_HeaderList(FileName)
    
    Execute_col = Get_ColumnNumber(HeaderList, ColumnName)       # get column no.
    
    sheet = book.sheet_by_index(0)
    val = sheet.cell(RowNo,Execute_col).value                   # Value in Execute column
    return val                                                  # Returning value in 'Execute' column

#----------------------------------------------------------
# Description:      Function to return value in required column of excel file 
# Input Parameters: Filename - Path of excel file
# 		    RowNo - Row no in excel file
#                   ColumnName - name of column
# Return Values:    Value in required column of excel file
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Get_ExcelValue_ColNo(FileName,RowNo,ColumnNo):
    # open excel file
    book = Open_Excel(FileName)

    sheet = book.sheet_by_index(0)
    try:
        val = sheet.cell(RowNo,int(ColumnNo)).value                   # Value in column
    except:
        raise ErrorOccured("Failed to get value from excel.")
    return val                                                  # Returning value 

#----------------------------------------------------------
# Description:      Function to create list of headers from excel file
# Input Parameters: Filename - Path of excel file
# Return Values:    List of headers
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Generate_HeaderList(FileName):

    print "in generate header"
    # open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)
    print "after open file"
    # get total number of columns
    TotalColumns = Get_NumberOfRowsColumns(FileName,"COLS")
    print "TotalColumns", TotalColumns
    list1 = []  # List to store headings

    # Creating list of headings
    for i in range(TotalColumns):
        x = sheet.cell(0,i).value                               # add excel heading to list
        list1.append(x)

    #print list1
    return list1                                                # return list of headings
    
#----------------------------------------------------------
# Description:      Function to return column number of required header
# Input Parameters: HeaderList - List of headers
#                   Header - Header whose column number is required
# Return Values:    Column number
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Get_ColumnNumber(HeaderList,Header):
    # Get index of heading from list
    ColumnNo = HeaderList.index(Header)                             # get index of required header from list
    #print ColumnNo
    return ColumnNo                                                 # return column no.

#----------------------------------------------------------
# Description:      Function to dynamically create XML string 
# Input Parameters: WebService - Name of WebService e.g. FD_PreQual
#                   Filename - Path of excel file
# 		    RowNo - Row no in excel file
# Return Values:    Request XML string
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Create_RequestString(WebService, FileName, RowNo, FolderPath, SchemafileName):
    # open excel file
    #print FileName
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)

    #print "Folder path: ",FolderPath
    
    # Calling a function to create list of headers from excel file
    HeaderList = []
    HeaderList = Generate_HeaderList(FileName)

    #print HeaderList
    DescColumn = Get_ColumnNumber(HeaderList, "Description")    # get decsription column from excel

    #print "DescColumn: ", DescColumn
    #print "Row no:",RowNo
    Desc = sheet.cell(RowNo,DescColumn).value                   # get decsription from excel

    #print "Desc: ",Desc
    
    #Schemafile = str(FolderPath) + "\\" + SchemafileName #Jenkins Adjustment
    Schemafile = str(FolderPath) + "//" + SchemafileName #Jenkins Adjustment
    #Schemafile = str(FolderPath) + "\\WebService\\" + SchemafileName

    #print FileName
    print "Schema file: ",Schemafile
    print "Description: **********"+ Desc + "**********"
    
    # Request will be created based on the web service name
    if WebService == 'PD_GetVehicle':
        # Calling function to create request string for PD - GetVehicles Web service
        sRequest = Create_PD_GetVehicles_Request(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath)
        print sRequest  
    elif WebService == 'PD_GetVehicleByChromeStyleId':
        # Create request for another webservice
        # Calling function to create request string for PD - GetVehicles Web service
        sRequest = Create_PD_GetVehicleByChromeStyleId_Request(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath)
        print sRequest 
    elif WebService == 'PD_GetMultiVehicle':
        sRequest = Create_PD_GetMultipleVehicles_Request(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath)
        print sRequest
    elif WebService == 'PD_GetMultiVehicleByChrome':
        sRequest = Create_PD_GetMultipleVehiclesByChrome_Request(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath)
        print sRequest
    elif WebService == 'PD_GetIncentivesByChromeMake':
        sRequest = Create_PD_GetIncentivesByChromeMake_Request(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath)
        print sRequest
    elif WebService == 'PD_GetIncentivesByALG':
        sRequest = Create_PD_GetIncentivesByALG_Request(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath)
        print sRequest
    elif WebService == 'PD_GetPayments':
        sRequest = Create_PD_GetPayments_Request(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath)
        print sRequest
    elif WebService == 'PD_GetMultiplePayments':
        sRequest = Create_PD_GetMultiplePayments_Request(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath)
        print sRequest
    elif WebService == 'PD_GetRates':
        sRequest = Create_PD_GetRates_Request(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath)
        print sRequest
    elif WebService == 'PD_GetResiduals':
        sRequest = Create_PD_GetResiduals_Request(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath)
        print sRequest
    elif WebService == 'FD_Prequal_1_0':
        sRequest = Create_FD_Prequal_Request_1_0(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath)
        print sRequest
    elif WebService == 'FD_Prequal_1_1':
        sRequest = Create_FD_Prequal_Request_1_1(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath)
        print sRequest
    elif WebService == 'FD_Lead_1_1':
        sRequest = Create_FD_Lead_Request_1_1(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath)
        print sRequest
    elif WebService == 'FD_CreditApp_1_1':
        sRequest = Create_FD_CreditApp_Request_1_1(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath)
        print sRequest
    elif WebService == 'FD_CreditBureau_1_1':
        sRequest = Create_FD_CreditBureau_Request_1_1(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath)
        print sRequest
    elif WebService == 'Phase_I_FD_Prequal_1_1':
        sRequest = Create_FD_Prequal_Request_1_1_Phase_I(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath)
        print sRequest
    elif WebService == 'Phase_I_FD_Lead_1_1':
        sRequest = Create_FD_Lead_Request_1_1_Phase_I(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath)
        print sRequest
    elif WebService == 'Phase_I_FD_CreditBureau_1_1':
        sRequest = Create_FD_CreditBureau_Request_1_1_Phase_I(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath)
        print sRequest
    elif WebService == 'Phase_I_FD_CreditApp_1_1':
        sRequest = Create_FD_CreditApp_Request_1_1_Phase_1(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath)
        print sRequest
    else:
        raise ErrorOccured("Failed to create request.")
    return sRequest                                             # return XML string

#----------------------------------------------------------
# Description:      Function to create request XML string for PD - GetVehicles web service
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row in excel file
#                   Schemafile - File path of schema file
# Return Values:    request XML string
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Create_PD_GetVehicles_Request(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath):
    # open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)
    
    # get column numbers from excel
    PartnerDealerId = Get_ColumnNumber(HeaderList, "PartnerDealerId")   # Calling a function to get column no of "PartnerDealerId" column
    PartnerId = Get_ColumnNumber(HeaderList, "PartnerId")
    ProfileId = Get_ColumnNumber(HeaderList, "ProfileId")
    ZipCode = Get_ColumnNumber(HeaderList, "ZipCode")
    IncludeRebates = Get_ColumnNumber(HeaderList, "IncludeRebates")
    InstalledOptions = Get_ColumnNumber(HeaderList, "InstalledOptions")
    VehicleCondition = Get_ColumnNumber(HeaderList, "VehicleCondition")
    Make = Get_ColumnNumber(HeaderList, "Make")
    Model = Get_ColumnNumber(HeaderList, "Model")
    Style = Get_ColumnNumber(HeaderList, "Style")
    Year = Get_ColumnNumber(HeaderList, "Year")
    
    # Open schema file
    try:
        f = open(Schemafile)
    except IOError, e:
        print e
        print "Problem while opening Schema file"
        raise ErrorOccured("Problem in opening file: " + Schemafile)        # Problem opening file
    except:
        raise ErrorOccured("Unknown error occured.")                        # handle any other type of error

    # everything is fine
    sRequest = f.read()         # get contenets of schema file in to string
    f.close()
    
    # Create request string
    # Calling function to add values to nodes in given schema file
    sRequest = SetXML(sRequest, "deal1:PartnerDealerId", str(sheet.cell(RowNo,PartnerDealerId).value), 0)   # XML request, node, node value, occurance
    sRequest = SetXML(sRequest, "deal1:PartnerId", str(sheet.cell(RowNo,PartnerId).value), 0)
    sRequest = SetXML(sRequest, "deal1:ProfileId", str(int(sheet.cell(RowNo,ProfileId).value)), 0)
    sRequest = SetXML(sRequest, "deal1:ZipCode", str(int(sheet.cell(RowNo,ZipCode).value)), 0)
    sRequest = SetXML(sRequest, "deal1:IncludeRebates", str(sheet.cell(RowNo,IncludeRebates).value), 0)
    sRequest = SetXML(sRequest, "arr:string", str(int(sheet.cell(RowNo,InstalledOptions).value)), 0)
    sRequest = SetXML(sRequest, "deal1:VehicleCondition", str(sheet.cell(RowNo,VehicleCondition).value), 0)
    sRequest = SetXML(sRequest, "deal1:Make", str(sheet.cell(RowNo,Make).value), 0)
    sRequest = SetXML(sRequest, "deal1:Model", str(sheet.cell(RowNo,Model).value), 0)
    sRequest = SetXML(sRequest, "deal1:Style", str(sheet.cell(RowNo,Style).value), 0)
    sRequest = SetXML(sRequest, "deal1:Year", str(int(sheet.cell(RowNo,Year).value)), 0)

    #print sRequest
    
    #ReqFilePath = str(FolderPath) + '\\Request\\'+ WebService +'Request'+ str(RowNo) +'.xml'
    #ReqFilePath = str(FolderPath) + "\\WebService\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    ReqFilePath = str(FolderPath) + "\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    
    f = open(ReqFilePath, 'w')
    f.write(sRequest)
    f.close()
    
    return sRequest                     # return request string
#----------------------------------------------------------
# Description:      Function to add node value to required node in xml 
# Input Parameters: sRequest - XML string 
#                   NodeName - Name of xml node whose value is to be set
#                   NodeValue - Value to be set for node
#                   Occurance - occurance of the node
# Return Value:     Updated request string
# Author:           Manisha Gadekar
#----------------------------------------------------------
def SetXML(sRequest, NodeName, NodeValue, Occurance):
    # parse xml string using xml.dom.minidom
    try:
        dom = parseString(sRequest)                             # parsing the xml string
    except:
        raise ErrorOccured("Invalid XML. Failed to parse XML input string: " + sRequest)          # not able to parse string
    
    # everything is fine
    try:
        node = dom.getElementsByTagName(NodeName)[Occurance]    # locating the node of interest
    except:
        raise ErrorOccured("Node not found in xml: "+ NodeName) # not able to locate node
    
    try:
        txt = dom.createTextNode(NodeValue.strip())         # creates text node
    except:
        raise ErrorOccured("Unable to add node value: ",NodeValue," to node: " + NodeName)     # problem creating text node
            
    node.appendChild(txt)                               # Add text to given node. Results in <NodeName>Nodevalue</NodeName>
    modified_str = dom.toxml()                          # modified xml string
    return modified_str                                 # return modified string

#----------------------------------------------------------
# Description:      Function to create request XML string for PD - GetVehiclesByChromeStyleId web service
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row in excel file
#                   Schemafile - File path of schema file
# Return Values:    request XML string
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Create_PD_GetVehicleByChromeStyleId_Request(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath):
    # open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)
    
    # get column numbers from excel
    PartnerDealerId = Get_ColumnNumber(HeaderList, "PartnerDealerId")   # Calling a function to get column no of "PartnerDealerId" column
    PartnerId = Get_ColumnNumber(HeaderList, "PartnerId")
    ProfileId = Get_ColumnNumber(HeaderList, "ProfileId")
    ZipCode = Get_ColumnNumber(HeaderList, "ZipCode")
    IncludeRebates = Get_ColumnNumber(HeaderList, "IncludeRebates")
    InstalledOptions = Get_ColumnNumber(HeaderList, "InstalledOptions")
    VehicleCondition = Get_ColumnNumber(HeaderList, "VehicleCondition")
    ChromeStyleId = Get_ColumnNumber(HeaderList, "ChromeStyleId")
    
    # Open schema file
    try:
        f = open(Schemafile)
    except IOError, e:
        print e
        print "Problem while opening Schema file"
        raise ErrorOccured("Problem in opening file: " + Schemafile)        # Problem opening file
    except:
        raise ErrorOccured("Unknown error occured.")                        # handle any other type of error

    # everything is fine
    sRequest = f.read()         # get contenets of schema file in to string
    f.close()
    
    # Create request string
    # Calling function to add values to nodes in given schema file
    sRequest = SetXML(sRequest, "deal1:PartnerDealerId", str(sheet.cell(RowNo,PartnerDealerId).value), 0)   # XML request, node, node value, occurance
    sRequest = SetXML(sRequest, "deal1:PartnerId", str(sheet.cell(RowNo,PartnerId).value), 0)
    sRequest = SetXML(sRequest, "deal1:ProfileId", str(int(sheet.cell(RowNo,ProfileId).value)), 0)
    sRequest = SetXML(sRequest, "deal1:ZipCode", str(int(sheet.cell(RowNo,ZipCode).value)), 0)
    sRequest = SetXML(sRequest, "deal1:IncludeRebates", str(sheet.cell(RowNo,IncludeRebates).value), 0)
    sRequest = SetXML(sRequest, "arr:string", str(int(sheet.cell(RowNo,InstalledOptions).value)), 0)
    sRequest = SetXML(sRequest, "deal1:VehicleCondition", str(sheet.cell(RowNo,VehicleCondition).value), 0)
    sRequest = SetXML(sRequest, "deal1:ChromeStyleId", str(sheet.cell(RowNo,ChromeStyleId).value), 0)
    
    #print sRequest
    
    #ReqFilePath = str(FolderPath) + '\\Request\\'+ WebService +'Request'+ str(RowNo+1) +'.xml'
    #ReqFilePath = str(FolderPath) + "\\WebService\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    ReqFilePath = str(FolderPath) + "\\Request\\" + WebService +'Request'+ str(RowNo+1) +'.xml'
    f = open(ReqFilePath, 'w')
    f.write(sRequest)
    f.close()
    
    return sRequest                     # return request string

#----------------------------------------------------------
# Description:      Function to create request XML string for PD - Get Multiple Vehicles web service
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row in excel file
#                   Schemafile - File path of schema file
# Return Values:    request XML string
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Create_PD_GetMultipleVehicles_Request(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath):
    # open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)
    
    # get column numbers from excel
    PartnerDealerId = Get_ColumnNumber(HeaderList, "PartnerDealerId")   # Calling a function to get column no of "PartnerDealerId" column
    PartnerId = Get_ColumnNumber(HeaderList, "PartnerId")
    ProfileId = Get_ColumnNumber(HeaderList, "ProfileId")
    ZipCode = Get_ColumnNumber(HeaderList, "ZipCode")
    IncludeRebates = Get_ColumnNumber(HeaderList, "IncludeRebates")
    GroupID = Get_ColumnNumber(HeaderList, "GroupID")
    InstalledOptions = Get_ColumnNumber(HeaderList, "InstalledOptions")
    VehicleCondition = Get_ColumnNumber(HeaderList, "VehicleCondition")
    Make = Get_ColumnNumber(HeaderList, "Make")
    Model = Get_ColumnNumber(HeaderList, "Model")
    Style = Get_ColumnNumber(HeaderList, "Style")
    Year = Get_ColumnNumber(HeaderList, "Year")
    
    # Open schema file
    try:
        f = open(Schemafile)
    except IOError, e:
        print e
        print "Problem while opening Schema file"
        raise ErrorOccured("Problem in opening file: " + Schemafile)        # Problem opening file
    except:
        raise ErrorOccured("Unknown error occured.")                        # handle any other type of error

    # everything is fine
    sRequest = f.read()         # get contenets of schema file in to string
    f.close()
    
    # Create request string
    # Calling function to add values to nodes in given schema file
    sRequest = SetXML(sRequest, "deal1:PartnerDealerId", str(sheet.cell(RowNo,PartnerDealerId).value), 0)   # XML request, node, node value, occurance
    sRequest = SetXML(sRequest, "deal1:PartnerId", str(sheet.cell(RowNo,PartnerId).value), 0)
    sRequest = SetXML(sRequest, "deal1:ProfileId", str(int(sheet.cell(RowNo,ProfileId).value)), 0)
    sRequest = SetXML(sRequest, "deal1:ZipCode", str(int(sheet.cell(RowNo,ZipCode).value)), 0)
    sRequest = SetXML(sRequest, "deal:IncludeRebates", str(sheet.cell(RowNo,IncludeRebates).value), 0)
    sRequest = SetXML(sRequest, "deal:GroupID", str(sheet.cell(RowNo,GroupID).value), 0)
    sRequest = SetXML(sRequest, "arr:string", str(int(sheet.cell(RowNo,InstalledOptions).value)), 0)
    sRequest = SetXML(sRequest, "deal2:Condition", str(sheet.cell(RowNo,VehicleCondition).value), 0)
    sRequest = SetXML(sRequest, "deal2:Make", str(sheet.cell(RowNo,Make).value), 0)
    sRequest = SetXML(sRequest, "deal2:Model", str(sheet.cell(RowNo,Model).value), 0)
    sRequest = SetXML(sRequest, "deal2:Style", str(sheet.cell(RowNo,Style).value), 0)
    sRequest = SetXML(sRequest, "deal2:Year", str(int(sheet.cell(RowNo,Year).value)), 0)

    #print sRequest
    
    #ReqFilePath = str(FolderPath) + '\\Request\\'+ WebService +'Request'+ str(RowNo) +'.xml'
    #ReqFilePath = str(FolderPath) + "\\WebService\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    ReqFilePath = str(FolderPath) + "\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    f = open(ReqFilePath, 'w')
    f.write(sRequest)
    f.close()
    
    return sRequest                     # return request string

#----------------------------------------------------------
# Description:      Function to create request XML string for PD - Get Multiple Vehicles By Chrome Style Id web service
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row in excel file
#                   Schemafile - File path of schema file
# Return Values:    request XML string
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Create_PD_GetMultipleVehiclesByChrome_Request(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath):
    # open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)
    
    # get column numbers from excel
    PartnerDealerId = Get_ColumnNumber(HeaderList, "PartnerDealerId")   # Calling a function to get column no of "PartnerDealerId" column
    PartnerId = Get_ColumnNumber(HeaderList, "PartnerId")
    ProfileId = Get_ColumnNumber(HeaderList, "ProfileId")
    ZipCode = Get_ColumnNumber(HeaderList, "ZipCode")
    IncludeRebates = Get_ColumnNumber(HeaderList, "IncludeRebates")
    ChromeStyleId = Get_ColumnNumber(HeaderList, "ChromeStyleId")
    Condition = Get_ColumnNumber(HeaderList, "Condition")
    GroupID = Get_ColumnNumber(HeaderList, "GroupID")
    InstalledOptions = Get_ColumnNumber(HeaderList, "InstalledOptions")
    
    # Open schema file
    try:
        f = open(Schemafile)
    except IOError, e:
        print e
        print "Problem while opening Schema file"
        raise ErrorOccured("Problem in opening file: " + Schemafile)        # Problem opening file
    except:
        raise ErrorOccured("Unknown error occured.")                        # handle any other type of error

    # everything is fine
    sRequest = f.read()         # get contenets of schema file in to string
    f.close()
    
    # Create request string
    # Calling function to add values to nodes in given schema file
    sRequest = SetXML(sRequest, "deal1:PartnerDealerId", str(sheet.cell(RowNo,PartnerDealerId).value), 0)   # XML request, node, node value, occurance
    sRequest = SetXML(sRequest, "deal1:PartnerId", str(sheet.cell(RowNo,PartnerId).value), 0)
    sRequest = SetXML(sRequest, "deal1:ProfileId", str(int(sheet.cell(RowNo,ProfileId).value)), 0)
    sRequest = SetXML(sRequest, "deal1:ZipCode", str(int(sheet.cell(RowNo,ZipCode).value)), 0)
    sRequest = SetXML(sRequest, "deal:IncludeRebates", str(sheet.cell(RowNo,IncludeRebates).value), 0)
    sRequest = SetXML(sRequest, "deal:ChromeStyleId", str(sheet.cell(RowNo,ChromeStyleId).value), 0)
    sRequest = SetXML(sRequest, "deal:Condition", str(sheet.cell(RowNo,Condition).value), 0)
    sRequest = SetXML(sRequest, "deal:GroupID", str(sheet.cell(RowNo,GroupID).value), 0)
    sRequest = SetXML(sRequest, "arr:string", str(int(sheet.cell(RowNo,InstalledOptions).value)), 0)
    

    #print sRequest
    
    #ReqFilePath = str(FolderPath) + '\\Request\\'+ WebService +'Request'+ str(RowNo) +'.xml'
    #ReqFilePath = str(FolderPath) + "\\WebService\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    ReqFilePath = str(FolderPath) + "\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    f = open(ReqFilePath, 'w')
    f.write(sRequest)
    f.close()
    
    return sRequest                     # return request string
#----------------------------------------------------------
# Description:      Function to create request XML string for PD - Get Incentives By Chrome Make web service
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row in excel file
#                   Schemafile - File path of schema file
# Return Values:    request XML string
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Create_PD_GetIncentivesByChromeMake_Request(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath):
    # open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)
    
    # get column numbers from excel
    PartnerDealerId = Get_ColumnNumber(HeaderList, "PartnerDealerId")   # Calling a function to get column no of "PartnerDealerId" column
    PartnerId = Get_ColumnNumber(HeaderList, "PartnerId")
    ProfileId = Get_ColumnNumber(HeaderList, "ProfileId")
    ZipCode = Get_ColumnNumber(HeaderList, "ZipCode")
    ChromeMake = Get_ColumnNumber(HeaderList, "ChromeMake")
    IncludeAlgCodesDesc = Get_ColumnNumber(HeaderList, "IncludeAlgCodesDesc")
    IncludeChromeStyleIds = Get_ColumnNumber(HeaderList, "IncludeChromeStyleIds")
    VehicleCondition = Get_ColumnNumber(HeaderList, "VehicleCondition")
    Year = Get_ColumnNumber(HeaderList, "Year")
    
    # Open schema file
    try:
        f = open(Schemafile)
    except IOError, e:
        print e
        print "Problem while opening Schema file"
        raise ErrorOccured("Problem in opening file: " + Schemafile)        # Problem opening file
    except:
        raise ErrorOccured("Unknown error occured.")                        # handle any other type of error

    # everything is fine
    sRequest = f.read()         # get contenets of schema file in to string
    f.close()
    
    # Create request string
    # Calling function to add values to nodes in given schema file
    sRequest = SetXML(sRequest, "deal1:PartnerDealerId", str(sheet.cell(RowNo,PartnerDealerId).value), 0)   # XML request, node, node value, occurance
    sRequest = SetXML(sRequest, "deal1:PartnerId", str(sheet.cell(RowNo,PartnerId).value), 0)
    sRequest = SetXML(sRequest, "deal1:ProfileId", str(int(sheet.cell(RowNo,ProfileId).value)), 0)
    sRequest = SetXML(sRequest, "deal1:ZipCode", str(int(sheet.cell(RowNo,ZipCode).value)), 0)
    sRequest = SetXML(sRequest, "deal1:ChromeMake", str(sheet.cell(RowNo,ChromeMake).value), 0)
    sRequest = SetXML(sRequest, "deal1:IncludeAlgCodesDesc", str(sheet.cell(RowNo,IncludeAlgCodesDesc).value), 0)
    sRequest = SetXML(sRequest, "deal1:IncludeChromeStyleIds", str(sheet.cell(RowNo,IncludeChromeStyleIds).value), 0)
    sRequest = SetXML(sRequest, "deal1:VehicleCondition", str(sheet.cell(RowNo,VehicleCondition).value), 0)
    sRequest = SetXML(sRequest, "deal1:Year", str(int(sheet.cell(RowNo,Year).value)), 0)
    
    #print sRequest
    
    #ReqFilePath = str(FolderPath) + '\\Request\\'+ WebService +'Request'+ str(RowNo) +'.xml'
    #ReqFilePath = str(FolderPath) + "\\WebService\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    ReqFilePath = str(FolderPath) + "\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    f = open(ReqFilePath, 'w')
    f.write(sRequest)
    f.close()
    
    return sRequest                     # return request string

#----------------------------------------------------------
# Description:      Function to create request XML string for PD - Get Incentives By Chrome Make web service
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row in excel file
#                   Schemafile - File path of schema file
# Return Values:    request XML string
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Create_PD_GetIncentivesByALG_Request(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath):
    # open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)
    
    # get column numbers from excel
    PartnerDealerId = Get_ColumnNumber(HeaderList, "PartnerDealerId")   # Calling a function to get column no of "PartnerDealerId" column
    PartnerId = Get_ColumnNumber(HeaderList, "PartnerId")
    ProfileId = Get_ColumnNumber(HeaderList, "ProfileId")
    ZipCode = Get_ColumnNumber(HeaderList, "ZipCode")
    MakeCode = Get_ColumnNumber(HeaderList, "MakeCode")
    ModelCode = Get_ColumnNumber(HeaderList, "ModelCode")
    StyleCode = Get_ColumnNumber(HeaderList, "StyleCode")
    VehicleCondition = Get_ColumnNumber(HeaderList, "VehicleCondition")
    YearCode = Get_ColumnNumber(HeaderList, "YearCode")
    
    # Open schema file
    try:
        f = open(Schemafile)
    except IOError, e:
        print e
        print "Problem while opening Schema file"
        raise ErrorOccured("Problem in opening file: " + Schemafile)        # Problem opening file
    except:
        raise ErrorOccured("Unknown error occured.")                        # handle any other type of error

    # everything is fine
    sRequest = f.read()         # get contenets of schema file in to string
    f.close()
    
    # Create request string
    # Calling function to add values to nodes in given schema file
    sRequest = SetXML(sRequest, "deal1:PartnerDealerId", str(sheet.cell(RowNo,PartnerDealerId).value), 0)   # XML request, node, node value, occurance
    sRequest = SetXML(sRequest, "deal1:PartnerId", str(sheet.cell(RowNo,PartnerId).value), 0)
    sRequest = SetXML(sRequest, "deal1:ProfileId", str(int(sheet.cell(RowNo,ProfileId).value)), 0)
    sRequest = SetXML(sRequest, "deal1:ZipCode", str(int(sheet.cell(RowNo,ZipCode).value)), 0)
    sRequest = SetXML(sRequest, "deal1:MakeCode", str(int(sheet.cell(RowNo,MakeCode).value)), 0)
    sRequest = SetXML(sRequest, "deal1:ModelCode", str(int(sheet.cell(RowNo,ModelCode).value)), 0)
    sRequest = SetXML(sRequest, "deal1:StyleCode", str(int(sheet.cell(RowNo,StyleCode).value)), 0)
    sRequest = SetXML(sRequest, "deal1:VehicleCondition", str(sheet.cell(RowNo,VehicleCondition).value), 0)
    sRequest = SetXML(sRequest, "deal1:YearCode", str(int(sheet.cell(RowNo,YearCode).value)), 0)
    
    #print sRequest
    
    #ReqFilePath = str(FolderPath) + '\\Request\\'+ WebService +'Request'+ str(RowNo) +'.xml'
    #ReqFilePath = str(FolderPath) + "\\WebService\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    ReqFilePath = str(FolderPath) + "\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    f = open(ReqFilePath, 'w')
    f.write(sRequest)
    f.close()
    
    return sRequest                     # return request string

#----------------------------------------------------------
# Description:      Function to create request XML string for PD - GetVehicles web service
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row in excel file
#                   Schemafile - File path of schema file
# Return Values:    request XML string
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Create_PD_GetPayments_Request(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath):
    # open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)
    
    # get column numbers from excel
    PartnerDealerId = Get_ColumnNumber(HeaderList, "PartnerDealerId")   # Calling a function to get column no of "PartnerDealerId" column
    PartnerId = Get_ColumnNumber(HeaderList, "PartnerId")
    ProfileId = Get_ColumnNumber(HeaderList, "ProfileId")
    ZipCode = Get_ColumnNumber(HeaderList, "ZipCode")
    AnnualMiles = Get_ColumnNumber(HeaderList, "AnnualMiles")
    CustomerTrade = Get_ColumnNumber(HeaderList, "CustomerTrade")
    DownPayment = Get_ColumnNumber(HeaderList, "DownPayment")
    FinanceType = Get_ColumnNumber(HeaderList, "FinanceType")
    #RebateIds = Get_ColumnNumber(HeaderList, "RebateIds")
	
    InstalledOptions = Get_ColumnNumber(HeaderList, "InstalledOptions")
	
    Term = Get_ColumnNumber(HeaderList, "Term")
	
    #AddedOptions = Get_ColumnNumber(HeaderList, "AddedOptions")
    A_Cost = Get_ColumnNumber(HeaderList, "A_Cost")
    A_Retail = Get_ColumnNumber(HeaderList, "A_Retail")
	
    #Msrp = Get_ColumnNumber(HeaderList, "Msrp")
    M_Cost = Get_ColumnNumber(HeaderList, "M_Cost")
    M_Retail = Get_ColumnNumber(HeaderList, "M_Retail")
	
    Odometer = Get_ColumnNumber(HeaderList, "Odometer")
	
    #Options = = Get_ColumnNumber(HeaderList, "Odometer")
    #FactoryOption = Get_ColumnNumber(HeaderList, "FactoryOption")
    Description = Get_ColumnNumber(HeaderList, "Description")
    Installed = Get_ColumnNumber(HeaderList, "Installed")
	
    #RemovedOptions = Get_ColumnNumber(HeaderList, "RemovedOptions")
    R_Cost = Get_ColumnNumber(HeaderList, "R_Cost")
    R_Retail = Get_ColumnNumber(HeaderList, "R_Retail")
	
    #vehicleCode = Get_ColumnNumber(HeaderList, "vehicleCode")	
    Condition_Code = Get_ColumnNumber(HeaderList, "Condition_Code")
    MakeCode = Get_ColumnNumber(HeaderList, "MakeCode")
    ModelCode = Get_ColumnNumber(HeaderList, "ModelCode")
    StyleCode = Get_ColumnNumber(HeaderList, "StyleCode")
    Year_Code = Get_ColumnNumber(HeaderList, "Year_Code")
	
    #VehicleDescription = Get_ColumnNumber(HeaderList, "VehicleDescription")
    Condition = Get_ColumnNumber(HeaderList, "Condition")
    Make = Get_ColumnNumber(HeaderList, "Make")
    Model = Get_ColumnNumber(HeaderList, "Model")
    Style = Get_ColumnNumber(HeaderList, "Style")
    Year = Get_ColumnNumber(HeaderList, "Year")
	
    AreAllFeesCapped = Get_ColumnNumber(HeaderList, "AreAllFeesCapped")
    CreditScore = Get_ColumnNumber(HeaderList, "CreditScore")
    IsTaxCapped = Get_ColumnNumber(HeaderList, "IsTaxCapped")
    OtherFees = Get_ColumnNumber(HeaderList, "OtherFees")
    TaxRate = Get_ColumnNumber(HeaderList, "TaxRate")
	
	    
    # Open schema file
    try:
        f = open(Schemafile)
    except IOError, e:
        print e
        print "Problem while opening Schema file"
        raise ErrorOccured("Problem in opening file: " + Schemafile)        # Problem opening file
    except:
        raise ErrorOccured("Unknown error occured.")                        # handle any other type of error

    # everything is fine
    sRequest = f.read()         # get contenets of schema file in to string
    f.close()
    
    # Create request string
    # Calling function to add values to nodes in given schema file
	
    sRequest = SetXML(sRequest, "deal1:PartnerDealerId", str(sheet.cell(RowNo,PartnerDealerId).value), 0)   # XML request, node, node value, occurance
    sRequest = SetXML(sRequest, "deal1:PartnerId", str(sheet.cell(RowNo,PartnerId).value), 0)
    sRequest = SetXML(sRequest, "deal1:ProfileId", str(int(sheet.cell(RowNo,ProfileId).value)), 0)
    sRequest = SetXML(sRequest, "deal1:ZipCode", str(int(sheet.cell(RowNo,ZipCode).value)), 0)
    sRequest = SetXML(sRequest, "deal1:AnnualMiles", str(int(sheet.cell(RowNo,AnnualMiles).value)), 0)
	
    sRequest = SetXML(sRequest, "arr:string", str(int(sheet.cell(RowNo,InstalledOptions).value)), 0)
	
    sRequest = SetXML(sRequest, "deal1:CustomerTrade", str(int(sheet.cell(RowNo,CustomerTrade).value)), 0)
    sRequest = SetXML(sRequest, "deal1:DownPayment", str(int(sheet.cell(RowNo,DownPayment).value)), 0)
    sRequest = SetXML(sRequest, "deal1:FinanceType", str(sheet.cell(RowNo,FinanceType).value), 0)
    #sRequest = SetXML(sRequest, "deal1:RebateIds", str(int(sheet.cell(RowNo,RebateIds).value)), 0)
    sRequest = SetXML(sRequest, "deal1:Term", str(int(sheet.cell(RowNo,Term).value)), 0)
	
    #sRequest = SetXML(sRequest, "deal1:Vehicle", str(sheet.cell(RowNo,Vehicle).value), 0)
	
    #sRequest = SetXML(sRequest, "deal2:AddedOptions", str(sheet.cell(RowNo,AddedOptions).value), 0)
    sRequest = SetXML(sRequest, "deal3:Cost", str(sheet.cell(RowNo,A_Cost).value), 0)
    sRequest = SetXML(sRequest, "deal3:Retail", str(sheet.cell(RowNo,A_Retail).value), 0)
	
    #sRequest = SetXML(sRequest, "deal2:Msrp", str(int(sheet.cell(RowNo,Msrp).value)), 0)
    sRequest = SetXML(sRequest, "deal3:Cost", str(sheet.cell(RowNo,M_Cost).value), 1)
    sRequest = SetXML(sRequest, "deal3:Retail", str(sheet.cell(RowNo,M_Retail).value), 1)
	
    sRequest = SetXML(sRequest, "deal2:Odometer", str(int(sheet.cell(RowNo,Odometer).value)), 0)
	
    #sRequest = SetXML(sRequest, "deal2:Options", str(sheet.cell(RowNo,Options).value), 0)
    #sRequest = SetXML(sRequest, "deal4:FactoryOption", str(sheet.cell(RowNo,FactoryOption).value), 0)
    sRequest = SetXML(sRequest, "deal4:Description", str(sheet.cell(RowNo,Description).value), 0)
    sRequest = SetXML(sRequest, "deal4:Installed", str(sheet.cell(RowNo,Installed).value), 0)
	
    #sRequest = SetXML(sRequest, "deal2:RemovedOptions", str(sheet.cell(RowNo,RemovedOptions).value), 0)
    sRequest = SetXML(sRequest, "deal3:Cost", str(sheet.cell(RowNo,R_Cost).value), 2)
    sRequest = SetXML(sRequest, "deal3:Retail", str(sheet.cell(RowNo,R_Retail).value), 2)
	
    #sRequest = SetXML(sRequest, "deal2:vehicleCode", str(sheet.cell(RowNo,vehicleCode).value), 0)
    sRequest = SetXML(sRequest, "deal3:Condition", str(sheet.cell(RowNo,Condition_Code).value), 0)
    sRequest = SetXML(sRequest, "deal3:MakeCode", str(int(sheet.cell(RowNo,MakeCode).value)), 0)
    sRequest = SetXML(sRequest, "deal3:ModelCode", str(int(sheet.cell(RowNo,ModelCode).value)), 0)
    sRequest = SetXML(sRequest, "deal3:StyleCode", str(int(sheet.cell(RowNo,StyleCode).value)), 0)
    sRequest = SetXML(sRequest, "deal3:Year", str(int(sheet.cell(RowNo,Year_Code).value)), 0)
	
    #sRequest = SetXML(sRequest, "deal2:VehicleDescription", str(sheet.cell(RowNo,VehicleDescription).value), 0)
    sRequest = SetXML(sRequest, "deal3:Condition", str(sheet.cell(RowNo,Condition).value), 1)
    sRequest = SetXML(sRequest, "deal3:Make", str(sheet.cell(RowNo,Make).value), 0)
    sRequest = SetXML(sRequest, "deal3:Model", str(sheet.cell(RowNo,Model).value), 0)
    sRequest = SetXML(sRequest, "deal3:Style", str(sheet.cell(RowNo,Style).value), 0)
    sRequest = SetXML(sRequest, "deal3:Year", str(sheet.cell(RowNo,Year).value), 1)
	
    sRequest = SetXML(sRequest, "deal:AreAllFeesCapped", str(sheet.cell(RowNo,AreAllFeesCapped).value), 0)
    sRequest = SetXML(sRequest, "deal:CreditScore", str(int(sheet.cell(RowNo,CreditScore).value)), 0)
    sRequest = SetXML(sRequest, "deal:IsTaxCapped", str(sheet.cell(RowNo,IsTaxCapped).value), 0)
    sRequest = SetXML(sRequest, "deal:OtherFees", str(float(sheet.cell(RowNo,OtherFees).value)), 0)
    sRequest = SetXML(sRequest, "deal:TaxRate", str(float(sheet.cell(RowNo,TaxRate).value)), 0)
	
    #print sRequest
    
    #ReqFilePath = str(FolderPath) + '\\Request\\'+ WebService +'Request'+ str(RowNo) +'.xml'
    #ReqFilePath = str(FolderPath) + "\\WebService\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    ReqFilePath = str(FolderPath) + "\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    f = open(ReqFilePath, 'w')
    f.write(sRequest)
    f.close()
    
    return sRequest                     # return request string

#---------------------------------------------------------------------------------------------------------------
# Description:      Function to create request XML string for PD - GetMultiplePayments web service
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row in excel file
#                   Schemafile - File path of schema file
# Return Values:    request XML string
# Author:           Manisha Gadekar
#------------------------------------------------------------------------------------------------------

def Create_PD_GetMultiplePayments_Request(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath):
    # open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)
    
    # get column numbers from excel
    PartnerDealerId = Get_ColumnNumber(HeaderList, "PartnerDealerId")   # Calling a function to get column no of "PartnerDealerId" column
    PartnerId = Get_ColumnNumber(HeaderList, "PartnerId")
    ProfileId = Get_ColumnNumber(HeaderList, "ProfileId")
    ZipCode = Get_ColumnNumber(HeaderList, "ZipCode")
    
    AnnualMiles = Get_ColumnNumber(HeaderList, "AnnualMiles")
    AreAllFeesCapped = Get_ColumnNumber(HeaderList, "AreAllFeesCapped")
    CustomerTrade = Get_ColumnNumber(HeaderList, "CustomerTrade")
    DownPayment = Get_ColumnNumber(HeaderList, "DownPayment")
    
    IsTaxCapped = Get_ColumnNumber(HeaderList, "IsTaxCapped")
    OtherFees = Get_ColumnNumber(HeaderList, "OtherFees")
    TaxRate = Get_ColumnNumber(HeaderList, "TaxRate")
    CreditScore = Get_ColumnNumber(HeaderList, "CreditScore")
    FinanceType = Get_ColumnNumber(HeaderList, "FinanceType")
    GroupId = Get_ColumnNumber(HeaderList, "GroupId")

    RebateList = Get_ColumnNumber(HeaderList, "RebateList")
    TermList = Get_ColumnNumber(HeaderList, "TermList")
    
    
    A_Cost = Get_ColumnNumber(HeaderList, "A_Cost")
    A_Retail = Get_ColumnNumber(HeaderList, "A_Retail")
    
    M_Cost = Get_ColumnNumber(HeaderList, "M_Cost")
    M_Retail = Get_ColumnNumber(HeaderList, "M_Retail")
    
    Odometer = Get_ColumnNumber(HeaderList, "Odometer")
    Description = Get_ColumnNumber(HeaderList, "Description")
    Installed = Get_ColumnNumber(HeaderList, "Installed")
    
    R_Cost = Get_ColumnNumber(HeaderList, "R_Cost")
    R_Retail = Get_ColumnNumber(HeaderList, "R_Retail")
    
    Condition_Code = Get_ColumnNumber(HeaderList, "Condition_Code")
    MakeCode = Get_ColumnNumber(HeaderList, "MakeCode")
    ModelCode = Get_ColumnNumber(HeaderList, "ModelCode")
    StyleCode = Get_ColumnNumber(HeaderList, "StyleCode")
    Year_Code = Get_ColumnNumber(HeaderList, "Year_Code")

    Condition = Get_ColumnNumber(HeaderList, "Condition")
    Make = Get_ColumnNumber(HeaderList, "Make")
    Model = Get_ColumnNumber(HeaderList, "Model")
    Style = Get_ColumnNumber(HeaderList, "Style")
    Year = Get_ColumnNumber(HeaderList, "Year")
    
    # Open schema file
    try:
        f = open(Schemafile)
    except IOError, e:
        print e
        print "Problem while opening Schema file"
        raise ErrorOccured("Problem in opening file: " + Schemafile)        # Problem opening file
    except:
        raise ErrorOccured("Unknown error occured.")                        # handle any other type of error

    # everything is fine
    sRequest = f.read()         # get contenets of schema file in to string
    f.close()
    
    # Create request string
    # Calling function to add values to nodes in given schema file
	
    sRequest = SetXML(sRequest, "deal1:PartnerDealerId", str(sheet.cell(RowNo,PartnerDealerId).value), 0)   # XML request, node, node value, occurance
    sRequest = SetXML(sRequest, "deal1:PartnerId", str(sheet.cell(RowNo,PartnerId).value), 0)
    sRequest = SetXML(sRequest, "deal1:ProfileId", str(int(sheet.cell(RowNo,ProfileId).value)), 0)
    
    sRequest = SetXML(sRequest, "deal1:ZipCode", str(int(sheet.cell(RowNo,ZipCode).value)), 0)
    sRequest = SetXML(sRequest, "deal:AnnualMiles", str(int(sheet.cell(RowNo,AnnualMiles).value)), 0)
    sRequest = SetXML(sRequest, "deal:AreAllFeesCapped", str(sheet.cell(RowNo,AreAllFeesCapped).value), 0)
    sRequest = SetXML(sRequest, "deal:CustomerTrade", str(int(sheet.cell(RowNo,CustomerTrade).value)), 0)
    sRequest = SetXML(sRequest, "deal:DownPayment", str(int(sheet.cell(RowNo,DownPayment).value)), 0)
    sRequest = SetXML(sRequest, "deal:IsTaxCapped", str(sheet.cell(RowNo,IsTaxCapped).value), 0)
    sRequest = SetXML(sRequest, "deal:OtherFees", str(float(sheet.cell(RowNo,OtherFees).value)), 0)
    sRequest = SetXML(sRequest, "deal:TaxRate", str(float(sheet.cell(RowNo,TaxRate).value)), 0)
    sRequest = SetXML(sRequest, "deal2:CreditScore", str(int(sheet.cell(RowNo,CreditScore).value)), 0)
    sRequest = SetXML(sRequest, "deal2:FinanceType", str(sheet.cell(RowNo,FinanceType).value), 0)
    sRequest = SetXML(sRequest, "deal2:GroupId", str(sheet.cell(RowNo,GroupId).value), 0)

    sRequest = SetXML(sRequest, "arr:int", str(sheet.cell(RowNo,RebateList).value), 0)
    sRequest = SetXML(sRequest, "arr:int", str(sheet.cell(RowNo,TermList).value), 1)

    sRequest = SetXML(sRequest, "deal3:Cost", str(sheet.cell(RowNo,A_Cost).value), 0)
    sRequest = SetXML(sRequest, "deal3:Retail", str(sheet.cell(RowNo,A_Retail).value), 0)
    sRequest = SetXML(sRequest, "deal3:Cost", str(sheet.cell(RowNo,M_Cost).value), 1)
    sRequest = SetXML(sRequest, "deal3:Retail", str(sheet.cell(RowNo,M_Retail).value), 1)
    sRequest = SetXML(sRequest, "deal2:Odometer", str(int(sheet.cell(RowNo,Odometer).value)), 0)
    sRequest = SetXML(sRequest, "deal4:Description", str(sheet.cell(RowNo,Description).value), 0)
    sRequest = SetXML(sRequest, "deal4:Installed", str(sheet.cell(RowNo,Installed).value), 0)

    sRequest = SetXML(sRequest, "deal3:Cost", str(sheet.cell(RowNo,R_Cost).value), 2)
    sRequest = SetXML(sRequest, "deal3:Retail", str(sheet.cell(RowNo,R_Retail).value), 2)

    sRequest = SetXML(sRequest, "deal3:Condition", str(sheet.cell(RowNo,Condition_Code).value), 0)
    sRequest = SetXML(sRequest, "deal3:MakeCode", str(int(sheet.cell(RowNo,MakeCode).value)), 0)
    sRequest = SetXML(sRequest, "deal3:ModelCode", str(int(sheet.cell(RowNo,ModelCode).value)), 0)
    sRequest = SetXML(sRequest, "deal3:StyleCode", str(int(sheet.cell(RowNo,StyleCode).value)), 0)
    sRequest = SetXML(sRequest, "deal3:Year", str(int(sheet.cell(RowNo,Year_Code).value)), 0)

    sRequest = SetXML(sRequest, "deal3:Condition", str(sheet.cell(RowNo,Condition).value), 1)
    sRequest = SetXML(sRequest, "deal3:Make", str(sheet.cell(RowNo,Make).value), 0)
    sRequest = SetXML(sRequest, "deal3:Model", str(sheet.cell(RowNo,Model).value), 0)
    sRequest = SetXML(sRequest, "deal3:Style", str(sheet.cell(RowNo,Style).value), 0)
    sRequest = SetXML(sRequest, "deal3:Year", str(int(sheet.cell(RowNo,Year).value)), 1)

        #print sRequest
    
    #ReqFilePath = str(FolderPath) + '\\Request\\'+ WebService +'Request'+ str(RowNo) +'.xml'
    #ReqFilePath = str(FolderPath) + "\\WebService\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    ReqFilePath = str(FolderPath) + "\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    f = open(ReqFilePath, 'w')
    f.write(sRequest)
    f.close()
    
    return sRequest                     # return request string

#---------------------------------------------------------------------------------------------------------   
# Description:      Function to create request XML string for PD - GetVehicles web service
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row in excel file
#                   Schemafile - File path of schema file
# Return Values:    request XML string
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Create_PD_GetRates_Request(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath):
    # open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)
    
    # get column numbers from excel
    PartnerDealerId = Get_ColumnNumber(HeaderList, "PartnerDealerId")   # Calling a function to get column no of "PartnerDealerId" column
    PartnerId = Get_ColumnNumber(HeaderList, "PartnerId")
    ProfileId = Get_ColumnNumber(HeaderList, "ProfileId")
    ZipCode = Get_ColumnNumber(HeaderList, "ZipCode")
    AnnualMiles = Get_ColumnNumber(HeaderList, "AnnualMiles")
    CustomerTrade = Get_ColumnNumber(HeaderList, "CustomerTrade")
    DownPayment = Get_ColumnNumber(HeaderList, "DownPayment")
    FinanceType = Get_ColumnNumber(HeaderList, "FinanceType")
	
    InstalledOptions = Get_ColumnNumber(HeaderList, "InstalledOptions")
	
    Term = Get_ColumnNumber(HeaderList, "Term")	
    
    A_Cost = Get_ColumnNumber(HeaderList, "A_Cost")
    A_Retail = Get_ColumnNumber(HeaderList, "A_Retail")
	    
    M_Cost = Get_ColumnNumber(HeaderList, "M_Cost")
    M_Retail = Get_ColumnNumber(HeaderList, "M_Retail")
	
    Odometer = Get_ColumnNumber(HeaderList, "Odometer")
    Description = Get_ColumnNumber(HeaderList, "Description")
    Installed = Get_ColumnNumber(HeaderList, "Installed")
	
    R_Cost = Get_ColumnNumber(HeaderList, "R_Cost")
    R_Retail = Get_ColumnNumber(HeaderList, "R_Retail")
		
    Condition_Code = Get_ColumnNumber(HeaderList, "Condition_Code")
    MakeCode = Get_ColumnNumber(HeaderList, "MakeCode")
    ModelCode = Get_ColumnNumber(HeaderList, "ModelCode")
    StyleCode = Get_ColumnNumber(HeaderList, "StyleCode")
    Year_Code = Get_ColumnNumber(HeaderList, "Year_Code")
	
    Condition = Get_ColumnNumber(HeaderList, "Condition")
    Make = Get_ColumnNumber(HeaderList, "Make")
    Model = Get_ColumnNumber(HeaderList, "Model")
    Style = Get_ColumnNumber(HeaderList, "Style")
    Year = Get_ColumnNumber(HeaderList, "Year")
	
    AreAllFeesCapped = Get_ColumnNumber(HeaderList, "AreAllFeesCapped")
    CreditScore = Get_ColumnNumber(HeaderList, "CreditScore")
    IsTaxCapped = Get_ColumnNumber(HeaderList, "IsTaxCapped")
    OtherFees = Get_ColumnNumber(HeaderList, "OtherFees")
    TaxRate = Get_ColumnNumber(HeaderList, "TaxRate")
	
	    
    # Open schema file
    try:
        f = open(Schemafile)
    except IOError, e:
        print e
        print "Problem while opening Schema file"
        raise ErrorOccured("Problem in opening file: " + Schemafile)        # Problem opening file
    except:
        raise ErrorOccured("Unknown error occured.")                        # handle any other type of error

    # everything is fine
    sRequest = f.read()         # get contenets of schema file in to string
    f.close()
    
    # Create request string
    # Calling function to add values to nodes in given schema file
    sRequest = SetXML(sRequest, "deal1:PartnerDealerId", str(sheet.cell(RowNo,PartnerDealerId).value), 0)   # XML request, node, node value, occurance
    sRequest = SetXML(sRequest, "deal1:PartnerId", str(sheet.cell(RowNo,PartnerId).value), 0)
    sRequest = SetXML(sRequest, "deal1:ProfileId", str(int(sheet.cell(RowNo,ProfileId).value)), 0)
    sRequest = SetXML(sRequest, "deal1:ZipCode", str(int(sheet.cell(RowNo,ZipCode).value)), 0)
    sRequest = SetXML(sRequest, "deal1:AnnualMiles", str(int(sheet.cell(RowNo,AnnualMiles).value)), 0)
	
    sRequest = SetXML(sRequest, "arr:string", str(int(sheet.cell(RowNo,InstalledOptions).value)), 0)
	
    sRequest = SetXML(sRequest, "deal1:CustomerTrade", str(int(sheet.cell(RowNo,CustomerTrade).value)), 0)
    sRequest = SetXML(sRequest, "deal1:DownPayment", str(int(sheet.cell(RowNo,DownPayment).value)), 0)
    sRequest = SetXML(sRequest, "deal1:FinanceType", str(sheet.cell(RowNo,FinanceType).value), 0)
    sRequest = SetXML(sRequest, "deal1:Term", str(int(sheet.cell(RowNo,Term).value)), 0)
    
    sRequest = SetXML(sRequest, "deal3:Cost", str(sheet.cell(RowNo,A_Cost).value), 0)
    sRequest = SetXML(sRequest, "deal3:Retail", str(sheet.cell(RowNo,A_Retail).value), 0)
	
    sRequest = SetXML(sRequest, "deal3:Cost", str(sheet.cell(RowNo,M_Cost).value), 1)
    sRequest = SetXML(sRequest, "deal3:Retail", str(sheet.cell(RowNo,M_Retail).value), 1)
	
    sRequest = SetXML(sRequest, "deal2:Odometer", str(int(sheet.cell(RowNo,Odometer).value)), 0)
    sRequest = SetXML(sRequest, "deal4:Description", str(sheet.cell(RowNo,Description).value), 0)
    sRequest = SetXML(sRequest, "deal4:Installed", str(sheet.cell(RowNo,Installed).value), 0)
    
    sRequest = SetXML(sRequest, "deal3:Cost", str(sheet.cell(RowNo,R_Cost).value), 2)
    sRequest = SetXML(sRequest, "deal3:Retail", str(sheet.cell(RowNo,R_Retail).value), 2)
	
    sRequest = SetXML(sRequest, "deal3:Condition", str(sheet.cell(RowNo,Condition_Code).value), 0)
    sRequest = SetXML(sRequest, "deal3:MakeCode", str(int(sheet.cell(RowNo,MakeCode).value)), 0)
    sRequest = SetXML(sRequest, "deal3:ModelCode", str(int(sheet.cell(RowNo,ModelCode).value)), 0)
    sRequest = SetXML(sRequest, "deal3:StyleCode", str(int(sheet.cell(RowNo,StyleCode).value)), 0)
    sRequest = SetXML(sRequest, "deal3:Year", str(int(sheet.cell(RowNo,Year_Code).value)), 0)
	
    sRequest = SetXML(sRequest, "deal3:Condition", str(sheet.cell(RowNo,Condition).value), 1)
    sRequest = SetXML(sRequest, "deal3:Make", str(sheet.cell(RowNo,Make).value), 0)
    sRequest = SetXML(sRequest, "deal3:Model", str(sheet.cell(RowNo,Model).value), 0)
    sRequest = SetXML(sRequest, "deal3:Style", str(sheet.cell(RowNo,Style).value), 0)
    sRequest = SetXML(sRequest, "deal3:Year", str(sheet.cell(RowNo,Year).value), 1)
	
    sRequest = SetXML(sRequest, "deal:AreAllFeesCapped", str(sheet.cell(RowNo,AreAllFeesCapped).value), 0)
    sRequest = SetXML(sRequest, "deal:CreditScore", str(int(sheet.cell(RowNo,CreditScore).value)), 0)
    sRequest = SetXML(sRequest, "deal:IsTaxCapped", str(sheet.cell(RowNo,IsTaxCapped).value), 0)
    sRequest = SetXML(sRequest, "deal:OtherFees", str(float(sheet.cell(RowNo,OtherFees).value)), 0)
    sRequest = SetXML(sRequest, "deal:TaxRate", str(float(sheet.cell(RowNo,TaxRate).value)), 0)
	

    #print sRequest
    
    #ReqFilePath = str(FolderPath) + '\\Request\\'+ WebService +'Request'+ str(RowNo) +'.xml'
    #ReqFilePath = str(FolderPath) + "\\WebService\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    ReqFilePath = str(FolderPath) + "\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    f = open(ReqFilePath, 'w')
    f.write(sRequest)
    f.close()
    
    return sRequest                     # return request string

#---------------------------------------------------------------------------------------------------
# Description:      Function to create request XML string for PD - GetVehicles web service
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row in excel file
#                   Schemafile - File path of schema file
# Return Values:    request XML string
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Create_PD_GetResiduals_Request(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath):
    # open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)
    
    # get column numbers from excel
    PartnerDealerId = Get_ColumnNumber(HeaderList, "PartnerDealerId")   # Calling a function to get column no of "PartnerDealerId" column
    PartnerId = Get_ColumnNumber(HeaderList, "PartnerId")
    ProfileId = Get_ColumnNumber(HeaderList, "ProfileId")
    ZipCode = Get_ColumnNumber(HeaderList, "ZipCode")
    AnnualMiles = Get_ColumnNumber(HeaderList, "AnnualMiles")
    CustomerTrade = Get_ColumnNumber(HeaderList, "CustomerTrade")
    DownPayment = Get_ColumnNumber(HeaderList, "DownPayment")
    FinanceType = Get_ColumnNumber(HeaderList, "FinanceType")
	
    InstalledOptions = Get_ColumnNumber(HeaderList, "InstalledOptions")
	
    Term = Get_ColumnNumber(HeaderList, "Term")	
    
    A_Cost = Get_ColumnNumber(HeaderList, "A_Cost")
    A_Retail = Get_ColumnNumber(HeaderList, "A_Retail")
	    
    M_Cost = Get_ColumnNumber(HeaderList, "M_Cost")
    M_Retail = Get_ColumnNumber(HeaderList, "M_Retail")
	
    Odometer = Get_ColumnNumber(HeaderList, "Odometer")
    Description = Get_ColumnNumber(HeaderList, "Description")
    Installed = Get_ColumnNumber(HeaderList, "Installed")
	
    R_Cost = Get_ColumnNumber(HeaderList, "R_Cost")
    R_Retail = Get_ColumnNumber(HeaderList, "R_Retail")
		
    Condition_Code = Get_ColumnNumber(HeaderList, "Condition_Code")
    MakeCode = Get_ColumnNumber(HeaderList, "MakeCode")
    ModelCode = Get_ColumnNumber(HeaderList, "ModelCode")
    StyleCode = Get_ColumnNumber(HeaderList, "StyleCode")
    Year_Code = Get_ColumnNumber(HeaderList, "Year_Code")
	
    Condition = Get_ColumnNumber(HeaderList, "Condition")
    Make = Get_ColumnNumber(HeaderList, "Make")
    Model = Get_ColumnNumber(HeaderList, "Model")
    Style = Get_ColumnNumber(HeaderList, "Style")
    Year = Get_ColumnNumber(HeaderList, "Year")
	
    AreAllFeesCapped = Get_ColumnNumber(HeaderList, "AreAllFeesCapped")
    CreditScore = Get_ColumnNumber(HeaderList, "CreditScore")
    IsTaxCapped = Get_ColumnNumber(HeaderList, "IsTaxCapped")
    OtherFees = Get_ColumnNumber(HeaderList, "OtherFees")
    TaxRate = Get_ColumnNumber(HeaderList, "TaxRate")
	
	    
    # Open schema file
    try:
        f = open(Schemafile)
    except IOError, e:
        print e
        print "Problem while opening Schema file"
        raise ErrorOccured("Problem in opening file: " + Schemafile)        # Problem opening file
    except:
        raise ErrorOccured("Unknown error occured.")                        # handle any other type of error

    # everything is fine
    sRequest = f.read()         # get contenets of schema file in to string
    f.close()
    
    # Create request string
    # Calling function to add values to nodes in given schema file
    sRequest = SetXML(sRequest, "deal1:PartnerDealerId", str(sheet.cell(RowNo,PartnerDealerId).value), 0)   # XML request, node, node value, occurance
    sRequest = SetXML(sRequest, "deal1:PartnerId", str(sheet.cell(RowNo,PartnerId).value), 0)
    sRequest = SetXML(sRequest, "deal1:ProfileId", str(int(sheet.cell(RowNo,ProfileId).value)), 0)
    sRequest = SetXML(sRequest, "deal1:ZipCode", str(int(sheet.cell(RowNo,ZipCode).value)), 0)
    sRequest = SetXML(sRequest, "deal1:AnnualMiles", str(int(sheet.cell(RowNo,AnnualMiles).value)), 0)
	
    sRequest = SetXML(sRequest, "arr:string", str(sheet.cell(RowNo,InstalledOptions).value), 0)
	
    sRequest = SetXML(sRequest, "deal1:CustomerTrade", str(int(sheet.cell(RowNo,CustomerTrade).value)), 0)
    sRequest = SetXML(sRequest, "deal1:DownPayment", str(int(sheet.cell(RowNo,DownPayment).value)), 0)
    sRequest = SetXML(sRequest, "deal1:FinanceType", str(sheet.cell(RowNo,FinanceType).value), 0)
    sRequest = SetXML(sRequest, "deal1:Term", str(int(sheet.cell(RowNo,Term).value)), 0)
    
    sRequest = SetXML(sRequest, "deal3:Cost", str(sheet.cell(RowNo,A_Cost).value), 0)
    sRequest = SetXML(sRequest, "deal3:Retail", str(sheet.cell(RowNo,A_Retail).value), 0)
	
    sRequest = SetXML(sRequest, "deal3:Cost", str(sheet.cell(RowNo,M_Cost).value), 1)
    sRequest = SetXML(sRequest, "deal3:Retail", str(sheet.cell(RowNo,M_Retail).value), 1)
	
    sRequest = SetXML(sRequest, "deal2:Odometer", str(int(sheet.cell(RowNo,Odometer).value)), 0)
    sRequest = SetXML(sRequest, "deal4:Description", str(sheet.cell(RowNo,Description).value), 0)
    sRequest = SetXML(sRequest, "deal4:Installed", str(sheet.cell(RowNo,Installed).value), 0)
    
    sRequest = SetXML(sRequest, "deal3:Cost", str(sheet.cell(RowNo,R_Cost).value), 2)
    sRequest = SetXML(sRequest, "deal3:Retail", str(sheet.cell(RowNo,R_Retail).value), 2)
	
    sRequest = SetXML(sRequest, "deal3:Condition", str(sheet.cell(RowNo,Condition_Code).value), 0)
    sRequest = SetXML(sRequest, "deal3:MakeCode", str(int(sheet.cell(RowNo,MakeCode).value)), 0)
    sRequest = SetXML(sRequest, "deal3:ModelCode", str(int(sheet.cell(RowNo,ModelCode).value)), 0)
    sRequest = SetXML(sRequest, "deal3:StyleCode", str(int(sheet.cell(RowNo,StyleCode).value)), 0)
    sRequest = SetXML(sRequest, "deal3:Year", str(int(sheet.cell(RowNo,Year_Code).value)), 0)
	
    sRequest = SetXML(sRequest, "deal3:Condition", str(sheet.cell(RowNo,Condition).value), 1)
    sRequest = SetXML(sRequest, "deal3:Make", str(sheet.cell(RowNo,Make).value), 0)
    sRequest = SetXML(sRequest, "deal3:Model", str(sheet.cell(RowNo,Model).value), 0)
    sRequest = SetXML(sRequest, "deal3:Style", str(sheet.cell(RowNo,Style).value), 0)
    sRequest = SetXML(sRequest, "deal3:Year", str(sheet.cell(RowNo,Year).value), 1)
	
    sRequest = SetXML(sRequest, "deal:AreAllFeesCapped", str(sheet.cell(RowNo,AreAllFeesCapped).value), 0)
    sRequest = SetXML(sRequest, "deal:CreditScore", str(int(sheet.cell(RowNo,CreditScore).value)), 0)
    sRequest = SetXML(sRequest, "deal:IsTaxCapped", str(sheet.cell(RowNo,IsTaxCapped).value), 0)
    sRequest = SetXML(sRequest, "deal:OtherFees", str(float(sheet.cell(RowNo,OtherFees).value)), 0)
    sRequest = SetXML(sRequest, "deal:TaxRate", str(float(sheet.cell(RowNo,TaxRate).value)), 0)
	

    #print sRequest
    
    #ReqFilePath = str(FolderPath) + '\\Request\\'+ WebService +'Request'+ str(RowNo) +'.xml'
    #ReqFilePath = str(FolderPath) + "\\WebService\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    ReqFilePath = str(FolderPath) + "\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    f = open(ReqFilePath, 'w')
    f.write(sRequest)
    f.close()
    
    return sRequest                     # return request string


#----------------------------------------------------------
# Description:      Function to create request XML string for FD - PreQualification web service for FD 1.0
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row in excel file
#                   Schemafile - File path of schema file
# Return Values:    request XML string
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Create_FD_Prequal_Request_1_0(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath):
    # open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)
    
    # get column numbers from excel
    PartnerId = Get_ColumnNumber(HeaderList, "PartnerId")   # Calling a function to get column no of "PartnerDealerId" column
    PartnerDealerId = Get_ColumnNumber(HeaderList, "PartnerDealerIds")
    CustomerConsent = Get_ColumnNumber(HeaderList, "CustomerConsent")
    FirstName = Get_ColumnNumber(HeaderList, "FirstName")
    MiddleInitial = Get_ColumnNumber(HeaderList, "MiddleInitial")
    LastName = Get_ColumnNumber(HeaderList, "LastName")
    SSN = Get_ColumnNumber(HeaderList, "SSN")
    AddressLine1 = Get_ColumnNumber(HeaderList, "AddressLine1")
    AddressLine2 = Get_ColumnNumber(HeaderList, "AddressLine2")
    City = Get_ColumnNumber(HeaderList, "City")
    State = Get_ColumnNumber(HeaderList, "State")
    ZipCode = Get_ColumnNumber(HeaderList, "ZipCode")
    TotalMonthsAtEmployment = Get_ColumnNumber(HeaderList, "TotalMonthsAtEmployment")
    MonthlyIncome = Get_ColumnNumber(HeaderList, "MonthlyIncome")
    
    # Open schema file
    try:
        f = open(Schemafile)
    except IOError, e:
        print e
        print "Problem while opening Schema file"
        raise ErrorOccured("Problem in opening file: " + Schemafile)        # Problem opening file
    except:
        raise ErrorOccured("Unknown error occured.")                        # handle any other type of error

    # everything is fine
    sRequest = f.read()         # get contenets of schema file in to string
    f.close()
    
    # Create request string
    # Calling function to add values to nodes in given schema file
    sRequest = SetXML(sRequest, "PartnerId", str(sheet.cell(RowNo,PartnerId).value), 0)
    sRequest = SetXML(sRequest, "a:string", str(int(sheet.cell(RowNo,PartnerDealerId).value)), 0)
    sRequest = SetXML(sRequest, "CustomerConsent", str(sheet.cell(RowNo,CustomerConsent).value), 0)
    sRequest = SetXML(sRequest, "a:FirstName", str(sheet.cell(RowNo,FirstName).value), 0)
    sRequest = SetXML(sRequest, "a:MiddleInitial", str(sheet.cell(RowNo,MiddleInitial).value), 0)
    sRequest = SetXML(sRequest, "a:LastName", str(sheet.cell(RowNo,LastName).value), 0)
    sRequest = SetXML(sRequest, "a:SSN", str(sheet.cell(RowNo,SSN).value), 0)
    sRequest = SetXML(sRequest, "a:AddressLine1", str(sheet.cell(RowNo,AddressLine1).value), 0)
    sRequest = SetXML(sRequest, "a:AddressLine2", str(sheet.cell(RowNo,AddressLine2).value), 0)
    sRequest = SetXML(sRequest, "a:City", str(sheet.cell(RowNo,City).value), 0)
    sRequest = SetXML(sRequest, "a:State", str(sheet.cell(RowNo,State).value), 0)
    sRequest = SetXML(sRequest, "a:ZipCode", str(int(sheet.cell(RowNo,ZipCode).value)), 0)
    sRequest = SetXML(sRequest, "a:TotalMonthsAtEmployment", str(int(sheet.cell(RowNo,TotalMonthsAtEmployment).value)), 0)
    sRequest = SetXML(sRequest, "a:MonthlyIncome", str(int(sheet.cell(RowNo,MonthlyIncome).value)), 0)

   #print sRequest
    
    #ReqFilePath = str(FolderPath) + '\\Request\\'+ WebService +'Request'+ str(RowNo) +'.xml'
    #ReqFilePath = str(FolderPath) + "\\WebService\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    ReqFilePath = str(FolderPath) +  "\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    f = open(ReqFilePath, 'w')
    f.write(sRequest)
    f.close()
    
    return sRequest                     # return request string

#----------------------------------------------------------
# Description:      Function to create request XML string for FD - PreQualification web service for FD 1.1
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row in excel file
#                   Schemafile - File path of schema file
# Return Values:    request XML string
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Create_FD_Prequal_Request_1_1(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath):
    # open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)
    
    # get column numbers from excel
    PartnerId = Get_ColumnNumber(HeaderList, "PartnerId")   # Calling a function to get column no of "PartnerDealerId" column
    PartnerDealerId = Get_ColumnNumber(HeaderList, "PartnerDealerIds")
    FirstName = Get_ColumnNumber(HeaderList, "FirstName")
    MiddleInitial = Get_ColumnNumber(HeaderList, "MiddleInitial")
    LastName = Get_ColumnNumber(HeaderList, "LastName")
    SSN = Get_ColumnNumber(HeaderList, "SSN")
    AddressLine1 = Get_ColumnNumber(HeaderList, "AddressLine1")
    AddressLine2 = Get_ColumnNumber(HeaderList, "AddressLine2")
    City = Get_ColumnNumber(HeaderList, "City")
    State = Get_ColumnNumber(HeaderList, "State")
    ZipCode = Get_ColumnNumber(HeaderList, "ZipCode")
    TotalMonthsEmployed = Get_ColumnNumber(HeaderList, "TotalMonthsEmployed")
    MonthlyIncome = Get_ColumnNumber(HeaderList, "MonthlyIncome")
    ConsentIndicator = Get_ColumnNumber(HeaderList, "ConsentIndicator")

    # get column no of Is_Applicant column
    IsApp = Get_ColumnNumber(HeaderList, "Is_Applicant")
    # get value of Is_Applicant
    IsApp_Val = sheet.cell(RowNo,IsApp).value
        
    # Open schema file
    try:
        f = open(Schemafile)
    except IOError, e:
        print e
        print "Problem while opening Schema file"
        raise ErrorOccured("Problem in opening file: " + Schemafile)        # Problem opening file
    except:
        raise ErrorOccured("Unknown error occured.")                        # handle any other type of error

    # everything is fine
    sRequest = f.read()         # get contenets of schema file in to string
    f.close()

    if IsApp_Val != 'Y':
        sRequest = sRequest.replace("""<Applicant xmlns:a="http://schemas.datacontract.org/2004/07/DealerTrack.DataContracts.PreQualification">
    <a:FirstName></a:FirstName>
    <a:MiddleInitial></a:MiddleInitial>
    <a:LastName></a:LastName>
    <a:SSN></a:SSN>
    <a:AddressLine1></a:AddressLine1>
    <a:AddressLine2></a:AddressLine2>
    <a:City></a:City>
    <a:State></a:State>
    <a:ZipCode></a:ZipCode>
    <a:TotalMonthsEmployed></a:TotalMonthsEmployed>
    <a:MonthlyIncome></a:MonthlyIncome>
  </Applicant>""","",1)
        print "*** Request without Applicant Section ***"
        print sRequest
    
    # Create request string
    # Calling function to add values to nodes in given schema file
    Cell_Val = sheet.cell(RowNo,PartnerId).value
    sRequest = SetXML(sRequest, "PartnerId", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,PartnerDealerId).value
    sRequest = SetXML(sRequest, "a:string",Cell_Val.encode('ascii','ignore') , 0)

    if IsApp_Val == 'Y':
        Cell_Val = sheet.cell(RowNo,FirstName).value
        sRequest = SetXML(sRequest, "a:FirstName", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,MiddleInitial).value
        sRequest = SetXML(sRequest, "a:MiddleInitial", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,LastName).value
        sRequest = SetXML(sRequest, "a:LastName", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,SSN).value
        sRequest = SetXML(sRequest, "a:SSN", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,AddressLine1).value
        sRequest = SetXML(sRequest, "a:AddressLine1", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,AddressLine2).value
        sRequest = SetXML(sRequest, "a:AddressLine2",Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,City).value
        sRequest = SetXML(sRequest, "a:City",Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,State).value
        sRequest = SetXML(sRequest, "a:State",Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,ZipCode).value
        sRequest = SetXML(sRequest, "a:ZipCode", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,TotalMonthsEmployed).value
        sRequest = SetXML(sRequest, "a:TotalMonthsEmployed", Cell_Val.encode('ascii','ignore'), 0)
    
        Cell_Val = sheet.cell(RowNo,MonthlyIncome).value 
        sRequest = SetXML(sRequest, "a:MonthlyIncome", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,ConsentIndicator).value
    sRequest = SetXML(sRequest, "ConsentIndicator", Cell_Val.encode('ascii','ignore'), 0)
    
   #print sRequest
    
    #ReqFilePath = str(FolderPath) + '\\Request\\'+ WebService +'Request'+ str(RowNo) +'.xml'
    #ReqFilePath = str(FolderPath) + "\\WebService\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    ReqFilePath = str(FolderPath) + "\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    f = open(ReqFilePath, 'w')
    f.write(sRequest)
    f.close()
    
    return sRequest                     # return request string
     
#----------------------------------------------------------
# Description:      Function to create request XML string for FD - Lead web service for FD 1.1
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row in excel file
#                   Schemafile - File path of schema file
# Return Values:    request XML string
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Create_FD_Lead_Request_1_1(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath):
    # open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)
    
    # get column numbers from excel
    # PrimaryApplicant section
    PartnerId = Get_ColumnNumber(HeaderList, "PartnerId")   # Calling a function to get column no of "PartnerDealerId" column
    PartnerDealerId = Get_ColumnNumber(HeaderList, "PartnerDealerIds")
    FinanceMethod = Get_ColumnNumber(HeaderList, "FinanceMethod")

    # Applicant Info
    FirstName = Get_ColumnNumber(HeaderList, "FirstName")
    MiddleInitial = Get_ColumnNumber(HeaderList, "MiddleInitial")
    LastName = Get_ColumnNumber(HeaderList, "LastName")
    Suffix = Get_ColumnNumber(HeaderList, "Suffix")
    SSN = Get_ColumnNumber(HeaderList, "SSN")
    DateOfBirth = Get_ColumnNumber(HeaderList, "DateOfBirth")
    EmailAddress = Get_ColumnNumber(HeaderList, "EmailAddress")
    HomePhone = Get_ColumnNumber(HeaderList, "HomePhone")
    AddressLine1 = Get_ColumnNumber(HeaderList, "AddressLine1")
    AddressLine2 = Get_ColumnNumber(HeaderList, "AddressLine2")
    City = Get_ColumnNumber(HeaderList, "City")
    State = Get_ColumnNumber(HeaderList, "State")
    ZipCode = Get_ColumnNumber(HeaderList, "ZipCode")
    
    HousingStatus = Get_ColumnNumber(HeaderList, "HousingStatus")
    MortgageOrRent = Get_ColumnNumber(HeaderList, "MortgageOrRent")
    TotalMonthsAtAddress = Get_ColumnNumber(HeaderList, "TotalMonthsAtAddress")
    TotalMonthsEmployed = Get_ColumnNumber(HeaderList, "TotalMonthsEmployed")

    MonthlyIncome = Get_ColumnNumber(HeaderList, "MonthlyIncome")
    OtherMonthlyIncome = Get_ColumnNumber(HeaderList, "OtherMonthlyIncome")
    OtherIncomeSource = Get_ColumnNumber(HeaderList, "OtherIncomeSource")

    # CoApp info
    CoApp_FirstName = Get_ColumnNumber(HeaderList, "Co_App_FirstName")
    CoApp_MiddleInitial = Get_ColumnNumber(HeaderList, "Co_App_MiddleInitial")
    CoApp_LastName = Get_ColumnNumber(HeaderList, "Co_App_LastName")
    CoApp_Suffix = Get_ColumnNumber(HeaderList, "Co_App_Suffix")
    CoApp_SSN = Get_ColumnNumber(HeaderList, "Co_App_SSN")
    CoApp_DateOfBirth = Get_ColumnNumber(HeaderList, "Co_App_DateOfBirth")
    CoApp_EmailAddress = Get_ColumnNumber(HeaderList, "Co_App_EmailAddress")
    CoApp_HomePhone = Get_ColumnNumber(HeaderList, "Co_App_HomePhone")
    CoApp_AddressLine1 = Get_ColumnNumber(HeaderList, "Co_App_AddressLine1")
    CoApp_AddressLine2 = Get_ColumnNumber(HeaderList, "Co_App_AddressLine2")
    CoApp_City = Get_ColumnNumber(HeaderList, "Co_App_City")
    CoApp_State = Get_ColumnNumber(HeaderList, "Co_App_State")
    CoApp_ZipCode = Get_ColumnNumber(HeaderList, "Co_App_ZipCode")
    
    CoApp_HousingStatus = Get_ColumnNumber(HeaderList, "Co_App_HousingStatus")
    CoApp_MortgageOrRent = Get_ColumnNumber(HeaderList, "Co_App_MortgageOrRent")
    CoApp_TotalMonthsAtAddress = Get_ColumnNumber(HeaderList, "Co_App_TotalMonthsAtAddress")
    CoApp_TotalMonthsEmployed = Get_ColumnNumber(HeaderList, "Co_App_TotalMonthsEmployed")

    CoApp_MonthlyIncome = Get_ColumnNumber(HeaderList, "Co_App_MonthlyIncome")
    CoApp_OtherMonthlyIncome = Get_ColumnNumber(HeaderList, "Co_App_OtherMonthlyIncome")
    CoApp_OtherIncomeSource = Get_ColumnNumber(HeaderList, "Co_App_OtherIncomeSource")

    
    # VehicleInfo section
    VehicleCondition = Get_ColumnNumber(HeaderList, "VehicleCondition")
    Year = Get_ColumnNumber(HeaderList, "Year")
    Make = Get_ColumnNumber(HeaderList, "Make")
    Model = Get_ColumnNumber(HeaderList, "Model")
    Trim = Get_ColumnNumber(HeaderList, "Trim")
    ChromeStyleId = Get_ColumnNumber(HeaderList, "ChromeStyleId")
    VIN = Get_ColumnNumber(HeaderList, "VIN")
    StockNumber = Get_ColumnNumber(HeaderList, "StockNumber")
    CertifiedUsed = Get_ColumnNumber(HeaderList, "CertifiedUsed")
    InteriorColor = Get_ColumnNumber(HeaderList, "InteriorColor")
    ExteriorColor = Get_ColumnNumber(HeaderList, "ExteriorColor")
    Mileage = Get_ColumnNumber(HeaderList, "Mileage")

    # LoanInfo section
    CashDown = Get_ColumnNumber(HeaderList, "CashDown")
    Term = Get_ColumnNumber(HeaderList, "Term")
    AmountRequesting = Get_ColumnNumber(HeaderList, "AmountRequesting")
    MonthlyPayment = Get_ColumnNumber(HeaderList, "MonthlyPayment")
    CreditType = Get_ColumnNumber(HeaderList, "CreditType")
    ApplicantType = Get_ColumnNumber(HeaderList, "ApplicantType")

    # TradeInVehicleInfo section
    Trade_Year = Get_ColumnNumber(HeaderList, "Trade_Year")
    Trade_Make = Get_ColumnNumber(HeaderList, "Trade_Make")
    Trade_Model = Get_ColumnNumber(HeaderList, "Trade_Model")
    Trade_Trim = Get_ColumnNumber(HeaderList, "Trade_Trim")
    Trade_Mileage = Get_ColumnNumber(HeaderList, "Trade_Mileage")
    Trade_ChromeStyleId = Get_ColumnNumber(HeaderList, "Trade_ChromeStyleId")
    Trade_TradeInPaid = Get_ColumnNumber(HeaderList, "Trade_TradeInPaid")
    
    Comments = Get_ColumnNumber(HeaderList, "Comments")
    PrequalificationReferenceNumber = Get_ColumnNumber(HeaderList, "PrequalificationReferenceNumber")

    # get column nos if sections present
    Is_Applicant = Get_ColumnNumber(HeaderList, "Is_Applicant")
    Is_CoApp = Get_ColumnNumber(HeaderList, "Is_CoApp")
    Is_Vehicle = Get_ColumnNumber(HeaderList, "Is_Vehicle")
    Is_Loan = Get_ColumnNumber(HeaderList, "Is_Loan")
    Is_Tradein = Get_ColumnNumber(HeaderList, "Is_Tradein")
    
    # Open schema file
    try:
        f = open(Schemafile)
    except IOError, e:
        print e
        print "Problem while opening Schema file"
        raise ErrorOccured("Problem in opening file: " + Schemafile)        # Problem opening file
    except:
        raise ErrorOccured("Unknown error occured.")                        # handle any other type of error

    # everything is fine
    sRequest = f.read()         # get contenets of schema file in to string
    f.close()

    Is_Applicant_Val = sheet.cell(RowNo,Is_Applicant).value
    Is_CoApp_Val = sheet.cell(RowNo,Is_CoApp).value
    Is_Vehicle_Val = sheet.cell(RowNo,Is_Vehicle).value
    Is_Loan_Val = sheet.cell(RowNo,Is_Loan).value
    Is_Tradein_Val = sheet.cell(RowNo,Is_Tradein).value
    

    # Change schema if fields are blank - only
    # partner ID
    Cell_Val = sheet.cell(RowNo,PartnerId).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<PartnerId></PartnerId>","""<PartnerId  i:nil="true" />""",1)

    # partner DEaler ID
    Cell_Val = sheet.cell(RowNo,PartnerDealerId).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<PartnerDealerId></PartnerDealerId>","""<PartnerDealerId i:nil="true" />""",1)

    # Finance Method
    Cell_Val = sheet.cell(RowNo,FinanceMethod).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<FinanceMethod></FinanceMethod>","""<FinanceMethod i:nil="true" />""",1)

    # PrequalificationReferenceNumber
    Cell_Val = sheet.cell(RowNo,PrequalificationReferenceNumber).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<PrequalificationReferenceNumber></PrequalificationReferenceNumber>","""<PrequalificationReferenceNumber i:nil="true" />""",1)  

    
    # Applicant info
    if Is_Applicant_Val == 'Y' and Is_CoApp_Val == 'Y':
        # First name
        FN_Val = sheet.cell(RowNo,FirstName).value
        if len(FN_Val) == 0:
            sRequest = sRequest.replace("<a:FirstName></a:FirstName>","""<a:FirstName i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,CoApp_FirstName).value
        if len(Cell_Val) == 0 and len(FN_Val)!=0 :
            sRequest = replacenth(sRequest,"<a:FirstName></a:FirstName>","""<a:FirstName  i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(FN_Val)==0 :
            sRequest = replacenth(sRequest,"<a:FirstName></a:FirstName>","""<a:FirstName  i:nil="true" />""",1)

        # middle initial
        MI_Val = sheet.cell(RowNo,MiddleInitial).value
        if len(MI_Val) == 0:
            sRequest = sRequest.replace("<a:MiddleInitial></a:MiddleInitial>","""<a:MiddleInitial i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,CoApp_MiddleInitial).value
        if len(Cell_Val) == 0 and len(MI_Val)!=0:
            sRequest = replacenth(sRequest,"<a:MiddleInitial></a:MiddleInitial>","""<a:MiddleInitial  i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(MI_Val)==0:
            sRequest = replacenth(sRequest,"<a:MiddleInitial></a:MiddleInitial>","""<a:MiddleInitial  i:nil="true" />""",1)


        # last name
        LN_Val = sheet.cell(RowNo,LastName).value
        if len(LN_Val) == 0: 
            sRequest = sRequest.replace("<a:LastName></a:LastName>","""<a:LastName i:nil="true" />""",1)
    
        Cell_Val = sheet.cell(RowNo,CoApp_LastName).value
        if len(Cell_Val) == 0 and len(LN_Val)!=0:
            sRequest = replacenth(sRequest,"<a:LastName></a:LastName>","""<a:LastName i:nil="true" />""",2)
        elif len(LN_Val) == 0 and len(LN_Val)==0: 
            sRequest = sRequest.replace("<a:LastName></a:LastName>","""<a:LastName i:nil="true" />""",1)


        # Suffix
        sf_Val = sheet.cell(RowNo,Suffix).value
        if len(sf_Val) == 0:
            sRequest = sRequest.replace("<a:Suffix></a:Suffix>","""<a:Suffix i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,CoApp_Suffix).value
        if len(Cell_Val) == 0 and len(sf_Val)!=0:
            sRequest = replacenth(sRequest,"<a:Suffix></a:Suffix>","""<a:Suffix i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(sf_Val)==0:
            sRequest = replacenth(sRequest,"<a:Suffix></a:Suffix>","""<a:Suffix i:nil="true" />""",1)

        # SSN
        ssn_Val = sheet.cell(RowNo,SSN).value
        if len(ssn_Val) == 0:
            sRequest = sRequest.replace("<a:SSN></a:SSN>","""<a:SSN i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,CoApp_SSN).value
        if len(Cell_Val) == 0 and len(ssn_Val)!=0:
            sRequest = replacenth(sRequest,"<a:SSN></a:SSN>","""<a:SSN i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(ssn_Val)==0:
            sRequest = replacenth(sRequest,"<a:SSN></a:SSN>","""<a:SSN i:nil="true" />""",1)


        # DateOfBirth
        dob_Val = sheet.cell(RowNo,DateOfBirth).value
        if len(dob_Val) == 0:
            sRequest = sRequest.replace("<a:DateOfBirth></a:DateOfBirth>","""<a:DateOfBirth i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,CoApp_DateOfBirth).value
        if len(Cell_Val) == 0 and len(dob_Val)!=0:
            sRequest = replacenth(sRequest,"<a:DateOfBirth></a:DateOfBirth>","""<a:DateOfBirth i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(dob_Val)==0:
            sRequest = replacenth(sRequest,"<a:DateOfBirth></a:DateOfBirth>","""<a:DateOfBirth i:nil="true" />""",1)

        # EmailAddress
        EM_Val = sheet.cell(RowNo,EmailAddress).value
        if len(EM_Val) == 0:
            sRequest = sRequest.replace("<a:EmailAddress></a:EmailAddress>","""<a:EmailAddress i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,CoApp_EmailAddress).value
        if len(Cell_Val) == 0 and len(EM_Val)!=0:
            sRequest = replacenth(sRequest,"<a:EmailAddress></a:EmailAddress>","""<a:EmailAddress i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(EM_Val)==0:
            sRequest = replacenth(sRequest,"<a:EmailAddress></a:EmailAddress>","""<a:EmailAddress i:nil="true" />""",1)

        # HomePhone
        hp_Val = sheet.cell(RowNo,HomePhone).value
        if len(hp_Val) == 0:
            sRequest = sRequest.replace("<a:HomePhone></a:HomePhone>","""<a:HomePhone i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,CoApp_HomePhone).value
        if len(Cell_Val) == 0 and len(hp_Val)!=0:
            sRequest = replacenth(sRequest,"<a:HomePhone></a:HomePhone>","""<a:HomePhone i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(hp_Val)==0:
            sRequest = replacenth(sRequest,"<a:HomePhone></a:HomePhone>","""<a:HomePhone i:nil="true" />""",1)
        
        # AddressLine1
        ad1_Val = sheet.cell(RowNo,AddressLine1).value
        if len(ad1_Val) == 0:
            sRequest = sRequest.replace("<a:AddressLine1></a:AddressLine1>","""<a:AddressLine1 i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,CoApp_AddressLine1).value
        if len(Cell_Val) == 0 and len(ad1_Val)!=0:
            sRequest = replacenth(sRequest,"<a:AddressLine1></a:AddressLine1>","""<a:AddressLine1 i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(ad1_Val)==0:
            sRequest = replacenth(sRequest,"<a:AddressLine1></a:AddressLine1>","""<a:AddressLine1 i:nil="true" />""",1)

        # AddressLine2
        Ad2_Val = sheet.cell(RowNo,AddressLine2).value
        if len(Ad2_Val) == 0:
            sRequest = sRequest.replace("<a:AddressLine2></a:AddressLine2>","""<a:AddressLine2 i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,CoApp_AddressLine2).value
        if len(Cell_Val) == 0 and len(Ad2_Val)!=0:
            sRequest = replacenth(sRequest,"<a:AddressLine2></a:AddressLine2>","""<a:AddressLine2 i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(Ad2_Val)==0:
            sRequest = replacenth(sRequest,"<a:AddressLine2></a:AddressLine2>","""<a:AddressLine2 i:nil="true" />""",1)

    
        # City
        ct_Val = sheet.cell(RowNo,City).value
        if len(ct_Val) == 0:
            sRequest = sRequest.replace("<a:City></a:City>","""<a:City i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,CoApp_City).value
        if len(Cell_Val) == 0 and len(ct_Val) != 0:
            sRequest = replacenth(sRequest,"<a:City></a:City>","""<a:City i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(ct_Val) == 0:
            sRequest = replacenth(sRequest,"<a:City></a:City>","""<a:City i:nil="true" />""",1)

        # State
        st_Val = sheet.cell(RowNo,State).value
        if len(st_Val) == 0:
            sRequest = sRequest.replace("<a:State></a:State>","""<a:State i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,CoApp_State).value
        if len(Cell_Val) == 0 and len(st_Val) != 0:
            sRequest = replacenth(sRequest,"<a:State></a:State>","""<a:State i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(st_Val) == 0:
            sRequest = replacenth(sRequest,"<a:State></a:State>","""<a:State i:nil="true" />""",1)

        # ZipCode
        zp_Val = sheet.cell(RowNo,ZipCode).value
        if len(zp_Val) == 0:
            sRequest = sRequest.replace("<a:ZipCode></a:ZipCode>","""<a:ZipCode i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,CoApp_ZipCode).value
        if len(Cell_Val) == 0 and len(zp_Val) != 0:
            sRequest = replacenth(sRequest,"<a:ZipCode></a:ZipCode>","""<a:ZipCode i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(zp_Val) == 0:
            sRequest = replacenth(sRequest,"<a:ZipCode></a:ZipCode>","""<a:ZipCode i:nil="true" />""",1)

        # HousingStatus
        hs_Val = sheet.cell(RowNo,HousingStatus).value
        if len(hs_Val) == 0:
            sRequest = sRequest.replace("<a:HousingStatus></a:HousingStatus>","""<a:HousingStatus i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,CoApp_HousingStatus).value
        if len(Cell_Val) == 0 and len(hs_Val) != 0:
            sRequest = replacenth(sRequest,"<a:HousingStatus></a:HousingStatus>","""<a:HousingStatus i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(hs_Val) == 0:
            sRequest = replacenth(sRequest,"<a:HousingStatus></a:HousingStatus>","""<a:HousingStatus i:nil="true" />""",1)
    

        # MortgageOrRent
        mr_Val = sheet.cell(RowNo,MortgageOrRent).value
        if len(mr_Val) == 0:
            sRequest = sRequest.replace("<a:MortgageOrRent></a:MortgageOrRent>","""<a:MortgageOrRent i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,CoApp_MortgageOrRent).value
        if len(Cell_Val) == 0 and len(mr_Val) != 0:
            sRequest = replacenth(sRequest,"<a:MortgageOrRent></a:MortgageOrRent>","""<a:MortgageOrRent i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(mr_Val) == 0:
            sRequest = replacenth(sRequest,"<a:MortgageOrRent></a:MortgageOrRent>","""<a:MortgageOrRent i:nil="true" />""",1)
    

        # TotalMonthsAtAddress
        tma_Val = sheet.cell(RowNo,TotalMonthsAtAddress).value
        if len(tma_Val) == 0:
            sRequest = sRequest.replace("<a:TotalMonthsAtAddress></a:TotalMonthsAtAddress>","""<a:TotalMonthsAtAddress i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,CoApp_TotalMonthsAtAddress).value
        if len(Cell_Val) == 0 and len(tma_Val) != 0:
            sRequest = replacenth(sRequest,"<a:TotalMonthsAtAddress></a:TotalMonthsAtAddress>","""<a:TotalMonthsAtAddress i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(tma_Val) == 0:
            sRequest = replacenth(sRequest,"<a:TotalMonthsAtAddress></a:TotalMonthsAtAddress>","""<a:TotalMonthsAtAddress i:nil="true" />""",1)

        # TotalMonthsEmployed
        tme_Val = sheet.cell(RowNo,TotalMonthsEmployed).value
        if len(tme_Val) == 0:
            sRequest = sRequest.replace("<a:TotalMonthsEmployed></a:TotalMonthsEmployed>","""<a:TotalMonthsEmployed i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,CoApp_TotalMonthsEmployed).value
        if len(Cell_Val) == 0 and len(tme_Val) != 0:
            sRequest = replacenth(sRequest,"<a:TotalMonthsEmployed></a:TotalMonthsEmployed>","""<a:TotalMonthsEmployed i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(tme_Val) == 0:
            sRequest = replacenth(sRequest,"<a:TotalMonthsEmployed></a:TotalMonthsEmployed>","""<a:TotalMonthsEmployed i:nil="true" />""",1)

        # MonthlyIncome
        mi_Val = sheet.cell(RowNo,MonthlyIncome).value
        if len(mi_Val) == 0:
            sRequest = sRequest.replace("<a:MonthlyIncome></a:MonthlyIncome>","""<a:MonthlyIncome i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,CoApp_MonthlyIncome).value
        if len(Cell_Val) == 0 and len(mi_Val) != 0:
            sRequest = replacenth(sRequest,"<a:MonthlyIncome></a:MonthlyIncome>","""<a:MonthlyIncome i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(mi_Val) == 0:
            sRequest = replacenth(sRequest,"<a:MonthlyIncome></a:MonthlyIncome>","""<a:MonthlyIncome i:nil="true" />""",1)

        # OtherMonthlyIncome
        omi_Val = sheet.cell(RowNo,OtherMonthlyIncome).value
        if len(omi_Val) == 0:
            sRequest = sRequest.replace("<a:OtherMonthlyIncome></a:OtherMonthlyIncome>","""<a:OtherMonthlyIncome i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,CoApp_OtherMonthlyIncome).value
        if len(Cell_Val) == 0 and len(omi_Val) != 0:
            sRequest = replacenth(sRequest,"<a:OtherMonthlyIncome></a:OtherMonthlyIncome>","""<a:OtherMonthlyIncome i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(omi_Val) == 0:
            sRequest = replacenth(sRequest,"<a:OtherMonthlyIncome></a:OtherMonthlyIncome>","""<a:OtherMonthlyIncome i:nil="true" />""",1)

        # OtherIncomeSource
        ois_Val = sheet.cell(RowNo,OtherIncomeSource).value
        if len(ois_Val) == 0:
            sRequest = sRequest.replace("<a:OtherIncomeSource></a:OtherIncomeSource>","""<a:OtherIncomeSource i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,CoApp_OtherIncomeSource).value
        if len(Cell_Val) == 0 and len(ois_Val) != 0:
            sRequest = replacenth(sRequest,"<a:OtherIncomeSource></a:OtherIncomeSource>","""<a:OtherIncomeSource i:nil="true" />""",2)
        if len(Cell_Val) == 0 and len(ois_Val) == 0:
            sRequest = replacenth(sRequest,"<a:OtherIncomeSource></a:OtherIncomeSource>","""<a:OtherIncomeSource i:nil="true" />""",1)

    # if No Applicant and only Co-App
    elif Is_Applicant_Val != 'Y' and Is_CoApp_Val == 'Y':
        sRequest = sRequest.replace("""<PrimaryApplicant xmlns:a="http://schemas.datacontract.org/2004/07/DealerTrack.DataContracts.Lead">       
    <a:FirstName></a:FirstName>
    <a:MiddleInitial></a:MiddleInitial>
    <a:LastName></a:LastName>
    <a:Suffix></a:Suffix>
    <a:EmailAddress></a:EmailAddress>
    <a:SSN></a:SSN>
    <a:DateOfBirth></a:DateOfBirth>
    <a:HomePhone></a:HomePhone>
    <a:AddressLine1></a:AddressLine1>
    <a:AddressLine2></a:AddressLine2>
    <a:City></a:City>
    <a:State></a:State>
    <a:ZipCode></a:ZipCode>
    <a:HousingStatus></a:HousingStatus>
    <a:MortgageOrRent></a:MortgageOrRent>
    <a:TotalMonthsAtAddress></a:TotalMonthsAtAddress>
    <a:TotalMonthsEmployed></a:TotalMonthsEmployed>
    <a:MonthlyIncome></a:MonthlyIncome>
    <a:OtherMonthlyIncome></a:OtherMonthlyIncome>
    <a:OtherIncomeSource></a:OtherIncomeSource>
 </PrimaryApplicant>""","",1)

    # if only Applicant and no Co-App
    elif Is_Applicant_Val == 'Y' and Is_CoApp_Val != 'Y':
        sRequest = sRequest.replace("""<CoApplicant xmlns:a="http://schemas.datacontract.org/2004/07/DealerTrack.DataContracts.Lead">
    <a:FirstName></a:FirstName>
    <a:MiddleInitial></a:MiddleInitial>
    <a:LastName></a:LastName>
    <a:Suffix></a:Suffix>
    <a:EmailAddress></a:EmailAddress>
    <a:SSN></a:SSN>
    <a:DateOfBirth></a:DateOfBirth>
    <a:HomePhone></a:HomePhone>
    <a:AddressLine1></a:AddressLine1>
    <a:AddressLine2></a:AddressLine2>
    <a:City></a:City>
    <a:State></a:State>
    <a:ZipCode></a:ZipCode>
    <a:HousingStatus></a:HousingStatus>
    <a:MortgageOrRent></a:MortgageOrRent>
    <a:TotalMonthsAtAddress></a:TotalMonthsAtAddress>
    <a:TotalMonthsEmployed></a:TotalMonthsEmployed>
    <a:MonthlyIncome></a:MonthlyIncome> 
    <a:OtherMonthlyIncome></a:OtherMonthlyIncome>
    <a:OtherIncomeSource></a:OtherIncomeSource>
 </CoApplicant>""","",1)

    # if only Applicant and no Co-App
    elif Is_Applicant_Val != 'Y' and Is_CoApp_Val != 'Y':
        sRequest = sRequest.replace("""<PrimaryApplicant xmlns:a="http://schemas.datacontract.org/2004/07/DealerTrack.DataContracts.Lead">       
    <a:FirstName></a:FirstName>
    <a:MiddleInitial></a:MiddleInitial>
    <a:LastName></a:LastName>
    <a:Suffix></a:Suffix>
    <a:EmailAddress></a:EmailAddress>
    <a:SSN></a:SSN>
    <a:DateOfBirth></a:DateOfBirth>
    <a:HomePhone></a:HomePhone>
    <a:AddressLine1></a:AddressLine1>
    <a:AddressLine2></a:AddressLine2>
    <a:City></a:City>
    <a:State></a:State>
    <a:ZipCode></a:ZipCode>
    <a:HousingStatus></a:HousingStatus>
    <a:MortgageOrRent></a:MortgageOrRent>
    <a:TotalMonthsAtAddress></a:TotalMonthsAtAddress>
    <a:TotalMonthsEmployed></a:TotalMonthsEmployed>
    <a:MonthlyIncome></a:MonthlyIncome>
    <a:OtherMonthlyIncome></a:OtherMonthlyIncome>
    <a:OtherIncomeSource></a:OtherIncomeSource>
 </PrimaryApplicant>
 <CoApplicant xmlns:a="http://schemas.datacontract.org/2004/07/DealerTrack.DataContracts.Lead">
    <a:FirstName></a:FirstName>
    <a:MiddleInitial></a:MiddleInitial>
    <a:LastName></a:LastName>
    <a:Suffix></a:Suffix>
    <a:EmailAddress></a:EmailAddress>
    <a:SSN></a:SSN>
    <a:DateOfBirth></a:DateOfBirth>
    <a:HomePhone></a:HomePhone>
    <a:AddressLine1></a:AddressLine1>
    <a:AddressLine2></a:AddressLine2>
    <a:City></a:City>
    <a:State></a:State>
    <a:ZipCode></a:ZipCode>
    <a:HousingStatus></a:HousingStatus>
    <a:MortgageOrRent></a:MortgageOrRent>
    <a:TotalMonthsAtAddress></a:TotalMonthsAtAddress>
    <a:TotalMonthsEmployed></a:TotalMonthsEmployed>
    <a:MonthlyIncome></a:MonthlyIncome> 
    <a:OtherMonthlyIncome></a:OtherMonthlyIncome>
    <a:OtherIncomeSource></a:OtherIncomeSource>
 </CoApplicant>""","",1)
    #print sRequest
    
    # vehicle info
    if Is_Vehicle_Val == 'Y' and Is_Tradein_Val == 'Y':
        Cell_Val = sheet.cell(RowNo,VehicleCondition).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:VehicleCondition></a:VehicleCondition>","""<a:VehicleCondition i:nil="true" />""",1)
   
        # year
        yr_Val = sheet.cell(RowNo,Year).value
        if len(yr_Val) == 0:
            sRequest = sRequest.replace("<a:Year></a:Year>","""<a:Year i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,Trade_Year).value
        if len(Cell_Val) == 0 and len(yr_Val) != 0:
            sRequest = sRequest.replace("<a:Year></a:Year>","""<a:Year i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(yr_Val) == 0:
            sRequest = sRequest.replace("<a:Year></a:Year>","""<a:Year i:nil="true" />""",1)

        #make
        mk_Val = sheet.cell(RowNo,Make).value
        if len(mk_Val) == 0:
            sRequest = sRequest.replace("<a:Make></a:Make>","""<a:Make i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,Trade_Make).value
        if len(Cell_Val) == 0 and len(mk_Val) != 0:
            sRequest = sRequest.replace("<a:Make></a:Make>","""<a:Make i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(mk_Val) == 0:
            sRequest = sRequest.replace("<a:Make></a:Make>","""<a:Make i:nil="true" />""",1)

        #model
        md_Val = sheet.cell(RowNo,Model).value
        if len(md_Val) == 0:
            sRequest = sRequest.replace("<a:Model></a:Model>","""<a:Model i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,Trade_Model).value
        if len(Cell_Val) == 0 and len(md_Val) != 0:
            sRequest = sRequest.replace("<a:Model></a:Model>","""<a:Model i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(md_Val) == 0:
            sRequest = sRequest.replace("<a:Model></a:Model>","""<a:Model i:nil="true" />""",1)

        # Trim
        trm_Val = sheet.cell(RowNo,Trim).value
        if len(trm_Val) == 0:
            sRequest = sRequest.replace("<a:Trim></a:Trim>","""<a:Trim i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,Trade_Trim).value
        if len(Cell_Val) == 0 and len(trm_Val) != 0:
            sRequest = sRequest.replace("<a:Trim></a:Trim>","""<a:Trim i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(trm_Val) == 0:
            sRequest = sRequest.replace("<a:Trim></a:Trim>","""<a:Trim i:nil="true" />""",1)

        # chrome style id
        crt_Val = sheet.cell(RowNo,ChromeStyleId).value
        if len(crt_Val) == 0:
            sRequest = sRequest.replace("<a:ChromeStyleId></a:ChromeStyleId>","""<a:ChromeStyleId i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,Trade_ChromeStyleId).value
        if len(Cell_Val) == 0 and len(crt_Val) != 0:
            sRequest = sRequest.replace("<a:ChromeStyleId></a:ChromeStyleId>","""<a:ChromeStyleId i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(crt_Val) == 0:
            sRequest = sRequest.replace("<a:ChromeStyleId></a:ChromeStyleId>","""<a:ChromeStyleId i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,VIN).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:VIN></a:VIN>","""<a:VIN i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,StockNumber).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:StockNumber></a:StockNumber>","""<a:StockNumber i:nil="true" />""",1)

    
        Cell_Val = sheet.cell(RowNo,CertifiedUsed).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:CertifiedUsed></a:CertifiedUsed>","""<a:CertifiedUsed i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,InteriorColor).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:InteriorColor></a:InteriorColor>","""<a:InteriorColor i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,ExteriorColor).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:ExteriorColor></a:ExteriorColor>","""<a:ExteriorColor i:nil="true" />""",1)

        # Mileage
        mlg_Val = sheet.cell(RowNo,Mileage).value
        if len(mlg_Val) == 0:
            sRequest = sRequest.replace("<a:Mileage></a:Mileage>","""<a:Mileage i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,Trade_Mileage).value
        if len(Cell_Val) == 0 and len(mlg_Val) != 0:
            sRequest = sRequest.replace("<a:Mileage></a:Mileage>","""<a:Mileage i:nil="true" />""",2)
        elif len(Cell_Val) == 0 and len(mlg_Val) == 0:
            sRequest = sRequest.replace("<a:Mileage></a:Mileage>","""<a:Mileage i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,Trade_TradeInPaid).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:TradeInPaid></a:TradeInPaid>","""<a:TradeInPaid i:nil="true" />""",1)

    #print sRequest
    
    # if No Vehicle section and trade in section is present
    if Is_Vehicle_Val != 'Y':
        sRequest = sRequest.replace("""<VehicleInfo xmlns:a="http://schemas.datacontract.org/2004/07/DealerTrack.DataContracts.Lead">
    <a:VehicleCondition></a:VehicleCondition>
    <a:Year></a:Year>
    <a:Make></a:Make>
    <a:Model></a:Model>
    <a:Trim></a:Trim>
    <a:ChromeStyleId></a:ChromeStyleId>
    <a:VIN></a:VIN>
    <a:StockNumber></a:StockNumber>
    <a:CertifiedUsed></a:CertifiedUsed>
    <a:InteriorColor></a:InteriorColor>
    <a:ExteriorColor></a:ExteriorColor>
    <a:Mileage></a:Mileage>
  </VehicleInfo>""","",1)

    # if Vehicle section is present and trade in section is NOT present
    if Is_Tradein_Val != 'Y':
        sRequest = sRequest.replace("""<TradeInVehicleInfo xmlns:a="http://schemas.datacontract.org/2004/07/DealerTrack.DataContracts.Lead">
    <a:Year></a:Year>
    <a:Make></a:Make>
    <a:Model></a:Model>
    <a:Trim></a:Trim>
    <a:Mileage></a:Mileage>
    <a:ChromeStyleId></a:ChromeStyleId>
    <a:TradeInPaid></a:TradeInPaid>
  </TradeInVehicleInfo>""","",1)

    #print sRequest
    
    # loan info
    if Is_Loan_Val == 'Y':
        Cell_Val = sheet.cell(RowNo,CashDown).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:CashDown></a:CashDown>","""<a:CashDown i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,Term).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:Term></a:Term>","""<a:Term i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,AmountRequesting).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:AmountRequesting></a:AmountRequesting>","""<a:AmountRequesting i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,MonthlyPayment).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:MonthlyPayment></a:MonthlyPayment>","""<a:MonthlyPayment i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,CreditType).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:CreditType></a:CreditType>","""<a:CreditType i:nil="true" />""",1)


        Cell_Val = sheet.cell(RowNo,ApplicantType).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:ApplicantType></a:ApplicantType>","""<a:ApplicantType i:nil="true" />""",1)
    else:
        sRequest = sRequest.replace("""<LoanInfo xmlns:a="http://schemas.datacontract.org/2004/07/DealerTrack.DataContracts.Lead">
   <a:CashDown></a:CashDown>
   <a:Term></a:Term>
   <a:AmountRequesting></a:AmountRequesting>
   <a:MonthlyPayment></a:MonthlyPayment>
   <a:CreditType></a:CreditType>
   <a:ApplicantType></a:ApplicantType>
  </LoanInfo>""","",1)
    
    Cell_Val = sheet.cell(RowNo,Comments).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<Comments></Comments>","""<Comments i:nil="true" />""",1)

    #print sRequest

    # *******************************************************************************************
    # Create request string
    # Calling function to add values to nodes in given schema file
    Cell_Val = sheet.cell(RowNo,PartnerId).value
    sRequest = SetXML(sRequest, "PartnerId", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,PartnerDealerId).value
    sRequest = SetXML(sRequest, "PartnerDealerId",Cell_Val.encode('ascii','ignore') , 0)

    Cell_Val = sheet.cell(RowNo,FinanceMethod).value
    sRequest = SetXML(sRequest, "FinanceMethod",Cell_Val.encode('ascii','ignore') , 0)


    if Is_Applicant_Val == 'Y':
        # Applicant info
    
        Cell_Val = sheet.cell(RowNo,FirstName).value
        sRequest = SetXML(sRequest, "a:FirstName", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,MiddleInitial).value
        sRequest = SetXML(sRequest, "a:MiddleInitial", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,LastName).value
        sRequest = SetXML(sRequest, "a:LastName", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Suffix).value
        sRequest = SetXML(sRequest, "a:Suffix", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,SSN).value
        sRequest = SetXML(sRequest, "a:SSN", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,DateOfBirth).value
        sRequest = SetXML(sRequest, "a:DateOfBirth", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,EmailAddress).value
        sRequest = SetXML(sRequest, "a:EmailAddress", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,HomePhone).value
        sRequest = SetXML(sRequest, "a:HomePhone", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,AddressLine1).value
        sRequest = SetXML(sRequest, "a:AddressLine1", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,AddressLine2).value
        sRequest = SetXML(sRequest, "a:AddressLine2",Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,City).value
        sRequest = SetXML(sRequest, "a:City",Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,State).value
        sRequest = SetXML(sRequest, "a:State",Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,ZipCode).value
        sRequest = SetXML(sRequest, "a:ZipCode", Cell_Val.encode('ascii','ignore'), 0)

        # changes done on 19 March - Additional fields
        Cell_Val = sheet.cell(RowNo,HousingStatus).value
        sRequest = SetXML(sRequest, "a:HousingStatus", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,MortgageOrRent).value
        sRequest = SetXML(sRequest, "a:MortgageOrRent", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,TotalMonthsAtAddress).value
        sRequest = SetXML(sRequest, "a:TotalMonthsAtAddress", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,TotalMonthsEmployed).value
        sRequest = SetXML(sRequest, "a:TotalMonthsEmployed", Cell_Val.encode('ascii','ignore'), 0)
    
        Cell_Val = sheet.cell(RowNo,MonthlyIncome).value
        sRequest = SetXML(sRequest, "a:MonthlyIncome", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,OtherMonthlyIncome).value
        sRequest = SetXML(sRequest, "a:OtherMonthlyIncome", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,OtherIncomeSource).value
        sRequest = SetXML(sRequest, "a:OtherIncomeSource", Cell_Val.encode('ascii','ignore'), 0)

        if Is_CoApp_Val == 'Y':
            # Co-Applicant info
            Cell_Val = sheet.cell(RowNo,CoApp_FirstName).value
            sRequest = SetXML(sRequest, "a:FirstName", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,CoApp_MiddleInitial).value
            sRequest = SetXML(sRequest, "a:MiddleInitial", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,CoApp_LastName).value
            sRequest = SetXML(sRequest, "a:LastName", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,CoApp_Suffix).value
            sRequest = SetXML(sRequest, "a:Suffix", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,CoApp_SSN).value
            sRequest = SetXML(sRequest, "a:SSN", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,CoApp_DateOfBirth).value
            sRequest = SetXML(sRequest, "a:DateOfBirth", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,CoApp_EmailAddress).value
            sRequest = SetXML(sRequest, "a:EmailAddress", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,CoApp_HomePhone).value
            sRequest = SetXML(sRequest, "a:HomePhone", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,CoApp_AddressLine1).value
            sRequest = SetXML(sRequest, "a:AddressLine1", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,CoApp_AddressLine2).value
            sRequest = SetXML(sRequest, "a:AddressLine2",Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,CoApp_City).value
            sRequest = SetXML(sRequest, "a:City",Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,CoApp_State).value
            sRequest = SetXML(sRequest, "a:State",Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,CoApp_ZipCode).value
            sRequest = SetXML(sRequest, "a:ZipCode", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,CoApp_HousingStatus).value
            sRequest = SetXML(sRequest, "a:HousingStatus", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,CoApp_MortgageOrRent).value
            sRequest = SetXML(sRequest, "a:MortgageOrRent", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,CoApp_TotalMonthsAtAddress).value
            sRequest = SetXML(sRequest, "a:TotalMonthsAtAddress", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,CoApp_TotalMonthsEmployed).value
            sRequest = SetXML(sRequest, "a:TotalMonthsEmployed", Cell_Val.encode('ascii','ignore'), 1)
    
            Cell_Val = sheet.cell(RowNo,CoApp_MonthlyIncome).value
            sRequest = SetXML(sRequest, "a:MonthlyIncome", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,CoApp_OtherMonthlyIncome).value
            sRequest = SetXML(sRequest, "a:OtherMonthlyIncome", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,CoApp_OtherIncomeSource).value
            sRequest = SetXML(sRequest, "a:OtherIncomeSource", Cell_Val.encode('ascii','ignore'), 1)
    # if Applicant section is not present
    else:
        # Co-Applicant info
        Cell_Val = sheet.cell(RowNo,CoApp_FirstName).value
        sRequest = SetXML(sRequest, "a:FirstName", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,CoApp_MiddleInitial).value
        sRequest = SetXML(sRequest, "a:MiddleInitial", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,CoApp_LastName).value
        sRequest = SetXML(sRequest, "a:LastName", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,CoApp_Suffix).value
        sRequest = SetXML(sRequest, "a:Suffix", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,CoApp_SSN).value
        sRequest = SetXML(sRequest, "a:SSN", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,CoApp_DateOfBirth).value
        sRequest = SetXML(sRequest, "a:DateOfBirth", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,CoApp_EmailAddress).value
        sRequest = SetXML(sRequest, "a:EmailAddress", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,CoApp_HomePhone).value
        sRequest = SetXML(sRequest, "a:HomePhone", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,CoApp_AddressLine1).value
        sRequest = SetXML(sRequest, "a:AddressLine1", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,CoApp_AddressLine2).value
        sRequest = SetXML(sRequest, "a:AddressLine2",Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,CoApp_City).value
        sRequest = SetXML(sRequest, "a:City",Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,CoApp_State).value
        sRequest = SetXML(sRequest, "a:State",Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,CoApp_ZipCode).value
        sRequest = SetXML(sRequest, "a:ZipCode", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,CoApp_HousingStatus).value
        sRequest = SetXML(sRequest, "a:HousingStatus", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,CoApp_MortgageOrRent).value
        sRequest = SetXML(sRequest, "a:MortgageOrRent", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,CoApp_TotalMonthsAtAddress).value
        sRequest = SetXML(sRequest, "a:TotalMonthsAtAddress", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,CoApp_TotalMonthsEmployed).value
        sRequest = SetXML(sRequest, "a:TotalMonthsEmployed", Cell_Val.encode('ascii','ignore'), 0)
    
        Cell_Val = sheet.cell(RowNo,CoApp_MonthlyIncome).value
        sRequest = SetXML(sRequest, "a:MonthlyIncome", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,CoApp_OtherMonthlyIncome).value
        sRequest = SetXML(sRequest, "a:OtherMonthlyIncome", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,CoApp_OtherIncomeSource).value
        sRequest = SetXML(sRequest, "a:OtherIncomeSource", Cell_Val.encode('ascii','ignore'), 0)

    if Is_Vehicle_Val == 'Y':
        # vehicle info
        Cell_Val = sheet.cell(RowNo,VehicleCondition).value
        sRequest = SetXML(sRequest, "a:VehicleCondition", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Year).value
        sRequest = SetXML(sRequest, "a:Year", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Make).value
        sRequest = SetXML(sRequest, "a:Make", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Model).value
        sRequest = SetXML(sRequest, "a:Model", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Trim).value
        sRequest = SetXML(sRequest, "a:Trim", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,ChromeStyleId).value
        sRequest = SetXML(sRequest, "a:ChromeStyleId", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,VIN).value
        sRequest = SetXML(sRequest, "a:VIN", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,StockNumber).value
        sRequest = SetXML(sRequest, "a:StockNumber", Cell_Val.encode('ascii','ignore'), 0)
    
        Cell_Val = sheet.cell(RowNo,CertifiedUsed).value
        sRequest = SetXML(sRequest, "a:CertifiedUsed", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,InteriorColor).value
        sRequest = SetXML(sRequest, "a:InteriorColor", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,ExteriorColor).value
        sRequest = SetXML(sRequest, "a:ExteriorColor", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Mileage).value
        sRequest = SetXML(sRequest, "a:Mileage", Cell_Val.encode('ascii','ignore'), 0)
    
    if Is_Loan_Val =='Y':
        # loan info
        Cell_Val = sheet.cell(RowNo,CashDown).value
        sRequest = SetXML(sRequest, "a:CashDown", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Term).value
        sRequest = SetXML(sRequest, "a:Term", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,AmountRequesting).value
        sRequest = SetXML(sRequest, "a:AmountRequesting", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,MonthlyPayment).value
        sRequest = SetXML(sRequest, "a:MonthlyPayment", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,CreditType).value
        sRequest = SetXML(sRequest, "a:CreditType", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,ApplicantType).value
        sRequest = SetXML(sRequest, "a:ApplicantType", Cell_Val.encode('ascii','ignore'), 0)

    if Is_Tradein_Val == 'Y' and Is_Vehicle_Val == 'Y':
        # trade in info
        Cell_Val = sheet.cell(RowNo,Trade_Year).value
        sRequest = SetXML(sRequest, "a:Year", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Trade_Make).value
        sRequest = SetXML(sRequest, "a:Make", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Trade_Model).value
        sRequest = SetXML(sRequest, "a:Model", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Trade_Trim).value
        sRequest = SetXML(sRequest, "a:Trim", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Trade_Mileage).value
        sRequest = SetXML(sRequest, "a:Mileage", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Trade_ChromeStyleId).value
        sRequest = SetXML(sRequest, "a:ChromeStyleId", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Trade_TradeInPaid).value
        sRequest = SetXML(sRequest, "a:TradeInPaid", Cell_Val.encode('ascii','ignore'), 0)

    if Is_Tradein_Val == 'Y' and Is_Vehicle_Val != 'Y':
        # trade in info
        Cell_Val = sheet.cell(RowNo,Trade_Year).value
        sRequest = SetXML(sRequest, "a:Year", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Trade_Make).value
        sRequest = SetXML(sRequest, "a:Make", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Trade_Model).value
        sRequest = SetXML(sRequest, "a:Model", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Trade_Trim).value
        sRequest = SetXML(sRequest, "a:Trim", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Trade_Mileage).value
        sRequest = SetXML(sRequest, "a:Mileage", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Trade_ChromeStyleId).value
        sRequest = SetXML(sRequest, "a:ChromeStyleId", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Trade_TradeInPaid).value
        sRequest = SetXML(sRequest, "a:TradeInPaid", Cell_Val.encode('ascii','ignore'), 0)
    
    Cell_Val = sheet.cell(RowNo,Comments).value
    sRequest = SetXML(sRequest, "Comments", Cell_Val.encode('ascii','ignore'), 0)

   #print sRequest
    
    #ReqFilePath = str(FolderPath) + '\\Request\\'+ WebService +'Request'+ str(RowNo) +'.xml'
    #ReqFilePath = str(FolderPath) + "\\WebService\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    ReqFilePath = str(FolderPath) + "\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    f = open(ReqFilePath, 'w')
    f.write(sRequest)
    f.close()
    
    return sRequest                     # return request string

def rreplace(s, old, new, count):
    return (s[::-1].replace(old[::-1], new[::-1], count))[::-1]

def findnth(source, target, n):
    num = 0
    start = -1
    while num < n:
        start = source.find(target, start+1)
        if start == -1: return -1
        num += 1
    return start

def replacenth(source, old, new, n):
    p = findnth(source, old, n)
    if n == -1: return source
    print source[:p] + new + source[p+len(old):] 
    return source[:p] + new + source[p+len(old):] 


#----------------------------------------------------------
# Description:      Function to create request XML string for FD - CreditApp web service for FD 1.1
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row in excel file
#                   Schemafile - File path of schema file
# Return Values:    request XML string
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Create_FD_CreditApp_Request_1_1(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath):
    # open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)
    
    # get column numbers from excel
    # PrimaryApplicant section
    PartnerId = Get_ColumnNumber(HeaderList, "PartnerId")   # Calling a function to get column no of "PartnerDealerId" column
    PartnerDealerId = Get_ColumnNumber(HeaderList, "PartnerDealerIds")
    FinanceMethod = Get_ColumnNumber(HeaderList, "FinanceMethod")

    FirstName = Get_ColumnNumber(HeaderList, "FirstName")
    MiddleInitial = Get_ColumnNumber(HeaderList, "MiddleInitial")
    LastName = Get_ColumnNumber(HeaderList, "LastName")
    Suffix = Get_ColumnNumber(HeaderList, "Suffix")
    EmailAddress = Get_ColumnNumber(HeaderList, "EmailAddress")
    SSN = Get_ColumnNumber(HeaderList, "SSN")
    DateOfBirth = Get_ColumnNumber(HeaderList, "DateOfBirth")
    DLNumber = Get_ColumnNumber(HeaderList, "DriverLicenseNumber")
    DLState = Get_ColumnNumber(HeaderList, "DriverLicenseState")
    HomePhone = Get_ColumnNumber(HeaderList, "HomePhone")
    AddressLine1 = Get_ColumnNumber(HeaderList, "AddressLine1")
    AddressLine2 = Get_ColumnNumber(HeaderList, "AddressLine2")
    City = Get_ColumnNumber(HeaderList, "City")
    State = Get_ColumnNumber(HeaderList, "State")
    ZipCode = Get_ColumnNumber(HeaderList, "ZipCode")
    HousingStatus = Get_ColumnNumber(HeaderList, "HousingStatus")
    MortgageOrRent = Get_ColumnNumber(HeaderList, "MortgageOrRent")
    TotalMonthsAtAddress = Get_ColumnNumber(HeaderList, "TotalMonthsAtAddress")

    # Prev Address info
    Prev_AddressLine1 = Get_ColumnNumber(HeaderList, "Prev_AddressLine1")
    Prev_AddressLine2 = Get_ColumnNumber(HeaderList, "Prev_AddressLine2")
    Prev_City = Get_ColumnNumber(HeaderList, "Prev_City")
    Prev_State = Get_ColumnNumber(HeaderList, "Prev_State")
    Prev_ZipCode = Get_ColumnNumber(HeaderList, "Prev_ZipCode")

    # Current Emp
    EmploymentStatus = Get_ColumnNumber(HeaderList, "EmploymentStatus")
    EmployedBy = Get_ColumnNumber(HeaderList, "EmployedBy")
    TotalMonthsEmployed = Get_ColumnNumber(HeaderList, "TotalMonthsEmployed")
    BusinessPhone = Get_ColumnNumber(HeaderList, "BusinessPhone")
    MonthlyIncome = Get_ColumnNumber(HeaderList, "MonthlyIncome")

    # Prev Emp
    Prev_EmploymentStatus = Get_ColumnNumber(HeaderList, "Prev_EmploymentStatus")
    Prev_EmployedBy = Get_ColumnNumber(HeaderList, "Prev_EmployedBy")
    Prev_TotalMonthsEmployed = Get_ColumnNumber(HeaderList, "Prev_TotalMonthsEmployed")
    Prev_BusinessPhone = Get_ColumnNumber(HeaderList, "Prev_BusinessPhone")

    OtherMonthlyIncome = Get_ColumnNumber(HeaderList, "OtherMonthlyIncome")
    OtherIncomeSource = Get_ColumnNumber(HeaderList, "OtherIncomeSource")
    ConsentIndicator = Get_ColumnNumber(HeaderList, "ConsentIndicator")

    # Co- Applicant section

    Coapp_FirstName = Get_ColumnNumber(HeaderList, "Coapp_FirstName")
    Coapp_MiddleInitial = Get_ColumnNumber(HeaderList, "Coapp_MiddleInitial")
    Coapp_LastName = Get_ColumnNumber(HeaderList, "Coapp_LastName")
    Coapp_Suffix = Get_ColumnNumber(HeaderList, "Coapp_Suffix")
    Coapp_EmailAddress = Get_ColumnNumber(HeaderList, "Coapp_EmailAddress")
    Coapp_SSN = Get_ColumnNumber(HeaderList, "Coapp_SSN")
    Coapp_DateOfBirth = Get_ColumnNumber(HeaderList, "Coapp_DateOfBirth")
    Coapp_DLNumber = Get_ColumnNumber(HeaderList, "Coapp_DriverLicenseNumber")
    Coapp_DLState = Get_ColumnNumber(HeaderList, "Coapp_DriverLicenseState")
    Coapp_HomePhone = Get_ColumnNumber(HeaderList, "Coapp_HomePhone")
    Coapp_AddressLine1 = Get_ColumnNumber(HeaderList, "Coapp_AddressLine1")
    Coapp_AddressLine2 = Get_ColumnNumber(HeaderList, "Coapp_AddressLine2")
    Coapp_City = Get_ColumnNumber(HeaderList, "Coapp_City")
    Coapp_State = Get_ColumnNumber(HeaderList, "Coapp_State")
    Coapp_ZipCode = Get_ColumnNumber(HeaderList, "Coapp_ZipCode")
    Coapp_HousingStatus = Get_ColumnNumber(HeaderList, "Coapp_HousingStatus")
    Coapp_MortgageOrRent = Get_ColumnNumber(HeaderList, "Coapp_MortgageOrRent")
    Coapp_TotalMonthsAtAddress = Get_ColumnNumber(HeaderList, "Coapp_TotalMonthsAtAddress")
    
    # Co-App Prev address
    Coapp_Prev_AddressLine1 = Get_ColumnNumber(HeaderList, "Coapp_Prev_AddressLine1")
    Coapp_Prev_AddressLine2 = Get_ColumnNumber(HeaderList, "Coapp_Prev_AddressLine2")
    Coapp_Prev_City = Get_ColumnNumber(HeaderList, "Coapp_Prev_City")
    Coapp_Prev_State = Get_ColumnNumber(HeaderList, "Coapp_Prev_State")
    Coapp_Prev_ZipCode = Get_ColumnNumber(HeaderList, "Coapp_Prev_ZipCode")

    # Co-App Emp
    Coapp_EmploymentStatus = Get_ColumnNumber(HeaderList, "Coapp_EmploymentStatus")
    Coapp_EmployedBy = Get_ColumnNumber(HeaderList, "Coapp_EmployedBy")
    Coapp_TotalMonthsEmployed = Get_ColumnNumber(HeaderList, "Coapp_TotalMonthsEmployed")
    Coapp_BusinessPhone = Get_ColumnNumber(HeaderList, "Coapp_BusinessPhone")
    Coapp_MonthlyIncome = Get_ColumnNumber(HeaderList, "Coapp_MonthlyIncome")
    
    # Co-App Prev Emp
    Coapp_Prev_EmploymentStatus = Get_ColumnNumber(HeaderList, "Coapp_Prev_EmploymentStatus")
    Coapp_Prev_EmployedBy = Get_ColumnNumber(HeaderList, "Coapp_Prev_EmployedBy")
    Coapp_Prev_TotalMonthsEmployed = Get_ColumnNumber(HeaderList, "Coapp_Prev_TotalMonthsEmployed")
    Coapp_Prev_BusinessPhone = Get_ColumnNumber(HeaderList, "Coapp_Prev_BusinessPhone")

    Coapp_OtherMonthlyIncome = Get_ColumnNumber(HeaderList, "Coapp_OtherMonthlyIncome")
    Coapp_OtherIncomeSource = Get_ColumnNumber(HeaderList, "Coapp_OtherIncomeSource")
    Relationship = Get_ColumnNumber(HeaderList, "Relationship")

    # VehicleInfo section
    VehicleCondition = Get_ColumnNumber(HeaderList, "VehicleCondition")
    Year = Get_ColumnNumber(HeaderList, "Year")
    Make = Get_ColumnNumber(HeaderList, "Make")
    Model = Get_ColumnNumber(HeaderList, "Model")
    Trim = Get_ColumnNumber(HeaderList, "Trim")
    ChromeStyleId = Get_ColumnNumber(HeaderList, "ChromeStyleId")
    VIN = Get_ColumnNumber(HeaderList, "VIN")
    StockNumber = Get_ColumnNumber(HeaderList, "StockNumber")
    CertifiedUsed = Get_ColumnNumber(HeaderList, "CertifiedUsed")
    

    # Product Info section
    CashSellingPrice = Get_ColumnNumber(HeaderList, "CashSellingPrice")
    SalesTax = Get_ColumnNumber(HeaderList, "SalesTax")
    Title = Get_ColumnNumber(HeaderList, "Title")
    CashDown = Get_ColumnNumber(HeaderList, "CashDown")
    Rebate = Get_ColumnNumber(HeaderList, "Rebate")
    CreditLifeIns = Get_ColumnNumber(HeaderList, "CreditLifeIns")
    Term = Get_ColumnNumber(HeaderList, "Term")
    AcquisitionFees = Get_ColumnNumber(HeaderList, "AcquisitionFees")
    InvoiceAmount = Get_ColumnNumber(HeaderList, "InvoiceAmount")
    Warranty = Get_ColumnNumber(HeaderList, "Warranty")
    MSRP = Get_ColumnNumber(HeaderList, "MSRP")
    EstimatedBalloonAmount = Get_ColumnNumber(HeaderList, "EstimatedBalloonAmount")
    EstimatedPayment = Get_ColumnNumber(HeaderList, "EstimatedPayment")
    UsedCarBook = Get_ColumnNumber(HeaderList, "UsedCarBook")
    Mileage = Get_ColumnNumber(HeaderList, "Mileage")
    UsedCarValue = Get_ColumnNumber(HeaderList, "UsedCarValue")
    OtherFees = Get_ColumnNumber(HeaderList, "OtherFees")
    WholesaleBookSource = Get_ColumnNumber(HeaderList, "WholesaleBookSource")
    WholesaleCondition = Get_ColumnNumber(HeaderList, "WholesaleCondition")
    WholesaleValueType = Get_ColumnNumber(HeaderList, "WholesaleValueType")
    WholesaleValue = Get_ColumnNumber(HeaderList, "WholesaleValue")
    NetTrade = Get_ColumnNumber(HeaderList, "NetTrade")
    
    # TradeInVehicleInfo section
    Trade_Year = Get_ColumnNumber(HeaderList, "Trade_Year")
    Trade_Make = Get_ColumnNumber(HeaderList, "Trade_Make")
    Trade_Model = Get_ColumnNumber(HeaderList, "Trade_Model")
    Trade_Trim = Get_ColumnNumber(HeaderList, "Trade_Trim")
    Trade_ChromeStyleId = Get_ColumnNumber(HeaderList, "Trade_ChromeStyleId")

    LienHolder = Get_ColumnNumber(HeaderList, "LienHolder")
    MonthlyPayment = Get_ColumnNumber(HeaderList, "MonthlyPayment")
    
    LeadComments = Get_ColumnNumber(HeaderList, "LeadComments")

    # reference numbers
    PrequalificationReferenceNumber = Get_ColumnNumber(HeaderList, "PrequalificationReferenceNumber")
    LeadReferenceNumber = Get_ColumnNumber(HeaderList, "LeadReferenceNumber")

    # Read if section present values
    IsApplicantSection = Get_ColumnNumber(HeaderList, "Is_Applicant")
    IsCoAppSection = Get_ColumnNumber(HeaderList, "Is_CoApp")
    IsVehicleSection = Get_ColumnNumber(HeaderList, "Is_Vehicle")
    IsProductSection = Get_ColumnNumber(HeaderList, "Is_Product")
    IsTradeinSection = Get_ColumnNumber(HeaderList, "Is_Tradein")
    
    # Open schema file
    try:
        f = open(Schemafile)
    except IOError, e:
        print e
        print "Problem while opening Schema file"
        raise ErrorOccured("Problem in opening file: " + Schemafile)        # Problem opening file
    except:
        raise ErrorOccured("Unknown error occured.")                        # handle any other type of error

    # everything is fine
    sRequest = f.read()         # get contenets of schema file in to string
    f.close()

    #VehicleContion_Val = sheet.cell(RowNo,VehicleCondition).value
    
    # Changes done for missing sections **** 14th March 2013
    # getting values from execl is section is present
    IsApplicant_Val = sheet.cell(RowNo,IsApplicantSection).value
    #print "Applicant: ", IsApplicant_Val
    IsCoApp_Val = sheet.cell(RowNo,IsCoAppSection).value
    #print "Co-app: ",IsCoApp_Val
    IsVehicle_Val = sheet.cell(RowNo,IsVehicleSection).value
    #print "Vehicle: ",IsVehicle_Val
    IsProduct_Val = sheet.cell(RowNo,IsProductSection).value
    #print "Product: ",IsProduct_Val
    IsTradein_Val = sheet.cell(RowNo,IsTradeinSection).value
    #print "trade in: ",IsTradein_Val

    if IsApplicant_Val=='Y' and IsCoApp_Val=='Y':
        
        # change the schema strings for below fields if the vale is null
    
        # PartnerId
        Cell_Val = sheet.cell(RowNo,PartnerId).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:PartnerId></a:PartnerId>","""<a:PartnerId  i:nil="true" />""",1)
    
        # DOBs
        Cell_Val = sheet.cell(RowNo,DateOfBirth).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:DateOfBirth></a:DateOfBirth>","""<a:DateOfBirth  i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,Coapp_DateOfBirth).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:DateOfBirth></a:DateOfBirth>","""<a:DateOfBirth  i:nil="true" />""",2)

        # suffix
        Cell_Val = sheet.cell(RowNo,Suffix).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:Suffix></a:Suffix>","""<a:Suffix  i:nil="true" />""",1)

        Cell_Val = sheet.cell(RowNo,Coapp_Suffix).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:Suffix></a:Suffix>","""<a:Suffix  i:nil="true" />""",2)
    
        # addressline 2
        Cell_Val = sheet.cell(RowNo,AddressLine2).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest, "<a:AddressLine2></a:AddressLine2>", """<a:AddressLine2 i:nil="true" />""", 1) 

        Cell_Val = sheet.cell(RowNo,Coapp_AddressLine2).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest, "<a:AddressLine2></a:AddressLine2>", """<a:AddressLine2 i:nil="true" />""", 3)

        # housing status
        Cell_Val = sheet.cell(RowNo,HousingStatus).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest, "<a:HousingStatus></a:HousingStatus>", """<a:HousingStatus i:nil="true" />""", 1)

        Cell_Val = sheet.cell(RowNo,Coapp_HousingStatus).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:HousingStatus></a:HousingStatus>","""<a:HousingStatus  i:nil="true" />""",2)

        # monthly income
        Cell_Val = sheet.cell(RowNo,MonthlyIncome).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest, "<a:MonthlyIncome></a:MonthlyIncome>", """<a:MonthlyIncome i:nil="true" />""", 1)
    
        Cell_Val = sheet.cell(RowNo,Coapp_MonthlyIncome).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:MonthlyIncome></a:MonthlyIncome>","""<a:MonthlyIncome  i:nil="true" />""",2)

        # Emp Status
        Cell_Val = sheet.cell(RowNo,EmploymentStatus).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest, "<a:EmploymentStatus></a:EmploymentStatus>", """<a:EmploymentStatus i:nil="true" />""", 1)

        Cell_Val = sheet.cell(RowNo,Coapp_EmploymentStatus).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:EmploymentStatus></a:EmploymentStatus>","""<a:EmploymentStatus  i:nil="true" />""",3)

        # Emp By
        Cell_Val = sheet.cell(RowNo,EmployedBy).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest, "<a:EmployedBy></a:EmployedBy>", """<a:EmployedBy i:nil="true" />""", 1)

        Cell_Val = sheet.cell(RowNo,Coapp_EmployedBy).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:EmployedBy></a:EmployedBy>","""<a:EmployedBy  i:nil="true" />""",3)

        # Total months Employed
        Cell_Val = sheet.cell(RowNo,TotalMonthsEmployed).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest, "<a:TotalMonthsEmployed></a:TotalMonthsEmployed>", """<a:TotalMonthsEmployed i:nil="true" />""", 1)

        Cell_Val = sheet.cell(RowNo,Coapp_TotalMonthsEmployed).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:TotalMonthsEmployed></a:TotalMonthsEmployed>","""<a:TotalMonthsEmployed  i:nil="true" />""",3)

        # BusinessPhone 
        Cell_Val = sheet.cell(RowNo,BusinessPhone).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest, "<a:BusinessPhone></a:BusinessPhone>", """<a:BusinessPhone i:nil="true" />""", 1)

        Cell_Val = sheet.cell(RowNo,Coapp_BusinessPhone).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:BusinessPhone></a:BusinessPhone>","""<a:BusinessPhone  i:nil="true" />""",3)

         #  Prev_EmploymentStatus 
        Cell_Val = sheet.cell(RowNo,Prev_EmploymentStatus).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest, "<a:EmploymentStatus></a:EmploymentStatus>", """<a:EmploymentStatus i:nil="true" />""", 2)

        Cell_Val = sheet.cell(RowNo,Coapp_Prev_EmploymentStatus).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:EmploymentStatus></a:EmploymentStatus>","""<a:EmploymentStatus  i:nil="true" />""",4)
        
        # Prev_EmployedBy
        Cell_Val = sheet.cell(RowNo,Prev_EmployedBy).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest, "<a:EmployedBy></a:EmployedBy>", """<a:EmployedBy i:nil="true" />""", 2)

        Cell_Val = sheet.cell(RowNo,Coapp_Prev_EmployedBy).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:EmployedBy></a:EmployedBy>","""<a:EmployedBy  i:nil="true" />""",4)
    
        # Prev_TotalMonthsEmployed
        Cell_Val = sheet.cell(RowNo,Prev_TotalMonthsEmployed).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest, "<a:TotalMonthsEmployed></a:TotalMonthsEmployed>", """<a:TotalMonthsEmployed i:nil="true" />""", 2)

        Cell_Val = sheet.cell(RowNo,Coapp_Prev_TotalMonthsEmployed).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:TotalMonthsEmployed></a:TotalMonthsEmployed>","""<a:TotalMonthsEmployed  i:nil="true" />""",4)
        
        #Prev_BusinessPhone
        Cell_Val = sheet.cell(RowNo,Prev_BusinessPhone).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest, "<a:BusinessPhone></a:BusinessPhone>", """<a:BusinessPhone i:nil="true" />""", 2)

        Cell_Val = sheet.cell(RowNo,Coapp_Prev_BusinessPhone).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:BusinessPhone></a:BusinessPhone>","""<a:BusinessPhone  i:nil="true" />""",4)

        # Other Monthly income
        Cell_Val = sheet.cell(RowNo,OtherMonthlyIncome).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest, "<a:OtherMonthlyIncome></a:OtherMonthlyIncome>", """<a:OtherMonthlyIncome i:nil="true" />""", 1)

        Cell_Val = sheet.cell(RowNo,Coapp_OtherMonthlyIncome).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:OtherMonthlyIncome></a:OtherMonthlyIncome>","""<a:OtherMonthlyIncome  i:nil="true" />""",2)

        # Relationship
        Cell_Val = sheet.cell(RowNo,Relationship).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest, "<a:Relationship></a:Relationship>", """<a:Relationship i:nil="true" />""", 1)

    # if No Applicant and only Co-App
    elif IsApplicant_Val!='Y' and IsCoApp_Val=='Y':
        sRequest = sRequest.replace("""<PrimaryApplicant xmlns:a="http://schemas.datacontract.org/2004/07/DealerTrack.DataContracts.CreditApp">
    <a:ApplicantInfo>
      <a:FirstName></a:FirstName>
      <a:MiddleInitial></a:MiddleInitial>
      <a:LastName></a:LastName>
      <a:Suffix></a:Suffix>
      <a:EmailAddress></a:EmailAddress>
      <a:SSN></a:SSN>
      <a:DateOfBirth></a:DateOfBirth>
      <a:DriverLicenseNumber></a:DriverLicenseNumber>
      <a:DriverLicenseState></a:DriverLicenseState>
      <a:HomePhone></a:HomePhone>
      <a:CurrentAddress>
        <a:AddressLine1></a:AddressLine1>
        <a:AddressLine2></a:AddressLine2>
        <a:City></a:City>
        <a:State></a:State>
        <a:ZipCode></a:ZipCode>
      </a:CurrentAddress>
      <a:HousingStatus></a:HousingStatus>
      <a:MortgageOrRent></a:MortgageOrRent>
      <a:TotalMonthsAtAddress></a:TotalMonthsAtAddress>
      <a:PreviousAddress>
        <a:AddressLine1></a:AddressLine1>
        <a:AddressLine2></a:AddressLine2>
        <a:City></a:City>
        <a:State></a:State>
        <a:ZipCode></a:ZipCode>
      </a:PreviousAddress>
      <a:CurrentEmploymentInfo>
        <a:EmploymentStatus></a:EmploymentStatus>
        <a:EmployedBy></a:EmployedBy>
        <a:TotalMonthsEmployed></a:TotalMonthsEmployed>
        <a:BusinessPhone></a:BusinessPhone>
      </a:CurrentEmploymentInfo>
      <a:MonthlyIncome></a:MonthlyIncome>
      <a:PreviousEmploymentInfo>
        <a:EmploymentStatus></a:EmploymentStatus>
        <a:EmployedBy></a:EmployedBy>
        <a:TotalMonthsEmployed></a:TotalMonthsEmployed>
        <a:BusinessPhone></a:BusinessPhone>
      </a:PreviousEmploymentInfo>
      <a:OtherMonthlyIncome></a:OtherMonthlyIncome>
      <a:OtherIncomeSource></a:OtherIncomeSource>
    </a:ApplicantInfo>
      <a:ConsentIndicator></a:ConsentIndicator>
    </PrimaryApplicant>""","",1)
        print "*** Request without Applicant Section ***"
        print sRequest
        
    # if only Applicant and no Co-App
    elif IsApplicant_Val=='Y' and IsCoApp_Val!='Y':
        sRequest = sRequest.replace("""<CoApplicant xmlns:a="http://schemas.datacontract.org/2004/07/DealerTrack.DataContracts.CreditApp">
    <a:ApplicantInfo>
      <a:FirstName></a:FirstName>
      <a:MiddleInitial></a:MiddleInitial>
      <a:LastName></a:LastName>
      <a:Suffix></a:Suffix>
      <a:EmailAddress></a:EmailAddress>
      <a:SSN></a:SSN>
      <a:DateOfBirth></a:DateOfBirth>
      <a:DriverLicenseNumber></a:DriverLicenseNumber>
      <a:DriverLicenseState></a:DriverLicenseState>
      <a:HomePhone></a:HomePhone>
      <a:CurrentAddress>
        <a:AddressLine1></a:AddressLine1>
        <a:AddressLine2></a:AddressLine2>
        <a:City></a:City>
        <a:State></a:State>
        <a:ZipCode></a:ZipCode>
      </a:CurrentAddress>
      <a:HousingStatus></a:HousingStatus>
      <a:MortgageOrRent></a:MortgageOrRent>
      <a:TotalMonthsAtAddress></a:TotalMonthsAtAddress>
	  <a:PreviousAddress>
	    <a:AddressLine1></a:AddressLine1>
        <a:AddressLine2></a:AddressLine2>
        <a:City></a:City>
        <a:State></a:State>
        <a:ZipCode></a:ZipCode>
      </a:PreviousAddress>
      <a:CurrentEmploymentInfo>
        <a:EmploymentStatus></a:EmploymentStatus>
        <a:EmployedBy></a:EmployedBy>
        <a:TotalMonthsEmployed></a:TotalMonthsEmployed>
        <a:BusinessPhone></a:BusinessPhone>
      </a:CurrentEmploymentInfo>
      <a:MonthlyIncome></a:MonthlyIncome>
      <a:PreviousEmploymentInfo>
        <a:EmploymentStatus></a:EmploymentStatus>
        <a:EmployedBy></a:EmployedBy>
        <a:TotalMonthsEmployed></a:TotalMonthsEmployed>
        <a:BusinessPhone></a:BusinessPhone>
      </a:PreviousEmploymentInfo>
      <a:OtherMonthlyIncome></a:OtherMonthlyIncome>
      <a:OtherIncomeSource></a:OtherIncomeSource>
    </a:ApplicantInfo>
      <a:Relationship></a:Relationship>          
    </CoApplicant>""","",1)
        print "*** Request without Co-AppSection ***"
        print sRequest

   
    # ****************************** Product Info ****************************************
    # if product info is present i.e = 'Y'
    if IsProduct_Val == 'Y':
        # CashSellingPrice
        Cell_Val = sheet.cell(RowNo,CashSellingPrice).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:CashSellingPrice></a:CashSellingPrice>", """<a:CashSellingPrice i:nil="true" />""",1)

        # Sales Tax
        Cell_Val = sheet.cell(RowNo,SalesTax).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:SalesTax></a:SalesTax>", """<a:SalesTax i:nil="true" />""",1)

        # Net Trade
        Cell_Val = sheet.cell(RowNo,NetTrade).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:NetTrade></a:NetTrade>", """<a:NetTrade i:nil="true" />""",1)
   
        # Title
        Cell_Val = sheet.cell(RowNo,Title).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:Title></a:Title>", """<a:Title i:nil="true" />""",1)

        # CashDown
        Cell_Val = sheet.cell(RowNo,CashDown).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:CashDown></a:CashDown>", """<a:CashDown i:nil="true" />""",1)

        # Rebate
        Cell_Val = sheet.cell(RowNo,Rebate).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:Rebate></a:Rebate>", """<a:Rebate i:nil="true" />""",1)

        # CreditLifeIns
        Cell_Val = sheet.cell(RowNo,CreditLifeIns).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:CreditLifeIns></a:CreditLifeIns>", """<a:CreditLifeIns i:nil="true" />""",1)

        # Term
        Cell_Val = sheet.cell(RowNo,Term).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:Term></a:Term>", """<a:Term i:nil="true" />""",1)

        # AcquisitionFees
        Cell_Val = sheet.cell(RowNo,AcquisitionFees).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:AcquisitionFees></a:AcquisitionFees>", """<a:AcquisitionFees i:nil="true" />""",1)

        # InvoiceAmount
        Cell_Val = sheet.cell(RowNo,InvoiceAmount).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:InvoiceAmount></a:InvoiceAmount>", """<a:InvoiceAmount i:nil="true" />""",1)

        # Warranty
        Cell_Val = sheet.cell(RowNo,Warranty).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:Warranty></a:Warranty>", """<a:Warranty i:nil="true" />""",1)

        # MSRP
        Cell_Val = sheet.cell(RowNo,MSRP).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:MSRP></a:MSRP>", """<a:MSRP i:nil="true" />""",1)

        # EstimatedBalloonAmount
        Cell_Val = sheet.cell(RowNo,EstimatedBalloonAmount).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:EstimatedBalloonAmount></a:EstimatedBalloonAmount>", """<a:EstimatedBalloonAmount i:nil="true" />""",1)

        # EstimatedPayment
        Cell_Val = sheet.cell(RowNo,EstimatedPayment).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:EstimatedPayment></a:EstimatedPayment>", """<a:EstimatedPayment i:nil="true" />""",1)

        # OtherFees
        Cell_Val = sheet.cell(RowNo,OtherFees).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:OtherFees></a:OtherFees>", """<a:OtherFees i:nil="true" />""",1)

        # Used car book
        Cell_Val = sheet.cell(RowNo,UsedCarBook).value
        if len(Cell_Val) != 0:
            sRequest = sRequest.replace("""<a:UsedCarBook  i:nil="true" />""","<a:UsedCarBook></a:UsedCarBook>",1)

        # Mileage
        Cell_Val = sheet.cell(RowNo,Mileage).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:Mileage></a:Mileage>", """<a:Mileage i:nil="true" />""",1)       

        # used car value
        Cell_Val = sheet.cell(RowNo,UsedCarValue).value
        if len(Cell_Val) != 0:
            sRequest = sRequest.replace("""<a:UsedCarValue i:nil="true" />""","<a:UsedCarValue></a:UsedCarValue>",1)

        # whole sale book source
        Cell_Val = sheet.cell(RowNo,WholesaleBookSource).value
        if len(Cell_Val) != 0:
            sRequest = sRequest.replace("""<a:WholesaleBookSource i:nil="true" />""","<a:WholesaleBookSource></a:WholesaleBookSource>",1)

        # WholesaleCondition
        Cell_Val = sheet.cell(RowNo,WholesaleCondition).value
        if len(Cell_Val) != 0:
            sRequest = sRequest.replace("""<a:WholesaleCondition i:nil="true" />""","<a:WholesaleCondition></a:WholesaleCondition>",1)

        # WholesaleValueType
        Cell_Val = sheet.cell(RowNo,WholesaleValueType).value
        if len(Cell_Val) != 0:
            sRequest = sRequest.replace("""<a:WholesaleValueType i:nil="true" />""","<a:WholesaleValueType></a:WholesaleValueType>",1)

        # WholesaleValue
        Cell_Val = sheet.cell(RowNo,WholesaleValue).value
        if len(Cell_Val) != 0:
            sRequest = sRequest.replace("""<a:WholesaleValue i:nil="true" />""","<a:WholesaleValue></a:WholesaleValue>",1)

    # if product info is not present i.e = 'N'
    else:
        sRequest = sRequest.replace("""<ProductInfo xmlns:a="http://schemas.datacontract.org/2004/07/DealerTrack.DataContracts.CreditApp">
    <a:CashSellingPrice></a:CashSellingPrice>
    <a:SalesTax></a:SalesTax>
    <a:Title></a:Title>
    <a:CashDown></a:CashDown>
    <a:Rebate></a:Rebate>
    <a:CreditLifeIns></a:CreditLifeIns>
    <a:Term></a:Term>
    <a:AcquisitionFees></a:AcquisitionFees>
    <a:InvoiceAmount></a:InvoiceAmount>
    <a:Warranty></a:Warranty>
    <a:MSRP></a:MSRP>
    <a:EstimatedBalloonAmount></a:EstimatedBalloonAmount>
    <a:EstimatedPayment></a:EstimatedPayment>
    <a:UsedCarBook  i:nil="true" />
    <a:Mileage></a:Mileage>
    <a:UsedCarValue i:nil="true" />
    <a:OtherFees></a:OtherFees>
    <a:WholesaleBookSource i:nil="true" />
    <a:WholesaleCondition i:nil="true" />
    <a:WholesaleValueType i:nil="true" />
    <a:WholesaleValue i:nil="true" />
    <a:NetTrade></a:NetTrade>
  </ProductInfo>""","",1)
        print "*** Request without Product info ***"
        print sRequest

    # ****************************** Vehicle Info ****************************************
    # if Vehicle info is present i.e = 'Y'
    if IsVehicle_Val == 'Y':
        # VehicleCondition
        Cell_Val = sheet.cell(RowNo,VehicleCondition).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:VehicleCondition></a:VehicleCondition>", """<a:VehicleCondition i:nil="true" />""",1)

        # Veh Year - Trade year
        Cell_Val = sheet.cell(RowNo,Year).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest, "<a:Year></a:Year>", """<a:Year i:nil="true" />""", 1)

        #Cell_Val = sheet.cell(RowNo,Trade_Year).value
        #if len(Cell_Val) == 0:
         #   sRequest = replacenth(sRequest,"<a:Year></a:Year>","""<a:Year  i:nil="true" />""",2)

        # Veh make - trade make
        Cell_Val = sheet.cell(RowNo,Make).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest, "<a:Make></a:Make>", """<a:Make i:nil="true" />""", 1)

        #Cell_Val = sheet.cell(RowNo,Trade_Make).value
        #if len(Cell_Val) == 0:
         #   sRequest = replacenth(sRequest,"<a:Make></a:Make>","""<a:Make  i:nil="true" />""",2)

        # Veh Model - trade Model
        Cell_Val = sheet.cell(RowNo,Model).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest, "<a:Model></a:Model>", """<a:Model i:nil="true" />""", 1)

        #Cell_Val = sheet.cell(RowNo,Trade_Model).value
        #if len(Cell_Val) == 0:
         #   sRequest = replacenth(sRequest,"<a:Model></a:Model>","""<a:Model  i:nil="true" />""",2)
        
        # Veh Trim - trade Trim
        Cell_Val = sheet.cell(RowNo,Trim).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest, "<a:Trim></a:Trim>", """<a:Trim i:nil="true" />""", 1)

        #Cell_Val = sheet.cell(RowNo,Trade_Trim).value
        #if len(Cell_Val) == 0:
         #   sRequest = replacenth(sRequest,"<a:Trim></a:Trim>","""<a:Trim  i:nil="true" />""",2)

        # Veh ChromeStyleId - Trade ChromeStyleId
        Cell_Val = sheet.cell(RowNo,ChromeStyleId).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest, "<a:ChromeStyleId></a:ChromeStyleId>", """<a:ChromeStyleId i:nil="true" />""", 1)

        #Cell_Val = sheet.cell(RowNo,Trade_ChromeStyleId).value
        #if len(Cell_Val) == 0:
         #   sRequest = replacenth(sRequest,"<a:ChromeStyleId></a:ChromeStyleId>","""<a:ChromeStyleId  i:nil="true" />""",2)

        # veh - VIN
        Cell_Val = sheet.cell(RowNo,VIN).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:VIN></a:VIN>", """<a:VIN i:nil="true" />""",1)

        # Veh - StockNumber
        Cell_Val = sheet.cell(RowNo,StockNumber).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:StockNumber></a:StockNumber>", """<a:StockNumber i:nil="true" />""",1)

        # Veh - CertifiedUsed
        Cell_Val = sheet.cell(RowNo,CertifiedUsed).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:CertifiedUsed></a:CertifiedUsed>", """<a:CertifiedUsed i:nil="true" />""",1)

    # if Vehicle info is not present i.e = 'N'
    else:
        sRequest = sRequest.replace("""<VehicleInfo xmlns:a="http://schemas.datacontract.org/2004/07/DealerTrack.DataContracts.CreditApp">
    <a:VehicleCondition></a:VehicleCondition>
    <a:Year></a:Year>
    <a:Make></a:Make>
    <a:Model></a:Model>
    <a:Trim></a:Trim>
    <a:ChromeStyleId></a:ChromeStyleId>
    <a:VIN></a:VIN>
    <a:StockNumber></a:StockNumber>
    <a:CertifiedUsed></a:CertifiedUsed>
  </VehicleInfo>""","",1)
        print "*** Request without Vehicle info ***"
        print sRequest
        

    # Trade in section is present
    if IsTradein_Val == 'Y':
        # trade year
        Cell_Val = sheet.cell(RowNo,Trade_Year).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:Year></a:Year>","""<a:Year  i:nil="true" />""",2)

        #trade make
        Cell_Val = sheet.cell(RowNo,Trade_Make).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:Make></a:Make>","""<a:Make  i:nil="true" />""",2)

        #trade in model
        Cell_Val = sheet.cell(RowNo,Trade_Model).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:Model></a:Model>","""<a:Model  i:nil="true" />""",2)

        #trade in model
        Cell_Val = sheet.cell(RowNo,Trade_Trim).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:Trim></a:Trim>","""<a:Trim  i:nil="true" />""",2)

        # trade in chrome style id
        Cell_Val = sheet.cell(RowNo,Trade_ChromeStyleId).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:ChromeStyleId></a:ChromeStyleId>","""<a:ChromeStyleId  i:nil="true" />""",2)

        #  Trade In Vechicle LienHolder
        Cell_Val = sheet.cell(RowNo,LienHolder).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:LienHolder></a:LienHolder>", """<a:LienHolder i:nil="true" />""",1)

        #  Trade In Vechicle MonthlyPayment
        Cell_Val = sheet.cell(RowNo,MonthlyPayment).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:MonthlyPayment></a:MonthlyPayment>", """<a:MonthlyPayment i:nil="true" />""",1)
    else:
        #print "original req schema: "
        #print sRequest
        sRequest = sRequest.replace("""<TradeInVehicleInfo xmlns:a="http://schemas.datacontract.org/2004/07/DealerTrack.DataContracts.CreditApp">
    <a:Year></a:Year>
    <a:Make></a:Make>
    <a:Model></a:Model>
    <a:Trim></a:Trim>
    <a:ChromeStyleId></a:ChromeStyleId>
    <a:LienHolder></a:LienHolder>
    <a:MonthlyPayment></a:MonthlyPayment>
  </TradeInVehicleInfo>""","",1)
        print "*** Request without Trade in Vehicle info ***"
        print sRequest

    # ConsentIndicator
    Cell_Val = sheet.cell(RowNo,ConsentIndicator).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:ConsentIndicator></a:ConsentIndicator>", """<a:ConsentIndicator i:nil="true" />""",1)

    # reference numbers
    # Prequal
    Cell_Val = sheet.cell(RowNo,PrequalificationReferenceNumber).value
    #print "PreQ ref no: ",Cell_Val
    if len(Cell_Val) != 0:
        sRequest = sRequest.replace("""<PrequalificationReferenceNumber i:nil="true" />""","<PrequalificationReferenceNumber></PrequalificationReferenceNumber>",1)
        
    # Lead
    Cell_Val = sheet.cell(RowNo,LeadReferenceNumber).value
    #print "Lead ref no: ",Cell_Val
    if len(Cell_Val) != 0:
        sRequest = sRequest.replace("""<LeadReferenceNumber i:nil="true" />""","<LeadReferenceNumber></LeadReferenceNumber>",1)
        

    #print "************************** Request Before values ***********************"
#    print sRequest
 #   print "***********************************************************************"
    #if VehicleContion_Val != 'New':
     #   New_Req = sRequest.replace("""<a:UsedCarBook  i:nil="true" />""","<a:UsedCarBook></a:UsedCarBook>")
      #  New_Req = New_Req.replace("""<a:UsedCarValue i:nil="true" />""","<a:UsedCarValue></a:UsedCarValue>")
       # New_Req = New_Req.replace("""<a:WholesaleBookSource i:nil="true" />""","<a:WholesaleBookSource></a:WholesaleBookSource>")
       # New_Req = New_Req.replace("""<a:WholesaleCondition i:nil="true" />""","<a:WholesaleCondition></a:WholesaleCondition>")
        #New_Req = New_Req.replace("""<a:WholesaleValueType i:nil="true" />""","<a:WholesaleValueType></a:WholesaleValueType>")
        #New_Req = New_Req.replace("""<a:WholesaleValue i:nil="true" />""","<a:WholesaleValue></a:WholesaleValue>")
        #sRequest = New_Req
    
    
    # Create request string
    # Calling function to add values to nodes in given schema file
    Cell_Val = sheet.cell(RowNo,PartnerId).value
    #print "Partner ID: ",Cell_Val
    sRequest = SetXML(sRequest, "PartnerId", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,PartnerDealerId).value
    sRequest = SetXML(sRequest, "PartnerDealerId",Cell_Val.encode('ascii','ignore') , 0)

    Cell_Val = sheet.cell(RowNo,FinanceMethod).value
    sRequest = SetXML(sRequest, "FinanceMethod",Cell_Val.encode('ascii','ignore') , 0)

    if IsApplicant_Val == 'Y':
        
        # App info
        Cell_Val = sheet.cell(RowNo,FirstName).value
        sRequest = SetXML(sRequest, "a:FirstName", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,MiddleInitial).value
        sRequest = SetXML(sRequest, "a:MiddleInitial", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,LastName).value
        sRequest = SetXML(sRequest, "a:LastName", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Suffix).value
        sRequest = SetXML(sRequest, "a:Suffix", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,EmailAddress).value
        sRequest = SetXML(sRequest, "a:EmailAddress", Cell_Val.encode('ascii','ignore'), 0)
    
        Cell_Val = sheet.cell(RowNo,SSN).value
        sRequest = SetXML(sRequest, "a:SSN", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,DateOfBirth).value
        sRequest = SetXML(sRequest, "a:DateOfBirth", Cell_Val.encode('ascii','ignore'), 0)
   
        Cell_Val = sheet.cell(RowNo,DLNumber).value
        sRequest = SetXML(sRequest, "a:DriverLicenseNumber", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,DLState).value
        sRequest = SetXML(sRequest, "a:DriverLicenseState", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,HomePhone).value
        sRequest = SetXML(sRequest, "a:HomePhone", Cell_Val.encode('ascii','ignore'), 0)

        # App Current address
        Cell_Val = sheet.cell(RowNo,AddressLine1).value
        sRequest = SetXML(sRequest, "a:AddressLine1", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,AddressLine2).value
        sRequest = SetXML(sRequest, "a:AddressLine2",Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,City).value
        sRequest = SetXML(sRequest, "a:City",Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,State).value
        sRequest = SetXML(sRequest, "a:State",Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,ZipCode).value
        sRequest = SetXML(sRequest, "a:ZipCode", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,HousingStatus).value
        sRequest = SetXML(sRequest, "a:HousingStatus", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,MortgageOrRent).value
        sRequest = SetXML(sRequest, "a:MortgageOrRent", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,TotalMonthsAtAddress).value
        sRequest = SetXML(sRequest, "a:TotalMonthsAtAddress", Cell_Val.encode('ascii','ignore'), 0)

        # App Prev address
        Cell_Val = sheet.cell(RowNo,Prev_AddressLine1).value
        sRequest = SetXML(sRequest, "a:AddressLine1", Cell_Val.encode('ascii','ignore'), 1)    
    
        Cell_Val = sheet.cell(RowNo,Prev_AddressLine2).value
        sRequest = SetXML(sRequest, "a:AddressLine2", Cell_Val.encode('ascii','ignore'), 1)
    
        Cell_Val = sheet.cell(RowNo,Prev_City).value
        sRequest = SetXML(sRequest, "a:City", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Prev_State).value
        sRequest = SetXML(sRequest, "a:State", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Prev_ZipCode).value
        sRequest = SetXML(sRequest, "a:ZipCode", Cell_Val.encode('ascii','ignore'), 1)

        # App Current Emp
        Cell_Val = sheet.cell(RowNo,EmploymentStatus).value
        sRequest = SetXML(sRequest, "a:EmploymentStatus", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,EmployedBy).value
        sRequest = SetXML(sRequest, "a:EmployedBy", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,TotalMonthsEmployed).value
        sRequest = SetXML(sRequest, "a:TotalMonthsEmployed", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,BusinessPhone).value
        sRequest = SetXML(sRequest, "a:BusinessPhone", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,MonthlyIncome).value
        sRequest = SetXML(sRequest, "a:MonthlyIncome", Cell_Val.encode('ascii','ignore'), 0)

        # App Prev Emp
        Cell_Val = sheet.cell(RowNo,Prev_EmploymentStatus).value
        sRequest = SetXML(sRequest, "a:EmploymentStatus", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Prev_EmployedBy).value
        sRequest = SetXML(sRequest, "a:EmployedBy", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Prev_TotalMonthsEmployed).value
        sRequest = SetXML(sRequest, "a:TotalMonthsEmployed", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Prev_BusinessPhone).value
        sRequest = SetXML(sRequest, "a:BusinessPhone", Cell_Val.encode('ascii','ignore'), 1)

        # Others

        Cell_Val = sheet.cell(RowNo,OtherMonthlyIncome).value
        sRequest = SetXML(sRequest, "a:OtherMonthlyIncome", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,OtherIncomeSource).value
        sRequest = SetXML(sRequest, "a:OtherIncomeSource", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,ConsentIndicator).value
        sRequest = SetXML(sRequest, "a:ConsentIndicator", Cell_Val.encode('ascii','ignore'), 0)

        # Co-App *******************************************************************************************
        if IsCoApp_Val =='Y':
            # Co-App info
            Cell_Val = sheet.cell(RowNo,Coapp_FirstName).value
            sRequest = SetXML(sRequest, "a:FirstName", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,Coapp_MiddleInitial).value
            sRequest = SetXML(sRequest, "a:MiddleInitial", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,Coapp_LastName).value
            sRequest = SetXML(sRequest, "a:LastName", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,Coapp_Suffix).value
            sRequest = SetXML(sRequest, "a:Suffix", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,Coapp_EmailAddress).value
            sRequest = SetXML(sRequest, "a:EmailAddress", Cell_Val.encode('ascii','ignore'), 1)
    
            Cell_Val = sheet.cell(RowNo,Coapp_SSN).value
            sRequest = SetXML(sRequest, "a:SSN", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,Coapp_DateOfBirth).value
            sRequest = SetXML(sRequest, "a:DateOfBirth", Cell_Val.encode('ascii','ignore'), 1)
   
            Cell_Val = sheet.cell(RowNo,Coapp_DLNumber).value
            sRequest = SetXML(sRequest, "a:DriverLicenseNumber", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,Coapp_DLState).value
            sRequest = SetXML(sRequest, "a:DriverLicenseState", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,Coapp_HomePhone).value
            sRequest = SetXML(sRequest, "a:HomePhone", Cell_Val.encode('ascii','ignore'), 1)

            # Co-App Current address
            Cell_Val = sheet.cell(RowNo,Coapp_AddressLine1).value
            sRequest = SetXML(sRequest, "a:AddressLine1", Cell_Val.encode('ascii','ignore'), 2)

            Cell_Val = sheet.cell(RowNo,Coapp_AddressLine2).value
            sRequest = SetXML(sRequest, "a:AddressLine2",Cell_Val.encode('ascii','ignore'), 2)

            Cell_Val = sheet.cell(RowNo,Coapp_City).value
            sRequest = SetXML(sRequest, "a:City",Cell_Val.encode('ascii','ignore'), 2)

            Cell_Val = sheet.cell(RowNo,Coapp_State).value
            sRequest = SetXML(sRequest, "a:State",Cell_Val.encode('ascii','ignore'), 2)

            Cell_Val = sheet.cell(RowNo,Coapp_ZipCode).value
            sRequest = SetXML(sRequest, "a:ZipCode", Cell_Val.encode('ascii','ignore'), 2)

            Cell_Val = sheet.cell(RowNo,Coapp_HousingStatus).value
            sRequest = SetXML(sRequest, "a:HousingStatus", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,Coapp_MortgageOrRent).value
            sRequest = SetXML(sRequest, "a:MortgageOrRent", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,Coapp_TotalMonthsAtAddress).value
            sRequest = SetXML(sRequest, "a:TotalMonthsAtAddress", Cell_Val.encode('ascii','ignore'), 1)

            # Co-App Prev address
            Cell_Val = sheet.cell(RowNo,Coapp_Prev_AddressLine1).value
            sRequest = SetXML(sRequest, "a:AddressLine1", Cell_Val.encode('ascii','ignore'), 3)    
    
            Cell_Val = sheet.cell(RowNo,Coapp_Prev_AddressLine2).value
            sRequest = SetXML(sRequest, "a:AddressLine2", Cell_Val.encode('ascii','ignore'), 3)
    
            Cell_Val = sheet.cell(RowNo,Coapp_Prev_City).value
            sRequest = SetXML(sRequest, "a:City", Cell_Val.encode('ascii','ignore'), 3)

            Cell_Val = sheet.cell(RowNo,Coapp_Prev_State).value
            sRequest = SetXML(sRequest, "a:State", Cell_Val.encode('ascii','ignore'), 3)

            Cell_Val = sheet.cell(RowNo,Coapp_Prev_ZipCode).value
            sRequest = SetXML(sRequest, "a:ZipCode", Cell_Val.encode('ascii','ignore'), 3)

            # Co-App Current Emp
            Cell_Val = sheet.cell(RowNo,Coapp_EmploymentStatus).value
            sRequest = SetXML(sRequest, "a:EmploymentStatus", Cell_Val.encode('ascii','ignore'), 2)

            Cell_Val = sheet.cell(RowNo,Coapp_EmployedBy).value
            sRequest = SetXML(sRequest, "a:EmployedBy", Cell_Val.encode('ascii','ignore'), 2)

            Cell_Val = sheet.cell(RowNo,Coapp_TotalMonthsEmployed).value
            sRequest = SetXML(sRequest, "a:TotalMonthsEmployed", Cell_Val.encode('ascii','ignore'), 2)

            Cell_Val = sheet.cell(RowNo,Coapp_BusinessPhone).value
            sRequest = SetXML(sRequest, "a:BusinessPhone", Cell_Val.encode('ascii','ignore'), 2)

            Cell_Val = sheet.cell(RowNo,Coapp_MonthlyIncome).value
            sRequest = SetXML(sRequest, "a:MonthlyIncome", Cell_Val.encode('ascii','ignore'), 1)

            # Co-App Prev Emp
            Cell_Val = sheet.cell(RowNo,Coapp_Prev_EmploymentStatus).value
            sRequest = SetXML(sRequest, "a:EmploymentStatus", Cell_Val.encode('ascii','ignore'), 3)

            Cell_Val = sheet.cell(RowNo,Coapp_Prev_EmployedBy).value
            sRequest = SetXML(sRequest, "a:EmployedBy", Cell_Val.encode('ascii','ignore'), 3)

            Cell_Val = sheet.cell(RowNo,Coapp_Prev_TotalMonthsEmployed).value
            sRequest = SetXML(sRequest, "a:TotalMonthsEmployed", Cell_Val.encode('ascii','ignore'), 3)

            Cell_Val = sheet.cell(RowNo,Coapp_Prev_BusinessPhone).value
            sRequest = SetXML(sRequest, "a:BusinessPhone", Cell_Val.encode('ascii','ignore'), 3)

            # Co-App Others

            Cell_Val = sheet.cell(RowNo,Coapp_OtherMonthlyIncome).value
            sRequest = SetXML(sRequest, "a:OtherMonthlyIncome", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,Coapp_OtherIncomeSource).value
            sRequest = SetXML(sRequest, "a:OtherIncomeSource", Cell_Val.encode('ascii','ignore'), 1)

            Cell_Val = sheet.cell(RowNo,Relationship).value
            sRequest = SetXML(sRequest, "a:Relationship", Cell_Val.encode('ascii','ignore'), 0)

    # Applicant section is missing
    else:
        Cell_Val = sheet.cell(RowNo,Coapp_FirstName).value
        sRequest = SetXML(sRequest, "a:FirstName", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Coapp_MiddleInitial).value
        sRequest = SetXML(sRequest, "a:MiddleInitial", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Coapp_LastName).value
        sRequest = SetXML(sRequest, "a:LastName", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Coapp_Suffix).value
        sRequest = SetXML(sRequest, "a:Suffix", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Coapp_EmailAddress).value
        sRequest = SetXML(sRequest, "a:EmailAddress", Cell_Val.encode('ascii','ignore'), 0)
    
        Cell_Val = sheet.cell(RowNo,Coapp_SSN).value
        sRequest = SetXML(sRequest, "a:SSN", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Coapp_DateOfBirth).value
        sRequest = SetXML(sRequest, "a:DateOfBirth", Cell_Val.encode('ascii','ignore'), 0)
   
        Cell_Val = sheet.cell(RowNo,Coapp_DLNumber).value
        sRequest = SetXML(sRequest, "a:DriverLicenseNumber", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Coapp_DLState).value
        sRequest = SetXML(sRequest, "a:DriverLicenseState", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Coapp_HomePhone).value
        sRequest = SetXML(sRequest, "a:HomePhone", Cell_Val.encode('ascii','ignore'), 0)

        # Co-App Current address
        Cell_Val = sheet.cell(RowNo,Coapp_AddressLine1).value
        sRequest = SetXML(sRequest, "a:AddressLine1", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Coapp_AddressLine2).value
        sRequest = SetXML(sRequest, "a:AddressLine2",Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Coapp_City).value
        sRequest = SetXML(sRequest, "a:City",Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Coapp_State).value
        sRequest = SetXML(sRequest, "a:State",Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Coapp_ZipCode).value
        sRequest = SetXML(sRequest, "a:ZipCode", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Coapp_HousingStatus).value
        sRequest = SetXML(sRequest, "a:HousingStatus", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Coapp_MortgageOrRent).value
        sRequest = SetXML(sRequest, "a:MortgageOrRent", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Coapp_TotalMonthsAtAddress).value
        sRequest = SetXML(sRequest, "a:TotalMonthsAtAddress", Cell_Val.encode('ascii','ignore'), 0)

            # Co-App Prev address
        Cell_Val = sheet.cell(RowNo,Coapp_Prev_AddressLine1).value
        sRequest = SetXML(sRequest, "a:AddressLine1", Cell_Val.encode('ascii','ignore'), 1)    
    
        Cell_Val = sheet.cell(RowNo,Coapp_Prev_AddressLine2).value
        sRequest = SetXML(sRequest, "a:AddressLine2", Cell_Val.encode('ascii','ignore'), 1)
    
        Cell_Val = sheet.cell(RowNo,Coapp_Prev_City).value
        sRequest = SetXML(sRequest, "a:City", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Coapp_Prev_State).value
        sRequest = SetXML(sRequest, "a:State", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Coapp_Prev_ZipCode).value
        sRequest = SetXML(sRequest, "a:ZipCode", Cell_Val.encode('ascii','ignore'), 1)

        # Co-App Current Emp
        Cell_Val = sheet.cell(RowNo,Coapp_EmploymentStatus).value
        sRequest = SetXML(sRequest, "a:EmploymentStatus", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Coapp_EmployedBy).value
        sRequest = SetXML(sRequest, "a:EmployedBy", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Coapp_TotalMonthsEmployed).value
        sRequest = SetXML(sRequest, "a:TotalMonthsEmployed", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Coapp_BusinessPhone).value
        sRequest = SetXML(sRequest, "a:BusinessPhone", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Coapp_MonthlyIncome).value
        sRequest = SetXML(sRequest, "a:MonthlyIncome", Cell_Val.encode('ascii','ignore'), 0)

        # Co-App Prev Emp
        Cell_Val = sheet.cell(RowNo,Coapp_Prev_EmploymentStatus).value
        sRequest = SetXML(sRequest, "a:EmploymentStatus", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Coapp_Prev_EmployedBy).value
        sRequest = SetXML(sRequest, "a:EmployedBy", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Coapp_Prev_TotalMonthsEmployed).value
        sRequest = SetXML(sRequest, "a:TotalMonthsEmployed", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Coapp_Prev_BusinessPhone).value
        sRequest = SetXML(sRequest, "a:BusinessPhone", Cell_Val.encode('ascii','ignore'), 1)

        # Co-App Others

        Cell_Val = sheet.cell(RowNo,Coapp_OtherMonthlyIncome).value
        sRequest = SetXML(sRequest, "a:OtherMonthlyIncome", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Coapp_OtherIncomeSource).value
        sRequest = SetXML(sRequest, "a:OtherIncomeSource", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Relationship).value
        sRequest = SetXML(sRequest, "a:Relationship", Cell_Val.encode('ascii','ignore'), 0)
        
#    if IsApplicantSection == 'Y' and IsCoAppSection != 'Y':
        
   
    # *************************************************************************************************

    if IsVehicle_Val == 'Y':
        # Vehicle info
        Cell_Val = sheet.cell(RowNo,VehicleCondition).value
        sRequest = SetXML(sRequest, "a:VehicleCondition", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Year).value
        sRequest = SetXML(sRequest, "a:Year", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Make).value
        sRequest = SetXML(sRequest, "a:Make", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Model).value
        sRequest = SetXML(sRequest, "a:Model", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Trim).value
        sRequest = SetXML(sRequest, "a:Trim", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,ChromeStyleId).value
        sRequest = SetXML(sRequest, "a:ChromeStyleId", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,VIN).value
        sRequest = SetXML(sRequest, "a:VIN", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,StockNumber).value
        sRequest = SetXML(sRequest, "a:StockNumber", Cell_Val.encode('ascii','ignore'), 0)
    
        Cell_Val = sheet.cell(RowNo,CertifiedUsed).value
        sRequest = SetXML(sRequest, "a:CertifiedUsed", Cell_Val.encode('ascii','ignore'), 0)

    # *************************************************************************************************
    if IsProduct_Val == 'Y':
        # Product Info
        Cell_Val = sheet.cell(RowNo,CashSellingPrice).value
        sRequest = SetXML(sRequest, "a:CashSellingPrice", Cell_Val.encode('ascii','ignore'), 0)
    
        Cell_Val = sheet.cell(RowNo,SalesTax).value
        sRequest = SetXML(sRequest, "a:SalesTax", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Title).value
        sRequest = SetXML(sRequest, "a:Title", Cell_Val.encode('ascii','ignore'), 0)
    
        Cell_Val = sheet.cell(RowNo,CashDown).value
        sRequest = SetXML(sRequest, "a:CashDown", Cell_Val.encode('ascii','ignore'), 0)
    
        Cell_Val = sheet.cell(RowNo,Rebate).value
        sRequest = SetXML(sRequest, "a:Rebate", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,CreditLifeIns).value
        sRequest = SetXML(sRequest, "a:CreditLifeIns", Cell_Val.encode('ascii','ignore'), 0)
    
        Cell_Val = sheet.cell(RowNo,Term).value
        sRequest = SetXML(sRequest, "a:Term", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,AcquisitionFees).value
        sRequest = SetXML(sRequest, "a:AcquisitionFees", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,InvoiceAmount).value
        sRequest = SetXML(sRequest, "a:InvoiceAmount", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Warranty).value
        sRequest = SetXML(sRequest, "a:Warranty", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,MSRP).value
        sRequest = SetXML(sRequest, "a:MSRP", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,EstimatedBalloonAmount).value
        sRequest = SetXML(sRequest, "a:EstimatedBalloonAmount", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,EstimatedPayment).value
        sRequest = SetXML(sRequest, "a:EstimatedPayment", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,UsedCarBook).value
        sRequest = SetXML(sRequest, "a:UsedCarBook", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Mileage).value
        sRequest = SetXML(sRequest, "a:Mileage", Cell_Val.encode('ascii','ignore'), 0)
    
        Cell_Val = sheet.cell(RowNo,UsedCarValue).value
        sRequest = SetXML(sRequest, "a:UsedCarValue", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,OtherFees).value
        sRequest = SetXML(sRequest, "a:OtherFees", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,WholesaleBookSource).value
        sRequest = SetXML(sRequest, "a:WholesaleBookSource", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,WholesaleCondition).value
        sRequest = SetXML(sRequest, "a:WholesaleCondition", Cell_Val.encode('ascii','ignore'), 0)
    
        Cell_Val = sheet.cell(RowNo,WholesaleValueType).value
        sRequest = SetXML(sRequest, "a:WholesaleValueType", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,WholesaleValue).value
        sRequest = SetXML(sRequest, "a:WholesaleValue", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,NetTrade).value
        sRequest = SetXML(sRequest, "a:NetTrade", Cell_Val.encode('ascii','ignore'), 0)


    if IsVehicle_Val == 'Y' and IsTradein_Val == 'Y':
        # trade in vehicle section
        Cell_Val = sheet.cell(RowNo,Trade_Year).value
        sRequest = SetXML(sRequest, "a:Year", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Trade_Make).value
        sRequest = SetXML(sRequest, "a:Make", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Trade_Model).value
        sRequest = SetXML(sRequest, "a:Model", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Trade_Trim).value
        sRequest = SetXML(sRequest, "a:Trim", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Trade_ChromeStyleId).value
        sRequest = SetXML(sRequest, "a:ChromeStyleId", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,LienHolder).value
        sRequest = SetXML(sRequest, "a:LienHolder", Cell_Val.encode('ascii','ignore'), 0)
    
        Cell_Val = sheet.cell(RowNo,MonthlyPayment).value
        sRequest = SetXML(sRequest, "a:MonthlyPayment", Cell_Val.encode('ascii','ignore'), 0)

    # No Vehicle section only trade in section
    if IsVehicle_Val != 'Y' and IsTradein_Val == 'Y':
        # trade in vehicle section
        Cell_Val = sheet.cell(RowNo,Trade_Year).value
        sRequest = SetXML(sRequest, "a:Year", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Trade_Make).value
        sRequest = SetXML(sRequest, "a:Make", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Trade_Model).value
        sRequest = SetXML(sRequest, "a:Model", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Trade_Trim).value
        sRequest = SetXML(sRequest, "a:Trim", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Trade_ChromeStyleId).value
        sRequest = SetXML(sRequest, "a:ChromeStyleId", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,LienHolder).value
        sRequest = SetXML(sRequest, "a:LienHolder", Cell_Val.encode('ascii','ignore'), 0)
    
        Cell_Val = sheet.cell(RowNo,MonthlyPayment).value
        sRequest = SetXML(sRequest, "a:MonthlyPayment", Cell_Val.encode('ascii','ignore'), 0)
    
    Cell_Val = sheet.cell(RowNo,LeadComments).value
    sRequest = SetXML(sRequest, "LeadComments", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,PrequalificationReferenceNumber).value
    sRequest = SetXML(sRequest, "PrequalificationReferenceNumber", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,LeadReferenceNumber).value
    sRequest = SetXML(sRequest, "LeadReferenceNumber", Cell_Val.encode('ascii','ignore'), 0)

    #print sRequest
    
    #ReqFilePath = str(FolderPath) + '\\Request\\'+ WebService +'Request'+ str(RowNo) +'.xml'
    #ReqFilePath = str(FolderPath) + "\\WebService\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    ReqFilePath = str(FolderPath) + "\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    f = open(ReqFilePath, 'w')
    f.write(sRequest)
    f.close()
    
    return sRequest                     # return request string


#----------------------------------------------------------
# Description:      Function to create request XML string for FD - Credit Bureau web service for FD 1.1
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row in excel file
#                   Schemafile - File path of schema file
# Return Values:    request XML string
# Author:           Purva Wakode
#----------------------------------------------------------
def Create_FD_CreditBureau_Request_1_1(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath):

    # open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)

    # get column numbers from excel
    PartnerId = Get_ColumnNumber(HeaderList, "PartnerId")                       # Calling a function to get column no of "PartnerDealerId" column
    PartnerDealerId = Get_ColumnNumber(HeaderList, "PartnerDealerId")
    FirstName = Get_ColumnNumber(HeaderList, "FirstName")
    MiddleInitial = Get_ColumnNumber(HeaderList, "MiddleInitial")
    LastName = Get_ColumnNumber(HeaderList, "LastName")
    SSN = Get_ColumnNumber(HeaderList, "SSN")
    DateOfBirth = Get_ColumnNumber(HeaderList, "DateOfBirth")
    HomePhone = Get_ColumnNumber(HeaderList, "HomePhone")
    StreetNumber = Get_ColumnNumber(HeaderList, "StreetNumber")
    StreetName = Get_ColumnNumber(HeaderList, "StreetName")
    StreetType = Get_ColumnNumber(HeaderList, "StreetType")
    City = Get_ColumnNumber(HeaderList, "City")
    State = Get_ColumnNumber(HeaderList, "State")
    ZipCode = Get_ColumnNumber(HeaderList, "ZipCode")
    ApplicantConsent = Get_ColumnNumber(HeaderList, "ApplicantConsent")
    Co_App_FirstName = Get_ColumnNumber(HeaderList, "Co_App_FirstName")
    Co_App_MiddleInitial = Get_ColumnNumber(HeaderList, "Co_App_MiddleInitial")
    Co_App_LastName = Get_ColumnNumber(HeaderList, "Co_App_LastName")
    Co_App_SSN = Get_ColumnNumber(HeaderList, "Co_App_SSN")
    Co_App_DateOfBirth = Get_ColumnNumber(HeaderList, "Co_App_DateOfBirth")
    Co_App_HomePhone = Get_ColumnNumber(HeaderList, "Co_App_HomePhone")
    Co_App_StreetNumber = Get_ColumnNumber(HeaderList, "Co_App_StreetNumber")
    Co_App_StreetName = Get_ColumnNumber(HeaderList, "Co_App_StreetName")
    Co_App_StreetType = Get_ColumnNumber(HeaderList, "Co_App_StreetType")
    Co_App_City = Get_ColumnNumber(HeaderList, "Co_App_City")
    Co_App_State = Get_ColumnNumber(HeaderList, "Co_App_State")
    Co_App_ZipCode = Get_ColumnNumber(HeaderList, "Co_App_ZipCode")
    Co_App_Consent = Get_ColumnNumber(HeaderList, "Co_App_Consent")
    BureauType_0 = Get_ColumnNumber(HeaderList, "BureauType_0")
    BureauType_1 = Get_ColumnNumber(HeaderList, "BureauType_1")
    BureauType_2 = Get_ColumnNumber(HeaderList, "BureauType_2")
    PartnerReferenceId = Get_ColumnNumber(HeaderList, "PartnerReferenceId")

    # if nodes exists columns
    Is_App = Get_ColumnNumber(HeaderList, "Is_Applicant")
    Is_Coapp = Get_ColumnNumber(HeaderList, "Is_CoApp")

    # modified to send request to single bureau - 6/7/2013
    # Get No. of bureaus
    No_Boreaus = Get_ColumnNumber(HeaderList, "No_of_Bureaus")
    
    # Open schema file
    try:
        f = open(Schemafile)
    except IOError, e:
        print e
        print "Problem while opening Schema file"
        raise ErrorOccured("Problem in opening file: " + Schemafile)        # Problem opening file
    except:
        raise ErrorOccured("Unknown error occured.")                        # handle any other type of error

    # everything is fine
    sRequest = f.read()         # get contenets of schema file in to string
    f.close()

    # check if App node present

    Is_App_Val = sheet.cell(RowNo,Is_App).value
    Is_Coapp_Val = sheet.cell(RowNo,Is_Coapp).value
    
    #print "Original"
    #print sRequest
    
    # change the schema strings for below fields if the value is null
    # PartnerId
    Cell_Val = sheet.cell(RowNo,PartnerId).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:PartnerId></a:PartnerId>","""<a:PartnerId  i:nil="true" />""",1)
        
    # PartnerId
    Cell_Val = sheet.cell(RowNo,PartnerDealerId).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:PartnerDealerId></a:PartnerDealerId>","""<a:PartnerDealerId  i:nil="true" />""",1)

    if Is_App_Val == 'Y' and Is_Coapp_Val =='Y':      
        # FirstName
        Cell_Val = sheet.cell(RowNo,FirstName).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:FirstName></a:FirstName>","""<a:FirstName  i:nil="true" />""",1)
        
        # MiddleInitial
        Cell_Val = sheet.cell(RowNo,MiddleInitial).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:MiddleInitial></a:MiddleInitial>","""<a:MiddleInitial  i:nil="true" />""",1)
        #print "after middlename blank"
        #print sRequest
        # LastName
        Cell_Val = sheet.cell(RowNo,LastName).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:LastName></a:LastName>","""<a:LastName  i:nil="true" />""",1)
        
        # SSN
        Cell_Val = sheet.cell(RowNo,SSN).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:SSN></a:SSN>","""<a:SSN  i:nil="true" />""",1)
        
        # DateOfBirth
        Cell_Val = sheet.cell(RowNo,DateOfBirth).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:DateOfBirth></a:DateOfBirth>","""<a:DateOfBirth  i:nil="true" />""",1)

        # HomePhone  
        Cell_Val = sheet.cell(RowNo,HomePhone).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:HomePhone></a:HomePhone>","""<a:HomePhone  i:nil="true" />""",1)
        
        # StreetNumber
        Cell_Val = sheet.cell(RowNo,StreetNumber).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:StreetNumber></a:StreetNumber>","""<a:StreetNumber  i:nil="true" />""",1)
        
        # StreetName
        Cell_Val = sheet.cell(RowNo,StreetName).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:StreetName></a:StreetName>","""<a:StreetName  i:nil="true" />""",1)
            
        # StreetType
        Cell_Val = sheet.cell(RowNo,StreetType).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:StreetType></a:StreetType>","""<a:StreetType  i:nil="true" />""",1)
        
        # City
        Cell_Val = sheet.cell(RowNo,City).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:City></a:City>","""<a:City  i:nil="true" />""",1)
        
        # State
        Cell_Val = sheet.cell(RowNo,State).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:State></a:State>","""<a:State  i:nil="true" />""",1)
        
        # ZipCode
        Cell_Val = sheet.cell(RowNo,ZipCode).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:ZipCode></a:ZipCode>","""<a:ZipCode  i:nil="true" />""",1)

         # ApplicantConsent
        Cell_Val = sheet.cell(RowNo,ApplicantConsent).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:ApplicantConsent></a:ApplicantConsent>","""<a:ApplicantConsent  i:nil="true" />""",1)
   
        # Co_App_FirstName
        Cell_Val = sheet.cell(RowNo,Co_App_FirstName).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:FirstName></a:FirstName>","""<a:FirstName  i:nil="true" />""",2)

        #print "------------befor co-app midlle name blank----------------------------------------"
        #print sRequest
        #print "----------------------------------------------------"
        
        # Co_App_MiddleInitial
        App_MiddleName = sheet.cell(RowNo,MiddleInitial).value
        
        Cell_Val = sheet.cell(RowNo,Co_App_MiddleInitial).value
        if len(Cell_Val) == 0 and len(App_MiddleName) !=0:
            sRequest = replacenth(sRequest,"<a:MiddleInitial></a:MiddleInitial>","""<a:MiddleInitial  i:nil="true" />""",2)
        if len(Cell_Val) == 0 and len(App_MiddleName) ==0:
            sRequest = replacenth(sRequest,"<a:MiddleInitial></a:MiddleInitial>","""<a:MiddleInitial  i:nil="true" />""",1)

        #print "----------------------------------------------------"
        #print sRequest
        #print "----------------------------------------------------"
        
        # Co_App_LastName
        Cell_Val = sheet.cell(RowNo,Co_App_LastName).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:LastName></a:LastName>","""<a:LastName  i:nil="true" />""",2)
        
        # Co_App_SSN
        Cell_Val = sheet.cell(RowNo,Co_App_SSN).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:SSN></a:SSN>","""<a:SSN  i:nil="true" />""",2)
        
        # Co_App_DateOfBirth
        Cell_Val = sheet.cell(RowNo,Co_App_DateOfBirth).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:DateOfBirth></a:DateOfBirth>","""<a:DateOfBirth  i:nil="true" />""",2)
        
        # Co_App_HomePhone
        Cell_Val = sheet.cell(RowNo,Co_App_HomePhone).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:HomePhone></a:HomePhone>","""<a:HomePhone  i:nil="true" />""",2)
        
        # Co_App_StreetNumber
        Cell_Val = sheet.cell(RowNo,Co_App_StreetNumber).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:StreetNumber></a:StreetNumber>","""<a:StreetNumber  i:nil="true" />""",2)
        
        # Co_App_StreetName
        Cell_Val = sheet.cell(RowNo,Co_App_StreetName).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:StreetName></a:StreetName>","""<a:StreetName  i:nil="true" />""",2)
        
        # Co_App_StreetType
        Cell_Val = sheet.cell(RowNo,Co_App_StreetType).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:StreetType></a:StreetType>","""<a:StreetType  i:nil="true" />""",2)

        # Co_App_City
        Cell_Val = sheet.cell(RowNo,Co_App_City).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:City></a:City>","""<a:City  i:nil="true" />""",2)

        # Co_App_State
        Cell_Val = sheet.cell(RowNo,Co_App_State).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:State></a:State>","""<a:State  i:nil="true" />""",2)

        # Co_App_ZipCode
        Cell_Val = sheet.cell(RowNo,Co_App_ZipCode).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:ZipCode></a:ZipCode>","""<a:ZipCode  i:nil="true" />""",2)

        # Co_App_Consent
        Cell_Val = sheet.cell(RowNo,Co_App_Consent).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:CoApplicantConsent></a:CoApplicantConsent>","""<a:CoApplicantConsent  i:nil="true" />""",1)
    
        print sRequest

    # if No Applicant and only Co-App
    elif Is_App_Val != 'Y' and Is_Coapp_Val=='Y':      
        sRequest = sRequest.replace("""<PrimaryApplicant xmlns:a="http://schemas.datacontract.org/2004/07/DealerTrack.DataContracts.CustomerEnquiry">
<a:FirstName></a:FirstName>
<a:MiddleInitial></a:MiddleInitial>
<a:LastName></a:LastName>
<a:SSN></a:SSN>
<a:DateOfBirth></a:DateOfBirth>
<a:HomePhone></a:HomePhone>
<a:StreetNumber></a:StreetNumber>
<a:StreetName></a:StreetName>
<a:StreetType></a:StreetType>
<a:City></a:City>
<a:State></a:State>
<a:ZipCode></a:ZipCode>
</PrimaryApplicant>
<ApplicantConsent></ApplicantConsent>""","",1)
        print "*** Request without Applicant Section ***"
        print sRequest

    # if only Applicant and No Co-App
    elif Is_App_Val == 'Y' and Is_Coapp_Val!='Y':
        sRequest = sRequest.replace("""<CoApplicant xmlns:a="http://schemas.datacontract.org/2004/07/DealerTrack.DataContracts.CustomerEnquiry">
<a:FirstName></a:FirstName>
<a:MiddleInitial></a:MiddleInitial>
<a:LastName></a:LastName>
<a:SSN></a:SSN>
<a:DateOfBirth></a:DateOfBirth>
<a:HomePhone></a:HomePhone>
<a:StreetNumber></a:StreetNumber>
<a:StreetName></a:StreetName>
<a:StreetType></a:StreetType>
<a:City></a:City>
<a:State></a:State>
<a:ZipCode></a:ZipCode>
</CoApplicant>
<CoApplicantConsent></CoApplicantConsent>""","",1)
        print "*** Request without Co-Applicant Section ***"
        print sRequest

    #if both are not present
    elif Is_App_Val != 'Y' and Is_Coapp_Val!='Y':
        sRequest = sRequest.replace("""<PrimaryApplicant xmlns:a="http://schemas.datacontract.org/2004/07/DealerTrack.DataContracts.CustomerEnquiry">
<a:FirstName></a:FirstName>
<a:MiddleInitial></a:MiddleInitial>
<a:LastName></a:LastName>
<a:SSN></a:SSN>
<a:DateOfBirth></a:DateOfBirth>
<a:HomePhone></a:HomePhone>
<a:StreetNumber></a:StreetNumber>
<a:StreetName></a:StreetName>
<a:StreetType></a:StreetType>
<a:City></a:City>
<a:State></a:State>
<a:ZipCode></a:ZipCode>
</PrimaryApplicant>
<ApplicantConsent></ApplicantConsent>
<CoApplicant xmlns:a="http://schemas.datacontract.org/2004/07/DealerTrack.DataContracts.CustomerEnquiry">
<a:FirstName></a:FirstName>
<a:MiddleInitial></a:MiddleInitial>
<a:LastName></a:LastName>
<a:SSN></a:SSN>
<a:DateOfBirth></a:DateOfBirth>
<a:HomePhone></a:HomePhone>
<a:StreetNumber></a:StreetNumber>
<a:StreetName></a:StreetName>
<a:StreetType></a:StreetType>
<a:City></a:City>
<a:State></a:State>
<a:ZipCode></a:ZipCode>
</CoApplicant>
<CoApplicantConsent></CoApplicantConsent>""","",1)
        print "*** Request without Applicant and CoApp Section ***"
        print sRequest

    # modified to send request to single bureau - 6/7/2013
    Cell_Val = sheet.cell(RowNo,No_Boreaus).value
#    print "no. of bureaus: ",Cell_Val
    if Cell_Val == '3':
        #print "in case for 3"
        # BureauType_0
        Cell_Val = sheet.cell(RowNo,BureauType_0).value
        if len(Cell_Val) == 0:
            sRequest = sRequest.replace("<a:BureauType></a:BureauType>","""<a:BureauType  i:nil="true" />""",1)

        # BureauType_1
        Cell_Val = sheet.cell(RowNo,BureauType_1).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:BureauType></a:BureauType>","""<a:BureauType  i:nil="true" />""",1)

        # BureauType_2
        Cell_Val = sheet.cell(RowNo,BureauType_2).value
        if len(Cell_Val) == 0:
            sRequest = replacenth(sRequest,"<a:BureauType></a:BureauType>","""<a:BureauType  i:nil="true" />""",1)

    elif Cell_Val == '1':
        print "in case for 1 bureau"
        sRequest = sRequest.replace("""<a:BureauType></a:BureauType>
<a:BureauType></a:BureauType>
<a:BureauType></a:BureauType>""","""<a:BureauType></a:BureauType>""",1)

    elif Cell_Val == '0':
        #print "in case for 1 bureau"
        sRequest = sRequest.replace("""<a:BureauType></a:BureauType>
<a:BureauType></a:BureauType>
<a:BureauType></a:BureauType>""","""<a:BureauType>None</a:BureauType>""",1)

    #print sRequest

    
    # PartnerReferenceId
    Cell_Val = sheet.cell(RowNo,PartnerReferenceId).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:PartnerReferenceId></a:PartnerReferenceId>","""<a:PartnerReferenceId  i:nil="true" />""",1)
   
    # Create request string
    # Calling function to add values to nodes in given schema file

    Cell_Val = sheet.cell(RowNo,PartnerId).value
    sRequest = SetXML(sRequest, "PartnerId", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,PartnerDealerId).value
    sRequest = SetXML(sRequest, "PartnerDealerId",Cell_Val.encode('ascii','ignore') , 0)

    if Is_App_Val == 'Y' and Is_Coapp_Val == 'Y':
        Cell_Val = sheet.cell(RowNo,FirstName).value
        sRequest = SetXML(sRequest, "a:FirstName", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,MiddleInitial).value
        sRequest = SetXML(sRequest, "a:MiddleInitial", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,LastName).value
        sRequest = SetXML(sRequest, "a:LastName", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,SSN).value
        sRequest = SetXML(sRequest, "a:SSN", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,DateOfBirth).value
        sRequest = SetXML(sRequest, "a:DateOfBirth", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,HomePhone).value
        sRequest = SetXML(sRequest, "a:HomePhone",Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,StreetNumber).value
        sRequest = SetXML(sRequest, "a:StreetNumber",Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,StreetName).value
        sRequest = SetXML(sRequest, "a:StreetName",Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,StreetType).value
        sRequest = SetXML(sRequest, "a:StreetType", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,City).value
        sRequest = SetXML(sRequest, "a:City", Cell_Val.encode('ascii','ignore'), 0)
    
        Cell_Val = sheet.cell(RowNo,State).value 
        sRequest = SetXML(sRequest, "a:State", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,ZipCode).value
        sRequest = SetXML(sRequest, "a:ZipCode", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,ApplicantConsent).value
        sRequest = SetXML(sRequest, "ApplicantConsent", Cell_Val.encode('ascii','ignore'), 0)

        
        Cell_Val = sheet.cell(RowNo,Co_App_FirstName).value
        sRequest = SetXML(sRequest, "a:FirstName", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Co_App_MiddleInitial).value
        sRequest = SetXML(sRequest, "a:MiddleInitial", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Co_App_LastName).value
        sRequest = SetXML(sRequest, "a:LastName", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Co_App_SSN).value
        sRequest = SetXML(sRequest, "a:SSN", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Co_App_DateOfBirth).value
        sRequest = SetXML(sRequest, "a:DateOfBirth", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Co_App_HomePhone).value
        sRequest = SetXML(sRequest, "a:HomePhone", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Co_App_StreetNumber).value
        sRequest = SetXML(sRequest, "a:StreetNumber", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Co_App_StreetName).value
        sRequest = SetXML(sRequest, "a:StreetName", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Co_App_StreetType).value
        sRequest = SetXML(sRequest, "a:StreetType", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Co_App_City).value
        sRequest = SetXML(sRequest, "a:City", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Co_App_State).value
        sRequest = SetXML(sRequest, "a:State", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Co_App_ZipCode).value
        sRequest = SetXML(sRequest, "a:ZipCode", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,Co_App_Consent).value
        sRequest = SetXML(sRequest, "CoApplicantConsent", Cell_Val.encode('ascii','ignore'), 0)

        print "App = Y, CoApp=Y"
        print sRequest
    # Applicant section is missing
    elif Is_App_Val != 'Y' and Is_Coapp_Val == 'Y':
        Cell_Val = sheet.cell(RowNo,Co_App_FirstName).value
        sRequest = SetXML(sRequest, "a:FirstName", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Co_App_MiddleInitial).value
        sRequest = SetXML(sRequest, "a:MiddleInitial", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Co_App_LastName).value
        sRequest = SetXML(sRequest, "a:LastName", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Co_App_SSN).value
        sRequest = SetXML(sRequest, "a:SSN", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Co_App_DateOfBirth).value
        sRequest = SetXML(sRequest, "a:DateOfBirth", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Co_App_HomePhone).value
        sRequest = SetXML(sRequest, "a:HomePhone", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Co_App_StreetNumber).value
        sRequest = SetXML(sRequest, "a:StreetNumber", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Co_App_StreetName).value
        sRequest = SetXML(sRequest, "a:StreetName", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Co_App_StreetType).value
        sRequest = SetXML(sRequest, "a:StreetType", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Co_App_City).value
        sRequest = SetXML(sRequest, "a:City", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Co_App_State).value
        sRequest = SetXML(sRequest, "a:State", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Co_App_ZipCode).value
        sRequest = SetXML(sRequest, "a:ZipCode", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,Co_App_Consent).value
        sRequest = SetXML(sRequest, "CoApplicantConsent", Cell_Val.encode('ascii','ignore'), 0)

    elif Is_App_Val == 'Y' and Is_Coapp_Val != 'Y':
        Cell_Val = sheet.cell(RowNo,FirstName).value
        sRequest = SetXML(sRequest, "a:FirstName", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,MiddleInitial).value
        sRequest = SetXML(sRequest, "a:MiddleInitial", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,LastName).value
        sRequest = SetXML(sRequest, "a:LastName", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,SSN).value
        sRequest = SetXML(sRequest, "a:SSN", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,DateOfBirth).value
        sRequest = SetXML(sRequest, "a:DateOfBirth", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,HomePhone).value
        sRequest = SetXML(sRequest, "a:HomePhone",Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,StreetNumber).value
        sRequest = SetXML(sRequest, "a:StreetNumber",Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,StreetName).value
        sRequest = SetXML(sRequest, "a:StreetName",Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,StreetType).value
        sRequest = SetXML(sRequest, "a:StreetType", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,City).value
        sRequest = SetXML(sRequest, "a:City", Cell_Val.encode('ascii','ignore'), 0)
    
        Cell_Val = sheet.cell(RowNo,State).value 
        sRequest = SetXML(sRequest, "a:State", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,ZipCode).value
        sRequest = SetXML(sRequest, "a:ZipCode", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,ApplicantConsent).value
        sRequest = SetXML(sRequest, "ApplicantConsent", Cell_Val.encode('ascii','ignore'), 0)
        
    Cell_Val = sheet.cell(RowNo,No_Boreaus).value
    print "no. of bureaus: ",Cell_Val
    if Cell_Val == '3':
        Cell_Val = sheet.cell(RowNo,BureauType_0).value
        sRequest = SetXML(sRequest, "a:BureauType", Cell_Val.encode('ascii','ignore'), 0)

        Cell_Val = sheet.cell(RowNo,BureauType_1).value
        sRequest = SetXML(sRequest, "a:BureauType", Cell_Val.encode('ascii','ignore'), 1)

        Cell_Val = sheet.cell(RowNo,BureauType_2).value
        sRequest = SetXML(sRequest, "a:BureauType", Cell_Val.encode('ascii','ignore'), 2)
    elif Cell_Val == '1':
        Cell_Val = sheet.cell(RowNo,BureauType_0).value
        sRequest = SetXML(sRequest, "a:BureauType", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,PartnerReferenceId).value
    sRequest = SetXML(sRequest, "PartnerReferenceId", Cell_Val.encode('ascii','ignore'), 0)

    
    #print sRequest
    
    #ReqFilePath = str(FolderPath) + '\\Request\\'+ WebService +'Request'+ str(RowNo) +'.xml'
    #ReqFilePath = str(FolderPath) + "\\WebService\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    ReqFilePath = str(FolderPath) + "\\Request\\"+ WebService +'Request'+ str(RowNo+1) +'.xml'
    f = open(ReqFilePath, 'w')
    f.write(sRequest)
    f.close()
    
    return sRequest                     # return request string
    


#----------------------------------------------------------
# Description:      Function to send request XML to required web service 
# Input Parameters: WebService - Name of WebService e.g. FD_PreQual
#                   sRequest - Request XML string
#                   Username, Password - Authenticated username and password to send request
# Return Value:     Received response code and response XML string
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Send_Request(FileName, RowNo, WebService,sRequest,Username,Password,url,sSoapAction, FolderPath):
    # Open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)                                  # select required worksheet

    # Calling a function to create list of headers from excel file
    HeaderList = []
    HeaderList = Generate_HeaderList(FileName)


    # To check whether the WebService is "FD_Credit_Bureau" and if Yes then change the soap action according to Red Flag Status values in excel file.
    if WebService == "FD_CreditBureau_1_1":
        Flag_Status = Get_ExcelValue(FileName, RowNo, "REDFLAGSTATUS")

        if Flag_Status == "Yes":
            sSoapAction = "/CreditBureau/V1.1/?redFlagRequest=Yes"
        else:
            sSoapAction = "/CreditBureau/V1.1/?redFlagRequest=No"
#    print sSoapAction
    
    StepResult = "Not Executed"                                             # initialising step result
        
    if WebService == 'FD_Prequal_1_0' or WebService == 'FD_Prequal_1_1' or WebService == 'FD_Lead_1_1' or WebService == 'FD_CreditApp_1_1' or WebService == 'FD_CreditBureau_1_1' or WebService == 'Phase_I_FD_Lead_1_1' or WebService == 'Phase_I_FD_CreditBureau_1_1' or WebService == 'Phase_I_FD_CreditApp_1_1' or WebService == 'Phase_I_FD_Prequal_1_1':
        # details for FD Prequalificarion web service
        print "Web Service: ",WebService
        url = url + sSoapAction
        headers = {'Content-Type':'text/xml;charset=utf-8',
                   'User-Agent':'Agent: Jakarta Commons-HttpClient/3.1'}
    else:
        headers = {'Content-Type': 'text/xml;charset=UTF-8',
                    'SOAPAction': sSoapAction}

	#####################################################
    # Common code for sending request
    request = urllib2.Request(url,sRequest,headers)
    base64string = base64.encodestring('%s:%s' % (Username, Password))[:-1]
    authheader =  "Basic %s" % base64string
    request.add_header("Authorization", authheader)
    #RespFile = str(FolderPath)+'\\Response\\'+ WebService +'Response'+str(RowNo)+'.xml'
    #RespFile = str(FolderPath) + "\\WebService\\Response\\"+ WebService +'Response'+ str(RowNo) +'.xml'
    #RespFile = str(FolderPath) + "\\Response\\"+ WebService +'Response'+ str(RowNo) +'.xml' #Jenkins Adjustment
    RespFile = str(FolderPath) + "//Response//"+ WebService +'Response'+ str(RowNo) +'.xml' #Jenkins Adjustment
    print "Response file name: " + RespFile
    flag = 0
    try:
        response = urllib2.urlopen(request)                         # sending request
    except urllib2.URLError, e:                                     # error code
        if e.code == 404:
            StepResult = "Executed"
            LogToExcel(FileName,RowNo,"Send_Request_Result",StepResult)         # calling function to write result of step to excel file
            print e.read()
            f = open(RespFile, 'w')
            f.write(e.read())
            f.close()
            raise ErrorOccured("URL Error occurred while sending request: " + e.read())
        elif e.code == 401:
            StepResult = "Executed"
            LogToExcel(FileName,RowNo,"Send_Request_Result",StepResult)         # calling function to write result of step to excel file
            print e.read()
            f = open(RespFile, 'w')
            f.write(e.read())
            f.close()
            raise ErrorOccured("Unauthorized Access: " + e.read())
        else:
            StepResult = "Executed"
            LogToExcel(FileName,RowNo,"Send_Request_Result",StepResult)         # calling function to write result of step to excel file
            StatusCode = e.code
            SResponseText = e.read()
            f = open(RespFile, 'w')
            f.write(SResponseText)
            f.close()
            print SResponseText            
            flag = 1
    except HTTPError, e:
        StepResult = "Executed"
        LogToExcel(FileName,RowNo,"Send_Request_Result",StepResult)
        print e.read()
        f = open(RespFile, 'w')
        f.write(e.read())
        f.close()
        raise ErrorOccured("You have received HTTP error: " +  e.read())
    except:
        raise ErrorOccured("Unknown error occurred." )

    # Everything is fine, you got valid response
    if flag == 0:
        StatusCode = response.getcode()                             # get status code
        SResponseText = response.read()                             # get response XML text
        StepResult = "Executed"
        print 'You have received valid response'
        print 'Status Code:',StatusCode
        print 'Response Text: ',SResponseText
        print StepResult
        LogToExcel(FileName,RowNo,"Send_Request_Result",StepResult) # calling function to write result of step to excel file

    # write response to a file
    f = open(RespFile, 'w')
    f.write(SResponseText)
    f.close()
    
    return StatusCode,SResponseText                             # return response code and response text
#----------------------------------------------------------
# Description:      Function to send request XML to required web service with given Request XML
# Input Parameters: WebService - Request XML & Hearder as well authentication information
#                   sRequest - Request XML string
#                   Username, Password - Authenticated username and password to send request
# Return Value:     String values of Response Code, Error Text & Response XML
# Author:           Naveed Hasan
#----------------------------------------------------------
def Send_Request_WithRequestXML(WebService,url,sSoapAction,sRequest,Username,Password,FolderPath):
    
    StepResult = "Not Executed"                                             
    ErrorText = "Nothing"
    headers = {'Content-Type': 'text/xml;charset=UTF-8',
                    'SOAPAction': sSoapAction}
    request = urllib2.Request(url,sRequest,headers)
    base64string = base64.encodestring('%s:%s' % (Username, Password))[:-1]
    authheader =  "Basic %s" % base64string
    request.add_header("Authorization", authheader)
    #RespFile = str(FolderPath) + "\\Response\\"+ WebService +'Response'+ str(RowNo) +'.xml'
    RespFile = str(FolderPath) + "\\Response\\"+ WebService +'Response'+'.xml'
    print "Response file name: " + RespFile
    flag = 0
    try:
        response = urllib2.urlopen(request)                         # sending request
    except urllib2.URLError as e:                                     # error code
        if e.code == 404:
            StepResult = "Executed"
            #LogToExcel(FileName,RowNo,"Send_Request_Result",StepResult)         # calling function to write result of step to excel file
            print e.read()
            # f = open(RespFile, 'w')
            # f.write(e.read())
            # f.close()
            ErrorText = e.read()
            print ErrorText
            raise ErrorOccured("URL Error occurred while sending request: " + e.read())
        elif e.code == 401:
            StepResult = "Executed"
            #LogToExcel(FileName,RowNo,"Send_Request_Result",StepResult)         # calling function to write result of step to excel file
            print e.read()
            # f = open(RespFile, 'w')
            # f.write(e.read())
            # f.close()
            ErrorText = e.read()
            print ErrorText 
            raise ErrorOccured("Unauthorized Access: " + e.read())
        else:
            StepResult = "Executed"
            #LogToExcel(FileName,RowNo,"Send_Request_Result",StepResult)         # calling function to write result of step to excel file
            StatusCode = e.code
            SResponseText = e.read()
            # f = open(RespFile, 'w')
            # f.write(SResponseText)
            # f.close()
            print SResponseText            
            flag = 1
    except urllib2.HTTPError as e:
    #except e:
        StepResult = "Executed"
        #LogToExcel(FileName,RowNo,"Send_Request_Result",StepResult)
        print e.read()
        # f = open(RespFile, 'w')
        # f.write(e.read())
        # f.close()
        ErrorText = str(e)#e.read()
        print ErrorText
        raise ErrorOccured("You have received HTTP error: " +  e.read())
    except Exception as e:
        ErrorText = str(e)
        print ErrorText
        raise ErrorOccured("Unknown error occurred." + ErrorText )

    # Everything is fine, you got valid response
    if flag == 0:
        StatusCode = response.getcode()                             # get status code
        SResponseText = response.read()                             # get response XML text
        StepResult = "Executed"
        #print 'You have received valid response'
        #print 'Status Code:',StatusCode
        #print 'Response Text: ',SResponseText
        #print StepResult
        #LogToExcel(FileName,RowNo,"Send_Request_Result",StepResult) # calling function to write result of step to excel file

    # write response to a file
    f = open(RespFile, 'w')
    f.write(SResponseText)
    f.close()
    
    return StatusCode,ErrorText,SResponseText                             # return response code and response text
	
#----------------------------------------------------------
# Description:      Function to send request XML to required web service 
# Input Parameters: WebService - Name of WebService e.g. FD_PreQual
#                   sRequest - Request XML string
#                   Username, Password - Authenticated username and password to send request
# Return Value:     Received response code and response XML string
# Author:           Naveed Hasan
#----------------------------------------------------------
def Send_Request_WithRequestXML_FD(url,sSoapAction,sRequest,Username,Password,FolderPath):
    print url
    print sSoapAction
    print Username
    print Password
    # Open excel file
    #book = Open_Excel(FileName)
    #sheet = book.sheet_by_index(0)                                  # select required worksheet

    # Calling a function to create list of headers from excel file
    #HeaderList = []
    #HeaderList = Generate_HeaderList(FileName)


    # To check whether the WebService is "FD_Credit_Bureau" and if Yes then change the soap action according to Red Flag Status values in excel file.
    # if WebService == "FD_CreditBureau_1_1":
        # Flag_Status = Get_ExcelValue(FileName, RowNo, "REDFLAGSTATUS")

        # if Flag_Status == "Yes":
            # sSoapAction = "/CreditBureau/V1.1/?redFlagRequest=Yes"
        # else:
            # sSoapAction = "/CreditBureau/V1.1/?redFlagRequest=No"
#    print sSoapAction
    StepResult = "Not Executed"                                             
    ErrorText = "Nothing"
    #url = "https://webservices.qa.dealertrack.com"
    #sSoapAction = "/financedriver/v1.1/applicant/prequalify"
    StepResult = "Not Executed"                                             # initialising step result
        
   # if WebService == 'FD_Prequal_1_0' or WebService == 'FD_Prequal_1_1' or WebService == 'FD_Lead_1_1' or WebService == 'FD_CreditApp_1_1' or WebService == 'FD_CreditBureau_1_1' or WebService == 'Phase_I_FD_Lead_1_1' or WebService == 'Phase_I_FD_CreditBureau_1_1' or WebService == 'Phase_I_FD_CreditApp_1_1' or WebService == 'Phase_I_FD_Prequal_1_1':
        # details for FD Prequalificarion web service
        # print "Web Service: ",WebService
    url = url + sSoapAction
    print url
        #
    #headers = {'Content-Type':'text/xml;charset=utf-8','User-Agent':'Agent: Jakarta Commons-HttpClient/3.1'}
    # else:
    # headers = {'Content-Type': 'text/xml;charset=UTF-8',
                    # 'SOAPAction': sSoapAction}

    headers = {'Content-Type': 'application/xml','Host':'webservices.qa.dealertrack.com',
                     'User-Agent': 'Apache-HttpClient/4.11(java 1.5)'}


	#####################################################
    # Common code for sending request
    request = urllib2.Request(url,sRequest,headers)
    base64string = base64.encodestring('%s:%s' % (Username, Password))[:-1]
    authheader =  "Basic %s" % base64string
    request.add_header("Authorization", authheader)
    #RespFile = str(FolderPath)+'\\Response\\'+ WebService +'Response'+str(RowNo)+'.xml'
    #RespFile = str(FolderPath) + "\\WebService\\Response\\"+ WebService +'Response'+ str(RowNo) +'.xml'
    #RespFile = str(FolderPath) + "\\Response\\"+ WebService +'Response'+ str(RowNo) +'.xml'
	#RespFile = str(FolderPath) + "\\Response\\"+ WebService +'Response'+'.xml'
    #print "Response file name: " + RespFile
    flag = 0
    try:
        response = urllib2.urlopen(request)                         # sending request
    except urllib2.URLError, e:                                     # error code
        if e.code == 404:
            StepResult = "Executed"
            #LogToExcel(FileName,RowNo,"Send_Request_Result",StepResult)         # calling function to write result of step to excel file
            print e.read()
            #f = open(RespFile, 'w')
            # f.write(e.read())
            # f.close()
            raise ErrorOccured("URL Error occurred while sending request: " + e.read())
        elif e.code == 401:
            StepResult = "Executed"
            #LogToExcel(FileName,RowNo,"Send_Request_Result",StepResult)         # calling function to write result of step to excel file
            print e.read()
            #f = open(RespFile, 'w')
            # f.write(e.read())
            # f.close()
            raise ErrorOccured("Unauthorized Access: " + e.read())
        else:
            StepResult = "Executed"
            #LogToExcel(FileName,RowNo,"Send_Request_Result",StepResult)         # calling function to write result of step to excel file
            StatusCode = e.code
            SResponseText = e.read()
            #f = open(RespFile, 'w')
            # f.write(SResponseText)
            # f.close()
            print SResponseText            
            flag = 1
    except HTTPError, e:
        StepResult = "Executed"
        #LogToExcel(FileName,RowNo,"Send_Request_Result",StepResult)
        print e.read()
        #f = open(RespFile, 'w')
        # f.write(e.read())
        # f.close()
        raise ErrorOccured("You have received HTTP error: " +  e.read())
    except:
        raise ErrorOccured("Unknown error occurred." )

    # Everything is fine, you got valid response
    if flag == 0:
        StatusCode = response.getcode()                             # get status code
        SResponseText = response.read()                             # get response XML text
        StepResult = "Executed"
        print 'You have received valid response'
        print 'Status Code:',StatusCode
        print 'Response Text: ',SResponseText
        print StepResult
        #LogToExcel(FileName,RowNo,"Send_Request_Result",StepResult) # calling function to write result of step to excel file

    # write response to a file
    #f = open(RespFile, 'w')
    # f.write(SResponseText)
    # f.close()
    
    return StatusCode,ErrorText,SResponseText                             # return response code and response text
	
	
def Send_Request_WithRequestXML_ADFLead(url,sSoapAction,sRequest,Username,Password):
    print url
    print sSoapAction
    print Username
    print Password

    StepResult = "Not Executed"                                             
    ErrorText = "Nothing"
    StepResult = "Not Executed"                                             # initialising step result
        
    url = url + sSoapAction
    print url

    headers = {'Content-Type': 'text/plain'}

	#####################################################
    # Common code for sending request
    request = urllib2.Request(url,sRequest,headers)
    base64string = base64.encodestring('%s:%s' % (Username, Password))[:-1]
    authheader =  "Basic %s" % base64string
    request.add_header("Authorization", authheader)
    flag = 0
    try:
        response = urllib2.urlopen(request)                         # sending request
    except urllib2.URLError, e:                                     # error code
        if e.code == 404:
            StepResult = "Executed"
            print e.read()
            raise ErrorOccured("URL Error occurred while sending request: " + e.read())
        elif e.code == 401:
            StepResult = "Executed"
            print e.read()
            raise ErrorOccured("Unauthorized Access: " + e.read())
        else:
            StepResult = "Executed"
            StatusCode = e.code
            SResponseText = e.read()
            print SResponseText            
            flag = 1
    except HTTPError, e:
        StepResult = "Executed"
        print e.read()
        raise ErrorOccured("You have received HTTP error: " +  e.read())
    except:
        raise ErrorOccured("Unknown error occurred." )

    # Everything is fine, you got valid response
    if flag == 0:
        StatusCode = response.getcode()                             # get status code
        SResponseText = response.read()                             # get response XML text
        StepResult = "Executed"
        print 'You have received valid response'
        print 'Status Code:',StatusCode
        print 'Response Text: ',SResponseText
        print StepResult
    
    return StatusCode,ErrorText,SResponseText                             # return response code and response text	
	
	
def Send_Request_WithRequestXML_TD(url,sSoapAction,sRequest,Username,Password,FolderPath):
    print url
    print sSoapAction
    print Username
    print Password
    ErrorText = "Nothing"
    StepResult = "Not Executed"                                             # initialising step result        
    url = url + sSoapAction
    print url
        #
    headers = {'Content-Type':'application/xml','Cookie':'SMCHALLENGE=YES'}

	#####################################################
    # Common code for sending request
    request = urllib2.Request(url,sRequest,headers)
    base64string = base64.encodestring('%s:%s' % (Username, Password))[:-1]
    authheader =  "Basic %s" % base64string
    request.add_header("Authorization", authheader)
    flag = 0
    try:
        response = urllib2.urlopen(request)                         # sending request
    except urllib2.URLError, e:                                     # error code
        if e.code == 404:
            StepResult = "Executed"
            print e.read()
            raise ErrorOccured("URL Error occurred while sending request: " + e.read())
        elif e.code == 401:
            StepResult = "Executed"
            #LogToExcel(FileName,RowNo,"Send_Request_Result",StepResult)         # calling function to write result of step to excel file
            print e.read()
            raise ErrorOccured("Unauthorized Access: " + e.read())
        else:
            StepResult = "Executed"
            StatusCode = e.code
            SResponseText = e.read()
            print SResponseText            
            flag = 1
    except HTTPError, e:
        StepResult = "Executed"
        print e.read()
        raise ErrorOccured("You have received HTTP error: " +  e.read())
    except:
        raise ErrorOccured("Unknown error occurred." )

    # Everything is fine, you got valid response
    if flag == 0:
        StatusCode = response.getcode()                             # get status code
        SResponseText = response.read()                             # get response XML text
        StepResult = "Executed"
        print 'You have received valid response'
        print 'Status Code:',StatusCode
        print 'Response Text: ',SResponseText
        print StepResult
     
    return StatusCode,ErrorText,SResponseText                             # return response code and response text

	#----------------------------------------------------------
# Description:      Function to log result of request send to excel file for current row
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row number in excel file
#                   StepResult - Result to log "PASS" or "FAIL"
# Author:           Manisha Gadekar
#----------------------------------------------------------
def LogToExcel(FileName,RowNo,ColumnName,CellValue):
    # open excel file
    book = Open_Excel(FileName)
    # Calling a function to create list of headers from excel file
    HeaderList = []
    HeaderList = Generate_HeaderList(FileName)
    
    Result_col = Get_ColumnNumber(HeaderList, ColumnName)   # Calling a function to get column no of "Send_Request_Result" column
    
    workbook = copy(book)                                   # using copy utility to write to excel file
    w_sheet = workbook.get_sheet(0)                         # locating exact worksheet

    w_sheet.write(RowNo,Result_col,CellValue)               # using write method to write in cell
    
    workbook.save(FileName)                                 # save changes to excel file

def Write_To_Excel(FileName,RowNo,ColNo,CellValue):
    # open excel file
    book = Open_Excel(FileName)
    # Calling a function to create list of headers from excel file
    # HeaderList = []
    # HeaderList = Generate_HeaderList(FileName)
    
    # Result_col = Get_ColumnNumber(HeaderList, ColumnName)   # Calling a function to get column no of "Send_Request_Result" column
    style = easyxf('font: bold 1, color green;')
    workbook = copy(book)                                   # using copy utility to write to excel file
    w_sheet = workbook.get_sheet(0)                         # locating exact worksheet

    w_sheet.write(RowNo,ColNo,CellValue,style)               # using write method to write in cell
    
    workbook.save(FileName)                                 # save changes to excel file

        
#----------------------------------------------------------
# Description:      Function to validate response received
# Input Parameters: WebService - Name of WebService e.g. FD_PreQual
#                   FileName - Path of excel file
#                   RowNo - Current row number in excel file
#                   sResponseCode - Status code received after sending request
#                   sResponseText - Response text received after sending request. This is response to be validated.
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Validate_Response(WebService, FileName, RowNo, sResponseCode,sResponseText):
    #return WebService   # Here, actual and expected responses will be compared
    if WebService == 'PD_GetVehicle' or 'PD_GetVehicleByChromeStyleId':
        sResult = Validate_Response_PDGetV(WebService,FileName, RowNo, sResponseCode, sResponseText)
    elif WebService == 'PD_GetPayments':
        # details for other web service
        print WebService
    return sResult

#----------------------------------------------------------
# Description:      Function to validate response received for PD_Get Vehicles web service
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row number in excel file
#                   sResponseCode - Response code received after sending request. This is response to be validated.
#                   sResponseText - Response text received after sending request. This is response to be validated.
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Validate_Response_PDGetV(WebService,FileName, RowNo, sResponseCode, sResponseText):
    Result = "PASS"

    # Open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)                                  # select required worksheet

    # Calling a function to create list of headers from excel file
    HeaderList = []
    HeaderList = Generate_HeaderList(FileName)
    
    if sResponseCode == 200:    # response code received is 200. 
        if WebService=='PD_GetVehicle':
            ResponseSchemaFile = "C:\\python27\\DRS - Automation\\GetVehicleResponseSchema.xls" #response schema file name
            IncludeRebates_col = Get_ColumnNumber(HeaderList, "IncludeRebates")                   
            IncludeRebates_Val = sheet.cell(RowNo,IncludeRebates_col).value
        elif WebService=='PD_GetVehicleByChromeStyleId':
            ResponseSchemaFile = "C:\\python27\\DRS - Automation\\GetVehicleByChromeStyleIdResponseSchema.xls"
            IncludeRebates_col = Get_ColumnNumber(HeaderList, "IncludeRebates")                   
            IncludeRebates_Val = sheet.cell(RowNo,IncludeRebates_col).value

        #Result = Validate_Response_PDGetV_Valid(FileName, RowNo, sResponseCode, sResponseText)
        Result = Validate_Response_PDGetV_Valid_common(FileName, RowNo, sResponseCode, sResponseText,ResponseSchemaFile,IncludeRebates_Val)
        
    else:   # any other response code. invalid response case.
        Result = Validate_Response_InValid(FileName, RowNo, sResponseCode, sResponseText)
        
    #print Result
    return Result
    
#----------------------------------------------------------
# Description:      Function to validate response received for PD_Get Vehicles web service. Response code received is 200.
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row number in excel file
#                   sResponseCode - Response code received after sending request. This is response to be validated.
#                   sResponseText - Response text received after sending request. This is response to be validated.
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Validate_Response_PDGetV_Valid(FileName, RowNo, sResponseCode, sResponseText):
    Result = "PASS"
   
    # Open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)                                  # select required worksheet

    # Calling a function to create list of headers from excel file
    HeaderList = []
    HeaderList = Generate_HeaderList(FileName)
    
    try:
        dom = parseString(sResponseText)        # parsing the xml string
    except xml.parsers.expat.ExpatError, e:     # this will make sure that the xml is valid and there is no tag mismatch
        print e.args[0]
        raise
    except:
        raise ErrorOccured("Invalid XML. Failed to parse XML input string: " + sRequest)          # not able to parse string
    
    #----------------------------------------------------------------------------------------
    # check s:Envelope exists
    
    Result = "PASS"
    col = Get_ColumnNumber(HeaderList, "Node1") # s:Envelope
    Root_val = sheet.cell(RowNo,col).value
    ResultNodeExist = []
    try:
        RootNode = dom.getElementsByTagName(Root_val)[0] # s:Envelope
    except:
        Result = "FAIL"
        print "Expected node as entered in excel column Node1: " + Root_val
        raise ErrorOccured("Required node not found in XML response.")

    print "Required Root/top node found: " +"<"+ Root_val + ">"

    # check s:Body exists
    if Result == "PASS":    # Node1 found
        # Check for node s:Body
        # Calling a function to check if node exists, parent has correct child and get value of node if not empty
        Node2_col = Get_ColumnNumber(HeaderList, "Node2") 
        Node2_val = sheet.cell(RowNo,Node2_col).value
        
        ResultNodeExist = CheckNodeExist(RootNode,0,Root_val,Node2_val)

        # Receive values retured from above function
        Result = ResultNodeExist[0]             # PASS or FAIL
        BodyNode = ResultNodeExist[1]           # s:Body node
               
        if Result == "PASS":        # Node2 found
            # check for GetVehiclesResponse node
            # Calling a function to check if node exists, parent has correct child and get value of node if not empty
            Node3_col = Get_ColumnNumber(HeaderList, "Node3") 
            Node3_val = sheet.cell(RowNo,Node3_col).value
            
            ResultNodeExist = CheckNodeExist(BodyNode,0,Node2_val,Node3_val)
            # Receive values retured from above function
            Result = ResultNodeExist[0]
            GetVehResponseNode = ResultNodeExist[1] # GetVehicleResponse node

            if Result == "PASS":        # Node3 found
                # check for GetVehiclesResult node
                # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                Node4_col = Get_ColumnNumber(HeaderList, "Node4") 
                Node4_val = sheet.cell(RowNo,Node4_col).value
                ResultNodeExist = CheckNodeExist(GetVehResponseNode,0,Node3_val,Node4_val)
                # Receive values retured from above function
                Result = ResultNodeExist[0]
                GetVehResultNode = ResultNodeExist[1]   #GetVehicleResult node
            
                if Result == "PASS":    # Node4 found - GetVehiclesResult
                    # check for Node5 - a:VehicleRebates
                    node5_col = Get_ColumnNumber(HeaderList, "Node5")                   
                    node5_val = sheet.cell(RowNo,node5_col).value
                       
                    ResultNodeExist = CheckNodeExist(GetVehResultNode,0,Node4_val,node5_val)
                    # Receive values retured from above function
                    Result = ResultNodeExist[0]
                    #print "Result: ", Result
                    if Result =="FAIL":
                        raise ErrorOccured("Expected node not found: "+node5_val)
                    
                    VehicleRebatesNode = GetVehResultNode.getElementsByTagName(node5_val) 
                
                    CountNodes = len(VehicleRebatesNode)      # count of "a:VehicleRebates" nodes
                    print "Total number of <a:VehicleRebates> Nodes: ", CountNodes

                    # repeat following for each <a:VehicleRebates> node
                    for i in range(0, CountNodes):
                        print "======================================================================================================================="
                        print "<a:VehicleRebates> node no.: ", i+1
                        print "======================================================================================================================="
                            
                        # Get current VehicleRebates node
                        GetVehNodeNode = GetVehResultNode.childNodes[i]
                        
                        Node6_col = Get_ColumnNumber(HeaderList, "Node6") # a:Rebates
                        Node6_val = sheet.cell(RowNo,Node6_col).value
                            
                        try:
                            RebatesNode = GetVehNodeNode.childNodes[0]     # get a:Rebates
                        except:
                            print "Node not found: " +  Node6_val
                            Result = "FAIL"
                            raise
                        # check that a:Rebates is child of a: vehicles
                        Result = CheckChildNodeExist(Node6_val, RebatesNode, node5_val)

                        if Result == "PASS":    # <a:Rebates> is child of <a:VehicleRebates> node
                            # check value in IncludeRebates column
                            IncludeRebates_col = Get_ColumnNumber(HeaderList, "IncludeRebates")                   
                            IncludeRebates_Val = sheet.cell(RowNo,IncludeRebates_col).value

                            # get child nodes of a:Rebates
                            IncentiveDetailNode = RebatesNode.childNodes
                            CountChildNodes = len(IncentiveDetailNode)
                            #print CountChildNodes
                            
                            if IncludeRebates_Val == "false":
                                if CountChildNodes == 0:
                                    Result = "PASS"
                                    print "You have received correct response."
                                    print "Include Rebates was set to false in input, so <a:Rebates> is expected to be empty"
                                    print "Expected child nodes in <a:Rebates>: 0"
                                    print "Actual no. of child nodes found in <a:Rebates>: ",CountChildNodes
                                else:
                                    Result = "FAIL"
                                    print "You have received incorrect response"
                                    print "Include Rebates was set to false in input, so <a:Rebates> is expected to be empty"
                                    print "Expected child nodes in <a:Rebates>: 0"
                                    print "Actual no. of child nodes found in <a:Rebates>: ",CountChildNodes
                                    raise 
                            else: #<deal1:IncludeRebates> = true
                                # repeat following for each <a:IncentiveDetails> node
                                print "Total number of <a:IncentiveDetails> Nodes: ", CountChildNodes
                                # Repeat following for each <a:IncentiveDetails> node
                                for j in range(0,CountChildNodes):
                                    # Calling a function to check if node exists, parent has correct child and get value of node if not empty

                                    Node7_col = Get_ColumnNumber(HeaderList, "Node7") 
                                    Node7_val = sheet.cell(RowNo,Node7_col).value 
                                    ResultNodeExist = CheckNodeExist(RebatesNode,j,Node6_val,Node7_val)
            
                                    Result = ResultNodeExist[0]
                                    IncentiveDetailNode = ResultNodeExist[1]
                                    if Result == "FAIL":
                                        raise ErrorOccured("Required node is not found: "+ Node7_val)
                
                                    print "****************************************************************"
                                    print "<a:IncentiveDetails> node no.: " ,j+1
                                    print "****************************************************************"
                                    
                                    # check a:Decsription node (Node8)
                                    # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                    Node8_col = Get_ColumnNumber(HeaderList, "Node8") 
                                    Node8_val = sheet.cell(RowNo,Node8_col).value 
                                    ResultNodeExist = CheckNodeExist(IncentiveDetailNode,0,Node7_val,Node8_val)
            
                                    Result = ResultNodeExist[0]
                                    DescNode = ResultNodeExist[1]
                                    if Result == "FAIL":
                                        raise ErrorOccured("Required node is not found: "+ Node8_val)
                                    #-----------------------------------------------------
                                    # check a:IncentiveId node (Node9)
                                    # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                    Node9_col = Get_ColumnNumber(HeaderList, "Node9") 
                                    Node9_val = sheet.cell(RowNo,Node9_col).value 
                                    ResultNodeExist = CheckNodeExist(IncentiveDetailNode,1,Node7_val,Node9_val)
            
                                    Result = ResultNodeExist[0]
                                    IncentiveIdNode = ResultNodeExist[1]
                                    if Result == "FAIL":
                                        raise ErrorOccured("Required node is not found: "+ Node9_val)
                                    #-----------------------------------------------------
                                    # check a:ProgramName node (Node10)
                                    # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                    Node10_col = Get_ColumnNumber(HeaderList, "Node10") 
                                    Node10_val = sheet.cell(RowNo,Node10_col).value 
                                    ResultNodeExist = CheckNodeExist(IncentiveDetailNode,2,Node7_val,Node10_val)
            
                                    Result = ResultNodeExist[0]
                                    ProgramNameNode = ResultNodeExist[1]
                                    if Result == "FAIL":
                                        raise ErrorOccured("Required node is not found: "+ Node10_val)
                                    #-----------------------------------------------------
                                    # check a:Cash node (Node11)
                                    # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                    Node11_col = Get_ColumnNumber(HeaderList, "Node11") 
                                    Node11_val = sheet.cell(RowNo,Node11_col).value 
                                    ResultNodeExist = CheckNodeExist(IncentiveDetailNode,3,Node7_val,Node11_val)
            
                                    Result = ResultNodeExist[0]
                                    CashNode = ResultNodeExist[1]
                                    if Result == "FAIL":
                                        raise ErrorOccured("Required node is not found: "+ Node11_val)
                                    #-----------------------------------------------------
                                    # check a:ChromeStyleIds node (Node12)
                                    # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                    Node12_col = Get_ColumnNumber(HeaderList, "Node12") 
                                    Node12_val = sheet.cell(RowNo,Node12_col).value 
                                    ResultNodeExist = CheckNodeExist(IncentiveDetailNode,4,Node7_val,Node12_val)
                                    
                                    Result = ResultNodeExist[0]
                                    ChromeStyleIdsNode = ResultNodeExist[1]
                                    if Result == "FAIL":
                                        raise ErrorOccured("Required node is not found: "+ Node12_val)
                                    #print "*****"
                                    #-----------------------------------------------------
                                    # check a:ConditionCategory node (Node13)
                                    # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                    Node13_col = Get_ColumnNumber(HeaderList, "Node13") 
                                    Node13_val = sheet.cell(RowNo,Node13_col).value 
                                    ResultNodeExist = CheckNodeExist(IncentiveDetailNode,5,Node7_val,Node13_val)
            
                                    Result = ResultNodeExist[0]
                                    ConditionCategoryNode = ResultNodeExist[1]
                                    if Result == "FAIL":
                                        raise ErrorOccured("Required node is not found: "+ Node13_val)
                                    #-----------------------------------------------------
                                    # check a:EndDate node (Node14)
                                    # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                    Node14_col = Get_ColumnNumber(HeaderList, "Node14") 
                                    Node14_val = sheet.cell(RowNo,Node14_col).value 
                                    ResultNodeExist = CheckNodeExist(IncentiveDetailNode,6,Node7_val,Node14_val)
            
                                    Result = ResultNodeExist[0]
                                    EndDateNode = ResultNodeExist[1]
                                    if Result == "FAIL":
                                        raise ErrorOccured("Required node is not found: "+ Node14_val)
                                    #-----------------------------------------------------
                                    # check a:Exclusions node (Node15)
                                    # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                    Node15_col = Get_ColumnNumber(HeaderList, "Node15") 
                                    Node15_val = sheet.cell(RowNo,Node15_col).value 
                                    ResultNodeExist = CheckNodeExist(IncentiveDetailNode,7,Node7_val,Node15_val)
            
                                    Result = ResultNodeExist[0]
                                    ExclusionsNode = ResultNodeExist[1]
                                    
                                    if Result == "PASS":    #   node15 found - a:Exclusions
                                        b_intNode = ExclusionsNode.childNodes
                                        countIntNodes = len(b_intNode)
                                        if countIntNodes == 0:      # count can be 0.
                                            print "<a:Exclusions> node does not have any child <b:int> nodes"
                                        else:
                                            col = Get_ColumnNumber(HeaderList, "Node16") # b:int
                                            Node16_val = sheet.cell(RowNo,col).value
                                            Result = CheckChildNodeExist(Node16_val, b_intNode, Node15_val)
                                            if Result == "PASS":
                                                print "Total number of <b:int> nodes: ",countIntNodes
                                                for k in range (0,countIntNodes):
                                                    b_intNode = ExclusionsNode.childNodes[k]
                                                    GetNodeValue(b_intNode,"b:int")
                                            else:
                                                raise ErrorOccured("Required child node not found: "+Node16_val)
                                    else:    
                                        raise ErrorOccured("Required node not found in XML response: " + Node15_val)
                                        
                                    #-----------------------------------------------------
                                    # check a:FinanceMethod node (Node17)
                                    # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                    Node17_col = Get_ColumnNumber(HeaderList, "Node17") 
                                    Node17_val = sheet.cell(RowNo,Node17_col).value 
                                    ResultNodeExist = CheckNodeExist(IncentiveDetailNode,8,Node7_val,Node17_val)
            
                                    Result = ResultNodeExist[0]
                                    FinanceMethodNode = ResultNodeExist[1]
                                    if Result == "FAIL":
                                        raise ErrorOccured("Required node is not found: "+ Node17_val)
                                    #-----------------------------------------------------
                                    # check a:FundingSourceName node (Node18)
                                    # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                    Node18_col = Get_ColumnNumber(HeaderList, "Node18") 
                                    Node18_val = sheet.cell(RowNo,Node18_col).value 
                                    ResultNodeExist = CheckNodeExist(IncentiveDetailNode,9,Node7_val,Node18_val)
            
                                    Result = ResultNodeExist[0]
                                    FundingSourceNameNode = ResultNodeExist[1]
                                    if Result == "FAIL":
                                        raise ErrorOccured("Required node is not found: "+ Node18_val)
                                    #-----------------------------------------------------
                                    # check a:IncentiveType node (Node19)
                                    # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                    Node19_col = Get_ColumnNumber(HeaderList, "Node19") 
                                    Node19_val = sheet.cell(RowNo,Node19_col).value 
                                    ResultNodeExist = CheckNodeExist(IncentiveDetailNode,10,Node7_val,Node19_val)
            
                                    Result = ResultNodeExist[0]
                                    IncentiveTypeNode = ResultNodeExist[1]
                                    if Result == "FAIL":
                                        raise ErrorOccured("Required node is not found: "+ Node19_val)
                                    #-----------------------------------------------------
                                    # check a:IncentiveVehicles node (Node20)
                                    # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                    Node20_col = Get_ColumnNumber(HeaderList, "Node20") 
                                    Node20_val = sheet.cell(RowNo,Node20_col).value 
                                    ResultNodeExist = CheckNodeExist(IncentiveDetailNode,11,Node7_val,Node20_val)
            
                                    Result = ResultNodeExist[0]
                                    IncentiveVehiclesNode = ResultNodeExist[1]
                                    if Result == "FAIL":
                                        raise ErrorOccured("Required node is not found: "+ Node20_val)
                                    #-----------------------------------------------------
                                    # check a:IsStandardRebate node (Node21)
                                    # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                    Node21_col = Get_ColumnNumber(HeaderList, "Node21") 
                                    Node21_val = sheet.cell(RowNo,Node21_col).value 
                                    ResultNodeExist = CheckNodeExist(IncentiveDetailNode,12,Node7_val,Node21_val)
            
                                    Result = ResultNodeExist[0]
                                    IsStandardRebateNode = ResultNodeExist[1]
                                    if Result == "FAIL":
                                        raise ErrorOccured("Required node is not found: "+ Node21_val)
                                    #-----------------------------------------------------
                                    # check a:ProgramId node (Node22)
                                    # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                    Node22_col = Get_ColumnNumber(HeaderList, "Node22") 
                                    Node22_val = sheet.cell(RowNo,Node22_col).value 
                                    ResultNodeExist = CheckNodeExist(IncentiveDetailNode,13,Node7_val,Node22_val)
            
                                    Result = ResultNodeExist[0]
                                    ProgramIdNode = ResultNodeExist[1]
                                    if Result == "FAIL":
                                        raise ErrorOccured("Required node is not found: "+ Node22_val)
                                    #-----------------------------------------------------
                                    # check a:StartDate node (Node23)
                                    # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                    Node23_col = Get_ColumnNumber(HeaderList, "Node23") 
                                    Node23_val = sheet.cell(RowNo,Node23_col).value 
                                    ResultNodeExist = CheckNodeExist(IncentiveDetailNode,14,Node7_val,Node23_val)
            
                                    Result = ResultNodeExist[0]
                                    StartDateNode = ResultNodeExist[1]
                                    if Result == "FAIL":
                                        raise ErrorOccured("Required node is not found: "+ Node23_val)
                                    #-----------------------------------------------------
                                    # check a:TermRebates node (Node24)
                                    # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                    Node24_col = Get_ColumnNumber(HeaderList, "Node24") 
                                    Node24_val = sheet.cell(RowNo,Node24_col).value 
                                    ResultNodeExist = CheckNodeExist(IncentiveDetailNode,15,Node7_val,Node24_val)
            
                                    Result = ResultNodeExist[0]
                                    TermRebatesNode = ResultNodeExist[1]

                                    if Result == "PASS":    #   node24 found - a:TermRebates
                                        # get value if not empty
                                        #GetNodeValue(TermRebatesNode,"Term Rebates")
                                        IncentiveTermRangeNode = TermRebatesNode.childNodes
                                        CountIncTermRange = len(IncentiveTermRangeNode)
                                        if CountIncTermRange == 0:
                                            print "There are no child items for <a:TermRebates>"
                                        else:
                                            print "Total number of <a:IncentiveTermRange> nodes: ", CountIncTermRange
                                            for l in range (0,CountIncTermRange):
                                                # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                                Node25_col = Get_ColumnNumber(HeaderList, "Node25") 
                                                Node25_val = sheet.cell(RowNo,Node25_col).value 
                                                ResultNodeExist = CheckNodeExist(TermRebatesNode,l,Node7_val,Node25_val)
                                    
                                                Result = ResultNodeExist[0]
                                                IncentiveTermRangeNode = ResultNodeExist[1]
                                                if Result == "FAIL":
                                                    raise ErrorOccured("Required node is not found: "+ Node25_val)
                                                
                                                print "+++++++++++++++++++++++++++++++++++++"
                                                print "<a:IncentiveTermRange> node no.: " ,l+1
                                                print "+++++++++++++++++++++++++++++++++++++"
                                                if Result == "PASS":
                                                    # check for a:Apr
                                                    Node26_col = Get_ColumnNumber(HeaderList, "Node26") 
                                                    Node26_val = sheet.cell(RowNo,Node26_col).value 
                                                    ResultNodeExist = CheckNodeExist(IncentiveTermRangeNode,0,Node25_val,Node26_val)
            
                                                    Result = ResultNodeExist[0]
                                                    a_AprNode = ResultNodeExist[1]
                                                    if Result == "FAIL":
                                                        raise ErrorOccured("Required node is not found: "+ Node26_val)
                                                    #-----------------------------------------------
                                                    # check for a:CustomerCash
                                                    # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                                    Node27_col = Get_ColumnNumber(HeaderList, "Node27") 
                                                    Node27_val = sheet.cell(RowNo,Node27_col).value 
                                                    ResultNodeExist = CheckNodeExist(IncentiveTermRangeNode,1,Node25_val,Node27_val)
            
                                                    Result = ResultNodeExist[0]
                                                    a_CustomerCashNode = ResultNodeExist[1]
                                                    if Result == "FAIL":
                                                        raise ErrorOccured("Required node is not found: "+ Node27_val)
                                                    #-----------------------------------------------
                                                    # check for a:EndTerm
                                                    Node28_col = Get_ColumnNumber(HeaderList, "Node28") 
                                                    Node28_val = sheet.cell(RowNo,Node28_col).value 
                                                    ResultNodeExist = CheckNodeExist(IncentiveTermRangeNode,2,Node25_val,Node28_val)
            
                                                    Result = ResultNodeExist[0]
                                                    a_EndTermNode = ResultNodeExist[1]
                                                    if Result == "FAIL":
                                                        raise ErrorOccured("Required node is not found: "+ Node28_val)
                                                    #-----------------------------------------------
                                                    # check for a:MoneyFactor
                                                    # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                                    Node29_col = Get_ColumnNumber(HeaderList, "Node29") 
                                                    Node29_val = sheet.cell(RowNo,Node29_col).value 
                                                    ResultNodeExist = CheckNodeExist(IncentiveTermRangeNode,3,Node25_val,Node29_val)
                                                    
                                                    Result = ResultNodeExist[0]
                                                    a_MoneyFactorNode = ResultNodeExist[1]
                                                    if Result == "FAIL":
                                                        raise ErrorOccured("Required node is not found: "+ Node29_val)
                                                    #-----------------------------------------------
                                                    # check for a:StartTerm
                                                    # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                                    Node30_col = Get_ColumnNumber(HeaderList, "Node30") 
                                                    Node30_val = sheet.cell(RowNo,Node30_col).value 
                                                    ResultNodeExist = CheckNodeExist(IncentiveTermRangeNode,4,Node25_val,Node30_val)

                                                    Result = ResultNodeExist[0]
                                                    a_StartTermNode = ResultNodeExist[1]
                                                    if Result == "FAIL":
                                                        raise ErrorOccured("Required node is not found: "+ Node30_val)
                                                      
                                                else:
                                                    raise ErrorOccured("Required node not found in XML response: " + Node25_val)
                                    else:
                                        raise ErrorOccured("Required node not found in XML response: " + Node24_val)
                        else: #a:Rebates not found
                            raise ErrorOccured("Required node not found in XML response: " + Node6_val)

                        #Going to check for a:Vehicle node
                        print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
                        print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"                        
                        
                        # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                        Node31_col = Get_ColumnNumber(HeaderList, "Node31") 
                        Node31_val = sheet.cell(RowNo,Node31_col).value 
                        ResultNodeExist = CheckNodeExist(GetVehNodeNode,1,node5_val,Node31_val)
            
                        Result = ResultNodeExist[0]
                        VehicleNode = ResultNodeExist[1]
                        
                        if Result == "PASS":    # a:Vehicle Node found
                            #check for a:AddedOptions node
                            # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                            Node32_col = Get_ColumnNumber(HeaderList, "Node32") 
                            Node32_val = sheet.cell(RowNo,Node32_col).value 
                            ResultNodeExist = CheckNodeExist(VehicleNode,0,Node31_val,Node32_val)
            
                            Result = ResultNodeExist[0]
                            a_AddedOptionsNode = ResultNodeExist[1]
                            
                            if Result == "PASS":    #   node32 found - a:AddedOptions
                                # check for child nodes
                                # b:Cost
                                # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                Node33_col = Get_ColumnNumber(HeaderList, "Node33") 
                                Node33_val = sheet.cell(RowNo,Node33_col).value 
                                ResultNodeExist = CheckNodeExist(a_AddedOptionsNode,0,Node32_val,Node33_val)
                            
                                Result = ResultNodeExist[0]
                                b_CostNode = ResultNodeExist[1]
                                if Result == "FAIL":
                                    raise ErrorOccured("Required node is not found: "+ Node33_val)
                                #--------------------------------
                                # b:Retail
                                # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                Node34_col = Get_ColumnNumber(HeaderList, "Node34") 
                                Node34_val = sheet.cell(RowNo,Node34_col).value 
                                ResultNodeExist = CheckNodeExist(a_AddedOptionsNode,1,Node32_val,Node34_val)
            
                                Result = ResultNodeExist[0]
                                b_RetailNode = ResultNodeExist[1]
                                if Result == "FAIL":
                                    raise ErrorOccured("Required node is not found: "+ Node34_val)
                            else:
                                raise ErrorOccured("Required node not found in XML response: " + Node32_val)
                            #--------------------------------------------------
                            #check for a:Msrp node
                            # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                            Node35_col = Get_ColumnNumber(HeaderList, "Node35") 
                            Node35_val = sheet.cell(RowNo,Node35_col).value 
                            ResultNodeExist = CheckNodeExist(VehicleNode,1,Node31_val,Node35_val) 
            
                            Result = ResultNodeExist[0]
                            a_MsrpNode = ResultNodeExist[1]
                            
                            if Result == "PASS":    #   node35 found - a:MSRP
                                # check for child nodes
                                # b:Cost
                                # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                Node36_col = Get_ColumnNumber(HeaderList, "Node36") 
                                Node36_val = sheet.cell(RowNo,Node36_col).value 
                                ResultNodeExist = CheckNodeExist(a_MsrpNode,0,Node35_val,Node36_val)
            
                                Result = ResultNodeExist[0]
                                b_CostNode = ResultNodeExist[1]
                                if Result == "FAIL":
                                    raise ErrorOccured("Required node is not found: "+ Node36_val)
                                #--------------------------------
                                # b:Retail
                                # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                Node37_col = Get_ColumnNumber(HeaderList, "Node37") 
                                Node37_val = sheet.cell(RowNo,Node37_col).value 
                                ResultNodeExist = CheckNodeExist(a_MsrpNode,1,Node35_val,Node37_val)
            
                                Result = ResultNodeExist[0]
                                b_RetailNode = ResultNodeExist[1]
                                if Result == "FAIL":
                                    raise ErrorOccured("Required node is not found: "+ Node37_val)
                            else:
                                raise ErrorOccured("Required node not found in XML response: " + Node35_val) #msrp
                                
                            #--------------------------------------------------
                            #check for a:Odometer node
                            # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                            Node38_col = Get_ColumnNumber(HeaderList, "Node38") 
                            Node38_val = sheet.cell(RowNo,Node38_col).value 
                            ResultNodeExist = CheckNodeExist(VehicleNode,2,Node31_val,Node38_val)
                                
                            Result = ResultNodeExist[0]
                            a_OdometerNode = ResultNodeExist[1]
                            if Result == "FAIL":
                                    raise ErrorOccured("Required node is not found: "+ Node38_val)
                            #--------------------------------------------------
                            #check for a:Options node
                            # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                            Node39_col = Get_ColumnNumber(HeaderList, "Node39") 
                            Node39_val = sheet.cell(RowNo,Node39_col).value 
                            ResultNodeExist = CheckNodeExist(VehicleNode,3,Node31_val,Node39_val)
            
                            Result = ResultNodeExist[0]
                            a_OptionsNode = ResultNodeExist[1]
                            if Result == "FAIL":
                                    raise ErrorOccured("Required node is not found: "+ Node39_val)
                            #--------------------------------------------------
                            #check for a:RemovedOptions node
                            # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                            Node40_col = Get_ColumnNumber(HeaderList, "Node40") 
                            Node40_val = sheet.cell(RowNo,Node40_col).value 
                            ResultNodeExist = CheckNodeExist(VehicleNode,4,Node31_val,Node40_val)
            
                            Result = ResultNodeExist[0]
                            a_RemovedOptionsNode = ResultNodeExist[1]
                            
                            if Result == "PASS":    #   node35 found - a:RemovedOptions
                                # check for child nodes
                                # b:Cost
                                # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                Node41_col = Get_ColumnNumber(HeaderList, "Node41") 
                                Node41_val = sheet.cell(RowNo,Node41_col).value 
                                ResultNodeExist = CheckNodeExist(a_RemovedOptionsNode,0,Node40_val,Node41_val)
            
                                Result = ResultNodeExist[0]
                                b_CostNode = ResultNodeExist[1]
                                if Result == "FAIL":
                                    raise ErrorOccured("Required node is not found: "+ Node41_val)
                                #--------------------------------
                                # b:Retail
                                # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                Node42_col = Get_ColumnNumber(HeaderList, "Node42") 
                                Node42_val = sheet.cell(RowNo,Node42_col).value 
                                ResultNodeExist = CheckNodeExist(a_RemovedOptionsNode,1,Node40_val,Node42_val)
            
                                Result = ResultNodeExist[0]
                                b_RetailNode = ResultNodeExist[1]
                                if Result == "FAIL":
                                    raise ErrorOccured("Required node is not found: "+ Node42_val)
                            else:
                                raise ErrorOccured("Required node not found in XML response: " + Node40_val)
                            #--------------------------------------------------
                            #check for a:VehicleCode node
                            # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                            Node43_col = Get_ColumnNumber(HeaderList, "Node43") 
                            Node43_val = sheet.cell(RowNo,Node43_col).value 
                            ResultNodeExist = CheckNodeExist(VehicleNode,5,Node31_val,Node43_val)
                                
                            Result = ResultNodeExist[0]
                            a_VehicleCodeNode = ResultNodeExist[1]
                            
                            if Result == "PASS":    #   node43 found - a:VehicleCode
                                # check for child nodes
                                # b:Condition
                                # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                Node44_col = Get_ColumnNumber(HeaderList, "Node44") 
                                Node44_val = sheet.cell(RowNo,Node44_col).value 
                                ResultNodeExist = CheckNodeExist(a_VehicleCodeNode,0,Node43_val,Node44_val)
            
                                Result = ResultNodeExist[0]
                                b_ConditionNode = ResultNodeExist[1]
                                if Result == "FAIL":
                                    raise ErrorOccured("Required node is not found: "+ Node44_val)
                                #--------------------------------
                                # b:MakeCode
                                # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                Node45_col = Get_ColumnNumber(HeaderList, "Node45") 
                                Node45_val = sheet.cell(RowNo,Node45_col).value 
                                ResultNodeExist = CheckNodeExist(a_VehicleCodeNode,1,Node43_val,Node45_val)
            
                                Result = ResultNodeExist[0]
                                b_MakeCodeNode = ResultNodeExist[1]
                                if Result == "FAIL":
                                    raise ErrorOccured("Required node is not found: "+ Node45_val)
                                #--------------------------------
                                # b:ModelCode
                                # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                Node46_col = Get_ColumnNumber(HeaderList, "Node46") 
                                Node46_val = sheet.cell(RowNo,Node46_col).value 
                                ResultNodeExist = CheckNodeExist(a_VehicleCodeNode,2,Node43_val,Node46_val)
            
                                Result = ResultNodeExist[0]
                                b_ModelCodeNode = ResultNodeExist[1]
                                if Result == "FAIL":
                                    raise ErrorOccured("Required node is not found: "+ Node46_val)
                                #--------------------------------
                                # b:StyleCode
                                # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                Node47_col = Get_ColumnNumber(HeaderList, "Node47") 
                                Node47_val = sheet.cell(RowNo,Node47_col).value 
                                ResultNodeExist = CheckNodeExist(a_VehicleCodeNode,3,Node43_val,Node47_val)
            
                                Result = ResultNodeExist[0]
                                b_StyleCodeNode = ResultNodeExist[1]
                                if Result == "FAIL":
                                    raise ErrorOccured("Required node is not found: "+ Node47_val)
                                #--------------------------------
                                # b:Year
                                # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                Node48_col = Get_ColumnNumber(HeaderList, "Node48") 
                                Node48_val = sheet.cell(RowNo,Node48_col).value 
                                ResultNodeExist = CheckNodeExist(a_VehicleCodeNode,4,Node43_val,Node48_val)
            
                                Result = ResultNodeExist[0]
                                b_YearNode = ResultNodeExist[1]
                                if Result == "FAIL":
                                    raise ErrorOccured("Required node is not found: "+ Node48_val)
                            else:
                                raise ErrorOccured("Required node not found in XML response: " + Node43_val)

                            #--------------------------------------------------
                            #check for a:VehicleDescription node
                            # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                            Node49_col = Get_ColumnNumber(HeaderList, "Node49") 
                            Node49_val = sheet.cell(RowNo,Node49_col).value 
                            ResultNodeExist = CheckNodeExist(VehicleNode,6,Node31_val,Node49_val)
            
                            Result = ResultNodeExist[0]
                            a_VehicleDescNode = ResultNodeExist[1]
                                
                            if Result == "PASS":    #   node49 found - a:VehicleDescription
                                # check for child nodes
                                # b:Condition
                                # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                Node50_col = Get_ColumnNumber(HeaderList, "Node50") 
                                Node50_val = sheet.cell(RowNo,Node50_col).value 
                                ResultNodeExist = CheckNodeExist(a_VehicleDescNode,0,Node49_val,Node50_val)
            
                                Result = ResultNodeExist[0]
                                b_ConditionNode = ResultNodeExist[1]
                                if Result == "FAIL":
                                    raise ErrorOccured("Required node is not found: "+ Node50_val)
                                #--------------------------------
                                # b:Make
                                # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                Node51_col = Get_ColumnNumber(HeaderList, "Node51") 
                                Node51_val = sheet.cell(RowNo,Node51_col).value 
                                ResultNodeExist = CheckNodeExist(a_VehicleDescNode,1,Node49_val,Node51_val)
            
                                Result = ResultNodeExist[0]
                                b_MakeNode = ResultNodeExist[1]
                                if Result == "FAIL":
                                    raise ErrorOccured("Required node is not found: "+ Node51_val)
                                #--------------------------------
                                # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                # b:Model
                                Node52_col = Get_ColumnNumber(HeaderList, "Node52") 
                                Node52_val = sheet.cell(RowNo,Node52_col).value 
                                ResultNodeExist = CheckNodeExist(a_VehicleDescNode,2,Node49_val,Node52_val)
            
                                Result = ResultNodeExist[0]
                                b_ModelNode = ResultNodeExist[1]
                                if Result == "FAIL":
                                    raise ErrorOccured("Required node is not found: "+ Node52_val)
                                #--------------------------------
                                # b:Style
                                # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                Node53_col = Get_ColumnNumber(HeaderList, "Node53") 
                                Node53_val = sheet.cell(RowNo,Node53_col).value 
                                ResultNodeExist = CheckNodeExist(a_VehicleDescNode,3,Node49_val,Node53_val)
            
                                Result = ResultNodeExist[0]
                                b_StyleNode = ResultNodeExist[1]
                                if Result == "FAIL":
                                    raise ErrorOccured("Required node is not found: "+ Node53_val)
                                #--------------------------------
                                # Calling a function to check if node exists, parent has correct child and get value of node if not empty
                                Node54_col = Get_ColumnNumber(HeaderList, "Node54") 
                                Node54_val = sheet.cell(RowNo,Node54_col).value 
                                ResultNodeExist = CheckNodeExist(a_VehicleDescNode,4,Node49_val,Node54_val)
            
                                Result = ResultNodeExist[0]
                                b_YearNode = ResultNodeExist[1]
                                if Result == "FAIL":
                                    raise ErrorOccured("Required node is not found: "+ Node54_val)
                            else:
                                raise ErrorOccured("Required node not found in XML response: "+Node49_val)
                else:       # Node4 not found
                    raise ErrorOccured("Required node not found in XML response: "+ Node4_val)
            else:       # Node3 not found
                raise ErrorOccured("Required node not found in XML response: "+ Node3_val)
        else:       # Node2 not found
            raise ErrorOccured("Required node not found in XML response: "+ Node2_val)
    else:       # Node1 (Root) not found
        print "Required value in excel column Node1 is: " + Root_val
        raise ErrorOccured("Expected value as entered in excel column Node1 is: " + Root_val)
    return Result

#----------------------------------------------------------
# Description:      Function to validate response received for PD_Get Vehicles web service. Response code received is 200.
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row number in excel file
#                   Node - Parent node
#                   NodeName - child node name
# Author:           Manisha Gadekar
#----------------------------------------------------------
def CheckNodeExist(ParentNode,ChildNodeNo,ParentNodeName,NodeDescription):
    Result = "PASS"
    
    Node_val = str(NodeDescription)       # get excel value
    
    try:
        Node = ParentNode.childNodes[ChildNodeNo]   # get child node of required occurance
    except:
        Result = "FAIL"

    if Result == "PASS":
        # Calling a function to check that the child item created is required child item
        Result = CheckChildNodeExist(Node_val, Node, ParentNodeName)
        if Result == "PASS":    
            # get value if not empty
            GetNodeValue(Node,Node_val)
        else:
            #print "Expected node as entered in excel column " + ExcelCol + " is: " + Node_val
            Result = "FAIL"
    
    return Result, Node

#----------------------------------------------------------
# Description:      Function to check if Parent node has required child node
# Input Parameters: ChildNodeName - Name of child node from excel
#                   ChildNode - Child node
#                   ParentNodeName - Parent node name
# Author:           Manisha Gadekar
#----------------------------------------------------------
def CheckChildNodeExist(ChildNodeName, ChildNode, ParentNodeName):
    Result = "PASS"
    # Child node will be like - <DOM Element: s:Envelope at 0x2453a48>
    # so checking if child node name from excel is part of "Child node"
    nameofNode = str(ChildNode)
    
    if ChildNodeName in nameofNode and ChildNodeName!= "" and ChildNodeName!= " ":
        print "<" + ParentNodeName + "> has child node: <" + ChildNodeName + ">"
        Result = "PASS"
    else:
        Result = "FAIL"
        #raise ErrorOccured("Child node not found: " + ChildNodeName)
    return Result

#----------------------------------------------------------
# Description:      Function to get value of xml node 
# Input Parameters: Node - Node whose value is required
#                   Nodename - Name of the node
# Author:           Manisha Gadekar
#----------------------------------------------------------   
def GetNodeValue(Node, Nodename):
    # check if text node is not empty
    if Node.firstChild != None:
        Req_Val = Node.firstChild.nodeValue # get value of text
        if Req_Val!= None:
            print "Value of <"+ Nodename + "> is: " + str(Req_Val)
            print "*****"
    else:
        print "Value of <"+ Nodename + ">: " + Nodename +" has no text."
        print "*****"

#----------------------------------------------------------
# Description:      Function to validate response received. Error Message is expected. 
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row number in excel file
#                   sResponseCode - Response code
#                   sResponseText - Response text
# Author:           Manisha Gadekar
#----------------------------------------------------------         
def Validate_Response_InValid(FileName, RowNo, sResponseCode, sResponseText):
    Result = "PASS"
    # Open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)
    
    # Calling a function to create list of headers from excel file
    HeaderList = []
    HeaderList = Generate_HeaderList(FileName)
    
    col = Get_ColumnNumber(HeaderList, "ExpectedErrorResponse") # ExpectedErrorReponse column
    ExpectedError = sheet.cell(RowNo,col).value

    pos = sResponseText.find(ExpectedError) # check if expected error is present in actual response

    if pos == -1:           # string not found
        Result = "FAIL"
        print "You have not received correct error message."
        print "Expected error: " + ExpectedError
        print "****************************************"
        print "Actual received error: " + sResponseText
        raise ErrorOccured("You have not received correct error message.\n\n"+"Expected error message: " +ExpectedError+ "\n\nActual error message: " +sResponseText )
    else:                   # string found
        print "You have received correct error message."
        print "Expected error: " + ExpectedError
        print "****************************************"
        print "Actual received error: " + sResponseText
   
    return Result

#----------------------------------------------------------
# Description:      Function to validate response received for PD_Get Vehicles web service. Response code received is 200.
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row number in excel file
#                   sResponseCode - Response code received after sending request. This is response to be validated.
#                   sResponseText - Response text received after sending request. This is response to be validated.
#                   ResponseSchemaFile - path of excel file where expected xml schema of response is stored
#                   IncludeRebates_Val - value of IncludeRebates - true or false
# Author:           Manisha Gadekar
#----------------------------------------------------------            
def Validate_Response_PDGetV_Valid_common(FileName, RowNo, sResponseCode, sResponseText, ResponseSchemaFile,IncludeRebates_Val):
    Result = "PASS"
    book = Open_Excel(ResponseSchemaFile)
    noRows = Get_NumberOfRowsColumns(ResponseSchemaFile,"ROWS")
    print "total rows: ",noRows
    
    try:
        dom = parseString(sResponseText)        # parsing the xml string
    except xml.parsers.expat.ExpatError, e:     # this will make sure that the xml is valid and there is no tag mismatch
        print e.args[0]
        raise
    except:
        raise ErrorOccured("Invalid XML. Failed to parse XML input string: " + sRequest)          # not able to parse string
    book = Open_Excel(ResponseSchemaFile)
    sheet = book.sheet_by_index(0)
    
    Root_Val = sheet.cell(1,1).value
    try:
        RootNode = dom.getElementsByTagName(Root_Val)
        #RootNode = dom.getElementsByTagName("s:Envelope")
    except:
        raise ErrorOccured("Root Node not found")
  
    for i in range(1, noRows):          # run on all excel rows
        print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
        print "Row no: ",i
        nodeval = sheet.cell(i,1).value     # read value of node from excel
        nodeval=str(nodeval.strip())
        try:
            nodeList = dom.getElementsByTagName(nodeval) # node can appear mutiple times in xml
        except:
            raise ErrorOccured ("Node not found."+ nodeval)
        print "Node found: " + nodeval
        print nodeList
        noofnodes = len(nodeList)   # get total number of nodes
        if noofnodes == 0:      #node not found
            if IncludeRebates_Val.strip() in "false":
                if nodeval == "a:IncentiveDetail" or nodeval == "a:Exclusions" or nodeval =="a:TermRebates" or nodeval == "a:IncentiveTermRange":
                    print "You have received correct response."
                    print "Include rebates was set to false, so expected no. of children for <a:Rebates> is 0"
                    print "<" + nodeval + "> should not be present"
                else:
                    Result = "FAIL"
                    raise ErrorOccured("Expected node is not found: "+ nodeval)
            else:
                Result = "FAIL"
                raise ErrorOccured("Expected node is not found: "+ nodeval)
        
        print "No of nodes: ", noofnodes
        no_childs = int(sheet.cell(i,2).value)  # read no. of children from excel
        req_childlist = []
        for j in range(3, 3+no_childs):
            req_childlist.append(str(sheet.cell(i,j).value))    # create list of req child nodes
        # loop to run on all instances of a node
        for j in range(0,noofnodes):
            print "--------------------------------------------------------------------------------------"
            print "Node no.: " ,j+1
            node = nodeList[j]      # get node from node list
            print "Node: ",node
            
            childList = node.childNodes     # get list of children of node from xml
            
            print "Expected child list: ",req_childlist
            print "Actual Child List: ", childList
            
            cnt = 0     # count for storing no of times child node appear in xml
            for k in range(0,len(req_childlist)):
                req_val = req_childlist[k]  # get expected value from list created above
                print "*******************"
                print "Looking for child: ",req_val
                if IncludeRebates_Val.strip() in "false":   # if include rebates is false, <a:IncentiveDetail> should not be present
                    if req_val=="a:IncentiveDetail":
                        if len(childList)==0:
                            print "You have received correct response."
                            print "Include rebates was set to false, so expected no. of children for <a:Rebates> is 0"
                            print "Actual no of childeren for <a:Rebates>: ",len(childList)
                        else:
                            print "Include rebates was set to false, so expected no. of children for <a:Rebates> is 0"
                            print "Actual no of childeren for <a:Rebates>: ",len(childList)
                            Result = "FAIL"
                            raise ErrorOccured("You have not received correct response.")
                flag = 0
                for l in range(0,len(childList)):
                    act_val = childList[l]
                    if req_val in str(act_val) and req_val!="" and req_val!=" ":
                        flag = 1
                        cnt = cnt + 1
                        print "Actual child: ", str(act_val)
                        print "*** PASS - Required child node found ***"
                        if act_val.firstChild != None and act_val.firstChild.nodeValue!= None:
                            print "Node Value: ", act_val.firstChild.nodeValue
                        else:
                            print "Node has no text."
                if flag == 0:   # req child node not found
                    if req_val=="b:int" or req_val== "a:IncentiveTermRange" or req_val== "a:IncentiveDetail":
                        Result = "PASS"
                    else:
                        Result = "FAIL"
                        print "Expected child node: ", req_val
                        raise ErrorOccured("Expected child node not found: " +req_val)
            print "Child node found count: ",cnt
            if cnt == 0:    # no child found
                if req_val=="b:int" or req_val== "a:IncentiveTermRange" or req_val== "a:IncentiveDetail":
                    print "This is fine. No child node found: ",req_val
                else:
                    Result = "FAIL"
                    raise ErrorOccured("Expected child node not found: " +req_val)
    return(Result)








#working copy
def Validate_Response_PDGetV_Valid_new1(FileName, RowNo, sResponseCode, sResponseText, ResponseSchemaFile,IncludeRebates_Val):
    Result = "PASS"
    book = Open_Excel(ResponseSchemaFile)
    noRows = Get_NumberOfRowsColumns(ResponseSchemaFile,"ROWS")
    print "total rows: ",noRows
    
    try:
        dom = parseString(sResponseText)        # parsing the xml string
    except xml.parsers.expat.ExpatError, e:     # this will make sure that the xml is valid and there is no tag mismatch
        print e.args[0]
        raise
    except:
        raise ErrorOccured("Invalid XML. Failed to parse XML input string: " + sRequest)          # not able to parse string
    book = Open_Excel(ResponseSchemaFile)
    sheet = book.sheet_by_index(0)
    
    Root_Val = sheet.cell(1,1).value
    try:
        RootNode = dom.getElementsByTagName(Root_Val)
        #RootNode = dom.getElementsByTagName("s:Envelope")
    except:
        raise ErrorOccured("Root Node not found")
  
    for i in range(1, noRows):
        print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
        print "Row no: ",i
        
        nodeval = sheet.cell(i,1).value
        nodeval=str(nodeval.strip())
        #print nodeval
        try:
            nodeList = dom.getElementsByTagName(nodeval)
        except:
            raise ErrorOccured ("Node not found."+ nodeval)
        print "Node found: " + nodeval
        print nodeList
        noofnodes = len(nodeList)
        
        print "No of nodes: ", noofnodes
        no_childs = int(sheet.cell(i,2).value)
        req_childlist = []
        for j in range(3, 3+no_childs):
            req_childlist.append(str(sheet.cell(i,j).value))

        for j in range(0,noofnodes):
            print "--------------------------------------------------------------------------------------"
            print "Node no.: " ,j+1
            node = nodeList[j]
            print "Node: ",node
            childList = node.childNodes
            
            print "Expected child list: ",req_childlist
            print "Actual Child List: ", childList
            cnt = 0
            for k in range(0,len(req_childlist)):
                req_val = req_childlist[k]
                print "*******************"
                print "Looking for child: ",req_val
                for l in range(0,len(childList)):
                    act_val = childList[l]
                    if req_val in str(act_val) and req_val!="" and req_val!=" ":
                        cnt = cnt + 1
                        print "Actual child: ", str(act_val)
                        print "*** PASS - Required child node found ***"
                        if act_val.firstChild != None and act_val.firstChild.nodeValue!= None:
                            print "Node Value: ", act_val.firstChild.nodeValue
                        else:
                            print "Node has no text."
                    
                        
            print "Child node found count: ",cnt

    return(Result)            

#----------------------------------------------------------
# Description:      Function to create request XML string for FD - Lead web service for FD 1.1
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row in excel file
#                   Schemafile - File path of schema file
# Return Values:    request XML string
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Create_FD_Lead_Request_1_1_Phase_I(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath):
    # open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)
    
    # get column numbers from excel
    # PrimaryApplicant section
    PartnerId = Get_ColumnNumber(HeaderList, "PartnerId")   # Calling a function to get column no of "PartnerDealerId" column
    PartnerDealerId = Get_ColumnNumber(HeaderList, "PartnerDealerIds")
    FinanceMethod = Get_ColumnNumber(HeaderList, "FinanceMethod")
    FirstName = Get_ColumnNumber(HeaderList, "FirstName")
    MiddleInitial = Get_ColumnNumber(HeaderList, "MiddleInitial")
    LastName = Get_ColumnNumber(HeaderList, "LastName")
    Suffix = Get_ColumnNumber(HeaderList, "Suffix")
    SSN = Get_ColumnNumber(HeaderList, "SSN")
    DateOfBirth = Get_ColumnNumber(HeaderList, "DateOfBirth")
    EmailAddress = Get_ColumnNumber(HeaderList, "EmailAddress")
    HomePhone = Get_ColumnNumber(HeaderList, "HomePhone")
    AddressLine1 = Get_ColumnNumber(HeaderList, "AddressLine1")
    AddressLine2 = Get_ColumnNumber(HeaderList, "AddressLine2")
    City = Get_ColumnNumber(HeaderList, "City")
    State = Get_ColumnNumber(HeaderList, "State")
    ZipCode = Get_ColumnNumber(HeaderList, "ZipCode")
    TotalMonthsEmployed = Get_ColumnNumber(HeaderList, "TotalMonthsEmployed")
    MonthlyIncome = Get_ColumnNumber(HeaderList, "MonthlyIncome")
    
    # VehicleInfo section
    VehicleCondition = Get_ColumnNumber(HeaderList, "VehicleCondition")
    Year = Get_ColumnNumber(HeaderList, "Year")
    Make = Get_ColumnNumber(HeaderList, "Make")
    Model = Get_ColumnNumber(HeaderList, "Model")
    Trim = Get_ColumnNumber(HeaderList, "Trim")
    ChromeStyleId = Get_ColumnNumber(HeaderList, "ChromeStyleId")
    VIN = Get_ColumnNumber(HeaderList, "VIN")
    StockNumber = Get_ColumnNumber(HeaderList, "StockNumber")
    CertifiedUsed = Get_ColumnNumber(HeaderList, "CertifiedUsed")
    InteriorColor = Get_ColumnNumber(HeaderList, "InteriorColor")
    ExteriorColor = Get_ColumnNumber(HeaderList, "ExteriorColor")
    Mileage = Get_ColumnNumber(HeaderList, "Mileage")

    # LoanInfo section
    CashDown = Get_ColumnNumber(HeaderList, "CashDown")
    Term = Get_ColumnNumber(HeaderList, "Term")
    AmountRequesting = Get_ColumnNumber(HeaderList, "AmountRequesting")
    MonthlyPayment = Get_ColumnNumber(HeaderList, "MonthlyPayment")
    CreditType = Get_ColumnNumber(HeaderList, "CreditType")
    ApplicantType = Get_ColumnNumber(HeaderList, "ApplicantType")

    # TradeInVehicleInfo section
    Trade_Year = Get_ColumnNumber(HeaderList, "Trade_Year")
    Trade_Make = Get_ColumnNumber(HeaderList, "Trade_Make")
    Trade_Model = Get_ColumnNumber(HeaderList, "Trade_Model")
    Trade_Trim = Get_ColumnNumber(HeaderList, "Trade_Trim")
    Trade_Mileage = Get_ColumnNumber(HeaderList, "Trade_Mileage")
    Trade_ChromeStyleId = Get_ColumnNumber(HeaderList, "Trade_ChromeStyleId")
    Trade_TradeInPaid = Get_ColumnNumber(HeaderList, "Trade_TradeInPaid")
    
    Comments = Get_ColumnNumber(HeaderList, "Comments")

    # Open schema file
    try:
        f = open(Schemafile)
    except IOError, e:
        print e
        print "Problem while opening Schema file"
        raise ErrorOccured("Problem in opening file: " + Schemafile)        # Problem opening file
    except:
        raise ErrorOccured("Unknown error occurred.")                        # handle any other type of error

    # everything is fine
    sRequest = f.read()         # get contenets of schema file in to string
    f.close()
    
    # Create request string
    # Calling function to add values to nodes in given schema file
    Cell_Val = sheet.cell(RowNo,PartnerId).value
    sRequest = SetXML(sRequest, "PartnerId", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,PartnerDealerId).value
    sRequest = SetXML(sRequest, "PartnerDealerId",Cell_Val.encode('ascii','ignore') , 0)

    Cell_Val = sheet.cell(RowNo,FinanceMethod).value
    sRequest = SetXML(sRequest, "FinanceMethod",Cell_Val.encode('ascii','ignore') , 0)

    Cell_Val = sheet.cell(RowNo,FirstName).value
    sRequest = SetXML(sRequest, "a:FirstName", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,MiddleInitial).value
    sRequest = SetXML(sRequest, "a:MiddleInitial", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,LastName).value
    sRequest = SetXML(sRequest, "a:LastName", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,Suffix).value
    sRequest = SetXML(sRequest, "a:Suffix", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,SSN).value
    sRequest = SetXML(sRequest, "a:SSN", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,DateOfBirth).value
    sRequest = SetXML(sRequest, "a:DateOfBirth", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,EmailAddress).value
    sRequest = SetXML(sRequest, "a:EmailAddress", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,HomePhone).value
    sRequest = SetXML(sRequest, "a:HomePhone", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,AddressLine1).value
    sRequest = SetXML(sRequest, "a:AddressLine1", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,AddressLine2).value
    sRequest = SetXML(sRequest, "a:AddressLine2",Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,City).value
    sRequest = SetXML(sRequest, "a:City",Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,State).value
    sRequest = SetXML(sRequest, "a:State",Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,ZipCode).value
    sRequest = SetXML(sRequest, "a:ZipCode", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,TotalMonthsEmployed).value
    sRequest = SetXML(sRequest, "a:TotalMonthsEmployed", Cell_Val.encode('ascii','ignore'), 0)
    
    Cell_Val = sheet.cell(RowNo,MonthlyIncome).value
    sRequest = SetXML(sRequest, "a:MonthlyIncome", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,VehicleCondition).value
    sRequest = SetXML(sRequest, "a:VehicleCondition", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,Year).value
    sRequest = SetXML(sRequest, "a:Year", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,Make).value
    sRequest = SetXML(sRequest, "a:Make", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,Model).value
    sRequest = SetXML(sRequest, "a:Model", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,Trim).value
    sRequest = SetXML(sRequest, "a:Trim", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,ChromeStyleId).value
    sRequest = SetXML(sRequest, "a:ChromeStyleId", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,VIN).value
    sRequest = SetXML(sRequest, "a:VIN", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,StockNumber).value
    print Cell_Val
    sRequest = SetXML(sRequest, "a:StockNumber", Cell_Val.encode('ascii','ignore'), 0)
    
    Cell_Val = sheet.cell(RowNo,CertifiedUsed).value
    sRequest = SetXML(sRequest, "a:CertifiedUsed", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,InteriorColor).value
    sRequest = SetXML(sRequest, "a:InteriorColor", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,ExteriorColor).value
    sRequest = SetXML(sRequest, "a:ExteriorColor", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,Mileage).value
    sRequest = SetXML(sRequest, "a:Mileage", Cell_Val.encode('ascii','ignore'), 0)
    
    Cell_Val = sheet.cell(RowNo,CashDown).value
    sRequest = SetXML(sRequest, "a:CashDown", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,Term).value
    sRequest = SetXML(sRequest, "a:Term", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,AmountRequesting).value
    sRequest = SetXML(sRequest, "a:AmountRequesting", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,MonthlyPayment).value
    sRequest = SetXML(sRequest, "a:MonthlyPayment", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,CreditType).value
    sRequest = SetXML(sRequest, "a:CreditType", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,ApplicantType).value
    sRequest = SetXML(sRequest, "a:ApplicantType", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,Trade_Year).value
    sRequest = SetXML(sRequest, "a:Year", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Trade_Make).value
    sRequest = SetXML(sRequest, "a:Make", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Trade_Model).value
    sRequest = SetXML(sRequest, "a:Model", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Trade_Trim).value
    sRequest = SetXML(sRequest, "a:Trim", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Trade_Mileage).value
    sRequest = SetXML(sRequest, "a:Mileage", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Trade_ChromeStyleId).value
    sRequest = SetXML(sRequest, "a:ChromeStyleId", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Trade_TradeInPaid).value
    sRequest = SetXML(sRequest, "a:TradeInPaid", Cell_Val.encode('ascii','ignore'), 0)
    
    Cell_Val = sheet.cell(RowNo,Comments).value
    sRequest = SetXML(sRequest, "Comments", Cell_Val.encode('ascii','ignore'), 0)

   #print sRequest
    
    ReqFilePath = str(FolderPath) + '\\Request\\'+ WebService +'Request'+ str(RowNo) +'.xml'
    f = open(ReqFilePath, 'w')
    f.write(sRequest)
    f.close()
    
    return sRequest                     # return request string

#----------------------------------------------------------
# Description:      Function to create request XML string for FD - Credit Bureau web service for FD 1.1
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row in excel file
#                   Schemafile - File path of schema file
# Return Values:    request XML string
# Author:           Purva Wakode
#----------------------------------------------------------
def Create_FD_CreditBureau_Request_1_1_Phase_I(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath):

    # open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)

    # get column numbers from excel
    PartnerId = Get_ColumnNumber(HeaderList, "PartnerId")                       # Calling a function to get column no of "PartnerDealerId" column
    PartnerDealerId = Get_ColumnNumber(HeaderList, "PartnerDealerId")
    FirstName = Get_ColumnNumber(HeaderList, "FirstName")
    MiddleInitial = Get_ColumnNumber(HeaderList, "MiddleInitial")
    LastName = Get_ColumnNumber(HeaderList, "LastName")
    SSN = Get_ColumnNumber(HeaderList, "SSN")
    DateOfBirth = Get_ColumnNumber(HeaderList, "DateOfBirth")
    HomePhone = Get_ColumnNumber(HeaderList, "HomePhone")
    StreetNumber = Get_ColumnNumber(HeaderList, "StreetNumber")
    StreetName = Get_ColumnNumber(HeaderList, "StreetName")
    StreetType = Get_ColumnNumber(HeaderList, "StreetType")
    City = Get_ColumnNumber(HeaderList, "City")
    State = Get_ColumnNumber(HeaderList, "State")
    ZipCode = Get_ColumnNumber(HeaderList, "ZipCode")
    ApplicantConsent = Get_ColumnNumber(HeaderList, "ApplicantConsent")
    Co_App_FirstName = Get_ColumnNumber(HeaderList, "Co_App_FirstName")
    Co_App_MiddleInitial = Get_ColumnNumber(HeaderList, "Co_App_MiddleInitial")
    Co_App_LastName = Get_ColumnNumber(HeaderList, "Co_App_LastName")
    Co_App_SSN = Get_ColumnNumber(HeaderList, "Co_App_SSN")
    Co_App_DateOfBirth = Get_ColumnNumber(HeaderList, "Co_App_DateOfBirth")
    Co_App_HomePhone = Get_ColumnNumber(HeaderList, "Co_App_HomePhone")
    Co_App_StreetNumber = Get_ColumnNumber(HeaderList, "Co_App_StreetNumber")
    Co_App_StreetName = Get_ColumnNumber(HeaderList, "Co_App_StreetName")
    Co_App_StreetType = Get_ColumnNumber(HeaderList, "Co_App_StreetType")
    Co_App_City = Get_ColumnNumber(HeaderList, "Co_App_City")
    Co_App_State = Get_ColumnNumber(HeaderList, "Co_App_State")
    Co_App_ZipCode = Get_ColumnNumber(HeaderList, "Co_App_ZipCode")
    Co_App_Consent = Get_ColumnNumber(HeaderList, "Co_App_Consent")
    BureauType_0 = Get_ColumnNumber(HeaderList, "BureauType_0")
    BureauType_1 = Get_ColumnNumber(HeaderList, "BureauType_1")
    BureauType_2 = Get_ColumnNumber(HeaderList, "BureauType_2")
    PartnerReferenceId = Get_ColumnNumber(HeaderList, "PartnerReferenceId")


    # Open schema file
    try:
        f = open(Schemafile)
    except IOError, e:
        print e
        print "Problem while opening Schema file"
        raise ErrorOccured("Problem in opening file: " + Schemafile)        # Problem opening file
    except:
        raise ErrorOccured("Unknown error occurred.")                        # handle any other type of error

    # everything is fine
    sRequest = f.read()         # get contenets of schema file in to string
    f.close()

    # change the schema strings for below fields if the vale is null
    # PartnerId
    Cell_Val = sheet.cell(RowNo,PartnerId).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:PartnerId></a:PartnerId>","""<a:PartnerId  i:nil="true" />""",1)
        
    # PartnerId
    Cell_Val = sheet.cell(RowNo,PartnerDealerId).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:PartnerDealerId></a:PartnerDealerId>","""<a:PartnerDealerId  i:nil="true" />""",1)
        
    # PartnerId
    Cell_Val = sheet.cell(RowNo,FirstName).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:FirstName></a:FirstName>","""<a:FirstName  i:nil="true" />""",1)
        
    # MiddleInitial
    Cell_Val = sheet.cell(RowNo,MiddleInitial).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:MiddleInitial></a:MiddleInitial>","""<a:MiddleInitial  i:nil="true" />""",1)
        
    # LastName
    Cell_Val = sheet.cell(RowNo,LastName).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:LastName></a:LastName>","""<a:LastName  i:nil="true" />""",1)
        
    # SSN
    Cell_Val = sheet.cell(RowNo,SSN).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:SSN></a:SSN>","""<a:SSN  i:nil="true" />""",1)
        
    # DateOfBirth
    Cell_Val = sheet.cell(RowNo,DateOfBirth).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:DateOfBirth></a:DateOfBirth>","""<a:DateOfBirth  i:nil="true" />""",1)

    # HomePhone  
    Cell_Val = sheet.cell(RowNo,HomePhone).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:HomePhone></a:HomePhone>","""<a:HomePhone  i:nil="true" />""",1)
        
    # StreetNumber
    Cell_Val = sheet.cell(RowNo,StreetNumber).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:StreetNumber></a:StreetNumber>","""<a:StreetNumber  i:nil="true" />""",1)
        
    # StreetName
    Cell_Val = sheet.cell(RowNo,StreetName).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:StreetName></a:StreetName>","""<a:StreetName  i:nil="true" />""",1)
        
    # StreetType
    Cell_Val = sheet.cell(RowNo,StreetType).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:StreetType></a:StreetType>","""<a:StreetType  i:nil="true" />""",1)
        
    # City
    Cell_Val = sheet.cell(RowNo,City).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:City></a:City>","""<a:City  i:nil="true" />""",1)
        
    # State
    Cell_Val = sheet.cell(RowNo,State).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:State></a:State>","""<a:State  i:nil="true" />""",1)
        
    # ZipCode
    Cell_Val = sheet.cell(RowNo,ZipCode).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:ZipCode></a:ZipCode>","""<a:ZipCode  i:nil="true" />""",1)
        
    # ApplicantConsent
    Cell_Val = sheet.cell(RowNo,ApplicantConsent).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:ApplicantConsent></a:ApplicantConsent>","""<a:ApplicantConsent  i:nil="true" />""",1)
        
    # Co_App_FirstName
    Cell_Val = sheet.cell(RowNo,Co_App_FirstName).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:FirstName></a:FirstName>","""<a:FirstName  i:nil="true" />""",2)
        
    # Co_App_MiddleInitial
    Cell_Val = sheet.cell(RowNo,Co_App_MiddleInitial).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:MiddleInitial></a:MiddleInitial>","""<a:MiddleInitial  i:nil="true" />""",2)
        
    # Co_App_LastName
    Cell_Val = sheet.cell(RowNo,Co_App_LastName).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:LastName></a:LastName>","""<a:LastName  i:nil="true" />""",2)
        
    # Co_App_SSN
    Cell_Val = sheet.cell(RowNo,Co_App_SSN).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:SSN></a:SSN>","""<a:SSN  i:nil="true" />""",2)
        
    # Co_App_DateOfBirth
    Cell_Val = sheet.cell(RowNo,Co_App_DateOfBirth).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:DateOfBirth></a:DateOfBirth>","""<a:DateOfBirth  i:nil="true" />""",2)
        
    # Co_App_HomePhone
    Cell_Val = sheet.cell(RowNo,Co_App_HomePhone).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:HomePhone></a:HomePhone>","""<a:HomePhone  i:nil="true" />""",2)
        
    # Co_App_StreetNumber
    Cell_Val = sheet.cell(RowNo,Co_App_StreetNumber).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:StreetNumber></a:StreetNumber>","""<a:StreetNumber  i:nil="true" />""",2)
        
    # Co_App_StreetName
    Cell_Val = sheet.cell(RowNo,Co_App_StreetName).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:StreetName></a:StreetName>","""<a:StreetName  i:nil="true" />""",2)
        
    # Co_App_StreetType
    Cell_Val = sheet.cell(RowNo,Co_App_StreetType).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:StreetType></a:StreetType>","""<a:StreetType  i:nil="true" />""",2)

    # Co_App_City
    Cell_Val = sheet.cell(RowNo,Co_App_City).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:City></a:City>","""<a:City  i:nil="true" />""",2)

    # Co_App_State
    Cell_Val = sheet.cell(RowNo,Co_App_State).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:State></a:State>","""<a:State  i:nil="true" />""",2)

    # Co_App_ZipCode
    Cell_Val = sheet.cell(RowNo,Co_App_ZipCode).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:ZipCode></a:ZipCode>","""<a:ZipCode  i:nil="true" />""",2)

    # Co_App_Consent
    Cell_Val = sheet.cell(RowNo,Co_App_Consent).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:CoApplicantConsent></a:CoApplicantConsent>","""<a:CoApplicantConsent  i:nil="true" />""",1)

    # BureauType_0
    Cell_Val = sheet.cell(RowNo,BureauType_0).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:BureauType></a:BureauType>","""<a:BureauType  i:nil="true" />""",1)

    # BureauType_1
    Cell_Val = sheet.cell(RowNo,BureauType_1).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:BureauType></a:BureauType>","""<a:BureauType  i:nil="true" />""",1)

    # BureauType_2
    Cell_Val = sheet.cell(RowNo,BureauType_2).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:BureauType></a:BureauType>","""<a:BureauType  i:nil="true" />""",1)

    # PartnerReferenceId
    Cell_Val = sheet.cell(RowNo,PartnerReferenceId).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:PartnerReferenceId></a:PartnerReferenceId>","""<a:PartnerReferenceId  i:nil="true" />""",1)

   
     

    
    # Create request string
    # Calling function to add values to nodes in given schema file

    Cell_Val = sheet.cell(RowNo,PartnerId).value
    sRequest = SetXML(sRequest, "PartnerId", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,PartnerDealerId).value
    sRequest = SetXML(sRequest, "PartnerDealerId",Cell_Val.encode('ascii','ignore') , 0)

    Cell_Val = sheet.cell(RowNo,FirstName).value
    sRequest = SetXML(sRequest, "a:FirstName", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,MiddleInitial).value
    sRequest = SetXML(sRequest, "a:MiddleInitial", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,LastName).value
    sRequest = SetXML(sRequest, "a:LastName", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,SSN).value
    sRequest = SetXML(sRequest, "a:SSN", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,DateOfBirth).value
    sRequest = SetXML(sRequest, "a:DateOfBirth", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,HomePhone).value
    sRequest = SetXML(sRequest, "a:HomePhone",Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,StreetNumber).value
    sRequest = SetXML(sRequest, "a:StreetNumber",Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,StreetName).value
    sRequest = SetXML(sRequest, "a:StreetName",Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,StreetType).value
    sRequest = SetXML(sRequest, "a:StreetType", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,City).value
    sRequest = SetXML(sRequest, "a:City", Cell_Val.encode('ascii','ignore'), 0)
    
    Cell_Val = sheet.cell(RowNo,State).value 
    sRequest = SetXML(sRequest, "a:State", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,ZipCode).value
    sRequest = SetXML(sRequest, "a:ZipCode", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,ApplicantConsent).value
    sRequest = SetXML(sRequest, "ApplicantConsent", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,Co_App_FirstName).value
    sRequest = SetXML(sRequest, "a:FirstName", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Co_App_MiddleInitial).value
    sRequest = SetXML(sRequest, "a:MiddleInitial", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Co_App_LastName).value
    sRequest = SetXML(sRequest, "a:LastName", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Co_App_SSN).value
    sRequest = SetXML(sRequest, "a:SSN", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Co_App_DateOfBirth).value
    sRequest = SetXML(sRequest, "a:DateOfBirth", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Co_App_HomePhone).value
    sRequest = SetXML(sRequest, "a:HomePhone", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Co_App_StreetNumber).value
    sRequest = SetXML(sRequest, "a:StreetNumber", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Co_App_StreetName).value
    sRequest = SetXML(sRequest, "a:StreetName", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Co_App_StreetType).value
    sRequest = SetXML(sRequest, "a:StreetType", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Co_App_City).value
    sRequest = SetXML(sRequest, "a:City", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Co_App_State).value
    sRequest = SetXML(sRequest, "a:State", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Co_App_ZipCode).value
    sRequest = SetXML(sRequest, "a:ZipCode", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Co_App_Consent).value
    sRequest = SetXML(sRequest, "CoApplicantConsent", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,BureauType_0).value
    sRequest = SetXML(sRequest, "a:BureauType", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,BureauType_1).value
    sRequest = SetXML(sRequest, "a:BureauType", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,BureauType_2).value
    sRequest = SetXML(sRequest, "a:BureauType", Cell_Val.encode('ascii','ignore'), 2)

    Cell_Val = sheet.cell(RowNo,PartnerReferenceId).value
    sRequest = SetXML(sRequest, "PartnerReferenceId", Cell_Val.encode('ascii','ignore'), 0)

    
    #print sRequest
    
    ReqFilePath = str(FolderPath) + '\\Request\\'+ WebService +'Request'+ str(RowNo) +'.xml'
    f = open(ReqFilePath, 'w')
    f.write(sRequest)
    f.close()
    
    return sRequest                     # return request string
    
#----------------------------------------------------------
# Description:      Function to create request XML string for FD - CreditApp web service for FD 1.1
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row in excel file
#                   Schemafile - File path of schema file
# Return Values:    request XML string
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Create_FD_CreditApp_Request_1_1_Phase_1(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath):
    # open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)
    
    # get column numbers from excel
    # PrimaryApplicant section
    PartnerId = Get_ColumnNumber(HeaderList, "PartnerId")   # Calling a function to get column no of "PartnerDealerId" column
    PartnerDealerId = Get_ColumnNumber(HeaderList, "PartnerDealerIds")
    FinanceMethod = Get_ColumnNumber(HeaderList, "FinanceMethod")

    FirstName = Get_ColumnNumber(HeaderList, "FirstName")
    MiddleInitial = Get_ColumnNumber(HeaderList, "MiddleInitial")
    LastName = Get_ColumnNumber(HeaderList, "LastName")
    Suffix = Get_ColumnNumber(HeaderList, "Suffix")
    EmailAddress = Get_ColumnNumber(HeaderList, "EmailAddress")
    SSN = Get_ColumnNumber(HeaderList, "SSN")
    DateOfBirth = Get_ColumnNumber(HeaderList, "DateOfBirth")
    DLNumber = Get_ColumnNumber(HeaderList, "DriverLicenseNumber")
    DLState = Get_ColumnNumber(HeaderList, "DriverLicenseState")
    HomePhone = Get_ColumnNumber(HeaderList, "HomePhone")
    AddressLine1 = Get_ColumnNumber(HeaderList, "AddressLine1")
    AddressLine2 = Get_ColumnNumber(HeaderList, "AddressLine2")
    City = Get_ColumnNumber(HeaderList, "City")
    State = Get_ColumnNumber(HeaderList, "State")
    ZipCode = Get_ColumnNumber(HeaderList, "ZipCode")
    HousingStatus = Get_ColumnNumber(HeaderList, "HousingStatus")
    MortgageOrRent = Get_ColumnNumber(HeaderList, "MortgageOrRent")
    TotalMonthsAtAddress = Get_ColumnNumber(HeaderList, "TotalMonthsAtAddress")

    # Prev Address info
    Prev_AddressLine1 = Get_ColumnNumber(HeaderList, "Prev_AddressLine1")
    Prev_AddressLine2 = Get_ColumnNumber(HeaderList, "Prev_AddressLine2")
    Prev_City = Get_ColumnNumber(HeaderList, "Prev_City")
    Prev_State = Get_ColumnNumber(HeaderList, "Prev_State")
    Prev_ZipCode = Get_ColumnNumber(HeaderList, "Prev_ZipCode")

    # Current Emp
    EmploymentStatus = Get_ColumnNumber(HeaderList, "EmploymentStatus")
    EmployedBy = Get_ColumnNumber(HeaderList, "EmployedBy")
    TotalMonthsEmployed = Get_ColumnNumber(HeaderList, "TotalMonthsEmployed")
    BusinessPhone = Get_ColumnNumber(HeaderList, "BusinessPhone")
    MonthlyIncome = Get_ColumnNumber(HeaderList, "MonthlyIncome")

    # Prev Emp
    Prev_EmploymentStatus = Get_ColumnNumber(HeaderList, "Prev_EmploymentStatus")
    Prev_EmployedBy = Get_ColumnNumber(HeaderList, "Prev_EmployedBy")
    Prev_TotalMonthsEmployed = Get_ColumnNumber(HeaderList, "Prev_TotalMonthsEmployed")
    Prev_BusinessPhone = Get_ColumnNumber(HeaderList, "Prev_BusinessPhone")

    OtherMonthlyIncome = Get_ColumnNumber(HeaderList, "OtherMonthlyIncome")
    OtherIncomeSource = Get_ColumnNumber(HeaderList, "OtherIncomeSource")
    ConsentIndicator = Get_ColumnNumber(HeaderList, "ConsentIndicator")

    # Co- Applicant section

    Coapp_FirstName = Get_ColumnNumber(HeaderList, "Coapp_FirstName")
    Coapp_MiddleInitial = Get_ColumnNumber(HeaderList, "Coapp_MiddleInitial")
    Coapp_LastName = Get_ColumnNumber(HeaderList, "Coapp_LastName")
    Coapp_Suffix = Get_ColumnNumber(HeaderList, "Coapp_Suffix")
    Coapp_EmailAddress = Get_ColumnNumber(HeaderList, "Coapp_EmailAddress")
    Coapp_SSN = Get_ColumnNumber(HeaderList, "Coapp_SSN")
    Coapp_DateOfBirth = Get_ColumnNumber(HeaderList, "Coapp_DateOfBirth")
    Coapp_DLNumber = Get_ColumnNumber(HeaderList, "Coapp_DriverLicenseNumber")
    Coapp_DLState = Get_ColumnNumber(HeaderList, "Coapp_DriverLicenseState")
    Coapp_HomePhone = Get_ColumnNumber(HeaderList, "Coapp_HomePhone")
    Coapp_AddressLine1 = Get_ColumnNumber(HeaderList, "Coapp_AddressLine1")
    Coapp_AddressLine2 = Get_ColumnNumber(HeaderList, "Coapp_AddressLine2")
    Coapp_City = Get_ColumnNumber(HeaderList, "Coapp_City")
    Coapp_State = Get_ColumnNumber(HeaderList, "Coapp_State")
    Coapp_ZipCode = Get_ColumnNumber(HeaderList, "Coapp_ZipCode")
    Coapp_HousingStatus = Get_ColumnNumber(HeaderList, "Coapp_HousingStatus")
    Coapp_MortgageOrRent = Get_ColumnNumber(HeaderList, "Coapp_MortgageOrRent")
    Coapp_TotalMonthsAtAddress = Get_ColumnNumber(HeaderList, "Coapp_TotalMonthsAtAddress")
    
    # Co-App Prev address
    Coapp_Prev_AddressLine1 = Get_ColumnNumber(HeaderList, "Coapp_Prev_AddressLine1")
    Coapp_Prev_AddressLine2 = Get_ColumnNumber(HeaderList, "Coapp_Prev_AddressLine2")
    Coapp_Prev_City = Get_ColumnNumber(HeaderList, "Coapp_Prev_City")
    Coapp_Prev_State = Get_ColumnNumber(HeaderList, "Coapp_Prev_State")
    Coapp_Prev_ZipCode = Get_ColumnNumber(HeaderList, "Coapp_Prev_ZipCode")

    # Co-App Emp
    Coapp_EmploymentStatus = Get_ColumnNumber(HeaderList, "Coapp_EmploymentStatus")
    Coapp_EmployedBy = Get_ColumnNumber(HeaderList, "Coapp_EmployedBy")
    Coapp_TotalMonthsEmployed = Get_ColumnNumber(HeaderList, "Coapp_TotalMonthsEmployed")
    Coapp_BusinessPhone = Get_ColumnNumber(HeaderList, "Coapp_BusinessPhone")
    Coapp_MonthlyIncome = Get_ColumnNumber(HeaderList, "Coapp_MonthlyIncome")
    
    # Co-App Prev Emp
    Coapp_Prev_EmploymentStatus = Get_ColumnNumber(HeaderList, "Coapp_Prev_EmploymentStatus")
    Coapp_Prev_EmployedBy = Get_ColumnNumber(HeaderList, "Coapp_Prev_EmployedBy")
    Coapp_Prev_TotalMonthsEmployed = Get_ColumnNumber(HeaderList, "Coapp_Prev_TotalMonthsEmployed")
    Coapp_Prev_BusinessPhone = Get_ColumnNumber(HeaderList, "Coapp_Prev_BusinessPhone")

    Coapp_OtherMonthlyIncome = Get_ColumnNumber(HeaderList, "Coapp_OtherMonthlyIncome")
    Coapp_OtherIncomeSource = Get_ColumnNumber(HeaderList, "Coapp_OtherIncomeSource")
    Relationship = Get_ColumnNumber(HeaderList, "Relationship")

    # VehicleInfo section
    VehicleCondition = Get_ColumnNumber(HeaderList, "VehicleCondition")
    Year = Get_ColumnNumber(HeaderList, "Year")
    Make = Get_ColumnNumber(HeaderList, "Make")
    Model = Get_ColumnNumber(HeaderList, "Model")
    Trim = Get_ColumnNumber(HeaderList, "Trim")
    ChromeStyleId = Get_ColumnNumber(HeaderList, "ChromeStyleId")
    VIN = Get_ColumnNumber(HeaderList, "VIN")
    StockNumber = Get_ColumnNumber(HeaderList, "StockNumber")
    CertifiedUsed = Get_ColumnNumber(HeaderList, "CertifiedUsed")
    

    # Product Info section
    CashSellingPrice = Get_ColumnNumber(HeaderList, "CashSellingPrice")
    SalesTax = Get_ColumnNumber(HeaderList, "SalesTax")
    Title = Get_ColumnNumber(HeaderList, "Title")
    CashDown = Get_ColumnNumber(HeaderList, "CashDown")
    Rebate = Get_ColumnNumber(HeaderList, "Rebate")
    CreditLifeIns = Get_ColumnNumber(HeaderList, "CreditLifeIns")
    Term = Get_ColumnNumber(HeaderList, "Term")
    AcquisitionFees = Get_ColumnNumber(HeaderList, "AcquisitionFees")
    InvoiceAmount = Get_ColumnNumber(HeaderList, "InvoiceAmount")
    Warranty = Get_ColumnNumber(HeaderList, "Warranty")
    MSRP = Get_ColumnNumber(HeaderList, "MSRP")
    EstimatedBalloonAmount = Get_ColumnNumber(HeaderList, "EstimatedBalloonAmount")
    EstimatedPayment = Get_ColumnNumber(HeaderList, "EstimatedPayment")
    UsedCarBook = Get_ColumnNumber(HeaderList, "UsedCarBook")
    Mileage = Get_ColumnNumber(HeaderList, "Mileage")
    UsedCarValue = Get_ColumnNumber(HeaderList, "UsedCarValue")
    OtherFees = Get_ColumnNumber(HeaderList, "OtherFees")
    WholesaleBookSource = Get_ColumnNumber(HeaderList, "WholesaleBookSource")
    WholesaleCondition = Get_ColumnNumber(HeaderList, "WholesaleCondition")
    WholesaleValueType = Get_ColumnNumber(HeaderList, "WholesaleValueType")
    WholesaleValue = Get_ColumnNumber(HeaderList, "WholesaleValue")
    NetTrade = Get_ColumnNumber(HeaderList, "NetTrade")
    
    # TradeInVehicleInfo section
    Trade_Year = Get_ColumnNumber(HeaderList, "Trade_Year")
    Trade_Make = Get_ColumnNumber(HeaderList, "Trade_Make")
    Trade_Model = Get_ColumnNumber(HeaderList, "Trade_Model")
    Trade_Trim = Get_ColumnNumber(HeaderList, "Trade_Trim")
    Trade_ChromeStyleId = Get_ColumnNumber(HeaderList, "Trade_ChromeStyleId")

    LienHolder = Get_ColumnNumber(HeaderList, "LienHolder")
    MonthlyPayment = Get_ColumnNumber(HeaderList, "MonthlyPayment")
    
    LeadComments = Get_ColumnNumber(HeaderList, "LeadComments")

    # reference numbers
    PrequalificationReferenceNumber = Get_ColumnNumber(HeaderList, "PrequalificationReferenceNumber")
    LeadReferenceNumber = Get_ColumnNumber(HeaderList, "LeadReferenceNumber")

    

    # Open schema file
    try:
        f = open(Schemafile)
    except IOError, e:
        print e
        print "Problem while opening Schema file"
        raise ErrorOccured("Problem in opening file: " + Schemafile)        # Problem opening file
    except:
        raise ErrorOccured("Unknown error occurred.")                        # handle any other type of error

    # everything is fine
    sRequest = f.read()         # get contenets of schema file in to string
    f.close()

    #VehicleContion_Val = sheet.cell(RowNo,VehicleCondition).value
    #New_Req=""

    # change the schema strings for below fields if the vale is null
    
    # PartnerId
    Cell_Val = sheet.cell(RowNo,PartnerId).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:PartnerId></a:PartnerId>","""<a:PartnerId  i:nil="true" />""",1)
    
    # DOBs
    Cell_Val = sheet.cell(RowNo,DateOfBirth).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:DateOfBirth></a:DateOfBirth>","""<a:DateOfBirth  i:nil="true" />""",1)

    Cell_Val = sheet.cell(RowNo,Coapp_DateOfBirth).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:DateOfBirth></a:DateOfBirth>","""<a:DateOfBirth  i:nil="true" />""",2)

    # suffix
    Cell_Val = sheet.cell(RowNo,Suffix).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:Suffix></a:Suffix>","""<a:Suffix  i:nil="true" />""",1)

    Cell_Val = sheet.cell(RowNo,Coapp_Suffix).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:Suffix></a:Suffix>","""<a:Suffix  i:nil="true" />""",2)

    # addressline 2
    Cell_Val = sheet.cell(RowNo,AddressLine2).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest, "<a:AddressLine2></a:AddressLine2>", """<a:AddressLine2 i:nil="true" />""", 1) 

    Cell_Val = sheet.cell(RowNo,Coapp_AddressLine2).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest, "<a:AddressLine2></a:AddressLine2>", """<a:AddressLine2 i:nil="true" />""", 3)

    # housing status
    Cell_Val = sheet.cell(RowNo,HousingStatus).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest, "<a:HousingStatus></a:HousingStatus>", """<a:HousingStatus i:nil="true" />""", 1)

    Cell_Val = sheet.cell(RowNo,Coapp_HousingStatus).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:HousingStatus></a:HousingStatus>","""<a:HousingStatus  i:nil="true" />""",2)

    # monthly income
    Cell_Val = sheet.cell(RowNo,MonthlyIncome).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest, "<a:MonthlyIncome></a:MonthlyIncome>", """<a:MonthlyIncome i:nil="true" />""", 1)

    Cell_Val = sheet.cell(RowNo,Coapp_MonthlyIncome).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:MonthlyIncome></a:MonthlyIncome>","""<a:MonthlyIncome  i:nil="true" />""",2)

    # Emp Status
    Cell_Val = sheet.cell(RowNo,EmploymentStatus).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest, "<a:EmploymentStatus></a:EmploymentStatus>", """<a:EmploymentStatus i:nil="true" />""", 1)

    Cell_Val = sheet.cell(RowNo,Coapp_EmploymentStatus).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:EmploymentStatus></a:EmploymentStatus>","""<a:EmploymentStatus  i:nil="true" />""",3)

    # Emp By
    Cell_Val = sheet.cell(RowNo,EmployedBy).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest, "<a:EmployedBy></a:EmployedBy>", """<a:EmployedBy i:nil="true" />""", 1)

    Cell_Val = sheet.cell(RowNo,Coapp_EmployedBy).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:EmployedBy></a:EmployedBy>","""<a:EmployedBy  i:nil="true" />""",3)

    # Total months Employed
    Cell_Val = sheet.cell(RowNo,TotalMonthsEmployed).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest, "<a:TotalMonthsEmployed></a:TotalMonthsEmployed>", """<a:TotalMonthsEmployed i:nil="true" />""", 1)

    Cell_Val = sheet.cell(RowNo,Coapp_TotalMonthsEmployed).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:TotalMonthsEmployed></a:TotalMonthsEmployed>","""<a:TotalMonthsEmployed  i:nil="true" />""",3)

    # BusinessPhone 
    Cell_Val = sheet.cell(RowNo,BusinessPhone).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest, "<a:BusinessPhone></a:BusinessPhone>", """<a:BusinessPhone i:nil="true" />""", 1)

    Cell_Val = sheet.cell(RowNo,Coapp_BusinessPhone).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:BusinessPhone></a:BusinessPhone>","""<a:BusinessPhone  i:nil="true" />""",3)

     #  Prev_EmploymentStatus 
    Cell_Val = sheet.cell(RowNo,Prev_EmploymentStatus).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest, "<a:EmploymentStatus></a:EmploymentStatus>", """<a:EmploymentStatus i:nil="true" />""", 2)

    Cell_Val = sheet.cell(RowNo,Coapp_Prev_EmploymentStatus).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:EmploymentStatus></a:EmploymentStatus>","""<a:EmploymentStatus  i:nil="true" />""",4)
        
    # Prev_EmployedBy
    Cell_Val = sheet.cell(RowNo,Prev_EmployedBy).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest, "<a:EmployedBy></a:EmployedBy>", """<a:EmployedBy i:nil="true" />""", 2)

    Cell_Val = sheet.cell(RowNo,Coapp_Prev_EmployedBy).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:EmployedBy></a:EmployedBy>","""<a:EmployedBy  i:nil="true" />""",4)
    
    # Prev_TotalMonthsEmployed
    Cell_Val = sheet.cell(RowNo,Prev_TotalMonthsEmployed).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest, "<a:TotalMonthsEmployed></a:TotalMonthsEmployed>", """<a:TotalMonthsEmployed i:nil="true" />""", 2)

    Cell_Val = sheet.cell(RowNo,Coapp_Prev_TotalMonthsEmployed).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:TotalMonthsEmployed></a:TotalMonthsEmployed>","""<a:TotalMonthsEmployed  i:nil="true" />""",4)
        
    #Prev_BusinessPhone
    Cell_Val = sheet.cell(RowNo,Prev_BusinessPhone).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest, "<a:BusinessPhone></a:BusinessPhone>", """<a:BusinessPhone i:nil="true" />""", 2)

    Cell_Val = sheet.cell(RowNo,Coapp_Prev_BusinessPhone).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:BusinessPhone></a:BusinessPhone>","""<a:BusinessPhone  i:nil="true" />""",4)

    # Other Monthly income
    Cell_Val = sheet.cell(RowNo,OtherMonthlyIncome).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest, "<a:OtherMonthlyIncome></a:OtherMonthlyIncome>", """<a:OtherMonthlyIncome i:nil="true" />""", 1)

    Cell_Val = sheet.cell(RowNo,Coapp_OtherMonthlyIncome).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:OtherMonthlyIncome></a:OtherMonthlyIncome>","""<a:OtherMonthlyIncome  i:nil="true" />""",2)

    # Relationship
    Cell_Val = sheet.cell(RowNo,Relationship).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest, "<a:Relationship></a:Relationship>", """<a:Relationship i:nil="true" />""", 1)

   
    # ****************************** Product Info ****************************************
    # CashSellingPrice
    Cell_Val = sheet.cell(RowNo,CashSellingPrice).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:CashSellingPrice></a:CashSellingPrice>", """<a:CashSellingPrice i:nil="true" />""",1)

    # Sales Tax
    Cell_Val = sheet.cell(RowNo,SalesTax).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:SalesTax></a:SalesTax>", """<a:SalesTax i:nil="true" />""",1)

    # Net Trade
    Cell_Val = sheet.cell(RowNo,NetTrade).value
    if len(Cell_Val) == 0:
        #sRequest = replacenth(sRequest, "<a:NetTrade></a:NetTrade>", """<a:NetTrade i:nil="true" />""", 1)
        sRequest = sRequest.replace("<a:NetTrade></a:NetTrade>", """<a:NetTrade i:nil="true" />""",1)
   
    # Title
    Cell_Val = sheet.cell(RowNo,Title).value
    if len(Cell_Val) == 0:
        #sRequest = replacenth(sRequest, "<a:Title></a:Title>", """<a:Title i:nil="true" />""", 1)
        sRequest = sRequest.replace("<a:Title></a:Title>", """<a:Title i:nil="true" />""",1)

    # CashDown
    Cell_Val = sheet.cell(RowNo,CashDown).value
    if len(Cell_Val) == 0:
        #sRequest = replacenth(sRequest, "<a:CashDown></a:CashDown>", """<a:CashDown i:nil="true" />""", 1)
        sRequest = sRequest.replace("<a:CashDown></a:CashDown>", """<a:CashDown i:nil="true" />""",1)

    # Rebate
    Cell_Val = sheet.cell(RowNo,Rebate).value
    if len(Cell_Val) == 0:
        #sRequest = replacenth(sRequest, "<a:Rebate></a:Rebate>", """<a:Rebate i:nil="true" />""", 1)
        sRequest = sRequest.replace("<a:Rebate></a:Rebate>", """<a:Rebate i:nil="true" />""",1)

    # CreditLifeIns
    Cell_Val = sheet.cell(RowNo,CreditLifeIns).value
    if len(Cell_Val) == 0:
        #sRequest = replacenth(sRequest, "<a:CreditLifeIns></a:CreditLifeIns>", """<a:CreditLifeIns i:nil="true" />""", 1)
        sRequest = sRequest.replace("<a:CreditLifeIns></a:CreditLifeIns>", """<a:CreditLifeIns i:nil="true" />""",1)

    # Term
    Cell_Val = sheet.cell(RowNo,Term).value
    if len(Cell_Val) == 0:
        #sRequest = replacenth(sRequest, "<a:Term></a:Term>", """<a:Term i:nil="true" />""", 1)
        sRequest = sRequest.replace("<a:Term></a:Term>", """<a:Term i:nil="true" />""",1)

    # AcquisitionFees
    Cell_Val = sheet.cell(RowNo,AcquisitionFees).value
    if len(Cell_Val) == 0:
        #sRequest = replacenth(sRequest, "<a:AcquisitionFees></a:AcquisitionFees>", """<a:AcquisitionFees i:nil="true" />""", 1)
        sRequest = sRequest.replace("<a:AcquisitionFees></a:AcquisitionFees>", """<a:AcquisitionFees i:nil="true" />""",1)

    # InvoiceAmount
    Cell_Val = sheet.cell(RowNo,InvoiceAmount).value
    if len(Cell_Val) == 0:
        #sRequest = replacenth(sRequest, "<a:InvoiceAmount></a:InvoiceAmount>", """<a:InvoiceAmount i:nil="true" />""", 1)
        sRequest = sRequest.replace("<a:InvoiceAmount></a:InvoiceAmount>", """<a:InvoiceAmount i:nil="true" />""",1)

    # Warranty
    Cell_Val = sheet.cell(RowNo,Warranty).value
    if len(Cell_Val) == 0:
        #sRequest = replacenth(sRequest, "<a:Warranty></a:Warranty>", """<a:Warranty i:nil="true" />""", 1)
        sRequest = sRequest.replace("<a:Warranty></a:Warranty>", """<a:Warranty i:nil="true" />""",1)

    # MSRP
    Cell_Val = sheet.cell(RowNo,MSRP).value
    if len(Cell_Val) == 0:
        #sRequest = replacenth(sRequest, "<a:Warranty></a:Warranty>", """<a:Warranty i:nil="true" />""", 1)
        sRequest = sRequest.replace("<a:MSRP></a:MSRP>", """<a:MSRP i:nil="true" />""",1)

    # EstimatedBalloonAmount
    Cell_Val = sheet.cell(RowNo,EstimatedBalloonAmount).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:EstimatedBalloonAmount></a:EstimatedBalloonAmount>", """<a:EstimatedBalloonAmount i:nil="true" />""",1)

    # EstimatedPayment
    Cell_Val = sheet.cell(RowNo,EstimatedPayment).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:EstimatedPayment></a:EstimatedPayment>", """<a:EstimatedPayment i:nil="true" />""",1)

    # OtherFees
    Cell_Val = sheet.cell(RowNo,OtherFees).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:OtherFees></a:OtherFees>", """<a:OtherFees i:nil="true" />""",1)

    # Used car book
    Cell_Val = sheet.cell(RowNo,UsedCarBook).value
    if len(Cell_Val) != 0:
        sRequest = sRequest.replace("""<a:UsedCarBook  i:nil="true" />""","<a:UsedCarBook></a:UsedCarBook>",1)

    # Mileage
    Cell_Val = sheet.cell(RowNo,Mileage).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:Mileage></a:Mileage>", """<a:Mileage i:nil="true" />""",1)       

    # used car value
    Cell_Val = sheet.cell(RowNo,UsedCarValue).value
    if len(Cell_Val) != 0:
        sRequest = sRequest.replace("""<a:UsedCarValue i:nil="true" />""","<a:UsedCarValue></a:UsedCarValue>",1)

    # whole sale book source
    Cell_Val = sheet.cell(RowNo,WholesaleBookSource).value
    if len(Cell_Val) != 0:
        sRequest = sRequest.replace("""<a:WholesaleBookSource i:nil="true" />""","<a:WholesaleBookSource></a:WholesaleBookSource>",1)

    # WholesaleCondition
    Cell_Val = sheet.cell(RowNo,WholesaleCondition).value
    if len(Cell_Val) != 0:
        sRequest = sRequest.replace("""<a:WholesaleCondition i:nil="true" />""","<a:WholesaleCondition></a:WholesaleCondition>",1)

    # WholesaleValueType
    Cell_Val = sheet.cell(RowNo,WholesaleValueType).value
    if len(Cell_Val) != 0:
        sRequest = sRequest.replace("""<a:WholesaleValueType i:nil="true" />""","<a:WholesaleValueType></a:WholesaleValueType>",1)

    # WholesaleValue
    Cell_Val = sheet.cell(RowNo,WholesaleValue).value
    if len(Cell_Val) != 0:
        sRequest = sRequest.replace("""<a:WholesaleValue i:nil="true" />""","<a:WholesaleValue></a:WholesaleValue>",1)

    # ****************************** Vehicle Info ****************************************
    # VehicleCondition
    Cell_Val = sheet.cell(RowNo,VehicleCondition).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:VehicleCondition></a:VehicleCondition>", """<a:VehicleCondition i:nil="true" />""",1)

    # Veh Year - Trade year
    Cell_Val = sheet.cell(RowNo,Year).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest, "<a:Year></a:Year>", """<a:Year i:nil="true" />""", 1)

    Cell_Val = sheet.cell(RowNo,Trade_Year).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:Year></a:Year>","""<a:Year  i:nil="true" />""",2)

    # Veh make - trade make
    Cell_Val = sheet.cell(RowNo,Make).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest, "<a:Make></a:Make>", """<a:Make i:nil="true" />""", 1)

    Cell_Val = sheet.cell(RowNo,Trade_Make).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:Make></a:Make>","""<a:Make  i:nil="true" />""",2)

    # Veh Model - trade Model
    Cell_Val = sheet.cell(RowNo,Model).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest, "<a:Model></a:Model>", """<a:Model i:nil="true" />""", 1)

    Cell_Val = sheet.cell(RowNo,Trade_Model).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:Model></a:Model>","""<a:Model  i:nil="true" />""",2)
        
    # Veh Trim - trade Trim
    Cell_Val = sheet.cell(RowNo,Trim).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest, "<a:Trim></a:Trim>", """<a:Trim i:nil="true" />""", 1)

    Cell_Val = sheet.cell(RowNo,Trade_Trim).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:Trim></a:Trim>","""<a:Trim  i:nil="true" />""",2)

    # Veh ChromeStyleId - Trade ChromeStyleId
    Cell_Val = sheet.cell(RowNo,ChromeStyleId).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest, "<a:ChromeStyleId></a:ChromeStyleId>", """<a:ChromeStyleId i:nil="true" />""", 1)

    Cell_Val = sheet.cell(RowNo,Trade_ChromeStyleId).value
    if len(Cell_Val) == 0:
        sRequest = replacenth(sRequest,"<a:ChromeStyleId></a:ChromeStyleId>","""<a:ChromeStyleId  i:nil="true" />""",2)

    # veh - VIN
    Cell_Val = sheet.cell(RowNo,VIN).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:VIN></a:VIN>", """<a:VIN i:nil="true" />""",1)

    # Veh - StockNumber
    Cell_Val = sheet.cell(RowNo,StockNumber).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:StockNumber></a:StockNumber>", """<a:StockNumber i:nil="true" />""",1)

    # Veh - CertifiedUsed
    Cell_Val = sheet.cell(RowNo,CertifiedUsed).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:CertifiedUsed></a:CertifiedUsed>", """<a:CertifiedUsed i:nil="true" />""",1)

    #  Trade In Vechicle LienHolder
    Cell_Val = sheet.cell(RowNo,LienHolder).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:LienHolder></a:LienHolder>", """<a:LienHolder i:nil="true" />""",1)

    #  Trade In Vechicle MonthlyPayment
    Cell_Val = sheet.cell(RowNo,MonthlyPayment).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:MonthlyPayment></a:MonthlyPayment>", """<a:MonthlyPayment i:nil="true" />""",1)

    # ConsentIndicator
    Cell_Val = sheet.cell(RowNo,ConsentIndicator).value
    if len(Cell_Val) == 0:
        sRequest = sRequest.replace("<a:ConsentIndicator></a:ConsentIndicator>", """<a:ConsentIndicator i:nil="true" />""",1)

    # reference numbers
    # Prequal
    Cell_Val = sheet.cell(RowNo,PrequalificationReferenceNumber).value
    print "PreQ ref no: ",Cell_Val
    if len(Cell_Val) != 0:
        sRequest = sRequest.replace("""<PrequalificationReferenceNumber i:nil="true" />""","<PrequalificationReferenceNumber></PrequalificationReferenceNumber>",1)
        
    # Lead
    Cell_Val = sheet.cell(RowNo,LeadReferenceNumber).value
    print "Lead ref no: ",Cell_Val
    if len(Cell_Val) != 0:
        sRequest = sRequest.replace("""<LeadReferenceNumber i:nil="true" />""","<LeadReferenceNumber></LeadReferenceNumber>",1)
        

    #print "************************** Request Before values ***********************"
#    print sRequest
 #   print "***********************************************************************"
    #if VehicleContion_Val != 'New':
     #   New_Req = sRequest.replace("""<a:UsedCarBook  i:nil="true" />""","<a:UsedCarBook></a:UsedCarBook>")
      #  New_Req = New_Req.replace("""<a:UsedCarValue i:nil="true" />""","<a:UsedCarValue></a:UsedCarValue>")
       # New_Req = New_Req.replace("""<a:WholesaleBookSource i:nil="true" />""","<a:WholesaleBookSource></a:WholesaleBookSource>")
       # New_Req = New_Req.replace("""<a:WholesaleCondition i:nil="true" />""","<a:WholesaleCondition></a:WholesaleCondition>")
        #New_Req = New_Req.replace("""<a:WholesaleValueType i:nil="true" />""","<a:WholesaleValueType></a:WholesaleValueType>")
        #New_Req = New_Req.replace("""<a:WholesaleValue i:nil="true" />""","<a:WholesaleValue></a:WholesaleValue>")
        #sRequest = New_Req
    
    
    # Create request string
    # Calling function to add values to nodes in given schema file
    Cell_Val = sheet.cell(RowNo,PartnerId).value
    #print "Partner ID: ",Cell_Val
    sRequest = SetXML(sRequest, "PartnerId", Cell_Val.encode('ascii','ignore'), 0)

    #print "************** Add Partner ID"
    #print sRequest

    Cell_Val = sheet.cell(RowNo,PartnerDealerId).value
    sRequest = SetXML(sRequest, "PartnerDealerId",Cell_Val.encode('ascii','ignore') , 0)

    Cell_Val = sheet.cell(RowNo,FinanceMethod).value
    sRequest = SetXML(sRequest, "FinanceMethod",Cell_Val.encode('ascii','ignore') , 0)

    # App info
    Cell_Val = sheet.cell(RowNo,FirstName).value
    sRequest = SetXML(sRequest, "a:FirstName", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,MiddleInitial).value
    sRequest = SetXML(sRequest, "a:MiddleInitial", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,LastName).value
    sRequest = SetXML(sRequest, "a:LastName", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,Suffix).value
    #if len(Cell_Val) == 0:
     #   sRequest = sRequest.replace("<a:Suffix></a:Suffix>","""<a:Suffix  i:nil="true" />""",1)
    #else:
#        sRequest = SetXML(sRequest, "a:Suffix", Cell_Val.encode('ascii','ignore'), 0)
    sRequest = SetXML(sRequest, "a:Suffix", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,EmailAddress).value
    sRequest = SetXML(sRequest, "a:EmailAddress", Cell_Val.encode('ascii','ignore'), 0)
    
    Cell_Val = sheet.cell(RowNo,SSN).value
    sRequest = SetXML(sRequest, "a:SSN", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,DateOfBirth).value
    #if len(Cell_Val) == 0:
     #   sRequest = sRequest.replace("<a:DateOfBirth></a:DateOfBirth>","""<a:DateOfBirth  i:nil="true" />""",1)
#    else:
 #       sRequest = SetXML(sRequest, "a:DateOfBirth", Cell_Val.encode('ascii','ignore'), 0)
    sRequest = SetXML(sRequest, "a:DateOfBirth", Cell_Val.encode('ascii','ignore'), 0)
   
    Cell_Val = sheet.cell(RowNo,DLNumber).value
    sRequest = SetXML(sRequest, "a:DriverLicenseNumber", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,DLState).value
    sRequest = SetXML(sRequest, "a:DriverLicenseState", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,HomePhone).value
    sRequest = SetXML(sRequest, "a:HomePhone", Cell_Val.encode('ascii','ignore'), 0)

    # App Current address
    Cell_Val = sheet.cell(RowNo,AddressLine1).value
    sRequest = SetXML(sRequest, "a:AddressLine1", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,AddressLine2).value
    #if len(Cell_Val) == 0:
     #   sRequest = replacenth(sRequest, "<a:AddressLine2></a:AddressLine2>", """<a:AddressLine2 i:nil="true" />""", 1)
#    else:
 #       sRequest = SetXML(sRequest, "a:AddressLine2",Cell_Val.encode('ascii','ignore'), 0)
    sRequest = SetXML(sRequest, "a:AddressLine2",Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,City).value
    sRequest = SetXML(sRequest, "a:City",Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,State).value
    sRequest = SetXML(sRequest, "a:State",Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,ZipCode).value
    sRequest = SetXML(sRequest, "a:ZipCode", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,HousingStatus).value
    sRequest = SetXML(sRequest, "a:HousingStatus", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,MortgageOrRent).value
    sRequest = SetXML(sRequest, "a:MortgageOrRent", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,TotalMonthsAtAddress).value
    sRequest = SetXML(sRequest, "a:TotalMonthsAtAddress", Cell_Val.encode('ascii','ignore'), 0)

    # App Prev address
    Cell_Val = sheet.cell(RowNo,Prev_AddressLine1).value
    sRequest = SetXML(sRequest, "a:AddressLine1", Cell_Val.encode('ascii','ignore'), 1)    
    
    Cell_Val = sheet.cell(RowNo,Prev_AddressLine2).value
    sRequest = SetXML(sRequest, "a:AddressLine2", Cell_Val.encode('ascii','ignore'), 1)
    
    Cell_Val = sheet.cell(RowNo,Prev_City).value
    sRequest = SetXML(sRequest, "a:City", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Prev_State).value
    sRequest = SetXML(sRequest, "a:State", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Prev_ZipCode).value
    sRequest = SetXML(sRequest, "a:ZipCode", Cell_Val.encode('ascii','ignore'), 1)

    # App Current Emp
    Cell_Val = sheet.cell(RowNo,EmploymentStatus).value
    sRequest = SetXML(sRequest, "a:EmploymentStatus", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,EmployedBy).value
    sRequest = SetXML(sRequest, "a:EmployedBy", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,TotalMonthsEmployed).value
    sRequest = SetXML(sRequest, "a:TotalMonthsEmployed", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,BusinessPhone).value
    sRequest = SetXML(sRequest, "a:BusinessPhone", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,MonthlyIncome).value
    sRequest = SetXML(sRequest, "a:MonthlyIncome", Cell_Val.encode('ascii','ignore'), 0)

    # App Prev Emp
    Cell_Val = sheet.cell(RowNo,Prev_EmploymentStatus).value
    sRequest = SetXML(sRequest, "a:EmploymentStatus", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Prev_EmployedBy).value
    sRequest = SetXML(sRequest, "a:EmployedBy", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Prev_TotalMonthsEmployed).value
    sRequest = SetXML(sRequest, "a:TotalMonthsEmployed", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Prev_BusinessPhone).value
    sRequest = SetXML(sRequest, "a:BusinessPhone", Cell_Val.encode('ascii','ignore'), 1)

    # Others

    Cell_Val = sheet.cell(RowNo,OtherMonthlyIncome).value
    sRequest = SetXML(sRequest, "a:OtherMonthlyIncome", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,OtherIncomeSource).value
    sRequest = SetXML(sRequest, "a:OtherIncomeSource", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,ConsentIndicator).value
    sRequest = SetXML(sRequest, "a:ConsentIndicator", Cell_Val.encode('ascii','ignore'), 0)

    # Co-App *******************************************************************************************
    # Co-App info
    Cell_Val = sheet.cell(RowNo,Coapp_FirstName).value
    sRequest = SetXML(sRequest, "a:FirstName", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Coapp_MiddleInitial).value
    sRequest = SetXML(sRequest, "a:MiddleInitial", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Coapp_LastName).value
    sRequest = SetXML(sRequest, "a:LastName", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Coapp_Suffix).value
    #if len(Cell_Val) == 0:
     #   sRequest = replacenth(sRequest,"<a:Suffix></a:Suffix>","""<a:Suffix  i:nil="true" />""",2)
    #else:
#        sRequest = SetXML(sRequest, "a:Suffix", Cell_Val.encode('ascii','ignore'), 1)
    sRequest = SetXML(sRequest, "a:Suffix", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Coapp_EmailAddress).value
    sRequest = SetXML(sRequest, "a:EmailAddress", Cell_Val.encode('ascii','ignore'), 1)
    
    Cell_Val = sheet.cell(RowNo,Coapp_SSN).value
    sRequest = SetXML(sRequest, "a:SSN", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Coapp_DateOfBirth).value
    #if len(Cell_Val) == 0:
     #   sRequest = replacenth(sRequest,"<a:DateOfBirth></a:DateOfBirth>","""<a:DateOfBirth  i:nil="true" />""",2)
#    else:
#        sRequest = SetXML(sRequest, "a:DateOfBirth", Cell_Val.encode('ascii','ignore'), 1)
    sRequest = SetXML(sRequest, "a:DateOfBirth", Cell_Val.encode('ascii','ignore'), 1)
   
    Cell_Val = sheet.cell(RowNo,Coapp_DLNumber).value
    sRequest = SetXML(sRequest, "a:DriverLicenseNumber", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Coapp_DLState).value
    sRequest = SetXML(sRequest, "a:DriverLicenseState", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Coapp_HomePhone).value
    sRequest = SetXML(sRequest, "a:HomePhone", Cell_Val.encode('ascii','ignore'), 1)

    # Co-App Current address
    Cell_Val = sheet.cell(RowNo,Coapp_AddressLine1).value
    sRequest = SetXML(sRequest, "a:AddressLine1", Cell_Val.encode('ascii','ignore'), 2)

    Cell_Val = sheet.cell(RowNo,Coapp_AddressLine2).value
    #if len(Cell_Val) == 0:
     #   sRequest = replacenth(sRequest, "<a:AddressLine2></a:AddressLine2>", """<a:AddressLine2 i:nil="true" />""", 3)
#    else:
 #       sRequest = SetXML(sRequest, "a:AddressLine2",Cell_Val.encode('ascii','ignore'), 2)
    sRequest = SetXML(sRequest, "a:AddressLine2",Cell_Val.encode('ascii','ignore'), 2)

    Cell_Val = sheet.cell(RowNo,Coapp_City).value
    sRequest = SetXML(sRequest, "a:City",Cell_Val.encode('ascii','ignore'), 2)

    Cell_Val = sheet.cell(RowNo,Coapp_State).value
    sRequest = SetXML(sRequest, "a:State",Cell_Val.encode('ascii','ignore'), 2)

    Cell_Val = sheet.cell(RowNo,Coapp_ZipCode).value
    sRequest = SetXML(sRequest, "a:ZipCode", Cell_Val.encode('ascii','ignore'), 2)

    Cell_Val = sheet.cell(RowNo,Coapp_HousingStatus).value
    sRequest = SetXML(sRequest, "a:HousingStatus", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Coapp_MortgageOrRent).value
    sRequest = SetXML(sRequest, "a:MortgageOrRent", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Coapp_TotalMonthsAtAddress).value
    sRequest = SetXML(sRequest, "a:TotalMonthsAtAddress", Cell_Val.encode('ascii','ignore'), 1)

    # Co-App Prev address
    Cell_Val = sheet.cell(RowNo,Coapp_Prev_AddressLine1).value
    sRequest = SetXML(sRequest, "a:AddressLine1", Cell_Val.encode('ascii','ignore'), 3)    
    
    Cell_Val = sheet.cell(RowNo,Coapp_Prev_AddressLine2).value
    sRequest = SetXML(sRequest, "a:AddressLine2", Cell_Val.encode('ascii','ignore'), 3)
    
    Cell_Val = sheet.cell(RowNo,Coapp_Prev_City).value
    sRequest = SetXML(sRequest, "a:City", Cell_Val.encode('ascii','ignore'), 3)

    Cell_Val = sheet.cell(RowNo,Coapp_Prev_State).value
    sRequest = SetXML(sRequest, "a:State", Cell_Val.encode('ascii','ignore'), 3)

    Cell_Val = sheet.cell(RowNo,Coapp_Prev_ZipCode).value
    sRequest = SetXML(sRequest, "a:ZipCode", Cell_Val.encode('ascii','ignore'), 3)

    # Co-App Current Emp
    Cell_Val = sheet.cell(RowNo,Coapp_EmploymentStatus).value
    sRequest = SetXML(sRequest, "a:EmploymentStatus", Cell_Val.encode('ascii','ignore'), 2)

    Cell_Val = sheet.cell(RowNo,Coapp_EmployedBy).value
    sRequest = SetXML(sRequest, "a:EmployedBy", Cell_Val.encode('ascii','ignore'), 2)

    Cell_Val = sheet.cell(RowNo,Coapp_TotalMonthsEmployed).value
    sRequest = SetXML(sRequest, "a:TotalMonthsEmployed", Cell_Val.encode('ascii','ignore'), 2)

    Cell_Val = sheet.cell(RowNo,Coapp_BusinessPhone).value
    sRequest = SetXML(sRequest, "a:BusinessPhone", Cell_Val.encode('ascii','ignore'), 2)

    Cell_Val = sheet.cell(RowNo,Coapp_MonthlyIncome).value
    sRequest = SetXML(sRequest, "a:MonthlyIncome", Cell_Val.encode('ascii','ignore'), 1)

    # Co-App Prev Emp
    Cell_Val = sheet.cell(RowNo,Coapp_Prev_EmploymentStatus).value
    sRequest = SetXML(sRequest, "a:EmploymentStatus", Cell_Val.encode('ascii','ignore'), 3)

    Cell_Val = sheet.cell(RowNo,Coapp_Prev_EmployedBy).value
    sRequest = SetXML(sRequest, "a:EmployedBy", Cell_Val.encode('ascii','ignore'), 3)

    Cell_Val = sheet.cell(RowNo,Coapp_Prev_TotalMonthsEmployed).value
    sRequest = SetXML(sRequest, "a:TotalMonthsEmployed", Cell_Val.encode('ascii','ignore'), 3)

    Cell_Val = sheet.cell(RowNo,Coapp_Prev_BusinessPhone).value
    sRequest = SetXML(sRequest, "a:BusinessPhone", Cell_Val.encode('ascii','ignore'), 3)

    # Co-App Others

    Cell_Val = sheet.cell(RowNo,Coapp_OtherMonthlyIncome).value
    sRequest = SetXML(sRequest, "a:OtherMonthlyIncome", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Coapp_OtherIncomeSource).value
    sRequest = SetXML(sRequest, "a:OtherIncomeSource", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Relationship).value
    sRequest = SetXML(sRequest, "a:Relationship", Cell_Val.encode('ascii','ignore'), 0)
   
    # *************************************************************************************************
    # Vehicle info
    Cell_Val = sheet.cell(RowNo,VehicleCondition).value
    sRequest = SetXML(sRequest, "a:VehicleCondition", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,Year).value
    sRequest = SetXML(sRequest, "a:Year", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,Make).value
    sRequest = SetXML(sRequest, "a:Make", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,Model).value
    sRequest = SetXML(sRequest, "a:Model", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,Trim).value
    sRequest = SetXML(sRequest, "a:Trim", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,ChromeStyleId).value
    sRequest = SetXML(sRequest, "a:ChromeStyleId", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,VIN).value
    sRequest = SetXML(sRequest, "a:VIN", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,StockNumber).value
    #print Cell_Val
    sRequest = SetXML(sRequest, "a:StockNumber", Cell_Val.encode('ascii','ignore'), 0)
    
    Cell_Val = sheet.cell(RowNo,CertifiedUsed).value
    sRequest = SetXML(sRequest, "a:CertifiedUsed", Cell_Val.encode('ascii','ignore'), 0)

    # *************************************************************************************************
    # Product Info
    Cell_Val = sheet.cell(RowNo,CashSellingPrice).value
    sRequest = SetXML(sRequest, "a:CashSellingPrice", Cell_Val.encode('ascii','ignore'), 0)
    
    Cell_Val = sheet.cell(RowNo,SalesTax).value
    sRequest = SetXML(sRequest, "a:SalesTax", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,Title).value
    sRequest = SetXML(sRequest, "a:Title", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,CashDown).value
    sRequest = SetXML(sRequest, "a:CashDown", Cell_Val.encode('ascii','ignore'), 0)
    
    Cell_Val = sheet.cell(RowNo,Rebate).value
    sRequest = SetXML(sRequest, "a:Rebate", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,CreditLifeIns).value
    sRequest = SetXML(sRequest, "a:CreditLifeIns", Cell_Val.encode('ascii','ignore'), 0)
    
    Cell_Val = sheet.cell(RowNo,Term).value
    sRequest = SetXML(sRequest, "a:Term", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,AcquisitionFees).value
    sRequest = SetXML(sRequest, "a:AcquisitionFees", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,InvoiceAmount).value
    sRequest = SetXML(sRequest, "a:InvoiceAmount", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,Warranty).value
    sRequest = SetXML(sRequest, "a:Warranty", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,MSRP).value
    sRequest = SetXML(sRequest, "a:MSRP", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,EstimatedBalloonAmount).value
    sRequest = SetXML(sRequest, "a:EstimatedBalloonAmount", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,EstimatedPayment).value
    sRequest = SetXML(sRequest, "a:EstimatedPayment", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,UsedCarBook).value
    #if len(Cell_Val) == 0:
     #   print "in if for used car book"
      #  sRequest = sRequest.replace("<a:UsedCarBook></a:UsedCarBook>","""<a:UsedCarBook  i:nil="true" />""",1)
#    else:
#        sRequest = SetXML(sRequest, "a:UsedCarBook", Cell_Val.encode('ascii','ignore'), 0)
    sRequest = SetXML(sRequest, "a:UsedCarBook", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,Mileage).value
    sRequest = SetXML(sRequest, "a:Mileage", Cell_Val.encode('ascii','ignore'), 0)
    
    Cell_Val = sheet.cell(RowNo,UsedCarValue).value
   # if len(Cell_Val) == 0:
    #    sRequest = sRequest.replace("<a:UsedCarValue></a:UsedCarValue>","""<a:UsedCarValue i:nil="true" />""",1)
#    else:
#        sRequest = SetXML(sRequest, "a:UsedCarValue", Cell_Val.encode('ascii','ignore'), 0)
    sRequest = SetXML(sRequest, "a:UsedCarValue", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,OtherFees).value
    sRequest = SetXML(sRequest, "a:OtherFees", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,WholesaleBookSource).value
    #if len(Cell_Val) == 0:
     #   sRequest = sRequest.replace("<a:WholesaleBookSource></a:WholesaleBookSource>","""<a:WholesaleBookSource i:nil="true" />""",1)
#    else:
#        sRequest = SetXML(sRequest, "a:WholesaleBookSource", Cell_Val.encode('ascii','ignore'), 0)
    sRequest = SetXML(sRequest, "a:WholesaleBookSource", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,WholesaleCondition).value
    #if len(Cell_Val) == 0:
     #   sRequest = sRequest.replace("<a:WholesaleCondition></a:WholesaleCondition>","""<a:WholesaleCondition i:nil="true" />""",1)
#    else:
#        sRequest = SetXML(sRequest, "a:WholesaleCondition", Cell_Val.encode('ascii','ignore'), 0)
    sRequest = SetXML(sRequest, "a:WholesaleCondition", Cell_Val.encode('ascii','ignore'), 0)
    
    Cell_Val = sheet.cell(RowNo,WholesaleValueType).value
    #if len(Cell_Val) == 0:
     #   sRequest = sRequest.replace("<a:WholesaleValueType></a:WholesaleValueType>","""<a:WholesaleValueType i:nil="true" />""",1)
#    else:
#        sRequest = SetXML(sRequest, "a:WholesaleValueType", Cell_Val.encode('ascii','ignore'), 0)
    sRequest = SetXML(sRequest, "a:WholesaleValueType", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,WholesaleValue).value
 #   if len(Cell_Val) == 0:
  #      sRequest = sRequest.replace("<a:WholesaleValue></a:WholesaleValue>","""<a:WholesaleValue i:nil="true" />""",1)
   # else:
#        sRequest = SetXML(sRequest, "a:WholesaleValue", Cell_Val.encode('ascii','ignore'), 0)
    sRequest = SetXML(sRequest, "a:WholesaleValue", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,NetTrade).value
    sRequest = SetXML(sRequest, "a:NetTrade", Cell_Val.encode('ascii','ignore'), 0)


    # trade in vehicle section
    Cell_Val = sheet.cell(RowNo,Trade_Year).value
    sRequest = SetXML(sRequest, "a:Year", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Trade_Make).value
    sRequest = SetXML(sRequest, "a:Make", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Trade_Model).value
    sRequest = SetXML(sRequest, "a:Model", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Trade_Trim).value
    sRequest = SetXML(sRequest, "a:Trim", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,Trade_ChromeStyleId).value
    sRequest = SetXML(sRequest, "a:ChromeStyleId", Cell_Val.encode('ascii','ignore'), 1)

    Cell_Val = sheet.cell(RowNo,LienHolder).value
    sRequest = SetXML(sRequest, "a:LienHolder", Cell_Val.encode('ascii','ignore'), 0)
    
    Cell_Val = sheet.cell(RowNo,MonthlyPayment).value
    sRequest = SetXML(sRequest, "a:MonthlyPayment", Cell_Val.encode('ascii','ignore'), 0)
    
    Cell_Val = sheet.cell(RowNo,LeadComments).value
    sRequest = SetXML(sRequest, "LeadComments", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,PrequalificationReferenceNumber).value
    sRequest = SetXML(sRequest, "PrequalificationReferenceNumber", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,LeadReferenceNumber).value
    sRequest = SetXML(sRequest, "LeadReferenceNumber", Cell_Val.encode('ascii','ignore'), 0)

    #print sRequest
    
    ReqFilePath = str(FolderPath) + '\\Request\\'+ WebService +'Request'+ str(RowNo) +'.xml'
    f = open(ReqFilePath, 'w')
    f.write(sRequest)
    f.close()
    
    return sRequest                     # return request string

#----------------------------------------------------------
# Description:      Function to create request XML string for FD - PreQualification web service for FD 1.1
# Input Parameters: FileName - Path of excel file
#                   RowNo - Current row in excel file
#                   Schemafile - File path of schema file
# Return Values:    request XML string
# Author:           Manisha Gadekar
#----------------------------------------------------------
def Create_FD_Prequal_Request_1_1_Phase_I(WebService, FileName, RowNo, HeaderList, Schemafile, FolderPath):
    # open excel file
    book = Open_Excel(FileName)
    sheet = book.sheet_by_index(0)
    
    # get column numbers from excel
    PartnerId = Get_ColumnNumber(HeaderList, "PartnerId")   # Calling a function to get column no of "PartnerDealerId" column
    PartnerDealerId = Get_ColumnNumber(HeaderList, "PartnerDealerIds")
    FirstName = Get_ColumnNumber(HeaderList, "FirstName")
    MiddleInitial = Get_ColumnNumber(HeaderList, "MiddleInitial")
    LastName = Get_ColumnNumber(HeaderList, "LastName")
    SSN = Get_ColumnNumber(HeaderList, "SSN")
    AddressLine1 = Get_ColumnNumber(HeaderList, "AddressLine1")
    AddressLine2 = Get_ColumnNumber(HeaderList, "AddressLine2")
    City = Get_ColumnNumber(HeaderList, "City")
    State = Get_ColumnNumber(HeaderList, "State")
    ZipCode = Get_ColumnNumber(HeaderList, "ZipCode")
    TotalMonthsEmployed = Get_ColumnNumber(HeaderList, "TotalMonthsEmployed")
    MonthlyIncome = Get_ColumnNumber(HeaderList, "MonthlyIncome")
    ConsentIndicator = Get_ColumnNumber(HeaderList, "ConsentIndicator")
    
    # Open schema file
    try:
        f = open(Schemafile)
    except IOError, e:
        print e
        print "Problem while opening Schema file"
        raise ErrorOccured("Problem in opening file: " + Schemafile)        # Problem opening file
    except:
        raise ErrorOccured("Unknown error occurred.")                        # handle any other type of error

    # everything is fine
    sRequest = f.read()         # get contenets of schema file in to string
    f.close()
    
    # Create request string
    # Calling function to add values to nodes in given schema file
    Cell_Val = sheet.cell(RowNo,PartnerId).value
    sRequest = SetXML(sRequest, "PartnerId", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,PartnerDealerId).value
    sRequest = SetXML(sRequest, "a:string",Cell_Val.encode('ascii','ignore') , 0)

    Cell_Val = sheet.cell(RowNo,FirstName).value
    sRequest = SetXML(sRequest, "a:FirstName", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,MiddleInitial).value
    sRequest = SetXML(sRequest, "a:MiddleInitial", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,LastName).value
    sRequest = SetXML(sRequest, "a:LastName", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,SSN).value
    sRequest = SetXML(sRequest, "a:SSN", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,AddressLine1).value
    sRequest = SetXML(sRequest, "a:AddressLine1", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,AddressLine2).value
    sRequest = SetXML(sRequest, "a:AddressLine2",Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,City).value
    sRequest = SetXML(sRequest, "a:City",Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,State).value
    sRequest = SetXML(sRequest, "a:State",Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,ZipCode).value
    sRequest = SetXML(sRequest, "a:ZipCode", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,TotalMonthsEmployed).value
    sRequest = SetXML(sRequest, "a:TotalMonthsEmployed", Cell_Val.encode('ascii','ignore'), 0)
    
    Cell_Val = sheet.cell(RowNo,MonthlyIncome).value 
    sRequest = SetXML(sRequest, "a:MonthlyIncome", Cell_Val.encode('ascii','ignore'), 0)

    Cell_Val = sheet.cell(RowNo,ConsentIndicator).value
    sRequest = SetXML(sRequest, "ConsentIndicator", Cell_Val.encode('ascii','ignore'), 0)
    
   #print sRequest
    
    ReqFilePath = str(FolderPath) + '\\Request\\'+ WebService +'Request'+ str(RowNo) +'.xml'
    f = open(ReqFilePath, 'w')
    f.write(sRequest)
    f.close()
    
    return sRequest                     # return request string

###############################################################################################
# Utilities to Read\Write to Text File using Standard Python Lib
#
################################################################################################
#def Deleate_Existing_Response_File(FilePath)

def GetFileCreationOrUpdateTime(fileRefWithFullPath):
   # st = os.stat(fileRefWithFullPath)
    fileInfo = []
    (mode, ino, dev, nlink, uid, gid, size, atime, mtime, ctime) = os.stat(fileRefWithFullPath)
    #timeInfo = time.asctime(time.localtime(st[ST_MTIME]))
    timeInfo = time.ctime(os.path.getctime(fileRefWithFullPath))
    print "Creation Time %s" % timeInfo
    fileInfo.append(timeInfo)	
    timeInfo = time.ctime(os.path.getmtime(fileRefWithFullPath))
    print "Modified Time %s" % timeInfo
    fileInfo.append(timeInfo)
    return fileInfo
	
def RenameAFile(dirpath, FileOriginalName,FileModifiedName):
    shutil.move(dirpath+FileOriginalName, dirpath+FileModifiedName)

def Delete_Existing_Response_File(RootPath, FileName):
 
        PATH = str(RootPath) + '\\' + 'Webservice\\Response\\' + str(FileName) 
        if path.isfile(PATH):
           try:
              os.remove(PATH)
           except OSError:
              pass
        return PATH 

def Store_To_File(FilePath, DataToAppend):
    with open(FilePath, 'a') as file:		
        DataToAppend += "\n"
        file.write(DataToAppend)
		
def Write_String_To_File(FilePath,StringToBeWritten):

    try:
    # This will create a new file or **overwrite an existing file**.
        f = open(FilePath, "w")
        try:
            #f.write(StringToBeWritten) # Write a string to a file
            f.writelines(StringToBeWritten) # Write a sequence of strings to a file
        finally:
            f.close()
    except IOError:
        pass
        
def Read_From_File_To_String(FilePath):
    # Read mode opens a file for reading only.
    stringReadFromFile = ""
    try: 
        
        f = open(FilePath, "r")
        try:
    # Read the entire contents of a file at once.
            stringReadFromFile = f.read() 
        finally:
            f.close()
    except IOError as e:
	    print str(e)
        #raise ErrorOccured("Invalid XML. Failed to read file into String " + str(e))
        #pass        
    sRequest = str(stringReadFromFile)
    return sRequest  

def Update_XML_Node_Value(XMLString,NodeRef,Value,Occurance):
    
    # parse xml string using xml.dom.minidom
    try:
        #dom = parse(XMLString)                             # XMLString : Either String of XML or XML File Ref
        dom = parseString(XMLString) 
    except:
        raise ErrorOccured("Invalid XML. Failed to parse XML input string: " + XMLString)
    # everything is fine
    #NodeName = dom.createTextNode(NodeRef.strip()) 
    NodeName = NodeRef	
    try:
        
#node = dom.getElementsByTagName(NodeName)[Occurance].childNodes[0]    # locating the node of interest
        node = dom.getElementsByTagName(NodeRef)[0].childNodes[0]    # locating the node of interest
    except:
        raise ErrorOccured("Node not found in xml: "+ NodeName) # not able to locate node
    
    try:
        node.nodeValue = Value
        modified_str = dom.toxml() 
        #txt = dom.createTextNode(Value.strip())         # creates text node
    except:
        raise ErrorOccured("Unable to add node value: ",Value," to node: " + node.tagName)     # problem creating text node
            
   # node.nodeValue = txt                                # Add text to given node. Results in <NodeName>Nodevalue</NodeName>
    #modified_str = dom.toxml()                          # modified xml string
    return modified_str                                 # return modified string

#----------------------------------------------------------
# Description:      Function to write string to console
# Input Parameters: 
# Author:           Sanjay Dubey
# Date:             25 Oct 13
#----------------------------------------------------------

def ConvertCSVToExcel(*listOfFiles):
    files = []
    files = listOfFiles
    for i in files:
        f=open(i, 'rb')
        g = csv.reader ((f), delimiter="~")
        wbk= xlwt.Workbook()
        sheet = wbk.add_sheet("Parsed Data")
    
	for rowi, row in enumerate(g):
            for coli, value in enumerate(row):
                sheet.write(rowi,coli,value)
	i = i[:-int(4)]
    wbk.save(i + '.xls')
	
def	OpenCSVFileInExcel(FileName):	
	Popen(FileName, shell=True)

	# def GetDataFromExcelFile(FileName, RowNo, HeaderList,FolderPath):    
    # book = Open_Excel(FileName)
    # sheet = book.sheet_by_index(0)    
	# PartnerId = Get_ColumnNumber(HeaderList, "PartnerId")   # Calling a function to get column no of "PartnerDealerId" column
    # PartnerDealerId = Get_ColumnNumber(HeaderList, "PartnerDealerIds")
    # FirstName = Get_ColumnNumber(HeaderList, "FirstName")
    # MiddleInitial = Get_ColumnNumber(HeaderList, "MiddleInitial")
    # LastName = Get_ColumnNumber(HeaderList, "LastName")
    # SSN = Get_ColumnNumber(HeaderList, "SSN")
    # AddressLine1 = Get_ColumnNumber(HeaderList, "AddressLine1")
    # AddressLine2 = Get_ColumnNumber(HeaderList, "AddressLine2")
    # City = Get_ColumnNumber(HeaderList, "City")
    # State = Get_ColumnNumber(HeaderList, "State")
    # ZipCode = Get_ColumnNumber(HeaderList, "ZipCode")
    # TotalMonthsEmployed = Get_ColumnNumber(HeaderList, "TotalMonthsEmployed")
    # MonthlyIncome = Get_ColumnNumber(HeaderList, "MonthlyIncome")
    # ConsentIndicator = Get_ColumnNumber(HeaderList, "ConsentIndicator")

def RetainExcelDocHeader(ExcelFileRef,SheetRef):
    Header = []
    FileName = trimFromRight(ExcelFileRef,4)
    book = open_workbook(FileName + '.xls')
    ws = book.sheet_by_name(SheetRef)
    for col_index in range(ws.ncols):
        x = ws.cell_value(0, col_index)	
        Header.append(x) 
    return Header
		
def	RetainExcelDocsColumnNames(ExcelFileRef,ColNum,CellRow,CellCol):
    files = [ExcelFileRef]
    for i in files:
        f=open(i, 'rb')
        g = csv.reader ((f), delimiter="~")
        wbk= xlwt.Workbook()
        sheet = wbk.add_sheet("Parsed Data")

    for rowi, row in enumerate(g):
       for coli, value in enumerate(row):
            sheet.write(rowi,coli,value)
    FileName = trimFromRight(ExcelFileRef,4)
    wbk.save(FileName + '.xls'	)
    book = open_workbook(FileName + '.xls')
    ws = book.sheet_by_name('Parsed Data')
    cell_value = ws.cell_value(0, 1)		#ws.cell_value(Row,Col)
    #columnName = colname(ColNum)
    # return columnName
    return cell_value
	
def	RetainRowValuesForACol(ExcelFileRef,ColName):
    Header = []	
    RowValues = []	
    Header = RetainExcelDocHeader(ExcelFileRef,"Parsed Data")
    index = 0
    col_index = 0
    row_index = 0
    for header in Header:
        if header == ColName:
           col_index = index 
        index = index + 1
    
	FileName = trimFromRight(ExcelFileRef,4)
    book = open_workbook(FileName + '.xls')
    ws = book.sheet_by_name("Parsed Data")
    #a = range(1,int(ws.nrows))
    #for row in range(1,int(ws.nrows)):
   # for row in range(int(ws.nrows)):
    Nrow = ws.nrows
    print  Nrow
    for row in range(Nrow):
        #print i
        x = ws.cell_value(row,col_index)
        RowValues.append(x)
    # for row in range(1, 5):
        # print row
    # for row in range(Nrow):		
        # print row
        #x = ws.cell_value(row,col_index)	
        #RowValues.append(x) 
    return RowValues
   
def Write_To_Console(Str):
    Str = '\n' + Str + '\n'
    sys.__stdout__.write(Str)

def	Create_Excel_File(Str):
	book = Workbook()
	sheet1 = book.add_sheet('Sheet1')
	book.save(Str)
	# book.save(TemporaryFile())

if __name__ =="__main__": 
    file = "C:\\Project\\Dev\\SL-Org\\testcases\\POC\\POC_Include.txt" 
    GetFileCreationOrUpdateTime(file)
	# #RetainRowValuesForACol(5)
	# '''
    # for x in ['1','2','3','4','5']:
        # #x = ws.cell_value(row,col_index)	
        # #RowValues.append(x) 
        # print 'test:',x
	# '''

def Send_Get_Request(url, Username, Password):
    request = urllib2.Request(url)
	
    base64string = base64.encodestring('%s:%s' % (Username,Password))[:-1]
    authheader =  "Basic %s" % base64string
    request.add_header("Authorization", authheader)
    
    flag = 0
    try:
        response = urllib2.urlopen(request)                         # sending request
    except urllib2.URLError, e:                                     # error code
        if e.code == 404:
            StepResult = "Executed"    
            print e.read()    
            raise ErrorOccured("URL Error occurred while sending request: " + e.read())
        elif e.code == 401:
            StepResult = "Executed"            
            print e.read()            
            raise ErrorOccured("Unauthorized Access: " + e.read())
        else:
            StepResult = "Executed"            
            StatusCode = e.code
            SResponseText = e.read()            
            print SResponseText            
            flag = 1
    except HTTPError, e:
        StepResult = "Executed"        
        print e.read()        
        raise ErrorOccured("You have received HTTP error: " +  e.read())
    except:
        raise ErrorOccured("Unknown error occurred." )   

    # Everything is fine, you got valid response
    if flag == 0:
        StatusCode = response.getcode()                             # get status code
        SResponseText = response.read()                             # get response XML text         
        StepResult = "Executed"
        print 'You have received valid response'
        print 'Status Code:',StatusCode
        print 'Response Text: ',SResponseText
        print StepResult        
    
    return StatusCode,SResponseText                             # return response code and response text	
	
def Send_Get_Request_PDF(url, Username, Password):
    request = urllib2.Request(url)
	
    base64string = base64.encodestring('%s:%s' % (Username,Password))[:-1]
    authheader =  "Basic %s" % base64string
    request.add_header("Authorization", authheader)
    
    flag = 0
    try:
        response = urllib2.urlopen(request)                         # sending request
    except urllib2.URLError, e:                                     # error code
        if e.code == 404:
            StepResult = "Executed"    
            print e.read()    
            raise ErrorOccured("URL Error occurred while sending request: " + e.read())
        elif e.code == 401:
            StepResult = "Executed"            
            print e.read()            
            raise ErrorOccured("Unauthorized Access: " + e.read())
        else:
            StepResult = "Executed"            
            StatusCode = e.code
            SResponseText = e.read()            
            print SResponseText            
            flag = 1
    except HTTPError, e:
        StepResult = "Executed"        
        print e.read()        
        raise ErrorOccured("You have received HTTP error: " +  e.read())
    except:
        raise ErrorOccured("Unknown error occurred." )   

    # Everything is fine, you got valid response
    if flag == 0:
        StatusCode = response.getcode()                             # get status code
        SResponseText = response.read()                             # get response XML text   
        SResponseHeaders = response.headers  
        StepResult = "Executed"
        print 'You have received valid response'
        print 'Status Code:',StatusCode
        print 'Response Text: ',SResponseText
        print StepResult        
    
    return StatusCode,SResponseText,SResponseHeaders				# return response code and response text	

def Send_TPP_Request(url,sRequest,Username,Password,FolderPath):
    print url
    print Username
    print Password

    StepResult = "Not Executed"                                             
    ErrorText = "Nothing"
    StepResult = "Not Executed"                                             # initialising step result
    
    print url
    
    headers = {'Content-Type': 'application/x-www-form-urlencoded',
                'Authorization': 'Basic YmFycG9zdDo4Z2d2cWh2cA=='}                                                       #'Cookie': 'SMCHALLENGE=YES'

    sRequest = "xml=" + urllib.quote_plus(sRequest)	
    #sRequest = urllib.urlencode(sRequest)
    
    #####################################################
    # Common code for sending request
    request = urllib2.Request(url,sRequest,headers)
   # base64string = base64.encodestring('%s:%s' % (Username, Password))[:-1]
   # authheader =  "Basic %s" % base64string
   # request.add_header("Authorization", authheader)
    
    flag = 0
    try:
        response = urllib2.urlopen(request)                         # sending request
    except urllib2.URLError, e:                                     # error code
        if e.code == 404:
            StepResult = "Executed"
    
            print e.read()
    
            raise ErrorOccured("URL Error occurred while sending request: " + e.read())
        elif e.code == 401:
            StepResult = "Executed"
            
            print e.read()
            
            raise ErrorOccured("Unauthorized Access: " + e.read())
        else:
            StepResult = "Executed"
            
            StatusCode = e.code
            SResponseText = e.read()
            
            print SResponseText            
            flag = 1
    except HTTPError, e:
        StepResult = "Executed"
        
        print e.read()
        
        raise ErrorOccured("You have received HTTP error: " +  e.read())
    except:
        raise ErrorOccured("Unknown error occurred." )

   

    # Everything is fine, you got valid response
    if flag == 0:
        StatusCode = response.getcode()                             # get status code
        SResponseText = response.read()                             # get response XML text

        
        
        StepResult = "Executed"
        print 'You have received valid response'
        print 'Status Code:',StatusCode
        print 'Response Text: ',SResponseText
        print StepResult
        
    
    return StatusCode,ErrorText,SResponseText                             # return response code and response text

def extract_between(text, sub1, sub2, nth=1):
    # prevent sub2 from being ignored if it's not there
    if sub2 not in text.split(sub1, nth)[-1]:
        return None
    return text.split(sub1, nth)[-1].split(sub2, nth)[0]

def CreateRandomNumberInRange(limit,start,end):
    str1=""
    for x in range(int(limit)):
      str2=random.randint(int(start),int(end))
      str1=str1+str(str2)
    return str1

def SendJsonRequest(url, str_req, Username, Password):
    # getting warnings like insecureplatformwarning. So adding below line.
    requests.packages.urllib3.disable_warnings()
    
    print url
    print str_req
    print Username
    print Password
    headers = {
               'Content-Type': 'application/json',
               'Cookie': 'SMCHALLENGE=YES',
    }

    try:
        json_response = requests.post(url,data=str_req,headers=headers,auth=(Username,Password))       # put or post method
    except Exception as e:
        ErrorText = str(e)
        print ErrorText
        raise ErrorOccured("Unknown error occurred. Please troubleshoot." + ErrorText)
        
    StatusCode = json_response.status_code
    SResponseText = json_response.text

    print StatusCode
    print SResponseText
    
    return StatusCode,SResponseText 

def ConvertJsonStringToDict(response_str):
    ret_dict=json.loads(response_str)
    return  ret_dict
    
def ConvertDictToJsonString(InputDict):
    ret_str=json.dumps(InputDict)
    return  ret_str

############### COL ################

def apply_customized_operation(value, operation):
    if operation == 'make_lower':
        return value.lower()

    elif operation == 'make_upper':
        return value.upper()

    elif operation == 'housing_status':
        value = value.upper()
        if value == 'OWNOUTRIGHT':
            return value[0:3] + ' ' + value[3:]
        elif value == 'MORTGAGE':
            return 'HOMEOWNER'

    elif operation == 'get_years':
        return int(value) / 12

    elif operation == 'get_months':
        return int(value) % 12

    elif operation == 'dob':
        return value[5:7] + value[8:10] + value[0:4]

    elif operation == 'street_number':
        # Sending some default value
        return '556'

    elif operation == 'street_name':
        return 'SE 01'

    elif operation == 'street_type':
        return 'ST'

    elif operation == 'finance_method':
        if value == 'Lease':
            return 'lease'
        return 'buy'

    elif operation == 'trade_in':
        if value == []:
            return 'TradeNo'
        else:
            return 'TradeYes'

    elif operation == 'vehicle_in':
        if value == []:
            return 'VehicleNo'
        else:
            return 'VehicleYes'

    elif operation == 'purchase_type':
        return value.lower() + 'Veh'

    elif operation == 'coapp_in':
        if value == []:
            return 'individually'
        return 'joint'

    elif operation == 'relationship':
        value = value.lower()
        if value not in ['spouse', 'relative']:
            return 'other'
        return value

    return value


def get_current_element_attribute(element, attrrib_name):
    return element.get(attrrib_name, False)


def create_string_from_dict(input_dict=None, str_dict=None):
    """Returns the dict into following string format
       "Key1: value1
       key2: value2"
    """
    if str_dict is None:
        str_dict = ''

    if input_dict is not None:

        for key, value in input_dict.items():
            str_dict = str_dict + key + ': ' + value + '\n'

    return str_dict



def create_ordered_dict():
    return collections.OrderedDict()


from urlparse import urlparse, parse_qs

def get_query_patrameter(urlstring, parameter):
    """Returns the parameter values from the provided urlstring.
    """
    #o = urlparse("LeadDetail.aspx?appId=17B9780617&amp;sid=WSP")

    o = urlparse(urlstring)
    q = parse_qs(o.query)
    return q[parameter][0]
