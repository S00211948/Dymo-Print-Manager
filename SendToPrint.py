import win32com.client
from json import load as jsload
from json import dumps as jsdumps
from os import path as ospath
from tkinter.filedialog import askopenfilename as tkopenfile
import sys

class DymoPrintService():
    # Create DYMO COM object
    label = None
    labelText = None
    template = None
    templateMap = None
    
    def __init__(self):
        self.label = win32com.client.Dispatch("Dymo.DymoAddIn")
        self.labelText = win32com.client.Dispatch("Dymo.DymoLabels")
        print("> Dymo Print Manager Initialized")

    def printLabelList(self,label_data):
        #for d in label_data:
            # Format data for printing
            #label_data = self.formatForPrinting(d)
            if len(label_data) > 0:
                for l in label_data:
                    if l != False:
                        print(l)
                        for mapping in self.templateMap: 
                            self.labelText.SetField(mapping['Label_Field'], l[mapping["Data_Field"]])
                        self.label.StartPrintJob()
                        self.label.Print(1, False)   # 1 copy, not asynchronously
                        self.label.EndPrintJob()

    def formatForPrinting(self,dataObject,isList=True):
        try:
            res_data=[]
            if isList:
                #Children Labels
                for i in range(1,7):
                    if isinstance(dataObject[f'Guest_{i}'],str) and dataObject[f'Guest_{i}'] != '':
                        res_data.append({"Employee_Name":f"{dataObject['Employee']} - {dataObject['ID']}","Visitor_Name":dataObject[f'Guest_{i}'],"Visitor_Age":dataObject[f'Guest_{i}_Age']})
                # Partner Label
                #if isinstance(dataObject[f'Partner'],str) and dataObject[f'Partner'] != '':
                 #       res_data.append({"Employee_Name":f"{dataObject['Employee']}","Visitor_Name":dataObject[f'Partner'],"Tour_Number":dataObject['Tour']})
            else:
                res_data.append({"Employee_Name":f"{dataObject['Employee']}","Employee_Address":dataObject['Address']})

            return res_data
        except Exception as e:
            return e
        
    def printListWithMap(self,templateName,mapName,dataObject,isList=True):
        self.readTemplate(templateName,mapName)
        for d in dataObject:
            data = self.formatForPrinting(d,isList)
            self.printLabelList(data)
        
    def readTemplate(self,templateName,mapName):
        # Get the directory of the current file
        filepath = self.get_path()
        config_path = ospath.join(filepath,'config.json')
        with open(config_path,'r') as config:
            data = jsload(config)
        self.template = data[templateName]
        self.templateMap = data[mapName]
        if not self.label.Open(self.template):
            raise Exception("Could not open label template")

    def setTemplate(self):
        label_file = tkopenfile(
            title="Select a File",
            filetypes=[("Label Files", "*.label")]
        )

        # Get the directory of the current file
        filepath = self.get_path()
        config_path = ospath.join(filepath,'config.json')
        if not self.label.Open(label_file):
            raise Exception("Could not open label template")
        else:
            with open(config_path,'w') as config:
                config.write(jsdumps({"template":filepath}))
            self.template = label_file
    
    def get_path(self):
        if hasattr(sys, 'frozen'):  
            # When running as .exe
            return ospath.dirname(ospath.realpath(sys.executable))
        else:
            # When running as .py
            return ospath.dirname(ospath.realpath(__file__))