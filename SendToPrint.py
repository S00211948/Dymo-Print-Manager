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
    
    def __init__(self):
        self.label = win32com.client.Dispatch("Dymo.DymoAddIn")
        self.labelText = win32com.client.Dispatch("Dymo.DymoLabels")
        self.readTemplate()
        # Load the template
        #if not self.label.Open(r"C:\Users\O89301\OneDrive - The Coca-Cola Company\Documents\DYMO Label\Labels\OpenDay-Basic.label"):
        if not self.label.Open(self.template):
            raise Exception("Could not open label template")

    def printLabelList(self,dataObj):
        for d in dataObj:
            # Format data for printing
            for l in self.formatForPrinting(d):
                if l != False:
                    print(l)  
                    self.labelText.SetField("Employee", l["Employee_Name"])
                    self.labelText.SetField("Visitor", l["Visitor_Name"])
                    self.labelText.SetField("Tour", l["Tour_Number"])
                    self.labelText.SetField("QRcode",l["ID"])
                    self.label.StartPrintJob()
                    self.label.Print(1, False)   # 1 copy, not asynchronously
                    self.label.EndPrintJob()

    def formatForPrinting(self,dataObject):
        try:
            res_data=[]
            for i in range(1,4):
                if isinstance(dataObject[f'Guest_{i}'],str) and dataObject[f'Guest_{i}'] != '':
                    res_data.append({"Employee_Name":f"{dataObject['Employee']}","Visitor_Name":dataObject[f'Guest_{i}'],"Tour_Number":dataObject['Tour'],"ID":dataObject['ID']})
            return res_data
        except Exception as e:
            return e
        
    def readTemplate(self):
        # Get the directory of the current file
        filepath = self.get_path()
        config_path = ospath.join(filepath,'config.json')
        with open(config_path,'r') as config:
            data = jsload(config)
        self.template = data['template']

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