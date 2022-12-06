
# Author: Nino Casupanan
# Cluster 3, BAS Team
import pandas
import os
import sys
from datetime import datetime
from argparse import ArgumentParser
import json
import jpype
import asposecells
from .components import getLogger

# This class is only package dependent no special complexity on this class.
class PyConverter:
    __class__ = "PyConverter"
    arguments = None
    parser = None
    fileName = None
    encoding = "utf-8-sig"
    logger = getLogger(__class__)
    
    def __init__(self):
        self.parser = ArgumentParser()
        self.arguments = None
        # Year, Date, Month - Hour Minute, Seconds
        self.fileName = datetime.now().strftime("%Y%d%m_%H_%M_%S")
        pass
    
    def cls(self):
        os.system('cls' if os.name=='nt' else 'clear')
        
    # Check arguments
    def parseArguments(self):
        
        # Validate arguments
        if len(sys.argv) <= 1:
            self.logger.critical("No argument passed. Please see -h for argument requirement.")
            exit() # raise Exception("No argument passed. Please see -h for argument requirement.")

        # Set expected arguments
        self.parser.add_argument("-t", "--type", help= "Choose the type of process. Possible values (e2j=Excel to JSON, j2e=JSON to Excel)")
        self.parser.add_argument("-e", "--excel", help = "Choose excel to convert.")
        self.parser.add_argument("-j", "--json", help = "Choose a file to convert.")
        self.parser.add_argument("-p", "--path", help = "Choose a path for the output.", default = os.path.dirname(__file__))
        self.parser.add_argument("-n", "--name", help = "Choose a name for the file.", default = self.fileName)
        self.parser.add_argument("-f", "--format", help= "Choose format of the json payload. (default, dataframe)", default="default")
        
        # For multiple parsing
        self.parser.add_argument("-jp", "--json-path", help = "Indicate the folder path of JSON to convert. Converted files will be in the same folder as the json files.")

        self.arguments = self.parser.parse_args()
        
        if self.arguments.json == None and self.arguments.json_path == None and self.arguments.type == None:
            raise Exception("Output path should not be empty..")
        
        if self.arguments.type != None and self.arguments.type == "e2j":
            self.convertExcel2Json()
        
        return self
    
    # Process
    def readAndConvert(self):
        if self.arguments.excel != None:
            return self
        
        # For multiple convertion of json to excel
        if self.arguments.json_path != None:
            directory = os.fsencode(self.arguments.json_path)
            dirFiles = os.listdir(directory)
            for file in dirFiles:
                filename = os.fsdecode(file)
                self.logger.info("\nConverting " + filename)
                
                if filename.endswith(".json"):
                    absFilePath = os.path.join(os.getcwd(), self.arguments.json_path, filename)
                    if os.path.isfile(absFilePath):
                        # set the name for multiple json convertion
                        self.fileName = filename.split(".")[0]
                        self.arguments.name = self.fileName
                        
                        # Force path to be the directory path.
                        if (self.arguments.path != None):
                            self.arguments.json_path = None
                        
                        # Set the json path
                        self.arguments.json = absFilePath
                        # Convertion
                        self.runConvertion()
        # For Single convertion of json to excel
        elif self.arguments.json != None:

            # Check if the given json path is a valid file or existing
            if os.path.isfile(self.arguments.json):
                try: 
                    # Run the convertion
                    self.runConvertion()
                except Exception as e: 
                    raise e
            else:
                raise Exception("File does not exist.")
        else:
            raise Exception("Please specify the file.")
        
    
    # Convertion of the json
    def runConvertion(self):
        dirPath = str(self.arguments.json_path if self.arguments.json_path != None else self.arguments.path)
        if self.arguments.name != None:
            filePath = os.path.join(dirPath, self.arguments.name + '.xlsx')
        else:
            filePath = os.path.join(dirPath, '%s-extraction.xlsx' % self.fileName)
            
        #
        if not os.path.exists(dirPath):
            os.makedirs(dirPath)
        
        # For generating excel with multiple sheet
        if self.arguments.format != "dataframe":
            writer = pandas.ExcelWriter(filePath, engine='xlsxwriter')
        
            # For bigger JSON size.
            writer.book.use_zip64()
            
            jsObj = json.load(open(self.arguments.json, encoding="utf8"))
        
            if type(jsObj) is dict:
                for sheetName in jsObj:
                    pandas.read_json(json.dumps(jsObj[sheetName]), encoding = self.encoding).to_excel(writer, index=False, sheet_name=sheetName)
            else:
                pandas.read_json(self.arguments.json, encoding = self.encoding).to_excel(writer, index=False)
                
            writer.close()
        # For generating excel with single sheet data
        else:
            jsonObj = json.load(open(self.arguments.json, encoding="utf8"))
            df = pandas.DataFrame(jsonObj)
            writer = pandas.ExcelWriter(filePath, engine='xlsxwriter')
        
            # For bigger JSON size.
            writer.book.use_zip64()
            
            # Convert the dataframe to an XlsxWriter Excel object.
            df.to_excel(writer, sheet_name='Sheet1', index=False)

            # Close the Pandas Excel writer and output the Excel file.
            writer.close()
        
        self.logger.info("Convertion completed. File: " + filePath)
        
    def convertExcel2Json(self):
        jpype.startJVM()
        from asposecells.api import Workbook, License
        
        self.logger.info("Converting to json " + self.arguments.excel)
        # load Excel file
        workbook = Workbook(self.arguments.excel)

        # Save Excel file as JSON
        workbook.save(self.arguments.excel.split(".")[0] + '.json')