from Entities.CJI3 import CJI3
from Entities.mod_planilhas import Planilhas
from Entities.utils import Paths
from Entities.dependencies.arguments import Arguments
from Entities.dependencies.config import Config
from datetime import datetime
import re
import os

class Execute:
    @staticmethod
    def start(with_sap:bool):
        paths = Paths(Config()['paths']['sharepoint'])
        
        if with_sap:
            CJI3(date=datetime.now()).gerar_relatorios_SAP(objeto=paths)
        
        result = {}
        p = os.path.join(os.getcwd(), 'Bases')
        for file in os.listdir(p):
            file_path = os.path.join(p, file)
            if os.path.isfile(file_path):
                if (c:=re.search(r'[A-z]{1}[0-9]{3}', file)):
                    result[c.group()] = file_path
        paths.files_extracted_from_sap = result
        
        #print(paths.files_extracted_from_sap)
            
        Planilhas.tratar_dados(paths)
    
    @staticmethod 
    def start_with_sap():
        Execute.start(with_sap=True)
        
    @staticmethod 
    def start_with_not_sap():
        Execute.start(with_sap=False)
        
if __name__ == "__main__":
    Arguments({
        'start_with_sap': Execute.start_with_sap,
        'start_with_not_sap': Execute.start_with_not_sap
    })