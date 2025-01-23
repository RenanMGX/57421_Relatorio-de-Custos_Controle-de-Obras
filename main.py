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
    def start():
        paths = Paths(Config()['paths']['sharepoint'])
        
        CJI3(date=datetime.now()).gerar_relatorios_SAP(objeto=paths)
        
        Planilhas.tratar_dados(paths)
        
if __name__ == "__main__":
    Arguments({
        'start': Execute.start
    })