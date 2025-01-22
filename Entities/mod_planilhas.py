import os
import multiprocessing
from typing import List, Dict
multiprocessing.freeze_support()
from Entities.dependencies.logs import Logs
from Entities.dependencies.functions import P
from Entities.tratar_dados import TratarDados
from Entities.utils import Paths

class Planilhas:
    @staticmethod
    def tratar_dados(paths:Paths):
        queues:List[multiprocessing.Queue] = []
        
        expurgos:dict = paths.expurgo.peps
        
        for centro in paths.lista_centros:
            if Planilhas.__tratar_errors(paths, centro):
                q = multiprocessing.Queue()
                multiprocessing.Process(target=TratarDados.start, args=(
                    q,
                    paths.files_to_modificated.get(centro),
                    paths.files_extracted_from_sap.get(centro),
                    expurgos.get(centro)
                )).start()
                
                queues.append(q)
                
            else:
                continue
            
        for queue in queues:
            print(queue.get())
        
    @staticmethod
    def __tratar_errors(paths:Paths, centro:str) -> bool:
        if (p:=paths.files_to_modificated.get(centro)):
            if not os.path.exists(p):
                Logs().register(status='Report', description=f"o arquivo '{p}' não foi encontrado!")
                return False
        else:
            Logs().register(status='Report', description=f"{centro=} não foi encontrado em .files_to_modificated")
            return False
        
        if (p:=paths.files_extracted_from_sap.get(centro)):
            if not os.path.exists(p):
                Logs().register(status='Report', description=f"o arquivo '{p}' não foi encontrado!")
                return False
        else:
            Logs().register(status='Report', description=f"{centro=} não foi encontrado em .files_extracted_from_sap")
            return False
        
        return True
        
        
        
                        
if __name__ == "__main__":
    pass

