import xlwings as xw
from xlwings.main import Sheet
from xlwings.main import App
import pandas as pd
import os
import shutil
from Entities.dependencies.logs import Logs, traceback
from Entities.dependencies.functions import Functions, P
import multiprocessing
multiprocessing.freeze_support()

def exec(*, file_extrac_from_sap:str, file_to_modificate:str, expurgos:list|None):
    df = pd.read_excel(file_extrac_from_sap)
    
    df = df.dropna(subset=['Elemento PEP'])
        
    df = df[
        (df['Classe de custo'].astype(str).str.startswith('41')) &
        (~df['Elemento PEP'].str.lower().str.contains('POSO'.lower())) &
        (~df['Elemento PEP'].str.lower().str.contains('POCRCIAI'.lower())) &
        (df['Denomin.da conta de contrapartida'].str.lower().str.replace(' ', '') != "ESTOQUE DE TERRENOS".lower().replace(' ', '')) &
        (df['Denomin.da conta de contrapartida'].str.lower().str.replace(' ', '') != "T.  EST. TERRENOS".lower().replace(' ', '')) &
        (df['Denomin.da conta de contrapartida'].str.lower().str.replace(' ', '') != "T. ESTOQUE INICIAL".lower().replace(' ', ''))
    ]
    
    if expurgos:
        df = df[
            ~df['Elemento PEP'].isin(expurgos)
        ]
    
    app = xw.App(visible=False)
    with app.books.open(file_to_modificate) as wb:
        ws:Sheet = wb.sheets['BD_SAP']
        
        ws.api.AutoFilter.ShowAllData()
        
        ws.range('A2', ws.cells.last_cell).clear_contents()
        
        
        num_rows = len(df)
        formulas =[[f'=EOMONTH(F{x},-1)+1',
        f'=MID(M{x},6,4)',
        f'=MID(M{x},6,6)',
        f'=MID(M{x},6,8)',
        f'=MID(Q{x},1,2)'
        ] for x in range(2, num_rows+2)]
        
        ws.range(f'A2:E{num_rows+1}').formula = formulas
        ws.range('F2').value = df.values
        
        wb.save()
    #app.quit()   
    Functions.fechar_excel(file_to_modificate)        

    return True

class TratarDados:
    @staticmethod
    def start(queue:multiprocessing.JoinableQueue, file_to_modificate:str, file_extrac_from_sap:str, expurgos:list|None):
        for _ in range(5):
            try:
                return queue.put(exec(file_extrac_from_sap=file_extrac_from_sap, file_to_modificate=file_to_modificate, expurgos=expurgos))
            except Exception as error:
                Logs().register(status='Report', description=f"Erro ao tratar dados do arquivo '{file_to_modificate}'", exception=traceback.format_exc())
                if _ == 4:
                    return queue.put(False)
        
                    
if __name__ == "__main__":
    pass
