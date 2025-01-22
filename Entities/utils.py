import os
import re
from typing import Dict
import pandas as pd

class Expurgo:
    """Classe responsável por processar dados de expurgo a partir de planilhas Excel."""

    @property
    def peps(self) -> dict:
        """
        Retorna um dicionário que mapeia cada centro para a lista de PEPs correspondente.
        """
        return self.__peps
    
    
    @property
    def path(self) -> str:
        """
        Retorna o caminho do arquivo Excel utilizado para armazenar os dados de expurgo.
        """
        return self.__path
        
    def __init__(self, path) -> None:
        """
        Inicializa a classe Expurgo com o caminho do arquivo Excel e realiza a extração dos dados.
        :param path: Caminho completo do arquivo Excel.
        """
        self.__path = path
        self.__df = pd.read_excel(path)
        self.__extrair()
        
    def __extrair(self):
        """
        Extrai e organiza as informações de PEPs associadas a cada centro, armazenando-as em um dicionário interno.
        """
        centros = self.__df['Código da Obra'].unique().tolist()
        result = {}
        for centro in centros:
            lista_peps = self.__df[
                self.__df['Código da Obra'] == centro
            ]['PEP'].tolist()
            result[centro] = lista_peps
            
        self.__peps = result
        
    def __repr__(self) -> str:
        """
        Retorna uma representação em string do caminho do arquivo.
        """
        return self.path

class Paths:
    """Classe responsável por localizar arquivos de relatório e expurgos em um diretório específico."""

    @property
    def nome_empreendimentos(self) -> dict:
        """
        Retorna um dicionário que relaciona cada centro ao nome do empreendimento.
        """
        return self.__nome_empreendimentos

    @property
    def expurgo(self) -> Expurgo:
        """
        Retorna uma instância da classe Expurgo, criada a partir do arquivo de expurgos encontrado.
        """
        return Expurgo(self.__expurgo)
    
    @property
    def files_to_modificated(self) -> Dict[str,str]:
        """
        Retorna um dicionário que mapeia cada centro ao respectivo caminho de relatório.
        """
        return self.__files_to_modificated
    
    @property
    def lista_centros(self) -> list:
        """
        Retorna a lista de centros presentes nos arquivos localizados.
        """
        return list(self.files_to_modificated.keys())
    
    def __init__(self, path:str):
        """
        Inicializa a classe Paths com o caminho base para busca de arquivos e executa a varredura inicial.
        :param path: Caminho do diretório onde estão os arquivos.
        """
        if not os.path.exists(path):
            raise FileNotFoundError(f"O diretório '{path}' não foi encontrado.")
        self.path = path
        self.files_extracted_from_sap:dict = {}
        self.exec()
        
    def exec(self):
        """
        Percorre os diretórios em busca de planilhas Excel, identificando arquivos de expurgo e relatórios de custos.
        """
        result = {}
        nome_empreendimentos = {}
        
        for root, dirs, files in os.walk(self.path):
            for file in files:
                file_path = os.path.join(root, file)
                
                if file.endswith((".xlsx", ".xls", ".xlsm")):
                    if "Expurgos.xlsx" == file:
                        self.__expurgo = file_path
                        
                    if re.search("Relatório de Custos", file):
                        if (centro:=re.search(r'[A-z]{1}[0-9]{3}', file)):
                            temp_name = file.split('-')[-1]
                            while temp_name[0] == " ":
                                temp_name = temp_name[1:]
                            temp_name = temp_name.split('.')[0]
                            
                            result[centro.group()] = file_path
                            nome_empreendimentos[centro.group()] = temp_name
                            
        
        self.__files_to_modificated = result
        self.__nome_empreendimentos = nome_empreendimentos
                    

        
if __name__ == "__main__":
    paths = Paths(f"C:\\Users\\renan.oliveira\\PATRIMAR ENGENHARIA S A\\Janela da Engenharia Controle de Obras - Relatório de Custos")
    
    print(f"{paths.expurgo.peps=}")
    #print(paths.nome_empreendimentos.keys())
    #print(paths.lista_centros)
