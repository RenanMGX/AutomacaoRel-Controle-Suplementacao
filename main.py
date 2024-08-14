from Entities.start_report import StartReport
from Entities.file_manipulate import FileManipulate
from Entities.dependencies.functions import _print
from getpass import getuser
import os
import sys

class Execute:
    @property
    def download_path(self) -> str:
        path:str = os.path.join(os.getcwd(), "downloads")
        if not os.path.exists(path):
            os.makedirs(path)    
        return path
    
    @property
    def destiny_path(self) -> str:
        return f"C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\RPA - Documentos\\RPA - Dados\\RPA - Controle - Suplementação"
    
    @property
    def new_name(self) -> str:
        return "ControleObra_Suplementacao.json"
    
    def __init__(self) -> None:
        self.__sap_start_report:StartReport = StartReport('SAP_PRD')
        #self.__file_manipulate: FileManipulate = FileManipulate()
    
    def start(self) -> None:
        file_path:str = self.__sap_start_report.extrair_rel(download_path=self.download_path)
        
        self.__file_manipulate:FileManipulate = FileManipulate(file_path)
        
        self.__file_manipulate.excelToJson(remove_duplicate=True).moveTo(self.destiny_path, new_name=self.new_name)
        
if __name__ == "__main__":
    argv:list = sys.argv
    if len(argv) > 1:
        if argv[1] == 'start':
            Execute().start()
    else:
        _print("Argumentos necessario 'start'")
    