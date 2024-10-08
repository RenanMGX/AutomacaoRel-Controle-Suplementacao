import traceback
from Entities.dependencies.logs import Logs
from Entities.start_report import StartReport
from Entities.file_manipulate import FileManipulate
from Entities.dependencies.functions import _print
from getpass import getuser
import os
import sys
from Entities.dependencies.config import Config

class Execute:
    @property
    def download_path(self) -> str:
        path:str = os.path.join(os.getcwd(), "downloads")
        if not os.path.exists(path):
            os.makedirs(path)    
        return path
    
    @property
    def destiny_path(self) -> str:
        return Config()['paths']['destiny_path']
    
    
    def __init__(self) -> None:
        self.__sap_start_report:StartReport = StartReport('SAP_PRD')
        #self.__file_manipulate: FileManipulate = FileManipulate()
    
    def start(self) -> None:
        file_path_suplementacao:str = self.__sap_start_report.extrair_rel_suplementacao(download_path=self.download_path)
        FileManipulate(file_path_suplementacao).excelToJson(remove_duplicate=True).moveTo(self.destiny_path, new_name="ControleObra_Suplementacao.json")
        
        file_path_reclassificacao:str = self.__sap_start_report.extrair_rel_reclassiicacao(download_path=self.download_path)
        FileManipulate(file_path_reclassificacao).excelToJson().moveTo(self.destiny_path, new_name="ControleObra_Recassificacao.json")
        
        self.__sap_start_report.fechar_sap()
        
    def test(self):
        print("testado")
        
if __name__ == "__main__":
    
    valid_argvs = {
        "start" : Execute().start,
        "teste" : Execute().test
    }
    
    def informativo():
        _print("Argumentos necessario: ")
        for key, value in valid_argvs.items():
            _print(key)
                
    argv:list = sys.argv
    if len(argv) > 1:
        if argv[1] in valid_argvs:
            try:
                valid_argvs[argv[1]]()
                Logs().register(status='Concluido', description="automação executou com exito!")
            except Exception as error:
                Logs().register(status='Error', description="algum erro aconteceu ao executar o script", exception=traceback.format_exc())
        else:
            informativo()
    else:
        informativo()