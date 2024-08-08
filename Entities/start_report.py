from typing import Literal
from .dependencies.sap import SAPManipulation
from .dependencies.credenciais import Credential
from .dependencies.logs import Logs
from .dependencies.functions import _print, os, datetime, Functions

class FolderNotFound(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)

class StartReport(SAPManipulation):
    def __init__(self, credential:Literal["SAP_PRD", "SAP_QAS", "SAP_PRD-Renan"]) -> None:
        crd:dict = Credential(credential).load()
        super().__init__(user=crd['user'], password=crd['password'], ambiente=crd['ambiente'])
        self.__log:Logs = Logs()

    @SAPManipulation.start_SAP
    def extrair_rel(self, *, path_destiny:str, file_name:str="ralatorio_suprimentos.xlsx") -> str:
        if not file_name.endswith(".xlsx"):
            file_name += ".xlsx"
        file_name = datetime.now().strftime(f"%Y%m%d-%H%M%S_{file_name}")
            
        if not os.path.exists(path_destiny):
            raise FolderNotFound(f"Caminho não encontrado: {path_destiny}")
        
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n start_report"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/txtD_SREPOVARI-REPORT").text = "AQICSYSTQV000033STATUS_SUPLEM="
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            self.session.findById("wnd[1]/usr/txtRS38R-DBACC").text = "99999"
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            # self.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").setCurrentCell(2,"CENTRO")
            # self.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectedRows = "2"
            self.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").contextMenu()
            self.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem("&XXL")
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = path_destiny
            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        except Exception as error:
            _print(f"Erro: {error}")
            self.__log.register(status='Error', description="erro ao gerar relatorio")
            return "None"
        
        Functions.fechar_excel(file_name)
        self.fechar_sap()
        return os.path.join(path_destiny, file_name)
        

if __name__ == "__main__":
    pass