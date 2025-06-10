from typing import Dict, Literal
from patrimar_dependencies.sap import SAPManipulation
from patrimar_dependencies.functions import _print, datetime, Functions
from botcity.maestro import * # type: ignore
import os


class FolderNotFound(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)

class StartReport(SAPManipulation):    
    def __init__(self, *, maestro:BotMaestroSDK, user:str, password:str, ambiente:str) -> None:
        self.maestro:BotMaestroSDK = maestro
        super().__init__(user=user, password=password, ambiente=ambiente, new_conection=True)

    @SAPManipulation.start_SAP
    def extrair_rel_suplementacao(self, *, 
                    download_path:str, 
                    file_name:str="ralatorio_suprimentos.xlsx"                    
                    ) -> str:
        if not file_name.endswith(".xlsx"):
            file_name += ".xlsx"
        file_name = datetime.now().strftime(f"%Y%m%d-%H%M%S_{file_name}")
            
        if not os.path.exists(download_path):
            raise FolderNotFound(f"Caminho não encontrado: {download_path}")
        
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
            self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = download_path
            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        except Exception as error:
            _print(f"Erro: {error}")            
            self.maestro.alert(
                task_id=self.maestro.get_execution().task_id,
                title="erro ao gerar relatorio",
                message=str(error),
                alert_type=AlertType.ERROR
            )            
            return "None"
        
        Functions.fechar_excel(file_name)
        return os.path.join(download_path, file_name)
    
    @SAPManipulation.start_SAP
    def extrair_rel_reclassiicacao(self, *, 
                    download_path:str, 
                    file_name:str="ralatorio_suprimentos_reclassificacao.xlsx"                    
                    ) -> str:
        if not file_name.endswith(".xlsx"):
            file_name += ".xlsx"
        file_name = datetime.now().strftime(f"%Y%m%d-%H%M%S_{file_name}")
            
        if not os.path.exists(download_path):
            raise FolderNotFound(f"Caminho não encontrado: {download_path}")
        
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n start_report"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/txtD_SREPOVARI-REPORT").text = "AQICSYSTQV000033RECLASS======="
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/txtSP$00001-LOW").text = "ZFI036"
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            self.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").contextMenu()
            self.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem("&XXL")
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = download_path
            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        except Exception as error:
            _print(f"Erro: {error}")
            self.maestro.alert(
                task_id=self.maestro.get_execution().task_id,
                title="erro ao gerar relatorio",
                message=str(error),
                alert_type=AlertType.ERROR
            )            
            
            
            return "None"
        
        Functions.fechar_excel(file_name)
        return os.path.join(download_path, file_name)
       

if __name__ == "__main__":
    pass