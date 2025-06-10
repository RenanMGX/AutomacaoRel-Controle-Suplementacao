"""
WARNING:

Please make sure you install the bot dependencies with `pip install --upgrade -r requirements.txt`
in order to get all the dependencies on your Python environment.

Also, if you are using PyCharm or another IDE, make sure that you use the SAME Python interpreter
as your IDE.

If you get an error like:
```
ModuleNotFoundError: No module named 'botcity'
```

This means that you are likely using a different Python interpreter than the one used to install the dependencies.
To fix this, you can either:
- Use the same interpreter as your IDE and install your bot with `pip install --upgrade -r requirements.txt`
- Use the same interpreter as the one used to install the bot (`pip install --upgrade -r requirements.txt`)

Please refer to the documentation for more information at
https://documentation.botcity.dev/tutorials/custom-automations/python-custom/
"""

# Import for integration with BotCity Maestro SDK
from botcity.maestro import * #type: ignore
import traceback
from patrimar_dependencies.gemini_ia import ErrorIA
from Entities.file_manipulate import FileManipulate
from Entities.start_report import StartReport
from datetime import datetime
import os

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False #type: ignore

class Process:
    def __init__(self) -> None:
        self.process:int = 0


class Execute:
    @property
    def download_path(self) -> str:
        path:str = os.path.join(os.getcwd(), "downloads")
        if not os.path.exists(path):
            os.makedirs(path)    
        return path
    
    
    @staticmethod
    def start():        
        crd_param = execution.parameters.get("crd")
        if not isinstance(crd_param, str):
            raise ValueError("Parâmetro 'crd_param' deve ser uma string representando o label da credencial.")
        
        download_path:str = os.path.join(os.getcwd(), "downloads")
        if not os.path.exists(download_path):
            os.makedirs(download_path)    
        
        destiny_path:str = f'C:\\Users\\{os.getlogin()}\\PATRIMAR ENGENHARIA S A\\RPA - Documentos\\RPA - Dados\\RPA - Controle - Suplementação'
        if not os.path.exists(destiny_path):
            raise FileNotFoundError(f"O caminho '{destiny_path}' não foi encontrado. Verifique se o diretório existe.")
        
        try:
            sap_start_report:StartReport = StartReport(
                                                        maestro=maestro,
                                                        user=maestro.get_credential(label=crd_param, key="user"),
                                                        password=maestro.get_credential(label=crd_param, key="password"),
                                                        ambiente=maestro.get_credential(label=crd_param, key="ambiente")
            )
            
            file_path_suplementacao:str = sap_start_report.extrair_rel_suplementacao(download_path=download_path)
            FileManipulate(file_path_suplementacao).excelToJson(remove_duplicate=True).moveTo(destiny_path, new_name="ControleObra_Suplementacao.json")
            process.process += 1
            
            file_path_reclassificacao:str = sap_start_report.extrair_rel_reclassiicacao(download_path=download_path)
            FileManipulate(file_path_reclassificacao).excelToJson().moveTo(destiny_path, new_name="ControleObra_Recassificacao.json")
            process.process += 1
        
        finally:
            sap_start_report.fechar_sap() #type: ignore


if __name__ == '__main__':
    maestro = BotMaestroSDK.from_sys_args()
    execution = maestro.get_execution()
    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")


    try:
        process = Process()
        Execute.start()
        
        maestro.finish_task(
                    task_id=execution.task_id,
                    status=AutomationTaskFinishStatus.SUCCESS,
                    message="Tarefa Extração Relatorio Controle - Suplementação finalizada com sucesso",
                    total_items=2, # Número total de itens processados
                    processed_items=process.process, # Número de itens processados com sucesso
                    failed_items=0 # Número de itens processados com falha
        )
        
    except Exception as error:
        ia_response = "Sem Resposta da IA"
        try:
            token = maestro.get_credential(label="GeminiIA-Token-Default", key="token")
            if isinstance(token, str):
                ia_result = ErrorIA.error_message(
                    token=token,
                    message=traceback.format_exc()
                )
                ia_response = ia_result.replace("\n", " ")
        except Exception as e:
            maestro.error(task_id=int(execution.task_id), exception=e)

        maestro.error(task_id=int(execution.task_id), exception=error, tags={"IA Response": ia_response})
