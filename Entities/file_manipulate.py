import os
import pandas as pd
from patrimar_dependencies.functions import Functions
import shutil
from time import sleep

class DirNotFoundError(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class FileManipulate:
    @property
    def file_path(self) -> str:
        return self.__file_path
        
    def __init__(self, file_path:str) -> None:
        self.__file_path:str = file_path
        
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"o arquivo '{self.file_path}' não foi encontrado!")
        
    def __str__(self) -> str:
        return self.file_path
    
    def excelToJson(self, *, remove_duplicate:bool=False):
        if (self.file_path.endswith("xlsx")) or (self.file_path.endswith("xls")):
            Functions.fechar_excel(self.file_path)
            df:pd.DataFrame = pd.read_excel(self.file_path)
            if remove_duplicate:
                df = df.drop_duplicates()
            os.unlink(self.file_path)
            dir_path = os.path.dirname(self.file_path)
            new_file = os.path.basename(self.file_path).split('.')[0] + ".json"
            new_file = os.path.join(dir_path, new_file)
            self.__file_path = new_file
            df.to_json(self.file_path, orient='records', date_format="iso")
            return self
        else:
            raise TypeError("é aceito apenas arquivos excel")
        
    def moveTo(self, destiny:str, *, new_name:str=""):
        if not os.path.exists(destiny):
            raise DirNotFoundError(f"o caminho '{destiny}' não foi encontrado!")
        
        if new_name:
            destiny = os.path.join(destiny, new_name)
        else:
            destiny = os.path.join(destiny, os.path.basename(self.file_path))    
        
        
        for _ in range(2):
            try:
                shutil.move(self.file_path, destiny)
                break
            except shutil.Error:
                os.unlink(destiny)
            sleep(1)
    
        self.__file_path = destiny
        
        return self


if __name__ == "__main__":
    pass