from typing import Dict
from getpass import getuser

default:Dict[str, Dict[str,object]] = {
    'credential': {
        'crd': 'SAP_PRD'
    },
    'log': {
        'hostname': 'Patrimar-RPA',
        'port' : 80,
        'token': ''
    },
    "paths": {
        'destiny_path': f"C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\RPA - Documentos\\RPA - Dados\\RPA - Controle - Suplementação"
    }
}

