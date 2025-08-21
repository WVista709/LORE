import os
import pandas as pd
from openpyxl import *
from openpyxl.utils import get_column_letter

#Caminho do excel
excel_path = os.path.join("/home/massani/Downloads", "DJE.xlsx")

#Adicionando uma nova coluna na aba de produtos
wb = load_workbook(excel_path)
ws = wb["COMPRAS PRODUTOS"]
ws[get_column_letter(ws.max_column + 1) + "1"] = "TIPO".upper()

def letra_cabecalho(cabecalho: str):
    coluna = None
    for col in range(1, ws.max_column + 1):
        cell_value = ws[f"{get_column_letter(col)}1"].value
        if cell_value == cabecalho:
            coluna = col
            print(f"Cabeçalho {cabecalho} da coluna {coluna}")
            return coluna
    
coluna_cfop = letra_cabecalho("CFOP")
coluna_tipo = letra_cabecalho("TIPO")
