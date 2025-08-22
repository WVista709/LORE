import os, sys
from openpyxl import *
from openpyxl.utils import get_column_letter
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.ensemble import RandomForestClassifier

def letra_cabecalho(cabecalho: str):
    coluna = None
    for col in range(1, ws.max_column + 1):
        cell_value = (ws[f"{get_column_letter(col)}1"].value).upper()
        if cell_value == cabecalho.upper():
            coluna = col
            print(f"Cabeçalho {cabecalho} da coluna {coluna}")
            return get_column_letter(coluna)
    
    print(f"Não existe o cabeçalho {cabecalho}")
    return coluna

def resource_path(relative_path):
    """Obtem o caminho absoluto para o recurso, funciona para dev e PyInstaller."""
    try:
        # PyInstaller cria uma pasta temporária e define essa variável
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

#MATRIZ TRIBUTARIA
path_matriz = resource_path("MatrizTributaria.xlsx")
wb_matriz = load_workbook(path_matriz)

#Excel que vai ser modificado
path_excel = os.path.join(r"C:\Users\Prime Contabil\Downloads", "TW CHECK 2025 07.xlsx")
wb = load_workbook(path_excel)
ws = wb["COMPRAS PRODUTOS"]

#Criando uma nova tabela
if letra_cabecalho("TIPO CFOP") == None:
    ws[get_column_letter(ws.max_column + 1) + "1"] = "TIPO CFOP"

#Pegando a letra da coluna
coluna_cfop = letra_cabecalho("CFOP")
coluna_ncm = letra_cabecalho("Classificação")
coluna_produto = letra_cabecalho("NOME PRODUTO")
coluna_tipo_cfop = letra_cabecalho("TIPO CFOP")

#Criando a aba PRODUTOS e copiado os produtos para la
wb.create_sheet("PRODUTOS")
ws_produtos = wb["PRODUTOS"]
ws_produtos["A1"] = "COMPRAS"
ws_produtos["A2"] = "PRODUTOS"
ws_produtos["B2"] = "NCM"
ws_produtos["C2"] = "CFOP"
ws_produtos["D2"] = "CLASSIFICAÇÃO"

#Pegando os produtos, ncm e cfop
produtos_unicos = set()
row_destino = 3

#Iterar linha por linha para manter a relação produto-NCM-CFOP
for row_num in range(2, ws.max_row + 1):
    produto = ws[f"{coluna_produto}{row_num}"].value
    ncm = ws[f"{coluna_ncm}{row_num}"].value
    cfop = ws[f"{coluna_cfop}{row_num}"].value
    
    # Criar uma chave única para evitar duplicatas
    chave_produto = (produto, ncm, cfop)
    
    if produto is not None and chave_produto not in produtos_unicos:
        produtos_unicos.add(chave_produto)
        
        # Escrever na planilha destino
        ws_produtos[f"A{row_destino}"] = produto
        ws_produtos[f"B{row_destino}"] = ncm
        ws_produtos[f"C{row_destino}"] = cfop
        row_destino += 1

wb.save("TESTE.xlsx")

# 1. Carregar a base tributaria
df_base = pd.read_excel('MatrizTributaria.xlsx')
X_train = df_base['PRODUTO'].astype(str)
y_train = df_base['CLASSIFICAÇÃO'].astype(str)

#ABa de cfop
df_cfop = pd.read_excel('MatrizTributaria.xlsx', sheet_name='CFOP') 

# 2. Vetorizar os nomes dos produtos
vectorizer = TfidfVectorizer()
X_train_vec = vectorizer.fit_transform(X_train)

# 3. Treinar o modelo
model = RandomForestClassifier()
model.fit(X_train_vec, y_train)

# 4. Carregar os produtos a classificar
df_teste = pd.read_excel('TESTE.xlsx', skiprows=1, sheet_name="PRODUTOS")
produtos_teste = df_teste['PRODUTOS'].astype(str)
X_test_vec = vectorizer.transform(produtos_teste)

probas = model.predict_proba(X_test_vec)
max_probas = probas.max(axis=1)

# Defina um limiar de confiança, ex: 0.6
limiar = 0.6
df_teste['CLASSIFICAÇÃO'] = model.predict(X_test_vec)
df_teste['CONFIANCA'] = max_probas

# 5. Prever as classificações
df_teste['CLASSIFICAÇÃO'] = model.predict(X_test_vec)

# Se a confiança for baixa, marque como "Não Classificado"
df_teste.loc[df_teste['CONFIANCA'] < limiar, 'CLASSIFICAÇÃO'] = 'Não Classificado'

df_teste.loc[df_teste['CFOP'] == 1556, 'CLASSIFICAÇÃO'] = 'USO E CONSUMO'

# 6. Salvar o resultado
df_teste.to_excel('TESTE_classificado_IA.xlsx', index=False)
print("Arquivo classificado salvo como TESTE_classificado_IA.xlsx")