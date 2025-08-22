import os, sys
from openpyxl import *
from openpyxl.utils import get_column_letter
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.ensemble import RandomForestClassifier

def letra_cabecalho(cabecalho: str, ws: str):
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

#Excel que vai ser modificado
path_excel = os.path.join(r"C:\Users\Prime Contabil\Downloads", "TW CHECK 2025 07.xlsx")
wb = load_workbook(path_excel)
ws_compras_produtos = wb["COMPRAS PRODUTOS"]
ws_vendas_produtos = wb["VENDAS PRODUTOS"]

#Criando uma nova tabela
if letra_cabecalho("TIPO CFOP", ws_compras_produtos) == None:
    ws_compras_produtos[get_column_letter(ws_compras_produtos.max_column + 1) + "1"] = "TIPO CFOP"

#Pegando a letra da coluna
coluna_compras_cfop = letra_cabecalho("CFOP", ws_compras_produtos)
coluna_compras_ncm = letra_cabecalho("Classificação", ws_compras_produtos)
coluna_compras_cest = letra_cabecalho("CEST", ws_compras_produtos)
coluna_compras_produto = letra_cabecalho("NOME PRODUTO", ws_compras_produtos)
coluna_compras_tipo_cfop = letra_cabecalho("TIPO CFOP", ws_compras_produtos)
coluna_compras_pis = letra_cabecalho("CST PIS", ws_compras_produtos)
coluna_compras_cofins = letra_cabecalho("CST COFINS", ws_compras_produtos)
coluna_compras_icms = letra_cabecalho("CST ICMS", ws_compras_produtos)

#Criando a aba PRODUTOS e copiado os produtos para la
wb.create_sheet("PRODUTOS")
wb.create_sheet("CONFIG")
aba_produtos = wb["PRODUTOS"]
aba_produtos["A1"] = "PRODUTOS"
aba_produtos["B1"] = "NCM"
aba_produtos["C1"] = "CFOP"
aba_produtos["D1"] = "ID-CFOP"
aba_produtos["E1"] = "CLASSIFICAÇÃO"
aba_produtos["F1"] = "CONFIANCA"
aba_produtos["G1"] = "CEST PIS"
aba_produtos["H1"] = "CEST COFINS"
aba_produtos["I1"] = "CEST ICMS"
#aba_produtos["J1"] = "CORREÇÃO PIS"
#aba_produtos["K1"] = "CORREÇÃO COFINS"
#aba_produtos["L1"] = "CORREÇÃO ICMS"

#Pegando os produtos, ncm e cfop
produtos_unicos = set()
row_destino = 2

#Iterar linha por linha para manter a relação produto-NCM-CFOP NO COMPRAS
for row_num in range(2, ws_compras_produtos.max_row + 1):
    produto = ws_compras_produtos[f"{coluna_compras_produto}{row_num}"].value
    ncm = ws_compras_produtos[f"{coluna_compras_ncm}{row_num}"].value
    cfop = ws_compras_produtos[f"{coluna_compras_cfop}{row_num}"].value
    pis = ws_compras_produtos[f"{coluna_compras_pis}{row_num}"].value
    cofins = ws_compras_produtos[f"{coluna_compras_cofins}{row_num}"].value
    icms = ws_compras_produtos[f"{coluna_compras_icms}{row_num}"].value
    
    #Criar uma chave única para evitar duplicatas
    chave_produto = (produto, ncm, cfop)
    
    if produto is not None and chave_produto not in produtos_unicos:
        produtos_unicos.add(chave_produto)
        
        #Escrever na planilha destino
        aba_produtos[f"A{row_destino}"] = produto
        aba_produtos[f"B{row_destino}"] = ncm
        aba_produtos[f"C{row_destino}"] = cfop
        aba_produtos[f"G{row_destino}"] = pis
        aba_produtos[f"H{row_destino}"] = cofins
        aba_produtos[f"I{row_destino}"] = icms
        row_destino += 1

#Iterar linha por linha para manter a relação produto-NCM-CFOP NO VENDAS
for row_num in range(2, ws_vendas_produtos.max_row + 1):
    produto = ws_vendas_produtos[f"{coluna_compras_produto}{row_num}"].value
    ncm = ws_vendas_produtos[f"{coluna_compras_ncm}{row_num}"].value
    cfop = ws_vendas_produtos[f"{coluna_compras_cfop}{row_num}"].value
    pis = ws_vendas_produtos[f"{coluna_compras_pis}{row_num}"].value
    cofins = ws_vendas_produtos[f"{coluna_compras_cofins}{row_num}"].value
    icms = ws_vendas_produtos[f"{coluna_compras_icms}{row_num}"].value
    
    #Criar uma chave única para evitar duplicatas
    chave_produto = (produto, ncm, cfop)
    
    if produto is not None and chave_produto not in produtos_unicos:
        produtos_unicos.add(chave_produto)
        
        #Escrever na planilha destino
        aba_produtos[f"A{row_destino}"] = produto
        aba_produtos[f"B{row_destino}"] = ncm
        aba_produtos[f"C{row_destino}"] = cfop
        aba_produtos[f"G{row_destino}"] = pis
        aba_produtos[f"H{row_destino}"] = cofins
        aba_produtos[f"I{row_destino}"] = icms
        row_destino += 1

#ws_produtos.insert_rows(idx=2, amount=10)
wb.save("TESTE.xlsx")
#Ate aqui OK

#Carregar a base tributaria
df_matriz_tributaria = pd.read_excel('MatrizTributaria.xlsx')
X_train = df_matriz_tributaria['PRODUTO'].astype(str)
y_train = df_matriz_tributaria['CLASSIFICAÇÃO'].astype(str)

#ABAS
df_planilha = pd.read_excel("TESTE.xlsx", sheet_name=None)
df_produtos = df_planilha['PRODUTOS']

#Vetorizando os produtos
vectorizer = TfidfVectorizer()
X_train_vec = vectorizer.fit_transform(X_train)

#Criando o modelo de treino
model = RandomForestClassifier()
model.fit(X_train_vec, y_train)

#Aplicando o modelo
produtos_teste = df_produtos['PRODUTOS'].astype(str)
X_test_vec = vectorizer.transform(produtos_teste)

probas = model.predict_proba(X_test_vec)
max_probas = probas.max(axis=1)

#Garantindo uma qualidade sobre a decisão
limiar = 0.6
df_produtos['CLASSIFICAÇÃO'] = model.predict(X_test_vec)
df_produtos['CONFIANCA'] = max_probas
df_produtos.loc[df_produtos['CONFIANCA'] < limiar, 'CLASSIFICAÇÃO'] = 'Não Classificado'
df_produtos.loc[df_produtos['CFOP'] == 1556, 'CLASSIFICAÇÃO'] = 'USO E CONSUMO'
#df_produtos.loc[df_produtos['ID-CFOP']]

# Carrega o workbook existente
wb = load_workbook("TESTE.xlsx", data_only=False)  # data_only=False mantém fórmulas
ws = wb["PRODUTOS"]

# Mapeia colunas por nome (assumindo cabeçalhos na primeira linha)
headers = [cell.value for cell in ws[1]]
col_idx = {name: i+1 for i, name in enumerate(headers)}

# Garante que colunas existam (cria cabeçalhos se faltarem)
def ensure_col(name):
    if name not in col_idx:
        ws.cell(row=1, column=ws.max_column+1, value=name)
        col_idx[name] = ws.max_column
ensure_col("CLASSIFICAÇÃO")
ensure_col("CONFIANCA")

# Escreve valores linha a linha (assumindo que df_produtos está alinhado à aba)
for i, row in df_produtos.iterrows():
    excel_row = i + 2  # +2 por causa do cabeçalho 1-based
    ws.cell(row=excel_row, column=col_idx["CLASSIFICAÇÃO"], value=row["CLASSIFICAÇÃO"])
    ws.cell(row=excel_row, column=col_idx["CONFIANCA"], value=float(row["CONFIANCA"]))

# Salva sem reescrever as outras abas (fórmulas preservadas)
wb.save("TESTE.xlsx")
