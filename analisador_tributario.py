import pandas as pd
import os
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.ensemble import IsolationForest
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_absolute_error
from sklearn.linear_model import LinearRegression

#Abrindo o excel
excel_path = r"C:\Users\Prime Contabil\Downloads\DJE.xlsx"
excel = pd.read_excel(excel_path, sheet_name="COMPRAS PRODUTOS", engine="openpyxl")

excel.columns = [c.strip().lower() for c in excel.columns]

coluna = "nome produto"  # troque pelo nome exato
serie = excel[coluna]              # Series (um vetor)
lista_valores = excel[coluna].tolist()  # como lista
print(serie.head())
