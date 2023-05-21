#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import streamlit as st
import re
import pandas as pd 
from functools import reduce
import numpy as np
import import openpyxl

def main():
    # Seu código atual aqui
    
    upload_button = st.file_uploader("Upload de arquivo")
    if upload_button is not None:
        file_path = upload_button.name
        extension = check_file_extension(file_path)
        if extension == "csv":
            st.write("O arquivo é um arquivo CSV.")
            df = pd.read_csv(upload_button)  # Para CSV
            id_dados(df)
        elif extension == "xlsx":
            st.write("O arquivo é um arquivo XLSX.")
            df = pd.read_excel(upload_button)  # Para XLSX
            id_dados(df, file_path)
        else:
            st.write("O arquivo não possui uma extensão válida.")

def check_file_extension(file_path):
    if re.search(r"\.xlsx$", file_path, re.IGNORECASE):
        return "xlsx"
    elif re.search(r"\.csv$", file_path, re.IGNORECASE):
        return "csv"
    else:
        return None

def id_dados(df, file_path):
    # Processar os dados e aplicar o método WISP
    
    df = pd.read_excel(file_path)
    
    colunas = df.columns.tolist()

    qnt_colunas = len(colunas)

    list_PorC = df.iloc[0].values.tolist()
    list_Weight = df.iloc[1].values.tolist()
    list_Weight.pop(0)
    list_PorC.pop(0)
    primeira_coluna = df.iloc[2:, 0].reset_index(drop=True)
    
    
    indices_remover = [0, 1]
    dfr = df.drop(indices_remover)
    dfr = dfr.drop(columns=dfr.columns[0]) 


    for i in range(1, qnt_colunas-1):
        dfr.iloc[:, i] = pd.to_numeric(dfr.iloc[:, i])

    df_vazio = pd.DataFrame()

    for j in range(0, qnt_colunas-1):
        for i in range(len(dfr)):
            maior_valor = dfr.iloc[:, j].max()
            valor = dfr.iloc[i, j]
            df_vazio.at[i, j] = valor/maior_valor


    dfff = pd.DataFrame()
    for i in range(len(df_vazio)):
        
        linha = df_vazio.iloc[i]  # Seleciona a primeira linha do DataFrame
        lista = linha.tolist()

        resultado = np.multiply(list_Weight, lista).tolist()
    
        dicionario = {f'{colunas[i+1]}': valor for i, valor in enumerate(resultado)}
    
        # Adiciona o dicionário como uma nova linha no DataFrame
        dfff = dfff.append(dicionario, ignore_index=True)


    dfff.insert(0, colunas[0], primeira_coluna)

    df_Ui = pd.DataFrame()

    for j in range(len(dfff)):
        resultado1 = []
        resultado2 = []
        for i in range(1, qnt_colunas):

            if list_PorC[i-1] == 'P':
                valor = dfff.iloc[j,i]
                resultado1.append(valor)
                #df_Ui.at[i, j] = resultado
    
            if list_PorC[i-1] == 'C':
                valor2 = dfff.iloc[j,i]
                resultado2.append(valor2)
        
        
        Uiwsd = reduce(lambda x, y: x + y, resultado1) - reduce(lambda x, y: x + y, resultado2)
        Uiwpd = reduce(lambda x, y: x * y, resultado1) - reduce(lambda x, y: x * y, resultado2)
        Uiwsr = reduce(lambda x, y: x + y, resultado1) / reduce(lambda x, y: x + y, resultado2)
        Uiwpr = reduce(lambda x, y: x * y, resultado1) / reduce(lambda x, y: x * y, resultado2)
    
        df_Ui.at[j, "Uiwsd"] = Uiwsd
        df_Ui.at[j, "Uiwpd"] = Uiwpd
        df_Ui.at[j, "Uiwsr"] = Uiwsr
        df_Ui.at[j, "Uiwpr"] = Uiwpr

    maior_valor_Uiwsd = df_Ui.iloc[:, 0].max()
    maior_valor_Uiwpd = df_Ui.iloc[:, 1].max()
    maior_valor_Uiwsr = df_Ui.iloc[:, 2].max()
    maior_valor_Uiwpr = df_Ui.iloc[:, 3].max()

    df_Uii = pd.DataFrame()

    df_Uii["Uiwsd"] = df_Ui["Uiwsd"] / (1 + maior_valor_Uiwsd)
    df_Uii["Uiwpd"] = df_Ui["Uiwpd"] / (1 + maior_valor_Uiwpd)
    df_Uii["Uiwsr"] = df_Ui["Uiwsr"] / (1 + maior_valor_Uiwsr)
    df_Uii["Uiwpr"] = df_Ui["Uiwpr"] / (1 + maior_valor_Uiwpr)

    df_Uii["Absolute Value"] = (df_Uii["Uiwsd"] + df_Uii["Uiwpd"] + df_Uii["Uiwsr"] + df_Uii["Uiwpr"])/4

    df_Uii.insert(0, colunas[0], primeira_coluna)

    df_Uii = df_Uii.sort_values(by="Absolute Value", ascending=False)
    df_Uii = df_Uii.reset_index(drop=True)

    df_Uii["Ranking"] = df_Uii.index + 1

    df_Uii.index += 1  # Incrementa o índice em 1 para começar de 1
    df_Uii.rename(columns={'index': 'Ranking'}, inplace=True)  # Renomeia a coluna do índice para "Ranking"
    df_Uii.reset_index(inplace=False)

    salvar(file_path)
    
    st.write('Result Table:\n', df_Uii)

def salvar(file_path):
    # Salvar o arquivo
    
    padrao = r'(.*)/'  # Regex para capturar todo o texto antes do último caractere '\'

    resultado = re.match(padrao, file_path)

    if resultado:
        caminho = resultado.group(1)
        df_Uii.to_excel(caminho + r'\File - Wisp Ranking.xlsx', index=False)
    else:
        st.write('Path not found.')

if __name__ == '__main__':
    main()

