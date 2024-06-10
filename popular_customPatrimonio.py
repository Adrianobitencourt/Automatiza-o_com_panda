import os
import pandas as pd
import requests, json, getpass
import numpy as np
import json
from datetime import datetime, timedelta
from openpyxl import load_workbook
from geopy.geocoders import Nominatim
import tkinter as tk

# Função para obter os valores de entrada do usuário
def get_input():
    company_id = int(entry_company_id.get())
    user_name = entry_user_name.get()
    password = entry_password.get()
    custom_patrimonio = entry_custom_patrimonio.get()

    main(company_id, user_name, password, custom_patrimonio)

def main(company_id, user_name, password, custom_patrimonio):
    resource = "/parse/login"
    
    def login(userName, password, serverurl, AppID):
        header = {'Content-Type': 'application/json',
                'X-Parse-Application-Id': AppID} 

        payload = {'username': userName, 
                'password': password}
        response_decoded_json = requests.post(serverurl + resource, json=payload, headers=header)
        response_dict = response_decoded_json.json()
        return(response_dict['sessionToken'])

    companyDf = pd.read_excel("company_Adriano.xlsx")  # Change directory
    company = companyDf[companyDf["companyId"] == company_id]

    db_url = company["endPoint"].values[0].replace('/parse/', '')
    appId = company["appId"].values[0]

    sessionToken = login(user_name, password, db_url, appId)

    header = {
        "Content-Type": "application/json",
        "X-Parse-Application-Id": appId,
        'X-Parse-Session-Token': sessionToken
    }

    # Cooler Query
    table = "/parse/classes/Cooler"
    field_list = [
        "coolerId",
        "usageStatus",
        "customPatrimonio",
        "controllerId",
        "oemSerial",
    ]

    cooler_df = pd.DataFrame()  # Will hold cooler data
    urlParams = {
        "keys": (",".join(field_list)),
        "limit": "1000000"
    }

    response = requests.get(db_url + table, headers=header, params=urlParams)
    cooler_data_json = response.json()["results"]
    cooler_df = pd.DataFrame.from_records(cooler_data_json, exclude=["createdAt", "updatedAt"])

    df_aofrio_abr = pd.read_excel('Andina - Relatório de Embarque 14.05.2024.xlsx', sheet_name='Ativos')

    nome_para_excluir = 'Scrapped'
    nome_para_excluir2 = 'Disabled'

    # Excluir todas as linhas com os nomes especificados
    df = cooler_df.loc[(cooler_df['usageStatus']!= nome_para_excluir) & (cooler_df['usageStatus']!= nome_para_excluir2)].copy()
    df.drop_duplicates('controllerId', inplace=True)

    # Gerar as colunas que serão utilizadas para fazer as consultas
    df['G'] = None
    df['H'] = None
    df['I'] = None
    df['J'] = None
    df['K'] = None

    # Juntando coolerId e Patrimonio mas o patrimonio tem prioridade
    def minha_funcao(row, customPatrimonio):
        if row[customPatrimonio]!= "-":
            return row[customPatrimonio]
        elif row['coolerId'] in df[customPatrimonio].values:
            return df.loc[df[customPatrimonio] == row['coolerId'], 'G'].iloc[0]
        else:
            return ''

    df['G'] = df.apply(minha_funcao, axis=1, customPatrimonio=custom_patrimonio)

    df_tem_custompatrimonio = df.loc[(df['G']!= '')].copy()

    df_tem_custompatrimonio.dropna(subset=['G'], inplace=True)

    def aplicar_formula(valor):
        valor_str = str(valor)
        return  valor_str[-7:]

    def converter_para_numero(valor):
        try:
            return float(valor)
        except ValueError:
            return valor

    def funcao_att(row, customPatrimonio):
        if row[customPatrimonio]!= "-":
            return row[customPatrimonio]
        elif row[customPatrimonio] == "-":
            return row['1º Patrimonial']
        else:
            return ''

    df_nao_tem_custompatrimonio = df.loc[(df['G'] == '' )].copy()

    df_nao_tem_custompatrimonio['G'] = df_nao_tem_custompatrimonio['coolerId'].apply(aplicar_formula)

    df_nao_tem_custompatrimonio['G'] = df_nao_tem_custompatrimonio['G'].apply(converter_para_numero)

    df_nao_tem_custompatrimonio.rename(columns={'G': 'Nr. Série'}, inplace=True)

    sub_df_aofrio = df_aofrio_abr[['Nr. Série', '1º Patrimonial', '2º Patrimonial']]

    sub_df_aofrio['Nr. Série'] = sub_df_aofrio['Nr. Série'].apply(converter_para_numero)

    df_nao_tem_custompatrimonio['Nr. Série'] = df_nao_tem_custompatrimonio['Nr. Série'].astype(str)

    sub_df_aofrio['Nr. Série'] = sub_df_aofrio['Nr. Série'].astype(str)

    # Realizar a junção dos DataFrames usando a coluna 'Nr. Série'
    df_merge = pd.merge(df_nao_tem_custompatrimonio, sub_df_aofrio, how='inner', on='Nr. Série')

    df_merge.drop(['I', 'J', 'K','H'], axis=1, inplace=True)

    df_concatenado2 = pd.concat([df_merge, df_tem_custompatrimonio], ignore_index=True)

    df_concatenado2.drop(['I', 'J', 'K','H', 'G'], axis=1, inplace=True)

    df_concatenado2.drop_duplicates('coolerId', inplace=True)

    df_concatenado = pd.concat([df_merge, df_nao_tem_custompatrimonio], ignore_index=True)

    # Remover linhas duplicadas pelo 'coolerId' e manter apenas as que a coluna '1º Patrimonial' não é NaN
    df_sem_duplicatas = df_concatenado.drop_duplicates(subset='coolerId')

    df_nan = df_sem_duplicatas[df_sem_duplicatas['1º Patrimonial'].isna()]

    df_nan_att = df_nan.copy()

    df_nan_att.drop(['I', 'J', 'K','H', '2º Patrimonial','1º Patrimonial', 'Nr. Série'], axis=1, inplace=True)

    df_desco_patri = pd.merge(df_nan_att, sub_df_aofrio, how='left', left_on='coolerId', right_on='1º Patrimonial')

    df_com_patri = df_desco_patri[df_desco_patri['1º Patrimonial'].notna()]

    df_concatenado3 = pd.concat([df_com_patri, df_concatenado2], ignore_index=True)

    df_patri_att = df_concatenado3.copy()  

    df_patri_att['customPatrimonioAtt'] = df_patri_att.apply(funcao_att, axis=1, customPatrimonio=custom_patrimonio)  

    df_patri_att.to_excel('customPatrimonioEncontrada.xlsx')

    df_sem_patri = df_desco_patri[df_desco_patri['1º Patrimonial'].isna()]

    df_sem_patri.to_excel('customPatrimonioNãoEncontrada.xlsx')

window = tk.Tk()

# Criar os campos de entrada
entry_company_id = tk.Entry(window)
entry_user_name = tk.Entry(window)
entry_password = tk.Entry(window, show="*")
entry_custom_patrimonio = tk.Entry(window)

# Criar os rótulos
label_company_id = tk.Label(window, text="Número da empresa:")
label_user_name = tk.Label(window, text="Nome de usuário:")
label_password = tk.Label(window, text="Senha:")
label_custom_patrimonio = tk.Label(window, text="Nome da coluna de patrimônio:")

# Adicionar os campos de entrada e rótulos à janela
label_company_id.grid(row=0, column=0)
entry_company_id.grid(row=0, column=1)

label_user_name.grid(row=1, column=0)
entry_user_name.grid(row=1, column=1)

label_password.grid(row=2, column=0)
entry_password.grid(row=2, column=1)

label_custom_patrimonio.grid(row=3, column=0)
entry_custom_patrimonio.grid(row=3, column=1)

# Criar o botão para obter os valores de entrada
button_get_input = tk.Button(window, text="Obter valores de entrada", command=get_input)
button_get_input.grid(row=4, column=0, columnspan=2)

# Iniciar o loop principal da janela
window.mainloop()