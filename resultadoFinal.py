import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import requests

# Módulo de login embutido
resource = "/parse/login"

#função para fazer o login do usuario
def login(userName, serverurl, AppID):
    password = 'Loldosloll@123' 
    header = {
        'Content-Type': 'application/json',
        'X-Parse-Application-Id': AppID
    }

    payload = {
        'username': userName, 
        'password': password
    }
    response_decoded_json = requests.post(serverurl + resource, json=payload, headers=header)
    response_dict = response_decoded_json.json()
    if 'sessionToken' in response_dict:
        return response_dict['sessionToken']
    else:
        raise Exception(f"Erro de login: {response_dict.get('error', 'Resposta inesperada do servidor')}")

# Função para gerar o DataFrame do cooler
def generate_cooler_df(db_url, header):
    table = "/parse/classes/Cooler"
    field_list = [
        "coolerId",
        "usageStatus",
        "customPatrimonio",
        "controllerId",
        "oemSerial",
    ]
    cooler_df = pd.DataFrame()
    urlParams = {
        "keys": (",".join(field_list)),
        "limit": "1000000"
    }
    response = requests.get(db_url + table, headers=header, params=urlParams)
    cooler_data_json = response.json()["results"]
    cooler_df = pd.DataFrame.from_records(cooler_data_json, exclude=["createdAt", "updatedAt"])
    return cooler_df

#faz uma filtragem da coluna coolerId com a custom patrimonio e compara esse dados e coloca em uma nova 
#coluna os dados tendo prioridade customPatrimonio
def minha_funcao(row, df, custom_patrimonio_column):
    if row[custom_patrimonio_column] != "-":
        return row[custom_patrimonio_column]
    elif row['coolerId'] in df[custom_patrimonio_column].values:
        return df.loc[df[custom_patrimonio_column] == row['coolerId'], custom_patrimonio_column].iloc[0]
    else:
        return row['coolerId']

#pega os ultimos 7 numeros de cada linha da coluna para comparar com a base 
def aplicar_formula(valor):
    valor_str = str(valor)
    return valor_str[-7:]

#converte a coluna para valores numericos
def converter_para_numero(valor):
    try:
        return float(valor)
    except ValueError:
        return valor

#indica quais dados dá match com a nossa base
def combinar_match(row):
    if row['I'] == 'Match' or row['J'] == 'Match':
        return 'Match'
    else:
        return '-'

def select_company_df():
    file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx;*.xls")])
    if file_path:
        entry_company_df.delete(0, tk.END)
        entry_company_df.insert(0, file_path)

def login_and_generate_df():
    global db_url
    global cooler_df
    global header

    header = {  # Definindo header globalmente
        "Content-Type": "application/json",
        "X-Parse-Application-Id": "",  # Preenchido após o login
        "X-Parse-Session-Token": ""    # Preenchido após o login
    }

    
    email = entry_email.get()
    company_id = int(entry_company_id.get())
    file_path = entry_company_df.get()

    if not file_path:
        messagebox.showerror("Erro", "Selecione o arquivo Excel do companyDf.")
        return

    # Ler o arquivo Excel companyDf
    try:
        company_df = pd.read_excel(file_path)
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao ler o arquivo Excel: {str(e)}")
        return

    # Selecionar os dados do companyId no DataFrame
    try:
        company_row = company_df.loc[company_df["companyId"] == company_id]
        db_url = company_row["endPoint"].iloc[0].replace('/parse/', '')

        # Adiciona o esquema "https://" se estiver ausente
        if not db_url.startswith("http://") and not db_url.startswith("https://"):
            db_url = "https://" + db_url

        appId = company_row["appId"].iloc[0]
        
        # Verifica se a URL final está correta
        print(f"URL final do endpoint: {db_url}")
        
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao obter os dados do companyId: {str(e)}")
        return

    # Realizar login
    try:
        session_token = login(email, db_url, appId)
        header['X-Parse-Session-Token'] = session_token
        header['X-Parse-Application-Id'] = appId
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao fazer login: {str(e)}")
        return

    # Consulta do cooler
    try:
        cooler_df = generate_cooler_df(db_url, header)
        messagebox.showinfo("Sucesso", "Login bem-sucedido e dados do cooler salvos.")
        print(cooler_df)  # Exibir os dados do cooler_df no console
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao gerar dados do cooler: {str(e)}")

def select_arquivo():
    global df_aofrio_abr
    global sheet_name_aofrio
    root.filename = filedialog.askopenfilename(initialdir="/", title="Selecione o arquivo Excel", filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
    if root.filename:
        label_arquivo.config(text=root.filename)
        sheet_name_aofrio = entry_sheet_aofrio.get()
        try:
            df_aofrio_abr = pd.read_excel(root.filename, sheet_name=sheet_name_aofrio)
        except Exception as e:
            label_arquivo.config(text="Erro ao ler o arquivo. Certifique-se de selecionar um arquivo Excel válido.")
            print("Erro ao ler o arquivo:", e)

def execute_codigo():
    global df_aofrio_abr
    global custom_patrimonio_column
    nome_para_excluir = 'Scrapped'
    nome_para_excluir2 = 'Disabled'
    custom_patrimonio_column = entry_coluna_patrimonio.get()
    
    try:
        if df_aofrio_abr is None:
            label_arquivo.config(text="Selecione um arquivo Excel antes de executar o código.")
            return
    except NameError:
        label_arquivo.config(text="Selecione um arquivo Excel antes de executar o código.")
        return
    
    df = cooler_df.loc[(cooler_df['usageStatus'] != nome_para_excluir) & (cooler_df['usageStatus'] != nome_para_excluir2)].copy()
    df.drop_duplicates('controllerId', inplace=True)

    df['G'] = None
    df['H'] = None
    df['I'] = None
    df['J'] = None
    df['K'] = None

    df['G'] = df.apply(lambda row: minha_funcao(row, df, custom_patrimonio_column), axis=1)
    df['H'] = df['G'].apply(aplicar_formula)
    df['H'] = df['H'].apply(converter_para_numero)

    sub_df_aofrio = df_aofrio_abr['Nº de série']

    df.rename(columns={'G': 'Nº de série'}, inplace=True)

    df_merge = pd.merge(df, df_aofrio_abr, how='inner', on='Nº de série')

    df['I'] = df['Nº de série'].isin(df_merge['Nº de série']).map({True: 'Match', False: '-'})

    df_aofrio_abr.rename(columns={'Nº de série': 'H'}, inplace=True)

    df_merge2 = pd.merge(df, df_aofrio_abr, how='inner', on='H')

    df['J'] = df['H'].isin(df_merge2['H']).map({True: 'Match', False: '-'})

    df['K'] = df.apply(combinar_match, axis=1)

    df.to_excel('veretidoFinal.xlsx')

    df_att = df.copy()
    df_att.drop(['I', 'J'], axis=1, inplace=True)
    
    nome_para_excluir = '-'
    df_dados_iguais = df_att.loc[(df_att['K'] != nome_para_excluir)].copy()
    df_dados_iguais.drop_duplicates('controllerId', inplace=True)
    df_dados_iguais.to_excel('dados_iguais.xlsx', index=False)
    
    nome_para_excluir2 = 'Match'
    df_dados_diferentes = df_att.loc[(df_att['K'] != nome_para_excluir2)].copy()
    df_dados_diferentes.drop_duplicates('controllerId', inplace=True)
    df_dados_diferentes.to_excel('dados_diferentes.xlsx', index=False)


# Interface de usuário
root = tk.Tk()
root.title("Login e Filtragem de Dados")

# Campos de entrada para o login
label_email = tk.Label(root, text="Email:")
label_email.pack()
entry_email = tk.Entry(root)
entry_email.pack(pady=5)

label_password = tk.Label(root, text="Senha:")
label_password.pack()
entry_password = tk.Entry(root, show="*")
entry_password.pack(pady=5)

label_company_id = tk.Label(root, text="ID do Banco ex(Andina = 92):")
label_company_id.pack()
entry_company_id = tk.Entry(root)
entry_company_id.pack(pady=5)

label_company_df = tk.Label(root, text="Selecione o arquivo Excel do companyDf:")
label_company_df.pack()

# Campo de entrada para exibir o caminho do arquivo selecionado
entry_company_df = tk.Entry(root)
entry_company_df.pack(pady=5)

# Botão para abrir a janela de seleção de arquivo
button_select_company_df = tk.Button(root, text="Selecionar Arquivo", command=select_company_df)
button_select_company_df.pack(pady=5)

# Botão de login e geração do DataFrame do cooler
button_login = tk.Button(root, text="Login e Gerar DataFrame", command=login_and_generate_df)
button_login.pack(pady=10)

# Seção para a seleção do arquivo e execução do código

label_instrucao_coluna_patrimonio = tk.Label(root, text="Digite o nome da coluna 'customPatrimonio':")
label_instrucao_coluna_patrimonio.pack()

entry_coluna_patrimonio = tk.Entry(root)
entry_coluna_patrimonio.pack(pady=5)

label_instrucao_sheet_aofrio = tk.Label(root, text="Digite o nome da aba da planilha Aofrio:")
label_instrucao_sheet_aofrio.pack()

entry_sheet_aofrio = tk.Entry(root)
entry_sheet_aofrio.pack(pady=5)

label_instrucao_arquivo = tk.Label(root, text="Selecione o arquivo Excel:")
label_instrucao_arquivo.pack()

label_arquivo = tk.Label(root, text="")
label_arquivo.pack(pady=5)

button_selecionar_arquivo = tk.Button(root, text="Selecionar Arquivo", command=select_arquivo)
button_selecionar_arquivo.pack()

button_executar = tk.Button(root, text="Executar Código", command=execute_codigo)
button_executar.pack(pady=10)

root.mainloop()    