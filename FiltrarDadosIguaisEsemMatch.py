import pandas as pd

cooler_df = pd.read_excel(r'c:\Users\Adriano.Bitencourt\Desktop\Estagio\Estagio\KoaBrasil.xlsx')

df_aofrio_abr = pd.read_excel(r'C:\Users\Adriano.Bitencourt\Desktop\Estagio\Estagio\Aofrio_Abr_.xlsx', sheet_name='base')

nome_para_excluir = 'Scrapped'
nome_para_excluir2 = 'Disabled'

# Excluir todas as linhas com os nomes especificados
df = cooler_df.loc[(cooler_df['usageStatus'] != nome_para_excluir) & (cooler_df['usageStatus'] != nome_para_excluir2)].copy()
df.drop_duplicates('controllerId', inplace=True)

# Gerar as colunas que serão utilizadas para fazer as consultas
df['G'] = None
df['H'] = None
df['I'] = None
df['J'] = None
df['K'] = None

# Juntando coolerId e Patrimonio mas o patrimonio tem prioridade
def minha_funcao(row):
    if row['customPatrimonio'] != "-":
        return row['customPatrimonio']
    elif row['coolerId'] in df['customPatrimonio'].values:
        return df.loc[df['customPatrimonio'] == row['coolerId'], 'G'].iloc[0]
    else:
        return row['coolerId']

# Função para pegar os últimos 7 dígitos do coolerId
def aplicar_formula(valor):
    valor_str = str(valor)
    return  valor_str[-7:]

# Função para converter valores para números, se possível
def converter_para_numero(valor):
    try:
        return float(valor)
    except ValueError:
        return valor

# Aplicar a função em uma nova coluna
df['G'] = df.apply(minha_funcao, axis=1)

# Aplicar a fórmula na coluna 'G' e armazenar o resultado na coluna 'H'
df['H'] = df['G'].apply(aplicar_formula)

# Converter os valores da coluna 'H' para números (se possível)
df['H'] = df['H'].apply(converter_para_numero)

sub_df_aofrio = df_aofrio_abr['Nº de série']

df.rename(columns={'G': 'Nº de série'}, inplace=True)

# Realizar a junção dos DataFrames usando a coluna 'Nº de série'
df_merge = pd.merge(df, df_aofrio_abr, how='inner', on='Nº de série')

# Adicionar os resultados correspondentes à coluna 'I'
df['I'] = df['Nº de série'].isin(df_merge['Nº de série']).map({True: 'Match', False: '-'})

df_aofrio_abr.rename(columns={'Nº de série': 'H'}, inplace=True)

df_merge2 = pd.merge(df, df_aofrio_abr, how='inner', on='H')

df['J'] = df['H'].isin(df_merge2['H']).map({True: 'Match', False: '-'})

#função para combinar as colunas i e j na coluna K para saber qual dar match com a base 
def combinar_match(row):
    if row['I'] == 'Match' or row['J'] == 'Match':
        return 'Match'
    else:
        return '-'

df['K'] = df.apply(combinar_match, axis=1)

df_att = df.copy()

df_att.drop(['I', 'J'], axis=1, inplace=True)

nome_para_excluir = '-'

df_dados_iguais = df_att.loc[(df_att['K'] != nome_para_excluir)].copy()

df_dados_iguais.drop_duplicates('controllerId', inplace=True)

df_dados_iguais.to_excel('dadosIguai.xlsx')

nome_para_excluir2 = 'Match'

df_dados_diferentes = df_att.loc[(df_att['K'] != nome_para_excluir2)].copy()

df_dados_diferentes.drop_duplicates('controllerId', inplace=True)

df_dados_diferentes.to_excel('semMatch.xlsx')
