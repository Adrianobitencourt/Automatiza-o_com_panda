import pandas as pd
import numpy as np

df_sem_matching = pd.read_excel('semMatch.xlsx')
df_sub_semmatch = pd.read_excel(r'C:\Users\Adriano.Bitencourt\Desktop\Estagio\Estagio\Aofrio_Abr_.xlsx', sheet_name='base')

num_columns = df_sem_matching.shape[1]

df_sem_matching = df_sem_matching.iloc[:, :num_columns-3]

def aplicar_formula(valor):
    valor_str = str(valor)
    return valor_str[:10]

def aplicar_formula2(valor):
    valor_str = str(valor)
    return valor_str[12:19]

def aplicar_formulaB(valor):
    valor_str = str(valor)
    return 'B'+valor_str[:9]

def aplicar_formulaB10(valor):
    valor_str = str(valor)
    return 'B'+valor_str[:10]

def aplicar_formulaC(valor):
    valor_str = str(valor)
    return 'C'+valor_str[:10]

def aplicar_formulaN(valor):
    valor_str = str(valor)
    return 'N'+valor_str[:10]

def aplicar_formulaC10(valor):
    valor_str = str(valor)
    return 'C'+valor_str[:9]

def converter_para_numero(valor):
    try:
        return float(valor)
    except ValueError:
        return valor

def combinar_matching(row):
    if row['MatchPorSlice2'] == 'Match' or row['MatchPorControllerIdB'] == 'Match'or row['MatchPorSlice'] == 'Match' or row['MatchPorControllerIdN'] == 'Match' or row['MatchPorControllerIdC'] == 'Match' or row['MatchPorControllerIdC10'] == 'Match' or row['MatchPorControllerIdB10'] == 'Match':
        return 'Match'
    else:
        return '-'


dfsubandina = df_sub_semmatch[['Nº de série', 'Nº série fabricante', 'N º controlador SAP']]

df_sem_matching['slice'] = None
df_sem_matching['MatchPorSlice'] = None
df_sem_matching['slice2'] = None
df_sem_matching['MatchPorSlice2'] = None
df_sem_matching['MatchPorControllerIdB'] = None
df_sem_matching['sliceControllerIdB'] = None
df_sem_matching['MatchPorControllerIdC'] = None
df_sem_matching['sliceControllerIdC'] = None
df_sem_matching['sliceControllerIdC10'] = None
df_sem_matching['MatchPorControllerIdC10'] = None
df_sem_matching['sliceControllerIdB10'] = None
df_sem_matching['MatchPorControllerIdB10'] = None
df_sem_matching['sliceControllerIdN'] = None
df_sem_matching['MatchPorControllerIdN'] = None
df_sem_matching['TodosOsMatch'] = None

df_sem_matching['slice'] = df_sem_matching['controllerId'].apply(aplicar_formula)
df_sem_matching['slice'] = df_sem_matching['slice'].apply(converter_para_numero)
df_sem_matching['slice'] = df_sem_matching['slice'].astype(str)
df_sem_matching['slice2'] = df_sem_matching['oemSerial'].apply(aplicar_formula2)
df_sem_matching['sliceControllerIdB'] = df_sem_matching['controllerId'].apply(aplicar_formulaB)
df_sem_matching['sliceControllerIdN'] = df_sem_matching['controllerId'].apply(aplicar_formulaN)
df_sem_matching['sliceControllerIdC'] = df_sem_matching['controllerId'].apply(aplicar_formulaC)
df_sem_matching['sliceControllerIdC10'] = df_sem_matching['controllerId'].apply(aplicar_formulaC10)
df_sem_matching['sliceControllerIdB10'] = df_sem_matching['controllerId'].apply(aplicar_formulaB10)

merge_por_slice = pd.merge(df_sem_matching, dfsubandina, how='inner', left_on='slice', right_on='Nº série fabricante')

# Adicionar os resultados correspondentes à coluna 'MatchPorSlice'
df_sem_matching['MatchPorSlice'] = np.where(df_sem_matching['slice'].isin(dfsubandina['N º controlador SAP']), 'Match', '-')
df_sem_matching['MatchPorControllerIdB'] = np.where(df_sem_matching['sliceControllerIdB'].isin(dfsubandina['N º controlador SAP']), 'Match', '-')
df_sem_matching['MatchPorControllerIdC'] = np.where(df_sem_matching['sliceControllerIdC'].isin(dfsubandina['N º controlador SAP']), 'Match', '-')
df_sem_matching['MatchPorControllerIdC10'] = np.where(df_sem_matching['sliceControllerIdC10'].isin(dfsubandina['N º controlador SAP']), 'Match', '-')
df_sem_matching['MatchPorControllerIdB10'] = np.where(df_sem_matching['sliceControllerIdB10'].isin(dfsubandina['N º controlador SAP']), 'Match', '-')
df_sem_matching['MatchPorControllerIdN'] = np.where(df_sem_matching['sliceControllerIdN'].isin(dfsubandina['N º controlador SAP']), 'Match', '-')
df_sem_matching['MatchPorSlice2'] = np.where(df_sem_matching['slice2'].isin(dfsubandina['Nº de série']), 'Match', '-')

df_sem_matching['TodosOsMatch'] = df_sem_matching.apply(combinar_matching, axis=1)

df_sem_matching.drop(['sliceControllerIdB', 'sliceControllerIdC', 'sliceControllerIdN', 'sliceControllerIdC10', 'sliceControllerIdB10', 'slice', 'slice2'], axis=1, inplace=True)

df_sem_matching.to_excel('Match_dos_sem_matching6.xlsx')