import pandas as pd


#ABRE TXT E SALVA XLSX LIMPANDO ENCOUDINGS
df_txt = pd.read_csv('ML84.txt', sep='\s+', skiprows=2, encoding='latin-1')
df_txt.to_excel('Rezando.xlsx', index=True)

# Ler o arquivo Excel
df = pd.read_excel('Rezando.xlsx')

# Inicializando variáveis
dados_alinhados_1 = ''
dados_alinhados_2 = ''

# Iterando pelas linhas do DataFrame
for idx, row in df.iterrows():
    valor = str(row['Unnamed: 0'])
    if valor.startswith('450') and len(valor) <= 10:
        dados_alinhados_1 = valor
        #Aqui queria pegar o valor existente na linha abaixo do dados_alinhados_1 e salvar em uma variavel
        dados_alinhados_1_subsequente = str(df.iloc[idx + 1]['Unnamed: 0']) if idx + 1 < len(df) else ''
    elif valor.startswith('100') and len(valor) <= 10:
        dados_alinhados_2 = valor
        # Se ambos os valores estiverem preenchidos, alinhar na linha atual
        if dados_alinhados_1 and dados_alinhados_2:
            df.loc[idx, 'Dados alinhados 1'] = dados_alinhados_1
            df.loc[idx, 'Dados alinhados 2'] = dados_alinhados_1_subsequente
            #df.loc[idx, 'Dados alinhados 3'] = dados_alinhados_1_subsequente
df = df.dropna(subset=['Unnamed: 2'])
df = df.drop(["Unnamed: 7", "Unnamed: 8","Unnamed: 9", "Item"], axis=1)
df = df.drop(df.index[0])
df = df.rename(columns={'Unnamed: 0': 'FolhRegSrv', 'Unnamed: 1': 'DT CRIACAO', 'Unnamed: 2': 'DT MODIFICACAO', 'Unnamed: 3': 'Valor Liquido', 'Unnamed: 4': '   Valor bruto', 'Unnamed: 5': 'Criado por', 'Unnamed: 6': 'Modificado por', 'Dados alinhados 1': 'PO', 'Dados alinhados 2': 'PO_ITEM'})
df.to_excel('SAP - ML84.xlsx', index=False)
print(df)