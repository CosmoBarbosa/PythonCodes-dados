"""
faça um codigo em python que leia uma planilha FTV
com varios itens na coluna Descrição_1 e compare os
com a outra planilha resultado_comparacao que tem a
coluna Descrição_R assim verificando os itens iguais,
e removendo da planilha que tem a coluna Descrição_R e
deixando apenas os itens que não são iguais
nesse codigo quero que gere tambem uma outa planilha
duplicados que arvazene a linha da planilha resultado_comparacao.xlsx
que a descrição_R é igual a descrição!

"""





import pandas as pd

# Leitura das planilhas
ftv_path = 'FTV.xlsx'  # Caminho para a planilha FTV
resultado_path = 'resultado_comparacao.xlsx'  # Caminho para a planilha resultado_comparacao

ftv_df = pd.read_excel(ftv_path)
resultado_df = pd.read_excel(resultado_path)

# Comparação das colunas
descricao_ftv = ftv_df['Descrição_1']
descricao_resultado = resultado_df['Descrição_R']

# Identificar duplicados (iguais) e não duplicados
duplicados = resultado_df[resultado_df['Descrição_R'].isin(descricao_ftv)]
resultado_filtrado = resultado_df[~resultado_df['Descrição_R'].isin(descricao_ftv)]

# Salvar as planilhas
output_filtrado_path = 'resultado_filtrado.xlsx'
output_duplicados_path = 'duplicados.xlsx'

resultado_filtrado.to_excel(output_filtrado_path, index=False)
duplicados.to_excel(output_duplicados_path, index=False)

print(f'Planilha filtrada salva em: {output_filtrado_path}')
print(f'Planilha de duplicados salva em: {output_duplicados_path}')
