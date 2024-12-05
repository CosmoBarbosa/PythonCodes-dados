"""

import pandas as pd

# Caminhos dos arquivos
input_file = "coincidem.xlsx"
output_file = "planilha_atualizada.xlsx"
duplicates_file = "duplicados.xlsx"

# Nome da coluna de descrição
description_column = "Descrição_1"

# Ler a planilha
df = pd.read_excel(input_file)

# Verificar se a coluna existe
if description_column not in df.columns:
    raise ValueError(f"A coluna '{description_column}' não foi encontrada na planilha.")

# Identificar duplicados com base na coluna de descrição
duplicated_rows = df[df.duplicated(subset=[description_column], keep=False)]

# Manter apenas uma instância de cada duplicado na planilha original
df_no_duplicates = df.drop_duplicates(subset=[description_column], keep='first')

# Salvar os resultados em planilhas separadas
df_no_duplicates.to_excel(output_file, index=False)
duplicated_rows.to_excel(duplicates_file, index=False)

print(f"Planilha original atualizada salva em: {output_file}")
print(f"Planilha com duplicados salva em: {duplicates_file}")
"""


import pandas as pd

# Caminho da planilha original
input_file = "coincidem.xlsx"
output_file = "coincidem_atualizada.xlsx"
duplicates_file = "duplicados.xlsx"

# Nome da coluna a ser analisada
column_name = "Descrição_2"

# Ler a planilha
df = pd.read_excel(input_file)

# Verificar se a coluna existe
if column_name not in df.columns:
    raise ValueError(f"A coluna '{column_name}' não foi encontrada na planilha.")

# Identificar duplicados com base na coluna
duplicated_rows = df[df.duplicated(subset=[column_name], keep=False)]

# Manter apenas uma instância de cada duplicado na planilha original
df_no_duplicates = df.drop_duplicates(subset=[column_name], keep='first')

# Salvar os resultados em planilhas separadas
df_no_duplicates.to_excel(output_file, index=False)
duplicated_rows.to_excel(duplicates_file, index=False)

print(f"Planilha original atualizada salva em: {output_file}")
print(f"Planilha com duplicados salva em: {duplicates_file}")
