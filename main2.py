import pandas as pd
from rapidfuzz import process, fuzz  # Para comparação de strings

def gerar_planilha_comparada(caminho_1, caminho_2, saida):
    # Carregar os dados das planilhas
    df1 = pd.read_excel(caminho_1)
    df2 = pd.read_excel(caminho_2)

    # Garantir que as colunas de interesse estejam nomeadas
    col_codigo_1 = "Código"
    col_descricao_1 = "Descrição"
    col_codigo_2 = "Código"
    col_descricao_2 = "Descrição"

    # Ajuste para os nomes das colunas se necessário
    if col_codigo_1 not in df1.columns or col_descricao_1 not in df1.columns:
        raise ValueError("As colunas da Planilha 1 não estão corretas ou foram nomeadas de forma diferente.")
    if col_codigo_2 not in df2.columns or col_descricao_2 not in df2.columns:
        raise ValueError("As colunas da Planilha 2 não estão corretas ou foram nomeadas de forma diferente.")

    # Garantir que as colunas de descrição sejam todas strings
    df1[col_descricao_1] = df1[col_descricao_1].astype(str)
    df2[col_descricao_2] = df2[col_descricao_2].astype(str)

    # Criar lista para o resultado
    resultados = []

    # Para cada item da planilha 1, encontrar a melhor correspondência na planilha 2
    for _, row1 in df1.iterrows():
        descricao_1 = row1[col_descricao_1]
        codigo_1 = row1[col_codigo_1]

        # Encontra a descrição mais próxima na Planilha 2
        melhor_correspondencia = process.extractOne(
            descricao_1,
            df2[col_descricao_2].tolist(),
            scorer=fuzz.token_sort_ratio
        )

        # Recuperar os dados do item correspondente
        descricao_2 = melhor_correspondencia[0]
        similaridade = melhor_correspondencia[1]
        codigo_2 = df2.loc[df2[col_descricao_2] == descricao_2, col_codigo_2].values[0]

        # Adicionar ao resultado
        resultados.append({
            "Código_1": codigo_1,
            "Descrição_1": descricao_1,
            "Código_2": codigo_2,
            "Descrição_2": descricao_2,
            "Similaridade": similaridade,
        })

    # Criar um DataFrame do resultado
    df_resultado = pd.DataFrame(resultados)

    # Salvar em um arquivo Excel
    df_resultado.to_excel(saida, index=False)
    print(f"Planilha gerada com sucesso: {saida}")

# Exemplo de uso
caminho_planilha_1 = "MENOR.xlsx"  # Substitua pelo caminho correto
caminho_planilha_2 = "COMPLETA.xlsx"  # Substitua pelo caminho correto
saida = "resultado_comparacao.xlsx"

gerar_planilha_comparada(caminho_planilha_1, caminho_planilha_2, saida)
