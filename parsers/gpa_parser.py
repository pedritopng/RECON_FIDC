# parsers/gpa_parser.py
import pandas as pd
from utils import limpar_valor


def processar(caminho_arquivo):
    """
    Lê e processa o relatório do fundo GPA.
    Retorna um DataFrame com as colunas padronizadas.
    """
    # Lê o arquivo CSV, que usa ponto e vírgula como separador
    df = pd.read_csv(caminho_arquivo, encoding='latin-1', delimiter=';')

    # Seleciona as colunas de interesse e as renomeia para o padrão do programa
    # 'Título' -> Documento
    # 'Razão Social Sacado' -> Nome do cliente
    # 'Vlr Original' -> Valor original do documento
    # 'Total Recdo' -> Valor total que foi pago/recebido
    df = df[['Título', 'Razão Social Sacado', 'Vlr Original', 'Total Recdo']].rename(columns={
        'Título': 'Documento',
        'Razão Social Sacado': 'Sacado_Fundo',
        'Vlr Original': 'Valor_Fundo_Original',
        'Total Recdo': 'Valor_Fundo_Pago'
    })

    # Limpa os valores monetários, convertendo-os para números
    df['Valor_Fundo_Original'] = df['Valor_Fundo_Original'].apply(limpar_valor)
    df['Valor_Fundo_Pago'] = df['Valor_Fundo_Pago'].apply(limpar_valor)

    # Garante que o número do documento seja tratado como texto
    df['Documento'] = df['Documento'].astype(str).str.strip()

    return df
