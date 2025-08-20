# parsers/diamante_parser.py
import pandas as pd
from utils import limpar_valor  # <-- MUDANÇA AQUI


def processar(caminho_arquivo):
    """
    Lê e processa o relatório da Diamante (estruturado).
    Retorna um DataFrame com as colunas: ['Documento', 'Sacado_Fundo', 'Valor_Fundo_Original', 'Valor_Fundo_Pago']
    """
    df = pd.read_csv(caminho_arquivo, encoding='latin-1', delimiter=',')

    # Renomeia as colunas para um padrão genérico
    df = df[['Documento', 'Sacado', 'Valor', 'Valor Pago']].rename(columns={
        'Valor': 'Valor_Fundo_Original',
        'Valor Pago': 'Valor_Fundo_Pago',
        'Sacado': 'Sacado_Fundo'
    })

    # Limpa os valores
    df['Valor_Fundo_Original'] = df['Valor_Fundo_Original'].apply(limpar_valor)
    df['Valor_Fundo_Pago'] = df['Valor_Fundo_Pago'].apply(limpar_valor)
    df['Documento'] = df['Documento'].astype(str).str.strip()

    return df
