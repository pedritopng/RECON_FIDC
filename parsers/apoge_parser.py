# parsers/apoge_parser.py
import pandas as pd
from utils import limpar_valor


def processar(caminho_arquivo):
    """
    Lê e processa o relatório do fundo APOGE.
    Retorna um DataFrame com as colunas padronizadas.
    """
    # Lê o arquivo CSV, ignorando a primeira linha que é um título
    # dtype={'Documento': str} garante que a coluna de documento seja lida como texto
    df = pd.read_csv(caminho_arquivo, encoding='latin-1', delimiter=';', skiprows=1, dtype={'Documento': str})

    # --- Lógica de Limpeza Definitiva (baseada na sugestão) ---
    # 1. Remove todas as linhas onde o 'Documento' é '0,00' ou '0', que são usadas para resumos.
    # O .str.strip() remove espaços em branco antes da comparação.
    if 'Documento' in df.columns:
        df = df[~df['Documento'].str.strip().isin(['0,00', '0'])]

    # 2. Como segurança, remove qualquer linha que ainda tenha a coluna 'Documento' vazia.
    df.dropna(subset=['Documento'], inplace=True)

    # Seleciona as colunas de interesse e as renomeia para o padrão do programa
    # 'Documento' -> Documento (precisará de limpeza)
    # 'Sacado' -> Nome do cliente
    # 'Valor Face' -> Valor original do documento
    # 'Valor Pago' -> Valor total que foi pago/recebido
    df = df[['Documento', 'Sacado', 'Valor Face', 'Valor Pago']].rename(columns={
        'Sacado': 'Sacado_Fundo',
        'Valor Face': 'Valor_Fundo_Original',
        'Valor Pago': 'Valor_Fundo_Pago'
    })

    # Limpa a coluna 'Documento' removendo o prefixo "DUP - "
    df['Documento'] = df['Documento'].str.replace('DUP - ', '', regex=False).str.strip()

    # Limpa os valores monetários, convertendo-os para números
    df['Valor_Fundo_Original'] = df['Valor_Fundo_Original'].apply(limpar_valor)
    df['Valor_Fundo_Pago'] = df['Valor_Fundo_Pago'].apply(limpar_valor)

    # Garante que o número do documento seja tratado como texto
    df['Documento'] = df['Documento'].astype(str)

    return df
