# parsers/nosso_relatorio_parser.py
import pandas as pd
import re
from utils import limpar_valor  # <-- MUDANÇA AQUI


def processar(caminho_arquivo):
    """
    Lê e processa o nosso relatório (semiestruturado).
    Retorna um DataFrame com as colunas: ['Documento', 'Sacado_Nosso', 'Valor_Nosso']
    """
    df = pd.read_csv(caminho_arquivo, header=None, encoding='latin-1', delimiter=';', on_bad_lines='warn',
                     engine='python')
    df.columns = [f'col_{i}' for i in range(df.shape[1])]

    dados_extraidos = []
    # Expressões regulares para extrair dados do campo de histórico
    regex_recebimento_padrao = re.compile(r'Recebimento cfe Dpl\s+(.*?)\s+-\s+(.*)')
    regex_recebimento_alt = re.compile(r'Recebimento cfe Dpl\s+([\w\d\/-]+)-(.*)')
    regex_recebimento_space = re.compile(r'Recebimento cfe Dpl\s+([\w\d\/-]+(?:-[\w\d]+)?)\s+([A-Za-z].*)')
    regex_reembolso_com_doc = re.compile(r'Reembolso Duplicata\s+([\w\d\/-]+)')
    regex_reembolso_sem_doc = re.compile(r'^Reembolso Duplicata$')
    regex_desconto = re.compile(r'^DESCONTO DUPL CFE BORDERO$')
    regex_pagamento = re.compile(r'Pagamento cfe dpl\.\s+(.*?)-DIAMANTE.*')

    for index, row in df.iterrows():
        historico = str(row['col_0']).strip()
        valor_str = str(row.get('col_1', '0'))
        documento, sacado = None, None

        match = (regex_recebimento_padrao.search(historico) or
                 regex_recebimento_alt.search(historico) or
                 regex_recebimento_space.search(historico))

        if match:
            documento, sacado = match.group(1).strip(), match.group(2).strip()
        elif (match := regex_pagamento.search(historico)):
            documento, sacado = match.group(1).strip(), "N/A (Pagamento)"
        elif (match := regex_reembolso_com_doc.search(historico)):
            documento, sacado = match.group(1).strip(), "N/A (Reembolso)"
        elif regex_reembolso_sem_doc.search(historico):
            documento, sacado = f"REEMBOLSO_SEM_DOC_{index}", "N/A (Reembolso sem doc)"
        elif regex_desconto.search(historico):
            documento, sacado = f"DESCONTO_BORDERO_{index}", "N/A (Desconto Bordero)"
        else:
            documento, sacado = (historico, "N/A (Lançamento Genérico)") if historico else (
                f"LANCAMENTO_VAZIO_LINHA_{index}", "N/A")

        if documento and (valor := limpar_valor(valor_str)) is not None and valor > 0:
            dados_extraidos.append({'Documento': documento, 'Sacado_Nosso': sacado, 'Valor_Nosso': valor})

    return pd.DataFrame(dados_extraidos)
