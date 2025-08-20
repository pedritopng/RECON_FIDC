import pandas as pd
import re

# --- Metadados do Parser ---
# O 'main.py' usará isso para identificar e carregar o parser automaticamente.
PARSER_INFO = {
    'name': 'Diamante',  # Nome que aparecerá na interface
    'function': 'processar_relatorio'  # Nome padrão da função que o main irá chamar
}


def limpar_valor(valor):
    """
    Converte um valor em formato de string (ex: "1.574,00") para um número float.
    """
    if isinstance(valor, str):
        valor_limpo = valor.replace('.', '').replace(',', '.')
        try:
            return float(valor_limpo)
        except (ValueError, TypeError):
            return None
    elif isinstance(valor, (int, float)):
        return float(valor)
    return None


def processar_relatorio(caminho_arquivo, thread_queue=None):
    """
    Lê e processa o relatório da Diamante (semiestruturado).
    Retorna um DataFrame com as colunas ['Documento', 'Sacado_Fundo', 'Valor_Fundo'].
    """
    if thread_queue:
        thread_queue.put(("progress", 15, "Iniciando leitura do relatório Diamante..."))

    df = pd.read_csv(caminho_arquivo, header=None, encoding='latin-1', delimiter=';', on_bad_lines='warn',
                     engine='python')
    df.columns = [f'col_{i}' for i in range(df.shape[1])]
    dados_extraidos = []

    # Expressões regulares para extrair dados do campo 'histórico'
    regex_recebimento_padrao = re.compile(r'Recebimento cfe Dpl\s+(.*?)\s+-\s+(.*)')
    regex_recebimento_alt = re.compile(r'Recebimento cfe Dpl\s+([\w\d\/-]+)-(.*)')
    regex_recebimento_space = re.compile(r'Recebimento cfe Dpl\s+([\w\d\/-]+(?:-[\w\d]+)?)\s+([A-Za-z].*)')
    regex_reembolso_com_doc = re.compile(r'Reembolso Duplicata\s+([\w\d\/-]+)')
    regex_reembolso_sem_doc = re.compile(r'^Reembolso Duplicata$')
    regex_desconto = re.compile(r'^DESCONTO DUPL CFE BORDERO$')
    regex_pagamento = re.compile(r'Pagamento cfe dpl\.\s+(.*?)-DIAMANTE.*')

    total_rows = len(df)
    for index, row in df.iterrows():
        historico = str(row['col_0']).strip()
        valor_str = str(row.get('col_1', '0'))
        documento, sacado = None, None

        # Tenta aplicar as regex em ordem
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
            # Lançamento genérico se nenhuma regex corresponder
            documento, sacado = (historico, "N/A (Lançamento Genérico)") if historico else (
                f"LANCAMENTO_VAZIO_LINHA_{index}", "N/A")

        if documento and (valor := limpar_valor(valor_str)) is not None and valor > 0:
            dados_extraidos.append({'Documento': documento, 'Sacado_Fundo': sacado, 'Valor_Fundo': valor})

        # Atualiza o progresso periodicamente para não sobrecarregar a fila
        if thread_queue and (index + 1) % 100 == 0:
            progress = 15 + int(20 * (index + 1) / total_rows)
            thread_queue.put(
                ("progress", progress, f"Processando linha {index + 1}/{total_rows} do relatório do fundo..."))

    if thread_queue:
        thread_queue.put(("progress", 35, "Relatório do fundo processado com sucesso."))

    return pd.DataFrame(dados_extraidos)
