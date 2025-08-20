# utils.py
import re
import locale

def limpar_valor(valor):
    """
    Converte um valor em formato de string (ex: "1.574,00") para um número float.
    """
    if isinstance(valor, str):
        # Remove caracteres de milhar e substitui a vírgula decimal por ponto
        valor_limpo = valor.replace('.', '').replace(',', '.')
        try:
            return float(valor_limpo)
        except (ValueError, TypeError):
            return None
    elif isinstance(valor, (int, float)):
        return float(valor)
    return None


def normalizar_documento(doc_str):
    """
    Normaliza o número do documento para um formato canônico para permitir a correspondência.
    Extrai o padrão 'num/num' ou 'num-num' mesmo que haja texto adicional.
    Ex: '58817/03-DME' se torna '58817/003'.
    """
    if not isinstance(doc_str, str):
        return str(doc_str)

    # Procura pelo padrão de documento (números com / ou - no meio)
    match = re.search(r'(\d+[\/-]\d+)', doc_str)
    if not match:
        return doc_str.strip()

    # Limpa e padroniza o documento encontrado
    doc_str_norm = match.group(1).replace('-', '/')
    partes = doc_str_norm.split('/')
    if len(partes) == 2:
        principal, parcela = partes
        # Garante que a parcela tenha 3 dígitos (ex: '3' vira '003')
        parcela_padded = parcela.zfill(3)
        return f"{principal.strip()}/{parcela_padded.strip()}"

    return doc_str.strip()

def configurar_locale():
    """
    Tenta configurar o locale para o padrão brasileiro para formatação de moeda.
    """
    try:
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    except locale.Error:
        try:
            locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil.1252')
        except locale.Error:
            print("Aviso: Locale 'pt_BR' não encontrado. A formatação de moeda pode não funcionar.")

