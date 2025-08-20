import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import locale
from openpyxl.styles import NamedStyle
import sys
import subprocess
import threading
import queue
import importlib.util
import glob


# --- INSTRUÇÕES ---
# 1. Mantenha a estrutura de pastas como antes:
#    /Reconciliador
#    |-- main.py (este arquivo)
#    |-- /parsers/
#        |-- __init__.py
#        |-- parser_diamante.py
# 2. Execute este arquivo 'main.py'. Uma janela de seleção aparecerá primeiro.

# --- Funções do Core ---

def load_parsers():
    """
    Encontra e carrega dinamicamente todos os módulos de parser da pasta 'parsers'.
    Esta função agora é independente da classe da App para ser usada na inicialização.
    """
    parsers = {}
    parser_folder = "parsers"
    # Procura por todos os arquivos .py dentro da pasta 'parsers'
    for path in glob.glob(os.path.join(parser_folder, "parser_*.py")):
        module_name = os.path.basename(path)[:-3]
        try:
            spec = importlib.util.spec_from_file_location(f"{parser_folder}.{module_name}", path)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)

            # Pega os metadados do arquivo de parser
            if hasattr(module, 'PARSER_INFO'):
                info = module.PARSER_INFO
                parser_name = info['name']
                parser_function_name = info['function']
                parser_function = getattr(module, parser_function_name)
                parsers[parser_name] = parser_function
                print(f"Parser '{parser_name}' carregado com sucesso.")
        except Exception as e:
            print(f"Erro ao carregar o parser {module_name}: {e}")
    return parsers


def normalizar_documento(doc_str):
    """
    Normaliza o número do documento para um formato canônico.
    """
    if not isinstance(doc_str, str):
        return str(doc_str)

    match = re.search(r'(\d+[\/-]\d+)', doc_str)
    if not match:
        return doc_str.strip()

    doc_str_norm = match.group(1).replace('-', '/')
    partes = doc_str_norm.split('/')
    if len(partes) == 2:
        principal, parcela = partes
        parcela_padded = parcela.zfill(3)
        return f"{principal.strip()}/{parcela_padded.strip()}"

    return doc_str.strip()


def processar_nosso_relatorio(caminho_arquivo, thread_queue=None):
    """
    Lê e processa o nosso relatório (estruturado).
    """
    if thread_queue:
        thread_queue.put(("progress", 40, "Processando nosso relatório local..."))

    from parsers.parser_diamante import limpar_valor

    df = pd.read_csv(caminho_arquivo, encoding='latin-1', delimiter=',')
    df = df[['Documento', 'Sacado', 'Valor', 'Valor Pago']].rename(columns={
        'Valor': 'Valor_Original_Nosso', 'Valor Pago': 'Valor_Pago_Nosso', 'Sacado': 'Sacado_Nosso'
    })
    df['Valor_Original_Nosso'] = df['Valor_Original_Nosso'].apply(limpar_valor)
    df['Valor_Pago_Nosso'] = df['Valor_Pago_Nosso'].apply(limpar_valor)
    df['Documento'] = df['Documento'].astype(str).str.strip()

    if thread_queue:
        thread_queue.put(("progress", 50, "Nosso relatório processado."))
    return df


def gerar_relatorio_excel(df_fundo_agg, df_nosso_agg, df_comparativo, caminho_saida):
    """
    Gera um relatório Excel detalhado com a análise completa.
    """
    with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
        try:
            locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
        except locale.Error:
            try:
                locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil.1252')
            except locale.Error:
                print("Aviso: Locale 'pt_BR' não encontrado para formatação de moeda.")

        df_ambos = df_comparativo[df_comparativo['_merge'] == 'both'].copy()
        df_ambos['Juros/Taxas (Nosso)'] = df_ambos['Valor_Pago_Nosso'] - df_ambos['Valor_Original_Nosso']
        df_ambos['Diferenca_Liquida'] = df_ambos['Valor_Pago_Nosso'] - df_ambos['Valor_Fundo']

        df_so_fundo = df_comparativo[df_comparativo['_merge'] == 'left_only'].copy()
        df_so_nosso = df_comparativo[df_comparativo['_merge'] == 'right_only'].copy()
        if not df_so_nosso.empty:
            df_so_nosso['Juros/Taxas (Nosso)'] = df_so_nosso['Valor_Pago_Nosso'] - df_so_nosso['Valor_Original_Nosso']

        sumario_data = {
            'Métrica': [
                'Documentos Únicos (Fundo)', 'Valor Total (Fundo)', '',
                'Documentos Únicos (Nosso)', 'Valor Original (Nosso)', 'Valor Pago (Nosso)',
                'Total Juros/Taxas (Nosso)', '',
                'Documentos Correspondentes', 'Documentos com Diferença de Valor',
                'Valor Total das Diferenças Líquidas', '',
                'Documentos Apenas no Rel. Fundo', 'Valor Total:', '',
                'Documentos Apenas no Rel. Nosso', 'Valor Total:', '',
                'VALIDAÇÃO FINAL', 'Diferença Real (Total Pago Nosso - Total Fundo)',
                'Diferença Calculada (Soma das Discrepâncias)'
            ],
            'Valor': [
                df_fundo_agg['Documento'].nunique(), df_fundo_agg['Valor_Fundo'].sum(), None,
                df_nosso_agg['Documento'].nunique(), df_nosso_agg['Valor_Original_Nosso'].sum(),
                df_nosso_agg['Valor_Pago_Nosso'].sum(), df_nosso_agg.get('Juros/Taxas (Nosso)', 0).sum(), None,
                len(df_ambos), len(df_ambos[df_ambos['Diferenca_Liquida'].abs() > 0.01]),
                df_ambos['Diferenca_Liquida'].sum(), None,
                len(df_so_fundo), df_so_fundo['Valor_Fundo'].sum(), None,
                len(df_so_nosso), df_so_nosso['Valor_Pago_Nosso'].sum(), None,
                "SUCESSO" if abs((df_nosso_agg['Valor_Pago_Nosso'].sum() - df_fundo_agg['Valor_Fundo'].sum()) - (
                            df_ambos['Diferenca_Liquida'].sum() - df_so_fundo['Valor_Fundo'].sum() + df_so_nosso[
                        'Valor_Pago_Nosso'].sum())) < 0.01 else "FALHA",
                df_nosso_agg['Valor_Pago_Nosso'].sum() - df_fundo_agg['Valor_Fundo'].sum(),
                df_ambos['Diferenca_Liquida'].sum() - df_so_fundo['Valor_Fundo'].sum() + df_so_nosso[
                    'Valor_Pago_Nosso'].sum()
            ]
        }
        pd.DataFrame(sumario_data).to_excel(writer, sheet_name='Sumario_Conciliacao', index=False)

        df_ambos[df_ambos['Diferenca_Liquida'].abs() > 0.01].to_excel(writer, sheet_name='Diferencas_de_Valor',
                                                                      index=False)
        df_so_fundo.to_excel(writer, sheet_name='Apenas_Rel_Fundo', index=False)
        df_so_nosso.to_excel(writer, sheet_name='Apenas_Rel_Nosso', index=False)

        workbook = writer.book
        currency_style = NamedStyle(name='currency_br', number_format='R$ #,##0.00')
        if 'currency_br' not in workbook.style_names:
            workbook.add_named_style(currency_style)

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                ws.column_dimensions[column].width = adjusted_width
            if sheet_name != 'Sumario_Conciliacao':
                for col_letter in ['C', 'D', 'E', 'F', 'G']:
                    for cell in ws[col_letter][1:]:
                        cell.style = 'currency_br'


class FundSelectionDialog:
    """
    Uma janela de diálogo simples para forçar o usuário a escolher um fundo na inicialização.
    """

    def __init__(self, parent, parser_names):
        self.parent = parent
        self.top = tk.Toplevel(parent)
        self.top.title("Selecionar Fundo")
        self.top.geometry("350x150")
        self.top.resizable(False, False)
        # Garante que esta janela fique na frente e capture o foco
        self.top.grab_set()

        self.selected_fund = None
        self.parser_names = parser_names
        self.selected_parser_var = tk.StringVar(value=parser_names[0] if parser_names else "")

        main_frame = ttk.Frame(self.top, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="Selecione o Fundo de Investimento:", font=("-size", 10, "bold")).pack(pady=(0, 10))

        parser_menu = ttk.OptionMenu(main_frame, self.selected_parser_var, self.selected_parser_var.get(),
                                     *self.parser_names)
        parser_menu.pack(fill=tk.X, pady=5)

        confirm_button = ttk.Button(main_frame, text="Confirmar", command=self.on_confirm)
        confirm_button.pack(pady=10)

        # Centraliza a janela
        self.top.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() / 2) - (self.top.winfo_width() / 2)
        y = parent.winfo_rooty() + (parent.winfo_height() / 2) - (self.top.winfo_height() / 2)
        self.top.geometry(f"+{int(x)}+{int(y)}")

    def on_confirm(self):
        self.selected_fund = self.selected_parser_var.get()
        self.top.destroy()


class ReconciliationApp:
    def __init__(self, root, parsers, selected_fund):
        self.root = root
        self.root.title(f"Reconciliador de Relatórios - {selected_fund}")
        self.root.geometry("650x450")

        self.fundo_path = tk.StringVar()
        self.nosso_path = tk.StringVar()
        self.output_path = ""
        self.thread_queue = queue.Queue()
        self.is_running = False

        self.parsers = parsers
        self.selected_parser_name = selected_fund

        # --- Widgets ---
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Label mostrando o fundo selecionado (não pode ser alterado)
        ttk.Label(main_frame, text="Fundo Selecionado:").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Label(main_frame, text=self.selected_parser_name, font=("-size", 10, "bold")).grid(row=0, column=1,
                                                                                               columnspan=2, sticky="w",
                                                                                               pady=5)

        # Seleção de Arquivos
        ttk.Label(main_frame, text=f"Relatório ({self.selected_parser_name}):").grid(row=1, column=0, sticky="w",
                                                                                     pady=2)
        ttk.Entry(main_frame, textvariable=self.fundo_path, state="readonly").grid(row=1, column=1, sticky="ew", padx=5)
        ttk.Button(main_frame, text="Selecionar...", command=lambda: self.select_file(self.fundo_path,
                                                                                      f"Selecione o relatório do Fundo {self.selected_parser_name}")).grid(
            row=1, column=2)

        ttk.Label(main_frame, text="Nosso Relatório:").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Entry(main_frame, textvariable=self.nosso_path, state="readonly").grid(row=2, column=1, sticky="ew", padx=5)
        ttk.Button(main_frame, text="Selecionar...",
                   command=lambda: self.select_file(self.nosso_path, "Selecione o nosso relatório")).grid(row=2,
                                                                                                          column=2)

        # Controles
        self.generate_button = ttk.Button(main_frame, text="Gerar Relatório", command=self.start_reconciliation_thread,
                                          state="disabled")
        self.generate_button.grid(row=3, column=0, columnspan=3, pady=10)

        self.progress_bar = ttk.Progressbar(main_frame, orient="horizontal", mode="determinate")
        self.progress_bar.grid(row=4, column=0, columnspan=3, sticky="ew", pady=5)

        self.log_text = tk.Text(main_frame, height=8, state="disabled", bg="#f0f0f0")
        self.log_text.grid(row=5, column=0, columnspan=3, sticky="nsew")

        self.open_button = ttk.Button(main_frame, text="Abrir Relatório Gerado", command=self.open_report,
                                      state="disabled")
        self.open_button.grid(row=6, column=0, columnspan=3, pady=10)

        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)

    def select_file(self, path_var, title):
        filepath = filedialog.askopenfilename(parent=self.root, title=title)
        if filepath:
            path_var.set(filepath)
            self.check_paths()

    def check_paths(self):
        if self.fundo_path.get() and self.nosso_path.get():
            self.generate_button.config(state="normal")
        else:
            self.generate_button.config(state="disabled")

    def log_message(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

    def start_reconciliation_thread(self):
        if self.is_running: return
        self.is_running = True
        self.generate_button.config(state="disabled")
        self.open_button.config(state="disabled")
        self.progress_bar['value'] = 0
        self.log_text.config(state="normal")
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state="disabled")

        self.thread = threading.Thread(target=self.run_reconciliation)
        self.thread.daemon = True
        self.thread.start()
        self.root.after(100, self.check_thread)

    def check_thread(self):
        try:
            while True:
                message = self.thread_queue.get(block=False)
                msg_type, *payload = message
                if msg_type == "progress":
                    self.progress_bar['value'] = payload[0]
                    self.log_message(payload[1])
                elif msg_type == "done":
                    self.is_running = False
                    self.open_button.config(state="normal")
                    messagebox.showinfo("Sucesso", "Relatório de reconciliação gerado com sucesso!", parent=self.root)
                    return
                elif msg_type == "error":
                    self.is_running = False
                    messagebox.showerror("Erro", payload[0], parent=self.root)
                    return
        except queue.Empty:
            pass
        finally:
            if self.is_running:
                self.root.after(100, self.check_thread)
            else:
                self.generate_button.config(state="normal")
                self.progress_bar['value'] = 0

    def run_reconciliation(self):
        try:
            self.thread_queue.put(("progress", 10, "Iniciando análise..."))

            parser_funcao = self.parsers[self.selected_parser_name]
            df_fundo = parser_funcao(self.fundo_path.get(), self.thread_queue)

            df_nosso = processar_nosso_relatorio(self.nosso_path.get(), self.thread_queue)

            self.thread_queue.put(("progress", 55, "Normalizando documentos para comparação..."))
            df_fundo['Documento_Norm'] = df_fundo['Documento'].apply(normalizar_documento)
            df_nosso['Documento_Norm'] = df_nosso['Documento'].apply(normalizar_documento)

            self.thread_queue.put(("progress", 65, "Agregando valores por documento..."))
            df_fundo_agg = df_fundo.groupby('Documento_Norm').agg(Valor_Fundo=('Valor_Fundo', 'sum'),
                                                                  Sacado_Fundo=('Sacado_Fundo', 'first')).reset_index()
            df_nosso_agg = df_nosso.groupby('Documento_Norm').agg(Valor_Original_Nosso=('Valor_Original_Nosso', 'sum'),
                                                                  Valor_Pago_Nosso=('Valor_Pago_Nosso', 'sum'),
                                                                  Sacado_Nosso=('Sacado_Nosso', 'first')).reset_index()
            if not df_nosso_agg.empty:
                df_nosso_agg['Juros/Taxas (Nosso)'] = df_nosso_agg['Valor_Pago_Nosso'] - df_nosso_agg[
                    'Valor_Original_Nosso']

            self.thread_queue.put(("progress", 75, "Cruzando informações dos dois relatórios..."))
            df_fundo_agg.rename(columns={'Documento_Norm': 'Documento'}, inplace=True)
            df_nosso_agg.rename(columns={'Documento_Norm': 'Documento'}, inplace=True)
            df_comparativo = pd.merge(df_fundo_agg, df_nosso_agg, on='Documento', how='outer', indicator=True)

            self.thread_queue.put(("progress", 85, "Gerando planilha Excel..."))
            pasta_saida = os.path.dirname(self.fundo_path.get())
            self.output_path = os.path.join(pasta_saida, f"Relatorio_Conciliacao_{self.selected_parser_name}.xlsx")
            gerar_relatorio_excel(df_fundo_agg, df_nosso_agg, df_comparativo, self.output_path)

            self.thread_queue.put(("progress", 100, "Análise concluída."))
            self.thread_queue.put(("done",))
        except PermissionError:
            self.thread_queue.put(("error",
                                   f"Não foi possível salvar o arquivo '{os.path.basename(self.output_path)}'.\n\nVerifique se o arquivo não está aberto e tente novamente."))
        except Exception as e:
            self.thread_queue.put(("error", f"Ocorreu um erro inesperado:\n\n{e}"))

    def open_report(self):
        if self.output_path and os.path.exists(self.output_path):
            try:
                if sys.platform == "win32":
                    os.startfile(self.output_path)
                elif sys.platform == "darwin":
                    subprocess.call(['open', self.output_path])
                else:
                    subprocess.call(['xdg-open', self.output_path])
            except Exception as e:
                messagebox.showerror("Erro ao Abrir", f"Não foi possível abrir o arquivo automaticamente.\n\nErro: {e}",
                                     parent=self.root)
        else:
            messagebox.showwarning("Aviso", "O relatório ainda não foi gerado ou não foi encontrado.", parent=self.root)


if __name__ == "__main__":
    # Carrega os parsers disponíveis
    all_parsers = load_parsers()

    if not all_parsers:
        messagebox.showerror("Erro Crítico",
                             "Nenhum parser foi encontrado na pasta 'parsers'.\nA aplicação será encerrada.")
    else:
        # Cria uma janela raiz temporária para o diálogo de seleção
        selection_root = tk.Tk()
        selection_root.withdraw()  # Esconde a janela raiz

        # Abre o diálogo de seleção
        dialog = FundSelectionDialog(selection_root, list(all_parsers.keys()))
        selection_root.wait_window(dialog.top)  # Pausa a execução até o diálogo ser fechado

        # Se um fundo foi selecionado, inicia a aplicação principal
        if dialog.selected_fund:
            selected_fund_name = dialog.selected_fund
            selection_root.destroy()  # Destrói a janela temporária

            # Cria a janela principal da aplicação
            main_app_root = tk.Tk()
            app = ReconciliationApp(main_app_root, all_parsers, selected_fund_name)
            main_app_root.mainloop()
        else:
            # Se o usuário fechou o diálogo, encerra o programa
            selection_root.destroy()

