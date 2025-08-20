# main.py
import os
import sys
import subprocess
import threading
import queue
import importlib
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# --- Correção para o ImportError ---
project_root = os.path.dirname(os.path.abspath(__file__))
if project_root not in sys.path:
    sys.path.insert(0, project_root)
# --- Fim da Correção ---

# Importa as funções dos nossos módulos
import pandas as pd
from utils import normalizar_documento
from excel_generator import gerar_relatorio_excel
# Importa os parsers específicos
from parsers import nosso_relatorio_parser


# --- Lógica para carregar parsers dinamicamente ---
def carregar_parsers_fundos():
    """
    Encontra todos os parsers de fundos na pasta 'parsers'.
    Retorna um dicionário com o nome do fundo e o módulo do parser.
    """
    parsers = {}
    parser_dir = os.path.join(project_root, "parsers")
    if not os.path.isdir(parser_dir):
        return parsers

    arquivos_parser = [f for f in os.listdir(parser_dir) if
                       f.endswith('.py') and not f.startswith('_') and 'nosso' not in f]

    for arquivo in arquivos_parser:
        nome_modulo = arquivo[:-3]
        nome_fundo = nome_modulo.replace("_parser", "").capitalize()
        try:
            modulo = importlib.import_module(f"parsers.{nome_modulo}")
            if hasattr(modulo, 'processar'):
                parsers[nome_fundo] = modulo
        except ImportError as e:
            print(f"Erro ao importar o parser '{nome_modulo}': {e}")

    return parsers


class ReconciliationApp:
    def __init__(self, root, parsers_disponiveis):
        self.root = root
        self.root.title("Reconciliador de Relatórios Contábeis")
        self.root.geometry("650x450")

        self.parsers = parsers_disponiveis
        self.nosso_path = tk.StringVar()
        self.fundo_path = tk.StringVar()
        self.fundo_selecionado = tk.StringVar()

        self.output_path = ""
        self.thread_queue = queue.Queue()
        self.is_running = False

        # --- Widgets ---
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Seleção do Nosso Relatório
        ttk.Label(main_frame, text="1. Nosso Relatório (interno):").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Entry(main_frame, textvariable=self.nosso_path, state="readonly").grid(row=0, column=1, columnspan=2,
                                                                                   sticky="ew", padx=5)
        ttk.Button(main_frame, text="Selecionar...",
                   command=lambda: self.select_file(self.nosso_path, "Selecione o Nosso Relatório")).grid(row=0,
                                                                                                          column=3)

        # Seleção do Fundo
        ttk.Label(main_frame, text="2. Selecione o Fundo:").grid(row=1, column=0, sticky="w", pady=2)
        self.fundo_combo = ttk.Combobox(main_frame, textvariable=self.fundo_selecionado, state="readonly")
        self.fundo_combo['values'] = list(self.parsers.keys())
        if self.fundo_combo['values']:
            self.fundo_combo.current(0)
        self.fundo_combo.grid(row=1, column=1, columnspan=2, sticky="ew", padx=5)

        # Seleção do Relatório do Fundo
        ttk.Label(main_frame, text="3. Relatório do Fundo (externo):").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Entry(main_frame, textvariable=self.fundo_path, state="readonly").grid(row=2, column=1, columnspan=2,
                                                                                   sticky="ew", padx=5)
        ttk.Button(main_frame, text="Selecionar...",
                   command=lambda: self.select_file(self.fundo_path, "Selecione o Relatório do Fundo")).grid(row=2,
                                                                                                             column=3)

        # Controles
        self.generate_button = ttk.Button(main_frame, text="Gerar Relatório", command=self.start_reconciliation_thread,
                                          state="disabled")
        self.generate_button.grid(row=3, column=0, columnspan=4, pady=15)

        self.progress_bar = ttk.Progressbar(main_frame, orient="horizontal", mode="determinate")
        self.progress_bar.grid(row=4, column=0, columnspan=4, sticky="ew", pady=5)

        self.log_text = tk.Text(main_frame, height=8, state="disabled", bg="#f0f0f0", wrap="word")
        self.log_text.grid(row=5, column=0, columnspan=4, sticky="nsew")

        self.open_button = ttk.Button(main_frame, text="Abrir Relatório Gerado", command=self.open_report,
                                      state="disabled")
        self.open_button.grid(row=6, column=0, columnspan=4, pady=10)

        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)

    def select_file(self, path_var, title):
        """
        Abre uma janela para selecionar um arquivo CSV.
        """
        filetypes = [("CSV files", "*.csv"), ("All files", "*.*")]
        filepath = filedialog.askopenfilename(parent=self.root, title=title, filetypes=filetypes)
        if filepath:
            path_var.set(filepath)
            self.check_paths()

    def check_paths(self):
        if self.nosso_path.get() and self.fundo_path.get() and self.fundo_selecionado.get():
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
                msg_type, msg_data = message

                if msg_type == "progress":
                    progress, text = msg_data
                    self.progress_bar['value'] = progress
                    self.log_message(text)
                elif msg_type == "done":
                    self.is_running = False
                    self.open_button.config(state="normal")
                    messagebox.showinfo("Sucesso", "Relatório de reconciliação gerado com sucesso!", parent=self.root)
                    return
                elif msg_type == "error":
                    self.is_running = False
                    messagebox.showerror("Erro", msg_data, parent=self.root)
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
            self.thread_queue.put(("progress", (10, "Processando nosso relatório...")))
            df_nosso = nosso_relatorio_parser.processar(self.nosso_path.get())

            nome_fundo_selecionado = self.fundo_selecionado.get()
            parser_modulo = self.parsers[nome_fundo_selecionado]
            self.thread_queue.put(("progress", (30, f"Processando relatório do fundo '{nome_fundo_selecionado}'...")))
            df_fundo = parser_modulo.processar(self.fundo_path.get())

            self.thread_queue.put(("progress", (50, "Normalizando documentos...")))
            df_nosso['Documento_Norm'] = df_nosso['Documento'].apply(normalizar_documento)
            df_fundo['Documento_Norm'] = df_fundo['Documento'].apply(normalizar_documento)

            self.thread_queue.put(("progress", (60, "Agregando valores por documento...")))
            df_nosso_agg = df_nosso.groupby('Documento_Norm').agg(Valor_Nosso=('Valor_Nosso', 'sum'),
                                                                  Sacado_Nosso=('Sacado_Nosso', 'first')).reset_index()
            df_fundo_agg = df_fundo.groupby('Documento_Norm').agg(Valor_Fundo_Original=('Valor_Fundo_Original', 'sum'),
                                                                  Valor_Fundo_Pago=('Valor_Fundo_Pago', 'sum'),
                                                                  Sacado_Fundo=('Sacado_Fundo', 'first')).reset_index()
            df_fundo_agg['Juros/Taxas (Fundo)'] = df_fundo_agg['Valor_Fundo_Pago'] - df_fundo_agg[
                'Valor_Fundo_Original']

            self.thread_queue.put(("progress", (70, "Cruzando informações dos relatórios...")))
            df_nosso_agg.rename(columns={'Documento_Norm': 'Documento'}, inplace=True)
            df_fundo_agg.rename(columns={'Documento_Norm': 'Documento'}, inplace=True)
            df_comparativo = pd.merge(df_nosso_agg, df_fundo_agg, on='Documento', how='outer', indicator=True)

            self.thread_queue.put(("progress", (80, "Gerando planilha Excel...")))
            pasta_saida = os.path.dirname(self.nosso_path.get())
            self.output_path = os.path.join(pasta_saida, f"Relatorio_Conciliacao_{nome_fundo_selecionado}.xlsx")
            gerar_relatorio_excel(df_nosso_agg, df_fundo_agg, df_comparativo, self.output_path)

            self.thread_queue.put(("progress", (100, "Análise concluída com sucesso!")))
            self.thread_queue.put(("done", None))
        except (PermissionError, ValueError) as e:
            self.thread_queue.put(("error", f"Erro ao processar arquivo:\n\n{e}"))
        except Exception as e:
            import traceback
            traceback.print_exc()
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
    if not os.path.isdir("parsers"):
        os.makedirs("parsers")

    available_parsers = carregar_parsers_fundos()

    app_root = tk.Tk()
    if not available_parsers:
        messagebox.showwarning("Nenhum Parser Encontrado",
                               "Nenhum arquivo de parser de fundo foi encontrado na pasta 'parsers'.\n\n"
                               "Crie arquivos .py para cada fundo dentro da pasta 'parsers' para continuar.",
                               parent=app_root)
        app_root.destroy()
    else:
        app = ReconciliationApp(app_root, available_parsers)
        app_root.mainloop()
