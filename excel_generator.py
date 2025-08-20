# excel_generator.py
import pandas as pd
from openpyxl.styles import NamedStyle
from utils import configurar_locale  # <-- MUDANÇA AQUI


def gerar_relatorio_excel(df_nosso_agg, df_fundo_agg, df_comparativo, caminho_saida):
    """
    Gera um relatório Excel detalhado com a análise completa.
    """
    configurar_locale()

    with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
        # --- Cálculos e Preparação dos DataFrames ---
        df_ambos = df_comparativo[df_comparativo['_merge'] == 'both'].copy()
        # Renomeia colunas para clareza no relatório final
        df_ambos.rename(columns={
            'Valor_Fundo_Original': 'Valor Original (Fundo)',
            'Valor_Fundo_Pago': 'Valor Pago (Fundo)',
            'Sacado_Fundo': 'Sacado (Fundo)'
        }, inplace=True)

        df_ambos['Juros/Taxas (Fundo)'] = df_ambos['Valor Pago (Fundo)'] - df_ambos['Valor Original (Fundo)']
        df_ambos['Diferenca_Liquida'] = df_ambos['Valor Pago (Fundo)'] - df_ambos['Valor_Nosso']

        df_so_nosso = df_comparativo[df_comparativo['_merge'] == 'left_only'].copy()

        df_so_fundo = df_comparativo[df_comparativo['_merge'] == 'right_only'].copy()
        df_so_fundo.rename(columns={
            'Valor_Fundo_Original': 'Valor Original (Fundo)',
            'Valor_Fundo_Pago': 'Valor Pago (Fundo)',
            'Sacado_Fundo': 'Sacado (Fundo)'
        }, inplace=True)
        df_so_fundo['Juros/Taxas (Fundo)'] = df_so_fundo['Valor Pago (Fundo)'] - df_so_fundo['Valor Original (Fundo)']

        # --- Aba de Sumário ---
        total_pago_fundo = df_fundo_agg['Valor_Fundo_Pago'].sum()
        total_nosso = df_nosso_agg['Valor_Nosso'].sum()
        diff_real = total_pago_fundo - total_nosso
        diff_calculada = df_ambos['Diferenca_Liquida'].sum() - df_so_nosso['Valor_Nosso'].sum() + df_so_fundo[
            'Valor Pago (Fundo)'].sum()

        sumario_data = {
            'Métrica': [
                'Documentos Únicos (Nosso Relatório)', 'Valor Total (Nosso)', '',
                'Documentos Únicos (Rel. Fundo)', 'Valor Original (Fundo)', 'Valor Pago (Fundo)',
                'Total Juros/Taxas (Fundo)', '',
                'Documentos Correspondentes', 'Documentos com Diferença de Valor',
                'Valor Total das Diferenças Líquidas', '',
                'Documentos Apenas no Nosso Relatório', 'Valor Total (Apenas Nosso)', '',
                'Documentos Apenas no Rel. Fundo', 'Valor Total (Apenas Fundo)', '',
                'VALIDAÇÃO FINAL', 'Diferença Real (Total Pago Fundo - Total Nosso)',
                'Diferença Calculada (Soma das Discrepâncias)'
            ],
            'Valor': [
                df_nosso_agg['Documento'].nunique(), total_nosso, None,
                df_fundo_agg['Documento'].nunique(), df_fundo_agg['Valor_Fundo_Original'].sum(), total_pago_fundo,
                df_fundo_agg['Juros/Taxas (Fundo)'].sum(), None,
                len(df_ambos), len(df_ambos[df_ambos['Diferenca_Liquida'].abs() > 0.01]),
                df_ambos['Diferenca_Liquida'].sum(), None,
                len(df_so_nosso), df_so_nosso['Valor_Nosso'].sum(), None,
                len(df_so_fundo), df_so_fundo['Valor Pago (Fundo)'].sum(), None,
                "SUCESSO" if abs(diff_real - diff_calculada) < 0.01 else "FALHA",
                diff_real,
                diff_calculada
            ]
        }
        pd.DataFrame(sumario_data).to_excel(writer, sheet_name='Sumario_Conciliacao', index=False)

        # --- Abas de Detalhes ---
        colunas_diferenca = ['Documento', 'Valor Original (Fundo)', 'Juros/Taxas (Fundo)', 'Valor Pago (Fundo)',
                             'Valor_Nosso', 'Diferenca_Liquida']
        df_ambos[df_ambos['Diferenca_Liquida'].abs() > 0.01][colunas_diferenca].to_excel(writer,
                                                                                         sheet_name='Diferencas_de_Valor',
                                                                                         index=False)

        df_so_nosso[['Documento', 'Sacado_Nosso', 'Valor_Nosso']].to_excel(writer,
                                                                           sheet_name='Apenas_no_Nosso_Relatorio',
                                                                           index=False)

        colunas_so_fundo = ['Documento', 'Sacado (Fundo)', 'Valor Original (Fundo)', 'Juros/Taxas (Fundo)',
                            'Valor Pago (Fundo)']
        df_so_fundo[colunas_so_fundo].to_excel(writer, sheet_name='Apenas_no_Rel_Fundo', index=False)

        # --- Formatação e Auto-ajuste ---
        workbook = writer.book
        currency_style = NamedStyle(name='currency_br', number_format='R$ #,##0.00')
        integer_style = NamedStyle(name='integer', number_format='#,##0')
        if 'currency_br' not in workbook.style_names:
            workbook.add_named_style(currency_style)
        if 'integer' not in workbook.style_names:
            workbook.add_named_style(integer_style)

        sheets_to_format = {
            'Diferencas_de_Valor': ['B', 'C', 'D', 'E', 'F'],
            'Apenas_no_Nosso_Relatorio': ['C'],
            'Apenas_no_Rel_Fundo': ['C', 'D', 'E']
        }
        for sheet_name, cols in sheets_to_format.items():
            ws = writer.sheets[sheet_name]
            ws.auto_filter.ref = ws.dimensions
            for col_letter in cols:
                for cell in ws[col_letter][1:]:
                    cell.style = 'currency_br'

        ws_sumario = writer.sheets['Sumario_Conciliacao']
        for cell in ws_sumario['B'][1:]:
            metric_cell = ws_sumario[f'A{cell.row}']
            if "Documentos" in str(metric_cell.value) or "Correspondentes" in str(metric_cell.value):
                cell.style = 'integer'
            elif isinstance(cell.value, (int, float)):
                cell.style = 'currency_br'

        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column_cells in worksheet.columns:
                max_length = max((len(str(cell.value)) for cell in column_cells if cell.value is not None), default=0)
                worksheet.column_dimensions[column_cells[0].column_letter].width = (max_length + 2)
