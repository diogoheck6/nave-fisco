import os
import pandas as pd
from models.excel_workbook import ExcelWorkbook

# Definindo o caminho absoluto utilizando uma string bruta
base_diretorio = r""

def gerar_excel(pintura_verde_entradas, pintura_verde_questor, entradas_sat, saidas_sat, entradas_questor, saidas_questor, only_in_sat, 
                only_in_sat_entradas, only_in_questor_entradas,
                fechamento_saidas, fechamento_entradas,
                teve_diferenca_saidas, teve_diferenca_entradas,
                cnpj_salvar, razao_social, only_in_questor_saidas, pintura_verde_questor_sat, intersecao_saidas, intersecao_entradas, 
                pintar_vermelho_saidas, pintar_vermelho_entradas, competencia):
    
    print('Iniciando geração de excel')

    # Converte para DataFrame se forem passados como lista
    try:
        if isinstance(intersecao_entradas, list):
            if intersecao_entradas:
                intersecao_entradas = pd.DataFrame(intersecao_entradas)
            else:
                intersecao_entradas = pd.DataFrame()
        
        if isinstance(intersecao_saidas, list):
            if intersecao_saidas:
                intersecao_saidas = pd.DataFrame(intersecao_saidas)
            else:
                intersecao_saidas = pd.DataFrame()
    except Exception as e:
        print(f"Erro ao converter listas para DataFrames: {e}")
        return None

    # print(f'intersecao_entradas DataFrame shape: {intersecao_entradas.shape}')
    # print(f'intersecao_saidas DataFrame shape: {intersecao_saidas.shape}')

    # Converte DataFrames para listas de dicionários
    list_intersecao_entradas = intersecao_entradas.to_dict(orient='records') if not intersecao_entradas.empty else []
    list_intersecao_saidas = intersecao_saidas.to_dict(orient='records') if not intersecao_saidas.empty else []

    try:
        filename = f"{razao_social} - {cnpj_salvar}.xlsx"
        caminho_completo = os.path.join(base_diretorio, filename)
        # print('caminho completo-->', caminho_completo)
        # print('competencia-->', competencia)
        excel = ExcelWorkbook(competencia, filename=caminho_completo)
        excel.create_new_workbook()
    except Exception as e:
        print(f"Erro ao criar workbook do Excel: {e}")
        return None

    # Processe cada adição de folha no bloco try-except
    try:
        # SAT Saídas
        print("Adicionando dados de SAT Saídas")
        excel.add_sheet("SAT Saídas")
        excel.populate_sheet('SAT Saídas', data=saidas_sat)    
        excel.centralize_non_numeric_texts('SAT Saídas')
        excel.format_numbers('SAT Saídas')
        excel.apply_header_filter('SAT Saídas')
        excel.apply_zebra_stripes('SAT Saídas')
        excel.auto_adjust_columns('SAT Saídas')

        # Questor Saídas
        print("Adicionando dados de Questor Saídas")
        excel.add_sheet("Questor Saídas")
        excel.populate_sheet('Questor Saídas', data=saidas_questor)
        excel.centralize_non_numeric_texts('Questor Saídas')
        excel.format_numbers('Questor Saídas')
        excel.apply_header_filter('Questor Saídas')
        excel.apply_zebra_stripes('Questor Saídas')
        excel.auto_adjust_columns('Questor Saídas')

        # SAT x Questor Saídas
        print("Adicionando comparação de SAT x Questor Saídas")
        excel.add_sheet("SAT x Questor Saídas")
        if not only_in_sat:
            excel.inserir_mensagem_sem_diferenca('SAT x Questor Saídas', 'saidas')
        else:
            excel.set_tab_color('SAT x Questor Saídas', 'red')
            excel.populate_sheet('SAT x Questor Saídas', data=only_in_sat)
            excel.centralize_non_numeric_texts('SAT x Questor Saídas')
            excel.format_numbers('SAT x Questor Saídas')
            excel.apply_header_filter('SAT x Questor Saídas')
            excel.apply_zebra_stripes('SAT x Questor Saídas')
            excel.auto_adjust_columns('SAT x Questor Saídas')

        # Questor Saídas x SAT
        print("Adicionando comparação de Questor Saídas x SAT")
        excel.add_sheet("Questor Saídas x SAT")
        if not only_in_questor_saidas:
            excel.inserir_mensagem_sem_diferenca('Questor Saídas x SAT', 'saidas questor')
        else:
            excel.set_tab_color('Questor Saídas x SAT', 'red')
            excel.populate_sheet('Questor Saídas x SAT', data=only_in_questor_saidas)
            excel.centralize_non_numeric_texts('Questor Saídas x SAT')
            excel.format_numbers('Questor Saídas x SAT')
            excel.apply_header_filter('Questor Saídas x SAT')
            excel.apply_zebra_stripes('Questor Saídas x SAT')
            excel.auto_adjust_columns('Questor Saídas x SAT')

        # Diferença Saídas
        print("Adicionando diferença de Saídas")
        excel.add_sheet("Diferença Saídas")
        if not list_intersecao_saidas:
            excel.inserir_mensagem_sem_diferenca('Diferença Saídas', 'msg_dif_saidas')
        else:
            excel.set_tab_color('Diferença Saídas', 'red')
            excel.populate_sheet("Diferença Saídas", data=list_intersecao_saidas)
            excel.centralize_non_numeric_texts("Diferença Saídas")
            excel.format_numbers("Diferença Saídas")
            excel.apply_header_filter('Diferença Saídas')
            excel.apply_zebra_stripes('Diferença Saídas')
            excel.auto_adjust_columns("Diferença Saídas")

        # Fechamento Saídas
        print("Adicionando fechamento de Saídas")
        excel.add_sheet("Fechamento Saídas")
        excel.populate_sheet('Fechamento Saídas', data=fechamento_saidas)
        excel.centralize_non_numeric_texts('Fechamento Saídas')
        excel.format_numbers('Fechamento Saídas')
        if teve_diferenca_saidas:
            excel.set_tab_color("Fechamento Saídas", 'red')
        excel.auto_adjust_columns('Fechamento Saídas')

        # SAT Entradas
        print("Adicionando dados de SAT Entradas")
        excel.add_sheet("SAT Entradas")
        excel.populate_sheet('SAT Entradas', data=entradas_sat)
        excel.centralize_non_numeric_texts('SAT Entradas')
        excel.format_numbers('SAT Entradas')
        excel.apply_header_filter('SAT Entradas')
        excel.apply_zebra_stripes('SAT Entradas')
        excel.auto_adjust_columns('SAT Entradas')

        # Questor Entradas
        print("Adicionando dados de Questor Entradas")
        excel.add_sheet("Questor Entradas")
        excel.populate_sheet('Questor Entradas', data=entradas_questor)
        excel.centralize_non_numeric_texts('Questor Entradas')
        excel.format_numbers('Questor Entradas')
        excel.apply_header_filter('Questor Entradas')
        excel.apply_zebra_stripes('Questor Entradas')
        excel.auto_adjust_columns('Questor Entradas')

        # SAT x Questor Entradas
        print("Adicionando comparação de SAT x Questor Entradas")
        excel.add_sheet("SAT x Questor Entradas")
        if not only_in_sat_entradas:
            excel.inserir_mensagem_sem_diferenca("SAT x Questor Entradas", 'entrada')
        else:
            excel.populate_sheet('SAT x Questor Entradas', data=only_in_sat_entradas)
            excel.centralize_non_numeric_texts('SAT x Questor Entradas')
            excel.format_numbers('SAT x Questor Entradas')
            excel.apply_header_filter('SAT x Questor Entradas')
            excel.apply_zebra_stripes('SAT x Questor Entradas')
            excel.auto_adjust_columns('SAT x Questor Entradas')
            if not pintura_verde_entradas:
                excel.set_tab_color('SAT x Questor Entradas', 'red')

        # Questor Entradas x SAT
        print("Adicionando comparação de Questor Entradas x SAT")
        excel.add_sheet("Questor Entradas x SAT")
        if not only_in_questor_entradas:
            excel.inserir_mensagem_sem_diferenca("Questor Entradas x SAT", 'questor')
        else:
            excel.populate_sheet('Questor Entradas x SAT', data=only_in_questor_entradas)
            excel.centralize_non_numeric_texts('Questor Entradas x SAT')
            excel.format_numbers('Questor Entradas x SAT')
            excel.apply_header_filter('Questor Entradas x SAT')
            excel.apply_zebra_stripes('Questor Entradas x SAT')
            excel.auto_adjust_columns('Questor Entradas x SAT')
            if not pintura_verde_questor:
                excel.set_tab_color('Questor Entradas x SAT', 'red')

        # Diferença Entradas
        print("Adicionando diferença de Entradas")
        excel.add_sheet("Diferença Entradas")
        if not list_intersecao_entradas:
            excel.inserir_mensagem_sem_diferenca('Diferença Entradas', 'msg_dif_entradas')
        else:
            excel.set_tab_color('Diferença Entradas', 'red')
            excel.populate_sheet("Diferença Entradas", data=list_intersecao_entradas)
            excel.centralize_non_numeric_texts("Diferença Entradas")
            excel.format_numbers("Diferença Entradas")
            excel.apply_header_filter('Diferença Entradas')
            excel.apply_zebra_stripes('Diferença Entradas')
            excel.auto_adjust_columns("Diferença Entradas")

        # Fechamento Entradas
        print("Adicionando fechamento de Entradas")
        excel.add_sheet("Fechamento Entradas")
        excel.populate_sheet('Fechamento Entradas', data=fechamento_entradas)
        excel.centralize_non_numeric_texts('Fechamento Entradas')
        excel.format_numbers('Fechamento Entradas')
        if teve_diferenca_entradas:
            excel.set_tab_color("Fechamento Entradas", 'red')
        excel.auto_adjust_columns('Fechamento Entradas')

        excel.save()

        print("Geração de excel concluída com sucesso")
        
    except Exception as e:
        print(f"Erro ao processar e criar excel: {e}")
        return None

    return caminho_completo