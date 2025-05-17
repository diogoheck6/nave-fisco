import os
import pandas as pd
from datetime import datetime
from decimal import Decimal
from firebird_connection.connection import PostgresConnection
from queries.lctofissai_query import get_lctofissai_query
from queries.lctofisent_query import get_lctofisent_query
from controller.read_files_html import ServiceLerHtmlCTE
from controller.ler_excel_cte import ServiceLerExcelCTE
import math


# FUNÇÃO DE UTILIDADE
def verificar_condicoes_para_pintura(only_in_sat_entradas, cnpj):
    # Verificar que cada entrada atende pelo menos uma das condições
    for item in only_in_sat_entradas:
        sit_cond = item.get('SIT', '').lower() == 'cancelado'  # Transformar para minúsculas
        es_cond = item.get('E/S') == 'E' and cnpj not in item.get('CHAVE', '')

        # Se uma nota não atender a nenhuma condição, retorna False
        if not (sit_cond or es_cond):
            return False

    # Se todas as notas atendem pelo menos uma condição, retorna True
    return True

def verificar_condicoes_para_pintura_questor(only_in_sat_entradas):
    # Verificar se existe alguma entrada na lista
    if only_in_sat_entradas:
        return False
    
    # Retornar True apenas se não houver nenhuma entrada
    return True

def sanitize_difference(diff, tolerance=1e-10):
    """Sanitiza a diferença para lidar com ruído de precisão de ponto flutuante"""
    if math.isnan(diff) or abs(diff) < tolerance:
        return 0.0
    return diff


# def sanitize_difference(diff, decimal_places=2):
#     """Sanitiza a diferença para lidar com ruído de precisão de ponto flutuante,
#     arredondando para o número especificado de casas decimais."""
#     if math.isnan(diff):
#         return 0.0
#     # Arredonda a diferença para o número especificado de casas decimais
#     return round(diff, decimal_places)


#  FUNÇÕES DE PROCESSAMENTO

def agregar_por_chave(dados):
    # Define colunas a sumariar e modos de agregação
    colunas_info = ['DATA', 'REF', 'MOD', 'TIPO', 'E/S', 'SIT', 'CHAVE', 'SERIE', 'FORNEC', 'CFOP','NUM NF']
    colunas_soma = ['VLR NF', 'VLR BC ICMS', 'VLR ICMS']

    def join_cfop(values):
        return ' | '.join(sorted(set(str(v) for v in values)))  # Converte cada valor para string antes de juntar

    return dados.groupby('CHAVE', as_index=False).agg({
        **{col: 'first' for col in colunas_info if col in dados.columns},
        'CFOP': join_cfop if 'CFOP' in dados.columns else 'first',  # Aplica a função para concatenar CFOPs
        **{col: 'sum' for col in colunas_soma if col in dados.columns},
        'VLR IPI': 'first' if 'VLR IPI' in dados.columns else 'sum'
    })


def realizar_intersecao(sat_agrupado, questor_agrupado):
    return pd.merge(sat_agrupado, questor_agrupado, on='CHAVE', suffixes=('_SAT', '_QUESTOR'))

def calcular_diferencas(intersecao):
    tolerancia = 1e-5
    for coluna in ['VLR NF', 'VLR BC ICMS', 'VLR ICMS', 'VLR IPI']:
        intersecao[f'DIF {coluna}'] = intersecao[f'{coluna}_SAT'] - intersecao[f'{coluna}_QUESTOR']
        intersecao[f'DIF {coluna} relevante'] = intersecao[f'DIF {coluna}'].abs() > tolerancia
    return intersecao


def inicializar_agrupamento(df):

    if not df.empty:
        return agregar_por_chave(df)
    else:
        return pd.DataFrame(columns=['CHAVE']) 


def processar_dados(entradas_sat, entradas_questor):

    sat_df = pd.DataFrame(entradas_sat)
    questor_df = pd.DataFrame(entradas_questor)    

    sat_agrupado = inicializar_agrupamento(sat_df)
    questor_agrupado = inicializar_agrupamento(questor_df)

    if sat_agrupado.empty or questor_agrupado.empty:
        return [], False, sat_agrupado, questor_agrupado

    # Interseção e cálculos de diferença apenas quando ambos os DataFrames têm dados
    intersecao = realizar_intersecao(sat_agrupado, questor_agrupado)
    intersecao_calculada = calcular_diferencas(intersecao)

    intersecao_filtrada = intersecao_calculada[
        intersecao_calculada[['DIF VLR NF relevante', 'DIF VLR BC ICMS relevante', 'DIF VLR ICMS relevante', 'DIF VLR IPI relevante']].any(axis=1)
    ].copy()
    pintar_vermelho = not intersecao_filtrada.empty

    colunas_reordenadas = [col for col in [
        'DATA_SAT', 'TIPO_SAT', 'CHAVE', 'FORNEC_QUESTOR', 'CFOP_QUESTOR', 'NUM NF_SAT',
        'VLR NF_SAT', 'VLR NF_QUESTOR', 'DIF VLR NF',
        'VLR BC ICMS_SAT', 'VLR BC ICMS_QUESTOR', 'DIF VLR BC ICMS',
        'VLR ICMS_SAT', 'VLR ICMS_QUESTOR', 'DIF VLR ICMS',
        'VLR IPI_SAT', 'VLR IPI_QUESTOR', 'DIF VLR IPI'
    ] if col in intersecao_filtrada.columns]

    return intersecao_filtrada[colunas_reordenadas].to_dict(orient='records'), pintar_vermelho, sat_agrupado, questor_agrupado


def realizar_intersecao(sat_agrupado, questor_agrupado):
    if sat_agrupado.empty or questor_agrupado.empty:
        return pd.DataFrame(columns=['CHAVE'])
    return pd.merge(sat_agrupado, questor_agrupado, on='CHAVE', suffixes=('_SAT', '_QUESTOR'))


def read_files_and_query(mais_de_500_ctes, cnpj, start_date, end_date):

    sanitized_cnpj = cnpj.replace('.', '').replace('/', '').replace('-', '')

    entradas_sat, saidas_sat = read_files_xlsx(sanitized_cnpj)
    
    html_data_entradas = []
    html_data_saidas = []

    if mais_de_500_ctes:
        ctes = ServiceLerHtmlCTE()
        html_data_entradas, html_data_saidas = ctes.ler_arquivos_html(sanitized_cnpj)

    if not mais_de_500_ctes:
        cte_normal = ServiceLerExcelCTE()
        html_data_entradas, html_data_saidas = cte_normal.ler_arquivos_excel(sanitized_cnpj)

    entradas_sat.extend(html_data_entradas)
    saidas_sat.extend(html_data_saidas)

    entradas_questor_rows = consultar_entradas_questor(cnpj, start_date, end_date)
    saidas_questor_rows = fazer_consulta_sql_saidas(cnpj, start_date, end_date)

    columns_map_entradas = {
        "DATAEMISSAO": "DATA",
        "CHAVENFEENT": "CHAVE",
        "NUMERONF": "NUM NF",
        "ESPECIENF": "TIPO",
        "SERIENF": "SERIE",
        'NOMEPESSOA': "FORNEC",
        'CODIGOCFOP': "CFOP",
        "VALORCONTABILIMPOSTO": "VLR NF",
        "BASECALCULOIMPOSTO": "VLR BC ICMS",
        "VALORIMPOSTO": "VLR ICMS",
        "VALORIPI": "VLR IPI"
    }

    columns_map_saidas = {
        "DATALCTOFIS": "DATA",
        "CHAVENFESAI": "CHAVE",
        "NUMERONF": "NUM NF",
        "ESPECIENF": "TIPO",
        "SERIENF": "SERIE",
        'NOMEPESSOA': "FORNEC",
        'CODIGOCFOP': "CFOP",
        "VALORCONTABILIMPOSTO": "VLR NF",
        "BASECALCULOIMPOSTO": "VLR BC ICMS",
        "VALORIMPOSTO": "VLR ICMS",
        "VALORIPI": "VLR IPI"
    }

    def process_questor_data(interesse_rows, columns_map):
        try:
            # Obtendo o cabeçalho da consulta
            header = [description[0].upper() for description in interesse_rows['cursor'].description]
            specified_header = [col for col in header if col in columns_map.keys()]
            converted_rows = []

            # Iteração sobre as linhas do resultado da consulta
            for row in interesse_rows['rows']:
                row_dict = {}

                for col in specified_header:
                    idx = header.index(col)
                    value = row[idx]
                    new_col = columns_map[col]

                    if new_col == 'VLR NF' or new_col == 'VLR BC ICMS' or new_col == 'VLR ICMS' or new_col == 'VLR IPI':
                        value = round(float(value), 2) if value is not None else 0.0
                    elif new_col == 'NUM NF':
                        value = str(value)
                    elif new_col == 'DATA':
                        if value is not None:
                            date_value = pd.to_datetime(value, errors='coerce')
                            if pd.notnull(date_value):
                                value = date_value.strftime('%d/%m/%Y')  # Formato brasileiro
                    row_dict[new_col] = value

                converted_rows.append(row_dict)

            return converted_rows

        except Exception as e:
            print(f"Erro ao processar resultados da consulta: {e}")
            return []

    entradas_questor = process_questor_data(entradas_questor_rows, columns_map_entradas)
    saidas_questor = process_questor_data(saidas_questor_rows, columns_map_saidas)


    entradas_sat_agrupadas = inicializar_agrupamento(pd.DataFrame(entradas_sat))
    entradas_questor_agrupadas = inicializar_agrupamento(pd.DataFrame(entradas_questor))
    saidas_sat_agrupadas = inicializar_agrupamento(pd.DataFrame(saidas_sat))
    saidas_questor_agrupadas = inicializar_agrupamento(pd.DataFrame(saidas_questor))

    # Verifique se existem dados antes de processá-los
    if not entradas_sat_agrupadas.empty and not entradas_questor_agrupadas.empty:
        intersecao_entradas, pintar_vermelho_entradas, entradas_sat_agrupadas, entradas_questor_agrupadas = processar_dados(entradas_sat_agrupadas, entradas_questor_agrupadas)
    else:
        intersecao_entradas, pintar_vermelho_entradas = pd.DataFrame(), False

    if not saidas_sat_agrupadas.empty and not saidas_questor_agrupadas.empty:
        intersecao_saidas, pintar_vermelho_saidas, saidas_sat_agrupadas, saidas_questor_agrupadas = processar_dados(saidas_sat_agrupadas, saidas_questor_agrupadas)
    else:
        intersecao_saidas, pintar_vermelho_saidas = pd.DataFrame(), False


    # Lógica de dicionários e operações
    sat_dict_agrupada = {row['CHAVE']: row for _, row in saidas_sat_agrupadas.iterrows()} if not saidas_sat_agrupadas.empty else {}
    questor_dict_agrupada = {row['CHAVE']: row for _, row in saidas_questor_agrupadas.iterrows()} if not saidas_questor_agrupadas.empty else {}

    only_in_sat = [sat_dict_agrupada[chave] for chave in sat_dict_agrupada if chave not in questor_dict_agrupada]
    only_in_questor_saidas = [questor_dict_agrupada[chave] for chave in questor_dict_agrupada if chave not in sat_dict_agrupada]
    pintura_verde_questor_sat = verificar_condicoes_para_pintura(only_in_questor_saidas, sanitized_cnpj)

    sat_dict_entradas_agrupada = {row['CHAVE']: row for _, row in entradas_sat_agrupadas.iterrows()} if not entradas_sat_agrupadas.empty else {}
    questor_dict_entradas_agrupada = {row['CHAVE']: row for _, row in entradas_questor_agrupadas.iterrows()} if not entradas_questor_agrupadas.empty else {}

    only_in_questor_entradas = [questor_dict_entradas_agrupada[chave] for chave in questor_dict_entradas_agrupada if chave not in sat_dict_entradas_agrupada]
    only_in_sat_entradas = [sat_dict_entradas_agrupada[chave] for chave in sat_dict_entradas_agrupada if chave not in questor_dict_entradas_agrupada]

    pintura_verde_entradas = verificar_condicoes_para_pintura(only_in_sat_entradas, sanitized_cnpj)
    pintura_verde_questor = verificar_condicoes_para_pintura(only_in_questor_entradas, sanitized_cnpj)

    fechamento_saidas, teve_diferenca_saidas = calcular_fechamento_saidas(saidas_sat, saidas_questor)
    fechamento_entradas, teve_diferenca_entradas = calcular_fechamento_entradas(entradas_sat, entradas_questor, sanitized_cnpj)

    return entradas_sat, saidas_sat, entradas_questor, saidas_questor, only_in_sat, only_in_sat_entradas, \
           only_in_questor_entradas, fechamento_saidas, fechamento_entradas, teve_diferenca_saidas, \
           teve_diferenca_entradas, pintura_verde_entradas, pintura_verde_questor, only_in_questor_saidas, pintura_verde_questor_sat, \
           intersecao_saidas, intersecao_entradas, pintar_vermelho_saidas, pintar_vermelho_entradas



def calcular_fechamento_entradas(entradas_sat, entradas_questor, cnpj):
    def soma_unico_ipi_por_chave(entradas):
        vistos = set()
        total_ipi = 0.0
        for entrada in entradas:
            chave = entrada.get('CHAVE')
            ipi = entrada.get('VLR IPI', 0.0)
            if chave not in vistos:
                total_ipi += ipi
                vistos.add(chave)
        return total_ipi

    entradas_sat_relevantes = [
        item for item in entradas_sat
        if not (item.get('SIT') == 'Cancelado' or (item.get('E/S') == 'E' and cnpj not in item.get('CHAVE', '')))
    ]

    total_sat = sum(item.get('VLR NF', 0.0) for item in entradas_sat_relevantes)
    bc_icms_sat = sum(item.get('VLR BC ICMS', 0.0) for item in entradas_sat_relevantes)
    icms_sat = sum(item.get('VLR ICMS', 0.0) for item in entradas_sat_relevantes)
    ipi_sat = soma_unico_ipi_por_chave(entradas_sat_relevantes)

    total_questor = sum(item.get('VLR NF', 0.0) for item in entradas_questor)
    bc_icms_questor = sum(item.get('VLR BC ICMS', 0.0) for item in entradas_questor)
    icms_questor = sum(item.get('VLR ICMS', 0.0) for item in entradas_questor)
    ipi_questor = soma_unico_ipi_por_chave(entradas_questor)

    dif_total = sanitize_difference(total_sat - total_questor)
    dif_bc_icms = sanitize_difference(bc_icms_sat - bc_icms_questor)
    dif_icms = sanitize_difference(icms_sat - icms_questor)
    dif_ipi = sanitize_difference(ipi_sat - ipi_questor)

    teve_diferenca = any([dif_total != 0.0, dif_bc_icms != 0.0, dif_icms != 0.0, dif_ipi != 0.0])

    fechamento = [
        {"": "", "VLR TOTAL": "", "BC ICMS": "", "ICMS": "", "IPI": ""},
        {},
        {"A": "SAT", "B": f"{total_sat:,.2f}", "C": f"{bc_icms_sat:,.2f}", "D": f"{icms_sat:,.2f}", "E": f"{ipi_sat:,.2f}"},
        {},
        {"A": "QUESTOR", "B": f"{total_questor:,.2f}", "C": f"{bc_icms_questor:,.2f}", "D": f"{icms_questor:,.2f}", "E": f"{ipi_questor:,.2f}"},
        {},
        {"A": "DIFERENÇA", "B": f"{dif_total:,.2f}", "C": f"{dif_bc_icms:,.2f}", "D": f"{dif_icms:,.2f}", "E": f"{dif_ipi:,.2f}"}
    ]

    return fechamento, teve_diferenca


def calcular_fechamento_saidas(saidas_sat, saidas_questor):
    def soma_unico_ipi_por_chave(entradas):
        vistos = set()
        total_ipi = 0.0
        for entrada in entradas:
            chave = entrada.get('CHAVE')
            ipi = entrada.get('VLR IPI', 0.0)
            if chave not in vistos:
                total_ipi += ipi
                vistos.add(chave)
        return total_ipi

    total_sat = sum(item.get('VLR NF', 0.0) for item in saidas_sat)
    bc_icms_sat = sum(item.get('VLR BC ICMS', 0.0) for item in saidas_sat)
    icms_sat = sum(item.get('VLR ICMS', 0.0) for item in saidas_sat)
    ipi_sat = soma_unico_ipi_por_chave(saidas_sat)

    total_questor = sum(item.get('VLR NF', 0.0) for item in saidas_questor)
    bc_icms_questor = sum(item.get('VLR BC ICMS', 0.0) for item in saidas_questor)
    icms_questor = sum(item.get('VLR ICMS', 0.0) for item in saidas_questor)
    ipi_questor = soma_unico_ipi_por_chave(saidas_questor)

    # print(total_sat, total_questor, bc_icms_sat, bc_icms_questor, icms_sat, icms_questor, ipi_sat, ipi_questor)
    # print(type(total_sat), type(total_questor), type(bc_icms_sat), type(bc_icms_questor), type(icms_sat), type(icms_questor), type(ipi_sat), type(ipi_questor))

    dif_total = sanitize_difference(total_sat - total_questor)
    dif_bc_icms = sanitize_difference(bc_icms_sat - bc_icms_questor)
    dif_icms = sanitize_difference(icms_sat - icms_questor)
    dif_ipi = sanitize_difference(ipi_sat - ipi_questor)

    # print(dif_total, dif_bc_icms, dif_icms, dif_ipi)    


    teve_diferenca = any([dif_total != 0.0, dif_bc_icms != 0.0, dif_icms != 0.0, dif_ipi != 0.0])
    

    fechamento = [
        {"": "", "VLR TOTAL": "", "BC ICMS": "", "ICMS": "", "IPI": ""},
        {},
        {"A": "SAT", "B": f"{total_sat:,.2f}", "C": f"{bc_icms_sat:,.2f}", "D": f"{icms_sat:,.2f}", "E": f"{ipi_sat:,.2f}"},
        {},
        {"A": "QUESTOR", "B": f"{total_questor:,.2f}", "C": f"{bc_icms_questor:,.2f}", "D": f"{icms_questor:,.2f}", "E": f"{ipi_questor:,.2f}"},
        {},
        {"A": "DIFERENÇA", "B": f"{dif_total:,.2f}", "C": f"{dif_bc_icms:,.2f}", "D": f"{dif_icms:,.2f}", "E": f"{dif_ipi:,.2f}"}
    ]

    return fechamento, teve_diferenca

import os
import pandas as pd


import os
import pandas as pd

def read_files_xlsx(cnpj):
    folder_path = os.path.join(os.getcwd())
    all_data = {'entrada': pd.DataFrame(), 'saida': pd.DataFrame()}  # Inicializando ambos como DataFrames vazios

    if not os.path.exists(folder_path):
        print(f"A pasta {folder_path} não foi encontrada.")
        return [], []  # Retornar listas vazias se o diretório não existir

    columns_map = {
        "DataEmissao": "DATA",
        "PeriodoDeReferencia": "REF",
        "ModeloDocumento": "MOD",
        "TipoDocumento": "TIPO",
        "TipoDeOperacaoEntradaOuSaida": "E/S",
        "Situacao": "SIT",
        "ChaveAcesso": "CHAVE",
        "SerieDocumento": "SERIE",
        "NumeroDocumento": "NUM NF",
        "ValorTotalNota": "VLR NF",
        "ValorBaseCalculoICMS": "VLR BC ICMS",
        "ValorTotalICMS": "VLR ICMS",
        "ValorIPI": "VLR IPI"
    }

    ordered_columns = [
        "DATA", "REF", "MOD", "TIPO", "E/S", "SIT", "CHAVE", "SERIE",
        "FORNEC", "CFOP", "NUM NF", "VLR NF", "VLR BC ICMS",
        "VLR ICMS", "VLR IPI"
    ]

    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(folder_path, filename)
            
            try:
                df = pd.read_excel(file_path, dtype=str, engine='openpyxl')
                
                is_saida = "CnpjOuCpfDoEmitente" in df.columns and df['CnpjOuCpfDoEmitente'].eq(cnpj).all()
                is_entrada = "CnpjOuCpfDoDestinatario" in df.columns and df['CnpjOuCpfDoDestinatario'].eq(cnpj).all()

                destinatario = df.get("NomeDestinatario", '')
                emitente = df.get("NomeEmitente", '')

                available_columns = set(df.columns).intersection(columns_map.keys())
                
                if not available_columns:
                    continue

                df = df[list(available_columns)]
                reduced_map = {col: columns_map[col] for col in available_columns}
                df.rename(columns=reduced_map, inplace=True)
              
                df['CFOP'] = ''

                if is_saida:
                    df['FORNEC'] = destinatario
                elif is_entrada:
                    df['FORNEC'] = emitente

                df = df[[col for col in ordered_columns if col in df.columns]]

                # Zerar valores se a coluna "SIT" for "Cancelado"
                if "SIT" in df.columns:
                    df.loc[df["SIT"] == "Cancelado", ["VLR NF", "VLR BC ICMS", "VLR ICMS", "VLR IPI"]] = 0.0

                for col in ["VLR NF", "VLR ICMS", "VLR BC ICMS", "VLR IPI"]:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce').round(2)

                if "DATA" in df.columns:
                    df["DATA"] = pd.to_datetime(df["DATA"], errors='coerce')

                if is_saida:
                    all_data['saida'] = pd.concat([all_data['saida'], df], ignore_index=True)
                    # >>> REMOVE DUPLICATAS EM 'saida'
                    if 'CHAVE' in all_data['saida'].columns:
                        all_data['saida'].drop_duplicates(subset=['CHAVE'], inplace=True, ignore_index=True)
                elif is_entrada:
                    all_data['entrada'] = pd.concat([all_data['entrada'], df], ignore_index=True)
                    # >>> REMOVE DUPLICATAS EM 'entrada'
                    if 'CHAVE' in all_data['entrada'].columns:
                        all_data['entrada'].drop_duplicates(subset=['CHAVE'], inplace=True, ignore_index=True)

            except Exception as e:
                print(f"Erro ao processar o arquivo {filename}: {e}")

    # Movimentações:
    if not all_data['saida'].empty:
        if "E/S" in all_data['saida'].columns and "CHAVE" in all_data['saida'].columns:
            entradas_movidas = all_data['saida'][
                (all_data['saida']["E/S"] == "E") &
                all_data['saida']["CHAVE"].str.contains(cnpj) &
                (all_data['entrada'].empty or ~all_data['saida']["CHAVE"].isin(all_data['entrada']["CHAVE"]))
            ]

            # Mesmo que 'entradas' comece vazio, ele será preenchido aqui
            all_data['entrada'] = pd.concat([all_data['entrada'], entradas_movidas], ignore_index=True)
            # >>> REMOVE DUPLICATAS EM 'entrada' DEPOIS DE MOVIMENTAÇÃO
            if 'CHAVE' in all_data['entrada'].columns:
                all_data['entrada'].drop_duplicates(subset=['CHAVE'], inplace=True, ignore_index=True)

            entradas_sat = all_data['entrada'].to_dict(orient='records')

            # Remove de 'saídas' as notas que foram movidas
            all_data['saida'] = all_data['saida'][
                ~all_data['saida']["CHAVE"].isin(entradas_movidas["CHAVE"])
            ]
            # >>> REMOVE DUPLICATAS EM 'saida' APÓS FILTRO (caso reste)
            if 'CHAVE' in all_data['saida'].columns:
                all_data['saida'].drop_duplicates(subset=['CHAVE'], inplace=True, ignore_index=True)

            saidas_sat = all_data['saida'].to_dict(orient='records')
        else:
            print("Colunas necessárias não estão presentes em 'saida' para realizar o filtro.")
            entradas_sat = all_data['entrada'].to_dict(orient='records')
            saidas_sat = all_data['saida'].to_dict(orient='records')
    else:
        entradas_sat = all_data['entrada'].to_dict(orient='records')
        saidas_sat = all_data['saida'].to_dict(orient='records')

    return entradas_sat, saidas_sat

def converter_query_rows_para_dict(interesse, columns_map):
    try:
        cursor_data = interesse['rows']
        header = [column[0].upper() for column in cursor_data['cursor'].description]
        specified_header = [col for col in header if col in interesse['cols_DB']]
        converted_rows = []
        for row in cursor_data['rows']:
            row_dict = {}
            for idx, col in enumerate(header):
                if col in specified_header:
                    value = row[idx]
                    new_col = columns_map[col]

                    # Aplicar transformação de dados e mapeamento
                    if new_col == 'VLR NF':
                        value = Decimal(value) if value is not None else None
                    elif new_col == 'NUM NF':
                        value = str(value)  # Garantir que o número NF seja uma string
                    elif new_col == 'DATA':
                        value = pd.to_datetime(value).date()

                    row_dict[new_col] = value
            converted_rows.append(row_dict)
        return converted_rows
    except Exception as e:
        print(f"Erro ao processar resultados da consulta: {e}")
        return []

def executar_consulta_obter_dados(insc_federal, start_date, end_date, query_function):
    pg_conn = PostgresConnection()
    connection = pg_conn.connect()
    query_data = {}

    if connection:
        try:
            cursor = connection.cursor()
            query = query_function()  # Certifique-se de já ter ajustado os ? para %s!
            cursor.execute(query, (insc_federal, insc_federal, start_date, end_date))
            rows = cursor.fetchall()
            query_data = {'cursor': cursor, 'rows': rows}
        except Exception as e:
            print(f"Erro durante execução da consulta: {e}")
        finally:
            pg_conn.disconnect(connection)

    return query_data


def consultar_entradas_questor(insc_federal, start_date, end_date):
    print('Efetuando consulta banco de dados Questor das Entradas')
    return executar_consulta_obter_dados(insc_federal, start_date, end_date, get_lctofisent_query)

def fazer_consulta_sql_saidas(insc_federal, start_date, end_date):
    print('Efetuando consulta banco de dados Questor das Saídas')
    return executar_consulta_obter_dados(insc_federal, start_date, end_date, get_lctofissai_query)

if __name__ == '__main__':
    # Chamada do fluxo completo
    entradas_sat, saidas_sat, entradas_questor, saidas_questor, only_in_sat = read_files_and_query('06.197.478/0001-30', '2024-12-01', '2024-12-31')

    # Processar os dados, como salvar no Excel ou analisar como necessário.