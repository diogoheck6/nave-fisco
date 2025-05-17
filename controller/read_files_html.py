import os
import zipfile
import pandas as pd
from bs4 import BeautifulSoup


class ServiceLerHtmlCTE:

    def __init__(self):
        self.folder_path = os.getcwd()

    def _unzip_files(self, zip_path):
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(self.folder_path)

    def _process_html(self, file_path):
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
            html = file.read()

        soup = BeautifulSoup(html, 'html.parser')
        tables = soup.find_all('table')

        all_data = []
        if tables:
            headers = [th.text.strip() for th in tables[0].find_all('tr')[0].find_all('td')]
            for row in tables[0].find_all('tr')[1:]:
                values = [td.text.strip() for td in row.find_all('td')]
                if len(values) == len(headers):
                    row_dict = dict(zip(headers, values))
                    all_data.append(row_dict)
        return all_data

    def ler_arquivos_html(self, cnpj_cliente):

        # Mapeamento correto das colunas de HTML para o formato padronizado
        column_map_html = {
            "DATA_EMISSO": "DATA",
            "REFERNCIA": "REF",
            # MOD falta e será adicionada depois como padrão 57
            "TIPO_CTE": "TIPO",
            # E/S falta e será definida como vazio inicialmente
            "SITUACAO": "SIT",
            "CHAVE_DE_ACESSO": "CHAVE",
            "SRIE": "SERIE",
            "NOME_EMITENTE": "FORNEC",
            "CFOP": "CFOP",
            "NMERO_CTE": "NUM NF",
            "VALOR_TOTAL_PREST": "VLR NF",
            "VALOR_BC_ICMS": "VLR BC ICMS",
            "VALOR_ICMS": "VLR ICMS",
            # VLR IPI
        }

        ordered_columns = [
            "DATA", "REF", "MOD", "TIPO", "E/S", "SIT", "CHAVE", "SERIE", 
               "FORNEC", "CFOP" ,"NUM NF", "VLR NF", "VLR BC ICMS", "VLR ICMS", "VLR IPI"
        ]

        entradas_sat, saidas_sat = [], []

        for filename in os.listdir(self.folder_path):
            
            if filename.endswith('.zip'):
                zip_file_path = os.path.join(self.folder_path, filename)
                try:
                    self._unzip_files(zip_file_path)
                    

                    for extracted_filename in os.listdir(self.folder_path):
                        if extracted_filename.endswith('.xls'):
                            xls_path = os.path.join(self.folder_path, extracted_filename)
                            html_filename = extracted_filename.replace('.xls', '.html')
                            html_path = os.path.join(self.folder_path, html_filename)
                            os.rename(xls_path, html_path)

                            html_data = self._process_html(html_path)
                            
                            # print("HTML Data Processado:", html_data)
                            
                            if html_data:
                                df = pd.DataFrame(html_data)

                                available_columns = set(df.columns).intersection(column_map_html.keys())
                                df = df[list(available_columns)]
                                df.rename(columns={col: column_map_html[col] for col in available_columns}, inplace=True)

                                

                                df['MOD'] = 57  # Adiciona a coluna MOD com valor fixo de 57
                                df['E/S'] = ''  # Adiciona a coluna E/S vazia
                                df['VLR IPI'] = 0

                                

                                df = df[[col for col in ordered_columns if col in df.columns]]

                                if "SIT" in df.columns:
                                    df.loc[df["SIT"] == "CANCELADO", ["VLR NF", "VLR BC ICMS", "VLR ICMS", "VLR IPI"]] = 0.0


                                for col in ["VLR NF", "VLR ICMS", "VLR BC ICMS"]:
                                    if col in df.columns:
                                        df[col] = pd.to_numeric(df[col].str.replace(',', '.'), errors='coerce').fillna(0).round(2)

                                if "DATA" in df.columns:
                                    # df["DATA"] = pd.to_datetime(df["DATA"], errors='coerce')
                                    df["DATA"] = pd.to_datetime(df["DATA"], errors='coerce').dt.strftime('%d/%m/%Y')

                                # Removendo pontos das chaves
                                if "CHAVE" in df.columns:
                                    df["CHAVE"] = df["CHAVE"].astype(str).apply(lambda x: x.replace('.', '') if pd.notnull(x) else x)

                                

                                # Determinando se é entrada ou saída
                                for _, row in df.iterrows():
                                
                                    if cnpj_cliente in row["CHAVE"]:
                                        saidas_sat.append(row)
                                    else:    
                                        entradas_sat.append(row)
                                

                except zipfile.BadZipFile:
                    print(f'Erro: {zip_file_path} não é um arquivo ZIP válido.')
                except Exception as e:
                    print(f'Erro ao processar o arquivo ZIP {zip_file_path}: {e}')

        entradas_sat_df = pd.DataFrame(entradas_sat, columns=ordered_columns)
        saidas_sat_df = pd.DataFrame(saidas_sat, columns=ordered_columns)

        # Para verificação, exporte o arquivo gerado
        entradas_sat_df.to_excel('entradas_sat_cte.xlsx', index=False)
        saidas_sat_df.to_excel('saidas_sat_cte.xlsx', index=False)
        print("Dados de entrada e saída exportados para verificação.")

        return entradas_sat_df.to_dict(orient='records'), saidas_sat_df.to_dict(orient='records')