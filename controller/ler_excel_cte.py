import os
import pandas as pd

class ServiceLerExcelCTE:
    
    def __init__(self):
        self.folder_path = os.getcwd()

    def ler_arquivos_excel(self, cnpj_cliente):
        column_map_excel = {
            "Data emissão": "DATA",
            "Referência": "REF",
            # MOD will be added with default 57
            "Tipo CTe": "TIPO",
            # E/S will be initially empty
            "Situação": "SIT",
            "Chave acesso": "CHAVE",
            "Série": "SERIE",
            "Nome emitente": "FORNEC",
            "CFOP": "CFOP",
            "Número CTe": "NUM NF",
            "Valor total prestação": "VLR NF",
            "Valor BC ICMS": "VLR BC ICMS",
            "Valor ICMS": "VLR ICMS",
        }

        ordered_columns = [
            "DATA", "REF", "MOD", "TIPO", "E/S", "SIT", "CHAVE", "SERIE", 
               "FORNEC", "CFOP" ,"NUM NF", "VLR NF", "VLR BC ICMS", "VLR ICMS", "VLR IPI"
        ]

        entradas_sat, saidas_sat = [], []

        for filename in os.listdir(self.folder_path):
            if filename.endswith('.xlsx') and 'CTe' in filename:
                file_path = os.path.join(self.folder_path, filename)

                try:
                    # Read Excel starting from row 3 and assume header is the 3rd row
                    df = pd.read_excel(file_path, header=2)

                    available_columns = set(df.columns).intersection(column_map_excel.keys())
                    df = df[list(available_columns)]
                    df.rename(columns={col: column_map_excel[col] for col in available_columns}, inplace=True)

                    df['MOD'] = 57  # Add fixed value for MOD
                    df['E/S'] = ''  # Add empty column for E/S indicator
                    df['VLR IPI'] = 0  # Assuming it's missing, adds with default 0

                    df = df[[col for col in ordered_columns if col in df.columns]]

                    if "SIT" in df.columns:
                        # Set values to 0 if the situation is "CANCELADO"
                        df.loc[df["SIT"] == "CANCELADO", ["VLR NF", "VLR BC ICMS", "VLR ICMS", "VLR IPI"]] = 0.0 

                    for col in ["VLR NF", "VLR ICMS", "VLR BC ICMS"]:
                        if col in df.columns:
                            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0).round(2)

                    if "DATA" in df.columns:
                        df["DATA"] = pd.to_datetime(df["DATA"], format='%d/%m/%Y %H:%M:%S', errors='coerce').dt.strftime('%d/%m/%Y')

                    if "CHAVE" in df.columns:
                        df["CHAVE"] = df["CHAVE"].astype(str).apply(lambda x: x.replace('.', '') if pd.notnull(x) else x)

                    # Determine if it's input or output based on CNPJ presence in the key
                    for _, row in df.iterrows():
                        if cnpj_cliente in row["CHAVE"]:
                            saidas_sat.append(row)
                        else:    
                            entradas_sat.append(row)

                except Exception as e:
                    print(f"Erro ao processar o arquivo {file_path}: {e}")

        entradas_sat_df = pd.DataFrame(entradas_sat, columns=ordered_columns)
        saidas_sat_df = pd.DataFrame(saidas_sat, columns=ordered_columns)

        # Export the generated file for verification
        entradas_sat_df.to_excel('entradas_sat_cte.xlsx', index=False)
        saidas_sat_df.to_excel('saidas_sat_cte.xlsx', index=False)
        print("Dados de entrada e saída exportados para verificação.")

        return entradas_sat_df.to_dict(orient='records'), saidas_sat_df.to_dict(orient='records')
