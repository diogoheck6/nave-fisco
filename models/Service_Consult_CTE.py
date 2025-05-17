from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from time import sleep
import os
import zipfile
from bs4 import BeautifulSoup
from models.db_sqlite_v2 import NumerosSolicitacaoCTE

class ServiceConsultCTE:
    
    def __init__(self, link, driver, wait,db_path=None):
        self.link = link
        self.driver = driver
        self.wait = wait
        if not db_path:
            db_path = os.path.join(os.getcwd(), 'NumerosSolicitacaoCTE.db')
        self.db = NumerosSolicitacaoCTE(db_path)
        self.folder_path = os.getcwd()

    def acessar_link(self):

        self.driver.get(self.link)

    def _unzip_files(self, zip_path):
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            # Extrair todos os arquivos diretamente para o diretório atual
            zip_ref.extractall(self.folder_path)

    def _process_html(self, file_path):

        with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
            html = file.read()

        soup = BeautifulSoup(html, 'html.parser')
        tables = soup.find_all('table')

        all_data = []
        if tables:
            # Extraindo as colunas (primeira linha <TR>)
            headers = [th.text.strip() for th in tables[0].find_all('tr')[0].find_all('td')]

            # Iterando sobre todas as linhas de dados, exceto a primeira
            for row in tables[0].find_all('tr')[1:]:
                values = [td.text.strip() for td in row.find_all('td')]

                # Criando um dicionário associando as colunas aos valores da linha
                if len(values) == len(headers):
                    row_dict = dict(zip(headers, values))
                    all_data.append(row_dict)

        return all_data



    def baixar_relatorios(self, cnpj, razao_social, mes_ano):
        # Recupera todos os IDs da competência específica para o CNPJ fornecido
        ids_cte = self.db.recuperar_ids(cnpj, mes_ano)

        if not ids_cte:
            print(f'Nenhum ID encontrado para -> {razao_social}:{cnpj} na competência {mes_ano}.')
            return

        for mensagem in ids_cte:
            # Com a lista de IDs para essa competência, continue com o workflow
            campo_input = self.wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//input[@id="ctl00_ctl00_Main_Main_txtNumeroSolicitacao"]')))
            campo_input.clear()
            campo_input.send_keys(mensagem)

            sleep(2)

            # Clica no botão de pesquisar
            self.wait.until(EC.presence_of_element_located(
                (By.XPATH, '//input[@type="submit"]'))).click()
            
            sleep(2)

            # Verifica o status após o clique
            try:
                status_element = self.wait.until(EC.presence_of_element_located(
                    (By.XPATH, '//tr[@class="dados"]/td[4]')))
                
                status = status_element.text.strip()

                if status != 'Processado com Sucesso':
                    print(f"Mensagem '{mensagem}' não processada com sucesso. Status: {status}")
                    continue

                print(f"Mensagem '{mensagem}' processada com sucesso. Tentando clicar no link...")

            except Exception as e:
                print(f"Erro ao verificar o status para a mensagem: {mensagem}. Erro: {e}")
                continue

            tentativas = 3
            for tentativa in range(tentativas):
                try:
                    link = self.wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '//tr[@class="dados"]/td/a')))
                    link.click()
                    print(f"Link clicado para a mensagem: {mensagem}")
                    break
                except Exception as e:
                    print(f"Erro ao tentar clicar no link para a mensagem: {mensagem}. Tentativa {tentativa + 1}/{tentativas}. Erro: {e}")
                    sleep(1)
            else:
                print(f"Falha ao tentar clicar no link após {tentativas} tentativas.")





    def ler_arquivos_html(self):

        all_results = {}

        # Procurando e descompactando arquivos .zip na pasta
        for filename in os.listdir(self.folder_path):
            if filename.endswith('.zip'):
                zip_file_path = os.path.join(self.folder_path, filename)

                try:
                    # Descompactando o arquivo zip diretamente na pasta atual
                    self._unzip_files(zip_file_path)

                    # Agora, vamos procurar pelos arquivos .xls na pasta extraída
                    for extracted_filename in os.listdir(self.folder_path):
                        if extracted_filename.endswith('.xls'):
                            xls_path = os.path.join(self.folder_path, extracted_filename)

                            # Renomeando a extensão de .xls para .html
                            html_filename = extracted_filename.replace('.xls', '.html')
                            html_path = os.path.join(self.folder_path, html_filename)
                            os.rename(xls_path, html_path)

                            # Processando o arquivo HTML gerado
                            html_data = self._process_html(html_path)

                            # Armazenando o resultado
                            if html_data:  # Só armazene se a conversão e leitura foram bem-sucedidas
                                all_results[html_filename] = html_data
                            else:
                                print(f"Sem dados para o arquivo {html_filename}.")
                except zipfile.BadZipFile:
                    print(f"Erro: {zip_file_path} não é um arquivo ZIP válido.")
                except Exception as e:
                    print(f"Erro ao processar o arquivo ZIP {zip_file_path}: {e}")

        # Exibindo os resultados de todos os arquivos HTML processados
        for html_filename, data in all_results.items():
            print(f"Dados extraídos de {html_filename}:")
            for row in data:

                chave_de_acesso = row['CHAVE_DE_ACESSO'].replace('.', '')
        
                # Criando um novo dicionário onde a CHAVE_DE_ACESSO é a chave
                # O valor será o dicionário com os dados dessa row
                result_dict = {chave_de_acesso: row}
                
                # Exibindo os resultados
                print(result_dict)





    

    