from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from models.db_sqlite_v2 import NumerosSolicitacaoCTE
import os


class MenuSolicitacaoCTE:

    def __init__(self, link, driver, wait, db_path=None):
        if not db_path:
            db_path = os.path.join(os.getcwd(), 'NumerosSolicitacaoCTE.db')
        self.link = link
        self.driver = driver
        self.wait = wait
        self.db = NumerosSolicitacaoCTE(db_path)

    def acessar_link(self):

        self.driver.get(self.link)

    def definir_tipo_pesquisa(self, tipo):

        campo_input = self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//div[@id="s2id_Body_Main_Main_sepBusca_selTipoPesquisa"]/a')))
    
        sleep(1)
        
        campo_input.click()

        sleep(1)

        select_result = self.wait.until(EC.presence_of_all_elements_located(
                (By.XPATH, '//div[@class="select2-result-label"]')))
                
        sleep(1)

        select_result[tipo].click()

    def definir_situacao(self):

        campo_input_sit = self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//div[@id="s2id_Body_Main_Main_sepBusca_selSituacao"]/a')))
    
        sleep(1)

        campo_input_sit.click()

        sleep(1)

        select_situacao = self.wait.until(EC.presence_of_all_elements_located(
                (By.XPATH, '//div[@class="select2-result-label"]')))
        
        sleep(1)
        
        select_situacao[0].click()

    
    def definir_periodo_inicial(self, mes_ano):
        self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//input[@id="Body_Main_Main_sepBusca_mopPeriodoInicio"]'))).send_keys(mes_ano)

    def definir_periodo_final(self, mes_ano):
        self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//input[@id="Body_Main_Main_sepBusca_mopPeriodoFim"]'))).send_keys(mes_ano)
        
    def definir_emitente_tomador(self, emitente_id):

        campo_input_cnpj = self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//div[@id="s2id_Body_Main_Main_sepBusca_ctbContribuinte_hid_single_ctbContribuinte_value"]/a')))
    
        sleep(1)

        campo_input_cnpj.click()

        sleep(1)

        campo_search_cnpj = self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//div[@id="select2-drop"]/div/input')))
    
        sleep(1)

        campo_search_cnpj.send_keys(emitente_id) 

        sleep(1)

        while(True):

            try:

                self.wait.until(EC.presence_of_element_located((By.XPATH, '//div[@id="select2-drop"]//li[contains(text(), "Nenhum resultado encontrado")]')))
                print('estou saindo, pois esse cliente nao existe...')
                sleep(2)
                break
  
            except:

                cliente_encontrado = self.wait.until(EC.presence_of_element_located((By.XPATH, f'//div[@class="select2-result-label clearfix" and contains(., "{emitente_id}")]')))
                # print('encontrei o cliente, oito e sete galera...')
                cliente_encontrado.click()
                sleep(2)
                break

    def _salvar_mensagem(self, cnpj, competencia, razao_social, operacao, mensagem_id=None):
        # Certifique-se de criar/adicionar cliente
        self.db.adicionar_cliente(cnpj, razao_social)

        # Se `mensagem_id` estiver presente ou gerado, associe-o
        if mensagem_id:
            if not self.db.id_existe(mensagem_id, operacao):
                self.db.adicionar_competencia(cnpj, competencia, operacao, mensagem_id)  # Escolha um nome de operação adequado
            else:
                print(f'ID {mensagem_id} já existe no banco de dados.')
        else:
            # Lidando com casos sem `mensagem_id`, permite salvar sem ele, ou log
            print(f'Competência {competencia} para CNPJ {cnpj} adicionada sem mensagem ID especificamente.')
            self.db.adicionar_competencia(cnpj, competencia, None, operacao="solicitacao_sem_id")
    
    def solicitar_relatorio(self, cnpj, competencia, razao_social, operacao, max_retries=3):

        try:
            # Clica no botão para solicitar o relatório
            self.wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//a[@id="Body_Main_Main_sepBusca_btnSolicitarRelatorio"]'))).click()

            # Tentativas para buscar a mensagem de sucesso ou erro
            retries = 0
            mensagem_id = None  # Inicializa `mensagem_id`
            while retries < max_retries:
                try:
                    # Tenta encontrar a mensagem de sucesso
                    msg_sucesso = self.wait.until(EC.presence_of_element_located(
                        (By.XPATH, '//div[@class="sat-vs"]//ul[@class="sat-vs-success"]//li[2]')))
                    
                    # Verifica se o texto da mensagem não está vazio antes de tentar processá-lo
                    if msg_sucesso.text.strip():  # Verifica se a mensagem não está vazia
                        mensagem_id = msg_sucesso.text.split(":")[1].strip().replace('.', '')
                        print(mensagem_id)
                        self._salvar_mensagem(cnpj, competencia, razao_social, operacao, mensagem_id)
                        break  # Sai do loop após encontrar a mensagem de sucesso
                    else:
                        print("Mensagem de sucesso está vazia, tentando novamente...")
                        retries += 1
                        sleep(2)  # Pausa para dar tempo para a página atualizar

                except TimeoutException:
                    retries += 1
                    if retries >= max_retries:
                        print("Falha ao encontrar a mensagem de sucesso após várias tentativas.")
                        break  # Caso atinja o número máximo de tentativas
                    print(f"Tentando novamente... Tentativa {retries}/{max_retries}")
                    sleep(2)  # Pausa para dar tempo para a página atualizar

                try:
                    # Se não encontrou sucesso, tenta buscar a mensagem de erro
                    msg_erro = self.wait.until(EC.presence_of_element_located(
                        (By.XPATH, '//div[@class="sat-vs"]//ul[@class="sat-vs-error"]//li[2]')))
                    print(f"Erro encontrado: {msg_erro.text}")
                    break  # Sai do loop após encontrar a mensagem de erro
                except TimeoutException:
                    retries += 1
                    if retries >= max_retries:
                        print("Falha ao encontrar a mensagem de erro após várias tentativas.")
                        break  # Caso atinja o número máximo de tentativas
                    print(f"Tentando novamente... Tentativa {retries}/{max_retries}")
                    sleep(2)  # Pausa para dar tempo para a página atualizar

        except Exception as e:
            print(f"Erro ao buscar relatório: {e}")

            
    