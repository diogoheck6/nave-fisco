from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from time import sleep
import os

folder_path = os.path.join(os.getcwd())

class ServiceConsultCTE:

    def __init__(self, link, driver, wait):
        self.link = link
        self.driver = driver
        self.wait = wait
        self.testado = False

    def acessar_link(self):
        self.driver.get(self.link)

    def limpar_inputs(self):

        self.limpar_emitente()
        sleep(1)
        self.limpar_destinatario()
        sleep(1)
        

    def selecionar_nfce(self):

        campo_input = self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//div[@id="s2id_Body_Main_Main_sepConsultaNfpe_selTipoDocumento"]/a')))
    
        campo_input.click()
    
        sleep(1)

        select_result = self.wait.until(EC.presence_of_all_elements_located(
            (By.XPATH, '//div[@class="select2-result-label"]')))
    

        select_result[1].click()
    
    def selecionar_nfe(self):

        campo_input = self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//div[@id="s2id_Body_Main_Main_sepConsultaNfpe_selTipoDocumento"]/a')))
    
        campo_input.click()
    
        sleep(1)

        select_result = self.wait.until(EC.presence_of_all_elements_located(
            (By.XPATH, '//div[@class="select2-result-label"]')))
    

        select_result[0].click()


    def selecionar_emitente(self, cnpj):
        self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//input[@id="Body_Main_Main_sepConsultaNfpe_ctl10_idnEmitente_MaskedField"]'))).send_keys(cnpj)
        

    def limpar_emitente(self):
        self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//input[@id="Body_Main_Main_sepConsultaNfpe_ctl10_idnEmitente_MaskedField"]'))).clear()
        
        
    def selecionar_destinatario(self, cnpj):
        self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//input[@id="Body_Main_Main_sepConsultaNfpe_ctl11_idnDestinatario_MaskedField"]'))).send_keys(cnpj)
        
    
    def limpar_destinatario(self):
        self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//input[@id="Body_Main_Main_sepConsultaNfpe_ctl11_idnDestinatario_MaskedField"]'))).clear()
        

    def inserir_periodo(self, data_inicial):
        self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//input[@placeholder="mm/aaaa"]'))).send_keys(data_inicial)
    
    def limpar_data_emissao_inicial(self):
        self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//input[@id="Body_Main_Main_sepConsultaNfpe_datDataInicial"]'))).clear()
        
    def inserir_data_emissao_final(self, data_final):
        self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//input[@id="Body_Main_Main_sepConsultaNfpe_datDataFinal"]'))).send_keys(data_final)
    
    def limpar_data_emissao_final(self):
        self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//input[@id="Body_Main_Main_sepConsultaNfpe_datDataFinal"]'))).clear()
        

    def verificar_se_esta_disponivel_input_captha(self):
        try:
            sleep(1)
            self.wait.until(EC.visibility_of_element_located(
                                        (By.XPATH, '//input[@placeholder="Digite o texto"]')))     
            print('Captcha precisa ser resolvido')
            return True
        except:
            print('Captcha está resolvido, pode serguir normalmente!')
            
            return False        

    def definir_tipo_pesquisa(self):

        campo_input = self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//div[@id="s2id_Body_Main_Main_sepBusca_selTipoPesquisa"]/a')))
    
        sleep(1)
        
        campo_input.click()

        sleep(1)

        select_result = self.wait.until(EC.presence_of_all_elements_located(
                (By.XPATH, '//div[@class="select2-result-label"]')))
                
        sleep(1)

        select_result[1].click()       
        
                    
        
    def clicar_botao_consultar(self):
                    
        
        self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//a[@id="Body_Main_Main_sepBusca_btnPesquisar"]'))).click()
    
    
    def clicar_botao_buscar(self):
                    
        
        self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//a[@id="Body_Main_Main_grpCTe_btn0"]'))).click()
    
    def baixar_excel(self):
                    
        
        botao_baixar = self.wait.until(EC.presence_of_all_elements_located(
            (By.XPATH, '//ul[@aria-labelledby="Body_Main_Main_grpCTe_btn0"]//li')))
        
        botao_baixar[0].click()

    
    def realizar_operacao(self, max_retries=1):
    
        try:
            # Clicar no botão "Consultar"
            self.clicar_botao_consultar()
            
            # Inicia a contagem de tentativas
            retries = 0
            sucesso_na_operacao = False

            while retries < max_retries and not sucesso_na_operacao:
                # Espera curta após clicar no botão de consultar
                sleep(1)
                
                try:
                    # Tentar fazer as ações de buscar e baixar
                    print(f"Tentativa {retries + 1}/{max_retries}: tentando buscar e baixar Excel.")
                    
                    self.clicar_botao_buscar()

                    # Espera após a ação de buscar
                    sleep(1)

                    self.baixar_excel()

                    print("Download do Excel realizado com sucesso.")
                    sucesso_na_operacao = True  # Marca como sucesso após operações bem-sucedidas

                except Exception as e:
                    # print(f"Erro ao tentar buscar e baixar: {e}")
                    retries += 1

                    if retries < max_retries:
                        print(f"Tentativa {retries + 1}/{max_retries} em transcurso após falha. Tentando novamente...")
                        sleep(2)  # Pausa adicional antes de tentar novamente
                
            if not sucesso_na_operacao:
                print("Sem movimento de Ctes.")

        except Exception as e:
            print(f"Erro na operação geral: {e}")

 
        
        

    def pesquisar_cnpj(self, emitente_id):

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

    

    def consultar_ctes(self, cnpj, mmaaaa):
        
        self.acessar_link()

        sleep(1)

        self.inserir_periodo(mmaaaa)

        sleep(1)

        self.pesquisar_cnpj(cnpj)

        sleep(1)

        self.realizar_operacao()

        sleep(1)

        self.definir_tipo_pesquisa()

        sleep(1)

        self.realizar_operacao()
        
        sleep(1)
    
    



# //div[@class="sat-vs"]
    





