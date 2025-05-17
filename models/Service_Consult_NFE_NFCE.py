import tkinter as tk
import requests
from anticaptchaofficial.imagecaptcha import imagecaptcha
from selenium.webdriver.support import expected_conditions as EC
from tkinter import messagebox
from selenium.webdriver.common.by import By
from time import sleep

import os

folder_path = os.path.join(os.getcwd())

class ServiceConsultNFENFCE:

    def __init__(self, link, driver, wait):
        self.link = link
        self.driver = driver
        self.wait = wait
        self.captcha_resolvido = False  

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
        

    def inserir_data_emissao_inicial(self, data_inicial):
        self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//input[@id="Body_Main_Main_sepConsultaNfpe_datDataInicial"]'))).send_keys(data_inicial)
    
    def limpar_data_emissao_inicial(self):
        self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//input[@id="Body_Main_Main_sepConsultaNfpe_datDataInicial"]'))).clear()
        
    def inserir_data_emissao_final(self, data_final):
        self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//input[@id="Body_Main_Main_sepConsultaNfpe_datDataFinal"]'))).send_keys(data_final)
    
    def limpar_data_emissao_final(self):
        self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//input[@id="Body_Main_Main_sepConsultaNfpe_datDataFinal"]'))).clear()
        


    def solve_captcha(self, driver, wait):

        while True:
            # Baixar a imagem do captcha
            src_img = wait.until(EC.presence_of_element_located(
                    (By.XPATH, '//div[@class="sat-captcha-image-container"]/img'))).get_attribute('src')
            print("URL da imagem do captcha:", src_img)

            response = requests.get(src_img)
            with open('captcha.jpeg', 'wb') as f:
                f.write(response.content)

            # Resolver o captcha
            solver = imagecaptcha()
            solver.set_verbose(1)
            solver.set_key("")  # Substitua pela sua chave da API AntiCaptcha
            

            captcha_text = solver.solve_and_return_solution('captcha.jpeg')

            if captcha_text != 0:
                print("Texto do captcha:", captcha_text)
                return captcha_text
            else:
                print("Erro ao resolver o captcha:", solver.error_code)
                
                if solver.error_code == "ERROR_NO_SLOT_AVAILABLE":
                    # Aguarde um curto período antes de tentar novamente
                    sleep(10)
                else:
                    # Para outros erros, tente novamente imediatamente
                    continue

    def inserir_captcha(self, captcha):
        captcha_input = self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//input[@placeholder="Digite o texto"]')))  # Ajuste o XPath conforme necessário
        captcha_input.send_keys(captcha)


    def verificar_se_esta_disponivel_input_captha(self):
        try:
            sleep(1)
            self.wait.until(EC.visibility_of_element_located(
                (By.XPATH, '//input[@placeholder="Digite o texto"]')))
            print('Captcha precisa ser resolvido')
            return True
        except Exception:
            print('Captcha não está visível, pode seguir normalmente!')
            return False

    def clicar_botao_exportar(self, max_retries=3):
        # Caso o CAPTCHA ja tenha sido resolvido, simplesmente clique no botão exportar
        if self.captcha_resolvido:
            print("Captcha já resolvido. Clicando no botão exportar sem verificar CAPTCHA.")
            self._executar_exportacao()
            return

        tentativas_captcha = 0

        while tentativas_captcha < max_retries:
            try:
                self._executar_exportacao()

                # Após o clique, verificar se o CAPTCHA precisa ser resolvido
                if self.verificar_se_esta_disponivel_input_captha():
                    print('Resolvendo captcha')
                    captcha = self.solve_captcha(self.driver, self.wait)

                    print('Inserindo Captcha')
                    self.inserir_captcha(captcha)
                    sleep(1)

                else:
                    print('Captcha resolvido corretamente. Prosseguindo com a exportação...')
                    self.captcha_resolvido = True  # Marca como resolvido
                    break  # Sai do loop, encaminhando a resolução e exportação

            except Exception as e:
                print(f"Um erro ocorreu ao tentar exportar: {e}")

            tentativas_captcha += 1
        
        if not self.captcha_resolvido:
            print("Máximo de tentativas para resolver o captcha alcançado.")
            self.encerrar_aplicacao()

    def _executar_exportacao(self):
        """Método auxiliar para encapsular o clique no botão de exportação."""
        print('Clicando no botão exportar')
        self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//a[@id="Body_Main_Main_sepConsultaNfpe_btnExportar"]'))).click()
        sleep(2)

    def encerrar_aplicacao(self):
        root = tk.Tk()
        root.withdraw()  # Oculta a janela principal do tkinter
        print("Encerrando a aplicação devido ao excesso de tentativas no CAPTCHA...")
        messagebox.showwarning("Operação Cancelada", "Não foi possível resolver o CAPTCHA após várias tentativas. O aplicativo será encerrado.")
        root.quit()  # Fecha o loop principal e encerra o Tkinter

    # Métodos solve_captcha e inserir_captcha seguem aqui...
