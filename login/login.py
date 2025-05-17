# login.py
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from time import sleep


def login(driver, wait, username, password):
    # Acessar a página inicial do site
    driver.get('https://sat.sef.sc.gov.br')

    # Digitar usuário
    wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//input[@id="Body_pnlMain_tbxUsername"]'))).send_keys(username)

    sleep(1)

    # Digitar senha
    wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//input[@id="Body_pnlMain_tbxUserPassword"]'))).send_keys(password)
    
    sleep(1)

    # Clicar no botão Entrar
    wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//a[@id="Body_pnlMain_btnLogin"]'))).click()
    
    sleep(1)

    print("Login realizado com sucesso!")
