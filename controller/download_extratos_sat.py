from browser.connection import start_browser
from login.login import login
from time import sleep
from models.Service_Consult_NFE_NFCE import ServiceConsultNFENFCE
from controller.baixar_nfe import baixar_nfes
from controller.baixar_nfce import baixar_nfce
import os


PastaDownloads = os.path.join(os.getcwd())



def fazer_download(CPF, SENHA, cnpj, data_inicial, data_final):

    try:
        
        driver, wait = start_browser()

        
        login(driver, wait, CPF, SENHA)

        sleep(1)
        
        servico_consulta_nfe_nfce = ServiceConsultNFENFCE(
            link='https://sat.sef.sc.gov.br/tax.NET/Sat.NFe.Web/Consultas/ConsultaOnlineCC.aspx',
            driver=driver,
            wait=wait
        )

        servico_consulta_nfe_nfce.acessar_link()
        
        sleep(1)

        servico_consulta_nfe_nfce.limpar_inputs()

        sleep(1)

        
        baixar_nfes(servico_consulta_nfe_nfce, cnpj, data_inicial, data_final, driver, wait)

        sleep(1)

        servico_consulta_nfe_nfce.limpar_inputs()

        sleep(1)
        
        baixar_nfce(servico_consulta_nfe_nfce, cnpj, data_inicial, data_final, driver, wait)

        sleep(5)

        print('Extratods baixados com Sucesso!!!')


    finally:
        
        # driver.quit()
        pass
        




if __name__ == '__main__':
    
    fazer_download()

