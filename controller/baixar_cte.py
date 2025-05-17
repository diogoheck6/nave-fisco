# SERVIÇO DE CONSULTA DE CTE ATRAVÉS DOS NÚMEROS DE REQUISIÇÃO JÁ SOLICITADOS NO SERVIÇO DE REQUISIÇÃO DE RELATÓRIOS DE CTE
from browser.connection import start_browser
from login.login import login
from time import sleep
from models.Service_Consult_CTE import ServiceConsultCTE

CPF = ''
SENHA = ''


def baixar_relat_ctes(driver, wait, cnpj, razao_social, mes_ano):
    

    servico_consulta_cte = ServiceConsultCTE(
        link = 'https://sat.sef.sc.gov.br/tax.NET/tax.Net.ReportGenerator/ConsultaSolicitacaoPrivado.aspx',
        driver=driver,
        wait=wait
    )

    servico_consulta_cte.acessar_link()

    sleep(1)

    servico_consulta_cte.baixar_relatorios(cnpj, razao_social, mes_ano)

    sleep(1)

 
