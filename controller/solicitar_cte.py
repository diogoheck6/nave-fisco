# App para efetuar o download dos extratos dos CTEs do site do Sefaz - SC
from browser.connection import start_browser
from login.login import login
from time import sleep
from models.solicitacao_ctes import MenuSolicitacaoCTE

CPF = ''
SENHA = ''
EMITENTE = 0
TOMADOR = 1



def solicitar_ctes(driver, wait, cnpj, razao_social, mes_ano):

    try:



        sleep(1)

        requisitar_cte = MenuSolicitacaoCTE(
            link='https://sat.sef.sc.gov.br/tax.NET/Sat.dfe.cte.web/Relatorios/RelatorioCTe.aspx',
            driver=driver,
            wait = wait
        )



        # ENTRAR NO MENU DE SOLICITAÇÃO DE RELATÓRIOS DO CT-e
        print('Acessando site consultar CTE')
        requisitar_cte.acessar_link()
        sleep(1)

        print('Selecionar pesquisa para emitente')
        # DEFINIR O TIPO DE PESQUISA (CONSULTA POR EMITENTE=0 OU TOMADOR=1)
        requisitar_cte.definir_tipo_pesquisa(tipo=EMITENTE)

        sleep(1)

        print('Selecionar situação')
        # DEFINIR SITUAÇÃO (SOMENTE AUTORIZADOS, SOMENTE CANCELADOS, SOMENTE DENEGADOS)
        requisitar_cte.definir_situacao()

        sleep(1)

        print('inserir período inicial')
        # DEFINIR O PERÍODO INICIAL
        requisitar_cte.definir_periodo_inicial(mes_ano=mes_ano)

        sleep(1)

        print('inserir período final')
        # DEFINIR O PERÍODO FINAL
        requisitar_cte.definir_periodo_final(mes_ano=mes_ano)

        sleep(1)

        print('inserir CNPJ')
        # BUSCAR O EMITENTE/TOMADOR
        requisitar_cte.definir_emitente_tomador(emitente_id=cnpj)

        sleep(1)

        print('Solicitar relatório')
        requisitar_cte.solicitar_relatorio(cnpj=cnpj, competencia=mes_ano, razao_social=razao_social, operacao='saida')

        sleep(1)

        print('Mudar pesquisa para tomador')
        requisitar_cte.definir_tipo_pesquisa(tipo=TOMADOR)

        sleep(1)

        requisitar_cte.solicitar_relatorio(cnpj=cnpj, competencia=mes_ano, razao_social=razao_social, operacao='entrada')

    


    finally:

        pass



