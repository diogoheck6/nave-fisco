from datetime import datetime
from browser.connection import start_browser
from login.login import login
from models.Service_Consult_NFE_NFCE import ServiceConsultNFENFCE
from controller.baixar_nfe import baixar_nfes
from controller.baixar_nfce import baixar_nfce
from controller.solicitar_cte import solicitar_ctes
from controller.read_files_xlsx_v1 import read_files_and_query
from controller.gerar_excel_ import gerar_excel
from controller.baixar_cte import baixar_relat_ctes
from models.db_sqlite_v2 import NumerosSolicitacaoCTE
import os

class ProcessarDados:
    def __init__(self, cpf, senha, cnpj, data_inicial, data_final, dicionario_clientes):
        self.cpf = cpf
        self.senha = senha
        self.cnpj = cnpj
        self.data_inicial_text = data_inicial
        self.data_inicial = datetime.strptime(data_inicial, '%d/%m/%Y').strftime('%d%m%Y')
        self.data_final = datetime.strptime(data_final, '%d/%m/%Y').strftime('%d%m%Y')
        self.start_date = datetime.strptime(data_inicial, '%d/%m/%Y').strftime('%Y-%m-%d')
        self.end_date = datetime.strptime(data_final, '%d/%m/%Y').strftime('%Y-%m-%d')
        self.dicionario_clientes = dicionario_clientes

    def executar(self):
        
        print('CNPJ--->', self.cnpj)
        if self.cnpj not in self.dicionario_clientes:
            raise Exception("CNPJ não encontrado nos registros.")

        obj_base = NumerosSolicitacaoCTE()
        self.remover_arquivos()

        razao_social = self.dicionario_clientes[self.cnpj]
        print(f"Processando CNPJ: {self.cnpj}, Razão Social: {razao_social}")
        
        from browser.connection import start_browser
        from login.login import login

        driver, wait = start_browser()
        login(driver, wait, self.cpf, self.senha)

        print('data inicial', self.data_inicial)

        dia, mes, ano = self.data_inicial_text.split('/')
        mes_ano = f"{mes}/{ano}"


        if not obj_base.deve_pular_solicitacao_cte(self.cnpj, mes_ano):
            solicitar_ctes(driver, wait, self.cnpj, razao_social, mes_ano)

        servico_consulta_nfe_nfce = ServiceConsultNFENFCE(
            link='https://sat.sef.sc.gov.br/tax.NET/Sat.NFe.Web/Consultas/ConsultaOnlineCC.aspx',
            driver=driver,
            wait=wait
        )

        servico_consulta_nfe_nfce.acessar_link()
        baixar_nfes(servico_consulta_nfe_nfce, self.cnpj, self.data_inicial, self.data_final, driver, wait)
        servico_consulta_nfe_nfce.limpar_inputs()
        baixar_nfce(servico_consulta_nfe_nfce, self.cnpj, self.data_inicial, self.data_final, driver, wait)

        baixar_relat_ctes(driver, wait, self.cnpj, razao_social, mes_ano)

        print('Extratos Notas baixados com Sucesso \\o/\\o/')

        entradas_sat, saidas_sat, entradas_questor, saidas_questor, only_in_sat, only_in_sat_entradas, \
            fechamento_saidas, fechamento_entradas, teve_diferenca_saidas, teve_diferenca_entradas, pintura_verde_entradas = read_files_and_query(self.cnpj, self.start_date, self.end_date)

        driver.quit()

        print('gerando excel')
        caminho_completo = gerar_excel(pintura_verde_entradas, entradas_sat, saidas_sat, entradas_questor, saidas_questor, only_in_sat, only_in_sat_entradas, 
            fechamento_saidas, fechamento_entradas, teve_diferenca_saidas, teve_diferenca_entradas, 
                self.cnpj.replace('.', '').replace('/', '').replace('-', ''), razao_social)
        print('excel gerado!!!')

        return f"Relatório gerado com sucesso para {caminho_completo}"

    def remover_arquivos(self):
        extensoes = ['.xlsx', '.zip', '.html', '.xls', '.jpeg']
        for arquivo in os.listdir(os.getcwd()):
            if any(arquivo.endswith(ext) for ext in extensoes):
                os.remove(arquivo)