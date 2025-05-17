import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from dateutil.relativedelta import relativedelta
import threading
from ttkthemes import ThemedStyle
from browser.connection import start_browser
from login.login import login
from models.Service_Consult_NFE_NFCE import ServiceConsultNFENFCE
from controller.baixar_nfe import baixar_nfes
from time import sleep
from controller.baixar_nfce import baixar_nfce
from controller.solicitar_cte import solicitar_ctes
from controller.ler_planilha_clientes import ler_planilha_excel
from controller.read_files_xlsx_v1 import read_files_and_query
from controller.gerar_excel_ import gerar_excel
from controller.baixar_cte import baixar_relat_ctes
from models.db_sqlite_v2 import NumerosSolicitacaoCTE
from controller.consulta_cte import ServiceConsultCTE
import os
import sys


class RedirectText:
    def __init__(self, text_widget):
        self.output = text_widget

    def write(self, string):
        self.output.insert(tk.END, string)
        self.output.see(tk.END)

    def flush(self):
        pass


def remover_arquivos():
    extensoes = ['.xlsx', '.zip', '.html', '.xls', '.jpeg']
    diretorio_atual = os.getcwd()
    planilha_cliente = "Clientes Navecon - SC com IE.xlsx"
    
    for arquivo in os.listdir(diretorio_atual):
        # Verifica se o arquivo é do tipo a ser removido e não é a planilha de clientes
        if any(arquivo.endswith(ext) for ext in extensoes) and arquivo != planilha_cliente:
            caminho_arquivo = os.path.join(diretorio_atual, arquivo)
            os.remove(caminho_arquivo)
            # print(f"Removido: {arquivo}")


def limpar_log():
    log_area.delete('1.0', tk.END)

def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

def processar_dados(cpf, senha, cnpj_selected, data_inicial_text, data_final_text, dicionario_clientes, mais_de_500_ctes):

    try:

        print('Processar mais de 500 CTEs? ', 'Sim' if mais_de_500_ctes else 'Não')

        remover_arquivos()

        cnpj_razao = cnpj_selected.split(" - ")
        cnpj = cnpj_razao[0].strip() if len(cnpj_razao) > 0 else ""

        data_inicial = datetime.strptime(data_inicial_text, '%d/%m/%Y').strftime('%d%m%Y')
        data_final = datetime.strptime(data_final_text, '%d/%m/%Y').strftime('%d%m%Y')
        start_date = datetime.strptime(data_inicial_text, '%d/%m/%Y').strftime('%Y-%m-%d')
        end_date = datetime.strptime(data_final_text, '%d/%m/%Y').strftime('%Y-%m-%d')

        if cnpj in dicionario_clientes:

            razao_social = dicionario_clientes[cnpj]
            print(f"Processando CNPJ: {cnpj}, Razão Social: {razao_social}")

            driver, wait = start_browser()  # Simulação
            login(driver, wait, cpf, senha)  # Simulação

            dia, mes, ano = data_inicial_text.split('/')
            mes_ano = f"{mes}/{ano}"
            competencia = f"{mes}-{ano}"

            if mais_de_500_ctes:

                obj_base = NumerosSolicitacaoCTE()

                if not obj_base.deve_pular_solicitacao_cte(cnpj, mes_ano):
                    solicitar_ctes(driver, wait, cnpj, razao_social, mes_ano)  # Simulação

                baixar_relat_ctes(driver, wait, cnpj, razao_social, mes_ano)

            servico_consulta_nfe_nfce = ServiceConsultNFENFCE(
                link='https://sat.sef.sc.gov.br/tax.NET/Sat.NFe.Web/Consultas/ConsultaOnlineCC.aspx',
                driver=driver,
                wait=wait
            )

            servico_consulta_nfe_nfce.acessar_link()
            baixar_nfes(servico_consulta_nfe_nfce, cnpj, data_inicial, data_final, driver, wait)
            servico_consulta_nfe_nfce.limpar_inputs()
            baixar_nfce(servico_consulta_nfe_nfce, cnpj, data_inicial, data_final, driver, wait)

            if not mais_de_500_ctes:

                consultar_cte_normal = ServiceConsultCTE(
                    link='https://sat.sef.sc.gov.br/tax.net/Sat.Dfe.Cte.Web/ConsultaCte.aspx',
                    driver=driver,
                    wait=wait
                )

                consultar_cte_normal.consultar_ctes(cnpj, mes_ano)

            sleep(5)

            print('Extratos Notas baixados com Sucesso \\o/\\o/')
    
            entradas_sat, saidas_sat, entradas_questor, saidas_questor, only_in_sat, only_in_sat_entradas,\
                  only_in_questor_entradas, fechamento_saidas, fechamento_entradas, teve_diferenca_saidas, \
                     teve_diferenca_entradas, pintura_verde_entradas, pintura_verde_questor,only_in_questor_saidas, pintura_verde_questor_sat,\
                        intersecao_saidas, intersecao_entradas, pintar_vermelho_saidas, pintar_vermelho_entradas = \
                read_files_and_query(mais_de_500_ctes, cnpj, start_date, end_date)
            
            
            caminho_completo = gerar_excel(pintura_verde_entradas, pintura_verde_questor, entradas_sat, saidas_sat, entradas_questor, saidas_questor, only_in_sat,
                        only_in_sat_entradas, only_in_questor_entradas, fechamento_saidas, fechamento_entradas,
                        teve_diferenca_saidas,teve_diferenca_entradas, cnpj.replace('.', '').replace('/', '').replace('-', ''),
                        razao_social, only_in_questor_saidas, pintura_verde_questor_sat, intersecao_saidas, intersecao_entradas, \
                            pintar_vermelho_saidas, pintar_vermelho_entradas, competencia)
            

            driver.quit()
            messagebox.showinfo("Sucesso", f"Relatório gerado com sucesso")

        else:
            messagebox.showerror("Erro", "CNPJ não encontrado nos registros.")

    except Exception as e:
        if 'driver' in locals():
            pass
            driver.quit()
        messagebox.showerror("Erro", f"Erro ao processar dados: {str(e)}")
    finally:
        resetar_interface()
        driver.quit()
 
 

def iniciar_processamento():
    processar_btn.config(text="Processando...", state='disabled')
    cpf = cpf_entry.get()
    senha = senha_entry.get()
    cnpj = cnpj_combo.get()
    data_inicial = data_inicial_entry.get()
    data_final = data_final_entry.get()
    mais_de_500_ctes = cte_var.get() # Pega o valor do Checkbutton relacionado.
    
    # Usando thread para prosseguir com novos processamentos
    t = threading.Thread(target=processar_dados, args=(cpf, senha, cnpj, data_inicial, data_final, dicionario_clientes, mais_de_500_ctes), daemon=True)
    t.start()

    # Monitorar a thread de processamento
    monitor_thread(t)

def monitor_thread(t):
    if t.is_alive():
        app.after(100, lambda: monitor_thread(t))  # Atualiza após 100 ms
    else:
        resetar_interface()

def filtrar_clientes():
    criterio = search_entry.get().lower()
    valores_filtrados = [
        f"{cnpj} - {razao}" for cnpj, razao in dicionario_clientes.items()
        if criterio in cnpj.lower() or criterio in razao.lower()
    ]
    cnpj_combo['values'] = valores_filtrados
    if valores_filtrados:
        cnpj_combo.set(valores_filtrados[0])

def resetar_interface():
    processar_btn.config(text="Processar Dados", state='normal')

def minimizar_app():
    app.iconify()

def fechar_app():
    app.destroy()

# Inicialização da Aplicação
if __name__ == '__main__':
    CPF = '07954444945'
    SENHA = 'Navecon2025@'
    caminho_arquivo = r"\\192.168.5.254\arquivos\ARQUIVOS\06 - SETOR FISCAL\22- Automações\Clientes Navecon - SC com IE.xlsx"
    hoje = datetime.now()
    primeiro_dia_mes_anterior = (hoje - relativedelta(months=1)).replace(day=1)
    ultimo_dia_mes_anterior = (hoje.replace(day=1) - relativedelta(days=1))

    data_inicial_default = primeiro_dia_mes_anterior.strftime("%d/%m/%Y")
    data_final_default = ultimo_dia_mes_anterior.strftime("%d/%m/%Y")

    app = tk.Tk()
    app.title("NaveFisco")
    app.resizable(False, False)
    

    style = ThemedStyle(app)
    style.set_theme("equilux")
    app.configure(background='#1E1E1E')

    std_padx, std_pady = 20, 10
    entry_width = 50
    label_width = 25

    logo_image = tk.PhotoImage(file=resource_path("logo.png"))
    logo_label = ttk.Label(app, image=logo_image, background='#1E1E1E', anchor='center')
    logo_label.grid(row=0, column=0, columnspan=2, pady=20)

    fields = {
        "CPF": CPF, 
        "Senha": SENHA, 
        "Pesquisar Cliente": "", 
        "CNPJ - Razão Social": "", 
        "Data Inicial (dd/mm/aaaa)": data_inicial_default, 
        "Data Final (dd/mm/aaaa)": data_final_default
    }

    entries = {}
    cnpj_combo = None
    row_idx = 1

    for label_text, default in fields.items():
        label = ttk.Label(app, text=f"{label_text}:", background='#1E1E1E', foreground='white', width=label_width, anchor='e')
        label.grid(row=row_idx, column=0, padx=std_padx, pady=std_pady)
        
        
        if label_text == "CNPJ - Razão Social":
            # cnpj_combo = ttk.Combobox(app, state='readonly', width=entry_width)
            cnpj_combo = ttk.Combobox(app, state='readonly', width=48)
            entries[label_text] = cnpj_combo
            cnpj_combo.grid(row=row_idx, column=1, padx=std_padx, pady=std_pady)
            

        else:
            show_param = "*" if label_text == "Senha" else None
            entry = ttk.Entry(app, show=show_param, width=entry_width)
            entry.insert(0, default)
            entries[label_text] = entry
            entry.grid(row=row_idx, column=1, padx=std_padx, pady=std_pady)
            if label_text == "Pesquisar Cliente":
                entry.bind('<KeyRelease>', lambda event: filtrar_clientes())

        row_idx += 1

    cpf_entry = entries["CPF"]
    senha_entry = entries["Senha"]
    search_entry = entries["Pesquisar Cliente"]
    data_inicial_entry = entries["Data Inicial (dd/mm/aaaa)"]
    data_final_entry = entries["Data Final (dd/mm/aaaa)"]

    # dicionario_clientes = {}
    dicionario_clientes = ler_planilha_excel(caminho_arquivo)
    cnpj_combo['values'] = [f"{cnpj} - {razao}" for cnpj, razao in dicionario_clientes.items()]

    processar_btn = ttk.Button(app, text='Processar Dados', command=iniciar_processamento)
    processar_btn.grid(row=row_idx, column=0, columnspan=2, pady=20)

    log_area = tk.Text(app, wrap='word', foreground='green', background='black', height=10, width=entry_width)
    log_area.grid(row=row_idx+1, column=0, columnspan=2, padx=std_padx, pady=(std_pady, 10))

    redirect_std = RedirectText(log_area)
    sys.stdout = redirect_std

    limpar_log_btn = ttk.Button(app, text="Limpar Log", command=limpar_log)
    limpar_log_btn.grid(row=row_idx+2, column=0, columnspan=2, pady=(0, 10))

     # Crie um Frame para o rodapé
    rodape_frame = tk.Frame(app, background='#1E1E1E')
    rodape_frame.grid(row=row_idx+3, column=0, columnspan=2, pady=10, sticky='w')

    global cte_var
    # Variável BooleanVar para armazenar o estado do checkbutton
    cte_var = tk.BooleanVar()

    # Adicione o checkbutton no rodapé
    check_cte = ttk.Checkbutton(
        rodape_frame,
        text="+ 500 CTEs",
        variable=cte_var,
        style='TCheckbutton'
    )
    check_cte.grid(row=0, column=0, padx=std_padx, pady=std_pady)


    def on_enter(e, btn):
        btn.config(style='Highlight.TButton')
    
    def on_leave(e, btn):
        btn.config(style='TButton')

    style.configure('TButton', padding=5, relief='flat', borderwidth=0)
    style.configure('Highlight.TButton', background='red', foreground='white')

    app.mainloop()