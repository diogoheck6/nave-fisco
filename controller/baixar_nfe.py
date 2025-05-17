from time import sleep


def baixar_nfes(servico_consulta_nfe_nfce, cnpj, data_inicial, data_final, driver, wait):
    print('Baixando Notas de Saída NFE')

    print('Selecionando tipo de nota (NFe)')
    servico_consulta_nfe_nfce.selecionar_nfe()

    sleep(1)

    print('Inserindo CNPJ emitente')
    servico_consulta_nfe_nfce.selecionar_emitente(cnpj)

    sleep(1)

    print('Inserindo data inicial')
    servico_consulta_nfe_nfce.inserir_data_emissao_inicial(data_inicial)

    sleep(1)

    print('Inserindo data final')
    servico_consulta_nfe_nfce.inserir_data_emissao_final(data_final)

    sleep(1)

    print('Resolvendo Captcha')
    captcha = servico_consulta_nfe_nfce.solve_captcha(driver, wait)

    sleep(1)
    
    print('Inserindo Captcha')
    servico_consulta_nfe_nfce.inserir_captcha(captcha)
     
    sleep(1)

    print('Exportando relatório')
    servico_consulta_nfe_nfce.clicar_botao_exportar()

    # if servico_consulta_nfe_nfce.verificar_se_esta_disponivel_input_captha():
    sleep(1)

    print('Apagando CNPJ do emitente')
    servico_consulta_nfe_nfce.limpar_emitente()

    sleep(1)

    print('Inserindo CNPJ no destinatário')
    servico_consulta_nfe_nfce.selecionar_destinatario(cnpj)

    print('Baixando notas de Entrada NFE')

    sleep(1)

    servico_consulta_nfe_nfce.clicar_botao_exportar()