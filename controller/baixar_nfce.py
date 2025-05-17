from time import sleep


def baixar_nfce(servico_consulta_nfe_nfce, cnpj, data_inicial, data_final, driver, wait):

    print('Baixando Notas de Saída NFCe')

    print('Selecionando tipo de nota (NFCe)')
    servico_consulta_nfe_nfce.selecionar_nfce()

    sleep(1)

    print('Inserindo CNPJ emitente')
    servico_consulta_nfe_nfce.selecionar_emitente(cnpj)

    sleep(1)

    print('Baixando notas de Saida NFCe')

    servico_consulta_nfe_nfce.clicar_botao_exportar()
    
    sleep(5)

    