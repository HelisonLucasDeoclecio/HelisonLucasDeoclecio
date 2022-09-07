try:
    from functions import *

    task = importAgency()
    if task == False:
        raise Exception("Erro ao chamar a task importAgency")

    task = individualInvestments()
    if task == False:
        raise Exception("Erro ao chamar a task individualInvestments")

    contentPdf = readExtractPdf()
    if contentPdf == False:
        raise Exception("Erro ao chamar a task readExtractPdf")

except Exception as eErro:
    # TRATAMENTO DE ERRO, ENVIADO PARA LOG...
    print(eErro)