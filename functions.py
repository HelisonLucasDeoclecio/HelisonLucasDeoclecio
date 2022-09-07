def importAgency():
    """Função responsável por criar o arquivo de agências e despesas.
    
        * Inicialmente cria o arquivo de "Agências.xlsx" na pasta output com os cabeçalhos Agência e Despesas, na ordem.
        * Após, abre a página diretamente: IT Portfolio Dashboard, coleta todos os itens disponíveis no select de filtro de agências.
        * Por fim, percorre todo os itens do filtro para coletar os valores das despesas, além de registar na planilha.
    """
    try:
        # IMPORTAÇÕES
        from RPA.Browser.Selenium import Selenium
        from RPA.Excel.Files import Files

        # PREPARAÇÃO E CRIAÇÃO DO ARQUIVO EXCEL DE AGÊNCIAS
        excelFile = Files().create_workbook('output/Agências.xlsx', sheet_name='Agência')
        excelFile.set_cell_value(1, 'A', 'Agência')
        excelFile.set_cell_value(1, 'B', 'Despesas')

        # ABERTURA, COLETA DAS AGÊNCIAS E DESPESAS, REGISTRO NO ARQUIVO EXCEL
        browser = Selenium()
        browser.open_chrome_browser('https://www.itdashboard.gov/itportfoliodashboard')
        agencyItems = browser.get_list_items("id:agency-select", False)
        for item in agencyItems:
            browser.select_from_list_by_label("id:agency-select", item)
            browser.wait_until_element_is_not_visible('//*[@id="block-data-visualizer-content"]/div[2]/img')
            valueSpending = browser.get_text('//*[@class="it-spending"]/p/strong')
            
            numberEmptyRow = excelFile.find_empty_row()
            excelFile.set_cell_value(numberEmptyRow, 'A', item)
            excelFile.set_cell_value(numberEmptyRow, 'B', valueSpending)
            excelFile.save()
        
        excelFile.close()
        browser.close_all_browsers()
        
        return True
    
    except Exception as eGenericErro:
        # IMAGINA QUE A EXCEPTION -> vErro ESTÁ SENDO PASSADA PARA UM REGISTRADOR DE LOGS
        return False

def individualInvestments():
    """Função responsável por coletar os detalhes de inventimentos e registrá-los em planilha, além de baixar o PDF de casos de sucesso.
    
        * Inicialmente observa o arquivo 'config.json' para coletar em qual departamento irá coletar os investimentos.
        * Após, prepara a planilha "Agências.xlsx" criando uma nova aba contendo os dados coletados na página "Investment Details".
        * Por fim, nos investimentos em que há "Caso de sucesso" é aberto a página e salva em PDF também na pasta output.
    """
    try:
        # IMPORTAÇÕES
        from RPA.Browser.Selenium import Selenium
        from RPA.JSON import JSON
        from RPA.Excel.Files import Files

        jsonFile = JSON().load_json_from_file('config.json')
        agency = jsonFile['agency']

        # COLETA DOS NOMES DOS INVESTIDORES+ID    
        browser = Selenium()
        browser.open_chrome_browser('https://www.itdashboard.gov/investment-details')
        browser.select_from_list_by_label('id:agency-filter', agency)
        investments = str(browser.get_text('id:investment_select')).splitlines()
        browser.close_browser()

        # PREPARAÇÃO DE PLANILHA
        excelFile = Files().open_workbook('output/Agências.xlsx')
        excelFile.create_worksheet(agency)
        excelFile.set_cell_value(1, 'A', 'Investment Name')
        excelFile.set_cell_value(1, 'B', 'UII')
        excelFile.set_cell_value(1, 'C', 'Agency')
        excelFile.set_cell_value(1, 'D', 'Investment Type')
        excelFile.set_cell_value(1, 'E', 'IT Spending')
        excelFile.set_cell_value(1, 'F', 'CIO Rating')
        excelFile.set_cell_value(1, 'G', 'Date As Of')
        for investment in investments:
            # PREPARAÇÃO DOS DADOS
            nPosition1 = investment.find(' / ')+3
            nPosition2 = investment.find('<span')
            nameInvestment = investment[0:nPosition1-2]
            numberInvestment = investment[nPosition1:nPosition2]

            # VERIFICA SE O BOTÃO DE BUSINESS CASE EXISTE OU NÃO
            print('Abrindo a página: '+numberInvestment)
            browser.open_chrome_browser('https://www.itdashboard.gov/investment-details/'+numberInvestment, headless=True)
            businessCaseCondition = browser.is_element_visible('//*[@id="block-data-visualizer-content"]/div/a[1]')
            if businessCaseCondition == False:
                # NÃO EXISTINDO, COLETA OS DADOS COMUNS E REGISTRA EM PLANILHA
                agency = browser.get_text('//*[@id="block-data-visualizer-content"]/div/div[4]/div[1]/div[1]/span[2]')
                investmentType = browser.get_text('//*[@id="block-data-visualizer-content"]/div/div[4]/div[1]/div[3]/span[2]')
                itSpending = browser.get_text('//*[@id="block-data-visualizer-content"]/div/div[4]/div[2]/div[1]/span[2]')
                cioRating = browser.get_text('//*[@id="block-data-visualizer-content"]/div/div[4]/div[2]/div[2]/span[2]')
                dateAsOf = str(browser.get_text('//*[@id="block-data-visualizer-content"]/div/div[4]/div[2]/div[3]'))[12:]
                
                numberEmptyRow = excelFile.find_empty_row()
                excelFile.set_cell_value(numberEmptyRow, 'A', nameInvestment)
                excelFile.set_cell_value(numberEmptyRow, 'B', numberInvestment)
                excelFile.set_cell_value(numberEmptyRow, 'C', agency)
                excelFile.set_cell_value(numberEmptyRow, 'D', investmentType)
                excelFile.set_cell_value(numberEmptyRow, 'E', itSpending)
                excelFile.set_cell_value(numberEmptyRow, 'F', cioRating)
                excelFile.set_cell_value(numberEmptyRow, 'G', dateAsOf)
                excelFile.save()
            else:
                # ABRE A PÁGINA DE BUSINESS CASE
                browser.click_element('//*[@id="block-data-visualizer-content"]/div/a[1]')
                # IMPRIMIR A PÁGINA PARA GERAR O PDF
                browser.switch_window('NEW')
                browser.print_to_pdf(f'output/{numberInvestment}.pdf')

            browser.close_browser()
        
        return True
    
    except Exception as eGenericErro:
        # IMAGINA QUE A EXCEPTION -> vErro ESTÁ SENDO PASSADA PARA UM REGISTRADOR DE LOGS
        return False

def readExtractPdf():
    """Função responsável por extrair o nome do investidor e o UII dos PDF's de casos de sucesso obtidos da página.
    
        * Inicialmente pecorre todos os arquivos com formato PDF da pasta output.
        * Após, através de uma busca por palavras chave identifica e extra o conteúdo necessário.
    """
    try:
        # IMPORTAÇÕES
        from RPA.PDF import PDF
        import os

        files = os.listdir('output/')
        result = []
        for file in files:
            if file.endswith('.pdf'):
                contentPdf = str(PDF().get_text_from_pdf(f'output/{file}'))

                nPosition1 = contentPdf.find('Investment Name')+len('Investment Name')
                nPosition2 = contentPdf.find('Unique Investment Identiﬁer')
                investmentNamePdf = contentPdf[nPosition1:nPosition2]

                nPosition1 = contentPdf.find('Unique Investment Identiﬁer')+len('Unique Investment Identiﬁer')
                nPosition2 = contentPdf.find('Investment Description')
                uiiPdf = contentPdf[nPosition1:nPosition2]

                result.append({
                    file: [investmentNamePdf, uiiPdf]
                })
        return result
    
    except Exception as eGenericErro:
        # IMAGINA QUE A EXCEPTION -> vErro ESTÁ SENDO PASSADA PARA UM REGISTRADOR DE LOGS
        return False