#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Dec  6 01:14:18 2020

@author: rodrigo
"""

###### Pesquisa de registro no site da Anvisa

# Instalações Necessárias
#pip install -U selenium
#pip install pandas
#pip install openpyxl
#pip install xlrd

# Importações
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
import numpy as np

# Importar Lista de registros a serem consultados
ListaParaPesquisa = pd.read_excel("ListaParaPesquisaTeste.xlsx")
ListaParaPesquisa
len(ListaParaPesquisa)

# Escolher o Firefox
driver = webdriver.Firefox()

# Loop
print("Iniciando a Pesquisa")
for n in range(0,len(ListaParaPesquisa)):
    
    # Criar o DataFrame vazio para compilar os resultados
    DadosPesquisaAgregada = pd.DataFrame([],columns=['empresa',
                                  'cnpj',
                                  'processo',
                                  'categoria',
                                  'dataReg',
                                  'NomeComerc',
                                  'registro_P',
                                  'VencReg',
                                  'Principio',
                                  'Referencia',
                                  'CT',
                                  'ATC',
                                  'N_apres',
                                  'registro_A',
                                  'apresent',
                                  'forma',
                                  'substancia',
                                  'validade',
                                  'DT_Pub_A',
                                  'Complemento',
                                  'Embalagem',
                                  'LocalFab',
                                  'ViaAdm',
                                  'Conservacao',
                                  'restricao',
                                  'destinacao',
                                  'tarja',
                                  'Fracionado',
                                  'link'])
  
    try:
           
        # Pesquisa pelo registro de produto (9 dígitos)
        driver.get("https://consultas.anvisa.gov.br/#/medicamentos/")
        time.sleep(5)
        print(driver.title)
        search_bar = driver.find_element_by_id("txtNumeroRegistro")
        search_bar.clear()
        a = str(ListaParaPesquisa.iloc[n,0])
        search_bar.send_keys(a)
        search_bar.send_keys(Keys.RETURN)
        time.sleep(3)
        print(driver.current_url)
        
        # Mostrar apresentações
        search_2 = driver.find_element_by_xpath('//*[@id="containerTable"]/table/tbody/tr[2]/td[4]')
        search_2.click()
        time.sleep(8)
        print(driver.current_url)
        
        # Mostrar detalhes das apresentações
        search_3 = driver.find_element_by_tag_name("a.btn.btn-default.no-print.ng-scope")
        search_3.click()
        time.sleep(2)
        print(driver.current_url)
        
        # Copiar dados das apresentações
        i = 1
        while True:
            empresa     = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[1]/div[2]/table/tbody/tr[1]/td[1]').text
            cnpj        = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[1]/div[2]/table/tbody/tr[1]/td[2]').text
            processo    = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[1]/div[2]/table/tbody/tr[2]/td[1]').text
            categoria   = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[1]/div[2]/table/tbody/tr[2]/td[2]').text
            dataReg     = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[1]/div[2]/table/tbody/tr[2]/td[3]').text
            NomeComerc  = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[1]/div[2]/table/tbody/tr[3]/td[1]').text
            registro_P  = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[1]/div[2]/table/tbody/tr[3]/td[2]').text
            VencReg     = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[1]/div[2]/table/tbody/tr[3]/td[3]').text
            Principio   = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[1]/div[2]/table/tbody/tr[4]/td[1]').text
            Referencia  = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[1]/div[2]/table/tbody/tr[4]/td[2]').text
            CT          = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[1]/div[2]/table/tbody/tr[5]/td[1]').text
            ATC         = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[1]/div[2]/table/tbody/tr[5]/td[2]').text
                   
            N_apres     = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[3]/div[1]/table[' + str(i) + ']/tbody/tr[2]/td[1]').text
            registro_A  = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[3]/div[1]/table[' + str(i) + ']/tbody/tr[2]/td[3]').text
            apresent    = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[3]/div[1]/table[' + str(i) + ']/tbody/tr[2]/td[2]').text
            forma       = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[3]/div[1]/table[' + str(i) + ']/tbody/tr[2]/td[4]').text
            substancia  = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[3]/div[1]/table[' + str(i) + ']/tbody/tr[3]/td').text
            validade    = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[3]/div[1]/table[' + str(i) + ']/tbody/tr[2]/td[6]').text
            DT_Pub_A    = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[3]/div[1]/table[' + str(i) + ']/tbody/tr[2]/td[5]').text
            Complemento = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[3]/div[1]/table[' + str(i) + ']/tbody/tr[4]/td').text
            Embalagem   = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[3]/div[1]/table[' + str(i) + ']/tbody/tr[5]/td').text
            LocalFab    = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[3]/div[1]/table[' + str(i) + ']/tbody/tr[6]/td').text
            ViaAdm      = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[3]/div[1]/table[' + str(i) + ']/tbody/tr[7]/td').text
            Conservacao = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[3]/div[1]/table[' + str(i) + ']/tbody/tr[8]/td').text
            restricao   = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[3]/div[1]/table[' + str(i) + ']/tbody/tr[9]/td').text
            destinacao  = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[3]/div[1]/table[' + str(i) + ']/tbody/tr[10]/td').text
            tarja       = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[3]/div[1]/table[' + str(i) + ']/tbody/tr[11]/td').text
            Fracionado  = driver.find_element_by_xpath('/html/body/div[3]/div[1]/form/div[3]/div[1]/table[' + str(i) + ']/tbody/tr[12]/td').text
            
            link = driver.current_url
            
            j = 0
            
            DadosZerados = np.zeros([1,29])
            DadosPesquisa = pd.DataFrame(DadosZerados,columns=['empresa',
                                          'cnpj',
                                          'processo',
                                          'categoria',
                                          'dataReg',
                                          'NomeComerc',
                                          'registro_P',
                                          'VencReg',
                                          'Principio',
                                          'Referencia',
                                          'CT',
                                          'ATC',
                                          'N_apres',
                                          'registro_A',
                                          'apresent',
                                          'forma',
                                          'substancia',
                                          'validade',
                                          'DT_Pub_A',
                                          'Complemento',
                                          'Embalagem',
                                          'LocalFab',
                                          'ViaAdm',
                                          'Conservacao',
                                          'restricao',
                                          'destinacao',
                                          'tarja',
                                          'Fracionado',
                                          'link'])
    
            DadosPesquisa.iloc[j,0] = empresa
            DadosPesquisa.iloc[j,1] = cnpj
            DadosPesquisa.iloc[j,2] = processo
            DadosPesquisa.iloc[j,3] = categoria
            DadosPesquisa.iloc[j,4] = dataReg
            DadosPesquisa.iloc[j,5] = NomeComerc
            DadosPesquisa.iloc[j,6] = registro_P
            DadosPesquisa.iloc[j,7] = VencReg
            DadosPesquisa.iloc[j,8] = Principio
            DadosPesquisa.iloc[j,9] = Referencia
            DadosPesquisa.iloc[j,10] = CT
            DadosPesquisa.iloc[j,11] = ATC
            DadosPesquisa.iloc[j,12] = N_apres
            DadosPesquisa.iloc[j,13] = registro_A
            DadosPesquisa.iloc[j,14] = apresent
            DadosPesquisa.iloc[j,15] = forma
            DadosPesquisa.iloc[j,16] = substancia
            DadosPesquisa.iloc[j,17] = validade
            DadosPesquisa.iloc[j,18] = DT_Pub_A
            DadosPesquisa.iloc[j,19] = Complemento
            DadosPesquisa.iloc[j,20] = Embalagem
            DadosPesquisa.iloc[j,21] = LocalFab
            DadosPesquisa.iloc[j,22] = ViaAdm
            DadosPesquisa.iloc[j,23] = Conservacao
            DadosPesquisa.iloc[j,24] = restricao
            DadosPesquisa.iloc[j,25] = destinacao
            DadosPesquisa.iloc[j,26] = tarja
            DadosPesquisa.iloc[j,27] = Fracionado
            DadosPesquisa.iloc[j,28] = link
            
            DadosPesquisaAgregada = DadosPesquisaAgregada.append(DadosPesquisa)
        
            i = i + 1

    except:
        #print("Em processo de gravação de dados")

        # Criar objeto para leitura e selecionar planilha
        excel_reader = pd.ExcelFile('Resultados.xlsx')
        to_update = {"Planilha2": DadosPesquisaAgregada}
        
        # Criar objeto para escrita
        excel_writer = pd.ExcelWriter('Resultados.xlsx')
        sheet_df = excel_reader.parse("Planilha2")
        append_df = to_update.get("Planilha2")
        
        #concatenar com o que já existia
        sheet_df = pd.concat([sheet_df, DadosPesquisaAgregada]).drop_duplicates()
        
    	# Gravar no arquivo
        sheet_df.to_excel(excel_writer, "Planilha2", index=False)
        
        # Salvar e fechar arquivo
        excel_writer.save()    
        
        # Limpar o DataFrame para nova gravação
        DadosPesquisaAgregada.drop(columns=['empresa',
                                              'cnpj',
                                              'processo',
                                              'categoria',
                                              'dataReg',
                                              'NomeComerc',
                                              'registro_P',
                                              'VencReg',
                                              'Principio',
                                              'Referencia',
                                              'CT',
                                              'ATC',
                                              'N_apres',
                                              'registro_A',
                                              'apresent',
                                              'forma',
                                              'substancia',
                                              'validade',
                                              'DT_Pub_A',
                                              'Complemento',
                                              'Embalagem',
                                              'LocalFab',
                                              'ViaAdm',
                                              'Conservacao',
                                              'restricao',
                                              'destinacao',
                                              'tarja',
                                              'Fracionado',
                                              'link'])

        sheet_df.drop(columns=['empresa',
                               'cnpj',
                               'processo',
                               'categoria',
                               'dataReg',
                               'NomeComerc',
                               'registro_P',
                               'VencReg',
                               'Principio',
                               'Referencia',
                               'CT',
                               'ATC',
                               'N_apres',
                               'registro_A',
                               'apresent',
                               'forma',
                               'substancia',
                               'validade',
                               'DT_Pub_A',
                               'Complemento',
                               'Embalagem',
                               'LocalFab',
                               'ViaAdm',
                               'Conservacao',
                               'restricao',
                               'destinacao',
                               'tarja',
                               'Fracionado',
                               'link'])
    
        print("Em processo de gravação de dados")
        
    print("Próximo registro")
    
print("Fim da Gravação")
quit()