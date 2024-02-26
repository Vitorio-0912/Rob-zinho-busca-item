from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import os
import PySimpleGUI as sg
import time
import openpyxl
import pandas as pd

# criação da janela de busca do produto
layout = [
        [sg.Text('Informe o produto que deseja pesquisar o preço: '), sg.InputText()],
        [sg.Button('Pesquisar'), sg.Button("Cancelar")]
]

janela = sg.Window('Busca preços', layout)

#acessando o site e validando se o user pesquisou ou cancelou
while True:
    event, values = janela.read()
    if event == 'Cancelar' or event == sg.WINDOW_CLOSED:
        break
    elif event == 'Pesquisar':
        navegador = webdriver.Chrome()
        wait = WebDriverWait(navegador, 12)
        navegador.get('https://shopping.google.com.br/')
        
        #clicando na barra de pesquisa
        campo_busca = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@type="search"]'  )))
        campo_busca.click()
         
        
        #digitanto e pesquisando para o google não perceber o bot
        produto = values[0]
        for letra in produto:
            campo_busca.send_keys(letra)
            time.sleep(0.2)
            
        campo_busca.send_keys(Keys.ENTER)
        
        #extrair o nome dos jogos
        titulos = navegador.find_elements(By.XPATH, '//div[@class="sh-np__seller-container"]')

        #extrair o preço dos jogos
        precos = navegador.find_elements(By.XPATH, '//b[@class="translate-content"]') 
        
        
        # Criando a planilha
        workbook = openpyxl.Workbook()
        # Criando a página "Produtos"
        workbook.create_sheet('Produtos')
        # Seleciono a página "Produtos"
        sheet_produtos = workbook['Produtos']
        sheet_produtos['A1'].value = 'lojas'
        sheet_produtos['B1'].value = 'preços'
        
        # Inserir os títulos na planilha (assumindo que você tenha definido 'titulos' e 'precos' anteriormente)
        for titulo, preco in zip(titulos, precos):
            sheet_produtos.append([titulo.text, preco.text])
            
        # Salvar a planilha
        workbook.save('Produtos.xlsx')
        
        # Ler a planilha
        excel = pd.read_excel('Produtos.xlsx', sheet_name='Produtos')

        # Encontrar a linha correspondente ao preço mais baixo
        linha_preco_mais_baixo = excel.loc[excel['preços'].idxmin()]

        # Obter o nome da loja e o preço mais baixo
        loja_mais_baixa = linha_preco_mais_baixo['lojas']
        preco_mais_baixo = linha_preco_mais_baixo['preços']
        
        # Criar layout para a saida
        layout = [
            [sg.Text(f'A loja mais barata é {loja_mais_baixa} com o preço de {preco_mais_baixo}.')],
            [sg.Text('Acesse a planilha abaixo: ')],
            [sg.Button('Acessar', button_color=('white', '#0077cc'), enable_events=True, key='Planilha')],
            [sg.Button('OK')]
            ]
        
      
        
        break  

# Criar a janela
window = sg.Window('Detalhes do Produto', layout)

# Loop de evento para manter a janela aberta
while True:
    event, values = window.read()

    if event == sg.WINDOW_CLOSED or event == 'OK':
            break
    elif event == 'Planilha':
        nome_arquivo = 'Produtos.xlsx'

        # Verifica se o arquivo existe antes de tentar abri-lo
        if os.path.exists(nome_arquivo):
            os.system(f'start excel "{nome_arquivo}"')  # Abre o arquivo com o aplicativo padrão do Excel
        else:
            print(f'O arquivo {nome_arquivo} não foi encontrado.')
window.close()


janela.close()