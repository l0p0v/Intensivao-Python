import pandas as pd
from selenium import webdriver  # permite criar o navegador
from selenium.webdriver.common.by import By  # permite selecionar items no navegador
from selenium.webdriver.common.keys import Keys  # permite escrever no navegador
import os

CWD = os.getcwd()
ARQUIVO = 'Produtos Novo.xlsx'


def exibir_cotacao(cotacao: float):
    return str(round(cotacao, 2)).replace('.', ',')


# opções do navegador
options = webdriver.ChromeOptions()
options.add_argument("--headless")  # rodar em segundo plano

# abrir o navegador
navegador = webdriver.Chrome(options=options)

# Passo 1: Pegar a cotação do dólar
# entrar no Google
navegador.get('https://www.google.com/')

# pesquisar no Google por "cotação dólar"
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(
    'cotação do dólar')
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(
    Keys.ENTER)

# pegar a contação que tá no Google
cotacao_dolar = float(navegador.find_element(By.XPATH,
                                             '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div['
                                             '2]/span[1]').get_attribute('data-value'))
print(f'$1 = R${exibir_cotacao(cotacao_dolar)}')

# Passo 2: Pegar a cotação do euro
# entrar no Google
navegador.get('https://www.google.com/')

# pesquisar no Google por "cotação euro"
navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(
    'cotação do euro')
navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(
    Keys.ENTER)

# pegar a contação que tá no Google
cotacao_euro = float(navegador.find_element(By.XPATH,
                                            '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div['
                                            '2]/span[1]').get_attribute('data-value'))
print(f'€1 = R${exibir_cotacao(cotacao_euro)}')

# Passo 3: Pegar a cotação do ouro
navegador.get('https://www.melhorcambio.com/ouro-hoje')

# pegar a cotação que tá no site
cotacao_ouro = navegador.find_element('xpath', '//*[@id="comercial"]').get_attribute('value')
cotacao_ouro = float(cotacao_ouro.replace(',', '.'))
print(f'1g de ouro = R${exibir_cotacao(cotacao_ouro)}')

# Extra: pegar a cotação da libra esterlina
# entrar no Google
navegador.get('https://www.google.com')

# pesquisar no Google "cotação libra esterlina"
navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(
    'cotação libra esterlina')
navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(
    Keys.ENTER)

# pegar a cotação que tá no Google
cotacao_libra = float(navegador.find_element(By.XPATH,
                                             '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div['
                                             '2]/span[1]').get_attribute('data-value'))
print(f'£1 = R${exibir_cotacao(cotacao_libra)}')

navegador.quit()

# Agora vamos atualiza a nossa base de preços com as novas cotações

# Importando a base de dados

# Passo 4: Atualizar a base de dados (atualizando o preço de compra e o de venda)
df = pd.read_excel('Produtos.xlsx')
# display(df)

# Atualizando os preços e o cálculo do Preço Final

# atualizar a coluna de cotação
df.loc[df['Moeda'] == 'Dólar', 'Cotação'] = cotacao_dolar
df.loc[df['Moeda'] == 'Euro', 'Cotação'] = cotacao_euro
df.loc[df['Moeda'] == 'Ouro', 'Cotação'] = cotacao_ouro

# atualizar a coluna de preço de compra
df.loc[df['Moeda'] == 'Dólar', 'Preço de Compra'] = df['Preço Original'] * df['Cotação']
df.loc[df['Moeda'] == 'Euro', 'Preço de Compra'] = df['Preço Original'] * df['Cotação']
df.loc[df['Moeda'] == 'Ouro', 'Preço de Compra'] = df['Preço Original'] * df['Cotação']

# atualizar a coluna de preço de venda
df['Preço de Venda'] = df['Preço de Compra'] * df['Margem']

# Formatando valores no dataframe

df['Cotação'] = round(df['Cotação'], 2)
df['Preço de Compra'] = round(df['Preço de Compra'], 2)
df['Preço de Venda'] = round(df['Preço de Venda'], 2)

# display(df)

# Agora vamos exportar a nova base de preços atualizada

# Passo 5: Exportar a base de preços atualizada
df.to_excel(ARQUIVO, index=False)

print('\nNova base de dados criada!')
print(f'Arquivo salvo em {CWD}\\{ARQUIVO}')
input()
