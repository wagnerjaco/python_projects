
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import pandas as pd
import openpyxl
from matplotlib import pyplot as plt
import calendar
from datetime import timedelta
from datetime import datetime, timedelta
# verificar se o arquivo exieste e crialo 
import os

# defino o nome do arquivo na mesma pasta do codigo 
file_path = ("Conjunto de jogos.xlsx")
#verificação se ele existe caso não, ele cria
if not os.path.exists (file_path):
    inni_jogos =  {}
    conju_jogos = pd.DataFrame(inni_jogos)
    conju_jogos.to_excel(file_path, index=False)
    
# variavel que vai armazenar a seguencia de dia 
results = {}
# data inincial e final que a pesquisa ira percorrer 
dia = datetime(2023, 1, 9)
fim = datetime(2023, 1, 13)
#condição para repetir as intruções 
while dia <= fim:
    if calendar.weekday(dia.year, dia.month, dia.day) < 5: # se o dia é de segunda a sexta
        # para rodar o chrome em 2º plano
        chrome_options = Options()
        chrome_options.headless = True
        navegador = webdriver.Chrome(options=chrome_options)
        #navegador = webdriver.Chrome() # abrir navegador em primeiro plano 
        navegador.get("https://www.google.com/")
        search_bar = navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input')
        search_bar.send_keys(f'lotofacil {dia.strftime("%d/%m/%Y")}')
        search_bar.send_keys(Keys.RETURN)
        #verifica o tempo de resposta em caso de erro fechae repete
        try:
            WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.ID, "result-stats")))
        except TimeoutException:
            print("Timed out waiting for page to load")
            navegador.quit()
        
        resultados = []
        # condição de repetição para ler todos os XPath  na pesquisa
        for i in range(1,9): 
            xpaths = ['//*[@id="tsuid_27"]/span/div/div/div/div/div[2]/div/div[1]/div/div[1]/div/div[1]/span[1]',
          '//*[@id="tsuid_27"]/span/div/div/div/div/div[2]/div/div[1]/div/div[1]/div/div[1]/span[2]', 
          '//*[@id="tsuid_27"]/span/div/div/div/div/div[2]/div/div[1]/div/div[1]/div/div[1]/span[3]', 
          '//*[@id="tsuid_27"]/span/div/div/div/div/div[2]/div/div[1]/div/div[1]/div/div[1]/span[4]', 
          '//*[@id="tsuid_27"]/span/div/div/div/div/div[2]/div/div[1]/div/div[1]/div/div[1]/span[5]', 
          '//*[@id="tsuid_27"]/span/div/div/div/div/div[2]/div/div[1]/div/div[1]/div/div[2]/span[1]', 
          '//*[@id="tsuid_27"]/span/div/div/div/div/div[2]/div/div[1]/div/div[1]/div/div[2]/span[2]', 
          '//*[@id="tsuid_27"]/span/div/div/div/div/div[2]/div/div[1]/div/div[1]/div/div[2]/span[3]',
          '//*[@id="tsuid_27"]/span/div/div/div/div/div[2]/div/div[1]/div/div[1]/div/div[2]/span[4]', 
          '//*[@id="tsuid_27"]/span/div/div/div/div/div[2]/div/div[1]/div/div[1]/div/div[2]/span[5]', 
          '//*[@id="tsuid_27"]/span/div/div/div/div/div[2]/div/div[1]/div/div[1]/div/div[3]/span[1]', 
          '//*[@id="tsuid_27"]/span/div/div/div/div/div[2]/div/div[1]/div/div[1]/div/div[3]/span[2]', 
          '//*[@id="tsuid_27"]/span/div/div/div/div/div[2]/div/div[1]/div/div[1]/div/div[3]/span[3]', 
          '//*[@id="tsuid_27"]/span/div/div/div/div/div[2]/div/div[1]/div/div[1]/div/div[3]/span[4]', 
          '//*[@id="tsuid_27"]/span/div/div/div/div/div[2]/div/div[1]/div/div[1]/div/div[3]/span[5]',]
            
            results0 = []
        
        for xpath in xpaths:
            try:
                element = WebDriverWait(navegador, 10).until(
               EC.presence_of_element_located((By.XPATH, xpath)))
                results0.append(element.text)
            except TimeoutException:
                 print(f'element not found: {xpath}')
        #printe dos resultado para acompanhar
        print(results0)
        # criação da planilha com os dados e concactenação com a planilha base
        jogos = (results0 )
        visu = pd.read_excel(file_path)
        jogos_lidos = pd.DataFrame (jogos)
        lidos = pd.concat([jogos_lidos, visu])
        lidos.to_excel("Conjunto de jogos.xlsx",index=False)
        print(lidos)
        
        #verificação de seguincia da condição de pesquisado da data inicial a para a data final
        results[dia.strftime("%d/%m/%Y")] = resultados
        
        navegador.quit()
    
    dia += timedelta(days=1)
#le a planilha e altera o nome da primera coluna  para "numeros"
lidos = pd.read_excel("Conjunto de jogos.xlsx")
lidos.columns =["numeros"]
lidos.to_excel("Conjunto de jogos.xlsx",index=False)
print(lidos)

# define o nome do arquivo na mesma pasta do codigo
file_path = ("resultado_frequentes.xlsx")
#verifica se ja exite o caso não exista cria o arquivo
if not os.path.exists (file_path):
    inni_jogos =  {}
    conju_jogos = pd.DataFrame(inni_jogos)
    conju_jogos.to_excel(file_path, index=False)

# Importando dados
df = pd.read_excel("Conjunto de jogos.xlsx")

# Contando números repetidos
num_repetidos = df['numeros'].value_counts()

# Criando gráfico de barras
num_repetidos.plot(kind='bar')

# Criando o gráfico
plt.bar(num_repetidos.index, num_repetidos.values)
plt.xlabel("Número")
plt.ylabel("Quantidade de vezes repetido")
plt.title("Números mais repetidos na coluna 1")

# Salvando o gráfico como imagem
plt.savefig("grafico.png")

# Abrindo a planilha
wb = openpyxl.load_workbook("resultado_frequentes.xlsx")

# Selecionando a aba que deseja adicionar o gráfico
ws = wb["Sheet1"]

# Adicionando o gráfico como imagem
img = openpyxl.drawing.image.Image("grafico.png")
ws.column_dimensions[ws.cell(row=1, column=1).column_letter].width = 30
ws.row_dimensions[1].height = 30
ws.add_image(img, "A1")

# Salvando a planilha
wb.save("resultado_frequentes.xlsx")
