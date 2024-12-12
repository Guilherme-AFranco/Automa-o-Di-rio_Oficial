from selenium import webdriver
from selenium.webdriver.chrome.service import Service 
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC 
import time
import win32com.client as win32
from Gemini import *
from email.mime.text import MIMEText
from datetime import datetime

# service = Service(executable_path="chromedriver.exe") # Descomentar para uso no pc do Léo
service = Service(executable_path="E:\Backup_PC\Aplicativos\ChromeDrive\WebDriver\chromedriver.exe") # Descomentar para uso no pc do Gui
driver = webdriver.Chrome(service=service)
driver.maximize_window()

driver.get("https://in.gov.br/leiturajornal")

time.sleep(2)

WebDriverWait(driver, 2).until(
    EC.presence_of_all_elements_located((By.ID, "toggle-search-advanced"))
)

advance_search = driver.find_element(By.ID, "toggle-search-advanced")
advance_search.click()
 

time.sleep (1)
WebDriverWait(driver, 12).until(
    EC.presence_of_all_elements_located((By.ID, "do3"))
)

secao3 = driver.find_element(By.ID, "do3")
secao3.click()

time.sleep(2)

dia = driver.find_element(By.ID, "dia") # Voltar para "dia" dps (tinha poucas noticias com parametro "dia")
dia.click()

WebDriverWait(driver, 2).until(
    EC.presence_of_all_elements_located((By.CLASS_NAME, "form-control"))
)


search_area = driver.find_element(By.CLASS_NAME, "form-control")
search_area.clear()
search_area.send_keys("aeronave" + Keys.ENTER)
time.sleep(10)


## Obtendo as noticias ##
noticias = driver.find_elements(By.CLASS_NAME, "resultados-wrapper") # Obtem as classes de noticias existentes na aba
noticia_url = {}
titulo_dou = {} # Dicionario para colocar o titulo da notícia
texto_dou = {} # Dicionario para o Gemini trabalhar
for i, noticia in enumerate(noticias):
    try:
        link_element = noticia.find_element(By.TAG_NAME, "a") # Encontrar link
        noticia_url[f'Noticia {i}'] = link_element.get_attribute("href") # Coletar o link

        driver.execute_script("window.open(arguments[0]);", noticia_url[f'Noticia {i}']) # Abrir em nova aba
        driver.switch_to.window(driver.window_handles[1]) # Ir para a nova aba

        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "dou-paragraph"))
        ) # Esperar o texto da notícia carregar

        titulo_dou[f'Noticia {i}'] = driver.find_element(By.CLASS_NAME, "identifica").text # Coletar titulo
        texto_dou[f'Noticia {i}'] = driver.find_element(By.CLASS_NAME, "texto-dou").text # Coletar notícia

        print(f'\nTítulo da Notícia {i+1}:')
        print(titulo_dou[f'Noticia {i}'])
        print(f'Texto da Notícia {i+1}:')
        print(texto_dou[f'Noticia {i}'])

        driver.close() # Fechar a aba adicional
        driver.switch_to.window(driver.window_handles[0])  # Voltar à aba principal

        noticias = driver.find_elements(By.CLASS_NAME, "resultados-wrapper") # Atualizar a lista das notícias na pagina
    except Exception as e:
        print(f"Erro ao processar a notícia {i+1}: {e}") # Se der erro, ele avisa e ficamos tristes

driver.quit() # Sai da pagina da web
    

results = gemini_analysis(titulo_dou,texto_dou) # Faz a analise a partir de IA na função vista pelo código Gemini.py

for key, value in results.items(): # Loop para pegar todas as noticias do dicionario
    print(f"Título: {value['title']}") # Imprime o título
    if "response" in value: # Pega apenas a resposta gerada pela IA para printar
        print(f"Resposta: {value['response']}") # Imprime a resposta
        with open("noticias.txt", "w") as arquivo:
            arquivo.write(f"Título {value['title']}.\n{value['response']}\n\n\n")
    
    
    else:
        print(f"Erro: {value['error']}") # Se der erra, ficaremos tristes

titulo = []
body = []
open("noticias.txt", "w").close()


for i in range(0, int(len(titulo_dou))):
    titulo.append(f"Titulo: {titulo_dou[f'Noticia {i}']}.")
    body.append(f"{texto_dou[f'Noticia {i}']}.")
    with open("noticias.txt","a") as arquivo:
        arquivo.write(f"<p>{str(titulo[i])}<\p>\n")
        arquivo.write(f"<p>{str(body[i])}<\p>\n\n")
        arquivo.write("<p><\p>")
        


# outlook = win32.Dispatch('Outlook.Application') # cria integração com o outlook
# email = outlook.CreateItem(0) # Cria e-mail

# # Configurações do e-mail


# with open("noticias.txt","r") as file:
#    file_content = file.read()

# file_content.replace("\n", "<br>")

# email.To = "guilherme.franco@embraer.com.br;"
# email.Subject = "Resumo Diário Oficial"


# email.HTMLBody = f"""
#     <p> Bom dia! Segue o resumo do dirário oficial </p>
#     <p></p>
#     <p> {(file_content)} </p>
#     <p></p>
#     <p> Best regards| Atenciosamente, </p>
#     <p> Guilherme Franco & Leo Ferreira </p>   
# """   
# email.Send() 

# driver.quit()