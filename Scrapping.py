from selenium import webdriver
from selenium.webdriver.chrome.service import Service 
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC 
import time

def scrapping(secao,search):
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
        EC.presence_of_all_elements_located((By.ID, secao))
    )

    secao_n = driver.find_element(By.ID, secao)
    secao_n.click()

    time.sleep(2)

    dia = driver.find_element(By.ID, "dia") # Voltar para "dia" dps (tinha poucas noticias com parametro "dia")
    dia.click()

    WebDriverWait(driver, 2).until(
        EC.presence_of_all_elements_located((By.CLASS_NAME, "form-control"))
    )


    search_area = driver.find_element(By.CLASS_NAME, "form-control")
    search_area.clear()
    search_area.send_keys(search + Keys.ENTER)
    time.sleep(10)


    ## Obtendo as noticias ##
    noticias = driver.find_elements(By.CLASS_NAME, "resultados-wrapper") # Obtem as classes de noticias existentes na aba
    noticia_url = {}
    titulo_dou = {} # Dicionario para colocar o titulo da notícia
    texto_dou = {} # Dicionario para texto da noticia
    secao_dou ={} # Dicionario para numero da seção
    data_dou = {}# Dicionario para data de publicação

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
            data_dou[f'Noticia {i}'] = driver.find_element(By.CLASS_NAME, "publicado-dou-data").text # Coletar data
            secao_dou[f'Noticia {i}'] = driver.find_element(By.CLASS_NAME, "secao-dou").text # Coletar secao

            driver.close() # Fechar a aba adicional
            driver.switch_to.window(driver.window_handles[0])  # Voltar à aba principal

            noticias = driver.find_elements(By.CLASS_NAME, "resultados-wrapper") # Atualizar a lista das notícias na pagina
        except Exception as e:
            print(f"Erro ao processar a notícia {i+1}: {e}") # Se der erro, ele avisa e ficamos tristes

    driver.quit() # Sai da pagina da web

    return noticia_url, titulo_dou, texto_dou, secao_dou, data_dou