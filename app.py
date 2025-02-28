from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import requests

# Configuração do WebDriver (precisa do chromedriver instalado)
service = Service("caminho/para/chromedriver")  # Substitua pelo caminho correto
options = webdriver.ChromeOptions()

# Lista de sites para capturar
sites = [
    "https://www.google.com.br/maps/preview",
    
]

# Inicializa o WebDriver
driver = webdriver.Chrome(service=service, options=options)

# Tira screenshots e salva
screenshots = []
for i, site in enumerate(sites):
    driver.get(site)
    time.sleep(3)  # Espera carregar
    screenshot_path = f"screenshot_{i+1}.png"
    driver.save_screenshot(screenshot_path)
    screenshots.append(screenshot_path)

driver.quit()

# Webhook do Teams
TEAMS_WEBHOOK_URL = "https://teams.live.com/v2/?utm_source=OfficeWeb"

# Envia as imagens para o Teams
for screenshot in screenshots:
    with open(screenshot, "rb") as img_file:
        image_data = img_file.read()

    files = {
        "file": (screenshot, image_data, "image/png")
    }

    payload = {
        "text": f"Screenshot do site {sites[screenshots.index(screenshot)]}",
    }

    response = requests.post(TEAMS_WEBHOOK_URL, json=payload, files=files)

    if response.status_code == 200:
        print(f"Imagem {screenshot} enviada com sucesso!")
    else:
        print(f"Erro ao enviar {screenshot}: {response.text}")
