from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
import os

chrome_user_data_dir = r"C:\Users\Win10\AppData\Local\Google\Chrome\User Data"# Substitua pelo caminho correto

# Configuração do Selenium para se conectar ao navegador já aberto
options = webdriver.ChromeOptions()
options.add_experimental_option("debuggerAddress", "localhost:9222")  # Conecta à porta de depuração
options.add_argument(f"user-data-dir={chrome_user_data_dir}")  # Usar o diretório de dados do usuário

# Usando o WebDriver Manager para baixar o ChromeDriver automaticamente
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# 🔹 Abrir o Microsoft Teams Web
teams_url = "https://teams.microsoft.com/"
driver.get(teams_url)
sleep(15)  # Tempo para login manual (caso precise)

# 🔹 Abrir o chat do grupo específico
grupo_nome = "Checklist Adiq / Kalendae"  # Nome do grupo no Teams
search_box = driver.find_element(By.XPATH, '//input[@type="text"]')  # Campo de busca
search_box.send_keys(grupo_nome)
sleep(2)
search_box.send_keys(Keys.ENTER)
sleep(5)  # Aguarda carregar o chat

# 🔹 Gerar a mensagem do checklist
data_hora = "26/02/2025 - 22h00"  # Ajuste automático pode ser feito com datetime
mensagem = f"""ACC ADIQ: ⏰⏰ Checklist Ambientes {data_hora}

Legenda: 
✅ Ambiente OK 
🆘 Incidente em Andamento
"""

# 🔹 Enviar a mensagem inicial no chat
chat_box = driver.find_element(By.XPATH, '//div[@contenteditable="true"]')
chat_box.send_keys(mensagem)
chat_box.send_keys(Keys.ENTER)
sleep(2)

# 🔹 Lista de sites para tirar print
sites = {
    "Grafana - Arquivos": "https://exemplo6.com",
    
}

# 🔹 Criar pasta para salvar prints
screenshot_dir = os.path.join(os.getcwd(), "screenshots")
os.makedirs(screenshot_dir, exist_ok=True)

# 🔹 Tirar prints e salvar
screenshots = {}
for nome, url in sites.items():
    driver.get(url)
    sleep(5)  # Espera carregar
    screenshot_path = os.path.join(screenshot_dir, f"{nome.replace(' ', '_')}.png")
    driver.save_screenshot(screenshot_path)
    screenshots[nome] = screenshot_path

# 🔹 Voltar ao Teams Web para enviar os prints
driver.get(teams_url)
sleep(5)

# 🔹 Enviar cada print com o nome correto
for nome, path in screenshots.items():
    chat_box.send_keys(f"✅ {nome}")  # Nome do sistema
    chat_box.send_keys(Keys.ENTER)
    sleep(2)

    # 🔹 Clicar no botão de anexar
    anexar_botao = driver.find_element(By.XPATH, '//button[@title="Anexar"]')  # Botão de anexar arquivos
    anexar_botao.click()
    sleep(2)

    # 🔹 Localizar o campo de upload e enviar o caminho do arquivo
    input_upload = driver.find_element(By.XPATH, '//input[@type="file"]')
    input_upload.send_keys(path)  # Selenium agora faz o upload da imagem
    sleep(5)  # Aguarda o upload da imagem

    # 🔹 Pressionar ENTER para enviar a imagem
    chat_box.send_keys(Keys.ENTER)
    sleep(3)

# 🔹 Mensagem final de confirmação
chat_box.send_keys("✅ Checklist finalizado e prints enviados!")
chat_box.send_keys(Keys.ENTER)

print("🚀 Automação finalizada!")
sleep(5)
driver.quit()  # Fecha o navegador