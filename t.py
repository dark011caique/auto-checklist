import time
import subprocess
import pyautogui
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# 🔹 Configurar WebDriver
chrome_user_data_dir = r"C:\Users\Win10\AppData\Local\Google\Chrome\User Data"  # Substitua pelo caminho correto

options = webdriver.ChromeOptions()
options.add_experimental_option("debuggerAddress", "localhost:9222")  # Conecta à porta de depuração
options.add_argument(f"user-data-dir={chrome_user_data_dir}")  # Usar o diretório de dados do usuário

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# 🔹 Acessar o site e tirar o print
grafana_url = "https://grafana.com/"
driver.get(grafana_url)
time.sleep(3)

# 🔹 Tirar print com Selenium e salvar
screenshot_path = r"C:\Users\Win10\Pictures\screenshot.png"  # Escolha um caminho válido
driver.save_screenshot(screenshot_path)
print(f"✅ Screenshot salvo em {screenshot_path}!")

# 🔹 Copiar a imagem para a área de transferência (usando "clip.exe")
subprocess.run(f'clip < "{screenshot_path}"', shell=True)
print("✅ Imagem copiada para a área de transferência!")

# 🔹 Acessar o Microsoft Teams Web
teams_url = "https://teams.microsoft.com/"
driver.get(teams_url)
time.sleep(10)  # Tempo para login manual

# 🔹 Gerar a mensagem inicial (SEM emojis)
data_hora = "26/02/2025 - 22h00"  # Ajuste automático pode ser feito com datetime
mensagem = f"""ACC ADIQ: Checklist Ambientes {data_hora}

Legenda: 
[OK] Ambiente OK 
[!] Incidente em Andamento
"""

# 🔹 Encontrar o campo de mensagem no Teams e enviar o texto + imagem
try:
    chat_box = driver.find_element(By.XPATH, '//div[@contenteditable="true"]')
    chat_box.click()
    time.sleep(2)

    # 🔹 Enviar a mensagem de texto
    chat_box.send_keys(mensagem)
    time.sleep(1)

    # 🔹 Simular "Ctrl + V" para colar a imagem
    pyautogui.hotkey("ctrl", "v")
    time.sleep(2)

    print("✅ Mensagem e imagem enviadas no Teams!")
except Exception as e:
    print(f"❌ Erro ao enviar mensagem e imagem: {e}")

time.sleep(5)
driver.quit()
