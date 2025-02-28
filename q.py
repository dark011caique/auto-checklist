import time
import subprocess
import pyautogui
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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
screenshot_path = r"C:\Users\Win10\OneDrive\Imagens\Saved Pictures\screenshot.png"  # Escolha um caminho válido
driver.save_screenshot(screenshot_path)
print(f"✅ Screenshot salvo em {screenshot_path}!")

# 🔹 Copiar a imagem para a área de transferência (usando "clip.exe")
subprocess.run(f'clip < "{screenshot_path}"', shell=True)
print("✅ Imagem copiada para a área de transferência!")

# 🔹 Acessar o WhatsApp Web
whatsapp_url = "https://web.whatsapp.com/"
driver.get(whatsapp_url)

# 🔹 Aguardar a página carregar (aguardar até o campo de busca aparecer)
try:
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="side"]/div[1]/div/div[2]/div/div/div[1]/p')))
    print("✅ WhatsApp carregado com sucesso!")
except:
    print("❌ Erro: WhatsApp Web não carregou.")
    driver.quit()
    exit()

# 🔹 Encontrar e abrir a conversa desejada
contato = "Fabricante De Lágrimas"  # Substitua pelo nome do contato/grupo exato no WhatsApp
try:
    search_box = driver.find_element(By.XPATH, '//*[@id="side"]/div[1]/div/div[2]/div/div/div[1]/p')
    search_box.click()
    time.sleep(1)
    search_box.send_keys(contato)
    time.sleep(2)
    search_box.send_keys(Keys.ENTER)
    time.sleep(2)
    print(f"✅ Chat com '{contato}' aberto!")
except Exception as e:
    print(f"❌ Erro ao abrir a conversa: {e}")
    driver.quit()
    exit()

# 🔹 Gerar a mensagem inicial
data_hora = "26/02/2025 - 22h00"  # Ajuste automático pode ser feito com datetime
mensagem = f"""ACC ADIQ: Checklist Ambientes {data_hora}

Legenda: 
[OK] Ambiente OK 
[!] Incidente em Andamento
"""

# 🔹 Encontrar o campo de mensagem e enviar texto + imagem
try:
    chat_box = driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/div[2]/div[1]/p')
    chat_box.click()
    time.sleep(1)

    # 🔹 Enviar a mensagem de texto
    chat_box.send_keys(mensagem)
    time.sleep(1)

    # 🔹 Simular "Ctrl + V" para colar a imagem
    pyautogui.hotkey("ctrl", "v")
    time.sleep(2)

    # 🔹 Pressionar Enter para enviar tudo junto
    pyautogui.press("enter")

    print("✅ Mensagem e imagem enviadas no WhatsApp!")
except Exception as e:
    print(f"❌ Erro ao enviar mensagem e imagem: {e}")

time.sleep(5)
driver.quit()
