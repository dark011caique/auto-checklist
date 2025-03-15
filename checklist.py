from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
from PIL import ImageGrab
import pyautogui
import io
import win32clipboard  # Biblioteca para copiar a imagem para a área de transferência
from datetime import datetime, timedelta

# Obtendo a data e o horário atuais
data_hora = datetime.now()

# Arredondando o horário
if data_hora.minute > 30:
    data_hora = data_hora.replace(minute=0, second=0, microsecond=0) + timedelta(hours=1)
else:
    data_hora = data_hora.replace(minute=0, second=0, microsecond=0)

# Formatando a data e hora no formato desejado
data_formatada = data_hora.strftime("%d/%m/%Y - %H:%M")

# Criando a string com a data e hora arredondada
mensagem = f'ACC ADIQ: ⏰⏰ Checklist Ambientes {data_formatada}'

print(mensagem)



############################################################################################################## VARIAVEIS
t_fisico = "✅ [Transacional Físico](https://grafana-monitoring-hml-grafana-monitoring-hml.apps.svs.adiq.local/d/ee9atfdxwhg5cf/visao-geral-transacional-fisico-sniffer-dxc?orgId=1&from=now-30m&to=now&refresh=5s)"

t_ecommerce = "✅ [Transacional E-commerce](https://grafana-monitoring-hml-grafana-monitoring-hml.apps.svs.adiq.local/d/fe97u788lyneob/visao-geral-transacional-e-commerce?from=now-1h&to=now&orgId=1&refresh=5s)"

t_unificada = "✅ [Tela Unificada](https://adqtrjvpkbn01.adiq.local:5601/app/dashboards#/view/f1b3c7d4-22a4-4956-be94-1a8c6103d4d4)"

t_banese = "✅ [Transacional Hub On-Us Banese](https://adqtrjvpkbn01.adiq.local:5601/app/dashboards#/view/3cf45840-2021-11ee-a5b4-81e7ec0febaf)"

t_spftpass = "✅ [Transacional Softpass](https://adqtrjvpkbn01.adiq.local:5601/app/dashboards#/view/73b07c60-2395-11ef-a2fe-31640ea6f96c)"

t_pix = "✅ [Transacional PIX](https://grafana-ocp4.adiq.io/d/f5067f59-9f90-4e0f-b86c-8c2ba32fc3a8/monitor-pix-gw-bancario-geral?orgId=1&from=1740997996000&to=1741040658000&var-dataselecionada=2025-03-07)"

t_atquivos = "✅ [Grafana - Arquivos](https://grafana-ocp4.adiq.io/d/uAoTFkJMzr/monitoracao-arquivos-geral-ambiente-de-prd?orgId=1&var-DiaNaoUtil=1&var-JanelaMarcacao=Janela%201%20Ida&var-JanelaMarcacao=Janela%202%20Ida&var-JanelaMarcacao=Janela%203%20Ida&var-JanelaMarcacao=Janela%201%20Volta&var-JanelaMarcacao=Janela%202%20Volta&var-JanelaMarcacao=Janela%203%20Volta&var-DataMarcacao=2025-02-14&from=now-1h&to=now&refresh=10s)"

t_zabbix = "✅ [Zabbix](http://radar.adiq.local/zabbix/zabbix.php?action=dashboard.view)"

t_uptime = "✅ [Uptime](https://dashboard.uptimerobot.com/monitors)"

######################################################################################################### CONFIGURAÇÃO CHROME


# 🔹 Configurar o WebDriver
chrome_user_data_dir = r"C:\Users\Win10\AppData\Local\Google\Chrome\User Data"

options = webdriver.ChromeOptions()
options.add_experimental_option("debuggerAddress", "localhost:9222")  # Conectar ao Chrome aberto
options.add_argument(f"user-data-dir={chrome_user_data_dir}")

# Iniciar WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

######################################################################################################### GRAFANA FISICO


# 🔹 Acessar o site e tirar o print
grafana_fisico_url = "https://grafana-monitoring-hml-grafana-monitoring-hml.apps.svs.adiq.local/d/ee9atfdxwhg5cf/visao-geral-transacional-fisico-sniffer-dxc?orgId=1&from=now-30m&to=now&refresh=5s"
driver.get(grafana_fisico_url)
time.sleep(3)  # Tempo para carregar a página

# 🔹 Capturar apenas a janela do navegador (evita pegar a tela inteira)
fisico = ImageGrab.grab()
fisico_path = "fisico.png"
fisico.save(fisico_path)

######################################################################################################### GRAFANA ECOMMERCE


grafana_ecommerce_url = "https://grafana-monitoring-hml-grafana-monitoring-hml.apps.svs.adiq.local/d/fe97u788lyneob/visao-geral-transacional-e-commerce?from=now-1h&to=now&orgId=1&refresh=5s"
driver.get(grafana_ecommerce_url)
time.sleep(3)  # Tempo para carregar a página

# 🔹 Capturar apenas a janela do navegador (evita pegar a tela inteira)
ecommerce = ImageGrab.grab()
ecommerce_path = "ecommerce.png"
ecommerce.save(ecommerce_path)

time.sleep(2)
######################################################################################################### GRAFANA ARQUIVOS

grafana_arquivos_url = "https://grafana-ocp4.adiq.io/d/uAoTFkJMzr/monitoracao-arquivos-geral-ambiente-de-prd?orgId=1&var-DiaNaoUtil=1&var-JanelaMarcacao=Janela%201%20Ida&var-JanelaMarcacao=Janela%202%20Ida&var-JanelaMarcacao=Janela%203%20Ida&var-JanelaMarcacao=Janela%201%20Volta&var-JanelaMarcacao=Janela%202%20Volta&var-JanelaMarcacao=Janela%203%20Volta&var-DataMarcacao=2025-02-14&from=now-1h&to=now&refresh=10s"
driver.get(grafana_arquivos_url)
time.sleep(3)  # Tempo para carregar a página

# 🔹 Capturar apenas a janela do navegador (evita pegar a tela inteira)
arquivos = ImageGrab.grab()
arquivos_path = "arquivos.png"
arquivos.save(arquivos_path)

time.sleep(2)
######################################################################################################### TELA UNIFICADA

unificada_url = "https://adqtrjvpkbn01.adiq.local:5601/app/dashboards#/view/f1b3c7d4-22a4-4956-be94-1a8c6103d4d4"
driver.get(unificada_url)
time.sleep(10)  # Tempo para carregar a página

# 🔹 Capturar apenas a janela do navegador (evita pegar a tela inteira)
unificada = ImageGrab.grab()
unificada_path = "unificada.png"
unificada.save(unificada_path)

time.sleep(2)
#########################################################################################################  Transacional Hub On-Us Banese 

banese_url = "https://adqtrjvpkbn01.adiq.local:5601/app/dashboards#/view/3cf45840-2021-11ee-a5b4-81e7ec0febaf"
driver.get(banese_url)
time.sleep(10)  # Tempo para carregar a página

# 🔹 Capturar apenas a janela do navegador (evita pegar a tela inteira)
banese = ImageGrab.grab()
banese_path = "banese.png"
banese.save(banese_path)

time.sleep(2)
######################################################################################################### Transacional Softpass 

softpass_url = "https://adqtrjvpkbn01.adiq.local:5601/app/dashboards#/view/73b07c60-2395-11ef-a2fe-31640ea6f96c"
driver.get(softpass_url)
time.sleep(10)  # Tempo para carregar a página

# 🔹 Capturar apenas a janela do navegador (evita pegar a tela inteira)
softpass = ImageGrab.grab()
softpass_path = "softpass.png"
softpass.save(softpass_path)

time.sleep(2)

######################################################################################################### Transacional PIX 

pix_url = "https://grafana-ocp4.adiq.io/d/f5067f59-9f90-4e0f-b86c-8c2ba32fc3a8/monitor-pix-gw-bancario-geral?orgId=1&from=1740997996000&to=1741040658000&var-dataselecionada=2025-03-07"
driver.get(pix_url)
time.sleep(3)  # Tempo para carregar a página

# 🔹 Capturar apenas a janela do navegador (evita pegar a tela inteira)
pix = ImageGrab.grab()
pix_path = "pix.png"
pix.save(pix_path)

######################################################################################################### ZABBIX

zabbix_url = "http://radar.adiq.local/zabbix/zabbix.php?action=dashboard.view"
driver.get(zabbix_url)
time.sleep(3)  # Tempo para carregar a página

# 🔹 Capturar apenas a janela do navegador (evita pegar a tela inteira)
zabbix = ImageGrab.grab()
zabbix_path = "zabbix.png"
zabbix.save(zabbix_path)

time.sleep(2)

######################################################################################################### UPTIME

uptime_url = "https://dashboard.uptimerobot.com/monitors"
driver.get(uptime_url)
time.sleep(3)  # Tempo para carregar a página

# 🔹 Capturar apenas a janela do navegador (evita pegar a tela inteira)
uptime = ImageGrab.grab()
uptime_path = "uptime.png"
uptime.save(uptime_path)

time.sleep(2)


######################################################################################################### TEAMS

# 🔹 Acessar o Microsoft Teams Web
teams_url = "https://teams.microsoft.com/"
driver.get(teams_url)
time.sleep(20)  # Tempo para login manual
grupo = driver.find_element(By.XPATH, '//span[@title="Checklist Adiq / Kalendae"]')
grupo.click()


teste = driver.find_element(By.XPATH, '//button[@name="expand-compose"]')
teste.click()

# 🔹 Enviar mensagem no Teams
try:

    # Aguarda o campo de texto
    time.sleep(2)
    chat_box = driver.find_element(By.XPATH, '//div[@contenteditable="true"]')
    chat_box.click()
    time.sleep(2)
    

    # Digitar a mensagem
    chat_box.send_keys(mensagem )
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)
    chat_box.send_keys(t_fisico)

    pyautogui.press("enter")
    time.sleep(2)
    chat_box.send_keys(t_ecommerce)
    pyautogui.press("enter")
    time.sleep(2)
    chat_box.send_keys(t_unificada)
    pyautogui.press("enter")
    time.sleep(2)
    chat_box.send_keys(t_banese)
    pyautogui.press("enter")
    time.sleep(2)
    chat_box.send_keys(t_spftpass)
    pyautogui.press("enter")
    time.sleep(2)
    chat_box.send_keys(t_pix)
    pyautogui.press("enter")
    time.sleep(2)
    chat_box.send_keys(t_atquivos)
    pyautogui.press("enter")
    time.sleep(2)
    chat_box.send_keys(t_zabbix)
    pyautogui.press("enter")
    time.sleep(2)
    chat_box.send_keys(t_uptime)
    pyautogui.press("enter")
    time.sleep(2)
    pyautogui.press("enter")
    time.sleep(2)
    chat_box.send_keys("Legenda: ")
    pyautogui.press("enter")
    time.sleep(2)
    chat_box.send_keys("✅ Ambiente OK ")
    pyautogui.press("enter")
    time.sleep(2)
    chat_box.send_keys("❌ Incidente em Andamento")
    pyautogui.press("enter")
    time.sleep(2)
    pyautogui.press("enter")
    time.sleep(2)

#########################################################################################################

    chat_box.send_keys(t_fisico)
    pyautogui.press("enter")
    time.sleep(2)
        # 🔹 Copiar a imagem para a área de transferência
    output = io.BytesIO()
    fisico.save(output, format="BMP")
    data = output.getvalue()[14:]  # Remove os primeiros bytes do cabeçalho BMP
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

    print("✅ Screenshot copiado para a área de transferência!")

    # 🔹 Simular "Ctrl + V" para colar a imagem
    pyautogui.hotkey("ctrl", "v")
    time.sleep(2)

    # 🔹 Pressionar Enter para enviar
    pyautogui.press("enter")
    pyautogui.press("enter")

#########################################################################################################

    chat_box.send_keys(t_ecommerce)
    pyautogui.press("enter")
    time.sleep(2)

    # 🔹 Copiar a imagem para a área de transferência
    output = io.BytesIO()
    ecommerce.save(output, format="BMP")
    data = output.getvalue()[14:]  # Remove os primeiros bytes do cabeçalho BMP
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

    print("✅ Screenshot copiado para a área de transferência!")

    pyautogui.hotkey("ctrl", "v")
    time.sleep(2)

        # 🔹 Pressionar Enter para enviar
    pyautogui.press("enter")
    pyautogui.press("enter")

#########################################################################################################

    chat_box.send_keys(t_unificada)
    pyautogui.press("enter")
    time.sleep(2)

    # 🔹 Copiar a imagem para a área de transferência
    output = io.BytesIO()
    unificada.save(output, format="BMP")
    data = output.getvalue()[14:]  # Remove os primeiros bytes do cabeçalho BMP
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

    print("✅ Screenshot copiado para a área de transferência!")

    pyautogui.hotkey("ctrl", "v")
    time.sleep(2)

        # 🔹 Pressionar Enter para enviar
    pyautogui.press("enter")
    pyautogui.press("enter")

#########################################################################################################
    chat_box.send_keys(t_banese)
    pyautogui.press("enter")
    time.sleep(2)

    # 🔹 Copiar a imagem para a área de transferência
    output = io.BytesIO()
    banese.save(output, format="BMP")
    data = output.getvalue()[14:]  # Remove os primeiros bytes do cabeçalho BMP
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

    print("✅ Screenshot copiado para a área de transferência!")

    pyautogui.hotkey("ctrl", "v")
    time.sleep(2)

    pyautogui.press("enter")
    pyautogui.press("enter")
#########################################################################################################

    chat_box.send_keys(t_spftpass)
    pyautogui.press("enter")
    time.sleep(2)

    # 🔹 Copiar a imagem para a área de transferência
    output = io.BytesIO()
    softpass.save(output, format="BMP")
    data = output.getvalue()[14:]  # Remove os primeiros bytes do cabeçalho BMP
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

    print("✅ Screenshot copiado para a área de transferência!")

    pyautogui.hotkey("ctrl", "v")
    time.sleep(2)

    pyautogui.press("enter")
    pyautogui.press("enter")

#########################################################################################################

    chat_box.send_keys(t_pix)
    pyautogui.press("enter")
    time.sleep(2)

    # 🔹 Copiar a imagem para a área de transferência
    output = io.BytesIO()
    pix.save(output, format="BMP")
    data = output.getvalue()[14:]  # Remove os primeiros bytes do cabeçalho BMP
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

    print("✅ Screenshot copiado para a área de transferência!")

    pyautogui.hotkey("ctrl", "v")
    time.sleep(2)

    pyautogui.press("enter")
    pyautogui.press("enter")
#########################################################################################################

    chat_box.send_keys(t_atquivos)
    pyautogui.press("enter")
    time.sleep(2)

    # 🔹 Copiar a imagem para a área de transferência
    output = io.BytesIO()
    arquivos.save(output, format="BMP")
    data = output.getvalue()[14:]  # Remove os primeiros bytes do cabeçalho BMP
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

    print("✅ Screenshot copiado para a área de transferência!")

    pyautogui.hotkey("ctrl", "v")
    time.sleep(2)
    
    pyautogui.press("enter")
    pyautogui.press("enter")
#########################################################################################################

#########################################################################################################

    chat_box.send_keys(t_zabbix)
    pyautogui.press("enter")
    time.sleep(2)

    # 🔹 Copiar a imagem para a área de transferência
    output = io.BytesIO()
    zabbix.save(output, format="BMP")
    data = output.getvalue()[14:]  # Remove os primeiros bytes do cabeçalho BMP
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

    print("✅ Screenshot copiado para a área de transferência!")

    pyautogui.hotkey("ctrl", "v")
    time.sleep(2)

    pyautogui.press("enter")
    pyautogui.press("enter")
#########################################################################################################

    chat_box.send_keys(t_uptime)
    pyautogui.press("enter")
    time.sleep(2)

    # 🔹 Copiar a imagem para a área de transferência
    output = io.BytesIO()
    uptime.save(output, format="BMP")
    data = output.getvalue()[14:]  # Remove os primeiros bytes do cabeçalho BMP
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

    print("✅ Screenshot copiado para a área de transferência!")

    pyautogui.hotkey("ctrl", "v")
    time.sleep(2)


    print("✅ Mensagem e imagem enviadas no Teams!")

except Exception as e:
    print(f"❌ Erro ao enviar mensagem e imagem: {e}")

time.sleep(5)
driver.quit()  # Fecha o navegador
