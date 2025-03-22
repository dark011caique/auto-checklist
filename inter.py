from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time
from PIL import ImageGrab, Image
import pyautogui
import io
import win32clipboard  # Biblioteca para copiar a imagem para a área de transferência
from datetime import datetime, timedelta
from time import sleep
import customtkinter as ctk
import os



os.environ['TCL_LIBRARY'] = r"C:\Users\Caique\AppData\Local\Programs\Python\Python313\tcl\tcl8.6"


# Configuração da janela principal
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("ACC Adiq")
app.geometry("800x500")

# Layout principal
frame_sidebar = ctk.CTkFrame(app, width=200, corner_radius=0)
frame_sidebar.pack(side="left", fill="y")

frame_main = ctk.CTkFrame(app)
frame_main.pack(side="right", expand=True, fill="both", padx=20, pady=20)

# Título no sidebar
label_title = ctk.CTkLabel(frame_sidebar, text="ACC Adiq", font=("Arial", 18, "bold"))
label_title.pack(pady=20)

# Função para mostrar apenas o frame selecionado
def mostrar_frame(frame):
    for widget in frame_main.winfo_children():
        widget.pack_forget()  # Esconde todos os widgets no frame_main
    frame.pack(fill="both", expand=True, padx=20, pady=20)  # Mostra apenas o frame selecionado

# Funções para serem chamadas
def checklist():
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

    data_url = data_hora.strftime("%Y-%m-%d")

    px = f"https://grafana-ocp4.adiq.io/d/f5067f59-9f90-4e0f-b86c-8c2ba32fc3a8/monitor-pix-gw-bancario-geral?orgId=1&from=1740997996000&to=1741040658000&var-dataselecionada={data_url}"
    
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

    fisico_path = "fisico.png"
    driver.save_screenshot(fisico_path)

    time.sleep(1)

    # 🔹 Ajustes para os outros URLs
    grafana_ecommerce_url = "https://grafana-monitoring-hml-grafana-monitoring-hml.apps.svs.adiq.local/d/fe97u788lyneob/visao-geral-transacional-e-commerce?from=now-1h&to=now&orgId=1&refresh=5s"
    driver.get(grafana_ecommerce_url)
    time.sleep(3)

    ecommerce_path = "ecommerce.png"
    driver.save_screenshot(ecommerce_path)

    sleep(1)

    grafana_arquivos_url = "https://grafana-ocp4.adiq.io/d/uAoTFkJMzr/monitoracao-arquivos-geral-ambiente-de-prd?orgId=1&from=now-1h&to=now&refresh=10s"
    driver.get(grafana_arquivos_url)
    time.sleep(3)

    arquivos_path = "arquivos.png"
    driver.save_screenshot(arquivos_path)

    time.sleep(1)

    unificada_url = "https://adqtrjvpkbn01.adiq.local:5601/app/dashboards#/view/f1b3c7d4-22a4-4956-be94-1a8c6103d4d4"
    driver.get(unificada_url)
    time.sleep(8)

    unificada_path = "unificada.png"
    driver.save_screenshot(unificada_path)

    time.sleep(1)

    banese_url = "https://adqtrjvpkbn01.adiq.local:5601/app/dashboards#/view/3cf45840-2021-11ee-a5b4-81e7ec0febaf"
    driver.get(banese_url)
    time.sleep(8)

    banese_path = "banese.png"
    driver.save_screenshot(banese_path)

    time.sleep(1)

    softpass_url = "https://adqtrjvpkbn01.adiq.local:5601/app/dashboards#/view/73b07c60-2395-11ef-a2fe-31640ea6f96c"
    driver.get(softpass_url)
    time.sleep(8)

    softpass_path = "softpass.png"
    driver.save_screenshot(softpass_path)

    time.sleep(1)

    pix_url = px
    driver.get(pix_url)
    time.sleep(3)

    zoom = driver.find_element(By.XPATH, "//span[contains(@class, 'css-1riaxdn') and contains(text(), 'Zoom to data')]")
    zoom.click()
    sleep(3)

    pix_path = "pix.png"
    driver.save_screenshot(pix_path)

    time.sleep(1)

    zabbix_url = "http://radar.adiq.local/zabbix/zabbix.php?action=dashboard.view"
    driver.get(zabbix_url)
    time.sleep(3)

    zabbix_path = "zabbix.png"
    driver.save_screenshot(zabbix_path)

    time.sleep(1)

    uptime_url = "https://dashboard.uptimerobot.com/monitors"
    driver.get(uptime_url)
    time.sleep(3)

    uptime_path = "uptime.png"
    driver.save_screenshot(uptime_path)

    time.sleep(1)
    ######################################################################################################### TEAMS

    # 🔹 Acessar o Microsoft Teams Web
    teams_url = "https://teams.microsoft.com/"
    driver.get(teams_url)
    time.sleep(10)  # Tempo para login manual
    grupo = driver.find_element(By.XPATH, '//span[@title="Checklist Adiq / Kalendae"]')
    grupo.click()


    teste = driver.find_element(By.XPATH, '//button[@name="expand-compose"]')
    teste.click()

    # 🔹 Enviar mensagem no Teams
    try:

        # Aguarda o campo de texto
        time.sleep(1)
        chat_box = driver.find_element(By.XPATH, '//div[@contenteditable="true"]')
        chat_box.click()
        time.sleep(1)
        

        # Digitar a mensagem
        texto = f"""
        {mensagem}

        {t_fisico}
        {t_ecommerce}
        {t_unificada}
        {t_banese}
        {t_spftpass}
        {t_pix}
        {t_atquivos}
        {t_zabbix}
        {t_uptime}

        Legenda: 
        ✅ Ambiente OK
        ❌ Incidente em Andamento
        """

        chat_box.send_keys(texto)
        pyautogui.press("enter")  # Pressiona enter para enviar tudo de uma vez


    #########################################################################################################

        # 🔹 Função para copiar a imagem para a área de transferência e colar no chat
        def copiar_e_enviar_imagem(caminho_imagem, chat_texto):
            chat_box.send_keys(chat_texto)
            pyautogui.press("enter")
            time.sleep(1)

            # Abrir a imagem
            imagem = Image.open(caminho_imagem)
            output = io.BytesIO()
            imagem.save(output, format="BMP")
            data = output.getvalue()[14:]  # Remove os primeiros bytes do cabeçalho BMP
            output.close()

            # Copiar a imagem para a área de transferência
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
            win32clipboard.CloseClipboard()

            print(f"✅ Screenshot '{caminho_imagem}' copiado para a área de transferência!")

            # Colar no chat e enviar
            pyautogui.hotkey("ctrl", "v")
            time.sleep(1)
            pyautogui.press("enter")
            pyautogui.press("enter")

        # 🔹 Chamadas para cada imagem
        copiar_e_enviar_imagem("fisico.png", t_fisico)
        copiar_e_enviar_imagem("ecommerce.png", t_ecommerce)
        copiar_e_enviar_imagem("unificada.png", t_unificada)
        copiar_e_enviar_imagem("banese.png", t_banese)
        copiar_e_enviar_imagem("softpass.png", t_spftpass)
        copiar_e_enviar_imagem("pix.png", t_pix)
        copiar_e_enviar_imagem("arquivos.png", t_atquivos)
        copiar_e_enviar_imagem("zabbix.png", t_zabbix)
        copiar_e_enviar_imagem("uptime.png", t_uptime)

        print("✅ Mensagem e imagem enviadas no Teams!")

    except Exception as e:
        print(f"❌ Erro ao enviar mensagem e imagem: {e}")


def tarefa2():
    # 🔹 Configurar o WebDriver
    chrome_user_data_dir = r"C:\Users\Win10\AppData\Local\Google\Chrome\User Data"

    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", "localhost:9222")  # Conectar ao Chrome aberto
    options.add_argument(f"user-data-dir={chrome_user_data_dir}")

    # Iniciar WebDriver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    # 🔹 Acessar o Microsoft Teams Web
    teams_url = "https://teams.microsoft.com/"
    driver.get(teams_url)
    time.sleep(10)  # Tempo para login manual



    teste = driver.find_element(By.XPATH, '//button[@name="expand-compose"]')
    teste.click()

    # 🔹 Enviar mensagem no Teams
    try:

        # Aguarda o campo de texto
        time.sleep(1)
        chat_box = driver.find_element(By.XPATH, '//div[@contenteditable="true"]')
        chat_box.click()
        time.sleep(1)
        
        def copiar_e_enviar_imagem(caminho_imagem, chat_texto):
            pyautogui.hotkey("ctrl", "b")
            chat_box.send_keys(chat_texto)
            pyautogui.hotkey("ctrl", "b")
            pyautogui.press("enter")
            time.sleep(1)

                # Abrir a imagem
            imagem = Image.open(caminho_imagem)
            output = io.BytesIO()
            imagem.save(output, format="BMP")
            data = output.getvalue()[14:]  # Remove os primeiros bytes do cabeçalho BMP
            output.close()

                # Copiar a imagem para a área de transferência
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
            win32clipboard.CloseClipboard()

            print(f"✅ Screenshot '{caminho_imagem}' copiado para a área de transferência!")

                # Colar no chat e enviar
            pyautogui.hotkey("ctrl", "v")
            time.sleep(1)
            pyautogui.press("enter")
            pyautogui.press("enter")
        
        copiar_e_enviar_imagem("pix.png", "ACC INFORMA:")
    except Exception as e:
        print(f"❌ Erro ao enviar mensagem e imagem: {e}")

def tarefa3():
    print("Tarefa 3 realizada!")

def tarefa4():
    print("Tarefa 3 realizada!")

def tarefa5():
    print("Tarefa 3 realizada!")

# Frame do checklist
frame_checklist = ctk.CTkFrame(frame_main, fg_color="gray20", corner_radius=10)

label_checklist = ctk.CTkLabel(frame_checklist, text="Checklist", font=("Arial", 18, "bold"))
label_checklist.pack(pady=10)

# Checkboxes com variáveis para verificar o estado
checkbox_var1 = ctk.BooleanVar()
checkbox_var2 = ctk.BooleanVar()
checkbox_var3 = ctk.BooleanVar()
checkbox_var4 = ctk.BooleanVar()
checkbox_var5 = ctk.BooleanVar()

checkbox1 = ctk.CTkCheckBox(frame_checklist, text="Checklist", variable=checkbox_var1)
checkbox1.pack(pady=5)
checkbox2 = ctk.CTkCheckBox(frame_checklist, text="Pix", variable=checkbox_var2)
checkbox2.pack(pady=5)
checkbox3 = ctk.CTkCheckBox(frame_checklist, text="Comparativo", variable=checkbox_var3)
checkbox3.pack(pady=5)
checkbox4 = ctk.CTkCheckBox(frame_checklist, text="adiq+", variable=checkbox_var4)
checkbox4.pack(pady=5)
checkbox5 = ctk.CTkCheckBox(frame_checklist, text="Farol", variable=checkbox_var5)
checkbox5.pack(pady=5)

# Função para executar as tarefas associadas aos itens marcados
def executar_tarefas():
    if checkbox_var1.get():
        checklist()
        sleep(10)
    if checkbox_var2.get():
        tarefa2()
        sleep(10)
    if checkbox_var3.get():
        tarefa3()
        sleep(10)
    if checkbox_var4.get():
        tarefa4()
        sleep(10)
    if checkbox_var5.get():
        tarefa5()
        

# Botão dentro do checklist para rodar as funções
btn_executar = ctk.CTkButton(frame_checklist, text="Executar Tarefas", command=executar_tarefas)
btn_executar.pack(pady=20)

# Frame inicial
frame_home = ctk.CTkFrame(frame_main)

label_home = ctk.CTkLabel(frame_home, text="Home Page", font=("Arial", 24, "bold"))
label_home.pack(pady=20)

# Exibir o frame inicial por padrão
mostrar_frame(frame_home)

# frame whatsapp
frame_whatsapp= ctk.CTkFrame(frame_main, fg_color="gray20", corner_radius=10)

frame_whatsapp = ctk.CTkLabel(frame_whatsapp, text="whatsapp", font=("Arial", 18, "bold"))
frame_whatsapp.pack(pady=10)

# Botões da sidebar
btn_home = ctk.CTkButton(frame_sidebar, text="Home", fg_color="gray30", hover_color="gray40", command=lambda: mostrar_frame(frame_home))
btn_home.pack(fill="x", pady=5, padx=10)

btn_checklist = ctk.CTkButton(frame_sidebar, text="CheckList", fg_color="gray30", hover_color="gray40", command=lambda: mostrar_frame(frame_checklist))
btn_checklist.pack(fill="x", pady=5, padx=10)

btn_whatsapp = ctk.CTkButton(frame_sidebar, text="WhatsApp", fg_color="gray30", hover_color="gray40", command=lambda: mostrar_frame(frame_whatsapp))
btn_whatsapp.pack(fill="x", pady=5, padx=10)

# Botão do sistema
btn_system = ctk.CTkOptionMenu(frame_sidebar, values=["System"])
btn_system.pack(side="bottom", pady=10)

app.mainloop()
