import pandas as pd
import os
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time

# Caminho da planilha de aniversários
caminho_planilha = r"Z:\MARKETING\ENDOMARKETING\Aniversários\LISTA DE ANIVERSÁRIOS ATUALIZADA.xlsx"

# Caminho base para os arquivos de aniversário (imagens .png)
caminho_base_aniversarios = r"Z:\MARKETING\ENDOMARKETING\Aniversários\2025"

# Caminho para o perfil do Chrome
caminho_perfil_chrome = os.path.join(os.path.dirname(__file__), 'chrome_profile')

# Ler a planilha
try:
    df = pd.read_excel(caminho_planilha)
    # Certifique-se de que a coluna de data está no formato datetime
    df['Data de Nascimento'] = pd.to_datetime(df['Data de Nascimento'], format='%d/%m/%Y')
except Exception as e:
    print(f"Erro ao ler a planilha: {e}")
    exit()

# Obter a data atual
hoje = datetime.now()
dia_atual = hoje.day
mes_atual = hoje.month

# Função para enviar mensagem no WhatsApp com imagem usando selenium
def enviar_mensagem_whatsapp(nome, caminho_imagem):
    mensagem = f"Muitas felicidades, {nome}!"
    
    # Configuração do WebDriver com webdriver-manager
    servico = Service(ChromeDriverManager().install())

    # Configurações do Chrome para usar o perfil
    chrome_options = Options()
    chrome_options.add_argument(f"user-data-dir={caminho_perfil_chrome}")  # Define o diretório do perfil
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-software-rasterizer")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--no-sandbox")

    driver = webdriver.Chrome(service=servico, options=chrome_options)  # Inicializa o WebDriver com o perfil
    driver.get("https://web.whatsapp.com")
    
    try:
        # Aguarde o usuário escanear o QR Code (apenas na primeira execução)
        if not os.path.exists(os.path.join(caminho_perfil_chrome, "Default")):
            print("Por favor, escaneie o QR Code no WhatsApp Web.")
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.XPATH, '//div[@aria-label="Caixa de texto de mensagem"]'))
            )
            print("QR Code escaneado com sucesso!")
        
        # Navegue diretamente para o grupo
        driver.get("https://web.whatsapp.com/accept?code=JCw4UlichjgFJ3zRcCMMKS")
        
        # Aguarde o carregamento do grupo
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
        )
        print("Grupo carregado com sucesso!")
        
        # Enviar a imagem com a mensagem como legenda
        try:
            # Clique no botão "+" para abrir a seção de anexos
            botao_mais = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[1]/div/button/span'))
            )
            botao_mais.click()
            time.sleep(1)
            
            # Clique em "Fotos e vídeos"
            fotos_videos = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div/span[5]/div/ul/div/div/div[2]'))
            )
            fotos_videos.click()
            time.sleep(1)
            
            # Encontre o campo de upload de imagem
            campo_upload = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]'))
            )
            campo_upload.send_keys(caminho_imagem)
            time.sleep(2)
            
            # Adicione a mensagem como legenda da imagem
            legenda = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
            )
            legenda.send_keys(mensagem)
            time.sleep(1)
            
            # Clique no botão de enviar
            botao_enviar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]'))
            )
            botao_enviar.click()
            
            # Aguarde alguns segundos para garantir que a mensagem seja enviada
            time.sleep(5)  # Aumente o tempo de espera se necessário
            print(f"Mensagem enviada para {nome} com sucesso!")
        except Exception as e:
            print(f"Erro ao enviar imagem: {e}")
    except Exception as e:
        print(f"Erro ao enviar mensagem para {nome}: {e}")
    finally:
        # Fechar o navegador apenas após o envio da mensagem
        driver.quit()

# Função para encontrar a pasta do mês atual
def encontrar_pasta_mes(mes_atual):
    # Lista todas as pastas no diretório base
    pastas = os.listdir(caminho_base_aniversarios)
    for pasta in pastas:
        # Verifica se a pasta começa com o número do mês atual
        if pasta.startswith(f"{mes_atual:02d}."):
            return os.path.join(caminho_base_aniversarios, pasta)
    return None

# Encontrar a pasta do mês atual
pasta_mes_atual = encontrar_pasta_mes(mes_atual)

if pasta_mes_atual:
    print(f"Pasta do mês atual encontrada: {pasta_mes_atual}")
else:
    print(f"Pasta do mês atual não encontrada.")
    exit()

# Iterar sobre as linhas da planilha
for index, row in df.iterrows():
    nome = row['Nome']
    data_nascimento = row['Data de Nascimento']
    
    # Verificar se a data de nascimento é hoje
    if data_nascimento.day == dia_atual and data_nascimento.month == mes_atual:
        # Construir o caminho da imagem de aniversário
        arquivo_aniversario = f"Aniversário {nome}.png"  # Arquivo de imagem
        caminho_completo = os.path.join(pasta_mes_atual, arquivo_aniversario)
        
        # Verificar se o arquivo existe
        if os.path.exists(caminho_completo):
            print(f"Arquivo encontrado para {nome}: {caminho_completo}")
            # Enviar mensagem no WhatsApp com a imagem
            enviar_mensagem_whatsapp(nome, caminho_completo)
        else:
            print(f"Arquivo não encontrado para {nome}.")