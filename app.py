import threading
import tkinter as tk
from tkinter import scrolledtext
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import logging
import time
import re
from webdriver_manager.chrome import ChromeDriverManager
import pyttsx3  # Biblioteca para texto em fala
from queue import Queue  # Usaremos uma fila para processar os códigos

# Definições de variáveis globais
LOG_FILE = "log_cadastro_numeros.txt"
codigo_queue = Queue()  # Fila para armazenar os códigos lidos

# Inicializa a interface gráfica
root = tk.Tk()
root.title(f"Cadastrador Maercio CAST")

# Cria um ScrolledText para exibir logs na interface
log_area = scrolledtext.ScrolledText(root, width=50, height=20, state=tk.DISABLED, bg='light grey')
log_area.grid(row=3, column=0, columnspan=3, pady=5, padx=10, sticky='nsew')

# Configuração do logging para registrar logs em tempo real
logger = logging.getLogger("CadastroLogger")
logger.setLevel(logging.INFO)
logger.propagate = False  # Desativa a propagação para evitar duplicação

# Criar um FileHandler para registrar logs em um arquivo
if not any(isinstance(handler, logging.FileHandler) for handler in logger.handlers):
    file_handler = logging.FileHandler(LOG_FILE, mode='w', encoding='utf-8')
    file_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
    logger.addHandler(file_handler)

# Adicionar um TextHandler para atualizar a interface gráfica
class TextHandler(logging.Handler):
    def __init__(self, widget):
        logging.Handler.__init__(self)
        self.widget = widget

    def emit(self, record):
        def atualizar_text():
            msg = self.format(record)
            self.widget.config(state=tk.NORMAL)
            self.widget.insert(tk.END, msg + '\n')
            self.widget.yview(tk.END)
            self.widget.config(state=tk.DISABLED)
        root.after(0, atualizar_text)

# Adicionar o TextHandler se ainda não estiver adicionado
if not any(isinstance(handler, TextHandler) for handler in logger.handlers):
    text_handler = TextHandler(log_area)
    text_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
    logger.addHandler(text_handler)

# Função para registrar logs de forma segura na thread principal
def registrar_log(mensagem):
    def log():
        logger.info(mensagem)
    root.after(0, log)

# Função para esperar por um elemento
def waiting(driver, by, value, timeout=10):
    try:
        element = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
        return element
    except Exception:
        return None

# Função para configurar o navegador
def configurar_navegador():
    registrar_log("Configurando opções do navegador.")
    chrome_options = Options()
    chrome_options.add_argument("--force-device-scale-factor=1")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.set_window_size(1600, 1200)
    return driver

# Função para monitorar a tela e detectar quando está na tela de cadastro
def monitorar_tela(driver):
    try:
        registrar_log("Monitorando a navegação manual para detectar a tela de cadastro...")
        while True:
            input_field = waiting(driver, By.XPATH, "//input[@title='Digite ou Utilize um leitor de código de barras ou QRCode']", timeout=2)
            if input_field:
                registrar_log("Tela de cadastro detectada.")
                registrar_log("Posicionando o cursor no campo de input do programa.")
                entry_codigo.focus_set()  # Posiciona o cursor no campo de entrada do programa
                registrar_log("Aguardando entradas do leitor de código de barras.")
                break
            time.sleep(1)
    except Exception as e:
        registrar_log(f"Erro ao monitorar a tela: {str(e)}")

# Função para anunciar mensagens com áudio
def anunciar_mensagem(mensagem):
    engine = pyttsx3.init()
    engine.say(mensagem)
    engine.runAndWait()

# Função para capturar o código e adicionar à fila
def capturar_codigo(event):
    codigo_lido = event.widget.get()
    event.widget.delete(0, tk.END)
    
    match = re.search(r'\b\d{44}\b', codigo_lido)
    if match:
        id_nota = match.group(0)
        registrar_log(f"Código válido de 44 dígitos identificado: {id_nota}")
        codigo_queue.put(id_nota)  # Adiciona o código à fila
        anunciar_mensagem("ok")
    else:
        registrar_log("Nenhum código válido de 44 dígitos encontrado.")
        anunciar_mensagem("inválido")

# Função para verificar as mensagens de sucesso ou erro
def verificar_mensagem():
    msg = "Nenhuma msg de label encontrada"

    if driver.find_element(By.ID, "lblInfo"):
        msg = driver.find_element(By.ID, "lblInfo")
        registrar_log("Encontrou lblinfo")
    
    elif driver.find_element(By.ID, "lblErro"):
        msg = driver.find_element(By.ID, "lblErro")
        registrar_log("Encontrou lblErro")
  
    try:        
        msg = msg.text.lower()
        if "excedeu o prazo" in msg:      
            registrar_log("Nota já expirou.")
            anunciar_mensagem("Nota já expirou")
        elif "já existe" in msg:
            registrar_log("Nota já cadastrada no sistema.")
            anunciar_mensagem("Nota já cadastrada no sistema")
        elif "Doação registrada" in msg:
            registrar_log("Cadastrada com sucesso.")
            anunciar_mensagem("Cadastrada com sucesso")
        elif "possivel" or "possível" in msg:
            registrar_log("Parece que atingiu o limite de cadastros para esse usuário")
            anunciar_mensagem("Parece que atingiu o limite de cadastros para esse usuário")
        else:
            registrar_log("Erro não identificado: " + msg)
           
    except Exception as e:
        registrar_log(f"Erro ao verificar mensagens: {str(e)}")

# Função para processar os códigos da fila
def processar_fila():
    while True:
        if not codigo_queue.empty():
            id_nota = codigo_queue.get()  # Pega o próximo código da fila
            registrar_log(f"Processando o código: {id_nota}")
            
            input_field = driver.find_element(By.XPATH, "//input[@title='Digite ou Utilize um leitor de código de barras ou QRCode']")
            input_field.clear()

            # Posiciona o cursor no início do campo
            input_field.send_keys(Keys.HOME)
            time.sleep(0.1)  # Pequena pausa para garantir o posicionamento

            input_field.send_keys(id_nota)
            time.sleep(1)
            
            # Simula a tecla Enter
            input_field.send_keys(Keys.RETURN)
            verificar_mensagem()  # Verifica as mensagens de sucesso ou erro
        time.sleep(1)

# Função principal para iniciar o navegador e começar o monitoramento
def iniciar_navegador():
    global driver
    driver = configurar_navegador()
    driver.get("https://www.nfp.fazenda.sp.gov.br/login.aspx")
    registrar_log("Navegador iniciado. Por favor, faça login manualmente e navegue até a tela de cadastro.")
    
    # Inicia a thread para monitorar a tela
    threading.Thread(target=monitorar_tela, args=(driver,), daemon=True).start()

    # Inicia a thread para processar a fila de códigos
    threading.Thread(target=processar_fila, daemon=True).start()

# Interface gráfica para capturar códigos de barras
start_button = tk.Button(root, text="Iniciar Navegador", command=lambda: threading.Thread(target=iniciar_navegador).start())
start_button.grid(row=0, column=0, pady=5, padx=5)

entry_codigo = tk.Entry(root, width=50)
entry_codigo.grid(row=1, column=0, padx=10, pady=10)
entry_codigo.bind("<Return>", capturar_codigo)  # Captura quando o usuário pressiona Enter

root.grid_rowconfigure(3, weight=1)
root.grid_columnconfigure(0, weight=1)

# Mensagem de áudio após carregar a interface gráfica
def mensagem_audio_inicio():
    engine = pyttsx3.init()
    mensagem_fala = "Clique em iniciar navegador, navegue manualmente até a tela de cadastro e insira os códigos de barras."
    engine.say(mensagem_fala)
    engine.runAndWait()

root.after(300, lambda: threading.Thread(target=mensagem_audio_inicio).start())

root.mainloop()
