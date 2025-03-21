import threading
import tkinter as tk
from tkinter import scrolledtext, messagebox, simpledialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import logging
import time
import re
import pyttsx3
from queue import Queue

# Definições de variáveis globais
LOG_FILE = "log_cadastro_numeros.txt"
codigo_queue = Queue()  # Fila para armazenar os códigos lidos
codigos_erro = []  # Lista para armazenar códigos que deram erro
driver = None  # Variável global para o driver do navegador
SENHA_ENCERRAMENTO = "5510"  # Senha para forçar o encerramento do programa

# Controle de mensagens para evitar repetições
mensagens_exibidas = {
    "tela_cadastro_detectada": False,
    "login_necessario": False,
    "erro_navegacao": False
}

# Inicializa a interface gráfica
root = tk.Tk()
root.title(f"Cadastrador Maercio CAST - Melhorado")

# Contadores para estatísticas
contadores = {
    "sucesso": 0,
    "ja_cadastrada": 0,
    "expirada": 0,
    "erro": 0
}

# Cria um frame para os contadores
frame_stats = tk.Frame(root)
frame_stats.grid(row=0, column=0, columnspan=3, pady=5, padx=10, sticky='ew')

# Labels para contadores
lbl_sucesso = tk.Label(frame_stats, text="Sucesso: 0", width=15)
lbl_sucesso.grid(row=0, column=0, padx=5)

lbl_ja_cadastrada = tk.Label(frame_stats, text="Já Cadastradas: 0", width=15)
lbl_ja_cadastrada.grid(row=0, column=1, padx=5)

lbl_expirada = tk.Label(frame_stats, text="Expiradas: 0", width=15)
lbl_expirada.grid(row=0, column=2, padx=5)

lbl_erro = tk.Label(frame_stats, text="Erros: 0", width=15)
lbl_erro.grid(row=0, column=3, padx=5)

# Campo de entrada para os códigos
entry_codigo = tk.Entry(root, width=50)
entry_codigo.grid(row=1, column=0, columnspan=3, padx=10, pady=10)

# Botão para iniciar recadastro de códigos com erro
btn_recadastrar = tk.Button(root, text="Recadastrar Erros (0)", state=tk.DISABLED)
btn_recadastrar.grid(row=2, column=0, columnspan=3, pady=5)

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
def registrar_log(mensagem, exibir_apenas_uma_vez=False, chave_mensagem=None):
    """
    Registra uma mensagem no log.
    
    Args:
        mensagem: A mensagem a ser registrada
        exibir_apenas_uma_vez: Se True, a mensagem será exibida apenas uma vez
        chave_mensagem: Identificador único para a mensagem, usado quando exibir_apenas_uma_vez=True
    """
    if exibir_apenas_uma_vez and chave_mensagem:
        if mensagens_exibidas.get(chave_mensagem, False):
            return  # Não registra a mensagem se já foi exibida antes
        mensagens_exibidas[chave_mensagem] = True
        
    def log():
        logger.info(mensagem)
    root.after(0, log)

# Função para anunciar mensagens com áudio
def anunciar_mensagem(mensagem, exibir_apenas_uma_vez=False, chave_mensagem=None):
    """
    Anuncia uma mensagem com áudio.
    
    Args:
        mensagem: A mensagem a ser anunciada
        exibir_apenas_uma_vez: Se True, a mensagem será anunciada apenas uma vez
        chave_mensagem: Identificador único para a mensagem, usado quando exibir_apenas_uma_vez=True
    """
    if exibir_apenas_uma_vez and chave_mensagem:
        if mensagens_exibidas.get(chave_mensagem, False):
            return  # Não anuncia a mensagem se já foi anunciada antes
        mensagens_exibidas[chave_mensagem] = True
        
    try:
        engine = pyttsx3.init()
        engine.say(mensagem)
        engine.runAndWait()
    except Exception as e:
        registrar_log(f"Erro ao anunciar mensagem: {str(e)}")

# Função para atualizar os contadores na interface
def atualizar_contadores():
    def update():
        lbl_sucesso.config(text=f"Sucesso: {contadores['sucesso']}")
        lbl_ja_cadastrada.config(text=f"Já Cadastradas: {contadores['ja_cadastrada']}")
        lbl_expirada.config(text=f"Expiradas: {contadores['expirada']}")
        lbl_erro.config(text=f"Erros: {contadores['erro']}")
    
    root.after(0, update)

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
    
    try:
        # Usar o Selenium Manager integrado
        registrar_log("Tentando iniciar Chrome usando o driver integrado...")
        driver = webdriver.Chrome(options=chrome_options)
        registrar_log("Chrome iniciado com sucesso!")
        driver.set_window_size(1600, 1200)
        return driver
    except Exception as e:
        registrar_log(f"Não foi possível iniciar o Chrome: {str(e)}")
        
        # Mostrar mensagem para o usuário
        messagebox.showinfo("Ação necessária", 
            "Para resolver este problema:\n\n"
            "1. Baixe o ChromeDriver manualmente de https://chromedriver.chromium.org/downloads\n"
            "   Certifique-se de baixar a versão compatível com seu Chrome\n\n"
            "2. Descompacte o arquivo e coloque o chromedriver.exe na pasta do programa\n\n"
            "3. Reinicie o aplicativo")
        
        # Tenta inicializar com um caminho específico como último recurso
        try:
            registrar_log("Tentando iniciar com o chromedriver.exe na pasta do programa...")
            service = Service(executable_path="./chromedriver.exe")
            driver = webdriver.Chrome(service=service, options=chrome_options)
            driver.set_window_size(1600, 1200)
            return driver
        except Exception as e2:
            registrar_log(f"Falha final: {str(e2)}")
            raise

# Função para verificar a presença de um elemento
def elemento_presente(driver, by, value):
    try:
        driver.find_element(by, value)
        return True
    except:
        return False

# Função para verificar em qual tela o programa está
def verificar_tela_atual(driver):
    try:
        # Verifica se está na tela de login
        if elemento_presente(driver, By.ID, "UserName") and elemento_presente(driver, By.ID, "Password"):
            registrar_log("Detectada tela de login.", exibir_apenas_uma_vez=True, chave_mensagem="tela_login_detectada")
            return "login"
        
        # Verifica se está na tela de cadastro
        if elemento_presente(driver, By.XPATH, "//input[@title='Digite ou Utilize um leitor de código de barras ou QRCode']"):
            registrar_log("Detectada tela de cadastro.", exibir_apenas_uma_vez=True, chave_mensagem="tela_cadastro_detectada")
            return "cadastro"
        
        # Verifica se está na tela principal
        if "Principal.aspx" in driver.current_url:
            registrar_log("Detectada tela principal.", exibir_apenas_uma_vez=True, chave_mensagem="tela_principal_detectada")
            return "principal"
        
        # Outras possíveis telas podem ser adicionadas aqui
        
        # Se não identificou nenhuma tela conhecida
        registrar_log("Tela atual não identificada.", exibir_apenas_uma_vez=True, chave_mensagem="tela_nao_identificada")
        return "desconhecida"
    except Exception as e:
        registrar_log(f"Erro ao verificar tela atual: {str(e)}")
        return "erro"

# Função para navegar de volta à tela de cadastro
def voltar_para_tela_cadastro(driver):
    try:
        tela_atual = verificar_tela_atual(driver)
        
        # Se já estiver na tela de cadastro, não faz nada
        if tela_atual == "cadastro":
            return True
        
        # Se estiver na tela de login, faz o login
        if tela_atual == "login":
            registrar_log("Redirecionando: Necessário fazer login novamente.", exibir_apenas_uma_vez=True, chave_mensagem="login_necessario")
            anunciar_mensagem("Necessário fazer login novamente", exibir_apenas_uma_vez=True, chave_mensagem="login_necessario")
            return False  # Não tenta o processo automático, pede intervenção manual
        
        # Se estiver na tela principal
        if tela_atual == "principal" or tela_atual == "desconhecida":
            registrar_log("Tentando navegar de volta para a tela de cadastro...", exibir_apenas_uma_vez=False)
            
            # Verifica se o botão "Continuar" está presente e clica nele
            try:
                continuar_button = driver.find_element(By.ID, "btnContinuar")
                continuar_button.click()
                registrar_log("Botão 'Continuar' clicado.")
                time.sleep(2)
            except:
                registrar_log("Botão 'Continuar' não encontrado.")
            
            # Navega pelo menu
            try:
                # Clica em "Entidades"
                entidades_link = driver.find_element(By.XPATH, "//a[text()='Entidades']")
                entidades_link.click()
                registrar_log("Clicou em 'Entidades'.")
                time.sleep(2)
                
                # Clica em "Cadastramento de Cupons"
                cadastro_link = driver.find_element(By.XPATH, "//a[text()='Cadastramento de Cupons']")
                cadastro_link.click()
                registrar_log("Clicou em 'Cadastramento de Cupons'.")
                time.sleep(2)
                
                # Clica em "Prosseguir"
                prosseguir_button = driver.find_element(By.XPATH, "//input[@value='Prosseguir']")
                prosseguir_button.click()
                registrar_log("Clicou em 'Prosseguir'.")
                time.sleep(2)
                
                # Seleciona a entidade
                entidade_dropdown = Select(driver.find_element(By.ID, "ddlEntidadeFilantropica"))
                entidade_dropdown.select_by_visible_text("ASSOCIACAO WISE MADDNESS")
                registrar_log("Selecionou a entidade 'ASSOCIACAO WISE MADDNESS'.")
                time.sleep(2)
                
                # Clica em "Nova Nota"
                nova_nota_button = driver.find_element(By.XPATH, "//input[@value='Nova Nota']")
                nova_nota_button.click()
                registrar_log("Clicou em 'Nova Nota'.")
                time.sleep(2)
                
                # Simula ESC para fechar possíveis pop-ups
                action = ActionChains(driver)
                action.send_keys(Keys.ESCAPE).perform()
                registrar_log("Enviou tecla 'ESC'.")
                time.sleep(1)
                
                # Verifica se está na tela de cadastro
                if elemento_presente(driver, By.XPATH, "//input[@title='Digite ou Utilize um leitor de código de barras ou QRCode']"):
                    registrar_log("Navegou com sucesso para a tela de cadastro.")
                    # Resetamos a chave para permitir que a mensagem seja exibida novamente
                    mensagens_exibidas["tela_cadastro_detectada"] = False
                    return True
                else:
                    registrar_log("Falha ao navegar para a tela de cadastro.", exibir_apenas_uma_vez=True, chave_mensagem="falha_navegacao")
                    return False
            except Exception as e:
                registrar_log(f"Erro ao navegar para a tela de cadastro: {str(e)}", exibir_apenas_uma_vez=True, chave_mensagem="erro_navegacao")
                return False
        
        return False
    except Exception as e:
        registrar_log(f"Erro ao tentar voltar para a tela de cadastro: {str(e)}")
        return False

# Função para verificar as mensagens de feedback e atualizar contadores
def verificar_mensagem(driver):
    time.sleep(1)  # Aguarda um momento para a interface atualizar
    
    try:
        # Verificar mensagem de sucesso
        info_label = driver.find_element(By.ID, "lblInfo")
        if info_label and info_label.is_displayed():
            msg_info = info_label.text.lower()
            if "doação registrada" in msg_info:
                registrar_log("Cadastrada com sucesso.")
                anunciar_mensagem("ok")
                contadores["sucesso"] += 1
                atualizar_contadores()
                return "sucesso"
    except:
        pass
    
    try:
        # Verificar mensagem de erro
        erro_label = driver.find_element(By.ID, "lblErro")
        if erro_label and erro_label.is_displayed():
            msg_erro = erro_label.text.lower()
            if "excedeu o prazo" in msg_erro:
                registrar_log("Nota já expirou.")
                anunciar_mensagem("expirada")
                contadores["expirada"] += 1
                atualizar_contadores()
                return "expirada"
            elif "já existe" in msg_erro:
                registrar_log("Nota já cadastrada no sistema.")
                anunciar_mensagem("já cadastrada")
                contadores["ja_cadastrada"] += 1
                atualizar_contadores()
                return "ja_cadastrada"
            elif "não foi possível" in msg_erro:
                registrar_log("Erro: Não foi possível incluir o pedido.")
                anunciar_mensagem("erro")
                contadores["erro"] += 1
                atualizar_contadores()
                return "erro"
            else:
                registrar_log(f"Erro não identificado: {msg_erro}")
                anunciar_mensagem("erro")
                contadores["erro"] += 1
                atualizar_contadores()
                return "erro"
    except:
        pass
    
    # Verifica se ainda está na tela de cadastro
    if not elemento_presente(driver, By.XPATH, "//input[@title='Digite ou Utilize um leitor de código de barras ou QRCode']"):
        registrar_log("Não está mais na tela de cadastro. Tentando retornar...")
        if not voltar_para_tela_cadastro(driver):
            registrar_log("Não foi possível retornar à tela de cadastro automaticamente.")
            anunciar_mensagem("Intervenção manual necessária")
        return "fora_da_tela"
    
    # Se não encontrou nenhuma mensagem clara
    registrar_log("Não foi possível identificar o status da operação.")
    contadores["erro"] += 1
    atualizar_contadores()
    return "erro"

# Função para capturar o código e adicionar à fila
def capturar_codigo(event):
    codigo_lido = event.widget.get()
    event.widget.delete(0, tk.END)
    
    match = re.search(r'\b\d{44}\b', codigo_lido)
    if match:
        id_nota = match.group(0)
        registrar_log(f"Código válido de 44 dígitos identificado: {id_nota}")
        
        # Verifica se estamos na tela de cadastro antes de adicionar à fila
        if driver is not None:
            try:
                tela_atual = verificar_tela_atual(driver)
                if tela_atual != "cadastro":
                    registrar_log("Fora da tela de cadastro. Tentando voltar antes de processar o código...")
                    if voltar_para_tela_cadastro(driver):
                        registrar_log("Retornou à tela de cadastro. Processando o código.")
                        codigo_queue.put(id_nota)  # Adiciona o código à fila após voltar à tela
                    else:
                        registrar_log("Não foi possível voltar à tela de cadastro. Salvando código para processamento posterior.")
                        codigos_erro.append(id_nota)  # Salva para recadastro posterior
                        anunciar_mensagem("Código salvo para processamento posterior")
                else:
                    codigo_queue.put(id_nota)  # Adiciona o código à fila se já estiver na tela correta
            except Exception as e:
                registrar_log(f"Erro ao verificar tela: {str(e)}. Salvando código para processamento posterior.")
                codigos_erro.append(id_nota)  # Salva para recadastro posterior
                anunciar_mensagem("Código salvo para processamento posterior")
        else:
            registrar_log("Navegador não iniciado. Salvando código para processamento posterior.")
            codigos_erro.append(id_nota)  # Salva para recadastro posterior
            anunciar_mensagem("Navegador não iniciado. Código salvo.")
    else:
        registrar_log("Nenhum código válido de 44 dígitos encontrado.")
        anunciar_mensagem("inválido")

# Função para processar os códigos da fila
def processar_fila():
    while True:
        if not codigo_queue.empty():
            id_nota = codigo_queue.get()  # Pega o próximo código da fila
            registrar_log(f"Processando o código: {id_nota}")
            
            # Verifica se estamos na tela de cadastro
            tela_atual = verificar_tela_atual(driver)
            if tela_atual != "cadastro":
                registrar_log("Não estamos na tela de cadastro. Tentando retornar...")
                if not voltar_para_tela_cadastro(driver):
                    # Se não conseguiu voltar, coloca o código de volta na fila e aguarda
                    codigo_queue.put(id_nota)
                    registrar_log("Não foi possível retornar à tela de cadastro. Aguardando intervenção manual.")
                    anunciar_mensagem("Intervenção manual necessária")
                    time.sleep(5)  # Aguarda um pouco antes de tentar novamente
                    continue
            
            try:
                input_field = driver.find_element(By.XPATH, "//input[@title='Digite ou Utilize um leitor de código de barras ou QRCode']")
                input_field.clear()

                # Posiciona o cursor no início do campo
                input_field.send_keys(Keys.HOME)
                time.sleep(0.1)  # Pequena pausa para garantir o posicionamento

                input_field.send_keys(id_nota)
                time.sleep(0.5)
                
                # Simula a tecla Enter
                input_field.send_keys(Keys.RETURN)
                time.sleep(0.5)
                
                # Verifica o resultado
                resultado = verificar_mensagem(driver)
                
                # Se deu erro ou está fora da tela, adiciona à lista para recadastro
                if resultado == "erro" or resultado == "fora_da_tela":
                    codigos_erro.append(id_nota)
                    
                    # Se estiver fora da tela, tenta voltar para a tela de cadastro
                    if resultado == "fora_da_tela":
                        voltar_para_tela_cadastro(driver)
            except Exception as e:
                registrar_log(f"Erro ao processar código: {str(e)}")
                codigos_erro.append(id_nota)
                contadores["erro"] += 1
                atualizar_contadores()
                
                # Verifica se o navegador ainda está respondendo
                try:
                    driver.current_url  # Apenas para testar se o navegador responde
                except:
                    registrar_log("Navegador não está respondendo. Possível crash.")
                    anunciar_mensagem("Navegador não responde")
                    # Aqui você poderia implementar uma lógica para reiniciar o navegador
                
                # Tenta voltar para a tela de cadastro após o erro
                voltar_para_tela_cadastro(driver)
                
        time.sleep(0.5)

# Função para monitorar a tela e detectar quando está na tela de cadastro
def monitorar_tela(driver):
    try:
        registrar_log("Monitorando a navegação manual para detectar a tela de cadastro...", exibir_apenas_uma_vez=True, chave_mensagem="monitorando_navegacao")
        while True:
            input_field = waiting(driver, By.XPATH, "//input[@title='Digite ou Utilize um leitor de código de barras ou QRCode']", timeout=2)
            if input_field:
                registrar_log("Tela de cadastro detectada.", exibir_apenas_uma_vez=True, chave_mensagem="tela_cadastro_detectada")
                registrar_log("Posicionando o cursor no campo de input do programa.", exibir_apenas_uma_vez=True, chave_mensagem="posicionando_cursor")
                entry_codigo.focus_set()  # Posiciona o cursor no campo de entrada do programa
                registrar_log("Aguardando entradas do leitor de código de barras.", exibir_apenas_uma_vez=True, chave_mensagem="aguardando_entradas")
                break
            time.sleep(1)
    except Exception as e:
        registrar_log(f"Erro ao monitorar a tela: {str(e)}")

# Função para verificar periodicamente se ainda estamos na tela de cadastro
def monitor_tela_cadastro():
    ultima_verificacao = 0  # Timestamp da última verificação
    intervalo_verificacao = 5  # Verificar a cada 5 segundos
    
    while True:
        if driver is not None:
            try:
                # Verifica apenas se não estiver processando um código e se passou o intervalo
                tempo_atual = time.time()
                if (codigo_queue.empty() and 
                    len(codigos_erro) == 0 and 
                    tempo_atual - ultima_verificacao >= intervalo_verificacao):
                    
                    ultima_verificacao = tempo_atual
                    tela_atual = verificar_tela_atual(driver)
                    
                    if tela_atual != "cadastro":
                        # Reiniciamos o flag quando detectamos que saímos da tela
                        # para permitir que a mensagem seja exibida na próxima navegação bem-sucedida
                        mensagens_exibidas["tela_cadastro_detectada"] = False
                        
                        registrar_log("Detectado que saímos da tela de cadastro. Tentando retornar...", 
                                     exibir_apenas_uma_vez=True, 
                                     chave_mensagem="saimos_tela_cadastro")
                        voltar_para_tela_cadastro(driver)
            except Exception as e:
                registrar_log(f"Erro ao monitorar tela de cadastro: {str(e)}", 
                             exibir_apenas_uma_vez=True, 
                             chave_mensagem="erro_monitor_tela")
        time.sleep(5)  # Verificação a cada 5 segundos

# Função para exibir a quantidade de códigos salvos para recadastro
def atualizar_botao_recadastro():
    def update():
        if codigos_erro:
            btn_recadastrar.config(
                text=f"Recadastrar Erros ({len(codigos_erro)})", 
                state=tk.NORMAL
            )
        else:
            btn_recadastrar.config(
                text="Recadastrar Erros", 
                state=tk.DISABLED
            )
    root.after(0, update)

# Função para recadastrar códigos que deram erro
def recadastrar_codigos():
    if not codigos_erro:
        registrar_log("Não há códigos para recadastrar.")
        messagebox.showinfo("Recadastro", "Não há códigos para recadastrar.")
        return
        
    atualizar_botao_recadastro()  # Atualiza o botão imediatamente
    
    # Verifica se estamos na tela de cadastro antes de iniciar o recadastro
    tela_atual = verificar_tela_atual(driver)
    if tela_atual != "cadastro":
        registrar_log("Não estamos na tela de cadastro. Tentando retornar...")
        if not voltar_para_tela_cadastro(driver):
            registrar_log("Não foi possível retornar à tela de cadastro. Recadastro abortado.")
            messagebox.showwarning("Erro", "Não foi possível retornar à tela de cadastro. Por favor, navegue manualmente até a tela de cadastro e tente novamente.")
            anunciar_mensagem("Intervenção manual necessária")
            return
    
    registrar_log(f"Iniciando recadastro de {len(codigos_erro)} códigos que deram erro...")
    # Cria uma cópia da lista para não afetar a iteração
    codigos_para_recadastrar = codigos_erro.copy()
    codigos_erro.clear()
    
    sucessos_recadastro = 0
    
    for i, codigo in enumerate(codigos_para_recadastrar, 1):
        registrar_log(f"Recadastrando código {i}/{len(codigos_para_recadastrar)}: {codigo}")
        
        # Verifica novamente se estamos na tela de cadastro
        if verificar_tela_atual(driver) != "cadastro":
            registrar_log("Saímos da tela de cadastro. Tentando retornar...")
            if not voltar_para_tela_cadastro(driver):
                registrar_log("Não foi possível retornar à tela de cadastro. Interrompendo recadastro.")
                # Adiciona os códigos restantes de volta à lista de erros
                codigos_erro.extend(codigos_para_recadastrar[i-1:])
                break
        
        try:
            input_field = driver.find_element(By.XPATH, "//input[@title='Digite ou Utilize um leitor de código de barras ou QRCode']")
            input_field.clear()
            input_field.send_keys(Keys.HOME)
            time.sleep(0.1)
            input_field.send_keys(codigo)
            time.sleep(0.5)
            input_field.send_keys(Keys.RETURN)
            time.sleep(0.5)
            
            resultado = verificar_mensagem(driver)
            if resultado == "sucesso":
                sucessos_recadastro += 1
            elif resultado == "erro" or resultado == "fora_da_tela":
                # Se falhar novamente, coloca de volta na lista de erros
                codigos_erro.append(codigo)
                
                # Se estiver fora da tela, tenta voltar para a tela de cadastro
                if resultado == "fora_da_tela":
                    if not voltar_para_tela_cadastro(driver):
                        registrar_log("Não foi possível retornar à tela de cadastro. Interrompendo recadastro.")
                        # Adiciona os códigos restantes de volta à lista de erros
                        codigos_erro.extend(codigos_para_recadastrar[i:])
                        break
        except Exception as e:
            registrar_log(f"Erro ao recadastrar código: {str(e)}")
            codigos_erro.append(codigo)
            
            # Tenta voltar para a tela de cadastro após o erro
            if not voltar_para_tela_cadastro(driver):
                registrar_log("Não foi possível retornar à tela de cadastro. Interrompendo recadastro.")
                # Adiciona os códigos restantes de volta à lista de erros
                codigos_erro.extend(codigos_para_recadastrar[i:])
                break
    
    if codigos_erro:
        mensagem = f"Recadastro concluído. {sucessos_recadastro} códigos recadastrados com sucesso. {len(codigos_erro)} códigos ainda com erro."
        registrar_log(mensagem)
        messagebox.showinfo("Recadastro Concluído", mensagem)
    else:
        mensagem = f"Recadastro concluído com sucesso para todos os {sucessos_recadastro} códigos!"
        registrar_log(mensagem)
        messagebox.showinfo("Recadastro Concluído", mensagem)
        btn_recadastrar.config(state=tk.DISABLED)

# Classe para uma caixa de diálogo de senha personalizada com tamanho ajustável
class SenhaDialog(tk.Toplevel):
    def __init__(self, parent, title=None):
        tk.Toplevel.__init__(self, parent)
        self.transient(parent)
        
        if title:
            self.title(title)
        
        self.parent = parent
        self.result = None
        
        # Cria o corpo da janela
        body = tk.Frame(self, padx=20, pady=10)
        body.pack(fill=tk.BOTH, expand=True)
        
        # Configura a largura/altura da janela
        self.geometry("300x150")  # Ajustável: largura x altura
        
        # Cria e posiciona os widgets
        self.create_widgets(body)
        
        # Torna a janela modal (bloqueia interação com a janela pai)
        self.grab_set()
        
        # Posiciona a janela relativa à janela pai
        self.geometry("+%d+%d" % (parent.winfo_rootx() + 50,
                                   parent.winfo_rooty() + 50))
        
        # Define o comportamento para ao fechar a janela
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        
        # Configura para dar foco ao campo de senha
        self.senha_entry.focus_set()
        
        # Inicia o loop de espera
        self.wait_window(self)
    
    def create_widgets(self, master):
        # Mensagem na caixa de diálogo
        mensagem_label = tk.Label(master, text="Digite a senha para encerrar:", font=("Arial", 12))
        mensagem_label.pack(pady=10)
        
        # Campo de entrada da senha
        self.senha_entry = tk.Entry(master, show="*", width=20, font=("Arial", 12))
        self.senha_entry.pack(pady=10)
        self.senha_entry.bind("<Return>", self.ok)
        
        # Frame para os botões
        button_frame = tk.Frame(master)
        button_frame.pack(fill=tk.X, pady=10)
        
        # Botão OK
        ok_button = tk.Button(button_frame, text="OK", width=10, command=self.ok, font=("Arial", 10))
        ok_button.pack(side=tk.LEFT, padx=5, expand=True)
        
        # Botão Cancelar
        cancel_button = tk.Button(button_frame, text="Cancelar", width=10, command=self.cancel, font=("Arial", 10))
        cancel_button.pack(side=tk.RIGHT, padx=5, expand=True)
    
    def ok(self, event=None):
        self.result = self.senha_entry.get()
        self.destroy()
    
    def cancel(self, event=None):
        self.result = None
        self.destroy()

# Função para mostrar a caixa de diálogo de senha personalizada
def pedir_senha(parent, titulo):
    dialog = SenhaDialog(parent, titulo)
    return dialog.result

# Função para gerar relatório em arquivo TXT
def gerar_relatorio():
    try:
        # Nome do arquivo baseado na data e hora atual
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        nome_arquivo = f"relatorio_cadastro_{timestamp}.txt"
        
        registrar_log(f"Gerando relatório: {nome_arquivo}")
        
        with open(nome_arquivo, "w", encoding="utf-8") as arquivo:
            # Cabeçalho do relatório
            arquivo.write("=" * 50 + "\n")
            arquivo.write(f"RELATÓRIO DE CADASTRO DE NOTAS FISCAIS\n")
            arquivo.write(f"Data/Hora: {time.strftime('%d/%m/%Y %H:%M:%S')}\n")
            arquivo.write("=" * 50 + "\n\n")
            
            # Estatísticas gerais
            arquivo.write("ESTATÍSTICAS DE CADASTRO:\n\n")
            arquivo.write(f"Notas cadastradas com sucesso: {contadores['sucesso']}\n")
            arquivo.write(f"Notas já cadastradas no sistema: {contadores['ja_cadastrada']}\n")
            arquivo.write(f"Notas expiradas: {contadores['expirada']}\n")
            arquivo.write(f"Notas com erro: {contadores['erro']}\n")
            
            total_processadas = contadores['sucesso'] + contadores['ja_cadastrada'] + contadores['expirada'] + contadores['erro']
            arquivo.write(f"Total de notas processadas: {total_processadas}\n\n")
            
            # Calcular percentuais
            if total_processadas > 0:
                sucesso_percent = (contadores['sucesso'] / total_processadas) * 100
                ja_cadastrada_percent = (contadores['ja_cadastrada'] / total_processadas) * 100
                expirada_percent = (contadores['expirada'] / total_processadas) * 100
                erro_percent = (contadores['erro'] / total_processadas) * 100
                
                arquivo.write(f"Taxa de sucesso: {sucesso_percent:.2f}%\n")
                arquivo.write(f"Taxa de já cadastradas: {ja_cadastrada_percent:.2f}%\n")
                arquivo.write(f"Taxa de expiradas: {expirada_percent:.2f}%\n")
                arquivo.write(f"Taxa de erro: {erro_percent:.2f}%\n\n")
            
            # Notas com erro pendentes
            arquivo.write("NOTAS PENDENTES COM ERRO:\n\n")
            if codigos_erro:
                for i, codigo in enumerate(codigos_erro, 1):
                    arquivo.write(f"{i}. {codigo}\n")
                arquivo.write(f"\nTotal de notas pendentes: {len(codigos_erro)}\n")
            else:
                arquivo.write("Não há notas pendentes com erro.\n")
            
            # Rodapé do relatório
            arquivo.write("\n" + "=" * 50 + "\n")
            arquivo.write("Fim do Relatório\n")
        
        registrar_log(f"Relatório gerado com sucesso: {nome_arquivo}")
        return nome_arquivo
    except Exception as e:
        registrar_log(f"Erro ao gerar relatório: {str(e)}")
        return None

# Função para parar o processamento de forma segura
def parar_processamento_seguro():
    if codigos_erro:
        # Se existem códigos com erro, exibe um alerta
        resposta = messagebox.askquestion("Atenção", 
                         f"Existem {len(codigos_erro)} notas que precisam ser recadastradas!\n\n"
                         "Deseja realmente encerrar o programa?")
        
        if resposta == "yes":
            # Se o usuário confirma que quer fechar, pede a senha usando nossa caixa de diálogo personalizada
            senha = pedir_senha(root, "Segurança")
            
            if senha == SENHA_ENCERRAMENTO:
                # Gera relatório antes de encerrar
                nome_relatorio = gerar_relatorio()
                if nome_relatorio:
                    messagebox.showinfo("Relatório Gerado", 
                                       f"Um relatório foi gerado no arquivo:\n{nome_relatorio}")
                
                # Se a senha estiver correta, encerra o programa
                if driver:
                    try:
                        driver.quit()
                    except:
                        pass
                root.destroy()
            else:
                messagebox.showwarning("Senha Incorreta", 
                                      "Senha incorreta. O programa continuará em execução.")
    else:
        # Se não há códigos com erro, pede confirmação simples
        resposta = messagebox.askquestion("Confirmar Saída", 
                                         "Deseja realmente encerrar o programa?")
        if resposta == "yes":
            # Gera relatório antes de encerrar
            nome_relatorio = gerar_relatorio()
            if nome_relatorio:
                messagebox.showinfo("Relatório Gerado", 
                                   f"Um relatório foi gerado no arquivo:\n{nome_relatorio}")
            
            if driver:
                try:
                    driver.quit()
                except:
                    pass
            root.destroy()

# Função para lidar com o evento de tentativa de fechar a janela
def on_closing():
    if codigos_erro:
        # Se existem códigos com erro, exibe um alerta
        resposta = messagebox.askquestion("Atenção", 
                         f"Existem {len(codigos_erro)} notas que precisam ser recadastradas!\n\n"
                         "Deseja realmente encerrar o programa?")
        
        if resposta == "yes":
            # Se o usuário confirma que quer fechar, pede a senha usando nossa caixa de diálogo personalizada
            senha = pedir_senha(root, "Segurança")
            
            if senha == SENHA_ENCERRAMENTO:
                # Gera relatório antes de encerrar
                nome_relatorio = gerar_relatorio()
                if nome_relatorio:
                    messagebox.showinfo("Relatório Gerado", 
                                       f"Um relatório foi gerado no arquivo:\n{nome_relatorio}")
                
                # Se a senha estiver correta, encerra o programa
                if driver:
                    try:
                        driver.quit()
                    except:
                        pass
                root.destroy()
            else:
                messagebox.showwarning("Senha Incorreta", 
                                      "Senha incorreta. O programa continuará em execução.")
    else:
        # Se não há códigos com erro, pede confirmação simples
        resposta = messagebox.askquestion("Confirmar Saída", 
                                         "Deseja realmente encerrar o programa?")
        if resposta == "yes":
            # Gera relatório antes de encerrar
            nome_relatorio = gerar_relatorio()
            if nome_relatorio:
                messagebox.showinfo("Relatório Gerado", 
                                   f"Um relatório foi gerado no arquivo:\n{nome_relatorio}")
            
            if driver:
                try:
                    driver.quit()
                except:
                    pass
            root.destroy()

# Função para lidar com o evento de tentativa de fechar a janela
def on_closing():
    if codigos_erro:
        # Se existem códigos com erro, exibe um alerta
        resposta = messagebox.askquestion("Atenção", 
                         f"Existem {len(codigos_erro)} notas que precisam ser recadastradas!\n\n"
                         "Deseja realmente encerrar o programa?")
        
        if resposta == "yes":
            # Se o usuário confirma que quer fechar, pede a senha
            senha = tk.simpledialog.askstring("Segurança", 
                                             "Digite a senha para encerrar:", 
                                             show='*')
            
            if senha == SENHA_ENCERRAMENTO:
                # Se a senha estiver correta, encerra o programa
                if driver:
                    try:
                        driver.quit()
                    except:
                        pass
                root.destroy()
            else:
                messagebox.showwarning("Senha Incorreta", 
                                      "Senha incorreta. O programa continuará em execução.")
    else:
        # Se não há códigos com erro, pede confirmação simples
        resposta = messagebox.askquestion("Confirmar Saída", 
                                         "Deseja realmente encerrar o programa?")
        if resposta == "yes":
            if driver:
                try:
                    driver.quit()
                except:
                    pass
            root.destroy()

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
    
    # Inicia a thread para monitorar periodicamente se ainda estamos na tela de cadastro
    threading.Thread(target=monitor_tela_cadastro, daemon=True).start()

# Configuração dos botões
start_button = tk.Button(root, text="Iniciar Navegador", command=lambda: threading.Thread(target=iniciar_navegador).start())
start_button.grid(row=4, column=0, pady=5, padx=5)

# Adiciona botão para voltar à tela de cadastro manualmente
btn_voltar_cadastro = tk.Button(root, text="Voltar para Cadastro", 
                              command=lambda: threading.Thread(target=lambda: voltar_para_tela_cadastro(driver)).start())
btn_voltar_cadastro.grid(row=4, column=1, pady=5, padx=5)

# Adiciona botão para parar o processamento de forma segura
btn_parar = tk.Button(root, text="Encerrar Programa", command=parar_processamento_seguro)
btn_parar.grid(row=4, column=2, pady=5, padx=5)

# Configura o botão de recadastro
btn_recadastrar.config(command=lambda: threading.Thread(target=recadastrar_codigos).start())

# Configura redimensionamento da janela
root.grid_rowconfigure(3, weight=1)
root.grid_columnconfigure(0, weight=1)

# Configurar o protocolo de fechamento da janela
root.protocol("WM_DELETE_WINDOW", on_closing)

# Atualiza o botão de recadastro periodicamente
def monitor_botao_recadastro():
    while True:
        atualizar_botao_recadastro()
        time.sleep(2)

threading.Thread(target=monitor_botao_recadastro, daemon=True).start()

# Vincula o evento Return ao campo de entrada
entry_codigo.bind("<Return>", capturar_codigo)

# Mensagem de áudio após carregar a interface gráfica
def mensagem_audio_inicio():
    engine = pyttsx3.init()
    mensagem_fala = "Clique em iniciar navegador, navegue manualmente até a tela de cadastro e insira os códigos de barras."
    engine.say(mensagem_fala)
    engine.runAndWait()

root.after(300, lambda: threading.Thread(target=mensagem_audio_inicio).start())

root.mainloop()