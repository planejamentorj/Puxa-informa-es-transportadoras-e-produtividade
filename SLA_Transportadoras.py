import os
import json
import getpass
import tkinter as tk
from tkinter import simpledialog, messagebox
from tkcalendar import DateEntry
import os
import tkinter as tk
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import shutil
import win32com.client
import calendar
import pythoncom
import glob
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import time
import sys
import tkinter as tk
from tkinter import messagebox
import traceback
import getpass
import os
import json
import getpass
import tkinter as tk
from tkinter import simpledialog, messagebox
from tkcalendar import DateEntry
import os
import tkinter as tk
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import shutil
import win32com.client
import calendar
import pythoncom
import glob
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime


##### Login Sap #####
D


def excecao_global(exc_type, exc_value, exc_traceback):
    # Monta a mensagem de erro detalhada
    mensagem = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    mensagem = "Erro identificado durante a execução do script. Solicito que reinicie o processo."

    # Cria a janela oculta
    root = tk.Tk()
    root.withdraw()

    # Mostra popup de erro
    messagebox.showerror("Erro no programa", mensagem)

    root.destroy()


# Substitui o hook padrão de exceção
sys.excepthook = excecao_global




hoje = datetime.today()

# --- Data de ontem ---
data_ontem = hoje - timedelta(days=1)
data_ontem_str = data_ontem.strftime("%d%m%Y")

# --- Primeiro dia da última semana do mês passado ---
ano = hoje.year
mes = hoje.month

if mes == 1:
    mes_passado = 12
    ano_passado = ano - 1
else:
    mes_passado = mes - 1
    ano_passado = ano

# último dia do mês passado
ultimo_dia = calendar.monthrange(ano_passado, mes_passado)[1]
data_final = datetime(ano_passado, mes_passado, ultimo_dia)

# primeiro dia da última semana (6 dias antes do último dia)
semana_mes_passado = data_final - timedelta(days=6)
semana_mes_passado_str = semana_mes_passado.strftime("%d%m%Y")


# Data tarefas
hoje = datetime.today()
dia_semana = hoje.weekday()  # Segunda=0, Domingo=6

if dia_semana == 0:  # Segunda-feira
    # Sexta até Domingo
    inicio_periodo_tarefas = (hoje - timedelta(days=3)).strftime('%d.%m.%Y')  # Sexta
    fim_periodo_tarefas = (hoje - timedelta(days=1)).strftime('%d.%m.%Y')     # Domingo
else:
    ontem = hoje - timedelta(days=1)
    inicio_periodo_tarefas = ontem.strftime('%d.%m.%Y')
    fim_periodo_tarefas = ontem.strftime('%d.%m.%Y')



###########################################
# Data de hoje
mes_atual = datetime.now().strftime("%m")

hoje = datetime.today()
dia_semana = hoje.weekday()  # Segunda=0, Domingo=6

if dia_semana == 0:  # Segunda-feira
    # Sexta até Domingo
    inicio_periodo = (hoje - timedelta(days=3)).strftime('%d%m%Y')  # Sexta
    fim_periodo = (hoje - timedelta(days=1)).strftime('%d%m%Y')     # Domingo
else:
    ontem = hoje - timedelta(days=1)
    inicio_periodo = ontem.strftime('%d%m%Y')
    fim_periodo = ontem.strftime('%d%m%Y')

print("Início do período:", inicio_periodo)
print("Fim do período:", fim_periodo)



##########





# --- Configurar pastas e arquivo de credenciais ---
nome_pc = os.path.expanduser("~")
pasta_credenciais = os.path.join(
    fr"{nome_pc}\Labs TI\Dados Gmill - Planejamento\Planejamento - Controle de Indicadores\27 - Acessos\Senhas"
)
os.makedirs(pasta_credenciais, exist_ok=True)
arquivo_credenciais = os.path.join(pasta_credenciais, "credenciais.json")

# --- Função para carregar credenciais ---
def carregar_credenciais():
    if os.path.exists(arquivo_credenciais):
        with open(arquivo_credenciais, "r") as f:
            dados = json.load(f)
            # Se o arquivo armazenar uma string antiga, converte para dicionário
            if isinstance(dados, dict):
                return dados
            else:
                return {"planejamento.rj": dados, nome_pc: "Sm@708090"}
    else:
        # Arquivo não existe: cria dicionário padrão
        credenciais = {
            "planejamento.rj": "Eduardo@12052005",
            nome_pc: "Sm@708090"
        }
        with open(arquivo_credenciais, "w") as f:
            json.dump(credenciais, f, indent=4)
        return credenciais

# --- Função para salvar credenciais ---
def salvar_credenciais(credenciais):
    with open(arquivo_credenciais, "w") as f:
        json.dump(credenciais, f, indent=4)

# --- Detectar usuário e definir login SAP ---
usuario = getpass.getuser()
credenciais = carregar_credenciais()

if usuario == "planejamento.rj":
    sap_user = "EVIEIRA"
    sap_password = credenciais["planejamento.rj"]
else:
    sap_user = "svenancio"
    sap_password = credenciais[f"{nome_pc}"]

# --- Tkinter para alterar senha ---
def alterar_senha():
    global credenciais, sap_password
    nova_senha = simpledialog.askstring(
        "Alterar senha", f"Digite a nova senha para {sap_user}:", show="*"
    )
    if nova_senha:
        if usuario == "planejamento.rj":
            credenciais["planejamento.rj"] = nova_senha
            sap_password = nova_senha
        else:
            credenciais[f"{nome_pc}"] = nova_senha
            sap_password = nova_senha
        salvar_credenciais(credenciais)
        messagebox.showinfo("Sucesso", "Senha alterada com sucesso!")

# --- Funções ---
def alterar_senha():
    global sap_user
    nova_senha = simpledialog.askstring("🔑 Alterar Senha", f"Digite a nova senha para {sap_user}:", show="*")
    if nova_senha:
        messagebox.showinfo("✅ Sucesso", "Senha alterada com sucesso!")

def pegar_datas():
    global data_inicio_selecionada, data_final_selecionada
    data_inicio_selecionada = cal_inicio.get_date()
    data_final_selecionada = cal_final.get_date()
    messagebox.showinfo("📅 Datas Selecionadas", f"Início: {data_inicio_selecionada}\nFim: {data_final_selecionada}")
    root.destroy()

def mostrar_info():
    texto = """ℹ️ Informações:
- Confirme as datas corretamente.
- Alterar a senha SAP quando necessário.
- Clique em 'Confirmar Datas' para prosseguir."""
    messagebox.showinfo("Informações", texto)

# --- Janela principal ---
root = tk.Tk()
root.title("⚙️ Configuração SAP - Datas e Acessos")
root.geometry("600x440")
root.resizable(False, False)

# --- Frame datas ---
frame_datas = tk.LabelFrame(root, text="📅 Período", padx=15, pady=15, font=("Segoe UI", 11, "bold"))
frame_datas.pack(padx=20, pady=10, fill="x")

tk.Label(frame_datas, text="Data Início:", font=("Segoe UI", 10, "bold")).pack(anchor="w")
cal_inicio = DateEntry(frame_datas, date_pattern='dd/mm/yyyy', width=18)
cal_inicio.pack(pady=5)

tk.Label(frame_datas, text="Data Final:", font=("Segoe UI", 10, "bold")).pack(anchor="w")
cal_final = DateEntry(frame_datas, date_pattern='dd/mm/yyyy', width=18)
cal_final.pack(pady=5)

tk.Button(frame_datas, text="✅ Confirmar Datas", command=pegar_datas, bg="#4CAF50", fg="white",
          font=("Segoe UI", 10, "bold")).pack(pady=10)

# --- Frame senha ---
frame_senha = tk.LabelFrame(root, text="🔑 Configuração SAP", padx=15, pady=15, font=("Segoe UI", 11, "bold"))
frame_senha.pack(padx=20, pady=10, fill="x")

tk.Label(frame_senha, text=f"Usuário atual: {usuario}\nLogin SAP: {sap_user}", justify="left",
         font=("Segoe UI", 10)).pack(anchor="w", pady=5)
tk.Button(frame_senha, text="🔄 Trocar Senha", command=alterar_senha, bg="#2196F3", fg="white",
          font=("Segoe UI", 10, "bold")).pack(pady=5)

# --- Botão informações ---

btn_info =tk.Button(root, text="ℹ️ Mais informações", command=mostrar_info, bg="#FF9800", fg="white",
          font=("Segoe UI", 10, "bold"))
btn_info.place(relx=1.0, x=-24, y=24, anchor="ne")

root.mainloop()

print(f"Datas selecionadas: {data_inicio_selecionada} a {data_final_selecionada}")
print(f"Usuário SAP: {sap_user}")

# Caminhos das pastas
pasta1 = fr"{nome_pc}\Downloads\Supera dados CD-ES"
pasta2 = fr"{nome_pc}\Downloads\DATA HORA NF'S"

# Função para criar/substituir pasta
def criar_ou_substituir(pasta):
    if os.path.exists(pasta):
        shutil.rmtree(pasta)  # apaga a pasta antiga
    os.makedirs(pasta)        # cria a pasta nova
time.sleep(3)
# Criar/substituir pastas
criar_ou_substituir(pasta1)
time.sleep(3)
criar_ou_substituir(pasta2)

print("Pastas criadas/substituídas com sucesso!")



# --- Configura o navegador  km ---
navegador = webdriver.Chrome()
navegador.maximize_window()
navegador.implicitly_wait(600)  # espera até 600 segundos (10 minutos) por qualquer elemento
# --- Abre a página de login ---
navegador.get('https://kmtransportes.eslcloud.com.br/guests/sign_in')

# --- Preenche login ---
WebDriverWait(navegador, 10).until(
    EC.presence_of_element_located((By.ID, "guest_email"))
).send_keys("transportesrj@gmill.com.br")

navegador.find_element(By.ID, "guest_password").send_keys("12345678")

# --- Clica no botão de login ---
WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/form/div[2]/div/div[1]/input"))
).click()

# --- Navegação pelo menu ---
WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div/ul[2]/li[1]/a/span"))
).click()

WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div/ul[2]/li[1]/ul/li[3]/a/span"))
).click()

# --- Seleciona o campo e insere a data ---
campo = WebDriverWait(navegador, 10).until(
    EC.presence_of_element_located(
        (By.XPATH, "/html/body/div[4]/div/div[2]/div/div/div/div/div[2]/div[1]/div/div[2]/form/div[1]/div[1]/div/div[6]/div/span/input")
    )
)

campo.click()  # foca no campo
campo.send_keys(Keys.CONTROL, "a")  # seleciona tudo
campo.send_keys(Keys.BACKSPACE)  # apaga o conteúdo
campo.send_keys(data_inicio_selecionada.strftime('%d/%m/%Y'))  # insere a data de início
campo.send_keys(data_final_selecionada.strftime('%d/%m/%Y'))  # insere a data de início
# --- Pausa para conferência ---

time.sleep(2)
WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/div[8]/div[3]/div/button[1]"))
).click()

time.sleep(2)
WebDriverWait(navegador, 20).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div/div[2]/div/div/div/div/div[2]/div[2]/div[2]/div/div[2]/div/div[3]/div/table/thead/tr/th[12]/a"))
).click()


WebDriverWait(navegador, 1200).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/ul[1]/li/a/span"))
).click()


WebDriverWait(navegador, 20).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/ul[1]/li/ul/div[1]/div[1]/li/a/div[2]"))
).click()


WebDriverWait(navegador, 20).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/ul[1]/li/ul/div[3]"))
).click()


download_path = os.path.join(os.path.expanduser("~"), "Downloads")
arquivo_baixado = False
prefixo_arquivo = "notas-fiscais_"

print("Aguardando download do arquivo...")

while not arquivo_baixado:
    for arquivo in os.listdir(download_path):
        if arquivo.startswith(prefixo_arquivo) and arquivo.endswith(".xlsx"):
            if not arquivo.endswith(".crdownload") and not arquivo.endswith(".part"):
                print(f"Download concluído: {arquivo}")
                arquivo_baixado = True
                arquivo_final = arquivo
                break
    time.sleep(1)

# --- Fecha o navegador após o download ---
navegador.quit()
print("Navegador fechado.")

# --- Mover arquivo baixado para pasta destino e renomear ---
destino_path = fr"{nome_pc}\Labs TI\Dados Gmill - Planejamento\Planejamento - Controle de Indicadores\25 - Transportadoras SLA\Supera log\DADOS KM"
os.makedirs(destino_path, exist_ok=True)

# Define novo nome do arquivo
novo_nome = f"2025 - {mes_atual}.xlsx"

origem = os.path.join(download_path, arquivo_final)
destino = os.path.join(destino_path, novo_nome)

shutil.move(origem, destino)
print(f"Arquivo movido e renomeado para: {destino}")

print("Download concluído. Fechando navegador km...")


navegador.quit()

# --- Selenium Supera ---
navegador = webdriver.Chrome()
navegador.maximize_window()
navegador.get('https://supera.brudam.com.br/index.php?acao=expirada')

# Login
navegador.find_element(By.ID, "user").send_keys("WAGNER.LEANDRO")
navegador.find_element(By.ID, "password").send_keys("Ag@090605")
navegador.find_element(By.ID, "acessar").click()
time.sleep(3)

# Clique no botão da tela seguinte
navegador.find_element(By.CLASS_NAME, "texto_botao").click()

# Preenche código da opção e pressiona Enter
campo = navegador.find_element(By.ID, "codigo_opcao")
campo.send_keys("106")
campo.send_keys(Keys.ENTER)

# Preenche Data Início
data_inicio = navegador.find_element(By.CLASS_NAME, "brd-input-submit-required")
data_inicio.send_keys(data_inicio_selecionada.strftime('%d/%m/%Y'))
time.sleep(3)

# Preenche Data Final
data_final = navegador.find_element(
    By.XPATH,
    '/html/body/center/table/tbody/tr[3]/td/table/tbody/tr[2]/td[1]/table/tbody/tr/td/table/tbody/tr/td/form/table/tbody/tr[7]/td[2]/input'
)
data_final.send_keys(data_final_selecionada.strftime('%d/%m/%Y'))

# Clica no botão de buscar/confirmar
navegador.find_element(By.CLASS_NAME, "brd-button").click()
time.sleep(15)

# Clica em outro botão (corrigido para XPath)
navegador.find_element(
    By.XPATH,
    '/html/body/center/table/tbody/tr[3]/td/table/tbody/tr[2]/td[2]/table/tbody/tr[4]/td[2]/button'
).click()

time.sleep(3)

navegador.find_element(
    By.XPATH,
    '/html/body/div[6]/div[3]/div/button[1]'
).click()
time.sleep(2)
navegador.find_element(
    By.XPATH,
    '/html/body/div[7]/div[2]/table/tbody/tr[1]/td[1]/label/input'
).click()

time.sleep(2)
navegador.find_element(
    By.XPATH,
    '/html/body/div[7]/div[3]/div/button[2]'
).click()

# --- Esperar até o arquivo Excel ser baixado ---
download_path = os.path.join(os.path.expanduser("~"), "Downloads")
arquivo_baixado = False

print("Aguardando download do arquivo...")

while not arquivo_baixado:
    for arquivo in os.listdir(download_path):
        if arquivo.startswith("supera_pers_") and arquivo.endswith(".xlsx"):
            print(f"Arquivo encontrado: {arquivo}")
            arquivo_baixado = True
            break
    time.sleep(1)  # espera 1 segundo antes de checar novamente


# --- Pasta de destino ---
destino_path = fr"{nome_pc}\Downloads\Supera dados CD-ES"
os.makedirs(destino_path, exist_ok=True)

# --- Limpar primeira pasta ---
for f in os.listdir(destino_path):
    caminho = os.path.join(destino_path, f)
    if os.path.isfile(caminho):
        os.remove(caminho)
        print(f"Arquivo deletado: {caminho}")

time.sleep(2)

# --- Limpar também a segunda pasta ---
destino_path2 = fr"{nome_pc}\Downloads\DATA HORA NF'S"
for f in os.listdir(destino_path2):
    caminho = os.path.join(destino_path2, f)
    if os.path.isfile(caminho):
        os.remove(caminho)
        print(f"Arquivo deletado: {caminho}")


time.sleep(2)

# --- Mover arquivo baixado ---
for arquivo in os.listdir(download_path):
    if arquivo.startswith("supera_pers_") and arquivo.endswith(".xlsx"):
        origem = os.path.join(download_path, arquivo)
        destino = os.path.join(destino_path, arquivo)
        shutil.move(origem, destino)
        print(f"Arquivo movido para: {destino}")
        break



print("Download concluído. Fechando navegador...")


navegador.quit()

pythoncom.CoInitialize()

# Configurações
SAP_USER = "Svenancio"
SAP_PASS = "Sm@708090"
SAP_SYSTEM = "0.1. SAP GMILL PRD 2022"
SAP_LOGON_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"

pythoncom.CoInitialize()

SAP_USER = "Svenancio"
SAP_PASS = "Sm@708090"
SAP_SYSTEM = "0.1. SAP GMILL PRD 2022"
SAP_LOGON_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"

session = None

# -------------------------------
# Tenta conectar a uma sessão SAP já aberta
# -------------------------------
try:
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    print("SAP já estava aberto, fechando...")

    # -------------------------------
    # Fecha SAP usando os comandos que você forneceu
    # -------------------------------
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n/"
    session.findById("wnd[0]/tbar[0]/btn[15]").press()            # Botão Fechar
    time.sleep(2)
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()       # Confirma fechar
    print("SAP fechado com sucesso!")
    time.sleep(2)
except Exception:
    # SAP não estava aberto
    print("SAP não estava aberto.")

# -------------------------------
# Abre SAP Logon e conecta
# -------------------------------
os.startfile(SAP_LOGON_PATH)
time.sleep(5)  # espera abrir SAP Logon

SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.OpenConnection(SAP_SYSTEM, True)
session = connection.Children(0)
time.sleep(3)

# -------------------------------
# Faz login
# -------------------------------
session.findById("wnd[0]/usr/txtRSYST-BNAME").text = sap_user
session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = sap_password
session.findById("wnd[0]").sendVKey(0)  # Enter

print("Login realizado com sucesso! SAP pronto para transações.")

time.sleep(2)


# --- Automação SAP ---

session.findById("wnd[0]").maximize()
session.findById("wnd[0]/tbar[0]/okcd").text = "/n/"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]").maximize()
session.findById("wnd[0]/tbar[0]/okcd").text = "J1BNFE"
session.findById("wnd[0]").sendVKey(0)

session.findById("wnd[0]/tbar[1]/btn[17]").press()
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").setCurrentCell(8, "TEXT")
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "8"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()

session.findById("wnd[0]/usr/ctxtDATE0-LOW").text = f"{semana_mes_passado_str}"
session.findById("wnd[0]/usr/ctxtDATE0-HIGH").text = f"{data_ontem_str}"
session.findById("wnd[0]/usr/ctxtDATE0-HIGH").setFocus()
session.findById("wnd[0]/usr/ctxtDATE0-HIGH").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[8]").press()

session.findById("wnd[0]/usr/cntlNFE_CONTAINER/shellcont/shell").pressToolbarButton("&MB_VARIANT")
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").setCurrentCell(2, "TEXT")
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = "2"
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()

session.findById("wnd[0]/usr/cntlNFE_CONTAINER/shellcont/shell").setCurrentCell(4, "ACTION_TIME")
session.findById("wnd[0]/usr/cntlNFE_CONTAINER/shellcont/shell").contextMenu()
session.findById("wnd[0]/usr/cntlNFE_CONTAINER/shellcont/shell").selectContextMenuItem("&XXL")

session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/usr/ctxtDY_PATH").text = fr"{nome_pc}\Downloads\DATA HORA NF'S"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Notas ficais.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 11
session.findById("wnd[1]/tbar[0]/btn[11]").press()


### Dados de tarefas ###

session.findById("wnd[0]").maximize()

# --- ZSWMS_CHECKOUT_LOG 2005 ---
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n/"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/okcd").Text = "ZSWMS_CHECKOUT_LOG"
session.findById("wnd[0]").sendVKey(0)

session.findById("wnd[0]/usr/ctxtS_WHOUSE-LOW").Text = "2005"
session.findById("wnd[0]/usr/ctxtS_DATE-LOW").text = f"{inicio_periodo_tarefas}"
session.findById("wnd[0]/usr/ctxtS_DATE-HIGH").text = f"{fim_periodo_tarefas}"
session.findById("wnd[0]/usr/ctxtS_DATE-LOW").SetFocus()
session.findById("wnd[0]/usr/ctxtS_DATE-LOW").caretPosition = 10
session.findById("wnd[0]/tbar[1]/btn[8]").press()






# Selecionar Picking HU
session.findById("wnd[0]/usr/shell").setCurrentCell(13, "PICKING_HU")
session.findById("wnd[0]/usr/shell").contextMenu()
session.findById("wnd[0]/usr/shell").selectContextMenuItem("&XXL")
session.findById("wnd[1]/tbar[0]/btn[0]").press()

# Salvar arquivo
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = fr"{nome_pc}\Labs TI\Dados Gmill - Planejamento\Planejamento - Geral\SAMUEL.VENANCIO\PRODUTIVIDADE\DADOS\DADOS Check CD-RJ"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = f"{inicio_periodo_tarefas} - CHECK-OUT-RJ.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
session.findById("wnd[1]/tbar[0]/btn[11]").press()

# --- /scwm/mon 2005 ---
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n/scwm/mon"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/tbar[1]/btn[5]").press()

session.findById("wnd[1]/usr/ctxtP_LGNUM").Text = "2005"
session.findById("wnd[1]/usr/ctxtP_LGNUM").caretPosition = 4
session.findById("wnd[1]").sendVKey(0)

session.findById("wnd[1]/usr/ctxtP_MONIT").Text = "sap"
session.findById("wnd[1]/usr/ctxtP_MONIT").SetFocus()
session.findById("wnd[1]/usr/ctxtP_MONIT").caretPosition = 3
session.findById("wnd[1]/tbar[0]/btn[8]").press()

# Selecionar node e abrir detalhes
session.findById("wnd[0]/usr/shell/shellcont[0]/shell").selectedNode = "N000000034"
session.findById("wnd[0]/usr/shell/shellcont[0]/shell").doubleClickNode("N000000034")

session.findById("wnd[1]/tbar[0]/btn[17]").press()
session.findById("wnd[2]/usr/cntlALV_CONTAINER_1/shellcont/shell").setCurrentCell(3, "TEXT")
session.findById("wnd[2]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "3"
session.findById("wnd[2]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()

session.findById("wnd[1]/usr/ctxtP_CODFR").Text = f"{inicio_periodo_tarefas}"
session.findById("wnd[1]/usr/ctxtP_COTFR").Text = "00:00:00"
session.findById("wnd[1]/usr/ctxtP_CODTO").Text = f"{fim_periodo_tarefas}"
session.findById("wnd[1]/usr/ctxtP_COTTO").caretPosition = 6
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[1]/tbar[0]/btn[8]").press()

# Exportar DADOS-RJ
session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").setCurrentCell(5, "ACT_TYPE")
session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").contextMenu()
session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").selectContextMenuItem("&XXL")
session.findById("wnd[1]/tbar[0]/btn[0]").press()

session.findById("wnd[1]/usr/ctxtDY_PATH").Text = fr"{nome_pc}\Labs TI\Dados Gmill - Planejamento\Planejamento - Geral\SAMUEL.VENANCIO\PRODUTIVIDADE\DADOS\Dados CD-RJ"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = f"{inicio_periodo_tarefas} - DADOS-RJ.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 11
session.findById("wnd[1]/tbar[0]/btn[11]").press()

session.findById("wnd[0]").maximize()
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n/"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]").maximize()
session.findById("wnd[0]/tbar[0]/okcd").Text = "ZSWMS_CHECKOUT_LOG"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/ctxtS_WHOUSE-LOW").Text = "2001"
session.findById("wnd[0]/usr/ctxtS_DATE-LOW").text = f"{inicio_periodo_tarefas}"
session.findById("wnd[0]/usr/ctxtS_DATE-HIGH").text = f"{fim_periodo_tarefas}"
session.findById("wnd[0]/usr/ctxtS_DATE-LOW").SetFocus()
session.findById("wnd[0]/usr/ctxtS_DATE-LOW").caretPosition = 10
session.findById("wnd[0]/tbar[1]/btn[8]").press()
session.findById("wnd[0]/usr/shell").setCurrentCell(13, "PICKING_HU")
session.findById("wnd[0]/usr/shell").contextMenu()
session.findById("wnd[0]/usr/shell").selectContextMenuItem("&XXL")
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = fr"{nome_pc}\Labs TI\Dados Gmill - Planejamento\Planejamento - Geral\SAMUEL.VENANCIO\PRODUTIVIDADE\DADOS\DADOS Check CD-ES"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = f"{inicio_periodo_tarefas}  - CHECK-OUT-ES.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
session.findById("wnd[1]/tbar[0]/btn[11]").press()

session.findById("wnd[0]").maximize()
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n/scwm/mon"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/tbar[1]/btn[5]").press()
session.findById("wnd[1]/usr/ctxtP_LGNUM").Text = "2001"
session.findById("wnd[1]/usr/ctxtP_LGNUM").caretPosition = 4
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[1]/usr/ctxtP_MONIT").Text = "sap"
session.findById("wnd[1]/usr/ctxtP_MONIT").SetFocus()
session.findById("wnd[1]/usr/ctxtP_MONIT").caretPosition = 3
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]/usr/shell/shellcont[0]/shell").selectedNode = "N000000034"
session.findById("wnd[0]/usr/shell/shellcont[0]/shell").doubleClickNode("N000000034")
session.findById("wnd[1]/tbar[0]/btn[17]").press()
session.findById("wnd[2]/usr/cntlALV_CONTAINER_1/shellcont/shell").setCurrentCell(3,"TEXT")
session.findById("wnd[2]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "3"
session.findById("wnd[2]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
session.findById("wnd[1]/usr/ctxtP_CODFR").Text = f"{inicio_periodo_tarefas}"
session.findById("wnd[1]/usr/ctxtP_COTFR").Text = "00:00:00"
session.findById("wnd[1]/usr/ctxtP_CODTO").Text = f"{fim_periodo_tarefas}"
session.findById("wnd[1]/usr/ctxtP_CTDFR").SetFocus()
session.findById("wnd[1]/usr/ctxtP_CTDFR").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").setCurrentCell(5, "ACT_TYPE")
session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").contextMenu()
session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").selectContextMenuItem("&XXL")
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = fr"{nome_pc}\Labs TI\Dados Gmill - Planejamento\Planejamento - Geral\SAMUEL.VENANCIO\PRODUTIVIDADE\DADOS\\Dados CD-ES"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = f"{inicio_periodo_tarefas} - DADOS-ES.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 11
session.findById("wnd[1]/tbar[0]/btn[11]").press()

session.findById("wnd[0]").maximize()
session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").setCurrentCell(-1,"")
session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").selectAll()
session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarContextButton("&MB_VARIANT")
session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").selectContextMenuItem("&LOAD")
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellRow = 31
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = "31"
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").setCurrentCell(4,"CONFIRMED_DATE")
session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").contextMenu()
session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").selectContextMenuItem("&XXL")
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = fr"{nome_pc}\Labs TI\Dados Gmill - Planejamento\Planejamento - Geral\SAMUEL.VENANCIO\PRODUTIVIDADE\DADOS\DADOS MKT-ES"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = f"{inicio_periodo_tarefas} - DADOS-MKT.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[11]").press()





session.findById("wnd[0]/tbar[0]/btn[15]").press()
session.findById("wnd[0]/tbar[0]/btn[15]").press()  # Botão Fechar
time.sleep(1)
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()  # Confirma fechar

time.sleep(6)



# Conectar ao Excel

print("Fechando todas as planilhas do Excel...")
ret = os.system("taskkill /f /im excel.exe >nul 2>&1")

if ret == 0:
    print("Todas as planilhas do Excel foram fechadas.")
else:
    print("Nenhuma planilha do Excel estava aberta.")


time.sleep(6)

# Caminhos das pastas
Nfspasta = fr"{nome_pc}\Downloads\DATA HORA NF'S"
Entregapasta = fr"{nome_pc}\Downloads\Supera dados CD-ES"



print("Mesclando planilhas.")
# Lista arquivos .xlsx
arquivosNfs = glob.glob(os.path.join(Nfspasta, "*.xlsx"))
arquivosEntregas = glob.glob(os.path.join(Entregapasta, "*.xlsx"))

# Lê e concatena os arquivos de NF
dfNfs = pd.concat([pd.read_excel(arq) for arq in arquivosNfs], ignore_index=True)
dfNfs = dfNfs[['Número de nota fiscal eletrônica', 'Data de modificação', 'Hora da modificação']]

# Lê e concatena os arquivos de entregas
dfEntregas = pd.concat([pd.read_excel(arq) for arq in arquivosEntregas], ignore_index=True)

# Remove linhas que parecem notação científica (ex.: 1.23E+11)
dfEntregas = dfEntregas[~dfEntregas['NF/DOC'].str.contains(r'E\+', na=False, case=False)]

# Remove tudo que não for número
dfNfs['Número de nota fiscal eletrônica'] = dfNfs['Número de nota fiscal eletrônica'].astype(str).str.replace(r'[^0-9]', '', regex=True)
dfEntregas['NF/DOC'] = dfEntregas['NF/DOC'].astype(str).str.replace(r'[^0-9]', '', regex=True)

# 🔹 Limita a 9 dígitos e adiciona zeros à esquerda
dfNfs = dfNfs[dfNfs['Número de nota fiscal eletrônica'].str.len() <= 9]
dfEntregas = dfEntregas[dfEntregas['NF/DOC'].str.len() <= 9]

dfNfs['Número de nota fiscal eletrônica'] = dfNfs['Número de nota fiscal eletrônica'].str.zfill(9)
dfEntregas['NF/DOC'] = dfEntregas['NF/DOC'].str.zfill(9)

# 🔹 Mantém como texto para evitar problemas no Excel
# Faz o merge
dfFinal = pd.merge(
    dfEntregas,
    dfNfs[['Número de nota fiscal eletrônica', 'Data de modificação', 'Hora da modificação']],
    left_on='NF/DOC',
    right_on='Número de nota fiscal eletrônica',
    how='left'
)

# Remove a coluna duplicada
dfFinal = dfFinal.drop(columns=['Número de nota fiscal eletrônica'])

# Salva em Excel
caminho_saida = fr"{nome_pc}\Labs TI\Dados Gmill - Planejamento\Planejamento - Controle de Indicadores\25 - Transportadoras SLA\Supera log\Supera dados CD-ES\2025 - {mes_atual}.xlsx"
dfFinal.to_excel(caminho_saida, index=False)

root_final = tk.Tk()
root_final.withdraw()  # Oculta a janela principal
messagebox.showinfo("Concluído", "Script finalizado!.")
root_final.destroy()



print("Arquivo salvo")
