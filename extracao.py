import os
import time
import glob
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import win32com.client as win32
from dotenv import load_dotenv

# Carrega variáveis do arquivo .env
load_dotenv()

# === CONFIGURAÇÕES ===
URL = os.getenv("SEUFISIO_URL", "https://app.seufisio.com.br")
EMAIL = os.getenv("SEUFISIO_EMAIL")
SENHA = os.getenv("SEUFISIO_SENHA")

PASTA_DOWNLOAD = r"C:\Users\rafae\OneDrive\Área de Trabalho\Pilates_25\Dados_seu_fisio"
NOME_ARQUIVO_XLS = "atendimentos-gerais.xls"
NOME_FINAL_BASE = "atendimentos-gerais"

CREDENCIALS_PATH = os.getenv("GOOGLE_CREDENTIALS")
GOOGLE_SHEET_NAME = os.getenv("GOOGLE_SHEET_NAME")
GOOGLE_SHEET_ABA = os.getenv("GOOGLE_SHEET_ABA")

# Data de hoje no formato dd/mm/aaaa
ontem = (datetime.today() - timedelta(days=1)).strftime("%d/%m/%Y")
hoje = datetime.today().strftime("%d/%m/%Y")
DATA_INICIAL = ontem
DATA_FINAL = ontem

# Cria a pasta de download se não existir
os.makedirs(PASTA_DOWNLOAD, exist_ok=True)

# Configurações do navegador
options = Options()
options.add_experimental_option("prefs", {
    "download.default_directory": PASTA_DOWNLOAD,
    "download.prompt_for_download": False,
    "safebrowsing.enabled": True
})
driver = webdriver.Chrome(options=options)

# === INÍCIO DA AUTOMATIZAÇÃO ===
driver.get(URL)

# LOGIN
email_input = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="tela_login"]/div/div[2]/form/div[1]//input'))
)
email_input.send_keys(EMAIL)

senha_input = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="tela_login"]/div/div[2]/form/div[2]//input'))
)
senha_input.send_keys(SENHA)

botao_login = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="tela_login"]//button[@type="submit"]'))
)
botao_login.click()


# ESPERA ENTRAR
WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="q-app"]/div/div[1]/header/div[1]/div/div[1]/button[7]/span')))
driver.find_element(By.XPATH, '//*[@id="q-app"]/div/div[1]/header/div[1]/div/div[1]/button[7]/span').click()

# ATENDIMENTOS GERAIS
WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((
        By.XPATH, "//a[contains(@href, '/relatorio/atendimento')]"
    ))
).click()

# Preencher Data Inicial
data_inicial = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'input[aria-label="Data atendimento - inicial"]'))
)
data_inicial.click()
data_inicial.send_keys(Keys.CONTROL + "a")
data_inicial.send_keys(Keys.BACKSPACE)
data_inicial.send_keys(DATA_INICIAL)

# Preencher Data Final
data_final = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'input[aria-label="Data atendimento - final"]'))
)
data_final.click()
data_final.send_keys(Keys.CONTROL + "a")
data_final.send_keys(Keys.BACKSPACE)
data_final.send_keys(DATA_FINAL)

# CLICA EM FILTRAR
botao_filtrar = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(),'Filtrar')]]"))
)
driver.execute_script("arguments[0].scrollIntoView(true);", botao_filtrar)  # Garante que esteja visível
botao_filtrar.click()

# CLICA EM EXPORTAR XLS
botao_exportar_xls = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(),'Exportar XLS')]]"))
)
driver.execute_script("arguments[0].scrollIntoView(true);", botao_exportar_xls)  # Garante que esteja visível
botao_exportar_xls.click()

print("✅ Exportação realizada com sucesso!")

# === FUNÇÃO PARA ESPERAR E RENOMEAR ===
def esperar_e_renomear_arquivo(pasta, nome_inicial, nome_final_base, timeout=60):
    caminho_inicial = os.path.join(pasta, nome_inicial)
    data_ontem = (datetime.today() - timedelta(days=1)).strftime("%Y-%m-%d")
    nome_final = f"{nome_final_base}_{data_ontem}.xls"
    caminho_final = os.path.join(pasta, nome_final)

    tempo_inicial = time.time()

    while True:
        arquivo_existe = os.path.exists(caminho_inicial)
        arquivos_tmp = glob.glob(os.path.join(pasta, '*.tmp'))

        if arquivo_existe and not arquivos_tmp:
            os.rename(caminho_inicial, caminho_final)
            print(f"✔️ Arquivo renomeado para {nome_final}")
            return caminho_final

        if time.time() - tempo_inicial > timeout:
            raise TimeoutError(f"⏰ Timeout: Arquivo {nome_inicial} não encontrado na pasta dentro do tempo limite.")

        time.sleep(1)

# === AGUARDA E RENOMEIA O ARQUIVO ===
arquivo_baixado = esperar_e_renomear_arquivo(
    pasta=PASTA_DOWNLOAD,
    nome_inicial=NOME_ARQUIVO_XLS,
    nome_final_base=NOME_FINAL_BASE,
    timeout=60
)

def converter_xls_para_xlsx(caminho_arquivo):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    caminho_absoluto = os.path.abspath(caminho_arquivo)
    pasta, arquivo = os.path.split(caminho_absoluto)
    nome_sem_extensao = os.path.splitext(arquivo)[0]
    novo_arquivo = os.path.join(pasta, f"{nome_sem_extensao}.xlsx")

    wb = excel.Workbooks.Open(caminho_absoluto)
    wb.SaveAs(novo_arquivo, FileFormat=51)  # 51 = formato xlsx
    wb.Close()
    excel.Application.Quit()

    # Remove o arquivo .xls original
    if os.path.exists(caminho_absoluto):
        os.remove(caminho_absoluto)
        print(f"🗑️ Arquivo .xls removido: {caminho_absoluto}")

    print(f"✔️ Arquivo convertido para: {novo_arquivo}")
    return novo_arquivo

arquivo_convertido = converter_xls_para_xlsx(arquivo_baixado)

# === TRATAMENTO DE DADOS ===

# === Leitura do arquivo sem cabeçalho e remoção das primeiras 4 linhas ===
df = pd.read_excel(arquivo_convertido, header=None)
df = df.iloc[4:].reset_index(drop=True)

# === Usa a primeira linha após corte como cabeçalho ===
df.columns = df.iloc[0]
df = df[1:].reset_index(drop=True)

# === Excluir colunas pelos índices: F (5) e B (1) ===
df = df.drop(df.columns[[5, 1]], axis=1)

# Separação da coluna "Data/Hora" em "Data" e "Hora"
df[['Data', 'Hora']] = df['Data/Hora'].str.split(' ', expand=True)
df = df.drop(columns=["Data/Hora"])

# Reorganiza as colunas (opcional)
colunas_ordenadas = ['Cliente', 'Data', 'Hora', 'Tipo Atendimento', 'Status']
df = df[colunas_ordenadas]

# Remove duplicatas por Cliente + Data + Hora
df = df.drop_duplicates(subset=["Cliente", "Data", "Hora"])

# === Ordena pelos valores de 'Data' e depois 'Hora' ===
df = df.sort_values(by=['Data', 'Hora']).reset_index(drop=True)

# Salva no mesmo arquivo XLSX (opcional, pode não salvar se quiser ir direto ao Google Sheets)
df.to_excel(arquivo_convertido, index=False)

print(f"✅ Arquivo transformado e salvo com sucesso em: {arquivo_convertido}")

# === ENVIO AO GOOGLE SHEETS ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENCIALS_PATH, scope)
client = gspread.authorize(creds)

sheet = client.open(GOOGLE_SHEET_NAME).worksheet(GOOGLE_SHEET_ABA)

# Verifica se a planilha está vazia e adiciona o cabeçalho se necessário
if len(sheet.get_all_values()) == 0:
    sheet.append_row(df.columns.values.tolist(), value_input_option="USER_ENTERED")

# Acrescenta os dados ao final da planilha
sheet.append_rows(df.values.tolist(), value_input_option="USER_ENTERED")

print("✅ Processo finalizado com sucesso!")
