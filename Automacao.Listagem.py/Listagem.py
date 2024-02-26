from selenium import webdriver

import pyautogui

import time

from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.chrome.service import Service

from selenium.webdriver.chrome.options import Options

from selenium.webdriver.common.by import By

from selenium.webdriver.common.keys import Keys

from selenium.webdriver.common.action_chains import ActionChains

from selenium.common.exceptions import NoSuchElementException

def seleciona_empresa(xpath):
    empresa = driver.find_element(By.XPATH, '//*[@id="DivExperiencesContainer"]/div[1]/div[2]/div/div[2]/div')
    driver.execute_script("window.getSelection().selectAllChildren(arguments[0]);", empresa)
    pyautogui.hotkey("ctrl", "c")
    pyautogui.hotkey("alt", "tab")
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("tab")
    pyautogui.hotkey("alt", "tab")

    time.sleep(4)

def seleciona_profissao(xpath):
    profissao = driver.find_element(By.XPATH, '//*[@id="DivExperiencesContainer"]/div[1]/div[2]/div/div[1]')
    driver.execute_script("window.getSelection().selectAllChildren(arguments[0])", profissao)
    pyautogui.hotkey("ctrl", "c")
    pyautogui.hotkey("alt", "tab")
    time.sleep(2)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("tab")
    pyautogui.hotkey("alt", "tab")

    time.sleep(4)

def seleciona_Data(xpath):
    data = driver.find_element(By.XPATH, '//*[@id="DivExperiencesContainer"]/div[1]/div[1]/div/div[1]')
    driver.execute_script("window.getSelection().selectAllChildren(arguments[0])", data)
    pyautogui.hotkey("ctrl", "c")
    pyautogui.hotkey("alt", "tab")
    time.sleep(2)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("tab")
    pyautogui.hotkey("alt", "tab")

    time.sleep(4)

def seleciona_NumeroCell(xpath):
    try:
        numeroCell = driver.find_element(By.XPATH, '//*[@id="HeaderInfoContainer"]/div[6]/div[1]/a')
    except NoSuchElementException:
        try:
            numeroCell = driver.find_element(By.XPATH, '//*[@id="HeaderInfoContainer"]/div[6]/div[2]/a')    
        except:
            print("Nenhum dos números foi encontrado")
    driver.execute_script("window.getSelection().selectAllChildren(arguments[0])", numeroCell)
    pyautogui.hotkey("ctrl", "c")
    pyautogui.hotkey("alt", "tab")
    time.sleep(2)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter")
    pyautogui.press("enter")

time.sleep(5) # Tempo de carregamento

chrome_options = Options()

chrome_options.add_experimental_option("detach", True)

servico = Service(ChromeDriverManager().install())

driver = webdriver.Chrome(service=servico, options=chrome_options)

# Abre a página da web
driver.get("https://login.pandape.com/Account/Login?ReturnUrl=%2FSelect%3FreturnUrl%3Dhttps%253a%252f%252fpandape.infojobs.com.br%252fCompany%252fDashboard")
driver.maximize_window()

# Tempo de espera para o carregamento da página
time.sleep(5)

# Insere o e-mail
email_input = driver.find_element(By.ID, 'Username') # Localizador por ID
email_input.send_keys("abertura@contabilizei.com.br") # Input dos dados cadastrais

# Insere a senha
password_input = driver.find_element(By.ID, 'Password')
password_input.send_keys("123456")

# Pressiona o btn de enter
password_input.send_keys(Keys.ENTER)

# Tempo de espera para o carregamento da página
time.sleep(10)

# Clica no local relacionado às vagas
vagas_input = driver.find_element(By.XPATH, '//*[@id="LnkNavVacancy"]')
vagas_input.click()

# Clica no local de candidatos totais (Auxiliar de Produção)
pressiona_btn = driver.find_element(By.XPATH, '//*[@id="rowVacancy_880398"]/div[2]/div/div[2]/div[2]/div[1]/a/div').click() # Localizador por XPATH

# Primeiro nome + Input na Planilha
time.sleep(5)
primeiro_Nome = driver.find_element(By.XPATH, '//*[@id="HeaderInfoContainer"]/div[1]/div')
driver.execute_script("window.getSelection().selectAllChildren(arguments[0])", primeiro_Nome) # Seleciona o Termo orientado a XPATH escolhido
pyautogui.hotkey('ctrl', 'c')
time.sleep(2)
pyautogui.click(x=989, y=744) # Click na aba da planilha
pyautogui.press('space')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

pyautogui.click(x=1059, y=746)  

time.sleep(4) # Tempo de espera para o carregamento 

try:
    seleciona_empresa(xpath=1)
except NoSuchElementException:
    print("(1) Empresa não encontrada. Continuando com a operação")

try:
    seleciona_profissao(xpath=1)
except NoSuchElementException:
    print("(1) Profissão não encontrada. Continuando com a operação")

try:
    seleciona_Data(xpath=1)
except NoSuchElementException:
    print("(1) Data não encontrada. Continuando com a operação")

try:
    seleciona_NumeroCell(xpath=1)
except NoSuchElementException:
    pyautogui.press('enter')
    pyautogui.press('enter')    
    print("(1) Número de celular não encontrado. Continuando com a operação")
    print("")
    
# - - - - - 2
# Seleciona o campo para seguir para a próxima pessoa
pyautogui.hotkey('alt','tab')
segunda_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[2]/div/div[2]/a').click()

time.sleep(10) # Tempo de carregamento da página 

# Nome da Segunda Pessoa + Input na Planilha 
segundo_Nome = driver.find_element(By.XPATH, '//*[@id="HeaderInfoContainer"]/div[1]/div')
driver.execute_script("window.getSelection().selectAllChildren(arguments[0])", segundo_Nome) # Seleciona o Termo orientado a XPATH escolhido
time.sleep(2) # Tempo de espera - "Talvez retirar daqui"
pyautogui.hotkey('ctrl', 'c')
pyautogui.hotkey('alt', 'tab')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

pyautogui.hotkey('alt','tab')

time.sleep(4)

try:
    seleciona_empresa(xpath=2)
except NoSuchElementException:
    print("(2) Empresa não encontrada. Continuando com a operação")

try:
    seleciona_profissao(xpath=2)
except NoSuchElementException:
    print("(2) Profissão não encontrada. Continuando com a operação")

try:
    seleciona_Data(xpath=2)
except NoSuchElementException:
    print("(2) Data não encontrada. Continuando com a operação")

try:
    seleciona_NumeroCell(xpath=2)
except NoSuchElementException:
    pyautogui.press('enter')
    pyautogui.press('enter')
    print("(2) Número de celular não encontrado. Continuando com a operação")
    print("")
    
# - - - - - 3
pyautogui.hotkey('alt','tab')
terceira_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[3]/div/div[2]/a').click()

time.sleep(10) # Tempo de carregamento da página 

# Nome da Segunda Pessoa + Input na Planilha 
segundo_Nome = driver.find_element(By.XPATH, '//*[@id="HeaderInfoContainer"]/div[1]/div')
driver.execute_script("window.getSelection().selectAllChildren(arguments[0])", segundo_Nome) # Seleciona o Termo orientado a XPATH escolhido
time.sleep(2) # Tempo de espera - "Talvez retirar daqui"
pyautogui.hotkey('ctrl', 'c')
pyautogui.hotkey('alt', 'tab')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

pyautogui.hotkey('alt','tab')

time.sleep(4)

try:
    seleciona_empresa(xpath=3)
except NoSuchElementException:
    print("(3) Empresa não encontrada. Continuando com a operação")

try:
    seleciona_profissao(xpath=3)
except NoSuchElementException:
    print("(3) Profissão não encontrada. Continuando com a operação")

try:
    seleciona_Data(xpath=3)
except NoSuchElementException:
    print("(3) Data não encontrada. Continuando com a operação")

try:
    seleciona_NumeroCell(xpath=3)
except NoSuchElementException:
    pyautogui.press('enter')
    pyautogui.press('enter')
    print("(3) Número de celular não encontrado. Continuando com a operação")
    print("")

# - - - - - 4
pyautogui.hotkey('alt','tab')
quarta_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[4]/div/div[2]/a').click()

time.sleep(10) # Tempo de carregamento da página 

# Nome da Segunda Pessoa + Input na Planilha 
segundo_Nome = driver.find_element(By.XPATH, '//*[@id="HeaderInfoContainer"]/div[1]/div')
driver.execute_script("window.getSelection().selectAllChildren(arguments[0])", segundo_Nome) # Seleciona o Termo orientado a XPATH escolhido
time.sleep(2) # Tempo de espera - "Talvez retirar daqui"
pyautogui.hotkey('ctrl', 'c')
pyautogui.hotkey('alt', 'tab')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

pyautogui.hotkey('alt','tab')

time.sleep(4)

try:
    seleciona_empresa(xpath=4)
except NoSuchElementException:
    print("(4) Empresa não encontrada. Continuando com a operação")

try:
    seleciona_profissao(xpath=4)
except NoSuchElementException:
    print("(4) Profissão não encontrada. Continuando com a operação")

try:
    seleciona_Data(xpath=4)
except NoSuchElementException:
    print("(4) Data não encontrada. Continuando com a operação")

try:
    seleciona_NumeroCell(xpath=4)
except NoSuchElementException:
    pyautogui.press('enter')
    pyautogui.press('enter')
    print("(4) Número de celular não encontrado. Continuando com a operação")
    print("")

# - - - - - 5
pyautogui.hotkey('alt','tab')
quinta_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[5]/div/div[2]/a').click()

time.sleep(10) # Tempo de carregamento da página 

# Nome da Segunda Pessoa + Input na Planilha 
segundo_Nome = driver.find_element(By.XPATH, '//*[@id="HeaderInfoContainer"]/div[1]/div')
driver.execute_script("window.getSelection().selectAllChildren(arguments[0])", segundo_Nome) # Seleciona o Termo orientado a XPATH escolhido
time.sleep(2) # Tempo de espera - "Talvez retirar daqui"
pyautogui.hotkey('ctrl', 'c')
pyautogui.hotkey('alt', 'tab')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

pyautogui.hotkey('alt','tab')

time.sleep(4)

try:
    seleciona_empresa(xpath=5)
except NoSuchElementException:
    print("(5) Empresa não encontrada. Continuando com a operação")

try:
    seleciona_profissao(xpath=5)
except NoSuchElementException:
    print("(5) Profissão não encontrada. Continuando com a operação")

try:
    seleciona_Data(xpath=5)
except NoSuchElementException:
    print("(5) Data não encontrada. Continuando com a operação")

try:
    seleciona_NumeroCell(xpath=5)
except NoSuchElementException:
    pyautogui.press('enter')
    pyautogui.press('enter')
    print("(5) Número de celular não encontrado. Continuando com a operação")
    print("")

# - - - - - 6
pyautogui.hotkey('alt','tab')
sexta_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[6]/div/div[2]/a').click()

time.sleep(10) # Tempo de carregamento da página

# Nome da Segunda Pessoa + Input na Planilha 
segundo_Nome = driver.find_element(By.XPATH, '//*[@id="HeaderInfoContainer"]/div[1]/div')
driver.execute_script("window.getSelection().selectAllChildren(arguments[0])", segundo_Nome) # Seleciona o Termo orientado a XPATH escolhido
time.sleep(2) # Tempo de espera - "Talvez retirar daqui"
pyautogui.hotkey('ctrl', 'c')
pyautogui.hotkey('alt', 'tab')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

pyautogui.hotkey('alt','tab')

time.sleep(4)

try:
    seleciona_empresa(xpath=6)
except NoSuchElementException:
    print("(6) Empresa não encontrada. Continuando com a operação")

try:
    seleciona_profissao(xpath=6)
except NoSuchElementException:
    print("(6) Profissão não encontrada. Continuando com a operação")

try:
    seleciona_Data(xpath=6)
except NoSuchElementException:
    print("(6) Data não encontrada. Continuando com a operação")

try:
    seleciona_NumeroCell(xpath=6)
except NoSuchElementException:
    pyautogui.press('enter')
    pyautogui.press('enter')
    print("(6) Número de celular não encontrado. Continuando com a operação")
    print("")
    
# - - - - - 7
pyautogui.hotkey('alt','tab')
setima_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[7]/div/div[2]/a').click()

time.sleep(10) # Tempo de carregamento da página

# Nome da Segunda Pessoa + Input na Planilha 
segundo_Nome = driver.find_element(By.XPATH, '//*[@id="HeaderInfoContainer"]/div[1]/div')
driver.execute_script("window.getSelection().selectAllChildren(arguments[0])", segundo_Nome) # Seleciona o Termo orientado a XPATH escolhido
time.sleep(2) # Tempo de espera - "Talvez retirar daqui"
pyautogui.hotkey('ctrl', 'c')
pyautogui.hotkey('alt', 'tab')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

pyautogui.hotkey('alt','tab')

time.sleep(4)

try:
    seleciona_empresa(xpath=7)
except NoSuchElementException:
    print("(7) Empresa não encontrada. Continuando com a operação")

try:
    seleciona_profissao(xpath=7)
except NoSuchElementException:
    print("(7) Profissão não encontrada. Continuando com a operação")

try:
    seleciona_Data(xpath=7)
except NoSuchElementException:
    print("(7) Data não encontrada. Continuando com a operação")

try:
    seleciona_NumeroCell(xpath=7)
except NoSuchElementException:
    pyautogui.press('enter')
    pyautogui.press('enter')
    print("(7) Número de celular não encontrado. Continuando com a operação")
    print("")
    
# - - - - - 8
pyautogui.hotkey('alt','tab')
oitava_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[8]/div/div[2]/a').click()

time.sleep(10) # Tempo de carregamento da página

# Nome da Segunda Pessoa + Input na Planilha 
segundo_Nome = driver.find_element(By.XPATH, '//*[@id="HeaderInfoContainer"]/div[1]/div')
driver.execute_script("window.getSelection().selectAllChildren(arguments[0])", segundo_Nome) # Seleciona o Termo orientado a XPATH escolhido
time.sleep(2) # Tempo de espera - "Talvez retirar daqui"
pyautogui.hotkey('ctrl', 'c')
pyautogui.hotkey('alt', 'tab')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

pyautogui.hotkey('alt','tab')

time.sleep(4)

try:
    seleciona_empresa(xpath=8)
except NoSuchElementException:
    print("(8) Empresa não encontrada. Continuando com a operação")

try:
    seleciona_profissao(xpath=8)
except NoSuchElementException:
    print("(8) Profissão não encontrada. Continuando com a operação")

try:
    seleciona_Data(xpath=8)
except NoSuchElementException:
    print("(8) Data não encontrada. Continuando com a operação")

try:
    seleciona_NumeroCell(xpath=8)
except NoSuchElementException:
    pyautogui.press('enter')
    pyautogui.press('enter')
    print("(8) Número de celular não encontrado. Continuando com a operação")
    print("")

# - - - - - 9
pyautogui.hotkey('alt','tab')
nona_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[9]/div/div[2]/a').click()

time.sleep(10) # Tempo de carregamento da página

# Nome da Segunda Pessoa + Input na Planilha 
segundo_Nome = driver.find_element(By.XPATH, '//*[@id="HeaderInfoContainer"]/div[1]/div')
driver.execute_script("window.getSelection().selectAllChildren(arguments[0])", segundo_Nome) # Seleciona o Termo orientado a XPATH escolhido
time.sleep(2) # Tempo de espera - "Talvez retirar daqui"
pyautogui.hotkey('ctrl', 'c')
pyautogui.hotkey('alt', 'tab')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

pyautogui.hotkey('alt','tab')

time.sleep(4)

try:
    seleciona_empresa(xpath=9)
except NoSuchElementException:
    print("(9) Empresa não encontrada. Continuando com a operação")

try:
    seleciona_profissao(xpath=9)
except NoSuchElementException:
    print("(9) Profissão não encontrada. Continuando com a operação")

try:
    seleciona_Data(xpath=9)
except NoSuchElementException:
    print("(9) Data não encontrada. Continuando com a operação")

try:
    seleciona_NumeroCell(xpath=9)
except NoSuchElementException:
    pyautogui.press('enter')
    pyautogui.press('enter')
    print("(9) Número de celular não encontrado. Continuando com a operação")
    print("")

# - - - - - 10
pyautogui.hotkey('alt','tab')
pyautogui.click(x=1362, y=368)
time.sleep(2)
decima_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[10]/div/div[2]/a').click()

time.sleep(10) # Tempo de carregamento da página

# Nome da Segunda Pessoa + Input na Planilha 
segundo_Nome = driver.find_element(By.XPATH, '//*[@id="HeaderInfoContainer"]/div[1]/div')
driver.execute_script("window.getSelection().selectAllChildren(arguments[0])", segundo_Nome) # Seleciona o Termo orientado a XPATH escolhido
time.sleep(2) # Tempo de espera - "Talvez retirar daqui"
pyautogui.hotkey('ctrl', 'c')
pyautogui.hotkey('alt', 'tab')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

pyautogui.hotkey('alt','tab')

time.sleep(4)

try:
    seleciona_empresa(xpath=10)
except NoSuchElementException:
    print("(10) Empresa não encontrada. Continuando com a operação")

try:
    seleciona_profissao(xpath=10)
except NoSuchElementException:
    print("(10) Profissão não encontrada. Continuando com a operação")

try:
    seleciona_Data(xpath=10)
except NoSuchElementException:
    print("(10) Data não encontrada. Continuando com a operação")

try:
    seleciona_NumeroCell(xpath=10)
except NoSuchElementException:
    pyautogui.press('enter')
    pyautogui.press('enter')    
    print("(10) Número de celular não encontrado. Continuando com a operação")
    print("")

# - - - - - 11
pyautogui.hotkey('alt','tab')
decimaPrimeira_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[11]/div/div[2]/a')
decimaPrimeira_Pessoa.click()

time.sleep(10) # Tempo de carregamento da página

# Nome da Segunda Pessoa + Input na Planilha 
segundo_Nome = driver.find_element(By.XPATH, '//*[@id="HeaderInfoContainer"]/div[1]/div')
driver.execute_script("window.getSelection().selectAllChildren(arguments[0])", segundo_Nome) # Seleciona o Termo orientado a XPATH escolhido
time.sleep(2) # Tempo de espera - "Talvez retirar daqui"
pyautogui.hotkey('ctrl', 'c')
pyautogui.hotkey('alt', 'tab')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

pyautogui.hotkey('alt','tab')

time.sleep(4) 

try:
    seleciona_empresa(xpath=11)
except NoSuchElementException:
    print("(11) Empresa não encontrada. Continuando com a operação")

try:
    seleciona_profissao(xpath=11)
except NoSuchElementException:
    print("(11) Profissão não encontrada. Continuando com a operação")

try:
    seleciona_Data(xpath=11)
except NoSuchElementException:
    print("(11) Data não encontrada. Continuando com a operação")

try:
    seleciona_NumeroCell(xpath=11)
except NoSuchElementException:
    pyautogui.press('enter')
    pyautogui.press('enter')
    print("(11) Número de celular não encontrado. Continuando com a operação")
    print("")

# - - - - - 12
pyautogui.hotkey('alt','tab')
decimaSegunda_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[12]/div/div[2]/a').click()

time.sleep(10) # Tempo de carregamento da página

# Nome da Segunda Pessoa + Input na Planilha 
segundo_Nome = driver.find_element(By.XPATH, '//*[@id="HeaderInfoContainer"]/div[1]/div')
driver.execute_script("window.getSelection().selectAllChildren(arguments[0])", segundo_Nome) # Seleciona o Termo orientado a XPATH escolhido
time.sleep(2) # Tempo de espera - "Talvez retirar daqui"
pyautogui.hotkey('ctrl', 'c')
pyautogui.hotkey('alt', 'tab')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

pyautogui.hotkey('alt','tab')

time.sleep(4)

try:
    seleciona_empresa(xpath=12)
except NoSuchElementException:
    print("(12) Empresa não encontrada. Continuando com a operação")

try:
    seleciona_profissao(xpath=12)
except NoSuchElementException:
    print("(12) Profissão não encontrada. Continuando com a operação")

try:
    seleciona_Data(xpath=12)
except NoSuchElementException:
    print("(12) Data não encontrada. Continuando com a operação")

try:
    seleciona_NumeroCell(xpath=12)
except NoSuchElementException:
    pyautogui.press('enter')
    pyautogui.press('enter')
    print("(12) Número de celular não encontrado. Continuando com a operação")
    print("")
 
# - - - - - 13
pyautogui.hotkey('alt','tab')
decimaTerceira_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[13]/div/div[2]/a').click()

time.sleep(10) # Tempo de carregamento da página

# Nome da Segunda Pessoa + Input na Planilha 
segundo_Nome = driver.find_element(By.XPATH, '//*[@id="HeaderInfoContainer"]/div[1]/div')
driver.execute_script("window.getSelection().selectAllChildren(arguments[0])", segundo_Nome) # Seleciona o Termo orientado a XPATH escolhido
time.sleep(2) # Tempo de espera - "Talvez retirar daqui"
pyautogui.hotkey('ctrl', 'c')
pyautogui.hotkey('alt', 'tab')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

pyautogui.hotkey('alt','tab')

time.sleep(4)

try:
    seleciona_empresa(xpath=13)
except NoSuchElementException:
    print("(13) Empresa não encontrada. Continuando com a operação")

try:
    seleciona_profissao(xpath=13)
except NoSuchElementException:
    print("(13) Profissão não encontrada. Continuando com a operação")

try:
    seleciona_Data(xpath=13)
except NoSuchElementException:
    print("(13) Data não encontrada. Continuando com a operação")

try:
    seleciona_NumeroCell(xpath=13)
except NoSuchElementException:
    pyautogui.press('enter')
    pyautogui.press('enter')
    print("(13) Número de celular não encontrado. Continuando com a operação")
    print("")

# - - - - - 14
pyautogui.hotkey('alt','tab')
decimaQuarta_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[14]/div/div[2]/a').click()

time.sleep(10) # Tempo de carregamento da página

# Nome da Segunda Pessoa + Input na Planilha 
segundo_Nome = driver.find_element(By.XPATH, '//*[@id="HeaderInfoContainer"]/div[1]/div')
driver.execute_script("window.getSelection().selectAllChildren(arguments[0])", segundo_Nome) # Seleciona o Termo orientado a XPATH escolhido
time.sleep(2) # Tempo de espera - "Talvez retirar daqui"
pyautogui.hotkey('ctrl', 'c')
pyautogui.hotkey('alt', 'tab')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

pyautogui.hotkey('alt','tab')

time.sleep(4)

try:
    seleciona_empresa(xpath=14)
except NoSuchElementException:
    print("(14) Empresa não encontrada. Continuando com a operação")

try:
    seleciona_profissao(xpath=14)
except NoSuchElementException:
    print("(14) Profissão não encontrada. Continuando com a operação")

try:
    seleciona_Data(xpath=14)
except NoSuchElementException:
    print("(14) Data não encontrada. Continuando com a operação")

try:
    seleciona_NumeroCell(xpath=14)
except NoSuchElementException:
    pyautogui.press('enter')
    pyautogui.press('enter')
    print("(14) Número de celular não encontrado. Continuando com a operação")
    print("")

# - - - - - 15
pyautogui.hotkey('alt','tab')
decimaQuinta_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[15]/div/div[2]/a').click()

time.sleep(10) # Tempo de carregamento da página

# Nome da Segunda Pessoa + Input na Planilha 
segundo_Nome = driver.find_element(By.XPATH, '//*[@id="HeaderInfoContainer"]/div[1]/div')
driver.execute_script("window.getSelection().selectAllChildren(arguments[0])", segundo_Nome) # Seleciona o Termo orientado a XPATH escolhido
time.sleep(2) # Tempo de espera - "Talvez retirar daqui"
pyautogui.hotkey('ctrl', 'c')
pyautogui.hotkey('alt', 'tab')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

pyautogui.hotkey('alt','tab')

time.sleep(4)

try:
    seleciona_empresa(xpath=15)
except NoSuchElementException:
    print("(15) Empresa não encontrada. Continuando com a operação")

try:
    seleciona_profissao(xpath=15)
except NoSuchElementException:
    print("(15) Profissão não encontrada. Continuando com a operação")

try:
    seleciona_Data(xpath=15)
except NoSuchElementException:
    print("(15) Data não encontrada. Continuando com a operação")

try:
    seleciona_NumeroCell(xpath=15)
except NoSuchElementException:
    pyautogui.press('enter')
    pyautogui.press('enter')
    print("(15) Número de celular não encontrado. Continuando com a operação")
    print("")
 
# Aumentar a capacidade para mais (15)