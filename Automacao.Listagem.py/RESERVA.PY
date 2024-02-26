from selenium import webdriver

import pyautogui

import time

from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.chrome.service import Service

from selenium.webdriver.chrome.options import Options

from selenium.webdriver.common.by import By

from selenium.webdriver.common.keys import Keys

from selenium.webdriver.common.action_chains import ActionChains

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
    numeroCell = driver.find_element(By.XPATH, '//*[@id="HeaderInfoContainer"]/div[6]/div[1]/a')
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

# Clica no local de candidatos totais
pressiona_btn = driver.find_element(By.XPATH, '//*[@id="rowVacancy_895631"]/div[2]/div/div[2]/div[2]/div[1]/a/div/div[1]').click() # Localizador por XPATH

# Primeiro nome + Input na Planilha
time.sleep(5)
primeiro_Nome = driver.find_element(By.XPATH, '//*[@id="HeaderInfoContainer"]/div[1]/div')
driver.execute_script("window.getSelection().selectAllChildren(arguments[0])", primeiro_Nome) # Seleciona o Termo orientado a XPATH escolhido
pyautogui.hotkey('ctrl', 'c')
time.sleep(2)
pyautogui.click(x=990, y=737) # Click na aba da planilha
pyautogui.press('space')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

pyautogui.click(x=1059, y=746)  

time.sleep(4) # Tempo de espera para o carregamento 

seleciona_empresa(xpath=1)
seleciona_profissao(xpath=1)
seleciona_Data(xpath=1)
seleciona_NumeroCell(xpath=1)


# - - - - - 2
# Seleciona o campo para seguir para a próxima pessoa
#   pyautogui.hotkey('alt','tab')
#   segunda_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[2]/div/div[2]/a').click()

#   time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

# Nome da Segunda Pessoa + Input na Planilha 
#    segundo_Nome = driver.find_element(By.XPATH, '//*[@id="HeaderInfoContainer"]/div[1]/div')
#    driver.execute_script("window.getSelection().selectAllChildren(arguments[0])", segundo_Nome) # Seleciona o Termo orientado a XPATH escolhido
#    time.sleep(2) # Tempo de espera - "Talvez retirar daqui"
#    pyautogui.hotkey('ctrl', 'c')
#    pyautogui.hotkey('alt', 'tab')
#    pyautogui.hotkey('ctrl', 'v')
#    pyautogui.press('tab')

#    pyautogui.hotkey('alt','tab')

#    time.sleep(4)

#    seleciona_empresa(xpath=2)
#    seleciona_profissao(xpath=2)
#    seleciona_Data(xpath=2)
#    seleciona_NumeroCell(xpath=2)


# - - - - - 3
pyautogui.hotkey('alt','tab')
terceira_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[3]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=3)
seleciona_profissao(xpath=3)
seleciona_Data(xpath=3)
seleciona_NumeroCell(xpath=3)


# - - - - - 4
pyautogui.hotkey('alt','tab')
quarta_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[4]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=4)
seleciona_profissao(xpath=4)
seleciona_Data(xpath=4)
seleciona_NumeroCell(xpath=4)


# - - - - - 5
pyautogui.hotkey('alt','tab')
quinta_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[5]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=5)
seleciona_profissao(xpath=5)
seleciona_Data(xpath=5)
seleciona_NumeroCell(xpath=5)


# - - - - - 6
pyautogui.hotkey('alt','tab')
sexta_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[6]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=6)
seleciona_profissao(xpath=6)
seleciona_Data(xpath=6)
seleciona_NumeroCell(xpath=6)


# - - - - - 7
pyautogui.hotkey('alt','tab')
setima_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[7]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=7)
seleciona_profissao(xpath=7)
seleciona_Data(xpath=7)
seleciona_NumeroCell(xpath=7)


# - - - - - 8
pyautogui.hotkey('alt','tab')
oitava_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[8]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=8)
seleciona_profissao(xpath=8)
seleciona_Data(xpath=8)
seleciona_NumeroCell(xpath=8)


# - - - - - 9
pyautogui.hotkey('alt','tab')
nona_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[9]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=9)
seleciona_profissao(xpath=9)
seleciona_Data(xpath=9)
seleciona_NumeroCell(xpath=9)


# - - - - - 10
pyautogui.hotkey('alt','tab')
pyautogui.click(x=1362, y=368)
time.sleep(2)
decima_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[10]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=10)
seleciona_profissao(xpath=10)
seleciona_Data(xpath=10)
seleciona_NumeroCell(xpath=10)

time.sleep(10)  


# - - - - - 11
pyautogui.hotkey('alt','tab')
decimaPrimeira_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[11]/div/div[2]/a')
decimaPrimeira_Pessoa.click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=11)
seleciona_profissao(xpath=11)
seleciona_Data(xpath=11)
seleciona_NumeroCell(xpath=11)


# - - - - - 12
pyautogui.hotkey('alt','tab')
decimaSegunda_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[12]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=12)
seleciona_profissao(xpath=12)
seleciona_Data(xpath=12)
seleciona_NumeroCell(xpath=12)


# - - - - - 13
pyautogui.hotkey('alt','tab')
decimaTerceira_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[13]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=13)
seleciona_profissao(xpath=13)
seleciona_Data(xpath=13)
seleciona_NumeroCell(xpath=13)


# - - - - - 14
pyautogui.hotkey('alt','tab')
decimaQuarta_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[14]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=14)
seleciona_profissao(xpath=14)
seleciona_Data(xpath=14)
seleciona_NumeroCell(xpath=14)


# - - - - - 15
pyautogui.hotkey('alt','tab')
decimaQuinta_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[15]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=15)
seleciona_profissao(xpath=15)
seleciona_Data(xpath=15)
seleciona_NumeroCell(xpath=15)


# - - - - - 16
pyautogui.hotkey('alt','tab')
pyautogui.click()
decimaSexta_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[16]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=16)
seleciona_profissao(xpath=16)
seleciona_Data(xpath=16)
seleciona_NumeroCell(xpath=16)


# - - - - - 17 
pyautogui.hotkey('alt','tab')
decimaSetima_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[17]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=17)
seleciona_profissao(xpath=17)
seleciona_Data(xpath=17)
seleciona_NumeroCell(xpath=17)

time.sleep(5)

# - - - - - 18
pyautogui.hotkey('alt','tab')
decimaOitava_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[18]/div/div[]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=18)
seleciona_profissao(xpath=18)
seleciona_Data(xpath=18)
seleciona_NumeroCell(xpath=18)


# - - - - - 19
pyautogui.hotkey('alt','tab')
decimaNona_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[19]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=19)
seleciona_profissao(xpath=19)
seleciona_Data(xpath=19)
seleciona_NumeroCell(xpath=19)

time.sleep(10)

# - - - - - 20
pyautogui.hotkey('alt','tab')
Vigesima_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[20]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=20)
seleciona_profissao(xpath=20)
seleciona_Data(xpath=20)
seleciona_NumeroCell(xpath=20)


# - - - - - 21
pyautogui.hotkey('alt','tab')
VigesimaPrimeira_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[21]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=21)
seleciona_profissao(xpath=21)
seleciona_Data(xpath=21)
seleciona_NumeroCell(xpath=21)


# - - - - - 22
pyautogui.hotkey('alt','tab')
VigesimaSegunda_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[22]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=22)
seleciona_profissao(xpath=22)
seleciona_Data(xpath=22)
seleciona_NumeroCell(xpath=22)


# - - - - - 23
pyautogui.hotkey('alt','tab')
VigesimaTerceira_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[23]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=23)
seleciona_profissao(xpath=23)
seleciona_Data(xpath=23)
seleciona_NumeroCell(xpath=23)


# - - - - - 24
pyautogui.hotkey('alt','tab')
VigesimaQuarta_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[24]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=24)
seleciona_profissao(xpath=24)
seleciona_Data(xpath=24)
seleciona_NumeroCell(xpath=24)


# - - - - - 25
pyautogui.hotkey('alt','tab')
VigesimaQuinta_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[25]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=25)
seleciona_profissao(xpath=25)
seleciona_Data(xpath=25)
seleciona_NumeroCell(xpath=25)


# - - - - - 26
pyautogui.hotkey('alt','tab')
VigesimaSexta_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[26]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=26)
seleciona_profissao(xpath=26)
seleciona_Data(xpath=26)
seleciona_NumeroCell(xpath=26)


# - - - - - 27
pyautogui.hotkey('alt','tab')
VigesimaSetima_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[27]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=27)
seleciona_profissao(xpath=27)
seleciona_Data(xpath=27)
seleciona_NumeroCell(xpath=27)


# - - - - - 28
pyautogui.hotkey('alt','tab')
VigesimaOitava_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[28]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=28)
seleciona_profissao(xpath=28)
seleciona_Data(xpath=28)
seleciona_NumeroCell(xpath=28)


# - - - - - 29
pyautogui.hotkey('alt','tab')
VigesimaNona_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[29]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=29)
seleciona_profissao(xpath=29)
seleciona_Data(xpath=29)
seleciona_NumeroCell(xpath=29)


# - - - - - 30
pyautogui.hotkey('alt','tab')
Trigesima_Pessoa = driver.find_element(By.XPATH, '//*[@id="DivMatchesList"]/div[30]/div/div[2]/a').click()

time.sleep(4) # Tempo de carregamento da página (Talvez alterar para 5s)

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

seleciona_empresa(xpath=30)
seleciona_profissao(xpath=30)
seleciona_Data(xpath=30)
seleciona_NumeroCell(xpath=30)