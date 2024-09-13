from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pyodbc
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from datetime import datetime
date_format = f"%d/%m/%Y"


# Acesso ao Banco de Dados
database_file = 'C:\\Users\\01780742177\\Desktop\\BOT_TCE.accdb'
# Criando a conexão
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=' + database_file + ';'
)
conn = pyodbc.connect(conn_str)

def main ():
    BuscarRHNet ()

def BuscarRHNet ():
    browser = webdriver.Chrome()
    browser.get('https://tcenet.tce.go.gov.br')
    browser.maximize_window()
    time.sleep(3)
    browser.implicitly_wait(10)

    browser.get("https://aplicacoes.expresso.go.gov.br/")

    browser.maximize_window()

    browser.find_element(By.NAME, "usernameUserInput").send_keys("06343161171")
    browser.find_element(By.NAME,"password").send_keys("1072697ass")
    browser.find_element(By.XPATH,"//*[@id='loginForm']/button").click()

    for i in range(1, 101):
        div = browser.find_element(By.XPATH,"/html/body/app-root/app-main/div/div/div/app-sistemas/div/p-dataview/div/div[2]/div/div[" + str(i) + "]").text

        if "RHNet" in div:
            browser.find_element(By.XPATH,"/html/body/app-root/app-main/div/div/div/app-sistemas/div/p-dataview/div/div[2]/div/div[" + str(i) + "]").click()
            break

    browser.switch_to.frame("principal")
    browser.execute_script("arguments[0].click();", browser.find_element(By.XPATH, '//*[text()="Consulta Ocorrências"]'))

    cursor = conn.cursor()
    cursor.execute("SELECT * FROM TESTE WHERE Buscado IS NULL;")
    pessoas = cursor.fetchall()
    npessoa = 0
    for pessoa in pessoas:
        aux = 0
        browser.find_element(By.XPATH,"/html/body/form/center/table/tbody/tr/td[2]/input").clear()
        browser.find_element(By.XPATH,"/html/body/form/center/table/tbody/tr/td[2]/input").send_keys(pessoa[2])
        browser.find_element(By.XPATH,"//input[@value='Consultar CPF']").click()
        selectElement = browser.find_element(By.NAME,"codVinculo")
        # optionElements = selectElement.find_elements(By.TAG_NAME,"option")
        Select(selectElement).select_by_value(pessoa[1])
        # for optionElement in optionElements:
        #     if pessoa[1] in optionElement.text:
        #             optionElement.Click
        #             continue
        browser.find_element(By.XPATH,"//input[@value='Consultar']").click()
        tabela = browser.find_element(By.XPATH,"/html/body/form/center/div/div[2]/div/div[2]/table")
        linhas = tabela.find_elements(By.TAG_NAME, "tr")
        for linha in linhas:
            browser.implicitly_wait(1)
            colunas = linha.find_elements(By.TAG_NAME, "td")
            browser.implicitly_wait(10)
            if len(colunas) == 10:
                if colunas[4].text == "Afastamento":
                    aux += 1
                    a = datetime.strptime(colunas[6].text, date_format)
                    b = datetime.strptime(colunas[7].text, date_format)
                    delta = b - a
                    SQL_UPDATE = f"INSERT INTO TabelaAfastamentosSEDUC2022 (Nome, CPF, Vinculo, Identificador, TipoOcorrencia, DataInicio, DataTermino, QntDias) VALUES ('{pessoa[0]}','{pessoa[2]}','{pessoa[1]}','{colunas[1].text}','{colunas[3].text}','{colunas[6].text}','{colunas[7].text}', '{delta.days + 1}');"
                    cursor.execute(SQL_UPDATE)
                    conn.commit() 
        npessoa += 1
        print (f"{npessoa} de {len(pessoas)} - {aux} afastamentos encontrados")
        SQL_UPDATE = f"UPDATE TESTE SET Buscado = 'OK' WHERE CPF = '{pessoa[2]}';"
        cursor.execute(SQL_UPDATE)
        conn.commit() 
        browser.find_element(By.XPATH,"//input[@value='Cancelar']").click()
main()