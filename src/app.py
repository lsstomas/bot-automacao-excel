from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import openpyxl

# Definir driver
driver = webdriver.Edge()
driver.get("https://contabilidade-devaprender.netlify.app/")
sleep(3)

# Preencher campo de email
email = driver.find_element(By.XPATH, "//input[@id='email']")
sleep(1)
email.send_keys("admin@email.com")

# Preencher campo de senha
password = driver.find_element(By.XPATH, "//input[@id='senha']")
sleep(1)
password.send_keys("admin1234")

# Realizar login
button = driver.find_element(By.XPATH, "//button[@id='Entrar']")
sleep(1)
button.click()
sleep(4)

# Extrair dados da planilha
data = openpyxl.load_workbook("./empresas.xlsx")
data_page = data["dados empresas"]

# Preencher os campos de cadastro de empresas
for line in data_page.iter_rows(min_row=2, values_only=True):
    nome_empresa, email, telefone, endereco, cnpj, area_atuacao, quantidade_de_func, data_fundacao = line
    
    driver.find_element(By.ID, "nomeEmpresa").send_keys(nome_empresa)
    sleep(1)
    
    driver.find_element(By.ID, "emailEmpresa").send_keys(email)
    sleep(1)
    
    driver.find_element(By.ID, "telefoneEmpresa").send_keys(telefone)
    sleep(1)
    
    driver.find_element(By.ID, "enderecoEmpresa").send_keys(endereco)
    sleep(1)
    
    driver.find_element(By.ID, "cnpj").send_keys(cnpj)
    sleep(1)
    
    driver.find_element(By.ID, "areaAtuacao").send_keys(area_atuacao)
    sleep(1)
    
    driver.find_element(By.ID, "numeroFuncionarios").send_keys(quantidade_de_func)
    sleep(1)
    
    driver.find_element(By.ID, "dataFundacao").send_keys(data_fundacao)
    sleep(1)
    
    driver.find_element(By.ID, "Cadastrar").click()
    sleep(3)