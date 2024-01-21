from undetected_chromedriver import Chrome, ChromeOptions
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium import webdriver
from time import sleep
import os

url_relatorio = "https://www.bling.com.br/relatorio.estoque.consumo.php"

options = webdriver.ChromeOptions()
options.add_argument("--disable-notifications")
dir_path = os.getcwd()
profile = os.path.join(dir_path, "profile", "one")
options.add_argument(r"user-data-dir={}".format(profile))

nav = Chrome(driver_executable_path="./chromedriver.exe", options=options)


###################
# Xpaths de Login #
###################
link_google = '//*[@id="rso"]/div[1]/div/div/div/div/div/div/div/div[1]/div/span/a/h3'
btn_entrar = '//*[@id="dropdown"]/a'
email_usuario = '//*[@id="username"]'
senha = '//*[@id="senha"]'
btn_login = '//*[@id="login-buttons-site"]/button'

#######################
# Xpaths de Relatorio #
#######################
edt_mes = '//*[@id="mes"]'
opc_periodo = '//*[@id="filtros-relatorios"]/div[2]/div/div[2]/label'
edt_data_inicial = '//*[@id="dataIni"]'
edt_data_final = '//*[@id="dataFim"]'
filter_deposito = '//*[@id="idDeposito"]'
btn_visualizar = '//*[@id="btn_visualizar"]'


# def persistence_finder(*args):
#     if args:
#         if args[0] == 1:
#             while True:
#
#
#     return


if __name__ == "__main__":
    nav.get("https://www.google.com")
    sleep(3)
    nav.get("https://www.google.com/search?q=bling")
    sleep(3)
    nav.find_element(By.XPATH, link_google).click()
    sleep(3)
    nav.find_element(By.XPATH, btn_entrar).click()

    # persistence_finder(1, )
    input("finish")
    nav.close()
