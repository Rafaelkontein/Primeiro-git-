from selenium import webdriver
from time import sleep
import pyautogui
import openpyxl
import pandas as pd


class ChromeAuto:
    def __init__(self):
        self.driver_path = 'chromedriver'
        self.options = webdriver.ChromeOptions()
        self.options.add_argument('--profile-directory=1')
        self.chrome = webdriver.Chrome(
            self.driver_path,
            options=self.options
        )

    def clica_sign_in(self):
        try:
            btn_receptivo= self.chrome.find_element_by_css_selector('body > div > div.page-container > div.page-content-wrapper > div > div:nth-child(11) > div:nth-child(3) > a')
            btn_receptivo.click()
            sleep(1)
            inss=self.chrome.find_element_by_class_name('thumbnail')
            inss.click()
        except Exception as e:
            print('Erro ao clicar em Sign in:', e)




    def atualizar(self):
        try:
            pyautogui.press('f5')
        except Exception as b:
            print('Erro ao atualizar', b)

    def acessa(self, site):  # Aqui o site que vc vai entrar
        self.chrome.get(site)

    def sair(self):  # aqui é para sair do site
        self.chrome.quit()

    def faz_login(self):
        try:
            input_login = self.chrome.find_element_by_id('exten') # aqui vc etsá pegando o id do campo de escrever
            input_password = self.chrome.find_element_by_css_selector('#login-form > div:nth-child(3) > div > input')# aqui voce está pedindo =
            btn_entrar= self.chrome.find_element_by_css_selector('#login-form > div:nth-child(4)')
            sleep(2)
            input_login.send_keys('roger@2215')  # aqui voce está pedindo para escrever no login
            input_password.send_keys('123456')  # Aqui voce está falando para escrever a senha
            sleep(1)
            btn_entrar.click()
            sleep(1)
            pyautogui.press('enter')



        except Exception as e:
             print('Erro ao fazer login:', e)

    def pegar_dados(self):
        pedidos = openpyxl.load_workbook('alaa.xlsx')
        nomes_planilhas = pedidos.sheetnames  # aqui vc ta pegadno quantas paginas tem no excel
        planilhas1 = pedidos['Planilha1']  # Aqui vc ta pegando tudo que tem na planilha

        dados = []
        for camp in planilhas1['a']:  # aqui vc ta pegando exatamente oq tem na coluna b da planilha
            if camp.value is not None:
                dados.append(camp.value)
        cart_margin = []
        margim = []
        for index in range(len(dados)):
            try:
                sleep(1)
                dados1 = dados[index]
                sleep(1)
                campo_escrever = self.chrome.find_element_by_css_selector('#NR_BENEFICIO')
                btn= self.chrome.find_element_by_css_selector('#btn_buscaCpf')
                erro=self.chrome.find_element_by_id('content-dados-cliente')
                campo_escrever.clear()
                campo_escrever.send_keys(dados1)
                sleep(3)
                btn.click()
                sleep(4)
                financeiro = self.chrome.find_element_by_css_selector('#content-dados-cliente > div:nth-child(2) > div:nth-child(3) > div > div.progress-info').text
                margim.append(financeiro)
                sleep(3)
                cartao=self.chrome.find_element_by_css_selector('#content-dados-cliente > div:nth-child(2) > div:nth-child(4) > div > div.progress-info').text

                cart_margin.append(cartao)

                sleep(3)




            except Exception as e:
              margim.append('Sem item')
              cart_margin.append('Sem item')

        print(cart_margin)
        print(margim)
        import pandas as pd

        data = dados

        # Converta o dicionário em DataFrame
        df = pd.DataFrame(data)

        # Usando 'endereço' como o nome da coluna e igualando-a à lista
        df2 = df.assign(beneficio=dados,magim=margim,cartaoo=cart_margin)

        # observe o resultado
        df2.to_excel('grupo.xlsx')
if __name__ == '__main__':
    chrome = ChromeAuto()
    chrome.acessa('http://sistemavanguard.ddns.net:8091/vanguard/index.php ')

    chrome.faz_login()
    sleep(5)
    chrome.clica_sign_in()
    sleep(3)
    chrome.pegar_dados()
