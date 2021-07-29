from selenium.webdriver.common.keys import Keys
from selenium import webdriver
import pandas as pd
from time import sleep

x = pd.read_excel('/home/nero1dev/Codes/Selenium/Cnpj.ods')

for c in range(0, 3):

    cnpj = x['cnpj'] [c]

    browser = webdriver.Firefox()
    browser.get('https://agiliblue.agilicloud.com.br/portal/prefjuinamt/#/certidao')
    
    sleep(5)

    browser.find_element_by_xpath('/html/body/div[2]/div[1]/div/div[1]/div/div[3]/form/div[2]/div/label[2]').click()

    elemento = browser.find_element_by_xpath('//*[@id="cpfcnpjCertidoes"]')
    elemento.clear()
    elemento.send_keys(f'{cnpj}')

    browser.find_element_by_xpath('//*[@id="btnImprimirCertidaoDebitos"]').click()

    browser.find_element_by_xpath('/html/body/div[2]/div[1]/div/div[1]/div/div[3]/form/div[10]/div/div/div[2]/div/div/label[2]').click()

    browser.find_element_by_xpath('/html/body/div[2]/div[1]/div/div[1]/div/div[3]/form/div[10]/div/div/div[3]/button[1]').click()

    browser.quit()
