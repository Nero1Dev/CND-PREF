from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from PySimpleGUI import PySimpleGUI as sg  
from pyautogui import press, hotkey
from time import sleep

    
# Good themes: SystemDefault1 SystemDefault, Tanblue

sg.theme('SystemDefault')

layout = [
    [sg.Text('CNPJ:'), sg.Input(key='cnpj', size=(25, 1))],
    [sg.Radio('Apenas a identificada', "RADIO1", default=True,size=(18, 1), key='RADIO01')], 
    [sg.Radio('Todas as empresas', "RADIO1", default=False,size=(18, 1), key='RADIO02')],
    [sg.Input(size=(30,1)), sg.FolderBrowse('Browse')],
    [sg.Button('Baixar'), sg.Button('Quit')]
]

# Declaring window
janela = sg.Window('CND PREFEITURA', layout)

while True:

    events, values = janela.read()

    if events == sg.WIN_CLOSED or events == 'Quit':
        break
    
    if events == 'Browse':
        caminho = sg.popup_get_folder('Por favor onde salvar')
    
    if events == 'Baixar' and values['RADIO01'] == True:
        browser = webdriver.Firefox()
        
        browser.implicitly_wait(10)
        
        browser.get('https://agiliblue.agilicloud.com.br/portal/prefjuinamt/#/certidao')

        browser.find_element_by_xpath('/html/body/div[2]/div[1]/div/div[1]/div/div[3]/form/div[2]/div/label[2]').click()

        elemento = browser.find_element_by_xpath('//*[@id="cpfcnpjCertidoes"]')
        elemento.clear()
        elemento.send_keys(values['cnpj'])
        
        browser.find_element_by_xpath('//*[@id="btnImprimirCertidaoDebitos"]').click()

        browser.find_element_by_xpath('/html/body/div[2]/div[1]/div/div[1]/div/div[3]/form/div[10]/div/div/div[2]/div/div/label[2]').click()

        browser.find_element_by_xpath('/html/body/div[2]/div[1]/div/div[1]/div/div[3]/form/div[10]/div/div/div[3]/button[1]').click()
        
        hotkey('ctrl', 'p')
        sleep(0.5)
        for c in range(0, 6):
            press('tab')
        press('enter')
        #browser.quit()
    