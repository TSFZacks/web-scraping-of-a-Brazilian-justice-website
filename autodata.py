from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from time import sleep
import openpyxl
import subprocess

numero_oab = 46479

driver = webdriver.Chrome()
driver.get('https://pje-consulta-publica.tjmg.jus.br/')
sleep(30)

campo_oab = driver.find_element(By.XPATH,"//input[@id='fPP:Decoration:numeroOAB']")
campo_oab.send_keys(numero_oab)
dropdown_estados = driver.find_element(By.XPATH,"//select[@id='fPP:Decoration:estadoComboOAB']")
opcoes_estados = Select(dropdown_estados)
opcoes_estados.select_by_visible_text('PE')
botao_pesquisar = driver.find_element(By.XPATH,"//input[@id='fPP:searchProcessos']")
botao_pesquisar.click()
sleep(15)

processos = driver.find_elements(By.XPATH,"//b[@class='btn-block']")
workbook = openpyxl.Workbook()
for processo in processos:
    processo.click()
    sleep(10)
    janelas = driver.window_handles
    driver.switch_to.window(janelas[-1])
    driver.set_window_size(1920,1080)
    numero_processo = driver.find_elements(By.XPATH,"//div[@class='col-sm-12 ']")
    numero_processo = numero_processo[0]
    numero_processo = numero_processo.text

    data_distribuicao = driver.find_elements(By.XPATH,"//div[@class='value col-sm-12 ']")
    data_distribuicao = data_distribuicao[1]
    data_distribuicao = data_distribuicao.text

    movimentacoes = driver.find_elements(By.XPATH,"//div[@id='j_id132:processoEventoPanel_body']//tr[contains(@class,'rich-table-row')]//td//div//div//span")
    lista_movimentacoes = []
    for movimentacao in movimentacoes:
        lista_movimentacoes.append(movimentacao.text)

    sheet = workbook.create_sheet(numero_processo)
    sheet['A1'] = "Número Processo"
    sheet['B1'] = "Data Distribuição"
    sheet['C1'] = "Movimentações"
    sheet['A2'] = numero_processo
    sheet['B2'] = data_distribuicao

    for index, movimentacao in enumerate(lista_movimentacoes, start=2):
        sheet[f'C{index}'] = movimentacao

    driver.close()
    sleep(5)
    driver.switch_to.window(driver.window_handles[0])

# Salva todas as planilhas em um único arquivo Excel
workbook.remove(workbook['Sheet'])  # Remove a planilha padrão
workbook.save('dados.xlsx')

# Abre o arquivo Excel no final
excel_file = 'dados.xlsx'
subprocess.run(['start', excel_file], shell=True)