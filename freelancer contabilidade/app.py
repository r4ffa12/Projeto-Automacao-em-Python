import openpyxl
import pyperclip
import pyautogui
from time import sleep


workbook = openpyxl.load_workbook('produtos_ficticios (1).xlsx')
sheet_produtos = workbook['Produtos']

#EXEMPLO para todas as linha 
#nome_produto = linha[0].value
#    pyperclip.copy(nome_produto)
#    pyautogui.click(1678,389,duration=1)
#    pyautogui.hotkey('ctrl','v')

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1138,353,duration=1)
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1143,439,duration=1)
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1150,569,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1143,662,duration=1)
    pyautogui.hotkey('ctrl','v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1148,746,duration=1)
    pyautogui.hotkey('ctrl','v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1158,824,duration=1)
    pyautogui.hotkey('ctrl','v')


    pyautogui.click(1127,886)
    sleep(4)

#   preco = linha[6].value
#   pyperclip.copy(preco)
#   pyautogui.click()
#   pyautogui.hotkey('ctrl','v')
#
#   quantidade_em_estoque = linha[7].value
#   pyperclip.copy(quantidade_em_estoque)
#   pyautogui.click()
#   pyautogui.hotkey('ctrl','v')
#
#   data_de_validade = linha[8].value
#   pyperclip.copy(data_de_validade)
#   pyautogui.click()
#   pyautogui.hotkey('ctrl','v')
#
#   cor = linha[9].value
#   pyperclip.copy(cor)
#   pyautogui.click()
#   pyautogui.hotkey('ctrl','v')
#
#   tamanho = linha[10].value
#   pyperclip.copy(tamanho)
#   pyautogui.click()
#   pyautogui.hotkey('ctrl','v')
#
#   material = linha[11].value
#   pyperclip.copy(material)
#   pyautogui.click()
#   pyautogui.hotkey('ctrl','v')
#
#   fabricante = linha[12].value
#   pyperclip.copy(fabricante)
#   pyautogui.click()
#   pyautogui.hotkey('ctrl','v')
#
#   pais_origem = linha[13].value
#   pyperclip.copy(pais_origem)
#   pyautogui.click()
#   pyautogui.hotkey('ctrl','v')
#
#   obsevacoes = linha[14].value
#   pyperclip.copy(obsevacoes)
#   pyautogui.click()
#   pyautogui.hotkey('ctrl','v')
#
#   codigo_de_barras = linha[15].value
#   pyperclip.copy(codigo_de_barras)
#   pyautogui.click()
#   pyautogui.hotkey('ctrl','v')
#
#   localizacao_armazem = linha[16].value
#   pyperclip.copy(localizacao_armazem)
#   pyautogui.click()
#   pyautogui.hotkey('ctrl','v')


