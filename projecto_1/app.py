import openpyxl
import pyautogui as rato
import time


livros_planilhas = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas_planilhas = livros_planilhas['vendas']

for linha in vendas_planilhas.iter_rows(min_row=2, values_only=True):
    nome,produto, quantidade, categoria=linha

    #cliente
    rato.click(67,158, duration=1.5)
    rato.write(nome)
    time.sleep(1)

    #produto
    rato.click(80,220, duration=1.5)
    rato.write(produto)
    time.sleep(1)

    #quantidade
    rato.click(57,286, duration=1.5)
    rato.write(str(quantidade))
    time.sleep(1)

    #categoria
    rato.click(249,160, duration=1.5)
    rato.write(categoria)
    time.sleep(1)

    #salvar
    rato.click(238,286)
    time.sleep(2)
    
    #confirmar
    rato.click(711,463)
    time.sleep(2)





