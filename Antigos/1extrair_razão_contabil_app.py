import openpyxl
from openpyxl import load_workbook
import pyautogui
import time

pyautogui.PAUSE = 0.5

usuario = "bruno.polycarpo"
senha = "Bru!0704"
modulo_contabilidade = "34"

# 1 abrir totvs
pyautogui.press("win")
pyautogui.write("PROTHEUS ")
pyautogui.press("down")  
pyautogui.press("enter")               

# 2 apertar ok na janela
time.sleep(5)
pyautogui.press("enter")

# 3 digitar usuário (tab), senha (tab) 
time.sleep(10)
pyautogui.write(usuario)
pyautogui.press("tab")
pyautogui.write(senha)

# 4 apertar ok
pyautogui.press("enter")

# 5 tabular 3 vezes e colocar a contabilidade (34) no nº.
time.sleep(4)
pyautogui.press("tab")
pyautogui.press("tab")
pyautogui.write(modulo_contabilidade)

# 5 apertar ok
pyautogui.press("tab")
pyautogui.press("tab")
pyautogui.press("tab")
pyautogui.press("enter")
time.sleep(4)

#6 utilizar mouse para clicar no caminho para abrir razão contábil
pyautogui.click(49,504, duration= 1)
pyautogui.click(57,584, duration= 1)

#7 fazer repetição da planilha e preencher informações contábeis de acordo com a planilha
planilha_contabil = openpyxl.open('RAZAO_CONTABIL.xlsx')
contabil_sheet = planilha_contabil['contabil']

for linha in contabil_sheet.iter_rows(min_row=2):
    pyautogui.click(71,601, duration= 1)
    time.sleep(3)
    pyautogui.click(71,601, duration= 1)
    time.sleep(9)
    pyautogui.press("enter")
    time.sleep(8)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.write(linha[0].value) #data inicio
    pyautogui.write(linha[1].value) #data final             
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab') 
    pyautogui.write(linha[2].value) #EG Inicial
    pyautogui.press('tab')   
    pyautogui.write(linha[3].value) #EG Final
    pyautogui.click(1166,735, duration=1)
    time.sleep(1)
    pyautogui.click(716,484, duration=1)
    pyautogui.click(1045,450, duration=1)
    pyautogui.click(910,498, duration=1)
    pyautogui.click(887,345, duration=1)
    pyautogui.write(linha[4].value) #nome do arquivo
    pyautogui.click(1227,851, duration=1)
    time.sleep(2)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(5)
 

    #pyautogui.hotkey("alt", "tab")
    #pyautogui.hotkey("crtl", "shift", "s")
    #pyautogui.click(1404,771)
    #pyautogui.click(634,58)
    #pyautogui.write(os.path.join(pasta_salvar))
    #pyautogui.click(785,538)
    #time.sleep(3)
    #pyautogui.hotkey("alt", "tab")



    




