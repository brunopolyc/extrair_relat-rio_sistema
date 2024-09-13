from config import USUARIO_PROTHEUS, SENHA_PROTHEUS, MODULO_CONTABILIDADE
import openpyxl
import pyautogui
import datetime
import time
import tkinter as tk
from tkinter import messagebox

pyautogui.PAUSE = 0.5



def show_popup():
    # Cria uma janela pop-up com uma mensagem e um botão OK
    messagebox.showinfo("Informação", "Os Relatórios Razões foram Gerados.")


def abrir_totvs():
    while True:
        protheus = pyautogui.locateOnScreen("protheus.png")
        
        if protheus is not None:
            x_eixo, y_eixo = pyautogui.center(protheus)
            pyautogui.moveTo(x_eixo, y_eixo, 1)
            pyautogui.click()
            break  # Sai do loop após encontrar e clicar na imagem

        time.sleep(1)
    
def logar_totvs():
    time.sleep(5)
    pyautogui.press("enter")
    time.sleep(10)
    pyautogui.write(USUARIO_PROTHEUS)
    pyautogui.press("tab")
    pyautogui.write(SENHA_PROTHEUS)
    pyautogui.press("enter")
    time.sleep(4)
    pyautogui.press("tab", presses=2, interval=1)
    pyautogui.write(MODULO_CONTABILIDADE)
    pyautogui.press("tab")
    pyautogui.press("enter")
    time.sleep(4)
    
def abrir_razao_contabil():
    time.sleep(4)
    pyautogui.click(58,503, duration=1)
    pyautogui.click(51,584, duration=1)
    pyautogui.click(66,602, duration=1)

def parametros_planilha():
    planilha_contabil = openpyxl.load_workbook('RAZAO_CONTABIL.xlsx')
    contabil_sheet = planilha_contabil['contabil']

    for linha in contabil_sheet.iter_rows(min_row=2):
        pyautogui.click(63,603, duration=1)
        time.sleep(3)
        pyautogui.click(63,603, duration=1)
        time.sleep(12)
        pyautogui.press("enter")
        time.sleep(8)
        pyautogui.press('tab', presses=2, interval=1)
        pyautogui.write(linha[0].value.strftime('%d/%m/%Y'))  # data inicio
        pyautogui.write(linha[1].value.strftime('%d/%m/%Y'))  # data final
        pyautogui.press('tab', presses=8, interval=1)
        pyautogui.write(linha[2].value.upper())  # EG Inicial
        pyautogui.press('tab')
        pyautogui.write(linha[3].value.upper())  # EG Final
        pyautogui.click(1166, 735, duration=0.5)
        time.sleep(2)
        pyautogui.click(716, 484, duration=0.5)
        pyautogui.click(1045, 450, duration=0.5)
        pyautogui.click(910, 498, duration=0.5)
        pyautogui.click(887, 345, duration=0.5)
        pyautogui.write(linha[4].value.upper())  # nome do arquivo_razão
        pyautogui.click(1227, 851, duration=1)
        time.sleep(2)
        pyautogui.press("enter")
        time.sleep(1)
        pyautogui.press("enter")
        time.sleep(1)
        pyautogui.press("enter")
        time.sleep(5)
        pyautogui.hotkey('alt', 'tab')



if __name__ == "__main__":
    #abrir_totvs()
    #logar_totvs()   
    abrir_razao_contabil()
    parametros_planilha()
    show_popup()
