import openpyxl
import pyautogui
import time
import datetime
import tkinter as tk
from tkinter import messagebox
import os
import shutil

def show_popup():
    # Cria uma janela pop-up com uma mensagem e um botão OK
    messagebox.showinfo("Informação", "Os Relatórios Comparativos foram Gerados.")


pyautogui.PAUSE = 0.5

# 1 abrir aba do comparativo
def abrir_comparativo_contabil():
    pyautogui.hotkey('alt' , 'tab')
    #pyautogui.click(739,1055, duration= 1)
    pyautogui.click(67,625, duration= 1)
    
# 2 repetição do comparativo
def preencher_campos():
    planilha_contabil = openpyxl.load_workbook('RAZAO_CONTABIL.xlsx')
    contabil_sheet = planilha_contabil['contabil']

    for linha in contabil_sheet.iter_rows(min_row=2):

        pyautogui.click(82,782, duration= 1)
        time.sleep(10)
        pyautogui.click(82,782, duration= 1)
        time.sleep(4)
        pyautogui.click(729,484, duration= 1)
        pyautogui.click(1153,457, duration= 1)
        pyautogui.click(904,345, duration= 1) 
        time.sleep(2)
        pyautogui.write(linha[5].value)  # nome do arquivo_comparativo
        pyautogui.click(1039,850, duration= 1)
        pyautogui.click(1040,880, duration= 1)     
        time.sleep(3)    
        pyautogui.write(linha[0].value.strftime('%d/%m/%Y'))  # data inicio
        pyautogui.write(linha[1].value.strftime('%d/%m/%Y'))  # data final       
        pyautogui.write(linha[2].value)  # EG Inicial
        pyautogui.press('tab')
        pyautogui.write(linha[3].value)  # EG Final
        pyautogui.click(1163,735, duration= 1) 
        time.sleep(2)
        pyautogui.click(1222,854, duration= 1) 
        pyautogui.press("enter")
        time.sleep(2)
        pyautogui.press("enter")
        time.sleep(2)
        pyautogui.press("enter")
        time.sleep(2)
        pyautogui.hotkey('alt', 'tab')

def copiar_arquivos(origem, destino):
    # Cria a pasta de destino, se não existir
    if not os.path.exists(destino):
        os.makedirs(destino)
        print(f"A pasta {destino} foi criada.")
    else:
        print(f"A pasta {destino} já existe.")
    
    # Lista todos os arquivos na pasta de origem
    arquivos = os.listdir(origem)
    
    # Copia cada arquivo para a pasta de destino
    for arquivo in arquivos:
        caminho_origem = os.path.join(origem, arquivo)
        caminho_destino = os.path.join(destino, arquivo)
        shutil.copy2(caminho_origem, caminho_destino)
        print(f"Arquivo {arquivo} copiado para {destino}.")

# Caminho da pasta de origem
pasta_origem = r"C:\Users\bruno.polycarpo\AppData\Local\Temp\totvsprinter"
# Caminho completo da pasta de destino

# Solicita o nome da nova pasta de destino ao usuário
nome_pasta_destino = input("Digite o nome da nova pasta de destino: ")

pasta_destino = os.path.join(r"F:\4. CONTABILIDADE\CONTABILIDADE DOS CONSÓRCIOS\RELATÓRIOS BAIXADOS", nome_pasta_destino)
# Chama a função para copiar os arquivos passando os argumentos corretos
copiar_arquivos(pasta_origem, pasta_destino)


if __name__ == "__main__":
    abrir_comparativo_contabil()
    preencher_campos()
    show_popup()
    copiar_arquivos()




 





    




