from config import USUARIO_PROTHEUS, SENHA_PROTHEUS, MODULO_CONTABILIDADE
import openpyxl
import pyautogui
import time
from tkinter import messagebox

pyautogui.PAUSE = 0.5

def show_popup(message="Os Relatórios Razões foram Gerados."):
    messagebox.showinfo("Informação", message)

def localizar_e_clicar(imagem, duracao=1, timeout=30):
    for _ in range(timeout):
        localizacao = pyautogui.locateOnScreen(imagem)
        if localizacao:
            x, y = pyautogui.center(localizacao)
            pyautogui.moveTo(x, y, duracao)
            pyautogui.click()
            return True
        time.sleep(1)
    return False

def abrir_totvs():
    """Abre o aplicativo Protheus."""
    if not localizar_e_clicar("protheus.png"):
        raise Exception("Imagem do Protheus não encontrada.")

def logar_totvs():
    """Realiza o login no sistema Protheus."""
    time.sleep(5)
    pyautogui.press("enter")
    time.sleep(10)
    pyautogui.write(USUARIO_PROTHEUS)
    pyautogui.press("tab")
    pyautogui.write(SENHA_PROTHEUS)
    pyautogui.press("enter")
    time.sleep(4)
    pyautogui.press("tab", presses=2)
    pyautogui.write(MODULO_CONTABILIDADE)
    pyautogui.press("enter")
    time.sleep(4)

def abrir_razao_contabil():
    """Navega pelos menus para abrir a Razão Contábil."""
    coordenadas = [(58, 503), (51, 584), (66, 602)]
    for x, y in coordenadas:
        pyautogui.click(x, y, duration=1)
        time.sleep(1)

def preencher_parametros(linha):
    """Preenche os parâmetros de uma linha da planilha."""
    pyautogui.click(63, 603, duration=1)
    time.sleep(3)
    pyautogui.click(63, 603, duration=1)
    time.sleep(12)
    pyautogui.press("enter")
    time.sleep(8)
    pyautogui.press("tab", presses=2)
    pyautogui.write(linha[0].value.strftime('%d/%m/%Y'))  # Data início
    pyautogui.write(linha[1].value.strftime('%d/%m/%Y'))  # Data final
    pyautogui.press("tab", presses=8)
    pyautogui.write(linha[2].value.upper())  # EG Inicial
    pyautogui.press("tab")
    pyautogui.write(linha[3].value.upper())  # EG Final

def salvar_relatorio(nome_arquivo):
    """Executa os comandos de salvar o relatório."""
    botoes = [(1166, 735), (716, 484), (1045, 450), (910, 498), (887, 345)]
    for x, y in botoes:
        pyautogui.click(x, y, duration=0.5)
        time.sleep(0.5)
    
    pyautogui.write(nome_arquivo.upper())  # Nome do arquivo razão
    pyautogui.click(1227, 851, duration=1)
    time.sleep(2)
    pyautogui.press("enter", presses=3, interval=1)
    time.sleep(5)
    pyautogui.hotkey('alt', 'tab')

def parametros_planilha(caminho_planilha='RAZAO_CONTABIL.xlsx'):
    """Abre a planilha e preenche os parâmetros de cada linha."""
    planilha = openpyxl.load_workbook(caminho_planilha)
    sheet = planilha['contabil']
    
    for linha in sheet.iter_rows(min_row=2):
        preencher_parametros(linha)
        salvar_relatorio(linha[4].value)

if __name__ == "__main__":
    try:
        abrir_totvs()
        #logar_totvs()   
        abrir_razao_contabil()
        parametros_planilha()
        show_popup()
    except Exception as e:
        show_popup(f"Erro: {str(e)}")
