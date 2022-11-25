import pyautogui
from time import sleep

def acessar_rendimentos_variariveis():
    pyautogui.hotkey('alt', 'f')
    sleep(1)
    pyautogui.hotkey('r')
    sleep(2)
    pyautogui.click(25,126, duration=2)
    sleep(2)


def cadastro_ficha_rendimentos_variaveis(cod_empresa, cod_empregado, competencia, verba, valor, nova_empresa, novo_empregado):
    if nova_empresa:
        pyautogui.doubleClick(166,178, duration=2)
        pyautogui.write(cod_empresa)
        sleep(1)
        pyautogui.press('enter')
    if novo_empregado:
        pyautogui.write(cod_empregado)
        sleep(1)
        pyautogui.press('enter')
        sleep(1)
    else:
        pyautogui.press('enter')
        sleep(1)

    pyautogui.write(competencia)
    sleep(1)
    pyautogui.write(verba)
    sleep(1)
    pyautogui.press('enter')
    sleep(1)
    pyautogui.write(valor)
    sleep(1)
    pyautogui.press('enter')
    sleep(1)

    



