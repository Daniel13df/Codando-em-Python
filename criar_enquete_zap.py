import pyautogui
import time
import sys


def criar_enquete():

    msg = pyautogui.prompt("Essa é uma automação para criação de enquete no whatsapp.", "Automação Daniel Coelho")

    if msg == None:
            pyautogui.alert("Operação cancelada com sucesso!")
            sys.exit()

    pyautogui.alert("Primeiro clique na conversa dentro do whatsapp que deseja, clique no botão da enquete e depois confirme aqui.")
    time.sleep(3)
    pyautogui.typewrite("Nome da enquete.")
    pyautogui.press("Tab", presses=2)

    lista_de_opcoes = ["Option1", "Option2", "Option3", "Option4", "Option5", "Option6", "Option7", "Option8", "Option9", "Option10", "Option11", "Option12"]

    for opcoes in lista_de_opcoes:
        pyautogui.write(opcoes)
        pyautogui.press("Tab", presses=2)
    pyautogui.press("Tab")
    pyautogui.press("Enter")
    pyautogui.alert("Automação finalizada com sucesso.")
