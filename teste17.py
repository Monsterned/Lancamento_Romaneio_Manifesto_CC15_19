import re
import pyautogui
import os
import pyperclip
import pandas as pd
from datetime import datetime
from datetime import datetime, timedelta
import sys
from unidecode import unidecode
import tkinter as tk
from tkinter import messagebox
from datetime import datetime


caminho = os.getcwd() 
time = 1
agora = datetime.now()
agora_formatado = agora.strftime("%d/%m/%Y")

def click_image(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path)
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                center_x = position.left + position.width // 2
                center_y = position.top + position.height // 2
                pyautogui.click(center_x, center_y)
                print("Imagem foi encontrada na tela.")
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)

def show_success_message(msg):
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal
    messagebox.showinfo("Lancamento Manifesto", msg)  # Exibe o título e a mensagem
    root.destroy()  # Fecha a janela após exibir a mensagem

hora_atual = datetime.now().time()

if hora_atual.strftime("%H:%M") == "12:09":
    print("Já são 22:55, esperar 15 minutos")
    click_image('sair_sistema.png')
    click_image('sair_sistema_logoff.png')
    click_image('sair_sistema_logoff_sair.png')
    
    # Simulando o final do código
    if __name__ == "__main__":  # Corrigido o nome da condição

        # Aqui você pode adicionar seu código que será executado
        show_success_message("Pausado por ser 23 horas, clique no ok quando logar na tela inicial!")
else:
    print("Ainda não são 22:55.")
