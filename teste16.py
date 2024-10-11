import re
import pyautogui
import os
import pyperclip
import pandas as pd
from datetime import datetime, timedelta  # Corrigido, mantendo apenas uma importação de datetime
import sys
from unidecode import unidecode
import tkinter as tk
from tkinter import messagebox

# Carregar a planilha
Planilha_cc19 = pd.read_excel("Pasta2.xlsx")

# Verificar se "Veículo" é uma coluna válida antes de usar
if 'Veículo' not in Planilha_cc19.columns:
    print("A coluna 'Veículo' não foi encontrada na planilha.")
else:
    # Criar um DataFrame com os nomes das colunas na primeira linha
    df_colunas = pd.DataFrame([Planilha_cc19.columns], columns=Planilha_cc19.columns)

    # Concatenar o DataFrame com a nova primeira linha
    Planilha_cc19 = pd.concat([df_colunas, Planilha_cc19], ignore_index=True)

    # Localiza onde a coluna tem o valor "Veículo"
    index_veiculo = Planilha_cc19[Planilha_cc19['Veículo'] == 'Veículo'].index

    # Função para exibir uma mensagem de sucesso
    def show_success_message(msg):
        root = tk.Tk()
        root.withdraw()  # Oculta a janela principal
        messagebox.showinfo("Lancamento Romaneio", msg)  # Exibe o título e a mensagem
        root.destroy()  # Fecha a janela após exibir a mensagem

    # Itera sobre os índices e verifica o valor da linha abaixo de cada "Veículo"
    for idx in index_veiculo:
        if idx + 1 < len(Planilha_cc19):
            valor_abaixo = Planilha_cc19.loc[idx + 1, 'Veículo']
            print(f"Valor abaixo de 'Veículo' na linha {idx}: {valor_abaixo}")

            # Verifica se o valor abaixo é NaN
            if pd.isna(valor_abaixo):
                print(f"Não consta algum veículo na linha {idx + 1}.")
                # Exibe a mensagem de alerta
                show_success_message("Não consta algum veículo na relação.")
                sys.exit()  # Para a execução do código
            else:
                # Se o valor for válido, define a flag como True
                placa_encontrada = True
        else:
            print(f"Não há uma linha abaixo de 'Veículo' na linha {idx}.")

    # Se nenhuma placa válida foi encontrada, exibe uma mensagem
    if not placa_encontrada:
        show_success_message("Nenhuma placa válida encontrada após 'Veículo'.")

    # Simulando o final do código
    if __name__ == "__main__":
        sys.exit()  # Para a execução do código