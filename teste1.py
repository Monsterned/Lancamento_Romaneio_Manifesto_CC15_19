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

def iniciar_robo(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path)
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                print("Imagem foi encontrada na tela.")
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)

def finalizar_baixa(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                print("Imagem foi encontrada na tela.")
                break
        except Exception as e:
            click('ja_inserida_ocorrencia.png')
            click('yes.png')
            click('yes_marcado.png')
            click('ok.png')
            click('yes2.png')
            click('yes3.png')
            click('cancelar_ocorrencia.png')           
            print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(0.5)

def click_ok(image_path,image_path2,image_path3, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path)
    image_path2 = os.path.join(current_dir, caminho_imagem, image_path2)
    image_path3 = os.path.join(current_dir, caminho_imagem, image_path3)
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                center_x = position.left + position.width // 2
                center_y = position.top + position.height // 2
                pyautogui.click(center_x, center_y)
                print("OK foi encontrada na tela.")
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        try:
            position2 = pyautogui.locateOnScreen(image_path2, confidence=confidence)
            if position2:
                center_x = position2.left + position2.width // 2
                center_y = position2.top + position2.height // 2
                pyautogui.click(center_x, center_y)
                print("OK foi encontrada na tela.")
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        try:
            position3 = pyautogui.locateOnScreen(image_path3, confidence=confidence)
            if position3:
                center_x = position3.left + position3.width // 2
                center_y = position3.top + position3.height // 2
                pyautogui.click(center_x, center_y)
                print("OK foi encontrada na tela.")
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)

def click_image_salvar(image_path, image_path2, image_path3, image_path4, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path)
    image_path2 = os.path.join(current_dir, caminho_imagem, image_path2)
    image_path3 = os.path.join(current_dir, caminho_imagem, image_path3)
    image_path4 = os.path.join(current_dir, caminho_imagem, image_path4)
    aviso = False  # Variável de controle para saber se a imagem 4 foi encontrada
    tentativa = 0
    while True:
        try:
            # Procurar pela imagem 4
            position4 = pyautogui.locateOnScreen(image_path4, confidence=confidence)
            if position4:
                click_ok('ok_marcado.png','ok.png','ok2.png')
                pyautogui.sleep(1)
                if tentativa == 0:
                    click_info('transportador.png')
                    pyautogui.sleep(2)
                    pyautogui.write('129')
                    pyautogui.sleep(2)
                    pyautogui.press('tab')
                    pyautogui.sleep(2)
                    tentativa = 1
                else:
                    aviso = True  # Definir o aviso como True
                    return aviso  # Sair da função e retornar a variável
        except Exception as e:
            print("Imagem 4 não encontrada na tela. Aguardando...")
        try:
            # Procurar pela imagem 2
            position2 = pyautogui.locateOnScreen(image_path2, confidence=confidence)
            if position2:
                center_x = position2.left + position2.width // 2
                center_y = position2.top + position2.height // 2
                pyautogui.click(center_x, center_y)
                pyautogui.sleep(1)
        except Exception as e:
            print("Imagem 2 não encontrada na tela. Aguardando...")
        try:
            # Procurar pela imagem 1
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                center_x = position.left + position.width // 2
                center_y = position.top + position.height // 2
                pyautogui.click(center_x, center_y)
                pyautogui.sleep(1)
                print("Imagem foi encontrada na tela.")
        except Exception as e:
            print("Imagem 1 não encontrada na tela. Aguardando...")
        try:
            # Procurar pela imagem 3
            position3 = pyautogui.locateOnScreen(image_path3, confidence=confidence)
            if position3:
                print("Ainda não salvo. Aguardando...")
        except Exception as e:
            break
        pyautogui.sleep(1)

def imagem_encontrada(image_path, confidence=0.9, max_attempts=5):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = os.path.join(current_dir, 'IMAGENS')
    image_path = os.path.join(caminho_imagem, image_path)
    for attempt in range(max_attempts):
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                print("Nota encontrada na tela.")
                return True
        except Exception as e:
            print("Nota não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)
        
def verificar_campo(image_name, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = os.path.join(current_dir, 'IMAGENS')
    image_path = os.path.join(caminho_imagem, image_name)    
    while True:
        found = False
        for i in range(5): 
            try:
                position = pyautogui.locateOnScreen(image_path, confidence=confidence)
                if position:
                    print("Campo preenchido.")
                    found = True
                    break
            except Exception as e:
                print(f"Erro ao procurar o campo: {e}")
            print("Campo não preenchido. Aguardando...")
            pyautogui.sleep(1)       
        if found:
            break
        else:
            click_image('salvar_filial.png')
            for i in range(5):
                pyautogui.press("backspace")
            pyautogui.write("1")
            pyautogui.sleep(1)   
            pyautogui.press("tab")
            for i in range(5):
                pyautogui.press("backspace")
            pyautogui.write("8")
            pyautogui.sleep(1)   
            pyautogui.press("tab")
            for i in range(5):
                pyautogui.press("backspace")            
            pyautogui.write("1")
            pyautogui.sleep(1)   
            pyautogui.press("tab")

def click(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path)
    try:
        position = pyautogui.locateOnScreen(image_path, confidence=confidence)
        if position:
            center_x = position.left + position.width // 2
            center_y = position.top + position.height // 2
            pyautogui.click(center_x, center_y)
            print("Imagem foi encontrada na tela.")
    except Exception as e:
        print("Imagem não encontrada na tela. Aguardando...")
    pyautogui.sleep(1)

def tela_cancelamento(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path)
    while True:
        click('ok_marcado.png')
        click('ok.png')
        click('ok2.png')
        pyautogui.sleep(2)
        click('faturamento.png')
        pyautogui.sleep(1)
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

def confirmacao_doc_incluido(image_path,image_path2,image_path3, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path)
    image_path2 = os.path.join(current_dir, caminho_imagem, image_path2)
    image_path3 = os.path.join(current_dir, caminho_imagem, image_path3)
    doc_invalido = False
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                print("Documento incluido.")
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        try:
            position3 = pyautogui.locateOnScreen(image_path3, confidence=confidence)
            if position3:
                doc_invalido = True
                status = f'Documento invalido:{cte}'
                try:
                    df = pd.read_excel(file_name)
                except FileNotFoundError:
                    # Criar um DataFrame com as colunas ROMANEIO e PLACA, caso o arquivo não exista
                    df = pd.DataFrame(columns=['DATA','ROMANEIO', 'PLACA', 'CIDADE', 'STATUS'])

                # Adicionar o número do romaneio e a placa à planilha
                new_row = pd.DataFrame({'DATA': [agora_formatado],'ROMANEIO': [romaneio], 'PLACA': [placa], 'CIDADE': [cidade_mais_longe] , 'COD MOTORISTA': [cod_motorista], 'KM': [km_veiculo], 'STATUS': [status]})
                df = pd.concat([df, new_row], ignore_index=True)
                # Salvar a planilha
                df.to_excel(file_name, index=False)
                print(f"Número do romaneio {romaneio} e placa {placa} salvos na planilha {file_name}.")
                # **Salvar os dados para reenvio de forma correta**
                dados_para_reenvio.append({
                    'PLACA': placa,
                    'NOME MOTORISTA': motorista_lancamento,
                    'KM': km_veiculo,
                    'CTES': ctes_por_cidade  # Salvando como lista de listas
                }) 
                click_image('ok4.png')      
                click_image('guia_geral.png')
                pyautogui.sleep(5)
                return doc_invalido
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        try:
            position2 = pyautogui.locateOnScreen(image_path2, confidence=confidence)
            if position2:
                tela_cancelamento('faturamento_movimentacao.png')#caso nao ache clicar no ok e em faturamento de novo
                pyautogui.sleep(1)
                click_image('faturamento_movimentacao_entregas.png')
                pyautogui.sleep(1)
                click_info('campo_cte.png')
                pyautogui.write('1')
                pyautogui.sleep(0.5)
                pyautogui.press('tab')
                pyautogui.write('2')
                pyautogui.sleep(0.5)
                pyautogui.press('tab')
                pyautogui.write(str(cte))   
                pyautogui.sleep(0.5)
                pyautogui.press('tab')
                click_image('atualizar_entregas.png')
                pyautogui.sleep(3)
                if imagem_encontrada('nota_encontrada.png'):
                    click_image('salvar_filial.png')
                    pyautogui.sleep(time)
                    pyautogui.write("1")
                    pyautogui.sleep(time)       
                    click_image('salvar_ocorrencia.png')
                    pyautogui.write("8")
                    pyautogui.sleep(time)      
                    click_image('salvar_observ.png')
                    pyautogui.write("1")
                    pyautogui.sleep(time)           
                    pyautogui.press("tab")
                    verificar_campo('campo_filial.png')
                    verificar_campo('campo_observacao.png')
                    verificar_campo('campo_ocorrencia.png')
                    click_image('efetuar_entregas.png')
                    finalizar_baixa('campo_cte_vazio.png')
                    pyautogui.sleep(2)
                    click_image('voltar.png')
                    pyautogui.sleep(5)
                    click_pasta_amarela('pastinha_amarela.png', 'pastinha_amarela_marcada.png')
                    pyautogui.sleep(2)
                    click_info('filial.png')
                    pyautogui.sleep(0.5)
                    pyautogui.write('1')
                    pyautogui.press('tab')
                    pyautogui.sleep(2) 
                    for i in range(3):
                        pyautogui.press('backspace')
                    for i in range(3):
                        pyautogui.press('del')
                    pyautogui.write('2')
                    pyautogui.press('tab')
                    pyautogui.sleep(2)
                    pyautogui.write(str(cte))
                    pyautogui.press('tab')
                    if a == 0:
                        click_image('yes.png')
                    pyautogui.sleep(2)
                    confirmacao_doc_incluido('verdinho_cinza.png','aviso_ja_entregue.png')
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)

def click_pasta_amarela(image_path,image_path2, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path)
    image_path2 = os.path.join(current_dir, caminho_imagem, image_path2)
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
        try:
            position2 = pyautogui.locateOnScreen(image_path2, confidence=confidence)
            if position2:
                center_x = position2.left + position2.width // 2
                center_y = position2.top + position2.height // 2
                pyautogui.click(center_x, center_y)
                print("Imagem foi encontrada na tela.")
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)

def click_info(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                center_x = position.left + position.width // 2
                center_y = position.top + position.height // 2
                pyautogui.moveTo(center_x, center_y)  # Movendo o cursor para a posição da imagem               
                pyautogui.moveRel(45, 0)  # Movendo o cursor para cima
                pyautogui.click()  # Clicando no local da imagem
                print("Imagem foi encontrada na tela.")
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)

def click_registros(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                center_x = position.left + position.width // 2
                center_y = position.top + position.height // 2
                pyautogui.moveTo(center_x, center_y)  # Movendo o cursor para a posição da imagem               
                pyautogui.moveRel(- 120, -60)  # Movendo o cursor para cima
                pyautogui.click(clicks=2)  # Clicando no local da imagem
                print("Imagem foi encontrada na tela.")
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)

def quantidade_registros():
    pyautogui.sleep(5)   
    click_registros('aguardando_liberacao.png')
    pyperclip.copy('')  # Limpa o conteúdo copiado
    pyautogui.sleep(2)
    pyautogui.click(button='right')
    pyautogui.sleep(1)
    click_image('copy.png')
    pyautogui.sleep(0.5)
    try:
        text = pyperclip.paste()  # Obtém o texto da área de transferência
        qtd_registros = int(text)  # Converte o texto para um número inteiro
        print("Quantidade de manifestos:", qtd_registros)
    except ValueError:
        print("O conteúdo copiado não é um número válido.")
        qtd_registros = 0  # Define como None se a conversão falhar
    except Exception as e:
        print("Ocorreu um erro:", str(e))
        qtd_registros = 0  # Define como None para qualquer outro erro
    return qtd_registros

def aviso_motorista(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                center_x = position.left + position.width // 2
                center_y = position.top + position.height // 2
                pyautogui.moveTo(center_x, center_y)  # Movendo o cursor para a posição da imagem               
                pyautogui.moveRel(45, 0)  # Movendo o cursor para cima
                pyautogui.click()  # Clicando no local da imagem
                print("Campo de pesquisa motorista encontrada na tela.")
                break
        except Exception as e:
            print("Campo de pesquisa motorista nao encontrado")
            click('cabecalho_romaneio.png')
            click('ok_marcado.png')
            click('ok.png')
            click('ok2.png')
            click_info('motorista_romaneio.png')
            pyautogui.press('F2')
            pyautogui.sleep(5)
        pyautogui.sleep(1)

def novo_evento(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                print("Confirmação para inclusao de evento.")
                break
        except Exception as e:
            print("Não autorizado inclusão de evento.Aguardando...")
        pyautogui.sleep(1)

# Função para verificar se o nome do motorista base está contido no nome do lançamento
def nomes_se_correspondem(nome1, nome2):
    partes_nome1 = nome1.lower().split()  # Motorista no lançamento
    partes_nome2 = nome2.lower().split()  # Motorista na base
    # Verifica se pelo menos um nome do motorista_base está presente no motorista_lancamento
    for parte in partes_nome2:
        if parte in partes_nome1:
            return True  # Encontrou uma correspondência
    return False  # Nenhuma correspondência encontrada

# Carregar a planilha
base_nomes = pd.read_excel("BaseMotoristas.xlsx")
file_name_base = 'BaseMotoristas.xlsx'

# Função simulada para lançar um único CTe (substitua com a automação real)
def lancar_cte(placa, cte, cidade):
    # Aqui você pode adicionar o código para lançar o CTe (como enviar para um sistema, preencher um formulário, etc.)
    print(f"Lançando o CTe '{cte}' para a placa {placa}, cidade: {cidade}")

# Função simulada para lançar um único CTe (substitua com a automação real)
def lancar_cte(placa, cte, cidade):
    # Aqui você pode adicionar o código para lançar o CTe (como enviar para um sistema, preencher um formulário, etc.)
    print(f"Lançando o CTe '{cte}' para a placa {placa} cidade {cidade}")

# Função para verificar se uma célula contém um formato de placa
def is_placa(value):
    if pd.isna(value):
        return False
    value = value.replace(' ', '')
    placa_pattern = r'^[A-Z]{3} \d{4}$|^[A-Z]{3} \d{1}[A-Z]{1}\d{2}$|^[A-Z]{3} \d{2}[A-Z]$|^[A-Z]{3}-\d{4}$|^[A-Z]{3}\d{1}[A-Z]{1}\d{2}$|^[A-Z]{3}-\d{1}[A-Z]{1}\d{2}$|^[A-Z]{3}\d{4}$|^[A-Z]{3} \d{4}$'
    return bool(re.match(placa_pattern, str(value)))

Planilha_romaneio = pd.read_excel("planilha_romaneio.xlsx")
# Filtrar as linhas onde o valor da coluna "ROMANEIO" não seja "AVISO DE MOTORISTA" ou "VEICULO"
Planilha_romaneio = Planilha_romaneio[~Planilha_romaneio['ROMANEIO'].isin(['AVISO DE MOTORISTA OU VEICULO'])]

# Exemplo de salvar a planilha atualizada (opcional)
Planilha_romaneio.to_excel('planilha_romaneio.xlsx', index=False)

# Função para identificar cidades e placas
def processar_veiculo_ou_cidade(value):
    if is_placa(value):
        return value
    return None

def processar_cidade(value):
    if not is_placa(value) and pd.notna(value) and str(value).strip() != '':
        return value
    return None

# Definir a função para verificar se o valor é apenas numérico ou vazio
def is_numeric_or_empty(value):
    # Verifica se o valor é NaN, uma string vazia, ou contém apenas números
    if pd.isna(value) or (isinstance(value, str) and value.strip() == ''):
        return True
    if isinstance(value, (int, float)):
        return True
    elif isinstance(value, str) and value.isdigit():
        return True
    return False

# Função para exibir uma mensagem de sucesso
def show_success_message(msg):
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal
    messagebox.showinfo("Lancamento Romaneio", msg)  # Exibe o título e a mensagem
    root.destroy()  # Fecha a janela após exibir a mensagem

# Carregar a planilha
Planilha_cc19 = pd.read_excel("Base.xlsx")

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

    # Itera sobre os índices e verifica o valor da linha abaixo de cada "Veículo"
    for idx in index_veiculo:
        if idx + 1 < len(Planilha_cc19):
            valor_abaixo = Planilha_cc19.loc[idx + 1, 'Veículo']

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

# Ajustar a coluna [Depósito] para as linhas onde [Depósito] está vazio e [Veículo] está preenchido
Planilha_cc19.loc[Planilha_cc19['Depósito'].isna() & Planilha_cc19['Veículo'].notna(), 'Depósito'] = 'EMPRESA'

# Aplicar a função para filtrar linhas
Planilha_cc19 = Planilha_cc19[~Planilha_cc19['Depósito'].apply(is_numeric_or_empty)]

# Colunas que você quer manter
colunas_para_manter_cc19 = ['Veículo', 'Custo Quantidade Atividades']

# Verificar se as colunas existem e, se não existirem ou estiverem vazias, criar colunas vazias
for coluna in colunas_para_manter_cc19:
    if coluna not in Planilha_cc19.columns or Planilha_cc19[coluna].isna().all():
        Planilha_cc19[coluna] = pd.NA  # ou np.nan para criar uma coluna vazia

# Processar cada linha para preencher cidades e placas
Planilha_cc19['Cidade'] = Planilha_cc19['Veículo'].apply(processar_cidade)
Planilha_cc19['Placa'] = Planilha_cc19['Veículo'].apply(processar_veiculo_ou_cidade)

# Preencher as placas nas linhas onde a cidade está definida
Planilha_cc19['Placa'] = Planilha_cc19['Placa'].fillna(method='ffill')

# Remover a coluna original 'Veículo'
Planilha_cc19.drop(columns=['Veículo'], inplace=True)

# Criar a nova coluna 'Motorista' com valor NaN
Planilha_cc19['Motorista'] = pd.NA
# Preencher a coluna 'KM' com NaN
Planilha_cc19['KM'] = pd.NA

if not Planilha_cc19.empty:
    # Encontrar os índices onde 'Depósito' tem valor 'Bairro'
    bairro_indices = Planilha_cc19[Planilha_cc19['Depósito'].isin(['Depósito'])].index
    # Preencher a coluna 'Motorista' da linha correspondente com o valor da linha acima
    for idx in bairro_indices:
        if idx >= 0:
            motorista_lancamento = Planilha_cc19.loc[idx + 1, 'Depósito']
            Planilha_cc19.loc[idx, 'Motorista'] = motorista_lancamento
    
    # Preencher as placas nas linhas onde a cidade está definida
    Planilha_cc19['Motorista'] = Planilha_cc19['Motorista'].fillna(method='ffill')

    # Encontrar os índices onde 'Depósito' tem valor 'Bairro'
    bairro_indices = Planilha_cc19[Planilha_cc19['Depósito'].isin(['Depósito'])].index
    
    # Preencher a coluna 'KM' da linha correspondente com o valor da linha acima
    for idx in bairro_indices:
        if idx > 0:
            Planilha_cc19.loc[idx, 'KM'] = Planilha_cc19.loc[idx +11, 'Distância (km)']
    
    # Propagar os valores para baixo na coluna 'KM'
    Planilha_cc19['KM'] = Planilha_cc19['KM'].fillna(method='ffill')

# Filtrar as linhas onde a coluna 'Cidade' não é vazia
Planilha_cc19 = Planilha_cc19.dropna(subset=['Cidade'])
Planilha_cc19['KM'] = Planilha_cc19['KM'].fillna(1)
# Garantir que os valores sejam strings antes de dividir
Planilha_cc19['Custo Quantidade Atividades'] = Planilha_cc19['Custo Quantidade Atividades'].astype(str)

# Remover as linhas onde o valor da coluna é 'Custo Quantidade Atividades'
Planilha_cc19 = Planilha_cc19[Planilha_cc19['Custo Quantidade Atividades'] != 'Custo Quantidade Atividades']
Planilha_cc19 = Planilha_cc19[Planilha_cc19['Custo Quantidade Atividades'] != 'Pedido']
Planilha_cc19 = Planilha_cc19[Planilha_cc19['Custo Quantidade Atividades'] != '-']

# Agrupar por placa e cidade, e dividir CTes em lista
agrupado_por_placa_cc19 = Planilha_cc19.groupby(['Placa', 'Cidade', 'Motorista', 'KM' ])['Custo Quantidade Atividades'].apply(lambda x: ', '.join(x).split(', ')).reset_index()

# Verificar se algum dos DataFrames está vazio
if len(agrupado_por_placa_cc19) == 0:
    print("O DataFrame esta vazio. Parando a execução do código.")
    sys.exit()  # Para a execução do código

# Carregar a planilha
Planilha_cidades = pd.read_excel("cidade km - smr.xlsx")
Planilha_reenvio = pd.read_excel("reenvio_planilha.xlsx")
# Nome do arquivo Excel
file_name = 'planilha_romaneio.xlsx'

dados_para_reenvio = []

# Agrupar os dados
agrupado_combinado = pd.concat([
        agrupado_por_placa_cc19.rename(columns={
        'Placa': 'PLACA',
        'Cidade': 'CIDADE',
        'Custo Quantidade Atividades': 'CTE / OST',
        'Motorista': 'NOME MOTORISTA'
    })
])

# Agrupar CTes por placa e incluir o nome do motorista
agrupado_por_placa = agrupado_combinado.groupby(['PLACA', 'NOME MOTORISTA', 'KM']).apply(
    lambda x: x[['CIDADE', 'CTE / OST']].values.tolist()
).reset_index(name='CTES')

def selecionar_planilha():
    escolha = var_opcao.get()  # Captura a escolha do usuário
    root.quit()  # Fecha a janela após a escolha

# Configuração da janela principal
root = tk.Tk()
root.title("Seleção de Planilha")

# Variável para armazenar a escolha do usuário
var_opcao = tk.IntVar()

# Rótulo para instrução
label = tk.Label(root, text="Selecione a planilha que deseja executar:")
label.pack(pady=10)

# Botões de opção (Planilha Base ou Planilha Reenvio)
radio1 = tk.Radiobutton(root, text="Planilha Base", variable=var_opcao, value=1)
radio1.pack(anchor="w", padx=20)

radio2 = tk.Radiobutton(root, text="Planilha Reenvio", variable=var_opcao, value=2)
radio2.pack(anchor="w", padx=20)

# Botão para confirmar a escolha
botao_confirmar = tk.Button(root, text="Confirmar", command=selecionar_planilha)
botao_confirmar.pack(pady=10)

# Iniciar a janela tkinter
root.mainloop()

# Usar a escolha do usuário para definir qual planilha carregar
if var_opcao.get() == 1:
    print("Carregar Planilha Base")
    # Planilha_base = pd.read_excel('caminho/para/planilha_base.xlsx')
    Planilha_romaneio = agrupado_por_placa
elif var_opcao.get() == 2:
    print("Carregar Planilha Reenvio")
    # Planilha_reenvio = pd.read_excel('caminho/para/planilha_reenvio.xlsx')
    Planilha_romaneio = Planilha_reenvio
else:
    print("Nenhuma planilha selecionada.")

root.destroy()

pyautogui.sleep(2)
b = 0
c = False

iniciar_robo('faturamento.png')
pyautogui.sleep(5)
click_image('faturamento.png')
pyautogui.sleep(0.5)
click_image('faturamento_movimentacao.png')
pyautogui.sleep(0.5)
click_image('faturamento_movimentacao_romaneio.png')
pyautogui.sleep(0.5)
click_image('faturamento_movimentacao_romaneio_carregamento.png')
pyautogui.sleep(0.5)

for _, row in Planilha_romaneio.iterrows():
    hora_atual = datetime.now().time()

    if hora_atual.strftime("%H:%M") == "22:45":
        print("Já são 22:55, esperar 15 minutos")
        click_image('sair_sistema.png')
        click_image('sair_sistema_logoff.png')
        click_image('sair_sistema_logoff_sair.png')
        
        # Simulando o final do código
        if __name__ == "__main__":  # Corrigido o nome da condição

            # Aqui você pode adicionar seu código que será executado
            show_success_message("Pausado por ser 23 horas, aguardando 15 minutos para continuar")
            pyautogui.sleep(900)
            show_success_message("Liberado Logar no sistema, clique no ok quando abrir na tela inicial!")
    else:
        print("Ainda não são 22:55.")

    placa_lancamento = row['PLACA']
    motorista_lancamento = row['NOME MOTORISTA']
    km_veiculo = row['KM']
    ctes_por_cidade = row['CTES']
    if isinstance(ctes_por_cidade, str):
        # Convertendo string formatada em lista
        ctes_por_cidade = eval(ctes_por_cidade)  # Ou outra forma segura de conversão
    info_faltantes = False
    motorista_lancamento = motorista_lancamento.rstrip()
    if 'Ç' in motorista_lancamento:
        motorista_lancamento = motorista_lancamento.replace("Ç", "C")
    elif('ç' in motorista_lancamento):
        motorista_lancamento = motorista_lancamento.replace("ç", "c")
    motorista_lancamento = unidecode(motorista_lancamento)
    motorista_lancamento = ' '.join(motorista_lancamento.split())
    placa_lancamento = str(placa_lancamento)
    placa_lancamento = placa_lancamento.translate(str.maketrans('', '', '*-_& '))
    km = 0
    # Criar um dicionário para agrupar CTes por cidade
    ctes_por_cidade_dict = {}
    for cidade, ctes in ctes_por_cidade:
        if isinstance(ctes, str):
            # Se ctes for uma string, divida em uma lista
            ctes = [cte.strip() for cte in ctes.split(',') if cte.strip()]
        elif isinstance(ctes, list):
            # Se ctes já é uma lista, apenas limpe os espaços
            ctes = [cte.strip() for cte in ctes if cte.strip()]
        else:
            # Caso ctes não seja nem string nem lista, ignore
            continue
        if cidade not in ctes_por_cidade_dict:
            ctes_por_cidade_dict[cidade] = []
        ctes_por_cidade_dict[cidade].extend(ctes)
        cidade = cidade.upper()
        
        if cidade in Planilha_cidades['Cidade Destino'].values:
            base_km = Planilha_cidades.loc[Planilha_cidades[Planilha_cidades['Cidade Destino'] == cidade].index.values, 'KM (origem Sumaré)'].values[0]
        if base_km > km:
            km = base_km
            cidade_mais_longe = cidade
        
    print(f'motorista:{motorista_lancamento} placa: {placa_lancamento} cidade mais longe:{cidade_mais_longe} km:{km_veiculo}')
    
    for cidade, ctes in ctes_por_cidade_dict.items():
        # Usar set() para remover duplicatas
        ctes_unicos = set(ctes)
        for cte in ctes_unicos:
            if cte:  # Verificar se não está vazio
                # Lançar cada CTe único para a placa e cidade
                try:
                    cte = int(float(cte))
                except:
                    print('Faltando informações')
                    info_faltantes = True
                    break

    if info_faltantes ==True:
        romaneio = 'INFORMAÇÕES FALTANTES'
        # Verificar se o arquivo já existe
        try:
            df = pd.read_excel(file_name)
        except FileNotFoundError:
            # Criar um DataFrame com as colunas ROMANEIO e PLACA, caso o arquivo não exista
            df = pd.DataFrame(columns=['DATA', 'ROMANEIO', 'PLACA', 'CIDADE'])
        # Adicionar o número do romaneio e a placa à planilha
        new_row = pd.DataFrame({'DATA': [agora_formatado], 'ROMANEIO': [romaneio], 'PLACA': [placa_lancamento], 'CIDADE': [cidade_mais_longe]})
        df = pd.concat([df, new_row], ignore_index=True)
        # Salvar a planilha
        df.to_excel(file_name, index=False)
        print(f"Número do romaneio {romaneio} e placa {placa_lancamento} salvos na planilha {file_name}.")
        continue  # Pula para a próxima iteração do for

    #COPIAR A PLACA TAMBEM E SALVAR
    click_image('incluir.png')
    pyautogui.sleep(2)
    novo_evento('campo_romaneio_automatico.png')
    click_info('setor.png')
    pyautogui.write('2')
    pyautogui.press('tab')
    click_info('campo_veiculo.png')
    pyautogui.press('F2')
    click_info('placa.png')
    pyautogui.write(str(placa_lancamento))
    pyautogui.sleep(2)
    pyautogui.press('tab')
    click_info('situacao_veiculo.png')
    pyautogui.sleep(1)
    click_image('situacao_veiculo_normal.png')
    pyautogui.sleep(1)
    click_image('atualizar.png')
    pyautogui.sleep(2)
    click_image('selecionar.png')
    pyautogui.sleep(2)
    pyautogui.press('tab')
    pyautogui.sleep(4)
    click('ok_marcado.png')
    click('ok.png')
    click('ok2.png')
    click_image('cabecalho_romaneio.png')
    click_info('campo_veiculo.png')

    pyperclip.copy('')  # Limpa o conteúdo copiado
    pyautogui.sleep(2)
    pyautogui.click(button='right')
    pyautogui.sleep(1)
    click_image('copy.png')
    pyautogui.sleep(0.5)
    try:
        text = pyperclip.paste()
        placa = str(text)
        print("Codigo da Placa:", placa)   
    except ValueError:
        print("O conteúdo copiado não é um número válido.")
    except Exception as e:
        print("Ocorreu um erro:", str(e))
    if placa == '':
        placa = placa_lancamento

    pyperclip.copy('')  # Limpa o conteúdo copiado
    click_info('motorista_romaneio.png')
    pyautogui.sleep(2)
    pyautogui.click(button='right')
    pyautogui.sleep(1)
    click_image('copy.png')
    pyautogui.sleep(0.5)
    try:
        text = pyperclip.paste()
        cod_motorista_lanc = int(text)
        print("Número do cod_motorista:", cod_motorista_lanc)
    except ValueError:
        print("O conteúdo copiado não é um número válido.")
        cod_motorista_lanc = 'Erro ao copiar cod_motorista'
    except Exception as e:
        print("Ocorreu um erro:", str(e))
    motorista_encontrado = False
    
    # Verifica se a placa está presente em base_nomes
    if placa in base_nomes['PLACA'].values:
        motorista_base = base_nomes.loc[base_nomes['PLACA'] == placa]
        for idx, row in motorista_base.iterrows():
            mot = row['MOTORISTA']
            cod_motorista_base = row['COD']
            # Comparando todos os nomes principais (case-insensitive)
            if nomes_se_correspondem(motorista_lancamento, mot):
                print(f"Os nomes são correspondentes: {motorista_lancamento} e {mot}")
                cod_motorista = cod_motorista_base
                print(f"Código do motorista encontrado: {cod_motorista}")
                motorista_encontrado = True
                if cod_motorista != cod_motorista_lanc:
                    click_image('cabecalho_romaneio.png')
                    print('codigo diferente alterar')
                    click_info('motorista_romaneio.png')
                    pyautogui.sleep(2)
                    for i in range(10):
                        pyautogui.press('backspace')
                    for i in range(10):
                        pyautogui.press('del')
                    pyautogui.write(str(cod_motorista))
                    pyautogui.sleep(2)
                    pyautogui.press('tab')
                    pyautogui.sleep(2)
                else: 
                    print('nao precisa alterar o codigo')
                break
            else:
                print(f"Os nomes são diferentes: {motorista_lancamento} e {mot}")

    if not motorista_encontrado:
        print(f"Motorista não encontrado para a placa {placa}. Adicionando novo motorista.")
        pyautogui.press('F2')
        pyautogui.sleep(5)        
        aviso_motorista('motorista_romaneio_nome.png')
        pyautogui.sleep(5)
        pyautogui.write(str(motorista_lancamento))
        pyautogui.sleep(5)
        click_image('atualizar.png')
        pyautogui.sleep(5)

        click_image('aguardando_liberacao.png')
        pyautogui.moveRel(0, - 100)  # Movendo o cursor para cima
        pyautogui.click()  # Clicando no local da imagem
        pyautogui.sleep(1)
        for i in range(10):
            pyautogui.press('down')
        pyautogui.sleep(5)

        qtd_registros = quantidade_registros()
        qtd_registros = int(qtd_registros)
        print(f'quantidade:{qtd_registros}')
        if qtd_registros > 1:
            romaneio = 'AJUSTAR NOME DO MOTORISTA'
            # Verificar se o arquivo já existe
            try:
                df = pd.read_excel(file_name)
            except FileNotFoundError:
                # Criar um DataFrame com as colunas ROMANEIO e PLACA, caso o arquivo não exista
                df = pd.DataFrame(columns=['DATA', 'ROMANEIO', 'PLACA', 'CIDADE'])
            # Adicionar o número do romaneio e a placa à planilha
            new_row = pd.DataFrame({'DATA': [agora_formatado], 'ROMANEIO': [romaneio], 'PLACA': [placa], 'CIDADE': [cidade_mais_longe]})
            df = pd.concat([df, new_row], ignore_index=True)
            # Salvar a planilha
            df.to_excel(file_name, index=False)
            print(f"Número do romaneio {romaneio} e placa {placa} salvos na planilha {file_name}.")
            # **Salvar os dados para reenvio de forma correta**
            dados_para_reenvio.append({
                'PLACA': placa,
                'NOME MOTORISTA': motorista_lancamento,
                'KM': km_veiculo,
                'CTES': ctes_por_cidade  # Salvando como lista de listas
            })
            continue  # Pula para a próxima iteração do for
            
        click_image('selecionar.png')
        pyautogui.sleep(2)
        pyautogui.press('tab')
        pyautogui.sleep(4)
        print('antes de ir pra tela de pesquisa')
        click('ok_marcado.png')
        click('ok.png')
        click('ok2.png')

        pyperclip.copy('')  # Limpa o conteúdo copiado
        click_info('motorista_romaneio.png')
        pyautogui.sleep(2)
        pyautogui.click(button='right')
        pyautogui.sleep(1)
        click_image('copy.png')
        pyautogui.sleep(0.5)
        try:
            text = pyperclip.paste()
            cod_motorista_cadastro = int(text)
            print("Número do cod_motorista:", cod_motorista_cadastro)
        except ValueError:
            print("O conteúdo copiado não é um número válido.")
            cod_motorista_cadastro = 'Erro ao copiar cod_motorista'
        except Exception as e:
            print("Ocorreu um erro:", str(e))
        # Verificar se o arquivo já existe
        try:
            df = pd.read_excel(file_name_base)
        except FileNotFoundError:
            # Criar um DataFrame com as colunas ROMANEIO e PLACA, caso o arquivo não exista
            df = pd.DataFrame(columns=['PLACA','MOTORISTA', 'COD'])

        if cod_motorista_cadastro != 'Erro ao copiar cod_motorista':
            # Adicionar o número do romaneio e a placa à planilha
            new_row = pd.DataFrame({'PLACA': [placa],'MOTORISTA': [motorista_lancamento], 'COD': [cod_motorista_cadastro]})
            df = pd.concat([df, new_row], ignore_index=True)
            
            # Salvar a planilha
            df.to_excel(file_name_base, index=False)
            print(f"Novo motorista adicionado: {motorista_lancamento}, Placa: {placa}, Código: {cod_motorista_cadastro}")
            cod_motorista = cod_motorista_cadastro
   
    aviso = click_image_salvar('salvar.png','ok_marcado.png', 'campo_romaneio_automatico.png','aviso.png')

    if aviso:
        print(f"Aviso encontrado. Pulando para o próximo.")
        romaneio = 'AVISO DE MOTORISTA OU VEICULO'
        # Verificar se o arquivo já existe
        try:
            df = pd.read_excel(file_name)
        except FileNotFoundError:
            # Criar um DataFrame com as colunas ROMANEIO e PLACA, caso o arquivo não exista
            df = pd.DataFrame(columns=['DATA', 'ROMANEIO', 'PLACA', 'CIDADE'])
        # Adicionar o número do romaneio e a placa à planilha
        new_row = pd.DataFrame({'DATA': [agora_formatado], 'ROMANEIO': [romaneio], 'PLACA': [placa], 'CIDADE': [cidade_mais_longe]})
        df = pd.concat([df, new_row], ignore_index=True)
        # Salvar a planilha
        df.to_excel(file_name, index=False)
        print(f"Número do romaneio {romaneio} e placa {placa} salvos na planilha {file_name}.")
        # **Salvar os dados para reenvio de forma correta**
        dados_para_reenvio.append({
            'PLACA': placa,
            'NOME MOTORISTA': motorista_lancamento,
            'KM': km_veiculo,
            'CTES': ctes_por_cidade  # Salvando como lista de listas
        })
        continue  # Pula para a próxima iteração do for

    #COPIAR CODIGO ROMANEIO E SALVAR NA PLANILHA
    pyperclip.copy('')  # Limpa o conteúdo copiado
    pyautogui.sleep(2)
    click_info('codigo_romaneio.png')
    pyautogui.sleep(2)
    pyautogui.click(button='right')
    pyautogui.sleep(1)
    click_image('copy.png')
    pyautogui.sleep(0.5)
    try:
        text = pyperclip.paste()
        romaneio = int(text)
        print("Número do ROMANEIO:", romaneio)
    except ValueError:
        print("O conteúdo copiado não é um número válido.")
        romaneio = 'Erro ao copiar codigo do ROMANEIO'
    except Exception as e:
        print("Ocorreu um erro:", str(e))

    click_image('guia_itens.png')
    pyautogui.sleep(2)

    a = 0
    for cidade, ctes in ctes_por_cidade_dict.items():
        # Usar set() para remover duplicatas
        ctes_unicos = set(ctes)
        for cte in ctes_unicos:
            if cte:  # Verificar se não está vazio
                # Lançar cada CTe único para a placa e cidade
                cte = int(float(cte))
                lancar_cte(placa, cte, cidade)
                click_pasta_amarela('pastinha_amarela.png', 'pastinha_amarela_marcada.png')
                pyautogui.sleep(2)
                if b == 0:
                    print('selecionar CTRC')
                    click_info('escolher_tipo_documento.png')
                    click_image('tipo_ctrc.png')
                    c = False
                    b +=1
                if c == True:
                    click_info('escolher_tipo_documento.png')
                    click_image('tipo_ctrc.png')
                    c = False
                if len(str(cte)) == 5:
                    click_info('escolher_tipo_documento.png')
                    click_image('tipo_ost.png')
                    c = True

                pyautogui.sleep(0.5)
                click_info('filial.png')
                pyautogui.sleep(0.5)
                pyautogui.write('1')
                pyautogui.sleep(0.5)
                pyautogui.press('tab')
                click_info('serie.png')
                pyautogui.sleep(2) 
                for i in range(3):
                    pyautogui.press('backspace')
                for i in range(3):
                    pyautogui.press('del')
                if c == True:
                    pyautogui.write('U')
                else:
                    pyautogui.write('2')
                pyautogui.sleep(0.5)
                pyautogui.press('tab')
                click_info('numero_doc.png')
                pyautogui.sleep(2)
                pyautogui.write(str(cte))
                pyautogui.sleep(0.5)
                pyautogui.press('tab')
                if a == 0:
                    click_image('yes.png')
                pyautogui.sleep(2)
                doc_invalido = confirmacao_doc_incluido('verdinho_cinza.png', 'aviso_ja_entregue.png', 'documento_situacao_invalida.png',)
                if doc_invalido:
                    print(f"Documento inválido para a placa {placa}. Pulando para a próxima.")
                    break  # Isso interrompe o loop interno, voltando para o loop externo
                pyautogui.sleep(1)
                a += 1
    if not doc_invalido:
        # Verificar se o arquivo já existe
        try:
            df = pd.read_excel(file_name)
        except FileNotFoundError:
            # Criar um DataFrame com as colunas ROMANEIO e PLACA, caso o arquivo não exista
            df = pd.DataFrame(columns=['DATA','ROMANEIO', 'PLACA', 'CIDADE'])

        # Adicionar o número do romaneio e a placa à planilha
        new_row = pd.DataFrame({'DATA': [agora_formatado],'ROMANEIO': [romaneio], 'PLACA': [placa], 'CIDADE': [cidade_mais_longe] , 'COD MOTORISTA': [cod_motorista], 'KM': [km_veiculo]})
        df = pd.concat([df, new_row], ignore_index=True)
        
        # Salvar a planilha
        df.to_excel(file_name, index=False)
        print(f"Número do romaneio {romaneio} e placa {placa} salvos na planilha {file_name}.")

        pyautogui.sleep(3)
        click_image('guia_geral.png')
        pyautogui.sleep(5)
        click_image('efetuar.png')
        click_image('no.png')
        click_image('ok.png')
        
    else:
        pyautogui.sleep(3)
        click_image('guia_geral.png')
        pyautogui.sleep(5)
        click_image('incluir.png')
        pyautogui.sleep(2)
        click_image('guia_geral.png')
        pyautogui.sleep(2)
    print("")
    pyautogui.sleep(5)

#Escrever o DataFrame vazio na planilha
Planilha_reenvio = pd.DataFrame(columns=Planilha_reenvio.columns)
Planilha_reenvio.to_excel('reenvio_planilha.xlsx', index=False)

# Se houver dados para reenvio, salvar em uma nova planilha
if dados_para_reenvio:
    # **Salvar a planilha de reenvio com os dados formatados corretamente**
    df_reenvio = pd.DataFrame(dados_para_reenvio)
    df_reenvio.to_excel('reenvio_planilha.xlsx', index=False)
    print("Dados para reenvio salvos com sucesso.")
else:
    print("Nenhuma condição atendida para reenvio.")
print("Automatização concluída com sucesso!")

click_image('voltar.png')

# Carrega a planilha Excel
arquivo = 'Base.xlsx'
df = pd.read_excel(arquivo)

# Mantém apenas a primeira linha
df_limpo = df.iloc[:0]

# Salva o resultado de volta no arquivo Excel
df_limpo.to_excel(arquivo, index=False)

show_success_message("Robô finalizado com sucesso")

