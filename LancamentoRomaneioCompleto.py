import re
import pyautogui
import os
import pyperclip
import pandas as pd
from datetime import datetime
from datetime import datetime, timedelta
import sys
from unidecode import unidecode

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
    while True:
        try:
            # Procurar pela imagem 4
            position4 = pyautogui.locateOnScreen(image_path4, confidence=confidence)
            if position4:
                click_ok('ok_marcado.png','ok.png','ok2.png')
                pyautogui.sleep(1)
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

def confirmacao_doc_incluido(image_path,image_path2, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path)
    image_path2 = os.path.join(current_dir, caminho_imagem, image_path2)
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                print("Documento incluido.")
                break
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
                if imagem_encontrada('nota_encontrada.png'):
                    click_image('salvar_filial.png')
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
                    click_image('yes2.png')
                    click_image('ok.png')
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
                break
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
    placa_pattern = r'^[A-Z]{3} \d{4}$|^[A-Z]{3} \d{1}[A-Z]{1}\d{2}$|^[A-Z]{3} \d{2}[A-Z]$|^[A-Z]{3}-\d{4}$|^[A-Z]{3}\d{1}[A-Z]{1}\d{2}$|^[A-Z]{3}-\d{1}[A-Z]{1}\d{2}$'
    return bool(re.match(placa_pattern, str(value)))

# Carregar a planilha
Planilha_cc19 = pd.read_excel("Pasta2.xlsx")
#print(len(Planilha_cc19))

# Função para identificar cidades e placas
def processar_veiculo_ou_cidade(value):
    if is_placa(value):
        return value
    return None

def processar_cidade(value):
    if not is_placa(value) and pd.notna(value) and str(value).strip() != '':
        return value
    return None

# Carregar a planilha
Planilha_cc15 = pd.read_excel("Pasta1.xlsx")

colunas_para_manter_cc15 = ['CTE / OST', 'CNPJ DESTINATARIO', 'CIDADE']

# Verificar se as colunas existem e, se não existirem ou estiverem vazias, criar colunas vazias
for coluna in colunas_para_manter_cc15:
    if coluna not in Planilha_cc15.columns or Planilha_cc15[coluna].isna().all():
        Planilha_cc15[coluna] = pd.NA  # ou np.nan para criar uma coluna vazia

# Criar uma nova coluna para o nome do motorista
Planilha_cc15['NOME MOTORISTA'] = Planilha_cc15['CNPJ DESTINATARIO'].apply(
    lambda x: x if isinstance(x, str) and x[0].isalpha() else ''
)

# Atualizar a coluna 'CNPJ DESTINATARIO' para deixar apenas os CNPJs
Planilha_cc15['CNPJ DESTINATARIO'] = Planilha_cc15['CNPJ DESTINATARIO'].apply(
    lambda x: x if isinstance(x, str) and x[0].isdigit() else ''
)

# Excluir as linhas onde 'NOME MOTORISTA' começa com 'AGENDAMENTO'
Planilha_cc15 = Planilha_cc15[~Planilha_cc15['NOME MOTORISTA'].str.startswith('AGENDAMENTO')]

# Preencher os valores acima da próxima linha preenchida
Planilha_cc15['NOME MOTORISTA'] = Planilha_cc15['NOME MOTORISTA'].replace('', pd.NA).bfill()

# Excluir as linhas onde 'CTE / OST' é NaN
Planilha_cc15 = Planilha_cc15.dropna(subset=['CTE / OST'])

# Criar uma nova coluna chamada 'PLACA' para armazen
# Criar uma nova coluna chamada 'PLACA' para armazenar os últimos 7 caracteres de 'NOME MOTORISTA'
Planilha_cc15['PLACA'] = Planilha_cc15['NOME MOTORISTA'].apply(lambda x: x[-7:] if isinstance(x, str) else '')

# Remover os últimos 7 caracteres da coluna 'NOME MOTORISTA'
Planilha_cc15['NOME MOTORISTA'] = Planilha_cc15['NOME MOTORISTA'].apply(lambda x: x[:-7] if isinstance(x, str) else x)

# Garantir que os valores na coluna 'CTE / OST' sejam strings
Planilha_cc15['CTE / OST'] = Planilha_cc15['CTE / OST'].astype(str)

Planilha_cc15 = Planilha_cc15[Planilha_cc15['CTE / OST'] != 'CTE / OST']

# Criar a nova coluna 'Motorista' com valor NaN
Planilha_cc15['KM'] = pd.NA
Planilha_cc15['KM'] = 1

# Agrupar por placa e cidade e combinar CTes em lista
agrupado_por_placa_cidade_cc15 = Planilha_cc15.groupby(['PLACA', 'CIDADE', 'NOME MOTORISTA','KM' ])['CTE / OST'].apply(lambda x: ', '.join(x).split(', ')).reset_index()

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
    bairro_indices = Planilha_cc19[Planilha_cc19['Depósito'] == 'Bairro'].index

    # Preencher a coluna 'Motorista' da linha correspondente com o valor da linha acima
    for idx in bairro_indices:
        if idx > 0:
            motorista = Planilha_cc19.loc[idx - 1, 'Depósito']
            Planilha_cc19.loc[idx, 'Motorista'] = motorista


    # Preencher as placas nas linhas onde a cidade está definida
    Planilha_cc19['Motorista'] = Planilha_cc19['Motorista'].fillna(method='ffill')

    # Encontrar os índices onde 'Depósito' tem valor 'Bairro'
    bairro_indices = Planilha_cc19[Planilha_cc19['Depósito'] == 'Bairro'].index

    # Preencher a coluna 'KM' da linha correspondente com o valor da linha acima
    for idx in bairro_indices:
        if idx > 0:
            Planilha_cc19.loc[idx, 'KM'] = Planilha_cc19.loc[idx - 1, 'Distância (km)']

    # Propagar os valores para baixo na coluna 'KM'
    Planilha_cc19['KM'] = Planilha_cc19['KM'].fillna(method='ffill')

# Filtrar as linhas onde a coluna 'Cidade' não é vazia
Planilha_cc19 = Planilha_cc19.dropna(subset=['Cidade'])

# Garantir que os valores sejam strings antes de dividir
Planilha_cc19['Custo Quantidade Atividades'] = Planilha_cc19['Custo Quantidade Atividades'].astype(str)

# Remover as linhas onde o valor da coluna é 'Custo Quantidade Atividades'
Planilha_cc19 = Planilha_cc19[Planilha_cc19['Custo Quantidade Atividades'] != 'Custo Quantidade Atividades']
Planilha_cc19 = Planilha_cc19[Planilha_cc19['Custo Quantidade Atividades'] != 'Pedido']
Planilha_cc19 = Planilha_cc19[Planilha_cc19['Custo Quantidade Atividades'] != '-']

# Agrupar por placa e cidade, e dividir CTes em lista
agrupado_por_placa_cc19 = Planilha_cc19.groupby(['Placa', 'Cidade', 'Motorista', 'KM' ])['Custo Quantidade Atividades'].apply(lambda x: ', '.join(x).split(', ')).reset_index()

# Verificar se algum dos DataFrames está vazio
if len(agrupado_por_placa_cc19) == 0 and len(agrupado_por_placa_cidade_cc15) == 0:
    print("Os dois DataFrames estam vazios. Parando a execução do código.")
    sys.exit()  # Para a execução do código

# Carregar a planilha
Planilha_cidades = pd.read_excel("cidade km - smr.xlsx")

# Nome do arquivo Excel
file_name = 'planilha_romaneio.xlsx'

# Agrupar os dados
agrupado_combinado = pd.concat([
    agrupado_por_placa_cidade_cc15,
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

#print(agrupado_por_placa)

# Automatizar o lançamento de cada CTe individualmente
for _, row in agrupado_por_placa.iterrows():
    placa = row['PLACA']
    motorista = row['NOME MOTORISTA']
    km_veiculo = row['KM']
    ctes_por_cidade = row['CTES']
    motorista = motorista.rstrip()
    if 'Ç' in motorista:
        motorista = motorista.replace("Ç", "C")
    elif('ç' in motorista):
        motorista = motorista.replace("ç", "c")
    motorista = unidecode(motorista)
    motorista = ' '.join(motorista.split())
    placa = str(placa)
    placa = placa.replace('*', '')
    placa = placa.replace('-', '')
    placa = placa.replace('_', '')
    placa = placa.replace('&', '')
    placa = placa.replace(' ', '')
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

        if cidade in Planilha_cidades['Cidade Destino'].values:
            base_km = Planilha_cidades.loc[Planilha_cidades[Planilha_cidades['Cidade Destino'] == cidade].index.values, 'KM (origem Sumaré)'].values[0]

        #print(f'cidade {cidade} km:{base_km}')
        if base_km > km:
            km = base_km
            cidade_mais_longe = cidade
        
    print(f'motorista:{motorista} placa: {placa} cidade mais longe:{cidade_mais_longe} km:{km_veiculo}')
    
    click_image('incluir.png')
    pyautogui.sleep(2)
    novo_evento('campo_romaneio_automatico.png')
    click_info('setor.png')
    pyautogui.write('2')
    pyautogui.press('tab')
    click_info('campo_veiculo.png')
    pyautogui.press('F2')
    click_info('placa.png')
    pyautogui.write(str(placa))
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

    click_info('motorista_romaneio.png')
    pyautogui.press('F2')
    pyautogui.sleep(5)
    
    
    aviso_motorista('motorista_romaneio_nome.png')



    pyautogui.sleep(5)
    pyautogui.write(str(motorista))
    pyautogui.sleep(5)
    click_image('atualizar.png')
    pyautogui.sleep(5)
    click_image('selecionar.png')
    pyautogui.sleep(2)
    pyautogui.press('tab')
    pyautogui.sleep(4)
    print('antes de ir pra tela de pesquisa')
    click('ok_marcado.png')
    click('ok.png')
    click('ok2.png')
    click_info('motorista_romaneio.png')
    pyautogui.sleep(2)
    pyautogui.click(button='right')
    pyautogui.sleep(1)
    click_image('copy.png')
    pyautogui.sleep(0.5)
    try:
        text = pyperclip.paste()
        cod_motorista = int(text)
        print("Número do cod_motorista:", cod_motorista)
               
    except ValueError:
        print("O conteúdo copiado não é um número válido.")
        cod_motorista = 'Erro ao copiar cod_motorista'
    except Exception as e:
        print("Ocorreu um erro:", str(e))
    
    aviso = click_image_salvar('salvar.png','ok_marcado.png', 'campo_romaneio_automatico.png','aviso.png')

    if aviso:
        print(f"Aviso encontrado. Pulando para o próximo.")
        romaneio = 'AVISO DE MOTORISTA OU VEICULO'
        # Verificar se o arquivo já existe
        try:
            df = pd.read_excel(file_name)
        except FileNotFoundError:
            # Criar um DataFrame com as colunas ROMANEIO e PLACA, caso o arquivo não exista
            df = pd.DataFrame(columns=['DATA','ROMANEIO', 'PLACA', 'CIDADE'])

        # Adicionar o número do romaneio e a placa à planilha
        new_row = pd.DataFrame({'DATA': [agora_formatado],'ROMANEIO': [romaneio], 'PLACA': [placa], 'CIDADE': [cidade_mais_longe] })
        df = pd.concat([df, new_row], ignore_index=True)
        # Salvar a planilha
        df.to_excel(file_name, index=False)
        print(f"Número do romaneio {romaneio} e placa {placa} salvos na planilha {file_name}.")
        continue  # Pula para a próxima iteração do for

    #COPIAR CODIGO ROMANEIO E SALVAR NA PLANILHA
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
                pyautogui.sleep(1)
                a += 1

    pyautogui.sleep(3)
    click_image('guia_geral.png')
    pyautogui.sleep(5)
    click_image('efetuar.png')
    click_image('no.png')
    click_image('ok.png')
    print("")
    pyautogui.sleep(5)

print("Automatização concluída com sucesso!")