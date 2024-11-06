import pandas as pd
import pyautogui
from time import sleep, time
import pytesseract
import cv2 #opencv
import numpy as np
import os
import keyboard

cont_cenario = 397
cont_geral = 0
check_distri = 0
modifica_altura_verifica = True #ativar ou desativar modificação de altura
coeficiente_eficientiza = 0.6 #Valor de eficientização combinado
atender_eficientiza = False #Ligar ou desligar eficientização
qtde_bracos = 4 #quantos braços vao ser testados
altura_modificada = False


# Path to save the screenshot
screenshot_dir = r"C:/Users/AdminDell/Desktop/Pictures_FastPPP/images"
if not os.path.exists(screenshot_dir):
    os.makedirs(screenshot_dir)

caminho = r"C:\Program Files\Tesseract-OCR"
pytesseract.pytesseract.tesseract_cmd = caminho + r'\tesseract.exe'

# Carregar os dados da planilha
df = pd.read_excel('Cadastro_Piloto_SM-PONTOS.xlsx', sheet_name='Cadastro IPSM')
# Verificar as colunas para encontrar os nomes corretos
print(df.columns)

# Adicionar colunas 'luminaria_escolhida' e 'angulo_escolhido' se não existirem
if 'luminaria_escolhida' not in df.columns:
    df['luminaria_escolhida'] = ""
if 'angulo_escolhido' not in df.columns:
    df['angulo_escolhido'] = ""
if 'cenario' not in df.columns:
    df['cenario'] = ""
if 'nova_altura' not in df.columns:
    df['nova_altura'] = ""

# Garantir que a coluna 'luminaria_escolhida' é do tipo object
df['luminaria_escolhida'] = df['luminaria_escolhida'].astype(object)

# Extrair dados das colunas "larg_passeio_opost", "largura_via" e "larg_passeio_adj"
larg_passeio_opost = df['larg_passeio_opost'].tolist()
largura_via = df['largura_via'].tolist()
larg_passeio_adj = df['larg_passeio_adj'].tolist()
entre_postes = df['entre_postes'].tolist()
altura_lum = df['altura_lum'].tolist()
angulo = df['angulo'].tolist()
poste_pista = df['poste_pista'].tolist()
comprimento_braco = df['comprimento_braco'].tolist()
distribuicao = df['distribuicao'].str.lower().tolist()      
# Converter a coluna 'qtde_faixas' para inteiros
# Preencher valores ausentes com 0 e converter a coluna 'qtde_faixas' para inteiros
df['qtde_faixas'] = df['qtde_faixas'].fillna(0).astype(int)
qtde_faixas = df['qtde_faixas'].tolist()
larg_canteiro_central = df['larg_canteiro_central'].tolist()
pendor = df['pendor'].tolist()
classe_via = df['classe_via'].str.lower().tolist() 
classe_passeio = df['classe_passeio'].str.lower().tolist() 
luminaria_antiga = df['luminaria_antiga'].tolist()
ip = df['ip'].tolist()



# Iterar sobre os valores extraídos e digitar no campo correspondente
for idx, (larg_passeio_oposto, larg_via, larg_passeio_adjacente, entre_postes_x, altura_lum_x, angulo_x, poste_pista_x, comprimento_braco_x, qtde_faixas_x, larg_canteiro_central_x, pendor_x, classe_via_x, classe_passeio_x, luminaria_antiga, ip) in enumerate(zip(larg_passeio_opost, largura_via, larg_passeio_adj, entre_postes, altura_lum, angulo, poste_pista, comprimento_braco, qtde_faixas, larg_canteiro_central, pendor, classe_via, classe_passeio, luminaria_antiga, ip)):
    
    print(type(ip))
    if idx > 0 and idx < len(ip) - 1:  # Garantir que não está no primeiro ou último índice
        if ip[idx - 1] == ip[idx] == ip[idx + 1]:
            print("Pontos iguais: ")
            print(ip[idx - 1])
            print(ip[idx])
            print(ip[idx + 1])


