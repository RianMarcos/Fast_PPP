import pandas as pd
import pyautogui
from time import sleep, time
import pytesseract
import cv2 #opencv
import numpy as np
import os
import keyboard

cont_cenario = 134
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
df = pd.read_excel('Cadastro_Piloto_SM_V02.xlsx', sheet_name='Cadastro IPSM')
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


#------------ABRINDO CENARIO PADRAO ITAJAI-------------
# 1 - ABRINDO ARQUIVO
#pyautogui.doubleClick(147, 423, duration=0.5)
#sleep(30)  # TEMPO ATÉ ABRIR E CARREGAR O DIALUX

def verifica_resultados_eficientiza(novo_braco, luminaria_escolhida, comprimento_braco_x):
    print("Testando tamanho de braço para satisfazer eficientização")

    atende_braco = False
    guia_planejamento = pyautogui.locateCenterOnScreen('guia_planejamento.png', confidence=0.6)
    pyautogui.click(guia_planejamento.x, guia_planejamento.y)
    sleep(1.5)
    pyautogui.doubleClick(528,364)

    #click indice tipos para que os proximos tabs estejam corretos
    indice_tipos = pyautogui.locateCenterOnScreen('indice_tipos.png', confidence=0.9)
    pyautogui.click(indice_tipos.x, indice_tipos.y)
    sleep(0.3)

    if(valida_central == 0):
        print("Canteiro central não encontrado, tabs ajustados para tal")
        tab_interate(34)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        sleep(1)
        pyautogui.write(str(novo_braco))
        otimizar = pyautogui.locateCenterOnScreen('otimizar.png', confidence=0.8)
        pyautogui.click(otimizar.x, otimizar.y)
        sleep(6)
        global luminaria_escolhida_eficientiza
        global luminaria_escolhida_eficientiza_int
        luminaria_escolhida_eficientiza = verifica_atendimento()
        print("Nova luminária escolhida: "+ luminaria_escolhida)
        print("TESTE DA VARIAVEL EFICIENTIZA, VALOR GUARDADO NELA É DE: ")
        print(eficientiza)
        luminaria_escolhida_eficientiza_int = int(luminaria_escolhida_eficientiza)
        luminaria_escolhida_int = int(luminaria_escolhida)
        if(luminaria_escolhida != "NAO ATENDE" and luminaria_escolhida_eficientiza_int <= eficientiza ):
            global verifica_modificacaoBraco
            verifica_modificacaoBraco = True   
        else:
            verifica_modificacaoBraco = False
        print("Luminaria escolhida na eficientização e luminatia antiga: ")
        print(luminaria_escolhida_eficientiza_int)
        print(luminaria_escolhida_int)
        if(luminaria_escolhida_eficientiza_int >= luminaria_escolhida_int):
            print("Manter luminaria ja escolhida anteriormente, bem como o braço ")  
            luminaria_escolhida = luminaria_escolhida
            comprimento_braco_x = comprimento_braco_x
            #ajustar aqui para configurar o cenario para escrever o braco antigo no dialux
            luminaria_escolhida = luminaria_escolhida
            comprimento_braco_x = comprimento_braco_x
            print("Ajustando altura")
            guia_planejamento = pyautogui.locateCenterOnScreen('guia_planejamento.png', confidence=0.6)
            pyautogui.click(guia_planejamento.x, guia_planejamento.y)
            sleep(1.5)
            pyautogui.doubleClick(528,364)
            #click indice tipos para que os proximos tabs estejam corretos
            indice_tipos = pyautogui.locateCenterOnScreen('indice_tipos.png', confidence=0.9)
            pyautogui.click(indice_tipos.x, indice_tipos.y)
            sleep(0.3)

            tab_interate(34)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('delete')
            sleep(0.3)
            pyautogui.write(str(comprimento_braco_x))
            sleep(0.3)
        else:
            print("Nova luminária e novo braço escolhidos serão passados para frente")
            luminaria_escolhida = luminaria_escolhida_eficientiza
            comprimento_braco_x = novo_braco 
            global braco_modificado_check #verificador de alteração de braço
            braco_modificado_check = True
            return True

    else:
        print("Canteiro central encontrado, tabs ajustados para tal (fazer ainda logica)")
    return verifica_modificacaoBraco

def ajuste_braco(comprimento_braco_x, luminaria_escolhida):
    print("Função ajuste de braço")
    i =0
    #luminaria_escolhida_eficientiza = luminaria_escolhida #? teste
    while(i < qtde_bracos): #verificar braço a braço qual atende com passo de 0.5
        
        if(i == 0): #se for a primeira vez do laço pega o valor do comprimento do braco da planilha
            if(comprimento_braco_x <= 3.5):
                novo_braco = comprimento_braco_x + 0.5
            elif(comprimento_braco_x == 3.5):
                novo_braco = comprimento_braco_x - 0.5
        else: #se for a segunda ou mais vezes pega o valor do novo braço e vai somando 
            if(comprimento_braco_x <= 3.5):
                novo_braco = novo_braco + 0.5
            elif(comprimento_braco_x == 3.5):
                novo_braco = novo_braco - 0.5
        print("Vai chamar a função:")
        teste_passa_braco = verifica_resultados_eficientiza(novo_braco, luminaria_escolhida, comprimento_braco_x)

        if(teste_passa_braco == True and luminaria_escolhida_eficientiza <= eficientiza):
            print("O braço escolhido satisfaz a eficientização com sucesso")
            print("Braço de tamanho: ")
            print(novo_braco)
            print("Obteve um melhor desempenho que o antigo: ")
            print(comprimento_braco_x)
            print(" ")
            comprimento_braco_x = novo_braco

            break
        elif(teste_passa_braco == True and luminaria_escolhida_eficientiza > eficientiza):
            print("O braço escolhido satisfaz a condição para passar no cenário com um melor desemepnho que o antigo, mas a luminaria nao atingiu a eficientização")
            #em teoria aqui precisa voltar pro while e testar outro braço
            #ainda falta fazer uma condição pra caso nao atenda com nenhum braço
        elif(teste_passa_braco == False):
            print("será necessário reajustar para devido atendimento")
        '''
        if(teste_passa_braco == True): #esta funcao que vai escrever o tamannho do braço e verificar se atende
            print("Braço de tamanho: ")
            print(novo_braco)
            print("Obteve um melhor desempenho que o antigo: ")
            print(comprimento_braco_x)
            print(" ")
            comprimento_braco_x = novo_braco
            #arrumar aqui tabs pra chegar no braço
            break
        '''
        i +=1  
    return luminaria_escolhida_eficientiza, novo_braco



def porcentagem_eficientiza(luminaria_escolhida, verifica_modificacaoH):
    print(luminaria_escolhida)
    luminaria_escolhida_int = int(luminaria_escolhida)
    luminaria_antiga_float = float(luminaria_antiga)
    global eficientiza
    eficientiza = luminaria_antiga_float - (coeficiente_eficientiza * luminaria_antiga_float)
    print("Potência antiga: ")
    print(luminaria_antiga)
    print("Porcentagem de eficientização definida: ") 
    print(coeficiente_eficientiza)
    print("Eficientização a ser atingida: ")
    print(eficientiza)

    
    if(luminaria_escolhida_int > eficientiza):
        print("Luminaria não atende a porcentagem de eficientização, ir para ajuste de braço")
        global novo_braco_eficientiza
        luminaria_escolhida_eficientiza, novo_braco_eficientiza = ajuste_braco(comprimento_braco_x, luminaria_escolhida)

        guia_planejamento = pyautogui.locateCenterOnScreen('guia_planejamento.png', confidence=0.6)
        pyautogui.click(guia_planejamento.x, guia_planejamento.y)
        sleep(1.5)
        pyautogui.doubleClick(528,364)
        if(verifica_modificacaoH == False):
            print("Altura ainda nao modificada, ajustando altura para melhor desempenho da eficientização")
            modifica_altura()
            sleep(0.9)
            otimizar = pyautogui.locateCenterOnScreen('otimizar.png', confidence=0.8)
            pyautogui.click(otimizar.x, otimizar.y)
            sleep(6)

            luminaria_escolhida = verifica_atendimento()
            print("Nova luminária escolhida: "+ luminaria_escolhida)

        if(luminaria_escolhida == "NAO ATENDE"):
  
            altura_modificada = False
            print("Modificação de altura nao foi o bastante para a atender a este cenário, iniciar modifcação de braço")
        elif(luminaria_escolhida != "NAO ATENDE" and luminaria_escolhida_int <= eficientiza):
            verifica_modificacaoH = True
            
            altura_modificada = True
            altura_float = float(altura)
            print("Modificação de altura bem sucedida, nova altura de instalação: ")
            print(altura_float)
            print("Nova luminaria escolhida atendendo a eficientização " + luminaria_escolhida)
    else:
        print("Não será necessário fazer modificações, a luminaria escolhida ja atende a eficientização")
        luminaria_escolhida_eficientiza = luminaria_escolhida
    return luminaria_escolhida_eficientiza, altura_modificada

def modifica_altura():
    altura_maxima = 8.5
    altura_minima = 7.5
    passo = 0.5
    print("Modificando a altura de instalação")
    sleep(0.5)
    guia_planejamento = pyautogui.locateCenterOnScreen('guia_planejamento.png', confidence=0.6)
    pyautogui.click(guia_planejamento.x, guia_planejamento.y)
    sleep(1.5)

    #click indice tipos para que os proximos tabs estejam corretos
    indice_tipos = pyautogui.locateCenterOnScreen('indice_tipos.png', confidence=0.9)
    pyautogui.click(indice_tipos.x, indice_tipos.y)
    sleep(0.3)
    
    if(larg_canteiro_central_x > 0):
     tab_interate(20)
    else:
     tab_interate(19)
    pyautogui.press('space')
    sleep(0.8)

    tab_interate(2)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(str(altura_minima))
    sleep(0.3)

    pyautogui.press('tab')
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(str(altura_maxima))
    sleep(0.3)

    pyautogui.press('tab')
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(str(passo))
    sleep(0.3)
    refatorar_altura_inst = True
def refatora_altura():
    print("Arrumando cenário padrão: ")

    guia_planejamento = pyautogui.locateCenterOnScreen('guia_planejamento.png', confidence=0.8)
    pyautogui.click(guia_planejamento.x, guia_planejamento.y)
    sleep(1.5)

    ruas = pyautogui.locateCenterOnScreen('ruas.png', confidence=0.8)
    pyautogui.click(ruas.x, ruas.y)
    sleep(1)

    luminaria = pyautogui.locateCenterOnScreen('luminaria.png', confidence=0.8)
    pyautogui.click(luminaria.x, luminaria.y)
    sleep(1)

    if(larg_canteiro_central_x >0):
        tab_interate(23)
    else:
        tab_interate(22)
    pyautogui.press('space')
    print("Cenário padrão ajustado")


def choose_luminaria():
    ruas = pyautogui.locateCenterOnScreen('ruas.png', confidence=0.7) #ir para ruas e voltar para luminarias para resetar tabs
    pyautogui.click(ruas.x, ruas.y)
    sleep(0.4)
    luminaria = pyautogui.locateCenterOnScreen('luminaria.png', confidence=0.6)
    pyautogui.click(luminaria.x, luminaria.y)
    sleep(0.4)

    #click area preta para fazer a rolagem
    x_wall= 1410
    y_wall = 567
    pyautogui.click(x_wall, y_wall)

    pyautogui.scroll(+800)
    sleep(1.5)

    otimizar = pyautogui.locateCenterOnScreen('otimizar.png', confidence=0.8)
    pyautogui.click(otimizar.x, otimizar.y)
    sleep(3)

def verifica_atendimento():
    aba_resultado = pyautogui.locateCenterOnScreen('aba_resultado.png', confidence=0.8)
    pyautogui.click(aba_resultado.x, aba_resultado.y)
    sleep(2.5)

    #tirar print para verificar qual atende
    left = 634
    top = 205
    width = 314
    height = 26
    # Capturar tela da área de resultados
    screenshot = pyautogui.screenshot(region=(left, top, width, height)) #captura luminaria
    screenshot_path = os.path.join(screenshot_dir, f"results_{cont_geral}.png")
    screenshot.save(screenshot_path)

    luminaria = pytesseract.image_to_string(screenshot_path, lang='por', config='--psm 7').strip()
    print(luminaria)

    if "26W" in luminaria:
        nome_luminaria = "026"
    elif "30W" in luminaria:
        nome_luminaria = "030"
    elif "40W" in luminaria:
        nome_luminaria = "040"
    elif "50W" in luminaria:
        nome_luminaria = "050"
    elif "55W" in luminaria:
        nome_luminaria = "055"
    elif "60W" in luminaria:
        nome_luminaria = "060"
    elif "70W" in luminaria:
        nome_luminaria = "070"
    elif "80W" in luminaria:
        nome_luminaria = "080"
    elif "90W" in luminaria:
        nome_luminaria = "090"
    elif "100W" in luminaria:
        nome_luminaria = "100"
    elif "110W" in luminaria:
        nome_luminaria = "110"
    elif "120W" in luminaria:
        nome_luminaria = "120"
    elif "130W" in luminaria:
        nome_luminaria = "130"
    elif "150W" in luminaria:
        nome_luminaria = "150"
    elif "160W" in luminaria:
        nome_luminaria = "160"
    elif "170W" in luminaria:
        nome_luminaria = "170"
    elif "180W" in luminaria:
        nome_luminaria = "180"
    elif "200W" in luminaria:
        nome_luminaria = "200"
    elif "220W" in luminaria:
        nome_luminaria = "220"
    elif "240W" in luminaria:
        nome_luminaria = "240"
    else:
        nome_luminaria = "NAO ATENDE"
    

    #registrar altura de instalação
    leftalt = 1150
    topalt = 206
    widthalt = 51
    heightalt = 22
    # Capturar tela da área de resultados
    screenshot = pyautogui.screenshot(region=(leftalt, topalt, widthalt, heightalt)) #captura altura instalação
    screenshot_path_altura = os.path.join(screenshot_dir, f"altura_{cont_geral}.png")
    screenshot.save(screenshot_path_altura)

    global altura
    altura = pytesseract.image_to_string(screenshot_path_altura, lang='por', config='--psm 7').strip()
    print("A altura de instalação é : " + altura)

    return nome_luminaria

def save_pdf_report():
    click_image('guia_documentacao.png', 0.6)
    sleep(1)
    click_image('exibir_doc.png', 0.6, double_click=True)
    sleep(53)
    click_image('guardar_como.png', 0.7)
    sleep(0.5)
    click_image('pdf.png', 0.7)
    sleep(0.5)
    click_image('ok_pdf.png', 0.8)
    sleep(2)
    click_image('documentos_w11.png', 0.9)
    sleep(0.5)
    click_image('teste_pasta.png', 0.6, double_click=True)
    sleep(0.5)
    click_image('salvar_pasta.png', 0.7)
    sleep(5.5)

def click_image(image_path, confidence=0.7, double_click=False):
    location = pyautogui.locateCenterOnScreen(image_path, confidence=confidence)
    if location:
        if double_click:
            pyautogui.doubleClick(location.x, location.y)
        else:
            pyautogui.click(location.x, location.y)
        sleep(0.5)
    else:
        print(f"Imagem {image_path} não encontrada.")


def to_upper_safe(texto):
    if not isinstance(texto, str):
        raise TypeError("A variável fornecida não é uma string")
    
    # Remover espaços em branco e caracteres invisíveis
    texto = texto.strip()
    
    try:
        texto_maiusculo = texto.upper()
        return texto_maiusculo
    except Exception as e:
        print(f"Erro ao converter para maiúsculas: {e}")
        return None

def exclui_passeio(check_passeio_adjacente, check_passeio_oposto):
    sleep(1)
    #abrir guia das ruas
    ruas = pyautogui.locateCenterOnScreen('ruas.png', confidence=0.8)
    pyautogui.click(ruas)   
    sleep(1.5)

    print("Excluindo passeios necessários")
    if(check_passeio_oposto == 0): #será necessário excluir primeiro passeio
        try:
            sleep(0.5)
            passeio1 = pyautogui.locateCenterOnScreen('passeio1.png', confidence=0.8)
            check_passeio1 = 1 if passeio1 is not None else 0  
        except pyautogui.ImageNotFoundException:
            check_passeio1 = 0
        if check_passeio1 == 1: 
            print("Passeio1 Encontrado")
            sleep(0.5)
            passeio1 = pyautogui.locateCenterOnScreen('passeio1.png', confidence=0.8)
            pyautogui.click(passeio1)
            sleep(1.9)
            remover = pyautogui.locateCenterOnScreen('remover2.png', confidence=0.8)
            pyautogui.click(remover) 
            sleep(0.7)
        else:
            print("Passeio1 já foi excluido") 
            sleep(0.5)
    else:
        print("Manter primeiro passeio")
    sleep(1)
    if(check_passeio_adjacente == 0): #será necessário excluir segundo passeio
        try:
            sleep(0.5)
            passeio2 = pyautogui.locateCenterOnScreen('passeio2.png', confidence=0.8)
            check_passeio2 = 1 if passeio2 is not None else 0  
        except pyautogui.ImageNotFoundException:
            check_passeio2 = 0
        if check_passeio2 == 1: 
            print("Passeio2 Encontrado")
            sleep(0.5)
            passeio2 = pyautogui.locateCenterOnScreen('passeio2.png', confidence=0.8)
            pyautogui.click(passeio2)
            sleep(1.9)
            remover = pyautogui.locateCenterOnScreen('remover2.png', confidence=0.8)
            pyautogui.click(remover) 
            sleep(0.8)
        else:
            print("Passeio2 já foi excluido") 
            sleep(0.5)
    else:
        print("Manter segundo passeio")

def verifica_add_passeio():
    #entra todo começo de loop para adicionar passeio se ainda nao tem 
    #verifica se ja existe os dois passeios, se nao existir adiciona 
    try:
        passeio1 = pyautogui.locateCenterOnScreen('first_passeio.png', confidence=0.9)
        check_passeio1 = 1 if passeio1 is not None else 0   
    except pyautogui.ImageNotFoundException:
        check_passeio1 = 0
    if check_passeio1 == 1: 
        print("Passeio1 Encontrado")
    else:
        add_passeio = pyautogui.locateCenterOnScreen('add_passeio.png', confidence=0.8)
        pyautogui.click(add_passeio)
        print("Passeio1 adicionado") 
        sleep(1)
        pyautogui.moveTo(1136, 1136)
        sleep(0.3)
        #clicar no primeiro passeio
        first_passeio = pyautogui.locateCenterOnScreen('first_passeio.png', confidence=0.8)
        pyautogui.click(first_passeio)
        sleep(1)
        tab_interate(1)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        name_passeio_1 = "Passeio 1 (C3)"
        pyautogui.write(str(name_passeio_1))
        tab_interate(4)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(str(name_passeio_1))
        sleep(0.3)
        tab_interate(1)
        pyautogui.press('left', presses=9)
        sleep(8)



    try:
        passeio2 = pyautogui.locateCenterOnScreen('passeio2.png', confidence=0.9)
        check_passeio2 = 1 if passeio2 is not None else 0   # Verifica se a imagem 'central.png' foi encontrada
    except pyautogui.ImageNotFoundException:
        check_passeio2 = 0
    if check_passeio2 == 1: 
        print("Passeio2 Encontrado")
    else:
        add_passeio = pyautogui.locateCenterOnScreen('add_passeio.png', confidence=0.8)
        pyautogui.click(add_passeio)
        print("Passeio2 adicionado")  
        sleep(2) 
        pyautogui.moveTo(1136, 1136)
        sleep(0.3)
        #clicar no primeiro passeio
        first_passeio = pyautogui.locateCenterOnScreen('first_passeio.png', confidence=0.8)
        pyautogui.click(first_passeio)
        sleep(1)

        #moficar nome
        tab_interate(1)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        nome_passeio_2 = "Passeio 2"
        pyautogui.write(str(nome_passeio_2))
        tab_interate(4)
        sleep(0.5)
        name_passeio = "PASSEIO"
        pyautogui.write(str(name_passeio))
        tab_interate(1)
        pyautogui.press('left', presses=9)
        sleep(8)

        #descer para último
        seta_baixo = pyautogui.locateCenterOnScreen('seta_baixo.png', confidence=0.9)
        pyautogui.click(seta_baixo.x, seta_baixo.y)
        pyautogui.click(seta_baixo.x, seta_baixo.y)
        pyautogui.click(seta_baixo.x, seta_baixo.y)
        pyautogui.click(seta_baixo.x, seta_baixo.y)
        pyautogui.click(seta_baixo.x, seta_baixo.y)
        pyautogui.click(seta_baixo.x, seta_baixo.y)
        sleep(5)


def classifica_vias_passeios():
    #click area preta para fazer a rolagem
    x_wall= 1410
    y_wall = 567
    pyautogui.click(x_wall, y_wall)

    pyautogui.scroll(-1000) 
    sleep(2)

    seta_passeio1 = pyautogui.locateCenterOnScreen('seta_passeio1.png', confidence=0.8)
    pyautogui.click(seta_passeio1)
    sleep(1)

    if(valida_central == 1  or valida_central_sem_distri== 1):
        seta_pista_rodagem2 = pyautogui.locateCenterOnScreen('seta_pista_rodagem2.png', confidence=0.9)
        pyautogui.click(seta_pista_rodagem2.x, seta_pista_rodagem2.y)

    sleep(1)
    seta_pista_rodagem1 = pyautogui.locateCenterOnScreen('seta_pista_rodagem1.png', confidence=0.9)
    pyautogui.click(seta_pista_rodagem1.x, seta_pista_rodagem1.y)
    sleep(1)
    seta_passeio2 = pyautogui.locateCenterOnScreen('seta_passeio2.png', confidence=0.9)
    pyautogui.click(seta_passeio2.x, seta_passeio2.y)

    #abrir a janela necessária
    seta_passeio1_closed = pyautogui.locateCenterOnScreen('seta_passeio1_closed.png', confidence=0.8)
    pyautogui.click(seta_passeio1_closed)
    sleep(1)

    #modificar em 
    em_parametro = pyautogui.locateCenterOnScreen('em_parametro.png', confidence=0.8)
    pyautogui.click(em_parametro)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(str(classe_passeio_em))

    #modificar uo
    uo_parametro = pyautogui.locateCenterOnScreen('uo_parametro.png', confidence=0.6)
    pyautogui.click(uo_parametro)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(str(classe_passeio_uo))

    #fechar janela modificada e passar para próxima 
    seta_passeio1 = pyautogui.locateCenterOnScreen('seta_passeio1.png', confidence=0.8)
    pyautogui.click(seta_passeio1)
    sleep(1)
    pyautogui.scroll(-300)
    sleep(1)
    #passar para proxima
    if(valida_central == 1 or valida_central_sem_distri== 1):
        #preencher valroes para canteiro central

        #abrir a janela necessária
        seta_pista2_closed = pyautogui.locateCenterOnScreen('seta_pista2_closed.png', confidence=0.9)
        pyautogui.click(seta_pista2_closed)
        sleep(1)

        #modificar em 
        em_parametro = pyautogui.locateCenterOnScreen('em_parametro.png', confidence=0.8)
        pyautogui.click(em_parametro)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(str(classe_via_em))

        #modificar uo
        uo_parametro = pyautogui.locateCenterOnScreen('uo_parametro.png', confidence=0.6)
        pyautogui.click(uo_parametro)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(str(classe_via_uo))
        pyautogui.scroll(-300)
        sleep(1)

        #fechar janela modificada e passar para próxima 
        seta_pista_rodagem2 = pyautogui.locateCenterOnScreen('seta_pista_rodagem2.png', confidence=0.9)
        pyautogui.click(seta_pista_rodagem2)
        sleep(1)
        pyautogui.scroll(-300)
        sleep(1)
    
    #abrir a janela necessária
    seta_pista1_closed = pyautogui.locateCenterOnScreen('seta_pista1_closed.png', confidence=0.9)
    pyautogui.click(seta_pista1_closed)
    sleep(1)

    #modificar em 
    em_parametro = pyautogui.locateCenterOnScreen('em_parametro.png', confidence=0.8)
    pyautogui.click(em_parametro)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(str(classe_via_em))
    sleep(0.2)
    pyautogui.scroll(-300)
    sleep(0.8)

    #modificar uo
    uo_parametro = pyautogui.locateCenterOnScreen('uo_parametro.png', confidence=0.6)
    pyautogui.click(uo_parametro)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(str(classe_via_uo))
    sleep(0.2)
    pyautogui.scroll(-300)
    sleep(1)

    #fechar janela modificada e passar para próxima 
    seta_pista_rodagem1 = pyautogui.locateCenterOnScreen('seta_pista_rodagem1.png', confidence=0.9)
    pyautogui.click(seta_pista_rodagem1)
    sleep(1)
    pyautogui.scroll(-600)
    sleep(1)
    
    #abrir a janela necessária
    seta_passeio2_closed = pyautogui.locateCenterOnScreen('seta_passeio2_closed.png', confidence=0.9)
    pyautogui.click(seta_passeio2_closed)
    sleep(1)
    pyautogui.scroll(-600)
    sleep(1)
    #modificar em 
    em_parametro = pyautogui.locateCenterOnScreen('em_parametro.png', confidence=0.8)
    pyautogui.click(em_parametro)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(str(classe_passeio_em))

    #modificar uo
    uo_parametro = pyautogui.locateCenterOnScreen('uo_parametro.png', confidence=0.6)
    pyautogui.click(uo_parametro)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(str(classe_passeio_uo))

    #-----abrindo todas as guias-----
    seta_pista1_closed = pyautogui.locateCenterOnScreen('seta_pista1_closed.png', confidence=0.9)
    pyautogui.click(seta_pista1_closed)
    sleep(1)
    #abrir a janela necessária
    seta_passeio1_closed = pyautogui.locateCenterOnScreen('seta_passeio1_closed.png', confidence=0.8)
    pyautogui.click(seta_passeio1_closed)
    sleep(1)

    if(valida_central == 1 or valida_central_sem_distri== 1):
        #abrir a janela necessária
        seta_pista2_closed = pyautogui.locateCenterOnScreen('seta_pista2_closed.png', confidence=0.9)
        pyautogui.click(seta_pista2_closed)
        sleep(1)

    pyautogui.scroll(-1000)
    #abrir todas as janelas novamente para verificar os checks (manter aberta)
    #lembrar de usar rolagem scroll

#função para verificar se possui canteiro central e fazer devido deslocamento de posição via 'tab'
def teste_central(x_img, y_img, tabs):
    try:
        img_central = pyautogui.locateCenterOnScreen('central.png', confidence=0.8)
        auxiliar = 1 if img_central is not None else 0   # Verifica se a imagem 'central.png' foi encontrada
    except pyautogui.ImageNotFoundException:
        auxiliar = 0
    if auxiliar == 1: 
        pyautogui.click(x_img, y_img)
        tab_interate(tabs)
        print("Canteiro central encontrado")
    else:
        pyautogui.click(x_img, y_img)
        tab_interate(tabs - 1)
        print("Canteiro central não encontrado")  

def check_all(screenshot_path, validacao_central):
    if not os.path.exists(screenshot_path):
        print(f"File not found: {screenshot_path}")
        return False

    image = cv2.imread(screenshot_path)
    if image is None:
        print(f"Failed to load image: {screenshot_path}")
        return False

    # Convert image to HSV color space
    hsv_image = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)

    # Define range for green color in HSV
    lower_green = np.array([35, 100, 100])
    upper_green = np.array([85, 255, 255])

    # Create a mask for green color
    mask = cv2.inRange(hsv_image, lower_green, upper_green)

    # Debugging: save the mask to visualize
    mask_path = screenshot_path.replace('.png', '_mask.png')
    cv2.imwrite(mask_path, mask)
    print(f"Mask saved to {mask_path}")

    # Find contours in the mask
    contours, _ = cv2.findContours(mask, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

    # Count the number of contours
    num_checks = len(contours)
    print(f"Number of checks found: {num_checks}")

    # Check if the number of contours (checks)
    if(check_passeio_oposto == 1 and check_passeio_adjacente ==1):
        if(validacao_central == 1):
            if num_checks >= 8:
                return True
            return False
        else:
            if num_checks >= 6:
                return True
            return False
    elif(check_passeio_oposto == 0 and check_passeio_adjacente == 1):   
        if(validacao_central == 1):
            if num_checks >= 6:
                return True
            return False
        else:
            if num_checks >= 4:
                return True
            return False
        
    elif(check_passeio_oposto == 1 and check_passeio_adjacente == 0):   
        if(validacao_central == 1):
            if num_checks >= 6:
                return True
            return False
        else:
            if num_checks >= 4:
                return True
            return False
        
    elif(check_passeio_oposto == 0 and check_passeio_adjacente == 0):   
        if(validacao_central == 1):
            if num_checks >= 4:
                return True
            return False
        else:
            if num_checks >= 2:
                return True
            return False  


def tab_interate(cont):
    i = 0
    while i < cont:
        pyautogui.press('tab')
        i += 1
    i = 0


# Iterar sobre os valores extraídos e digitar no campo correspondente
for idx, (larg_passeio_oposto, larg_via, larg_passeio_adjacente, entre_postes_x, altura_lum_x, angulo_x, poste_pista_x, comprimento_braco_x, qtde_faixas_x, larg_canteiro_central_x, pendor_x, classe_via_x, classe_passeio_x, luminaria_antiga, ip) in enumerate(zip(larg_passeio_opost, largura_via, larg_passeio_adj, entre_postes, altura_lum, angulo, poste_pista, comprimento_braco, qtde_faixas, larg_canteiro_central, pendor, classe_via, classe_passeio, luminaria_antiga, ip)):
    sleep(1.5)
    braco_modificado_check = False
    refatorar_altura_inst = False
    #braco_modificado_check = False #veriricador de alteração de braço
    verifica_modificacaoH = False #sempre que entrar no loop precisa estar em false pra conseguir entrar na modificação de altura da eficientização
    altura_modificada = False
    # Verifica se a tecla Shift está pressionada
    if keyboard.is_pressed('shift'):
        print('A tecla Shift está pressionada.')
    else:
        print('A tecla Shift não está pressionada.')
        
    if(classe_via_x == "v1" or classe_via_x == "V1"):
        classe_via_em = 30
        classe_via_uo = 0.4
        print(classe_via_em)
        print(classe_via_uo)
    elif(classe_via_x == "v2"):
        classe_via_em = 20
        classe_via_uo = 0.3
    elif(classe_via_x == "v3"):
        classe_via_em = 15
        classe_via_uo = 0.2
    elif(classe_via_x == "v4"):
        classe_via_em = 10
        classe_via_uo = 0.2
    elif(classe_via_x == "v5"):
        classe_via_em = 5
        classe_via_uo = 0.2

    if(classe_passeio_x == "p1"):
        classe_passeio_em = 20
        classe_passeio_uo = 0.3
    elif(classe_passeio_x == "p2"):
        classe_passeio_em = 10
        classe_passeio_uo = 0.25
    elif(classe_passeio_x == "p3"):
        classe_passeio_em = 5
        classe_passeio_uo = 0.2
    elif(classe_passeio_x == "p4"):
        classe_passeio_em = 3
        classe_passeio_uo = 0.2

    #Verificar se será necessário excluir ou adicionar um passeio
    if(larg_passeio_oposto == 0):
        check_passeio_oposto = 0
    else: 
        check_passeio_oposto = 1

    if(larg_passeio_adjacente == 0):
        check_passeio_adjacente = 0
    else: 
        check_passeio_adjacente = 1

    print("Distribuição: "+ distribuicao[idx])
    cont_geral += 1  # var para fazer a contagem de cenários 
    cont__str = str(cont_geral)  # var para fazer conversão de int para string e passar como parametro no nome do cenário

    cont_cenario += 1
    cont_cenario_str = str(cont_cenario)
    
    # Abrindo guia planejamento
    pyautogui.click(399, 82, duration=0.5)
    sleep(1)
    ruas = pyautogui.locateCenterOnScreen('ruas.png', confidence=0.6)
    pyautogui.click(ruas.x, ruas.y)
    sleep(1)

    verifica_add_passeio()

    auxiliar_1 = 0
    if distribuicao[cont_geral-1] == 'central' or distribuicao[cont_geral-1]== 'canteiro central' or distribuicao[cont_geral-1]==  'canteiro_central' or distribuicao[cont_geral-1]== 'central' or larg_canteiro_central_x != 0:
        #validar se ja existe canteiro central, se nao existir adicionar 
        try:
            faixa_central_1 = pyautogui.locateCenterOnScreen('faixa_central_1.png', confidence=0.8)
            auxiliar_1 = 1 if faixa_central_1 is not None else 0  
        except pyautogui.ImageNotFoundException:
            auxiliar_1 = 0
            print('Imagem da faixa central não encontrada')
        if auxiliar_1 == 1: 
            print("Faixa ja adicionada")
        else:
            adicionar_faixa_central = pyautogui.locateCenterOnScreen('adicionar_faixa_central.png', confidence=0.8)
            pyautogui.click(adicionar_faixa_central.x, adicionar_faixa_central.y)
            sleep(1)
            faixa_central_1 = pyautogui.locateCenterOnScreen('faixa_central_1.png', confidence=0.8)
            pyautogui.click(faixa_central_1.x, faixa_central_1.y)
            sleep(2)
            seta_baixo = pyautogui.locateCenterOnScreen('seta_baixo.png', confidence=0.9)
            pyautogui.click(seta_baixo.x, seta_baixo.y)
            sleep(1)
            adicionar_pista = pyautogui.locateCenterOnScreen('adicionar_pista.png', confidence=0.9)
            pyautogui.click(adicionar_pista.x, adicionar_pista.y)
            sleep(2)
            pista_de_rodagem2 = pyautogui.locateCenterOnScreen('pista_de_rodagem2.png', confidence=0.9)
            pyautogui.click(pista_de_rodagem2.x, pista_de_rodagem2.y)
            sleep(1.5)
            tab_interate(10)
            pyautogui.press('left')
            pyautogui.press('left')
            pyautogui.press('left')
            pyautogui.press('left')
            pyautogui.press('left')
            pyautogui.press('left')
            pyautogui.press('left')
            pyautogui.press('left')
            pyautogui.press('left')
            sleep(8)
            seta_baixo = pyautogui.locateCenterOnScreen('seta_baixo.png', confidence=0.9)
            pyautogui.click(seta_baixo.x, seta_baixo.y)
            sleep(1)

        #restante dos passos para inserir valores nos campos correspondentes 
        #PASSEIO1
        sleep(1.5)
        passeio1 = pyautogui.locateCenterOnScreen('passeio1.png', confidence=0.7)
        pyautogui.doubleClick(passeio1.x, passeio1.y)
        sleep(1.5)
        tab_interate(3)
        # Selecionar todo o texto existente e apagar
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        # Digitando o novo valor para larg_passeio_opost
        pyautogui.write(str(larg_passeio_oposto))
        sleep(1.5)

        #PISTA DE RODAGEM2
        tab_interate(11)
        sleep(1)
        pyautogui.press('Down')
        # Clicando no campo largura via (ajustar coordenadas conforme necessário)
        # pista1 = pyautogui.locateCenterOnScreen('pista1.png', confidence=0.7)
        # pyautogui.doubleClick(pista1.x, pista1.y)
        # sleep(1)
        tab_interate(1)
        nome_pista_2 = "Pista de rodagem 2"
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(str(nome_pista_2))
        sleep(0.5)
        tab_interate(5)
        # Selecionar todo o texto existente e apagar
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        # Digitando o novo valor para largura_via
        pyautogui.write(str(larg_via))
        sleep(1.2)

        tab_interate(1)
        sleep(0.7)
        print("A quantidade de faixas é: "+ str(qtde_faixas_x))
        sleep(0.5)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        sleep(0.7)
        # Digitando o novo valor para
        # Digitando qtde de faixas
        pyautogui.write(str(qtde_faixas_x))
        sleep(1.5)
   
        #CANTEIRO CENTRAL
        tab_interate(12) #validar se esta certo
        pyautogui.press('Down')
        tab_interate(3)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(str(larg_canteiro_central_x))

        #PISTA DE RODAGEM1
        pyautogui.hotkey('shift', 'tab')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.press('tab')
        pyautogui.press('Down')
        sleep(0.5)
        # Clicando no campo largura via (ajustar coordenadas conforme necessário)
        # pista1 = pyautogui.locateCenterOnScreen('pista1.png', confidence=0.7)
        # pyautogui.doubleClick(pista1.x, pista1.y)
        # sleep(1)
        tab_interate(6)
        # Selecionar todo o texto existente e apagar
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        # Digitando o novo valor para largura_via
        pyautogui.write(str(larg_via))
        sleep(1)
        tab_interate(1)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        sleep(0.5)
        print("A quantidade de faixas é: "+ str(qtde_faixas_x))
        sleep(0.5)
        # Digitando qtde de faixas
        pyautogui.write(str(qtde_faixas_x))
        sleep(1)
        
        #PASSEIO2
        pyautogui.hotkey('shift', 'tab')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.press('tab')
        pyautogui.press('Down')
        passeio2 = pyautogui.locateCenterOnScreen('passeio2.png', confidence=0.7)
        pyautogui.doubleClick(passeio2.x, passeio2.y)
        sleep(1)
        tab_interate(3)
        # Selecionar todo o texto existente e apagar
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        # Digitando o novo valor para larg_passeio_opost
        pyautogui.write(str(larg_passeio_adjacente))
        sleep(1.5)      

    else:
        try: 
            #antes de remover o passeio será necessário mover a distribuicao do canteiro central p/ qqlr outra
            luminaria = pyautogui.locateCenterOnScreen('luminaria.png', confidence=0.7)
            pyautogui.doubleClick(luminaria.x, luminaria.y)
            sleep(0.5)

            tab_interate(16)
            bilateral = pyautogui.locateCenterOnScreen('bilateral.png', confidence=0.8)
            pyautogui.doubleClick(bilateral.x, bilateral.y)
            sleep(0.5)  
            
            ruas = pyautogui.locateCenterOnScreen('ruas.png', confidence=0.8)
            pyautogui.doubleClick(ruas.x, ruas.y)
            sleep(0.5)  

            #removendo canteiro central e segunda via adicionada
            faixa_central_1 = pyautogui.locateCenterOnScreen('faixa_central_1.png', confidence=0.8)
            pyautogui.click(faixa_central_1.x, faixa_central_1.y)
            sleep(2)
            remover_central = pyautogui.locateCenterOnScreen('remover_central.png', confidence=0.9)
            pyautogui.click(remover_central.x, remover_central.y)
            sleep(2)

            pista_de_rodagem2 = pyautogui.locateCenterOnScreen('pista_de_rodagem2.png', confidence=0.9)
            pyautogui.click(pista_de_rodagem2.x, pista_de_rodagem2.y)
            sleep(0.8)
            remover = pyautogui.locateCenterOnScreen('remover.png', confidence=0.9)
            pyautogui.click(remover.x, remover.y)
            sleep(2)
        except pyautogui.ImageNotFoundException:
            print("Imagem do canteiro central nao encontrada")
    
        # Selecionando o passeio1
        #tab_interate(8)
        sleep(1.5)
        passeio1 = pyautogui.locateCenterOnScreen('passeio1.png', confidence=0.7)
        pyautogui.doubleClick(passeio1.x, passeio1.y)
        sleep(1)
        tab_interate(3)
        # Selecionar todo o texto existente e apagar
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        # Digitando o novo valor para larg_passeio_opost
        pyautogui.write(str(larg_passeio_oposto))
        sleep(1.5)

        ##--------------------- PARAMETROS RUA ---------------------
        tab_interate(11)
        sleep(1)
        pyautogui.press('Down')
        # Clicando no campo largura via (ajustar coordenadas conforme necessário)
    # pista1 = pyautogui.locateCenterOnScreen('pista1.png', confidence=0.7)
    # pyautogui.doubleClick(pista1.x, pista1.y)
    # sleep(1)
        tab_interate(6)
        # Selecionar todo o texto existente e apagar
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        # Digitando o novo valor para largura_via
        pyautogui.write(str(larg_via))
        sleep(1)

        tab_interate(1)
        sleep(0.7)
        print("A quantidade de faixas é: "+ str(qtde_faixas_x))
        sleep(0.5)
        # Digitando qtde de faixas
        pyautogui.write(str(qtde_faixas_x))
        sleep(1.5)
    
        ##--------------------- PARAMETROS PASSEIO ADJACENTE ---------------------
        tab_interate(12)
        sleep(1)
        pyautogui.press('Down')
        # Selecionando o passeio2
    # passeio2 = pyautogui.locateCenterOnScreen('passeio2.png', confidence=0.7)
        #pyautogui.doubleClick(passeio2.x, passeio2.y)
        #sleep(1)
        tab_interate(3)
        # Selecionar todo o texto existente e apagar
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        # Digitando o novo valor para larg_passeio_adjacente
        pyautogui.write(str(larg_passeio_adjacente))
        sleep(0.8)

    ##--------------------- PARAMETROS LUMINÁRIA ---------------------
    img = pyautogui.locateCenterOnScreen('luminaria.png', confidence=0.7)
    pyautogui.click(img.x, img.y)
    sleep(0.8)
    tab_interate(16)

    #clicar no bilateral para atualizar distruibuições e liberar canteiro central
    img_bilateral = pyautogui.locateCenterOnScreen('bilateral.png', confidence =0.7)
    pyautogui.click(img_bilateral.x, img_bilateral.y)
    sleep(2.5)
    img_uni = pyautogui.locateCenterOnScreen('unilateral_inferior.png', confidence =0.7)
    pyautogui.click(img_uni.x, img_uni.y)
    sleep(2.5)

    # Posicione o mouse sobre a scrollbar 
    #pyautogui.moveTo(492, 512)  # Ajuste as coordenadas conforme necessário
    #target = 917
    #scroll_to_position(target, 300)
    #sleep(1.5)

    #Distância entre postes entre_postes
    #postes = pyautogui.locateCenterOnScreen('entre_postes.png', confidence=0.6)
    #pyautogui.click(postes.x, postes.y)
    #Selecionando tipo de distruibuicao dos postes

    valida_central = 0
    valida_central_sem_distri = 0
    if(larg_canteiro_central_x != 0): #receber um aqui para que na classificação dos passeios ele entenda que possui um canteiro central
        valida_central_sem_distri = 1

    if distribuicao[cont_geral-1] == 'unilateral' or distribuicao[cont_geral-1] == 'unilateral inferior' or distribuicao[cont_geral-1] == 'unilateral_inferior' :
        img_uni = pyautogui.locateCenterOnScreen('unilateral_inferior.png', confidence =0.7)
        pyautogui.click(img_uni.x, img_uni.y)
        print("ENTROU NO UNILATERAL")
        sleep(1)
        x_img = img_uni.x #posicao da distruibuicao
        y_img = img_uni.y
        tabs = 6 #qtde de tabs para passar para o proximo
        teste_central(x_img, y_img, tabs)
        sleep(0.5)
  
    elif distribuicao[cont_geral-1] == 'bilateral' or distribuicao[cont_geral-1] == 'bilateral frontal':
        img_bilateral = pyautogui.locateCenterOnScreen('bilateral.png', confidence =0.7)
        pyautogui.click(img_bilateral.x, img_bilateral.y)
        print("ENTROU NO bilateral")
        sleep(1)
        x_img = img_bilateral.x #posicao da distruibuicao
        y_img = img_bilateral.y
        tabs = 4 #qtde de tabs para passar para o proximo
        teste_central(x_img, y_img, tabs)
        sleep(0.5)

    elif distribuicao[cont_geral-1]== 'bilateral_alternada' or distribuicao[cont_geral-1] == 'bilateral alternada' or distribuicao[cont_geral-1] == 'Bilateral Alternado' or distribuicao[cont_geral-1] == 'bilateral alternado':
        print("entrou no bilat alternada")
        img_bilateral_alternada = pyautogui.locateCenterOnScreen('bilateral_alternada.png', confidence=0.7)
        pyautogui.click(img_bilateral_alternada.x, img_bilateral_alternada.y)
        sleep(1)
        x_img = img_bilateral_alternada.x #posicao da distruibuicao
        y_img = img_bilateral_alternada.y
        tabs = 3 #qtde de tabs para passar para o proximo
        teste_central(x_img, y_img, tabs)
        sleep(0.5)

    elif distribuicao[cont_geral-1] == 'central' or distribuicao[cont_geral-1]== 'canteiro central' or distribuicao[cont_geral-1]==  'canteiro_central' or distribuicao[cont_geral-1]== 'central':
        valida_central = 1
        img_central = pyautogui.locateCenterOnScreen('central.png', confidence =0.8)
        pyautogui.click(img_central.x, img_central.y)
        print("Entrou na distri canteiro central")
        sleep(1)     
        tab_interate(2)
    

    sleep(0.5)
    #Distancia entre postes
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    # Digitando o novo valor para larg_passeio_adjacente
    pyautogui.write(str(entre_postes_x))
    sleep(3.5)
    #Altura do ponto de luz
    tab_interate(3)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(str(altura_lum_x))

    #se for canteiro central inserir duas luminarias por poste
    if(valida_central == 1):
        tab_interate(3)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(str(2))
    else:
        tab_interate(3)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(str(1))
    sleep(0.6)
    #Angulo
    tab_interate(2)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(str(angulo_x))
    
    if(valida_central == 0): #trava o pendor se nao for distri_central e insere valores para poste_pista e braco
        print("Entrou na validação")
        tab_interate(2)
        pyautogui.press('space')
        sleep(0.5)
        #distância poste-pista
        tab_interate(3)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(str(poste_pista_x))
        #comprimento do braço
        tab_interate(2)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(str(comprimento_braco_x))
    else: #se a distribuicao estiver no canteiro central, será necessário informar valor do pendor e até mesmo do deslocamento longitudinal
        tab_interate(3)     
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(str(pendor_x))
        #ADICIONAR AQUI DESLOCAMENTO LONGITUDINAL

    classifica_vias_passeios()
                                              
    if(check_passeio_adjacente == 0 or check_passeio_oposto == 0):
        print("entrou no IF que chama a funcao de exclusao")
        exclui_passeio(check_passeio_adjacente, check_passeio_oposto)

    #------------------------------------------#CHOOSE LUM-------------------------------------------#
    choose_luminaria()
    #--------------------------------------------------------------------------------------------------#
    #verifica_atendimento()

    luminaria_escolhida = verifica_atendimento()
    
    if(luminaria_escolhida == "NAO ATENDE" and modifica_altura_verifica == True):
        modifica_altura()
        verifica_modificacaoH = True
        sleep(0.9)
        otimizar = pyautogui.locateCenterOnScreen('otimizar.png', confidence=0.8)
        pyautogui.click(otimizar.x, otimizar.y)
        sleep(6)

        luminaria_escolhida = verifica_atendimento()
        print("Nova luminária escolhida: "+ luminaria_escolhida)

        if(luminaria_escolhida == "NAO ATENDE"):
            altura_modificada = False
            print("Modificação de altura nao foi o bastante para a atender a este cenário, iniciar modifcação de braço")
            refatorar_altura_inst = True
        else:
            altura_modificada = True
            altura_float = float(altura)
            print("Modificação de altura bem sucedida, nova altura de instalação: ")
            print(altura_float)
            
   

    if(atender_eficientiza == True):
        luminaria_escolhida, altura_modificada = porcentagem_eficientiza(luminaria_escolhida, verifica_modificacaoH)

        #arrumar parametros do cenário padrão caso a altura de instalção foi modificada
    if(altura_modificada == True or refatorar_altura_inst == True):
        refatora_altura()


    #---------------------------modificando nome do projeto---------------------------
    sleep(1.5)
    nome_projet = pyautogui.locateCenterOnScreen('nome_projeto.png', confidence=0.6)
    pyautogui.click(nome_projet.x, nome_projet.y)
    sleep(1)
    digitar_nome = pyautogui.locateCenterOnScreen('descri.png', confidence=0.6)
    pyautogui.click(digitar_nome.x, digitar_nome.y)
    sleep(0.4)
    pyautogui.hotkey('shift', 'tab')

    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    if(luminaria_escolhida != "NAO ATENDE" and altura_modificada == True and braco_modificado_check == True):
        print("Entrou no altura modificada e braco modificado")
        modify_name = "Santa Maria " + cont_cenario_str + " - " + ip + " - AGN7" + luminaria_escolhida + "D4 " + " - H" + str(altura_float) + " - BR" + str(novo_braco_eficientiza) #entra aqui se modificou a altura de instalção e braço
    elif(luminaria_escolhida != "NAO ATENDE" and altura_modificada == True):
        print("Entrou no altura modificada")
        modify_name = "Santa Maria " + cont_cenario_str + " - " + ip + " - AGN7" + luminaria_escolhida + "D4 " + " - H" + str(altura_float) #entra aqui se modificou a altura de instalção
    elif(luminaria_escolhida != "NAO ATENDE" and braco_modificado_check == True):
        print("braco modificado")
        modify_name = "Santa Maria " + cont_cenario_str + " - " + ip +" - AGN7" + luminaria_escolhida + "D4 " + " - BR" + str(comprimento_braco_x) 
    elif(luminaria_escolhida != "NAO ATENDE"):
        print("Entrou no altura nao modificada")
        modify_name = "Santa Maria " + cont_cenario_str + " - " + ip + " - AGN7" + luminaria_escolhida + "D4"
    else:
        modify_name = "Santa Maria " + cont_cenario_str + " - " + ip + " - " + luminaria_escolhida #entra aqui se nao atender
    #modify_name.upper() 
     
    to_upper_safe(modify_name)
    pyautogui.write(modify_name.upper()) 

    #pyautogui.write(str("Itajai " + cont__str + " " + luminaria_escolhida))
    sleep(2.3)

    #salvando pdf relatório
    save_pdf_report()

    #salvando arquivo editável
    ficheiro = pyautogui.locateCenterOnScreen('ficheiro.png', confidence=0.7)
    pyautogui.click(ficheiro.x, ficheiro.y)
    sleep(0.5)
    guardar_project = pyautogui.locateCenterOnScreen('guardar_project.png', confidence=0.8)
    pyautogui.click(guardar_project.x, guardar_project.y)
    sleep(1.5)
    documentos = pyautogui.locateCenterOnScreen('documentos_w11.png', confidence=0.7)
    pyautogui.click(documentos.x, documentos.y)
    sleep(0.3)
    teste = pyautogui.locateCenterOnScreen('teste_pasta.png', confidence=0.6)
    pyautogui.doubleClick(teste.x, teste.y)
    sleep(0.5)
    nome_editavel = pyautogui.locateCenterOnScreen('nome_editavel.png', confidence=0.7)
    pyautogui.doubleClick(nome_editavel.x, nome_editavel.y)
    sleep(1)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    if(luminaria_escolhida != "NAO ATENDE" and altura_modificada == True and braco_modificado_check == True):
        print("Entrou no altura modificada e braco modificado")
        project_name = "Santa Maria " + cont_cenario_str + " - "  + ip + " - AGN7" + luminaria_escolhida + "D4 " + " - H" + str(altura_float) + " - BR" + str(novo_braco_eficientiza) #entra aqui se modificou a altura de instalção e braço
    elif(luminaria_escolhida != "NAO ATENDE" and altura_modificada == True):
        print("Entrou no altura modificada")
        project_name = "Santa Maria " + cont_cenario_str + " - " + ip + " - AGN7" + luminaria_escolhida + "D4 " + " - H" + str(altura_float) #entra aqui se modificou a altura de instalção
    elif(luminaria_escolhida != "NAO ATENDE" and braco_modificado_check == True):
        print("braco modificado")
        project_name = "Santa Maria " + cont_cenario_str + " - " +  ip +" - AGN7" + luminaria_escolhida + "D4 " + " - BR" + str(comprimento_braco_x) 
    elif(luminaria_escolhida != "NAO ATENDE"):
        print("Entrou no altura nao modificada")
        project_name = "Santa Maria " + cont_cenario_str + " - " + ip + " - AGN7" + luminaria_escolhida + "D4"
    else:
        project_name = "Santa Maria " + cont_cenario_str + " - " + ip + " - " +  luminaria_escolhida #entra aqui se nao atender
    pyautogui.write(project_name.upper() + ".evo")
    
    pyautogui.moveTo(1136, 1136)
    sleep(0.8)
   # pyautogui.write(str("Itajai " + cont__str + " " + luminaria_escolhida + ".evo")) #nome arquivo
    salvar_pasta = pyautogui.locateCenterOnScreen('salvar_pasta.png', confidence=0.6)
    pyautogui.click(salvar_pasta.x, salvar_pasta.y)
    sleep(4)
    '''
    #arrumar parametros do cenário padrão caso a altura de instalção foi modificada
    if(verifica_modificacaoH  == True):
        refatora_altura()
    '''
    if(luminaria_escolhida == "NAO ATENDE"):
        df.at[idx, 'luminaria_escolhida'] =  luminaria_escolhida 
    else:    
        # Atualizar a planilha com a luminária escolhida e o ângulo
        df.at[idx, 'luminaria_escolhida'] = "AGN7" + luminaria_escolhida + "D4"

    df.at[idx, 'angulo_escolhido'] = angulo_x

          
    if(altura_modificada == True): #
        df.at[idx, 'nova_altura'] = altura_float
    else:
        df.at[idx, 'nova_altura'] = "Sem alterações"
 
    if(braco_modificado_check == True):
        df.at[idx, 'novo_braco'] = novo_braco_eficientiza
    else:
        df.at[idx, 'novo_braco'] = "Sem alterações"

    # Garantir que a coluna 'cenario' é do tipo object
    df['cenario'] = df['cenario'].astype(object)

    df.at[idx, 'cenario'] = modify_name

    # Identificar colunas a serem removidas
    colunas_remover = [col for col in df.columns if 'Material' in col]

    # Remover as colunas indesejadas
    df_limpo = df.drop(columns=colunas_remover)

    # Salvar a planilha atualizada  
    df_limpo.to_excel('Atualizada_Cadastro_Piloto_SM.xlsx', sheet_name='Cadastro IPSM', index=False)

    #colocar nova altura de instalação na planilha