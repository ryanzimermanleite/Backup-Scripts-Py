# -*- coding: utf-8 -*-
# ----------------------------------------------------------------------#
# Nome:     Robo Envia ECD                                              #
# Arquivo:  robo_ecd.py                                                 #
# Versão:   1.0.0                                                       #
# Modulo:   Dominio                                                     #
# Objetivo: Enviar ECD pro Leo                                          #
# Autor:    Ryan Zimerman Leite                                         #
# Data:     11/12/2023                                                  #
# ----------------------------------------------------------------------#

import pyperclip, time, pyautogui as p
import os, re
import PySimpleGUI as sg
from functools import wraps
from threading import Thread


def wait_img(img, pasta='imgs', conf=1.0, delay=1, timeout=20, debug=False):
    if debug:
        print('\tEsperando', img)

    aux = 0
    while True:
        box = find_img(img, pasta, conf=conf)
        if box:
            return box
        time.sleep(delay)

        if timeout < 0:
            continue
        if timeout == aux:
            break
        aux += 1

    return None


def click_img(img, pasta='imgs', conf=1.0, delay=1, timeout=20, button='left', clicks=1):
    img = os.path.join(pasta, img)
    for i in range(timeout):
        box = p.locateCenterOnScreen(img, confidence=conf)
        if box:
            p.click(p.locateCenterOnScreen(img, confidence=conf), button=button, clicks=clicks)
            return True
        time.sleep(delay)
    else:
        return False


def find_img(img, pasta='imgs', conf=1.0):
    path = os.path.join(pasta, img)
    return p.locateOnScreen(path, confidence=conf)


def login(empresa, window):
    # espera a tela inicial do domínio
    while not find_img('inicial.png', pasta='imgs_c', conf=0.9):
        time.sleep(1)

    p.click(833, 384)

    # espera abrir a janela de seleção de empresa
    while not find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
        p.press('f8')

    time.sleep(1)
    # clica para pesquisar empresa por código
    if find_img('codigo.png', pasta='imgs_c', conf=0.9):
        p.click(p.locateCenterOnScreen(r'imgs_c\codigo.png', confidence=0.9))
    p.write(empresa)
    time.sleep(3)

    # confirmar empresa
    p.hotkey('alt', 'a')
    # enquanto a janela estiver aberta verifica exceções
    while find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
        time.sleep(1)
        if find_img('sem_parametro.png', pasta='imgs_c', conf=0.9):
            window['-Mensagens-'].update('Parametro não cadastrado para esta empresa!')
            print('❌ Parametro não cadastrado para esta empresa')
            p.press('enter')
            time.sleep(1)
            while not find_img('parametros.png', pasta='imgs_c', conf=0.9):
                time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return False

        if find_img('nao_existe_parametro.png', pasta='imgs_c', conf=0.9):
            window['-Mensagens-'].update('Não existe parametro cadastrado para esta empresa!')
            print('❌ Não existe parametro cadastrado para esta empresa')
            p.press('enter')
            time.sleep(1)
            p.hotkey('alt', 'n')
            time.sleep(1)
            p.press('esc')
            time.sleep(1)
            p.hotkey('alt', 'n')
            while find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
                time.sleep(1)
            return False

        if find_img('empresa_nao_usa_sistema.png', pasta='imgs_c', conf=0.9) or find_img(
                'empresa_nao_usa_sistema_2.png', pasta='imgs_c', conf=0.9):
            window['-Mensagens-'].update('Empresa não está marcada para usar este sistema!')
            print('❌ Empresa não está marcada para usar este sistema')
            p.press('enter')
            time.sleep(1)
            p.press('esc', presses=5)
            while find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
                time.sleep(1)
            return False

        if find_img('fase_dois_do_cadastro.png', pasta='imgs_c', conf=0.9):
            p.hotkey('alt', 'n')
            time.sleep(1)
            p.hotkey('alt', 'n')

        if find_img('conforme_modulo.png', pasta='imgs_c', conf=0.9):
            p.press('enter')
            time.sleep(1)

        if find_img('aviso_regime.png', pasta='imgs_c', conf=0.9):
            p.hotkey('alt', 'n')
            time.sleep(1)

        if find_img('aviso.png', pasta='imgs_c', conf=0.9):
            p.hotkey('alt', 'o')
            time.sleep(1)

        if find_img('erro_troca_empresa.png', pasta='imgs_c', conf=0.9):
            p.press('enter')
            time.sleep(1)
            p.press('esc', presses=5, interval=1)
            login(empresa, window)

    if not verifica_empresa(empresa):
        window['-Mensagens-'].update('Empresa não encontrada!')
        print('❌ Empresa não encontrada')
        p.press('esc')
        return False

    p.press('esc', presses=5)
    time.sleep(1)

    return True


def verifica_empresa(empresa):
    erro = 'sim'
    while erro == 'sim':
        try:
            p.click(1258, 82)

            while True:
                try:
                    time.sleep(1)
                    p.hotkey('ctrl', 'c')
                    time.sleep(1)
                    p.hotkey('ctrl', 'c')
                    time.sleep(1)
                    cnpj_codigo = pyperclip.paste()
                    break
                except:
                    pass

            time.sleep(0.5)
            codigo = cnpj_codigo.split('-')
            codigo = str(codigo[1])
            codigo = codigo.replace(' ', '')

            if codigo != empresa:
                print(f'Código da empresa: {codigo}')
                print(f'Código encontrado no Domínio: {empresa}')
                return False
            else:
                return True
        except:
            erro = 'sim'


def barra_de_status(func):
    @wraps(func)
    def wrapper():
        sg.theme('GrayGrayGray')  # Define o tema do PySimpleGUI
        # sg.theme_previewer()
        # Layout da janela
        layout = [
            [sg.Text('Codigo:'),
             sg.Input(key='-codigo-', size=(4, 1)),
             sg.Text('Inicio:'),
             sg.Input(key='-data_inicio-', size=(9, 1)),
             sg.Text('Final:'),
             sg.Input(key='-data_final-', size=(9, 1)),
             sg.Text('Livro:'),
             sg.Input(key='-livro-', size=(2, 2)),
             sg.Text('Hash:'),
             sg.Input(key='-hash-', size=(52, 3), font=("Helvetica", 8)),
             sg.Radio("NORMAL", "ecd", key="-normal-", font=("Helvetica", 9)),
             sg.Radio("RETIFICAR", "ecd", key="-retificar-", font=("Helvetica", 9)),
             sg.Text('|'),
             sg.Button('RUN', key='-iniciar-', border_width=0, button_color='green1'),
             sg.Button('STOP', key='-stop-', border_width=0, button_color='yellow'),
             sg.Button('EXIT', key='-exit-', border_width=0, button_color='red'),
             sg.Text('', key='-Mensagens-', size=100)],
        ]

        # guarda a janela na variável para manipula-la
        screen_width, screen_height = sg.Window.get_screen_size()
        window = sg.Window('', layout, no_titlebar=True, location=(0, 0), size=(screen_width, 35), keep_on_top=True)

        def run_script_thread():
            try:
                # habilita e desabilita os botões conforme necessário
                window['-iniciar-'].update(disabled=True)

                # Chama a função que executa o script

                func(window, values)

                # habilita e desabilita os botões conforme necessário
                window['-iniciar-'].update(disabled=False)

                # apaga qualquer mensagem na interface
            # window['-Mensagens-'].update('')
            except:
                pass

        while True:
            # captura o evento e os valores armazenados na interface
            event, values = window.read()

            if event == sg.WIN_CLOSED or event == '-exit-':
                break

            elif event == '-iniciar-':
                # Cria uma nova thread para executar o script
                script_thread = Thread(target=run_script_thread)
                script_thread.start()

        window.close()

    return wrapper


def ecd_retificar(window, values):
    window['-Mensagens-'].update('ECD Retificar em execução!!')
    login(values['-codigo-'], window)
    time.sleep(1)

    while not find_img('sped_contabil.png', conf=0.9):
        p.hotkey('alt', 'r')
        time.sleep(0.5)
        p.press('f')
        time.sleep(0.5)
        p.press('s')
        time.sleep(2)
    time.sleep(1)
    p.write(str(values['-data_inicio-']))
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.write(str(values['-data_final-']))
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.write('T:\# ecd_robo\Arquivos TXT RETIFICAR')
    time.sleep(0.5)
    p.press('backspace', presses=40)
    p.write('T:\# ecd_robo\Arquivos TXT RETIFICAR')
    time.sleep(0.5)
    p.hotkey('alt', 'd')
    time.sleep(0.5)
    wait_img('outros_dados.png', conf=0.9, timeout=-1)
    time.sleep(0.5)
    p.press('tab', presses=3, interval=0.1)
    time.sleep(0.5)
    p.press('down')
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('down', presses=5, interval=0.1)
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    hash_num = str(values['-hash-']).replace('.', '')
    time.sleep(1)
    p.write(hash_num)
    time.sleep(0.5)
    p.press('tab', presses=2, interval=0.1)
    p.write(str(values['-data_final-']))
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.write(str(values['-data_final-']))
    time.sleep(1)

    if find_img('gerar_movimento.png', conf=0.9):
        p.press('tab', presses=11, interval=0.1)
    else:
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('space')
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('space')
        time.sleep(0.5)
        p.press('tab', presses=7, interval=0.1)

    time.sleep(0.5)
    p.press('right')
    time.sleep(0.5)
    p.press('right')
    time.sleep(0.5)
    p.press('right')
    time.sleep(0.5)
    p.press('backspace')
    time.sleep(0.5)
    p.press('backspace')
    time.sleep(0.5)
    p.press('right')
    time.sleep(0.5)
    p.press('right')
    p.write(str(values['-livro-']))
    time.sleep(1)

    if find_img('sem_evandro.png', conf=0.9):
        time.sleep(0.5)
        p.hotkey('alt', 'n')
        time.sleep(0.5)
        p.write('30782876889')
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        wait_img('importar_cadastro.png', conf=0.9, timeout=-1)
        time.sleep(0.5)
        p.hotkey('alt', 'y')
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('down', presses=15, interval=0.1)
        time.sleep(0.5)
        while not find_img('tela_1.png', conf=0.9):
            time.sleep(0.5)
            click_img('demonstrativos.png', conf=0.9)
            time.sleep(0.5)
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('right')
        time.sleep(0.5)

    else:
        while not find_img('tela_1.png', conf=0.9):
            time.sleep(0.5)
            click_img('demonstrativos.png', conf=0.9)
            time.sleep(0.5)
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('right')
        time.sleep(0.5)

    if find_img('gerar_img.png', conf=0.9):
        click_img('gerar.png', conf=0.9)
        time.sleep(0.5)
        p.press('right')
    else:
        click_img('gerar.png', conf=0.9)
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('space')
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('space')
        time.sleep(0.5)
        p.press('tab', presses=9, interval=0.1)
        time.sleep(0.5)
        p.press('right')

    time.sleep(1)

    while not find_img('selecione_arquivo.png', conf=0.9):
        time.sleep(1)
        click_img('3pontos_new.png', conf=0.9)
        time.sleep(1)
    time.sleep(2)

    while not find_img('arquivos_rtf.png', conf=0.9):
        time.sleep(0.5)
        p.press('tab', presses=4, interval=0.1)
        time.sleep(0.5)
        p.press('down')
        time.sleep(0.5)
        click_img('cliente_t.png', conf=0.9)
        time.sleep(0.5)
        wait_img('ecd_robo.png', conf=0.9, timeout=-1)
        time.sleep(0.5)
        click_img('ecd_robo.png', conf=0.9, clicks=2)
        time.sleep(0.5)
        wait_img('pasta_rtf.png', conf=0.9, timeout=-1)
        time.sleep(0.5)
        click_img('pasta_rtf.png', conf=0.9, clicks=2)
        time.sleep(1)
        p.press('tab')
        time.sleep(0.5)
        p.press('tab')
        time.sleep(1)
    time.sleep(1)
    p.write('ECD RETIFICAR_' + str(values['-codigo-']))
    time.sleep(0.5)
    p.press('down')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(1)
    p.press('tab', presses=6, interval=0.1)
    time.sleep(0.5)
    p.press('right')
    time.sleep(0.5)
    if find_img('opcoes.png', conf=0.9):
        p.hotkey('alt', 'o')
    else:
        p.press('tab', presses=2, interval=0.1)
        time.sleep(0.5)
        p.press('space')
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('space')
        time.sleep(0.5)
        p.hotkey('alt', 'o')

    while not find_img('spe2.png', conf=0.9):
        time.sleep(1)
        if find_img('resp.png', conf=0.9):
            time.sleep(1)
            p.press('enter')
            time.sleep(1)
        time.sleep(1)

        if find_img('todas_contas.png', conf=0.9):
            time.sleep(1)
            p.press('tab')
            time.sleep(1)
            p.press('enter')
        time.sleep(1)
    time.sleep(1)

    p.hotkey('alt', 'o')
    time.sleep(1)
    # DEU CERTO
    while not find_img('final_exportacao.png', conf=0.9):
        time.sleep(1)
        if find_img('grupo_contas.png', conf=0.9):
            time.sleep(1)
            p.hotkey('alt', 's')
            time.sleep(1)
        elif find_img('validacao.png', conf=0.9):
            time.sleep(1)
            p.hotkey('alt', 's')
            time.sleep(1)

    time.sleep(1)

    window['-Mensagens-'].update('ECD Retificar Finalizado!!!')


def ecd_normal(window, values, codigo):
    window['-Mensagens-'].update('ECD Normal em execução!!')
    login(values['-codigo-'], window)
    time.sleep(1)

    while not find_img('sped_contabil.png', conf=0.9):
        p.hotkey('alt', 'r')
        time.sleep(0.5)
        p.press('f')
        time.sleep(0.5)
        p.press('s')
        time.sleep(2)
    time.sleep(1)
    p.write(str(values['-data_inicio-']))
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.write(str(values['-data_final-']))
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.write('T:\# ecd_robo\Arquivos TXT NORMAL')
    time.sleep(0.5)
    p.press('backspace', presses=40)
    p.write('T:\# ecd_robo\Arquivos TXT NORMAL')
    time.sleep(0.5)
    p.hotkey('alt', 'd')
    time.sleep(0.5)
    wait_img('outros_dados.png', conf=0.9, timeout=-1)
    time.sleep(0.5)
    p.press('tab', presses=3, interval=0.2)
    time.sleep(0.5)
    p.press('up', presses=2)
    time.sleep(0.5)
    p.press('tab', presses=2, interval=0.2)
    time.sleep(0.5)
    p.write(str(values['-data_final-']))
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.write(str(values['-data_final-']))
    time.sleep(1)

    if find_img('gerar_movimento.png', conf=0.9):
        p.press('tab', presses=11, interval=0.1)
    else:
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('space')
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('space')
        time.sleep(0.5)
        p.press('tab', presses=7, interval=0.1)

    time.sleep(0.5)
    p.press('right')
    time.sleep(0.5)
    p.press('right')
    time.sleep(0.5)
    p.press('right')
    time.sleep(0.5)
    p.press('backspace')
    time.sleep(0.5)
    p.press('backspace')
    time.sleep(0.5)
    p.press('right')
    time.sleep(0.5)
    p.press('right')
    p.write(str(values['-livro-']))
    time.sleep(1)

    while not find_img('sem_evandro.png', conf=0.9):
        time.sleep(0.5)
        p.hotkey('alt', 'x')
        time.sleep(0.5)
    time.sleep(1)

    while not find_img('tela_1.png', conf=0.9):
        time.sleep(0.5)
        click_img('demonstrativos.png', conf=0.9)
        time.sleep(0.5)
    time.sleep(0.5)
    p.press('tab')
    time.sleep(1)

    if find_img('gerar_img.png', conf=0.9):
        time.sleep(0.5)
        click_img('op.png', conf=0.9)
        time.sleep(0.5)
    else:
        click_img('gerar.png', conf=0.9)
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('space')
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('space')
        time.sleep(0.5)
        click_img('op.png', conf=0.9)
        time.sleep(0.5)

    if find_img('opcoes.png', conf=0.9):
        p.hotkey('alt', 'o')
    else:
        p.press('tab')
        time.sleep(0.5)
        p.press('space')
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('space')
        p.hotkey('alt', 'o')
#
    time.sleep(5)

    if find_img('grupo_contas_normal.png', conf=0.9):
        time.sleep(1)
        p.hotkey('alt', 'n')  #
    time.sleep(3)

    p.hotkey('alt', 'o')

    while not find_img('validacao.png', conf=0.9):
        time.sleep(0.5)

    p.hotkey('alt', 's')

    time.sleep(1)
    while not find_img('final_exportacao2.png', conf=0.9):
        time.sleep(0.5)
    p.press('enter')
    time.sleep(1)

    window['-Mensagens-'].update('ECD Normal Finalizada!!')
    sped_normal('normal', codigo)
##
def sped_normal(tipo, codigo):
    click_img('sped_icone.png', conf=0.9)

    while not find_img('sped_titulo.png', conf=0.9):
        time.sleep(0.5)
        click_img('sped_icone.png', conf=0.9)
        time.sleep(1)

    while not find_img('sped_importar_escrituracao.png', conf=0.9):
        time.sleep(0.5)
        click_img('sped_icone_importar.png', conf=0.9)


    p.press('tab', presses=4, interval=0.2)
    time.sleep(0.5)
    p.press('down')
    time.sleep(0.5)
    p.press('up', presses=15)
    time.sleep(0.5)
    while not find_img('comum.png', conf=0.9):
        time.sleep(0.2)#
        p.press('down')

    click_img('comum2.png', conf=0.9)#

    while not find_img('img_ecd.png', conf=0.9):
        time.sleep(0.5)

    click_img('img_ecd.png', conf=0.9, clicks=2)

    while not find_img('ecd_robo2.png', conf=0.9):
        time.sleep(0.5)
    if tipo == 'normal':
        click_img('arquivo_txt_normal.png', conf=0.9, clicks=2)
    elif tipo == 'retificar':
        click_img('arquivo_txt_retificar.png', conf=0.9, clicks=2)
    arquivo = 'sped_diario0' + codigo + '.txt'
    time.sleep(1)
    p.press('tab')#
    time.sleep(0.5)
    p.write(arquivo)
    time.sleep(0.5)
    p.press('enter')

    while not find_img('importacao_blocos.png', conf=0.9):
        time.sleep(0.5)

    click_img('ok.png', conf=0.9)

    while not find_img('sped_aviso_sucesso.png', conf=0.9):
        time.sleep(0.5)

    p.hotkey('alt', 'n')

    while not find_img('passo_a_passo.png', conf=0.9):
        time.sleep(0.5)

    click_img('escrituracao.png', conf=0.9)
    time.sleep(0.5)
    click_img('recuperar_ecd.png', conf=0.9)

    while not find_img('titulo_recuperar_ecd.png', conf=0.9):
        time.sleep(0.5)

    click_img('localizar.png', conf=0.9)

    while not find_img('abrir_recuperar_ecd.png', conf=0.9):
        time.sleep(0.5)
    click_img('clique.png', conf=0.9)
    time.sleep(0.5)
    p.press('up', presses=15)
    time.sleep(0.5)
    while not find_img('comum.png', conf=0.9):
        time.sleep(0.2)
        p.press('down')

    while not find_img('comum5.png', conf=0.9):
        time.sleep(0.5)
    click_img('comum5.png', conf=0.9, clicks=2)

    while not find_img('img_ecd.png', conf=0.9):
        time.sleep(0.5)

    extrai_data_nome_txt(arquivo, codigo)
###

def extrai_data_nome_txt(arquivo, codigo):
    pasta = 'T:\# ecd_robo\Arquivos TXT NORMAL\\' + arquivo

    with open(pasta, 'r', errors="ignore") as f:
        dados = f.read()
    dados_arquivo = re.findall(r'\|([^|]*)', dados)
    data = dados_arquivo[3]
    razao = dados_arquivo[4]
    data = data[4:]
    razao = razao[:20]

    p.write(razao)
    time.sleep(0.5)
    p.press('enter')

    while not find_img('dp_contabil.png', conf=0.9):
        time.sleep(1)

    click_img('dp_contabil.png', conf=0.9, clicks=2)

    while not find_img('2022.png', conf=0.9):
        time.sleep(0.5)

    p.press('tab')
    time.sleep(0.5)

    data = int(data) - 1
    if data == 2020:
        time.sleep(0.5)
        pyperclip.copy('SPED CONTABIL 2020')
        time.sleep(0.5)
        p.hotkey('ctrl', 'v')
        time.sleep(0.5)

    if data == 2021:
        time.sleep(0.5)
        pyperclip.copy('SPED CONTABIL 2021')
        time.sleep(0.5)
        p.hotkey('ctrl', 'v')
        time.sleep(0.5)

    if data == 2022:
        time.sleep(0.5)
        pyperclip.copy('SPED CONTABIL 2022')
        time.sleep(0.5)
        p.hotkey('ctrl', 'v')
        time.sleep(0.5)

    p.press('enter')
    time.sleep(2)
    p.press('tab', presses=10, interval=0.2)

    while not find_img('txt.png', conf=0.9):
        time.sleep(0.5)
        p.press('down')
        time.sleep(0.5)
        p.press('up')
        time.sleep(0.5)
        p.press('enter')
        time.sleep(0.5)

    time.sleep(0.5)
    p.press('down')
    time.sleep(0.5)
    p.press('up')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)

    while not find_img('sucesso_ecd.png', conf=0.9):
        time.sleep(0.5)

    p.press('enter')
    time.sleep(1)
    click_img('botao_recuperar.png', conf=0.9, clicks=1)

    while not find_img('recuperado.png', conf=0.9):
        time.sleep(0.5)

    p.press('enter')
    time.sleep(0.5)
    click_img('escrituracao.png', conf=0.9)
    time.sleep(0.5)
    click_img('dados_escrituracao.png', conf=0.9)
    time.sleep(1)
    click_img('sped_escrituracao.png', conf=0.9)
    time.sleep(1)
    click_img('4_sig.png', conf=0.9)


    while not find_img('procurador.png', conf=0.9):
        time.sleep(0.5)

    click_img('procurador.png', conf=0.9)
    time.sleep(0.5)
    click_img('remover.png', conf=0.9)
    while not find_img('excluir.png', conf=0.9):
        time.sleep(0.5)
    p.hotkey('alt', 's')
    time.sleep(1)
    click_img('rpem.png', conf=0.9, clicks=2)
    while not find_img('registro.png', conf=0.9):
        time.sleep(0.5)
    p.press('tab', presses=12, interval=0.2)
    time.sleep(0.5)
    p.press('del')
    time.sleep(0.5)
    p.write('S')
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    click_img('salvar.png', conf=0.9)
    time.sleep(1)
    print('aeeeeeeeeeeeeeeeeeee')
    click_img('passo_a_passo2.png', conf=0.9)
    time.sleep(1)
    click_img('botao_validar.png', conf=0.9)
    while not find_img('leitor_pdf.png', conf=0.9):
        time.sleep(0.5)

    if find_img('validar_erro.png', conf=0.9):
        p.press('enter')
        time.sleep(1)
        click_img('fechar.png', conf=0.9)
        time.sleep(1)
        click_img('escrituracao.png', conf=0.9)###
        time.sleep(0.5)
        click_img('gerar_copia.png', conf=0.9)
        while not find_img('titulo_seguranca.png', conf=0.9):
            time.sleep(0.5)
        p.press('tab', presses=4, interval=0.2)
        time.sleep(0.5)
        p.press('down')
        time.sleep(0.5)
        p.press('up', presses=15)
        time.sleep(0.5)
        while not find_img('comum.png', conf=0.9):
            time.sleep(0.2)  #
            p.press('down')
        click_img('comum2.png', conf=0.9)  #
        while not find_img('img_ecd.png', conf=0.9):
            time.sleep(0.5)
        click_img('img_ecd.png', conf=0.9, clicks=2)

        time.sleep(2)
        time.sleep(0.5)
        click_img('pasta_copia.png', conf=0.9, clicks=2)

        while not find_img('ecd_robo3.png', conf=0.9):##
            print('aaaaaa')
            time.sleep(0.5)

        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)#
        data = data + 1
        pyperclip.copy('ECD ' + codigo + ' - ' + str(data))
        time.sleep(0.5)
        p.hotkey('ctrl', 'v')
        time.sleep(0.5)
        click_img('gravar.png', conf=0.9)


    print('fim!!!')###
@barra_de_status
def run(window, values):
    if values['-normal-'] == False and values['-retificar-'] == False:
        window['-Mensagens-'].update('Modelo ECD não Informado!')
    if values['-normal-'] == True:
        if values['-codigo-'] == '':
            window['-Mensagens-'].update('Codigo empresa não informado!')
        elif values['-data_inicio-'] == '':
            window['-Mensagens-'].update('Data Inicio não informado!')
        elif values['-data_final-'] == '':
            window['-Mensagens-'].update('Data Final não informado!')
        elif values['-livro-'] == '':
            window['-Mensagens-'].update('Número do Livro não informado!')
        else:
            codigo = values['-codigo-']
            ecd_normal(window, values, codigo)

    elif values['-retificar-'] == True:
        if values['-codigo-'] == '':
            window['-Mensagens-'].update('Codigo empresa não informado!')
        elif values['-data_inicio-'] == '':
            window['-Mensagens-'].update('Data Inicio não informado!')
        elif values['-data_final-'] == '':
            window['-Mensagens-'].update('Data Final não informado!')
        elif values['-livro-'] == '':
            window['-Mensagens-'].update('Número do Livro não informado!')
        elif values['-hash-'] == '':
            window['-Mensagens-'].update('Hash não informado!')
        else:
            ecd_retificar(window, values)
    print('Robo Em Finalizado!')

if __name__ == '__main__':
    print('Robo Em Execução!')
    run()