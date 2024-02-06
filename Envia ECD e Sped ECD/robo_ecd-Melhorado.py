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
from shutil import move
import pyperclip, time, os, re, pyautogui as p
import PySimpleGUI as sg
from functools import wraps
from threading import Thread
from sys import path
path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img


def login(empresa, window):
    # espera a tela inicial do domínio
    while not _find_img('inicial.png', pasta='imgs_c', conf=0.9):
        time.sleep(1)

    p.click(833, 384)

    # espera abrir a janela de seleção de empresa
    while not _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
        p.press('f8')

    time.sleep(1)
    # clica para pesquisar empresa por código
    if _find_img('codigo.png', pasta='imgs_c', conf=0.9):
        p.click(p.locateCenterOnScreen(r'imgs_c\codigo.png', confidence=0.9))
    p.write(empresa)
    time.sleep(3)

    # confirmar empresa
    p.hotkey('alt', 'a')
    # enquanto a janela estiver aberta verifica exceções
    while _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
        time.sleep(1)
        if _find_img('sem_parametro.png', pasta='imgs_c', conf=0.9):
            window['-Mensagens-'].update('Parametro não cadastrado para esta empresa!')
            print('❌ Parametro não cadastrado para esta empresa')
            p.press('enter')
            time.sleep(1)
            while not _find_img('parametros.png', pasta='imgs_c', conf=0.9):
                time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return False

        if _find_img('nao_existe_parametro.png', pasta='imgs_c', conf=0.9):
            window['-Mensagens-'].update('Não existe parametro cadastrado para esta empresa!')
            print('❌ Não existe parametro cadastrado para esta empresa')
            p.press('enter')
            time.sleep(1)
            p.hotkey('alt', 'n')
            time.sleep(1)
            p.press('esc')
            time.sleep(1)
            p.hotkey('alt', 'n')
            while _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
                time.sleep(1)
            return False

        if _find_img('empresa_nao_usa_sistema.png', pasta='imgs_c', conf=0.9) or _find_img(
                'empresa_nao_usa_sistema_2.png', pasta='imgs_c', conf=0.9):
            window['-Mensagens-'].update('Empresa não está marcada para usar este sistema!')
            print('❌ Empresa não está marcada para usar este sistema')
            p.press('enter')
            time.sleep(1)
            p.press('esc', presses=5)
            while _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
                time.sleep(1)
            return False

        if _find_img('fase_dois_do_cadastro.png', pasta='imgs_c', conf=0.9):
            p.hotkey('alt', 'n')
            time.sleep(1)
            p.hotkey('alt', 'n')

        if _find_img('conforme_modulo.png', pasta='imgs_c', conf=0.9):
            p.press('enter')
            time.sleep(1)

        if _find_img('aviso_regime.png', pasta='imgs_c', conf=0.9):
            p.hotkey('alt', 'n')
            time.sleep(1)

        if _find_img('aviso.png', pasta='imgs_c', conf=0.9):
            p.hotkey('alt', 'o')
            time.sleep(1)

        if _find_img('erro_troca_empresa.png', pasta='imgs_c', conf=0.9):
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

    while not _find_img('sped_contabil.png', conf=0.9):
        p.hotkey('alt', 'r', interval=0.5)
        p.press('f', interval=0.5)
        p.press('s', interval=2)

    time.sleep(1)
    p.write(str(values['-data_inicio-']))
    p.press('tab', interval=0.5)
    p.write(str(values['-data_final-']))
    p.press('tab', interval=0.5)
    p.write('T:\# ecd_robo\Arquivos TXT RETIFICAR')
    p.press('backspace', presses=40)
    p.write('T:\# ecd_robo\Arquivos TXT RETIFICAR')
    p.hotkey('alt', 'd', interval=0.5)
    _wait_img('outros_dados.png', conf=0.9, timeout=-1)
    p.press('tab', presses=3, interval=0.5)
    p.press('down', interval=0.5)
    p.press('tab', interval=0.5)
    p.press('down', presses=5, interval=0.1)
    time.sleep(0.5)
    p.press('tab', interval=0.5)
    hash_num = str(values['-hash-']).replace('.', '')
    p.write(hash_num)
    p.press('tab', presses=2, interval=0.5)
    p.write(str(values['-data_final-']))
    p.press('tab', interval=0.5)
    p.write(str(values['-data_final-']))
    time.sleep(1)

    if _find_img('gerar_movimento.png', conf=0.9):
        p.press('tab', presses=11, interval=0.1)
    else:
        p.press('tab',presses=2, interval=0.5)
        p.press('space', interval=0.5)
        p.press('tab', presses=2, interval=0.5)
        p.press('space', interval=0.5)
        p.press('tab', presses=7, interval=0.1)

    p.press('right', presses=3, interval=0.5)
    p.press('backspace',presses=2, interval=0.5)
    p.press('right',presses=2, interval=0.5)
    p.write(str(values['-livro-']))
    time.sleep(2)

    if _find_img('sem_evandro.png', conf=0.9):
        p.hotkey('alt', 'n', interval=0.5)
        p.write('30782876889')
        p.press('tab', interval=0.5)
        _wait_img('importar_cadastro.png', conf=0.9, timeout=-1)
        time.sleep(0.5)
        p.hotkey('alt', 'y', interval=0.5)

        p.press('tab', interval=0.5)
        p.press('down', presses=15, interval=0.1)
        time.sleep(0.5)
        while not _find_img('tela_1.png', conf=0.9):
            _click_img('demonstrativos.png', conf=0.9, interval=0.5)
        time.sleep(0.5)
        p.press('tab', interval=0.5)
        p.press('right', interval=0.5)


    else:
        while not _find_img('tela_1.png', conf=0.9):
            _click_img('demonstrativos.png', conf=0.9, interval=0.5)

        time.sleep(0.5)
        p.press('tab', interval=0.5)
        p.press('right', interval=0.5)

    if _find_img('gerar_img.png', conf=0.9):
        _click_img('gerar.png', conf=0.9, interval=0.5)
        p.press('right')
    else:
        _click_img('gerar.png', conf=0.9, interval=0.5)
        p.press('tab', interval=0.5)
        p.press('space', interval=0.5)
        p.press('tab', interval=0.5)
        p.press('space', interval=0.5)
        p.press('tab', presses=9, interval=0.1)
        p.press('right')

    time.sleep(1)

    while not _find_img('selecione_arquivo.png', conf=0.9):
        _click_img('3pontos_new.png', conf=0.9, interval=1)
    time.sleep(2)

    while not _find_img('arquivos_rtf.png', conf=0.9):
        time.sleep(0.5)
        p.press('tab', presses=4, interval=0.5)
        p.press('down', interval=0.5)
        _click_img('cliente_t.png', conf=0.9)
        _wait_img('ecd_robo.png', conf=0.9, timeout=-1)
        _click_img('ecd_robo.png', conf=0.9, clicks=2)
        _wait_img('pasta_rtf.png', conf=0.9, timeout=-1)
        _click_img('pasta_rtf.png', conf=0.9, clicks=2)
        time.sleep(1)
        p.press('tab', presses=2, interval=0.5)
        time.sleep(1)
    time.sleep(1)
    p.write('ECD RETIFICAR_' + str(values['-codigo-']))
    p.press('down', interval=0.5)
    p.press('enter', interval=1)
    p.press('tab', presses=6, interval=0.5)
    p.press('right', interval=0.5)
    if _find_img('opcoes.png', conf=0.9):
        p.hotkey('alt', 'o')
    else:
        p.press('tab', presses=2, interval=0.5)
        p.press('space', interval=0.5)
        p.press('tab', interval=0.5)
        p.press('space', interval=0.5)
        p.hotkey('alt', 'o')

    while not _find_img('spe2.png', conf=0.9):
        time.sleep(1)
        if _find_img('resp.png', conf=0.9):
            p.press('enter', interval=1)
        time.sleep(1)

        if _find_img('todas_contas.png', conf=0.9):
            p.press('tab', interval=1)
            p.press('enter', interval=1)
        time.sleep(1)
    time.sleep(1)

    p.hotkey('alt', 'o', interval=0.5)

    while not _find_img('final_exportacao.png', conf=0.9):
        time.sleep(1)
        if _find_img('grupo_contas.png', conf=0.9):
            p.hotkey('alt', 's', interval=1)
        elif _find_img('validacao.png', conf=0.9):
            p.hotkey('alt', 's', interval=1)
    time.sleep(1)

    window['-Mensagens-'].update('ECD Retificar Finalizado!!!')


def ecd_normal(window, values, codigo):
    window['-Mensagens-'].update('ECD Normal em execução!!')
    login(values['-codigo-'], window)
    time.sleep(1)

    while not _find_img('sped_contabil.png', conf=0.9):
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
    _wait_img('outros_dados.png', conf=0.9, timeout=-1)
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

    if _find_img('gerar_movimento.png', conf=0.9):
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
    p.press('backspace')##
    time.sleep(0.5)
    p.press('right')
    time.sleep(0.5)
    p.press('right')
    p.write(str(values['-livro-']))
    time.sleep(1)

    while not _find_img('sem_evandro.png', conf=0.9):
        time.sleep(0.5)
        p.hotkey('alt', 'x')
        time.sleep(0.5)
    time.sleep(1)

    while not _find_img('tela_1.png', conf=0.9):
        time.sleep(0.5)
        _click_img('demonstrativos.png', conf=0.9)
        time.sleep(0.5)
    time.sleep(0.5)
    p.press('tab')
    time.sleep(1)

    if _find_img('gerar_img.png', conf=0.9):
        time.sleep(0.5)
        _click_img('op.png', conf=0.9)
        time.sleep(0.5)
    else:
        _click_img('gerar.png', conf=0.9)
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('space')
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('space')
        time.sleep(0.5)
        _click_img('op.png', conf=0.9)
        time.sleep(0.5)

    if _find_img('opcoes.png', conf=0.9):
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

    if _find_img('grupo_contas_normal.png', conf=0.9):

        time.sleep(1)
        p.hotkey('alt', 'n')
    time.sleep(3)

    p.hotkey('alt', 'o')

    time.sleep(3)

    if _find_img('validacao.png', conf=0.9):
        time.sleep(0.5)
        p.hotkey('alt', 's')
    time.sleep(1)

    while not _find_img('final_exportacao2.png', conf=0.9):
        time.sleep(0.5)
    p.press('enter')
    time.sleep(1)#

    window['-Mensagens-'].update('ECD Normal Finalizada!!')
    sped_normal('normal', codigo)
##
def sped_normal(tipo, codigo):
    _click_img('sped_icone.png', conf=0.9)
    time.sleep(2)
    while not _find_img('sped_titulo.png', conf=0.9):
        time.sleep(1)
        _click_img('sped_icone.png', conf=0.9)
        time.sleep(2)

    while not _find_img('sped_importar_escrituracao.png', conf=0.9):#
        time.sleep(0.5)
        _click_img('sped_icone_importar.png', conf=0.9)

    p.press('tab', presses=4, interval=0.2)
    p.press('down', interval=0.5)
    p.press('up', presses=15)

    while not _find_img('comum.png', conf=0.9):
        time.sleep(0.2)
        p.press('down')

    _click_img('comum2.png', conf=0.9)#
    _click_img('img_ecd.png', conf=0.9, clicks=2)
    _wait_img('ecd_robo2.png', conf=0.9, timeout=-1)

    if tipo == 'normal':
        _click_img('arquivo_txt_normal.png', conf=0.9, clicks=2)
    elif tipo == 'retificar':
        _click_img('arquivo_txt_retificar.png', conf=0.9, clicks=2)

    arquivo = 'sped_diario0' + codigo + '.txt'
    time.sleep(1)
    p.press('tab')
    p.write(arquivo)
    p.press('enter')
    _wait_img('importacao_blocos.png', conf=0.9, timeout=-1)
    _click_img('ok.png', conf=0.9)
    _wait_img('sped_aviso_sucesso.png', conf=0.9, timeout=-1)
    p.hotkey('alt', 'n')
    _wait_img('passo_a_passo.png', conf=0.9, timeout=-1)
    _click_img('escrituracao.png', conf=0.9)
    _click_img('recuperar_ecd.png', conf=0.9)
    _wait_img('titulo_recuperar_ecd.png', conf=0.9, timeout=-1)
    extrai_data_nome_txt(arquivo, codigo)

def extrai_data_nome_txt(arquivo, codigo):
    pasta = 'T:\# ecd_robo\Arquivos TXT NORMAL\\' + arquivo

    with open(pasta, 'r', errors="ignore") as f:
        dados = f.read()
    dados_arquivo = re.findall(r'\|([^|]*)', dados)
    data = dados_arquivo[3]
    data_fim = dados_arquivo[3]
    data_inicio = dados_arquivo[2]
    razao = dados_arquivo[4]
    cnpj = dados_arquivo[5]
    data = data[4:]
    razao = razao[:10]
    nome_arquivo2 = gera_arquivo_receitabx(cnpj, data_inicio, data_fim, razao)
    _click_img('localizar.png', conf=0.9)
    _wait_img('abrir_recuperar_ecd.png', conf=0.9, timeout=-1)
    _click_img('clique.png', conf=0.9)
    p.press('up', presses=15)
#
    while not _find_img('comum.png', conf=0.9):
        time.sleep(0.2)
        p.press('down')

    _click_img('comum5.png', conf=0.9, clicks=2)
    _wait_img('img_ecd.png', conf=0.9, timeout=-1)

    p.write('# ecd_robo')
    p.press('enter')
    _click_img('img_receita.png', conf=0.9, clicks=2)
    time.sleep(3)
    p.press('tab')

    pyperclip.copy(nome_arquivo2)
    p.hotkey('ctrl', 'v')

    p.press('enter', interval=2)
    p.press('tab', presses=10, interval=0.2)

    while not _find_img('txt.png', conf=0.9):
        p.press('down', interval=0.3)
        p.press('up', interval=0.3)
        p.press('enter', interval=0.3)
        time.sleep(1)

    p.press('down', interval=0.3)
    p.press('up', interval=0.3)
    p.press('enter')
    _wait_img('sucesso_ecd.png', conf=0.9, timeout=-1)
    p.press('enter')
    _click_img('botao_recuperar.png', conf=0.9, clicks=1)
    _wait_img('recuperado.png', conf=0.9, timeout=-1)
    p.press('enter')
    _click_img('escrituracao.png', conf=0.9)
    _click_img('dados_escrituracao.png', conf=0.9)
    _click_img('sped_escrituracao.png', conf=0.9)
    _click_img('4_sig.png', conf=0.9)
    _click_img('procurador.png', conf=0.9)
    _click_img('remover.png', conf=0.9)
    _wait_img('excluir.png', conf=0.9, timeout=-1)
    p.hotkey('alt', 's')
    _click_img('rpem.png', conf=0.9, clicks=2)
    _wait_img('registro.png', conf=0.9, timeout=-1)
    p.press('tab', presses=12, interval=0.2)
    p.press('del', interval=0.5)
    p.write('S')
    p.press('tab', interval=0.5)
    _click_img('salvar.png', conf=0.9)
    _click_img('passo_a_passo2.png', conf=0.9)
    _click_img('botao_validar.png', conf=0.9)
    _wait_img('leitor_pdf.png', conf=0.9, timeout=-1)
    data = data + 1

    if _find_img('validar_erro.png', conf=0.9):
        salvar_copia(codigo, data, 'Erro')
    elif _find_img('finalizado_ok.png', conf=0.9):
        salvar_copia(codigo, data, 'Sucesso')

def salvar_copia(codigo, data, status):
    p.press('enter')
    _click_img('fechar.png', conf=0.9)
    _click_img('escrituracao.png', conf=0.9)
    _click_img('gerar_copia.png', conf=0.9)
    _wait_img('titulo_seguranca.png', conf=0.9, timeout=-1)
    p.press('tab', presses=4, interval=0.5)
    p.press('down', interval=0.5)
    p.press('up', presses=15)
    while not _find_img('comum.png', conf=0.9):
        time.sleep(0.2)
        p.press('down')
    _click_img('comum2.png', conf=0.9)
    _click_img('img_ecd.png', conf=0.9, clicks=2)
    _click_img('pasta_copia.png', conf=0.9, clicks=2)
    _wait_img('ecd_robo3.png', conf=0.9, timeout=-1)
    p.press('tab')
    pyperclip.copy('ECD ' + codigo + ' - Competência ' + str(data) + ' ' + status)
    p.hotkey('ctrl', 'v')
    _click_img('gravar.png', conf=0.9)

##
def gera_arquivo_receitabx(cnpj, data_inicio, data_fim, razao):#
    while not _find_img('receitabx.png', conf=0.9):
        _click_img('icone_bx.png', conf=0.9)
    p.press('up', interval=0.5)
    p.press('tab', presses=2, interval=0.2)
    p.press('down', interval=0.2)
    p.press('tab', interval=0.2)
    p.press('down', interval=0.2)
    p.press('tab', interval=0.2)
    p.write(cnpj)
    p.press('tab', presses=4, interval=0.2)
    p.press('enter', interval=0.5)
    _click_img('pesquisa_pequeno.png', conf=0.9)
    _wait_img('user.png', conf=0.9, timeout=-1)
    p.press('tab', presses=7, interval=0.2)
    p.press('down', presses=2, interval=0.2)
    p.press('tab', interval=0.2)
    p.press('down', interval=1)
    time.sleep(1)
    p.press('space')
    data_inicio_ano = data_inicio[4:]
    data_inicio_dia_mes = data_inicio[:4]
    data_inicio_ano = int(data_inicio_ano) - 1
    data_inicio_full = str(data_inicio_dia_mes) + str(data_inicio_ano)


    pyperclip.copy(data_inicio_full)
    p.hotkey('ctrl', 'v')
    p.press('tab', interval=0.2)
    time.sleep(1)
    p.press('space')
    data_fim_ano = data_fim[4:]
    data_fim_dia_mes = data_fim[:4]
    data_fim_ano = int(data_fim_ano) - 1
    data_fim_full = str(data_fim_dia_mes) + str(data_fim_ano)
    pyperclip.copy(data_fim_full)
    p.hotkey('ctrl', 'v')
    time.sleep(1)###
    p.press('tab')
    _click_img('pesquisar.png', conf=0.9)
    time.sleep(2)
    p.press('tab')
    _click_img('autenticada.png', conf=0.9)
    _click_img('quadrado.png', conf=0.9)
    _click_img('solicitar.png', conf=0.9)
    _wait_img('registrado_sucesso.png', conf=0.9, timeout=-1)
    p.press('enter')
    _click_img('acompanhamento.png', conf=0.9)
    p.press('tab', presses=8, interval=0.2)
    time.sleep(1)
    p.press('tab')
    _click_img('baixar.png', conf=0.9)
    _click_img('quadrado.png', conf=0.9)
    _click_img('botao_baixar.png', conf=0.9)
    time.sleep(5)
    _click_img('sair.png', conf=0.9)
    nome_arquivo = move_arquivo_receitabx(razao)
    _click_img('sped_icone.png', conf=0.9)
    while not _find_img('localizar.png', conf=0.9):
        time.sleep(1)
    return nome_arquivo
def move_arquivo_receitabx(razao):
    pasta = 'C:\\Users\\robo\Documents\Arquivos ReceitanetBX\\'
    destino = 'T:\# ecd_robo\Arquivo Receita Bx\\'
    for arquivo in os.listdir(pasta):
        if str(pasta) + str(arquivo) == 'C:\\Users\\robo\Documents\Arquivos ReceitanetBX\\tmp':
            break


        arquivo_renomeado = arquivo.split('-')
        nome_arquivo_formatado = arquivo_renomeado[0] + razao + '.txt'

        move(pasta + arquivo, destino + nome_arquivo_formatado)

    return str(nome_arquivo_formatado)
@barra_de_status
def run(window, values):
    if not (values['-codigo-'] and values['-data_inicio-'] and values['-data_final-'] and values['-livro-']):
        window['-Mensagens-'].update('Preencha todos os campos obrigatórios!')
    elif values['-retificar-'] and not values['-hash-']:
        window['-Mensagens-'].update('Hash não informado!')
    else:
        if values['-normal-']:
            ecd_normal(window, values, values['-codigo-'])
        elif values['-retificar-']:
            ecd_retificar(window, values)
    print('Robo Em Finalizado!')

if __name__ == '__main__':
    run()