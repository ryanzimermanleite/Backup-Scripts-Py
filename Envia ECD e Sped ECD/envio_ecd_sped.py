from shutil import move
import pyperclip, time, os, re, pyautogui as p
import PySimpleGUI as sg
from functools import wraps
from threading import Thread
from sys import path
path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img

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

def combinar_tecla_e_esperar(tecla1, tecla2, delay):
    p.hotkey(tecla1, tecla2)
    time.sleep(delay)
def apertar_tecla_e_esperar(tecla, delay=0, quantidade=1, intervalo=0):
    p.press(tecla, presses=quantidade, interval=intervalo)
    time.sleep(delay)

def escrever_e_esperar(texto, delay):
    p.write(texto)
    time.sleep(delay)

def envio_ecd_dominio(codigo, data_inicio, data_final, tipo, hash, livro):
    login(codigo)
    abre_janela_sped_contabil()
    escreve_data_inicio_e_final(data_inicio, data_final)
    escreve_caminho_arquivo_txt(tipo)
    abre_janela_outros_dados()
    define_finalidade_escrituracao_e_hash(tipo, hash)
    escreve_data_arquivamento_e_encerramento(data_final)
    verifica_checkbox_aba_geral()
    verifica_livro_aba_dados(livro)
    verifica_certificado_aba_dados(tipo)
    verifica_demonstrativo_gerar(tipo)
    if tipo == 'RETIFICAR':
        verifica_demonstrativo_arquivos_rtf(codigo)
    verifica_aba_opcoes()
    verifica_mensagens(tipo)

def envio_sped_contabil(tipo, codigo):
    abre_programa_sped_contabil()
    abre_janela_sped_importar()
    acessa_diretorio_comum()
    acessa_diretorio_arquivos_txt(tipo)
    arquivo = abre_arquivo_txt_importacao(codigo)
    verifica_importacao()
    abre_janela_recuperacao()
    data_inicio, data_fim, razao, cnpj = extrai_informacoes_txt(arquivo, tipo)
    arquivo_bx = gera_arquivo_receitabx(cnpj, data_inicio, data_fim, razao)
    abre_janela_localizar()
    abre_pasta_arquivo_receitabx()
    abre_e_valida_arquivo_receita_bx(arquivo_bx)
    remove_altera_procurador()
    status = valida_arquivo_recuperado()
    salvar_copia(codigo, data_inicio, status, tipo, arquivo_bx)

def gera_arquivo_receitabx(cnpj, data_inicio, data_fim, razao):
    abre_programa_receitabx()
    seleciona_certificado_perfil(cnpj)
    abre_janela_pesquisar()
    configura_arquivo()
    data_inicio_formatada, data_fim_formatada = formata_data(data_inicio, data_fim)
    escreve_data_e_pesquisa(data_inicio_formatada, data_fim_formatada)
    seleciona_arquivo_pesquisado()
    baixa_arquivo_acompanhamento()
    nome_arquivo_receitabx = move_arquivo_receitabx(razao)
    retorna_ao_sped()
    return nome_arquivo_receitabx


def abre_programa_receitabx():
    while not _find_img('receitabx.png', conf=0.9):
        _click_img('icone_bx.png', conf=0.9)

def seleciona_certificado_perfil(cnpj):
    p.press('up', interval=0.5)
    p.press('tab', presses=2, interval=0.2)
    p.press('down', interval=0.2)
    p.press('tab', interval=0.2)
    p.press('down', interval=0.2)
    p.press('tab', interval=0.2)
    p.write(cnpj)
    p.press('tab', presses=4, interval=0.2)
    p.press('enter', interval=0.5)

def abre_janela_pesquisar():
    _click_img('pesquisa_pequeno.png', conf=0.9)
    _wait_img('user.png', conf=0.9, timeout=-1)

def configura_arquivo():
    p.press('tab', presses=7, interval=0.2)
    p.press('down', presses=2, interval=0.2)
    p.press('tab', interval=0.2)
    p.press('down', interval=1)

def formata_data(data_inicio, data_fim):
    data_inicio_ano = data_inicio[4:]
    data_inicio_dia_mes = data_inicio[:4]
    data_inicio_ano = int(data_inicio_ano) - 1
    data_inicio_formatada = str(data_inicio_dia_mes) + str(data_inicio_ano)
    pyperclip.copy(str(data_inicio_dia_mes) + str(data_inicio_ano))
    data_fim_ano = data_fim[4:]
    data_fim_dia_mes = data_fim[:4]
    data_fim_ano = int(data_fim_ano) - 1
    data_fim_formatada = str(data_fim_dia_mes) + str(data_fim_ano)
    return data_inicio_formatada, data_fim_formatada

def escreve_data_e_pesquisa(data_inicio_formatada, data_fim_formatada):
    p.press('space')
    pyperclip.copy(data_inicio_formatada)
    p.hotkey('ctrl', 'v')
    p.press('tab', interval=0.5)
    p.press('space')
    pyperclip.copy(data_fim_formatada)
    p.hotkey('ctrl', 'v')
    p.press('tab', interval=0.5)
    _click_img('pesquisar.png', conf=0.9)

def seleciona_arquivo_pesquisado():
    p.press('tab')
    _click_img('autenticada.png', conf=0.9)
    _click_img('quadrado.png', conf=0.9)
    _click_img('solicitar.png', conf=0.9)
    _wait_img('registrado_sucesso.png', conf=0.9, timeout=-1)
    p.press('enter')

def baixa_arquivo_acompanhamento():
    _click_img('acompanhamento.png', conf=0.9)
    p.press('tab', presses=8, interval=0.2)
    time.sleep(1)
    p.press('tab')
    _click_img('baixar.png', conf=0.9)  #
    _click_img('quadrado.png', conf=0.9)
    _click_img('botao_baixar.png', conf=0.9)
    _wait_img('fim_baixar.png', conf=0.9, timeout=-1)
    _click_img('sair.png', conf=0.9)

def move_arquivo_receitabx(razao):
    pasta = 'C:\\Users\\robo\Documents\Arquivos ReceitanetBX\\'
    destino = 'T:\# ecd_robo\Arquivo Receita Bx\\'
    for arquivo in os.listdir(pasta):
        if str(pasta) + str(arquivo) == 'C:\\Users\\robo\Documents\Arquivos ReceitanetBX\\tmp':
            break
        arquivo_renomeado = arquivo.split('-')
        nome_arquivo_formatado = arquivo_renomeado[0] + ' - ' + razao + '.txt'
        move(pasta + arquivo, destino + nome_arquivo_formatado)
    return str(nome_arquivo_formatado)

def retorna_ao_sped():
    _click_img('sped_icone.png', conf=0.9)
    while not _find_img('localizar.png', conf=0.9):
        time.sleep(3)

def abre_janela_localizar():
    _click_img('localizar.png', conf=0.9)
    _wait_img('abrir_recuperar_ecd.png', conf=0.9, timeout=-1)
    _click_img('clique.png', conf=0.9)

def abre_pasta_arquivo_receitabx():
    p.press('up', presses=15)

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

def abre_e_valida_arquivo_receita_bx(arquivo_bx):
    pyperclip.copy(arquivo_bx)
    p.hotkey('ctrl', 'v')
    time.sleep(1)
    p.press('enter')
    _wait_img('sucesso_ecd.png', conf=0.9, timeout=-1)
    p.press('enter')

def remove_altera_procurador():
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

def valida_arquivo_recuperado():
    _click_img('passo_a_passo2.png', conf=0.9)
    _click_img('botao_validar.png', conf=0.9)
    while _find_img('validando.png', conf=0.9):
        time.sleep(0.5)
    time.sleep(3)
    _wait_img('ok2.png', conf=0.9, timeout=-1)
    if _find_img('validar_erro.png', conf=0.9):
        p.hotkey('alt', 'o')
        return 'Erro'

    elif _find_img('finalizado_ok.png', conf=0.9):
        p.hotkey('alt', 'o')
        return 'Sucesso'

def abre_programa_sped_contabil():
    while not _find_img('sped_icone_importar.png', conf=0.9):
        _click_img('sped_icone.png', conf=0.9)
        time.sleep(4)

def abre_janela_sped_importar():
    while not _find_img('sped_importar_escrituracao.png', conf=0.9):
        _click_img('sped_icone_importar.png', conf=0.9)

def acessa_diretorio_comum():
    p.press('tab', presses=4, interval=0.2)
    p.press('down', interval=0.5)
    p.press('up', presses=15)

    while not _find_img('comum.png', conf=0.9):
        apertar_tecla_e_esperar('down',0.2)
    _click_img('comum2.png', conf=0.9)

def acessa_diretorio_arquivos_txt(tipo):
    _click_img('img_ecd.png', conf=0.9, clicks=2)
    _wait_img('ecd_robo2.png', conf=0.9, timeout=-1)

    if tipo == 'NORMAL':
        _click_img('arquivo_txt_normal.png', conf=0.9, clicks=2)
    elif tipo == 'RETIFICAR':
        _click_img('arquivo_txt_retificar.png', conf=0.9, clicks=2)

def abre_arquivo_txt_importacao(codigo):
    arquivo = 'sped_diario' + str(codigo).rjust(5, '0') + '.txt'
    p.press('tab')
    p.write(arquivo)
    p.press('enter')
    return arquivo

def verifica_importacao():
    _wait_img('importacao_blocos.png', conf=0.9, timeout=-1)
    _click_img('ok.png', conf=0.9)
    _wait_img('sped_aviso_sucesso.png', conf=0.9, timeout=-1)
    p.hotkey('alt', 'n')

def abre_janela_recuperacao():
    _wait_img('passo_a_passo.png', conf=0.9, timeout=-1)
    _click_img('escrituracao.png', conf=0.9)
    _click_img('recuperar_ecd.png', conf=0.9)
    _wait_img('titulo_recuperar_ecd.png', conf=0.9, timeout=-1)

def extrai_informacoes_txt(arquivo, tipo):
    pasta = f'T:\# ecd_robo\Arquivos TXT {tipo}\\' + arquivo

    with open(pasta, 'r', errors="ignore") as f:
        dados = f.read()
    dados_arquivo = re.findall(r'\|([^|]*)', dados)

    data_inicio = dados_arquivo[2]
    data_fim = dados_arquivo[3]
    razao = dados_arquivo[4]
    cnpj = dados_arquivo[5]
    return data_inicio, data_fim, razao, cnpj

def abre_janela_sped_contabil():
    while not _find_img('sped_contabil.png', conf=0.9):
        combinar_tecla_e_esperar('alt', 'r', 0.5)
        apertar_tecla_e_esperar('f', 0.5)
        apertar_tecla_e_esperar('s', 3)

def escreve_data_inicio_e_final(data_inicio, data_final):
    escrever_e_esperar(data_inicio, 0.5)
    apertar_tecla_e_esperar('tab', 0.5)
    escrever_e_esperar(data_final, 0.5)

def escreve_caminho_arquivo_txt(tipo):
    apertar_tecla_e_esperar('tab', 0.5)
    escrever_e_esperar(f'T:\# ecd_robo\Arquivos TXT {tipo}', 0.5)
    p.press('backspace', presses=50)
    escrever_e_esperar(f'T:\# ecd_robo\Arquivos TXT {tipo}', 0.5)

def abre_janela_outros_dados():
    combinar_tecla_e_esperar('alt', 'd', 0.5)
    _wait_img('outros_dados.png', conf=0.9, timeout=-1)

def define_finalidade_escrituracao_e_hash(tipo, hash):
    p.press('tab', presses=3, interval=0.1)
    if tipo == 'NORMAL':
        apertar_tecla_e_esperar('up', 0.5)
        apertar_tecla_e_esperar('tab', 0.1, 2)
    if tipo == 'RETIFICAR':
        apertar_tecla_e_esperar('down', 0.5)
        apertar_tecla_e_esperar('tab', 0.5)
        apertar_tecla_e_esperar('down', 0.2, 5)
        apertar_tecla_e_esperar('tab', 0.5)
        hash_num = str(hash).replace('.', '')
        escrever_e_esperar(hash_num, 0.5)
        apertar_tecla_e_esperar('tab', 0.1, 2)

def escreve_data_arquivamento_e_encerramento(data_final):
    escrever_e_esperar(str(data_final), 0.5)
    apertar_tecla_e_esperar('tab', 0.5)
    escrever_e_esperar(str(data_final), 0.5)

def verifica_checkbox_aba_geral():
    if _find_img('gerar_movimento.png', conf=0.9):
        _click_img('aba_dados.png', conf=0.9)
    else:
        apertar_tecla_e_esperar('tab', 0.5, 2, 0.2)
        apertar_tecla_e_esperar('space', 0.5)
        apertar_tecla_e_esperar('tab', 0.5, 2, 0.2)
        apertar_tecla_e_esperar('space', 0.5)
        _click_img('aba_dados.png', conf=0.9)

def verifica_livro_aba_dados(livro):
    apertar_tecla_e_esperar('right', 0.5, 5, 0.2)
    apertar_tecla_e_esperar('backspace', 0.5, 5, 0.2)
    apertar_tecla_e_esperar('right', 0.5, 5, 0.2)
    escrever_e_esperar(str(livro), 0.5)

def verifica_certificado_aba_dados(tipo):
    if tipo == 'NORMAL':
        while not _find_img('sem_evandro.png', conf=0.9):
            combinar_tecla_e_esperar('alt', 'x', 0.5)
    if tipo == 'RETIFICAR':
        if _find_img('sem_evandro.png', conf=0.9):
            combinar_tecla_e_esperar('alt', 'n', 0.5)
            escrever_e_esperar('30782876889', 0.5)
            apertar_tecla_e_esperar('tab', 0.5)
            _wait_img('importar_cadastro.png', conf=0.9, timeout=-1)
            combinar_tecla_e_esperar('alt', 'y', 0.5)
            apertar_tecla_e_esperar('tab', 0.5)
            apertar_tecla_e_esperar('down', 0.5, 15, 0.1)

    while not _find_img('tela_1.png', conf=0.9):
        _click_img('demonstrativos.png', conf=0.9)
    apertar_tecla_e_esperar('tab', 0.5)

def verifica_demonstrativo_gerar():
    if _find_img('gerar_img.png', conf=0.9):
       return
    else:
        _click_img('gerar.png', conf=0.9)
        apertar_tecla_e_esperar('tab', 0.5)
        apertar_tecla_e_esperar('space', 0.5)
        apertar_tecla_e_esperar('tab', 0.5)
        apertar_tecla_e_esperar('space', 0.5)

def verifica_demonstrativo_arquivos_rtf(codigo):
    _click_img('botao_rtf.png', conf=0.9)
    while not _find_img('selecione_arquivo.png', conf=0.9):
        _click_img('3pontos_new.png', conf=0.9)
        time.sleep(3)
    while not _find_img('arquivos_rtf.png', conf=0.9):
        apertar_tecla_e_esperar('tab', 0.5, 4, 0.1)
        apertar_tecla_e_esperar('down', 0.5)
        _click_img('cliente_t.png', conf=0.9)
        _wait_img('ecd_robo.png', conf=0.9, timeout=-1)
        _click_img('ecd_robo.png', conf=0.9, clicks=2)
        _wait_img('pasta_rtf.png', conf=0.9, timeout=-1)
        _click_img('pasta_rtf.png', conf=0.9, clicks=2)
        time.sleep(1)
        apertar_tecla_e_esperar('tab', 0.5, 2, 0.5)
        time.sleep(2)

    p.write('ECD RETIFICAR_' + str(codigo))
    apertar_tecla_e_esperar('down', 0.5)
    apertar_tecla_e_esperar('enter', 0.5)

def verifica_aba_opcoes():
    _click_img('op.png', conf=0.9)
    if _find_img('opcoes.png', conf=0.9):
        p.hotkey('alt', 'o')
    else:
        apertar_tecla_e_esperar('tab', 0.5)
        apertar_tecla_e_esperar('space', 0.5)
        apertar_tecla_e_esperar('tab', 0.5)
        apertar_tecla_e_esperar('space', 0.5)
        combinar_tecla_e_esperar('alt', 'o', 5)

def verifica_mensagens(tipo):
    if tipo == 'NORMAL':
        if _find_img('grupo_contas_normal.png', conf=0.9):
            combinar_tecla_e_esperar('alt', 'n', 4)
        combinar_tecla_e_esperar('alt', 'o', 5)
        if _find_img('validacao.png', conf=0.9):
            combinar_tecla_e_esperar('alt', 's', 3)
        _wait_img('final_exportacao2.png', conf=0.9, timeout=-1)
        apertar_tecla_e_esperar('enter', 1)

    if tipo == 'RETIFICAR':
        while not _find_img('spe2.png', conf=0.9):
            if _find_img('resp.png', conf=0.9):
                apertar_tecla_e_esperar('enter', 2)

            if _find_img('todas_contas.png', conf=0.9):
                apertar_tecla_e_esperar('tab', 2)
                apertar_tecla_e_esperar('enter', 2)

        combinar_tecla_e_esperar('alt', 'o', 1)

        while not _find_img('final_exportacao.png', conf=0.9):
            if _find_img('grupo_contas.png', conf=0.9):
                combinar_tecla_e_esperar('alt', 's', 2)
            elif _find_img('validacao.png', conf=0.9):
                combinar_tecla_e_esperar('alt', 's', 2)

@barra_de_status
def run(window, values):
    codigo = values['-codigo-']
    data_inicio = values['-data_inicio-']
    data_final = values['-data_final-']
    hash = values['-hash-']
    livro = values['-livro-']

    if not (values['-codigo-'] and values['-data_inicio-'] and values['-data_final-'] and values['-livro-']):
        window['-Mensagens-'].update('Preencha todos os campos obrigatórios!')
    elif values['-retificar-'] and not values['-hash-']:
        window['-Mensagens-'].update('Hash não informado!')
    else:
        if values['-normal-']:
            envio_ecd_dominio(codigo, data_inicio, data_final, 'NORMAL', hash, livro)
            envio_sped_contabil('NORMAL', codigo)
        elif values['-retificar-']:
            envio_ecd_dominio(codigo, data_inicio, data_final, 'RETIFICAR', hash, livro)
            envio_sped_contabil('RETIFICAR', codigo)
    print('Robô Finalizado!')

if __name__ == '__main__':
    run()