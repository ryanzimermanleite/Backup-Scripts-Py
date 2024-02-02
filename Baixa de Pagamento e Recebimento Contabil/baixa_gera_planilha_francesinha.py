# -*- coding: utf-8 -*-
import shutil, time, os, pyautogui as p
from sys import path
path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv
import PySimpleGUI as sg
from pyautogui import alert
import pandas as pd
from xlrd import open_workbook

data_anterior = []

#
def run(codigo):
    if not login(codigo, 'Relatorio Erros'):
        return

    if tipo_pagar is True:
         # RELATORIO A PAGAR
         titulo_tela = 'titulo_contas_a_pagar.png'
         tecla = 'p'
         imagem = 'aberta_paga_parcial.png'
         titulo_relatorio = 'contas_a_pagar_fornecedor.png'
         tipo = 'A Pagar'
         caminho = 'Contas a Pagar'
         nome_relatorio = 'Relatorio Contas a Pagar'
         # RELATORIO FORNECEDORES
         titulo_tela2 = 'titulo_fornecedores.png'
         tecla2 = 'f'
         titulo_relatorio2 = 'relacao_fornecedores.png'
         tipo2 = 'Fornecedores'
         caminho2 = 'Relação Fornecedores'
         nome_relatorio2 = 'Relatorio Fornecedores'
         # EXTRAI DADOS EXCEL
         tipo3 = 'pagar_fornecedor'

    if tipo_receber is True:
        # RELATORIO A RECEBER
        titulo_tela = 'titulo_contas_a_receber.png'
        tecla = 'r'
        imagem = 'aberta_recebida_parcial.png'
        titulo_relatorio = 'contas_a_receber_por_cliente.png'
        tipo = 'A Receber'
        caminho = 'Contas a Receber'
        nome_relatorio = 'Relatorio Contas a Receber'
        # RELATORIO CLIENTES
        titulo_tela2 = 'titulo_clientes.png'
        tecla2 = 'c'
        titulo_relatorio2 = 'relacao_clientes.png'
        tipo2 = 'Clientes'
        caminho2 = 'Relação Clientes'
        nome_relatorio2 = 'Relatorio Clientes'
        # EXTRAI DADOS EXCEL
        tipo3 = 'receber_cliente'

    arquivo_pagar_receber = gera_relatorio_pagar_receber(codigo, titulo_tela, tecla, imagem, titulo_relatorio, tipo, caminho, nome_relatorio)
    arquivo_fornecedor_cliente = gera_relatorio_fornecedor_cliente(codigo, titulo_tela2, tecla2, titulo_relatorio2, tipo2, caminho2, nome_relatorio2)
    dados_pagar_receber = open_lista_dados(arquivo_pagar_receber)
    dados_fornecedor_cliente = open_lista_dados(arquivo_fornecedor_cliente)
    extrai_dados_excel(codigo, dados_pagar_receber, dados_fornecedor_cliente, tipo3)


def gera_relatorio_pagar_receber(codigo, titulo_tela, tecla, imagem, titulo_relatorio, tipo, caminho, nome_relatorio):
    tela_contas_pagar_receber(titulo_tela, tecla, imagem)
    time.sleep(1)
    p.hotkey('alt', 'o')
    status = verifica_possui_relatorio(titulo_relatorio)
    if status == 'ok':
        salvar_pdf(codigo, tipo)
        arquivo = mover_arquivo(codigo, caminho, tipo, nome_relatorio)
    return arquivo

def gera_relatorio_fornecedor_cliente(codigo, titulo_tela2, tecla2, titulo_relatorio2, tipo2, caminho2, nome_relatorio2):
    tela_clientes_fornecedor(titulo_tela2, tecla2)
    time.sleep(1)
    p.hotkey('alt', 'o')
    status = verifica_possui_relatorio(titulo_relatorio2)
    if status == 'ok':
        salvar_pdf(codigo, tipo2)
        arquivo = mover_arquivo(codigo, caminho2, tipo2, nome_relatorio2)
    return arquivo

def extrai_dados_excel(codigo, dados_pagar_receber, dados_fornecedor_cliente, tipo3):
    dados_cliente, status = open_lista_dados_cliente(input_excel_cliente)
    global data_anterior
    if status == 'ok':
        nota_cliente = []

        for linha_cliente in dados_cliente.itertuples():
            numero_nota = linha_cliente[1]
            data_pagamento = linha_cliente[2]
            valor_pagamento = linha_cliente[3]
            juros = linha_cliente[4]
            multa = linha_cliente[5]
            desconto = linha_cliente[6]
            nota_cliente.append([str(numero_nota), str(data_pagamento), str(valor_pagamento), str(juros), str(multa), str(desconto)])

            if tipo3 == 'pagar_fornecedor':
                nome_excel_encontrados = 'A Pagar Fornecedor'


            elif tipo3 == 'receber_cliente':
                nome_excel_encontrados = 'A Receber Cliente'

            gera_importação_dominio(dados_pagar_receber, dados_fornecedor_cliente, numero_nota, nome_excel_encontrados, data_pagamento, valor_pagamento, juros, multa, desconto, codigo)

        escreve_header(codigo, nome_excel_encontrados)
        gera_planilha_erros(codigo, nome_excel_encontrados)


def gera_importação_dominio(dados_pagar_receber, dados_fornecedor_cliente, numero_nota, nome_excel, data_pagamento, valor_pagamento, juros, multa, desconto, codigo):
    notas_pagar_receber = []

    global data_anterior

    for count1, linha1 in enumerate(range(dados_fornecedor_cliente.nrows), start=1):
        codigo1 = dados_fornecedor_cliente.cell_value(linha1, 0)
        cnpj = dados_fornecedor_cliente.cell_value(linha1, 10)
        codigo1 = (str(codigo1).split('.'))
        for count, linha2 in enumerate(range(dados_pagar_receber.nrows), start=1):
            data_atual = []
            documento = dados_pagar_receber.cell_value(linha2, 16)
            data_vencimento = dados_pagar_receber.cell_value(linha2, 19)
            codigo2 = dados_pagar_receber.cell_value(linha2, 31).split(' ')
            notas_pagar_receber.append(str(documento))

            if str(numero_nota) == str(documento):
                if codigo2[0] == codigo1[0]:

                    data_pagamento = str(data_pagamento)
                    data_atual.append([numero_nota, data_vencimento])
                    valor_pagamento = str(valor_pagamento).replace(".", ",")
                    juros = str(juros).replace(".", ",")
                    multa = str(multa).replace(".", ",")
                    desconto = str(desconto).replace(".", ",")

                    for elemento_x in data_atual:
                        if elemento_x not in data_anterior:
                            _escreve_relatorio_csv(f'{numero_nota};{cnpj};{data_vencimento};{data_pagamento};{valor_pagamento};{juros};{multa};{desconto};{numero_banco[1]}', nome=codigo + ' - ' + nome_excel)
                            data_anterior.append(elemento_x)
                            return
                        else:
                            continue

def gera_planilha_erros(codigo, nome_excel_encontrados):
    dados_resultado = []
    nota_data_valor = []

    dados_cliente, status = open_lista_dados_cliente(input_excel_cliente)

    for linha_cliente in dados_cliente.itertuples():
        numero_nota = linha_cliente[1]
        data_pagamento = linha_cliente[2]
        valor_pagamento = linha_cliente[3]
        juros = linha_cliente[4]
        multa = linha_cliente[5]
        desconto = linha_cliente[6]
        nota_data_valor.append([str(numero_nota), str(data_pagamento), str(valor_pagamento).replace(".", ","), str(juros).replace(".", ","), str(multa).replace(".", ","), str(desconto).replace(".", ",")])

    empresas = abre_arquivo_resultado(codigo, nome_excel_encontrados)
    if empresas is False:
        for linha_cliente in dados_cliente.itertuples():
            numero_nota = linha_cliente[1]
            data_pagamento = linha_cliente[2]
            valor_pagamento = linha_cliente[3]
            juros = linha_cliente[4]
            multa = linha_cliente[5]
            desconto = linha_cliente[6]
            cnpj = 'Não Encontrado'
            data_venc = 'Não Encontrado'
            _escreve_relatorio_csv(
                f'{str(numero_nota)};{cnpj};{data_venc};{str(data_pagamento)};{str(valor_pagamento).replace(".", ",")};{str(juros).replace(".", ",")};{str(multa).replace(".", ",")};{str(desconto).replace(".", ",")};{numero_banco[1]}',
                nome=codigo + ' - ' + nome_excel_encontrados + ' Não Encontrados')
        try:
            _escreve_header_csv('Número da Nota;'
                                'CPF/CNPJ do fornecedor;'
                                'DAta de Vencimento da parcela;'
                                'Data da baixa da parcela;'
                                'Valor pago;'
                                'Valor juros;'
                                'Valor multa;'
                                'Valor Desconto;'
                                'Código da conta banco;', nome=codigo + ' - ' + nome_excel_encontrados + ' Não Encontrados')
        except:
            alert(text=f'Planilha de notas não encontradas, sem dados.\n\nNão foi gerada!.')
    else:
        for count, empresa in enumerate(empresas, start=1):
            nota_resultado = empresa[0]
            data_baixa = empresa[3]
            valor_resultado = empresa[4]
            juros_resultado = empresa[5]
            multa_resultado = empresa[6]
            desconto_resultado = empresa[7]
            dados_resultado.append([str(nota_resultado), str(data_baixa), str(valor_resultado), str(juros_resultado), str(multa_resultado), str(desconto_resultado)])

        for elemento_y in nota_data_valor:
            if elemento_y not in dados_resultado:
                cnpj = 'Não Encontrado'
                data_venc = 'Não Encontrado'
                _escreve_relatorio_csv(f'{elemento_y[0]};{cnpj};{data_venc};{elemento_y[1]};{elemento_y[2]};{elemento_y[3]};{elemento_y[4]};{elemento_y[5]};{numero_banco[1]}', nome=codigo + ' - ' + nome_excel_encontrados + ' Não Encontrados')

        try:
            _escreve_header_csv('Número da Nota;'
                                'CPF/CNPJ do fornecedor;'
                                'DAta de Vencimento da parcela;'
                                'Data da baixa da parcela;'
                                'Valor pago;'
                                'Valor juros;'
                                'Valor multa;'
                                'Valor Desconto;'
                                'Código da conta banco;', nome=codigo + ' - ' + nome_excel_encontrados + ' Não Encontrados')
        except:
            alert(text=f'Planilha de notas não encontradas, sem dados.\n\nNão foi gerada!.')
def abre_arquivo_resultado(codigo, nome_excel_encontrados):
    file = 'V:\Setor Robô\Scripts Python\Domínio\Baixa de Pagamento e Recebimento Contabil\execução\\' + codigo + ' - ' + nome_excel_encontrados + '.csv'
    if not file:
        return False

    try:
        with open(file, 'r', encoding='latin-1') as f:
            dados = f.readlines()
    except Exception as e:
        return False

    return list(map(lambda x: tuple(x.replace('\n', '').split(';')), dados))



def open_lista_dados(input_excel):
    workbook = ''
    file = input_excel

    if not file:
        return False

    if file.endswith('.xls') or file.endswith('.XLS'):
        workbook = open_workbook(file)
        workbook = workbook.sheet_by_index(0)

    return workbook

def open_lista_dados_cliente(input_excel_cliente):
    try:
        arquivo_cliente = pd.read_excel(input_excel_cliente)
        return arquivo_cliente, 'ok'

    except:
        alert(text=f'Você selecionou uma planilha fora do padrão.')
        window['-INICIAR-'].update(disabled=False)
        return arquivo_cliente, 'erro'

def tela_clientes_fornecedor(titulo, tecla):
    while not _find_img(titulo, conf=0.9):
        p.hotkey('alt', 'a')
        time.sleep(0.5)
        p.press(tecla)
        time.sleep(5)
    time.sleep(1)
    p.press('l')
    while not _find_img('codigo_nome.png', conf=0.9):
        time.sleep(1)
    p.hotkey('alt', 'r')
    while not _find_img('listagem_relatorio.png', conf=0.9):
        time.sleep(1)
    _click_img('adicionar.png', conf=0.9)

def tela_contas_pagar_receber(titulo, tecla, aberta_parcial):
    while not _find_img(titulo, conf=0.9):
        p.hotkey('alt', 'r')
        time.sleep(0.5)
        p.press('g')
        time.sleep(0.5)
        p.press(tecla)
        time.sleep(2)
    _click_img(aberta_parcial, conf=0.9)

def verifica_possui_relatorio(titulo_relatorio):
    while not _find_img(titulo_relatorio, conf=0.9):
        time.sleep(1)
        if _find_img('sem_dados.png', conf=0.9):
            print('❌ Sem dados para emitir')
            p.press('enter')
            time.sleep(0.5)
            p.press('esc')
            time.sleep(0.5)
            return 'sem_dados'
    return 'ok'

def salvar_pdf(cod, tipo):
    p.click(833, 384)
    time.sleep(0.5)

    _click_img('salvar.png', conf=0.9)
    timer = 0
    _wait_img('salvar_relat.png', conf=0.9)
    time.sleep(0.5)
    _click_img('botao.png', conf=0.9)
    time.sleep(0.5)
    _click_img('planilha_cabecalho.png', conf=0.9)
    time.sleep(0.5)
    _click_img('3pontos.png', conf=0.9)
    time.sleep(0.5)

    while not _find_img('selecione_arquivo.png', conf=0.9):
        time.sleep(1)
        timer += 1
        if timer > 30:
            return False

    time.sleep(1)
    p.write(str(cod) + ' - ' + tipo + '.xls')
    time.sleep(0.5)

    if not _find_img('cliente_c_selecionado.png', pasta='imgs_c', conf=0.9):
        while not _find_img('cliente_c.png', pasta='imgs_c', conf=0.9) or _find_img('cliente_m.png', pasta='imgs_c',
                                                                                    conf=0.9):
            _click_img('botao.png', pasta='imgs_c', conf=0.9)
            time.sleep(3)

        _click_img('cliente_m.png', pasta='imgs_c', conf=0.9, timeout=1)
        _click_img('cliente_c.png', pasta='imgs_c', conf=0.9, timeout=1)
        time.sleep(5)

    time.sleep(1)
    p.hotkey('alt', 's')
    _wait_img('salvar_fechar.png', conf=0.9)
    time.sleep(1)
    p.hotkey('alt', 's')
    while not _find_img('img.png', conf=0.9):
        time.sleep(5)
        p.press('esc')
        if _find_img('gravar_dados.png', conf=0.9):
            p.hotkey('alt', 'n')
            time.sleep(1)
            p.press('esc')
    time.sleep(1)
    p.press('esc', presses=5, interval=0.2)
    return True

def login(cod, nome_relatorio):
    # espera a tela inicial do domínio
    while not _find_img('inicial.png', pasta='imgs_c', conf=0.9):
        time.sleep(1)

    p.click(833, 384)

    # espera abrir a window de seleção de empresa
    while not _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
        p.press('f8')

    time.sleep(1)
    # clica para pesquisar empresa por código
    if _find_img('codigo.png', pasta='imgs_c', conf=0.9):
        _click_img('codigo.png', pasta='imgs_c', conf=0.9)
    p.write(cod)
    time.sleep(3)

    # confirmar empresa
    p.hotkey('alt', 'a')
    # enquanto a window estiver aberta verifica exceções
    while _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
        time.sleep(1)

    p.press('esc', presses=5)
    time.sleep(1)
    return True

def mover_arquivo(cod, caminho, tipo, nome_relatorio):
    pasta_origem = 'C:\\'
    pasta_destino = 'V:\Setor Robô\Scripts Python\Domínio\Baixa de Pagamento e Recebimento Contabil\execução\\' + caminho + '\\'
    nome_arquivo = str(cod) + ' - ' + tipo + '.xls'
    p.press('esc', presses=5)
    time.sleep(1)
    try:
        shutil.move(os.path.join(pasta_origem, nome_arquivo), os.path.join(pasta_destino, nome_arquivo))
    except:
        pass
    arquivo = pasta_destino + nome_arquivo
    return arquivo


def escreve_header(codigo, nome_excel):
    try:
        _escreve_header_csv('Número da Nota;'
                            'CPF/CNPJ do fornecedor;'
                            'DAta de Vencimento da parcela;'
                            'Data da baixa da parcela;'
                            'Valor pago;'
                            'Valor juros;'
                            'Valor multa;'
                            'Valor Desconto;'
                            'Código da conta banco;', nome=codigo + ' - ' + nome_excel)
    except:
        alert(text=f'Planilha de notas não encontradas, sem dados.\n\nNão foi gerada!.')

sg.theme('GrayGrayGray')
bancos2 = ['Banco Neutro - 0', 'Banco Caixa - 5', 'Banco Itaú - 21', 'Santander - 1090']
modelo = ['Veiga & Postal', 'Druck Chemie', 'LockPipe']
layout = [

    [sg.Text("Código: "), sg.InputText(key="-codigo-", size=(5, 5))],
    [sg.FileBrowse('Arquivo', button_color='light yellow', key='-Abrir-', file_types=(('Planilhas Excel', '*.xlsx *.xls'),)),
     sg.InputText(key='-input_excel_cliente-', size=200, disabled=True)],
    [sg.Text('Tipo:   '), sg.Radio("Pagar", "tipo", enable_events=True, key="-pagar-", default=True),
     sg.Radio("Receber", "tipo", enable_events=True, key="-receber-")],
    [sg.Text('Bancos:'),
     sg.Combo(bancos2, font=("Helvetica", 10), expand_x=True, enable_events=True, readonly=False, key='-COMBO-')],
    [sg.Text('Modelos:'),
     sg.Combo(modelo, font=("Helvetica", 10), expand_x=True, enable_events=True, readonly=False, key='-COMBO-')],
    [sg.Text('Status: '), sg.InputText("", key="texto", disabled=True)],
    [sg.Button("Gerar", button_color='lightgreen', key='-gerar-', expand_x=True),
     sg.Button("Resultado", button_color='lightblue', expand_x=True),
     sg.Button("Sair", button_color='Light Coral', expand_x=True)]
]

window = sg.Window("DCA", layout, size=(230, 210))

while True:
    evento, values = window.read()

    try:
        input_excel_cliente = values['-input_excel_cliente-']
        codigo = values["-codigo-"]
        tipo_pagar = values['-pagar-']
        tipo_receber = values['-receber-']
        banco = values['-COMBO-']
        numero_banco = banco.split('- ')
    except:
        input_excel_cliente = 'Desktop'

    if evento == sg.WIN_CLOSED or evento == "Sair":
        break
    if evento == "-gerar-":
        if not input_excel_cliente or codigo == '':
            alert(text=f'Por favor selecione uma planilha do Cliente.' + '\n\n'
                                                                         f'Digite o código do domínio referente a Empresa.')
        else:
            window['-gerar-'].update(disabled=True)
            run(codigo)

        window["texto"].update(f"{codigo} - OK!")

window.close()