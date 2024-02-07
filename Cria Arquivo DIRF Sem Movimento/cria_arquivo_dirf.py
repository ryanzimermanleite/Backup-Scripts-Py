import pandas as pd
from sys import path
path.append(r'..\..\_comum')

def run():
    arquivo = pd.read_excel('V:\Setor Robô\Scripts Python\DIRF\Cria Arquivos DIRF\ignore\dados.xlsx')

    for linha in arquivo.itertuples():
        index = linha[0]
        codigo = str(linha[1])
        cnpj = str(linha[2])
        empresa = str(linha[3])
        cpf = linha[4]
        nome = str(linha[5])
        ddd = str(linha[6])
        telefone = str(linha[7])
        email = str(linha[8])#

        cpf_formatado = str(cpf).rjust(11, '0')
        cnpj_formatado = str(cnpj).rjust(14, '0')


        texto = f"""DIRF|2024|2023|N||B3VH8RQ|
RESPO|{cpf_formatado}|{nome}|{ddd}|{telefone}|||{email}|
DECPJ|{cnpj_formatado}|{empresa}|0|{cpf_formatado}|N|N|N|N|N|N|N|N||
IDREC|5706|
BPFDEC|{cpf_formatado}|{nome}||S|S|
RTIRF|1|||||||||||||
FIMDIRF|"""

        with open(f'execução\DIRF_{codigo}.txt', 'w') as arquivo:
            arquivo.write(texto)

if __name__ == '__main__':
    run()