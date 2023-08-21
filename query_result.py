import datetime
from xlsxwriter import Workbook
import tkinter as tk
from tkinter import filedialog

def abrir_arquivo():
    arquivo = filedialog.askopenfilename(title="Selecione um arquivo", filetypes=[("Arquivos de Texto", "*.txt"), ("Todos os Arquivos", "*.*")])
    if arquivo:
        print(f"Arquivo selecionado: {arquivo}")
    return arquivo


def query_to_excel():

    # modelo a trabalhar
    #                 +-----------------³----------------------------³--------------³-----------³ 
    # (cabeçalho)     ³   id_da_venda   ³      data_da_venda         ³   cliente    ³  produto  ³        
    #                 +-----------------³----------------------------³--------------³-----------³ 
    # (linha)       1_³    202300000001 ³ 2023-06-15-10.30.45.123456 ³ Cliente A    ³ Produto X ³        
    #                 +-----------------³----------------------------³--------------³-----------³ 

    # Crie uma janela principal
    root = tk.Tk()

    # Crie um botão que, quando clicado, abre a caixa de diálogo de abertura de arquivo
    botao = tk.Button(root, text="Abrir Arquivo", command=abrir_arquivo)
    botao.pack()

    # Inicie o loop principal da interface gráfica
    # root.mainloop()

    # arq_query_result = open('D:\\SC\\query_resut.txt', 'r')
    arq_query_result = open(abrir_arquivo(), 'r')

    query_result = arq_query_result.readlines()

    # despreza primeira coluna do header '#'
    header = Util.procura_header(query_result)[1:]
    dados = Util.procura_linhas(query_result)

    Util.gerar_excel(header, dados)


class Util:
    num_col = 0
    ARG_HEADER = ' ³ '
    ARG_LINHA = '_³ '

    @classmethod
    def quebra_linha(cls, linha):
        colunas = linha.split('³')
        # despreza ultima coluna que estara vazia
        return [coluna.strip() for coluna in colunas][:-1]
    
    @classmethod
    def quebra_header(cls, linha_header):
        colunas_header = []
        for i, coluna in enumerate(Util.quebra_linha(linha_header)):
            if coluna == '':
                if i == 0:
                    coluna = '#'
                else:
                    Util.num_col =+ 1 
                    coluna = 'COL_' + str(Util.num_col)
            colunas_header.append(coluna)

        return colunas_header

    @classmethod
    def procura_header(cls, query_result):

        header_colunas = []

        for linha in query_result:
            if Util.ARG_HEADER in linha and Util.ARG_LINHA not in linha:
                if len(header_colunas) == 0:
                    header_colunas = Util.quebra_header(linha)
                # se a primeira coluna igual a primeira coluna ja registrada, o header esta completo
                elif header_colunas[1] != Util.quebra_header(linha)[1]:
                    # despreza primeira coluna que estará em branco
                    header_colunas.extend(Util.quebra_header(linha)[1:])
                else:
                    break
        return header_colunas

    @classmethod
    def procura_linhas(cls, query_result):
        linhas = {}

        for linha in query_result:
            if Util.ARG_LINHA in linha:
                colunas = Util.quebra_linha(linha)
                # despreza ultimo caracter que seria '_'
                num_linha = int(colunas[0][:-1])
                if num_linha not in linhas:
                    linhas[num_linha] = colunas[1:]
                else:
                    linhas[num_linha] = linhas[num_linha] + colunas[1:]
        return linhas                

    @classmethod
    def definir_tamanho_planilha(cls, dados):
        linhas = len(dados)
        colunas = len(dados[1]) - 1
        return (linhas, colunas)
    
    @classmethod
    def gerar_csv(cls, header, dados):
        csv_result = []
        csv_result.append(';'.join(header) + '\n')
        for chave in dados:
            csv_result.append(';'.join(dados[chave]) + '\n')

        datahora = datetime.datetime.now().strftime("%d%m%Y_%H%M%S")
        nome_arquivo = 'arq_result{}.txt'.format(datahora)
        arq_result_csv = open(nome_arquivo, 'w')
        arq_result_csv.writelines(csv_result)

    @classmethod
    def gerar_excel(cls, header, dados):
        tamanho_planilha = Util.definir_tamanho_planilha(dados)

        # criar arquivo excel
        workbook = Workbook('query_result.xlsx')
        planilha = workbook.add_worksheet('QueryResult')

        # definir tamanho das colunas (col_ini, col_fim, tam)
        # planilha.set_column(0, tamanho_planilha[1], 15)
        
        # estilo_texto = workbook.add_format({'num_format': '@'})

        planilha.add_table(0, 0, tamanho_planilha[0], tamanho_planilha[1], 
                           {'data': list(dados.values()),
                            'columns': Util.header_para_xlsxwriter(header)
                            })
        
        planilha.autofit()

        # planilha.write(1, 1, header[0], estilo_texto)

        workbook.close()
    
    @classmethod
    def header_para_xlsxwriter(cls, header):
        # header_xlsx = []
        # for nome_coluna in header:
        #     print(nome_coluna)
        #     header_xlsx.append({'header': nome_coluna})
        # return header_xlsx
        return [{'header': nome_coluna} for nome_coluna in header]
        


if __name__ == '__main__':
    query_to_excel()