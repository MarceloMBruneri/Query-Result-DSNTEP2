import datetime, os
from xlsxwriter import Workbook
import tkinter as tk
from tkinter import filedialog
from tkinter import *
from PIL import Image, ImageTk
import webbrowser


def query_to_excel():

    # Crie uma janela principal
    root = Tk()
    root.title('Query Result')

    Application(root)

    # Inicie o loop principal da interface gráfica
    root.mainloop()



class Application():
    def __init__(self, master=None):
        self.arquivo_query = StringVar()
        self.path_saida = StringVar()
        # criar widget acoplado ao todo (master)
        self.container_titulo = Frame(master)
        self.container_titulo.pack()

        # criar widget acoplado ao todo (master)
        self.container_entrada = Frame(master, pady=20, padx=50)
        self.container_entrada.pack()

        # criar widget acoplado ao todo (master)
        self.container_saida = Frame(master, pady=20, padx=50)
        self.container_saida.pack()

        # criar label acoplado ao container_titulo
        self.titulo = Label(self.container_titulo, text='Query Result')
        self.titulo.pack()

        # criar label acoplado ao container_entrada
        self.lbl_entrada = Label(self.container_entrada, text='Arquivo de entrada:', padx=10)
        self.lbl_entrada.pack(side=LEFT)

        # criar box acoplado ao container_entrada
        self.arquivo_query_box = Entry(self.container_entrada, width=80, state=DISABLED, textvariable=self.arquivo_query)
        self.arquivo_query_box.pack(side=LEFT)

        # criar botão acoplado ao container_entrada
        self.abrir_arquivo = Button(self.container_entrada, text='Selecionar Arquivo', width=15)
        # outra forma de definir atributos:
        # self.abrir_arquivo["width"] = 15
        self.abrir_arquivo['command'] = self.buscar_arquivo
        self.abrir_arquivo.pack(side=LEFT)

        # criar label acoplado ao container_saida
        self.lbl_path = Label(self.container_saida, text='Destino:', padx=10)
        self.lbl_path.pack(side=LEFT)

        # criar box acoplado ao container_entrada
        self.path_saida_box = Entry(self.container_saida, width=50, state=DISABLED, textvariable=self.path_saida)
        self.path_saida_box.pack(side=LEFT)


        # criar botão acoplado ao container_saida
        self.run = Button(self.container_saida, text='Exportar Excel', state=DISABLED, width=15)
        # outra forma de definir atributos:
        self.run['command'] = self.exportar_excel
        self.run.pack(side=RIGHT)

        # Carregue o ícone de fonte do GitHub
        # github_icon = Image.open("github-mark.png")  # Substitua pelo caminho do seu arquivo .ttf
        # github_icon = github_icon.resize((32, 32), Image.ADAPTIVE)
        # github_icon = ImageTk.PhotoImage(github_icon)

        # # Crie o botão com o ícone do GitHub
        # botao = Button(master, text="Abrir Repositório", image=github_icon, command=self.abrir_repo, compound=tk.LEFT)
        # botao.pack(pady=20)

    def buscar_arquivo(self, event=None):
        arquivo = filedialog.askopenfilename(title="Selecione um arquivo", 
                                                    filetypes=[("Arquivos de Texto", "*.txt"), 
                                                               ("Todos os Arquivos", "*.*")])
        
        if arquivo:
            self.arquivo_query.set(arquivo) 
            print(f"Arquivo selecionado: {self.arquivo_query.get()}")
            self.run['state'] = 'active'
            self.path_saida.set(os.path.dirname(arquivo))
            

    def exportar_excel(self, event=None):
        # arq_query_result = open('D:\\SC\\query_resut.txt', 'r')
        arq_query_result = open(self.arquivo_query.get(), 'r')

        query_result = arq_query_result.readlines()

        dados = Util.procura_linhas(query_result)
        if dados:
            header = Util.procura_header(query_result, len(dados[1]))
        else:
            print('Dados não encontrados')

        Util.gerar_excel(header, dados, self.path_saida.get())
        
    def abrir_repo(self, event=None):
        url = "https://github.com/MarceloMBruneri/Query-Result-DSNTEP2"
        webbrowser.open_new(url)

class Util:
    # modelo a trabalhar
    #                 +-----------------³----------------------------³--------------³-----------³ 
    # (cabeçalho)     ³   id_da_venda   ³      data_da_venda         ³   cliente    ³  produto  ³        
    #                 +-----------------³----------------------------³--------------³-----------³ 
    # (linha)       1_³    202300000001 ³ 2023-06-15-10.30.45.123456 ³ Cliente A    ³ Produto X ³        
    #                 +-----------------³----------------------------³--------------³-----------³ 

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
        # despreza primeira coluna em branco
        for coluna in Util.quebra_linha(linha_header)[1:]:
            if coluna == '':
                Util.num_col =+ 1 
                coluna = 'COL_' + str(Util.num_col)
            colunas_header.append(coluna)

        return colunas_header

    @classmethod
    def procura_header(cls, query_result, max_colunas=999):

        header_colunas = []

        for linha in query_result:
            if Util.ARG_HEADER in linha and Util.ARG_LINHA not in linha:
                if len(header_colunas) == 0:
                    header_colunas = Util.quebra_header(linha)
                # se a primeira coluna igual a primeira coluna ja registrada, o header esta completo
                elif header_colunas[1] != Util.quebra_header(linha)[0] and len(header_colunas) < max_colunas:
                    header_colunas.extend(Util.quebra_header(linha))
                else:
                    break
        # despreza primeira coluna do header '#'
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
    def gerar_csv(cls, header, dados, path_arquivo):
        csv_result = []
        csv_result.append(';'.join(header) + '\n')
        for chave in dados:
            csv_result.append(';'.join(dados[chave]) + '\n')

        datahora = datetime.datetime.now().strftime("%d%m%Y_%H%M%S")
        nome_arquivo = '{}/arq_result{}.csv'.format(path_arquivo, datahora)
        arq_result_csv = open(nome_arquivo, 'w')
        arq_result_csv.writelines(csv_result)

    @classmethod
    def gerar_excel(cls, header, dados, path_arquivo):
        tamanho_planilha = Util.definir_tamanho_planilha(dados)

        # criar arquivo excel
        datahora = datetime.datetime.now().strftime("%d%m%Y_%H%M%S")
        nome_arquivo = '{}/arq_result-{}.xlsx'.format(path_arquivo, datahora)

        workbook = Workbook(nome_arquivo)
        planilha = workbook.add_worksheet('QueryResult')

        # definir tamanho das colunas (col_ini, col_fim, tam)
        # planilha.set_column(0, tamanho_planilha[1], 15)
        
        # estilo_texto = workbook.add_format({'num_format': '@'})

        print(header)
        planilha.add_table(0, 0, tamanho_planilha[0], tamanho_planilha[1], 
                           {'data': list(dados.values()),
                            'columns': Util.header_para_xlsxwriter(header)
                            })
        
        planilha.autofit()

        workbook.close()
        os.system(f'start excel "{nome_arquivo}"')
    
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