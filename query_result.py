import datetime

def query_to_excel():

    # modelo a trabalhar
    #                 +-----------------------------------------------------------
    # (cabeçalho)     ³    COD_BIN_INF      ³    COD_BIN_SUPR     ³ IND_BIN_ATVO ³       
    #                 +-----------------------------------------------------------
    # (linha)       1_³ 1234567890123456789 ³ 1234567890123456789 ³ A            ³       
    #                 +-----------------------------------------------------------
    arq_query_result = open('D:\\SC\\query_resut.txt', 'r')

    query_result = arq_query_result.readlines()

    # despreza primeira coluna do header '#'
    header = Util.procura_header(query_result)[1:]
    csv_result = []
    csv_result.append(';'.join(header) + '\n')

    dados = Util.procura_linhas(query_result)
    for chave in dados:
        csv_result.append(';'.join(dados[chave]) + '\n')

    datahora = datetime.datetime.now().strftime("%d%m%Y_%H%M%S")
    nome_arquivo = 'arq_result{}.txt'.format(datahora)
    arq_result_csv = open(nome_arquivo, 'w')
    arq_result_csv.writelines(csv_result)


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
                num_linha = colunas[0][:-1]
                if num_linha not in linhas:
                    linhas[num_linha] = colunas[1:]
                else:
                    linhas[num_linha] = linhas[num_linha] + colunas[1:]
        return linhas                


if __name__ == '__main__':
    query_to_excel()