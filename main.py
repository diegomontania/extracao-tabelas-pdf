import camelot
import os
import logging
from datetime import datetime

# setando caminhos
caminhoPastaLog = "./logs"

# setando log
def ConfiguracoesLog(caminhoPastaLog):
    # cria pasta de log
    if not os.path.exists(caminhoPastaLog):
        os.makedirs(caminhoPastaLog)

    # cria arquivo de log
    logging.basicConfig(handlers=[
        logging.FileHandler(filename="./" + caminhoPastaLog + "/" + datetime.today().strftime('%Y-%m-%d') + ".txt",
                            encoding='utf-8', mode='a+')],
        format="%(asctime)s %(name)s:%(levelname)s:%(message)s",
        datefmt="%F %A %T",
        level=logging.INFO)

    logger = logging.getLogger(__name__)
    return logger

# executa extracao de tabelas dos pdfs para exportar em .xlsx
def ExecutaExtracaoTabelasPdf(objetoLog, caminhoTodosArquivos, caminhoFinalExportacao):
    # recebe todos os arquivos da pasta
    todosOsArquivos = os.listdir(caminhoTodosArquivos)

    # logando o total de arquivos encontrados na pasta, recebendo o valor total de arquivos da lista e convertendo
    # pra string aqui
    objetoLog.info('Total arquivos encontrados ' + str(len(todosOsArquivos)))

    for arquivoAtual in todosOsArquivos:
        caminhoArquivoAtual = caminhoTodosArquivos + '\\' + arquivoAtual

        objetoLog.info('Lendo arquivo ' + arquivoAtual)

        # faz a extracao e retorna tabela, possivelmente ajeitar essa tabela para que fique melhor organizada
        tabelas = camelot.read_pdf(caminhoArquivoAtual, pages='all', encoding="iso8859-1")

        # exporta para .xlsx
        tabelas.export(caminhoFinalExportacao + '\\' + arquivoAtual.replace('.pdf', '') + '.xlsx', f='excel',
                       compress=False)

        objetoLog.info('Exportado com sucesso ' + arquivoAtual.replace('.pdf', '') + '.xlsx')


# comeco do programa
def main():
    try:
        print('Executando o programa')

        # recebe objeto de logging
        logger = ConfiguracoesLog(caminhoPastaLog)

        # recebe os parametros
        caminhoArquivos = "C:\\Users\\" + os.getlogin() + "\\Downloads\\TestePdfs"
        caminhoFinalExportacao = "C:\\Users\\" + os.getlogin() + "\\Downloads"

        # recebe o texto extraido do arquivo
        ExecutaExtracaoTabelasPdf(logger, caminhoArquivos, caminhoFinalExportacao)

        logger.info('Processo de extracao via python finalizado com sucesso.')
        print('Processo de extracao via python finalizado com sucesso')

    except Exception as ex:
        print(ex)
        print('Nao executado - Sem parametros de caminhos!')
        logger.error(ex)
        pass

if __name__ == "__main__":
    main()