from botcity.web.browsers.firefox import default_options
from webdriver_manager.firefox import GeckoDriverManager
from botcity.web import *
from datetime import datetime
from botcity.plugins.excel import *
import logging
from logging.handlers import RotatingFileHandler

class Bot:
    def bot(self):
        #  Logger Config Activity
        # Displayname: Diario Debug
        loggerBot = logging.getLogger("Cadastro de Forncedores")
        loggerBot.setLevel(logging.DEBUG)
        filelogging = RotatingFileHandler("fileLogging.log", maxBytes = 100000, backupCount = 10)
        filelogging.setLevel(logging.DEBUG)
        formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
        filelogging.setFormatter(formatter)
        loggerBot.addHandler(filelogging)

        # Open Browser Activity
        loggerBot.info("Start: Abrir o site de cadastro de Fornecedores")

        # Displayname: Abrir o site de cadastro de Fornecedores
        webDriverPath = GeckoDriverManager().install()
        webBot = WebBot()
        webBot.driver_path = webDriverPath
        webBot.browser = Browser.FIREFOX
        webBot.headless = False
        webBot.page_load_strategy = "Normal"
        webBotDef_options = default_options()
        webBot.options = webBotDef_options
        webBot.browse("https://jornadarpa.com.br/alunos/desafios/cadfor25/")

        loggerBot.info("End: Abrir o site de cadastro de Fornecedores")

        loggerBot.info("Start: Mapeamento dos elementos da pagina de Login")

        # DisplayName: Mapeamento dos elementos da pagina de Login

        # Sequence: Element list

        # Find Element Activity
        loggerBot.info("Start: Campo Email")

        # Displayname: Campo Email
        CampoEmail = webBot.find_element(selector="usuario", by=By.ID, waiting_time=1000, ensure_visible=False, ensure_clickable=False)

        loggerBot.info("End: Campo Email")

        # Find Element Activity
        loggerBot.info("Start: Campo Senha")

        # Displayname: Campo Senha
        CampoSenha = webBot.find_element(selector="senha", by=By.ID, waiting_time=1000, ensure_visible=False, ensure_clickable=False)

        loggerBot.info("End: Campo Senha")

        # Find Element Activity
        loggerBot.info("Start: Botão LGPD")

        # Displayname: Botão LGPD
        BotaoLGPD = webBot.find_element(selector="lgpd", by=By.ID, waiting_time=1000, ensure_visible=False, ensure_clickable=False)

        loggerBot.info("End: Botão LGPD")

        # Find Element Activity
        loggerBot.info("Start: Botão Login")

        # Displayname: Botão Login
        Botao = webBot.find_element(selector="btnLogin", by=By.ID, waiting_time=1000, ensure_visible=False, ensure_clickable=False)

        loggerBot.info("End: Botão Login")

        loggerBot.info("End: Mapeamento dos elementos da pagina de Login")

        loggerBot.info("Start: Entrada de dados de Login")

        # DisplayName: Entrada de dados de Login

        # Sequence: Action list

        # Type Into Activity
        loggerBot.info("Start: Digitação do campo Email")

        # Displayname: Digitação do campo Email
        CampoEmail.send_keys("participante@desafiosrpa.com.br")

        loggerBot.info("End: Digitação do campo Email")

        # Type Into Activity
        loggerBot.info("Start: Digitação do campo Senha")

        # Displayname: Digitação do campo Senha
        CampoSenha.send_keys("evento")

        loggerBot.info("End: Digitação do campo Senha")

        # Click Activity
        loggerBot.info("Start: Click in BotaoLGPD element")

        # Displayname: Click in BotaoLGPD element
        BotaoLGPD.click()

        loggerBot.info("End: Click in BotaoLGPD element")

        # Click Activity
        loggerBot.info("Start: Click in Botao element")

        # Displayname: Click in Botao element
        Botao.click()

        loggerBot.info("End: Click in Botao element")

        loggerBot.info("End: Entrada de dados de Login")

        # Read Excel Activity
        loggerBot.info("Start: Ler dados da planilha de Fornecedores")

        # Displayname: Ler dados da planilha de Fornecedores
        excelBot = BotExcelPlugin()
        file_or_path = "C:\\Users\\JEFFERSON LIMA\\Downloads\\Lista_exemplo.xlsx"

        listaFornecedores = excelBot.read(file_or_path=file_or_path).as_list(sheet="lista")[1:]
        loggerBot.info("End: Ler dados da planilha de Fornecedores")

        loggerBot.info("Start: Mapeamento dos elementos de cadastro dos Fornecedores")

        # DisplayName: Mapeamento dos elementos de cadastro dos Fornecedores

        # Sequence: Element list

        # Find Element Activity
        loggerBot.info("Start: Mapeamento do botão PF")

        # Displayname: Mapeamento do botão PF
        BotaoPF = webBot.find_element(selector="pf", by=By.ID, waiting_time=1000, ensure_visible=False, ensure_clickable=False)

        loggerBot.info("End: Mapeamento do botão PF")

        # Find Element Activity
        loggerBot.info("Start: Mapeamento do botão PJ")

        # Displayname: Mapeamento do botão PJ
        BotaoPJ = webBot.find_element(selector="pj", by=By.ID, waiting_time=1000, ensure_visible=False, ensure_clickable=False)

        loggerBot.info("End: Mapeamento do botão PJ")

        # Find Element Activity
        loggerBot.info("Start: Mapeamento do Nome/Razão")

        # Displayname: Mapeamento do Nome/Razão
        BotaoNomeRazao = webBot.find_element(selector="nomeRazao", by=By.ID, waiting_time=1000, ensure_visible=False, ensure_clickable=False)

        loggerBot.info("End: Mapeamento do Nome/Razão")

        # Find Element Activity
        loggerBot.info("Start: Mapeamento CPF/CNPJ")

        # Displayname: Mapeamento CPF/CNPJ
        BotaoCpfCnpj = webBot.find_element(selector="cpfCnpj", by=By.ID, waiting_time=1000, ensure_visible=False, ensure_clickable=False)

        loggerBot.info("End: Mapeamento CPF/CNPJ")

        # Find Element Activity
        loggerBot.info("Start: Mapeamento do botão Enviar")

        # Displayname: Mapeamento do botão Enviar
        BotaoEnviar = webBot.find_element(selector="btnEnviar", by=By.ID, waiting_time=1000, ensure_visible=False, ensure_clickable=False)

        loggerBot.info("End: Mapeamento do botão Enviar")

        loggerBot.info("End: Mapeamento dos elementos de cadastro dos Fornecedores")

        # ForEach Activity
        loggerBot.info("Start: ForEach")

        # Displayname: ForEach
        for linha in listaFornecedores:
            # Sequence: Body

            loggerBot.info("Start: Entrada dos dados do Fornecedor")

            # DisplayName: Entrada dos dados do Fornecedor

            # Sequence: Action list

            # Sequence: Conditional Structure

            # If Activity
            loggerBot.info("Start: If Condition")

            # Displayname: If Condition
            if linha[0]== "PF":
                # Sequence: Body

                # Click Activity
                loggerBot.info("Start: Click in BotaoPF element")

                # Displayname: Click in BotaoPF element
                BotaoPF.click()

                loggerBot.info("End: Click in BotaoPF element")


                loggerBot.info("End: If Condition")

            # Else Activity
            # Displayname: Else
            else:
                loggerBot.info("Start: Else")

                # Sequence: Body

                # Click Activity
                loggerBot.info("Start: Click in BotaoPJ ")

                # Displayname: Click in BotaoPJ 
                BotaoPJ.click()

                loggerBot.info("End: Click in BotaoPJ ")


                loggerBot.info("End: Else")

            # Type Into Activity
            loggerBot.info("Start: Type Into BotaoNomeRazao field")

            # Displayname: Type Into BotaoNomeRazao field
            BotaoNomeRazao.send_keys(linha[1])

            loggerBot.info("End: Type Into BotaoNomeRazao field")

            #  Write Log Activity
            # Displayname: WriteLog
            loggerBot.debug(linha[1])

            # Type Into Activity
            loggerBot.info("Start: Type Into BotaoCpfCnpj field")

            # Displayname: Type Into BotaoCpfCnpj field
            BotaoCpfCnpj.send_keys(linha[2])

            loggerBot.info("End: Type Into BotaoCpfCnpj field")

            # Click Activity
            loggerBot.info("Start: Click in BotaoEnviar element")

            # Displayname: Click in BotaoEnviar element
            BotaoEnviar.click()

            loggerBot.info("End: Click in BotaoEnviar element")

            loggerBot.info("End: Entrada dos dados do Fornecedor")


        loggerBot.info("End: ForEach")

        # Wait Activity
        loggerBot.info("Start: Wait")

        # Displayname: Wait
        webBot.wait(3000)

        loggerBot.info("End: Wait")



        logging.shutdown()

        return
if __name__ == '__main__':
    bot = Bot()
    bot.bot()