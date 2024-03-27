
### IMPORTS EXTERNOS BOTCITY ###

import os
import shutil
from os import path
from datetime import datetime
from twocaptcha import TwoCaptcha

### IMPORTS INTERNOS BOTCITY ###

import botcity.web
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from botcity.web import WebBot, Browser, By
from botcity.web.util import element_as_select
from botcity.web.parsers import table_to_dict
from botcity.plugins.excel import BotExcelPlugin
from botcity.plugins.files import BotFilesPlugin
from botcity.document_processing import *
from botcity.web.browsers.chrome import default_options

# Usando o Maestro SDK
from botcity.maestro import *

# Instância da classe Bot(WebBot)
class Bot(WebBot):
    def action(self, execution=None):
        
        bot = WebBot
        
        #--------------------------------#
        #   VARIÁVEIS LOCAIS E GLOBAIS   #
        #--------------------------------#
        
        # Variáveis globais

        url_site_prefeitura = "http://egov2.dourados.ms.gov.br/webAtendimento/"
        path_arquivo_img_captcha = r"PATH_IMG_CAPTCHA"
        api_key_twoCaptcha = "API_KEY"

        # Modo Headless
        self.headless = False
        
        # Navegador usado no processo
        self.browser = Browser.CHROME

        # Path do arquivo executável do chromedriver
        self.driver_path = r"C:\Botcity\bots\chromedriver.exe"

        # Path dos boletos baixados
        self.download_folder_path = r"PATH_BOLETOS"

        # Path das planilhas usadas no processo
        path_planilha_base = r'PATH_PLANILHA_BASE'
        path_planilha_resultados = r'PATH_PLANILHA_RESULTADOS'

        # Transforma as planilhas de Status, Resultados e Base de dados em listas
        planilha_status = BotExcelPlugin('Status da extração').read(path_planilha_resultados)
        lista_planilha_status = planilha_status.as_list()

        planilha_resultados = BotExcelPlugin('Resultados da extração').read(path_planilha_resultados)
        lista_planilha_resultados = planilha_resultados.as_list()

        planilha_base = BotExcelPlugin("IPTU").read(path_planilha_base)
        lista_planilha_base = planilha_base.as_list()

        # Armazena o número da linha que deve ser preenchida nas planilhas
        linha_saida_status_extracao = len(lista_planilha_status) + 1
        linha_saida_resultado_extracao = len(lista_planilha_resultados) + 1
        
        #-------------#
        #   MÉTODOS   #
        #-------------#

        def obtem_data_vencimento_cliente():
            data_referencia = input("Digite a data do último dia do mês: ")
            data_referencia_formatada = datetime.strptime(data_referencia, '%d/%m/%Y').date()
            dia_data_referencia  = int(data_referencia_formatada.day)
            mes_data_referencia  = int(data_referencia_formatada.month)
            ano_data_referencia  = int(data_referencia_formatada.year)
            dataReferencia = datetime(ano_data_referencia, mes_data_referencia, dia_data_referencia).date()

            return mes_data_referencia
        
        def obtem_data_vencimento_boleto(vencimento):
            data_vencimento_formatado = datetime.strptime(vencimento, '%d/%m/%Y').date()
            dia_vencimento_formatado  = int(data_vencimento_formatado.day)
            mes_vencimento_formatado  = int(data_vencimento_formatado.month)
            ano_vencimento_formatado  = int(data_vencimento_formatado.year)
            data_vencimento = datetime(ano_vencimento_formatado, mes_vencimento_formatado, dia_vencimento_formatado).date()

            return mes_vencimento_formatado
        
        def login_site_prefeitura(url_site_prefeitura):
            self.browse(url_site_prefeitura)
            self.wait(2000)

            # Entra no contexto do iframe "menu"
            menu_iframe = self.find_element("menu", By.ID)
            self.enter_iframe(menu_iframe)
        
        def verificacao_erro_http_500(contrato, inscricao_imobiliaria, linha_saida_status_extracao):
            menu_login_site_prefeitura = self.find_element("/html/body/div[2]/ul/li[2]/a", By.XPATH)
            if menu_login_site_prefeitura == None:
                planilha_resultados.set_active_sheet('Status da extração')
                planilha_resultados.set_cell("A", linha_saida_status_extracao, contrato, sheet="Status da extração")
                planilha_resultados.set_cell("B", linha_saida_status_extracao, inscricao_imobiliaria, sheet="Status da extração")
                planilha_resultados.set_cell("D", linha_saida_status_extracao, "Erro ao carregar site (Erro HTTP 500)", sheet="Status da extração")
                planilha_resultados.write(path_planilha_resultados)
                linha_saida_status_extracao += 1
                erro_http_500 = True
                
                return erro_http_500
            
        def preenche_campos_login(inscricao_imobiliaria):
            # Clica no menu "IPTU 2022" e sai do iframe
            self.find_element("/html/body/div[2]/ul/li[2]/a", By.XPATH).click()
            self.leave_iframe()
            
            # Entra no contexto do iframe "form"
            formFrame = self.find_element("form", By.ID)
            self.enter_iframe(formFrame)

            # Preenche campo cadastro código
            self.find_element("/html/body/div[1]/form/input[1]", By.XPATH).send_keys(inscricao_imobiliaria)
            
        def quebra_normal_captcha(api_key_twoCaptcha, path_arquivo_img_captcha):   
            solver = TwoCaptcha(api_key_twoCaptcha)
            try:
                img_captcha = self.find_element("/html/body/div[1]/form/table/tbody/tr/td[2]/img", By.XPATH)
                captcha = img_captcha.screenshot(path_arquivo_img_captcha)
                resultado_quebra_captcha = solver.normal(path_arquivo_img_captcha)
            
            except Exception as erroCaptcha:
                print(erroCaptcha)
                print('> Erro ao quebrar captcha!')
            else:
                tolken_captcha = resultado_quebra_captcha['code']
                
                # Apaga o arquivo "kaptcha.png" armezenado na pasta de boletos baixados
                if os.path.isfile(path_arquivo_img_captcha):
                    os.remove(path_arquivo_img_captcha)
            
            self.find_element("/html/body/div[1]/form/table/tbody/tr/td[1]/input", By.XPATH).send_keys(tolken_captcha)
            
            # Clica no botão para pesquisar
            self.find_element("/html/body/div[1]/form/input[3]", By.XPATH).click()

        def verificacao_erro_de_pesquisa(contrato, inscricao_imobiliaria, linha_saida_status_extracao):
            div_resposta_erro = self.find_element('/html/body/div/form/div[1]/p ', By.XPATH)

            if div_resposta_erro != None:
                texto_div_resposta_erro = self.find_element('/html/body/div/form/div[1]/p ', By.XPATH).text
                planilha_resultados.set_active_sheet('Status da extração')
                planilha_resultados.set_cell("A", linha_saida_status_extracao, contrato, sheet="Status da extração")
                planilha_resultados.set_cell("B", linha_saida_status_extracao, inscricao_imobiliaria, sheet="Status da extração")
                if 'C Ó D I G O    D E    S E G U R A N Ç A' in texto_div_resposta_erro:
                    planilha_resultados.set_cell("D", linha_saida_status_extracao, 'Falha ao quebrar captcha', sheet="Status da extração")
                else:
                    planilha_resultados.set_cell("D", linha_saida_status_extracao, texto_div_resposta_erro, sheet="Status da extração")
                planilha_resultados.write(path_planilha_resultados)
                div_resposta_erro = True
                
                return div_resposta_erro, texto_div_resposta_erro
            
        def verificacao_existencia_tabela(contrato, inscricao_imobiliaria, linha_saida_status_extracao):
            tabela_boletos = self.find_element('/html/body/div/form/div[1]/table[3]', By.XPATH)

            if tabela_boletos == None:
                planilha_resultados.set_active_sheet('Status da extração')
                planilha_resultados.set_cell("A", linha_saida_status_extracao, contrato, sheet="Status da extração")
                planilha_resultados.set_cell("B", linha_saida_status_extracao, inscricao_imobiliaria, sheet="Status da extração")
                planilha_resultados.set_cell("D", linha_saida_status_extracao, 'Erro ao pesquisar por inscrição (tabela de parecelas não gerada)', sheet="Status da extração")
                planilha_resultados.write(path_planilha_resultados)
                
                return False, linha_saida_status_extracao
            else:
                return tabela_boletos, True, linha_saida_status_extracao
            
        def marca_desmarca_boleto(contador_parcela):
            campo_selecao = self.find_element(f"/html/body/div/form/div[1]/table[3]/tbody/tr[{contador_parcela}]/td[1]/input", By.XPATH).click()

        def baixa_boleto():
            # Clica em "Gerar Guia DAM"                    
            files = BotFilesPlugin()

            with files.wait_for_file(directory_path=self.download_folder_path, file_extension=".pdf", timeout=8000):
                self.find_element("/html/body/div/form/div[2]/input[2]", By.XPATH,).click()

        def verificacao_erro_baixa_boleto(contrato, inscricao_imobiliaria, parcela, contador_parcela, linha_saida_status_extracao):
            campo_selecao = self.find_element(f"/html/body/div/form/div[1]/table[3]/tbody/tr[{contador_parcela}]/td[1]/input", By.XPATH, waiting_time=3000)
                        
            if campo_selecao == None:
                planilha_resultados.set_active_sheet('Status da extração')
                planilha_resultados.set_cell("A", linha_saida_status_extracao, contrato, sheet="Status da extração")
                planilha_resultados.set_cell("B", linha_saida_status_extracao, inscricao_imobiliaria, sheet="Status da extração")
                planilha_resultados.set_cell("C", linha_saida_status_extracao, parcela, sheet="Status da extração")
                planilha_resultados.set_cell("D", linha_saida_status_extracao, "Erro ao fazer download do boleto", sheet="Status da extração")
                planilha_resultados.write(path_planilha_resultados)

                return False

        def renomea_boleto():
            # Verifica se há duplicidade de boleto. Se não existe, excluí o arquivo original, copia o boleto baixado e renomea o boleto copiado.
            arquivoBaixado = self.get_last_created_file()
            print('> Download do boleto: SUCESSO\n')
            arquivoRenomeado = os.path.join(self.download_folder_path, f"Boleto_IPTU_{inscricao_imobiliaria} ({parcela}).pdf")            
            
            if os.path.exists(arquivoRenomeado):
                os.remove(arquivoBaixado)
                print('> ATENÇÃO - Inscrição imobiliária duplicada na base.\n')
                inscricao_duplicada = True
                return inscricao_duplicada
            else:                         
                arquivoCopiado = shutil.copy2(arquivoBaixado, arquivoRenomeado)
                self.wait(1500)
                os.remove(arquivoBaixado)
                print('> Renomeação do boleto: SUCESSO\n')
                return arquivoCopiado
        
        def extrai_dados_boleto():
            reader = PDFReader()                        
            parser = reader.read_file(arquivo_renomeado)
            _autenticacao_mecanica = parser.get_first_entry("Autenticação Mecânica")
            linha_digitavel = parser.read(_autenticacao_mecanica, -4.489247, -1.5, 3.650538, 4.333333)
            linha_digitavel = linha_digitavel.replace(".","")
            linha_digitavel = linha_digitavel.replace(" ","")

            for numero in linha_digitavel:
                codigo_barras_1 = linha_digitavel[:4]
                codigo_barras_2 = linha_digitavel[32:]
                codigo_barras_3 = linha_digitavel[4:9]+linha_digitavel[10:20]+linha_digitavel[21:31]
                codigo_barras = codigo_barras_1 + codigo_barras_2 + codigo_barras_3
                codigo_barras = codigo_barras.replace(".","")
                codigo_barras = codigo_barras.replace(" ","")

            return linha_digitavel, codigo_barras

        #------------------------#
        #   INÍCIO DO PROCESSO   #
        #------------------------#

        # Obtém o perído (dia/mês/ano) dos boletos que devem ser baixados (Referenciados pelo cliente)
        mes_data_referencia = obtem_data_vencimento_cliente()

        # Inicia um loop que intera sob a planilha base para baixar os boletos
        for linha in lista_planilha_base:
            if "Contrato" in str(linha[0]):
                continue
            else:
                contrato = linha[0]
                inscricao_imobiliaria = linha[1]

                # Printa os dados da iteração
                print("\n# DADOS ANALISADOS #\n")
                print("> Inscrição Imobiliária:", inscricao_imobiliaria)
                print("> Linha de saída - Status:", linha_saida_status_extracao)
                print("> Linha de saída - Resultados:", linha_saida_resultado_extracao)

                # Entra no site da prefeitura
                login_site_prefeitura(url_site_prefeitura)

                # Espera a resposta sobre erro HTTP 500
                resposta_verificacao_erro_http_500 = verificacao_erro_http_500(contrato, inscricao_imobiliaria, linha_saida_status_extracao)
                if resposta_verificacao_erro_http_500 == True:
                    linha_saida_status_extracao += 1
                    continue
                
                # Preenche os campos para fazer login
                preenche_campos_login(inscricao_imobiliaria)

                # Quebra o normal captcha e clica no botão "Pesquisar"
                quebra_normal_captcha(api_key_twoCaptcha, path_arquivo_img_captcha)

                # Espera o erro sobre pesquisa de inscrição e geração da tabela de boletos
                resposta_verificacao_erro_de_pesquisa = verificacao_erro_de_pesquisa(contrato, inscricao_imobiliaria, linha_saida_status_extracao)
                if resposta_verificacao_erro_de_pesquisa == None:
                    pass
                elif resposta_verificacao_erro_de_pesquisa[0] == True:
                    print(f'> Login Site: FALHA ({resposta_verificacao_erro_de_pesquisa[1]})\n')
                    linha_saida_status_extracao += 1
                    continue
                
                resposta_verificacao_existencia_tabela = verificacao_existencia_tabela(contrato, inscricao_imobiliaria, linha_saida_status_extracao)
                if resposta_verificacao_existencia_tabela == True:
                    pass
                elif resposta_verificacao_existencia_tabela[1] == False:
                    print('> Login Site: FALHA (Erro ao pesquisar por inscrição (tabela de parecelas não gerada))\n')
                    linha_saida_status_extracao += 1
                    continue

                print('Login site: SUCESSO\n')

                # Transforma a planilha de parcelas de IPTU em dicionário
                dicionario_parcela_tabela_boletos = table_to_dict(table=resposta_verificacao_existencia_tabela[0], has_header=False, skip_rows=1)

                # Inicializa/reinicia os valores das variáveis contadoras
                contador_parcela = 2
                contador_iteracoes = 0
                existencia_parcela = False

                # Inicia um loop que intera sob o dicionário para buscar e baixar os boletos disponíveis
                for linha in dicionario_parcela_tabela_boletos:
                    numero_boletos_disponiveis = len(dicionario_parcela_tabela_boletos)                               
                    # Percorre tabelaWeb e baixa os boletos
                    imposto = linha['col_1']                 
                    exercicio = linha['col_2']
                    parcela = linha['col_3']
                    parcela = parcela.replace(" ","")
                    vencimento = linha['col_4']
                    vencimento = vencimento.replace(" ","")
                    valorOriginal = linha['col_5']
                    valorAtualizado = linha['col_6']
                    juros = linha['col_7']
                    desconto = linha['col_8']
                    valorTotal = linha['col_9']

                    # Obtém e formata a data de vencimento do boleto (No site)
                    mes_vencimento_formatado = obtem_data_vencimento_boleto(vencimento)

                    # Filtra o boleto desejado pelo cliente
                    if mes_vencimento_formatado == mes_data_referencia:
                        # Seleciona o boleto
                        marca_desmarca_boleto(contador_parcela) 

                        # Baixa o boleto
                        baixa_boleto()

                        # Espera pelo erro ao tentar baixar boleto
                        resposta_verificacao_erro_baixa_boleto = verificacao_erro_baixa_boleto(contrato, inscricao_imobiliaria, parcela, contador_parcela, linha_saida_status_extracao)

                        if resposta_verificacao_erro_baixa_boleto == False:
                            print('> Baixar boleto: FALSE\n')
                            linha_saida_status_extracao += 1
                            break
                        else:
                            print('> Baixar boleto: SUCESSO\n') 

                        # Desmarca o boleto
                        marca_desmarca_boleto(contador_parcela)

                        # Renomea boleto
                        arquivo_renomeado = renomea_boleto()
                        
                        # Extrai informações dos boletos
                        resultado_extracao_dados_boleto = extrai_dados_boleto()

                        # Insere dados obtidos na planilha statusExtracao
                        planilha_resultados.set_active_sheet('Status da extração')
                        planilha_resultados.set_cell("A", linha_saida_status_extracao, contrato, sheet="Status da extração")
                        planilha_resultados.set_cell("B", linha_saida_status_extracao, inscricao_imobiliaria, sheet="Status da extração")
                        planilha_resultados.set_cell("C", linha_saida_status_extracao, parcela, sheet="Status da extração")
                        planilha_resultados.set_cell("D", linha_saida_status_extracao, "Boleto baixado", sheet="Status da extração")

                        # Insere dados obtidos na planilha resultadoStatus
                        planilha_resultados.set_active_sheet('Resultados da extração')
                        planilha_resultados.set_cell("A", linha_saida_resultado_extracao, contrato, sheet="Resultados da extração")
                        planilha_resultados.set_cell("B", linha_saida_resultado_extracao, inscricao_imobiliaria, sheet="Resultados da extração")
                        planilha_resultados.set_cell("C", linha_saida_resultado_extracao, imposto, sheet="Resultados da extração")
                        planilha_resultados.set_cell("D", linha_saida_resultado_extracao, exercicio, sheet="Resultados da extração")
                        planilha_resultados.set_cell("E", linha_saida_resultado_extracao, parcela, sheet="Resultados da extração")
                        planilha_resultados.set_cell("F", linha_saida_resultado_extracao, vencimento, sheet="Resultados da extração")
                        planilha_resultados.set_cell("G", linha_saida_resultado_extracao, valorOriginal, sheet="Resultados da extração")
                        planilha_resultados.set_cell("H", linha_saida_resultado_extracao, valorAtualizado, sheet="Resultados da extração")
                        planilha_resultados.set_cell("I", linha_saida_resultado_extracao, juros, sheet="Resultados da extração")
                        planilha_resultados.set_cell("J", linha_saida_resultado_extracao, desconto, sheet="Resultados da extração")
                        planilha_resultados.set_cell("K", linha_saida_resultado_extracao, valorTotal, sheet="Resultados da extração")
                        planilha_resultados.set_cell("L", linha_saida_resultado_extracao, resultado_extracao_dados_boleto[0], sheet="Resultados da extração")
                        planilha_resultados.set_cell("M", linha_saida_resultado_extracao, resultado_extracao_dados_boleto[1], sheet="Resultados da extração")
                        planilha_resultados.write(path_planilha_resultados)

                        # Incrementa as variáveis contadoras
                        linha_saida_status_extracao += 1
                        linha_saida_resultado_extracao += 1                        
                        contador_parcela += 1
                        contador_iteracoes += 1
                        existencia_parcela = True  

                    else:
                        contador_iteracoes += 1
                        contador_parcela += 1

                        if contador_iteracoes == numero_boletos_disponiveis and existencia_parcela != True:
                            # Insere dados obtidos na planilha statusExtracao
                            planilha_resultados.set_active_sheet('Status da extração')
                            planilha_resultados.set_cell("A", linha_saida_status_extracao, contrato, sheet="Status da extração")
                            planilha_resultados.set_cell("B", linha_saida_status_extracao, inscricao_imobiliaria, sheet="Status da extração")
                            planilha_resultados.set_cell("D", linha_saida_status_extracao, "Nenhum boleto deste mês disponível", sheet="Status da extração")
                            planilha_resultados.write(path_planilha_resultados)
                            linha_saida_status_extracao += 1
                            continue

if __name__ == '__main__':
    Bot.main()
