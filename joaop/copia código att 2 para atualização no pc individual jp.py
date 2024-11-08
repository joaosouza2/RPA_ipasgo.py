import logging
import pandas as pd
import time
import os
import re
import shutil
import datetime
import inspect  
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException,ElementClickInterceptedException
from selenium.webdriver.common.keys import Keys
from openpyxl import load_workbook
from selenium.webdriver.common.action_chains import ActionChains
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox


logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


#CONFIGURAÇÕES PARA RETIRADA DE LOGGINGS EXCESSIVOS DE EXECUÇÃO. 
####################################################################################
#class CustomFilter(logging.Filter):
#    def filter(self, record):
#        # Filtra apenas as mensagens de log do script principal
#        return record.name == '__main__'
#logger = logging.getLogger()
#logger.addFilter(CustomFilter())
#####################################################################################

def encontrar_arquivos_paciente(caminho_pasta, id_paciente):
#################################################################################################
    # Busca de relatórios /RM e RC para anexo de documentos dentro da automação, englobam os casos de não padronização de tamanho, espaço de caracteres, e formato de letras.
    # Não aborda o caso de escrita invertida
    # Os arquivos dentro do relatório clínico precisam de padronização para evitar erros futuros.
#################################################################################################
    p = Path(caminho_pasta)

    arquivos_rm = []
    arquivos_rc_fono = []
    arquivos_rc_psi = []
    arquivos_rc_to = []

    # Padrão para arquivos RM usando expressões regulares, permitindo espaços e traços opcionais
    padrao_rm = rf'(?i)^RM[\s-]*{id_paciente}.*\..+$'

    # Padrões para os três tipos de documentação usando expressões regulares, permitindo espaços e traços e tamanho de caracteres diferentes
    padrao_rc_fono = rf'(?i)^RC[\s-]*{id_paciente}[\s-]*.*FONO.*\..+$'
    padrao_rc_psi = rf'(?i)^RC[\s-]*{id_paciente}[\s-]*.*PSI.*\..+$'
    padrao_rc_to = rf'(?i)^RC[\s-]*{id_paciente}[\s-]*.*TO.*\..+$'

    logging.info(f"Verificando arquivos na pasta: {caminho_pasta}")

    for arquivo in p.iterdir():
        if arquivo.is_file():
            nome_arquivo = arquivo.name
            logging.info(f"Arquivo encontrado: {nome_arquivo}")

            # Verifica se é um arquivo RM
            if re.match(padrao_rm, nome_arquivo):
                arquivos_rm.append(arquivo)
                logging.info(f"Arquivo RM identificado: {arquivo}")
                continue

            # Verifica se é um arquivo RC e classifica de acordo com FONO, PSI ou TO
            if re.match(padrao_rc_fono, nome_arquivo):
                arquivos_rc_fono.append(arquivo)
                logging.info(f"Arquivo RC FONO identificado: {arquivo}")
                continue

            if re.match(padrao_rc_psi, nome_arquivo):
                arquivos_rc_psi.append(arquivo)
                logging.info(f"Arquivo RC PSI identificado: {arquivo}")
                continue

            if re.match(padrao_rc_to, nome_arquivo):
                arquivos_rc_to.append(arquivo)
                logging.info(f"Arquivo RC TO identificado: {arquivo}")
                continue

    return arquivos_rm, arquivos_rc_fono, arquivos_rc_psi, arquivos_rc_to




class BaseAutomation:
    def __init__(self):
        """Configurações gerais do WebDriver."""
        self.options = Options()
        self.options.add_argument("--start-maximized")
        self.driver = webdriver.Chrome(options=self.options)

#####################################################################################################
        #self.options = Options()
        #self.options.add_argument("--headless")  # Executa em modo headless
        #self.options.add_argument("--window-size=1366,768")  # Define a resolução de tela
        #self.options.add_argument("--disable-gpu")  # Desativa aceleração por GPU
        #self.options.add_argument("--no-sandbox")  # Ignora o sandboxing
        #self.options.add_argument("--disable-dev-shm-usage")  # Evita problemas de recursos compartilhados
        #self.driver = webdriver.Chrome(options=self.options)
########################################################################################################




    def wait_for_stability(self, timeout=10, check_interval=1):
        """Espera pela estabilidade da altura da página."""
        old_height = self.driver.execute_script("return document.body.scrollHeight;")
        for _ in range(timeout):
            time.sleep(check_interval)
            new_height = self.driver.execute_script("return document.body.scrollHeight;")
            if new_height == old_height:
                break
            old_height = new_height
        

    def safe_click(self, by_locator):
        """Tenta clicar no elemento várias vezes se for interceptado."""
        for _ in range(3):
            try:
                element = WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable(by_locator))
                element.click()
                logging.info(f"Elemento clicado com sucesso: {by_locator}")
                return
            except Exception as e:
                time.sleep(1)
        raise Exception("Não foi possível clicar no elemento após várias tentativas.")

    def acessar_com_reattempt(self, by_locator, attempts=3):
        """Tenta acessar um elemento várias vezes."""
        for attempt in range(attempts):
            try:
                element = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(by_locator))
                logging.info(f"Elemento encontrado: {by_locator}")
                return element
            except TimeoutException as e:
                logging.warning(f"Tentativa {attempt + 1} falhou. Tentando novamente...")
                time.sleep(2)
        raise Exception(f"Não foi possível acessar o elemento após {attempts} tentativas.") 

    def close(self):
        """Fecha o navegador."""
        logging.info("O navegador vai fechar de qualquer maneira. Avisei...")
        self.driver.quit()



class IpasgoAutomation(BaseAutomation):

    # Definição dos elementos de erro
    ERROR_ELEMENTS = {
        "numero_guia_do_prestador": '//*[@id="liNumGuiaPrincipalIgualNumGuiaPrestador"]',
        "guia_principal_invalida": '//*[@id="liNumGuiaPrincipalInvalido"]',
        "informe_o_beneficiario": '//*[@id="liBeneficiario"]',
        "informe_se_e_atendimento_rn": '//*[@id="liAtendimentoRN"]',
        "informe_os_dados_do_solicitante": '//*[@id="liDadosSolicitante"]',
        "informe_o_profissional_solicitante": '//*[@id="liProfissionalSolicitante"]',
        "informe_a_especialidade_CBO": '//*[@id="liCBO"]',
        "informe_carater_de_atendimento": '//*[@id="liCaraterAtendimento"]',
        "informe_a_data_de_solicitacao": '//*[@id="liDataSolicitacao"]',
        "informe_o_PMSCO": '//*[@id="liPCMSO"]',
        "informe_pelo_menos_um_procedimento": '//*[@id="liProcedimento"]',
        "informe_o_contratado_executante": '//*[@id="liContratadoExecutante"]',
        "informe_o_tipo_de_atendimento": '//*[@id="liTipoAtendimento"]',
        "informe_o_tipo_de_consulta": '//*[@id="liTipoConsulta"]',
        "pelo_menos_um_profissional_executante": '//*[@id="liProfissionalExecutante"]',
        "cancele_ou_confirme_o_procedimento": '//*[@id="liGridProcedimentosEmEdicao"]',
        "cancele_ou_confirme_profissional_executante": '//*[@id="liGridProfissionaisExecutanteEmEdicao"]',
        "insira_documento_anexo_a_guia": '//*[@id="liObrigaAnexo"]',
        "informe_o_numero_do_beneficiario": '//*[@id="liNumeroBeneficiario"]',
        "informe_o_regime_de_atendimento": '//*[@id="liRegimeAtendimento"]'
    }

    def __init__(self, login, password, start_row, end_row):
        super().__init__()
        self.login = login
        self.password = password
        self.start_row = start_row
        self.end_row = end_row
        self.row_index = self.start_row

    
        self.file_path = r"C:\Users\SUPERVISÃO ADM\Desktop\SOLICITACOES_AUTORIZACAO_FACPLAN_ATUALIZADO.xlsx" # Caminho do arquivo excel
        self.sheet_name = 'AUTORIZACOES' #aba da planilha que será usada para extração dos dados
        self.txt_file_path = os.path.join(r"C:\Users\SUPERVISÃO ADM\Desktop\números_guias_test.txt")  # Caminho do arquivo txt

        # Ler o arquivo Excel original
        self.df = pd.read_excel(
            self.file_path,
            sheet_name=self.sheet_name,
            header=0,
            dtype={'CARTEIRA': str} #coluna CARTEIRA permanecerá com os zeros a esquerda, visto que as demais colunas não seguem o mesmo padrão devido ao site do ipasgo.
        )

 
        self.df.columns = [col.upper().strip() for col in self.df.columns]# Normaliza os nomes das colunas para maiúsculas
        
        # Adiciona a coluna 'ERRO' se não existir
        if 'ERRO' not in self.df.columns:
            self.df['ERRO'] = ''
            
        #adiciona a coluna GUIA_COD para armazenar os numeros das guias soclitadas 
        if 'GUIA_COD' not in self.df.columns:
            self.df['GUIA_COD'] = ''

        #incializa pelo indice de linha atual
        self.row_index = self.start_row
        


    def get_excel_value(self, column_name):
        try:
            value = str(self.df[column_name].iloc[self.row_index])
            return value
        except KeyError:
            logging.error(f"A coluna '{column_name}' não foi encontrada no arquivo Excel.")
            return ""



    def acessar_portal_ipasgo(self):
        """Executa o fluxo principal do IPASGO."""
        try:
           
            self.driver.get("https://portalos.ipasgo.go.gov.br/Portal_Dominio/PrestadorLogin.aspx")
            self.wait_for_stability(timeout=10)

            matricula_input = self.acessar_com_reattempt((By.ID, "SilkUIFramework_wt13_block_wtUsername_wtUserNameInput2"))
            matricula_input.send_keys(self.login)

            senha_input = self.acessar_com_reattempt((By.ID, "SilkUIFramework_wt13_block_wtPassword_wtPasswordInput"))
            senha_input.send_keys(self.password)

            self.safe_click((By.ID, "SilkUIFramework_wt13_block_wtAction_wtLoginButton"))   
       

            # Verificar se o alerta está dentro de um iframe (opcional)
            try:
                # Localiza todos os iframes na página
                iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
                for iframe in iframes:
                    self.driver.switch_to.frame(iframe)
                    try:
                        fechar_alerta = WebDriverWait(self.driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, "//a[contains(@id, 'wt15')]/span[contains(@class, 'fa-close')]"))
                        )
                        fechar_alerta.click()
                        logging.info("Alerta detectado e fechado dentro de um iframe.")
                        self.driver.switch_to.default_content()
                        self.wait_for_stability(timeout=5)
                        break  # Sai do loop após fechar o alerta
                    except TimeoutException:
                        self.driver.switch_to.default_content()
                        continue
            except Exception as e:
                pass
            
            self.wait_for_stability(timeout=10)

            link_portal_webplan = self.acessar_com_reattempt((By.XPATH, "//*[@id='IpasgoTheme_wt16_block_wtMainContent_wtSistemas_ctl08_SilkUIFramework_wt36_block_wtActions_wtModulos_SilkUIFramework_wt9_block_wtContent_wtModuloPortalTable_ctl04_wt2']/span"))
            self.driver.execute_script("arguments[0].scrollIntoView(true);", link_portal_webplan)
            time.sleep(2)
            link_portal_webplan.click()

            WebDriverWait(self.driver, 20).until(EC.number_of_windows_to_be(2))
            self.driver.switch_to.window(self.driver.window_handles[1])

            self.acessar_com_reattempt((By.ID, "menuPrincipal"))

            time.sleep(3)

            self.acessar_guias()

        except Exception as e:
            pass



    def acessar_guias(self):
        """Acessa a aba 'Guias' no menu principal."""
        try:
            guias_tab = self.acessar_com_reattempt((By.CSS_SELECTOR, ".menuLink.grupo-menu-guias-icon"))
            guias_tab.click()

            time.sleep(3)

            self.acessar_guia_sadt()

        except Exception as e:
            pass



    def acessar_guia_sadt(self):
        """Acessa o 'Guia de SP/SADT' na aba 'Guias'."""
        try:
            guia_sadt_button = self.acessar_com_reattempt((By.CSS_SELECTOR, ".menuLink.guia-spsadt-icon"))
            guia_sadt_button.click()

            time.sleep(3)

        except Exception as e:
            pass


    def process_row(self):
        """Processando uma única linha do Excel por vez."""
        # Verifica se a linha já foi processada
        guia_cod = self.df.at[self.row_index, 'GUIA_COD']
        if pd.notna(guia_cod) and str(guia_cod).strip() != '':
            guia_cod_str = str(guia_cod).strip()
            # Verifica se o conteúdo é numérico
            if guia_cod_str.replace('.', '', 1).isdigit():
                # Linha já processada, exibe mensagem de aviso e pula para a próxima
                logging.warning(f"Linha {self.row_index + 2} já foi executada e a guia solicitada é {guia_cod_str}.")
                return
            else:
                pass
        else:
            pass

        try:
            # Lida com o alerta caso ele apareça
            self.lidar_com_alerta()

            # Preenche o número da carteira apenas após o alerta ser fechado
            self.preencher_numero_carteira()

            # Preenche o tipo de atendimento e quantas guias serão solicitadas
            self.preencher_carater_atendimento()

           #self.data_solicitacao()

            # Preenche o campo 'Indicação Clínica'
            self.preencher_indicacao_clinica()

            # Abre a aba 'Procedimentos'
            self.acessar_procedimentos()

            # Clicando em inserir
            self.clicar_inserir_e_preencher()

            # Preenchendo campo dos profissionais
            self.preencher_campo_profissionais()

            # Abre a aba 'Observações/Justificativa'
            self.preencher_observacao_justificativa()

            # Anexa o documento
            self.selecionar_pedido_medico()

            # Clica no botão 'Escolher arquivo' e faz o upload
            self.Anexando_RM()

            # Selecionando opção de anexo
            self.selecionar_relatorio_clinico()

            # Anexo relatório clínico
            self.Anexando_RC()

            # Salvando e confirmando a solicitação
            self.salvar_confirmar()

            #self.salvar_anotar_numero()  # Captura e salva o número da guia

        except Exception as e:
            logging.error(f"Erro ao processar a linha {self.row_index + 2}: {e}")
            try:
                nome_paciente = self.df['PACIENTE'].iloc[self.row_index]
            except KeyError:
                nome_paciente = 'Nome do paciente não encontrado'
            
            try:
                especialidade = self.df['ESPECIALIDADE'].iloc[self.row_index]
            except KeyError:
                especialidade = 'Especialidade não encontrada'
            # Salva a mensagem de erro no arquivo txt
            current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            with open(self.txt_file_path, 'a', encoding='utf-8') as f:
                f.write(f"{current_time} - Paciente: {nome_paciente} - Especialidade: {especialidade}- Erro ao processar linha.\n")

            self.salvar_numero_no_excel(lista_erros=f"Erro ao processar a linha")

            time.sleep(5)

            logging.info("Passando para a próxima linha após erro.")
            pass 


    def lidar_com_alerta(self):
        """Lida com possíveis alertas que possam aparecer na página."""
        try:
            alert_present = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "noty_top_layout_container"))
            )

            if alert_present:
                fechar_button = self.driver.find_element(By.ID, "button-1")
                fechar_button.click()
                time.sleep(2)
        except Exception:
            pass



    def preencher_numero_carteira(self):
        """Preenche o campo 'Número da Carteira' através do popup e valida o nome do beneficiário."""
        max_attempts = 2
        for attempt in range(max_attempts):
            try:
                numero_carteira = self.get_excel_value('PADRAO')
                nome_esperado = self.get_excel_value('PACIENTE')

                # Clicar no elemento para abrir o popup
                elemento_abrir_popup = self.acessar_com_reattempt((By.XPATH, '//*[@id="div_numeroDaCarteira"]/div/i'))
                elemento_abrir_popup.click()
                logging.info("Popup para inserção do número da carteira aberto.")

                # Aguarda o popup aparecer
                WebDriverWait(self.driver, 15).until(
                    EC.visibility_of_element_located((By.XPATH, '//*[@id="cartao"]'))
                )

                # Preenche o campo no popup
                campo_cartao = self.acessar_com_reattempt((By.XPATH, '//*[@id="cartao"]'))
                campo_cartao.clear()
                campo_cartao.send_keys(numero_carteira)
                campo_cartao.send_keys(Keys.ENTER)

                # Verifica se o nome do beneficiário foi preenchido corretamente
                WebDriverWait(self.driver, 180).until(
                    lambda driver: driver.find_element(By.XPATH, '//*[@id="nomeDoBeneficiario"]').get_attribute('value') != ''
                )

                # Valida se o nome do beneficiário é o esperado
                nome_preenchido = self.driver.find_element(By.XPATH, '//*[@id="nomeDoBeneficiario"]').get_attribute('value')
                if nome_preenchido.strip().lower() != nome_esperado.strip().lower():
                    logging.error("O nome do beneficiário preenchido não corresponde ao esperado.")
                    if attempt < max_attempts - 1:
                        logging.info(f"Tentativa {attempt + 1} falhou. Tentando novamente.")
                        continue  # Tentar novamente
                    else:
                        raise Exception("Falha ao preencher o número da carteira após múltiplas tentativas.")
                else:
                    logging.info("Nome do beneficiário validado com sucesso.")
                    break  # Sai do loop, pois teve sucesso

            except TimeoutException:
                logging.error("Tempo limite excedido ao esperar pelo popup ou pelo preenchimento do nome do beneficiário.")
                if attempt < max_attempts - 1:
                    logging.info(f"Tentativa {attempt + 1} falhou. Tentando novamente.")
                    continue  # Tentar novamente
                else:
                    raise
            except Exception as e:
                logging.error(f"Erro ao preencher o campo 'Número da Carteira': {e}")
                if attempt < max_attempts - 1:
                    logging.info(f"Tentativa {attempt + 1} falhou. Tentando novamente.")
                    continue  # Tentar novamente
                else:
                    raise



    def preencher_carater_atendimento(self):
        try:

            carater_atendimento_input = self.acessar_com_reattempt((By.ID, "caraterAtendimento"))

            carater_atendimento_input.click()

            carater_atendimento_input.send_keys(Keys.ARROW_DOWN)
            carater_atendimento_input.send_keys(Keys.ENTER)

            time.sleep(1)

        except Exception:
            pass


    def data_solicitacao(self):
        """Preenche o campo 'Data de Solicitação' com dados do Excel e verifica se foi preenchido corretamente."""
        try:
            data_solicitacao = self.df['DATA'].iloc[self.row_index]
            if pd.isnull(data_solicitacao):
                logging.error(f"Data de solicitação está vazia na linha {self.row_index + 2}.")
                return

            # Verifica se data_solicitacao é um Timestamp e formata para 'DD/MM/YYYY'
            if isinstance(data_solicitacao, pd.Timestamp):
                data_solicitacao_str = data_solicitacao.strftime('%d/%m/%Y')
            else:
                #converte a string para datetime, caso não esteja no formato Timestamp DD/MM/YYYY
                try:
                    data_solicitacao_parsed = pd.to_datetime(data_solicitacao, dayfirst=True)
                    data_solicitacao_str = data_solicitacao_parsed.strftime('%d/%m/%Y')
                except Exception as e:
                    logging.error(f"Erro ao converter a data: {e}")
                    return

            data_solicitacao_input = self.acessar_com_reattempt((By.XPATH, '//*[@id="dataSolicitacao"]'))
            data_solicitacao_input.click()  # Clique no campo para garantir que está focado
            data_solicitacao_input.clear()  # Limpa o texto existente
            data_solicitacao_input.send_keys(data_solicitacao_str)  # Insere a data
            logging.info(f" A 'Data de Solicitação' foi preenchida com sucesso: {data_solicitacao_str}")

            time.sleep(1)

            # Agora verifica se o campo realmente contém a data inserida
            valor_preenchido = data_solicitacao_input.get_attribute('value')
            if valor_preenchido != data_solicitacao_str:
                return
            else:
                pass

        except Exception:
            pass


    def preencher_indicacao_clinica(self):
        """Preenche o campo 'Indicação Clínica' com dados do Excel."""
        try:
            indicacao_clinica = self.get_excel_value('INDICACAO_CLINICA')

            # Acessa o campo de entrada 'Indicação Clínica'
            indicacao_clinica_input = self.acessar_com_reattempt((By.ID, "indicacaoClinica"))
            indicacao_clinica_input.clear()
            indicacao_clinica_input.send_keys(indicacao_clinica)

            # Espera até que o valor digitado esteja presente no campo de entrada
            WebDriverWait(self.driver, 10).until(
                lambda driver: indicacao_clinica_input.get_attribute('value') == indicacao_clinica
            )
            time.sleep(.5)

        except Exception as e:
            pass




    def acessar_procedimentos(self):
        """Clica na aba 'Procedimentos' e garante que o botão 'confirmarEdicaoDeProcedimento' esteja visível."""
        try:
            # Encontra e clica na aba 'Procedimentos'
            procedimentos_tab = self.acessar_com_reattempt((By.ID, "ui-accordion-accordion-header-2"))
            self.driver.execute_script("arguments[0].scrollIntoView(true);", procedimentos_tab)
            procedimentos_tab.click()
            logging.info("Aba 'Procedimentos' clicada.")

            # Espera até que o botão 'Incluir Procedimento' esteja visível, indicando que a aba abriu
            WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.ID, "incluirProcedimento"))
            )
            

            # Verifica se o elemento 'confirmarEdicaoDeProcedimento' está presente e visível
            confirmar_button_locator = (By.XPATH, '//*[@id="confirmarEdicaoDeProcedimento"]')
            confirmar_button = self.driver.find_element(*confirmar_button_locator)

            # Se o elemento não estiver visível, rola um pouco para cima e verifica novamente
            if not confirmar_button.is_displayed():
                self.driver.execute_script("window.scrollBy(0, -200);")  
                time.sleep(1) 
                if not confirmar_button.is_displayed():
                    pass

        except Exception as e:
            pass



    def clicar_inserir_e_preencher(self):
        """Clica no botão 'Inserir Procedimento' e preenche o campo subsequente com dados do Excel."""
        try:

            inserir_button = self.acessar_com_reattempt((By.ID, "incluirProcedimento"))

            inserir_button.click()

            time.sleep(1)

            procedimento = self.get_excel_value('STATUS').zfill(8)

            procedimento_input = self.acessar_com_reattempt((By.XPATH, '//*[@id="registroProcedimentoCodigo"]/input'))

            procedimento_input.send_keys(procedimento)

            time.sleep(2)

            procedimento_input.send_keys(Keys.RETURN)

            total = self.get_excel_value('TOTAL')

            total_input = self.acessar_com_reattempt((By.XPATH, '//*[@id="registroProcedimentoQuantidade"]/input'))
            total_input.clear()
            total_input.send_keys(total)

            time.sleep(1)

            total_input.send_keys(Keys.RETURN)

            logging.info(f" A 'QUANTIDADE' de consultas solicitadas foram: {total}")

            confirmar_button = self.acessar_com_reattempt((By.XPATH, '//*[@id="confirmarEdicaoDeProcedimento"]'))

            confirmar_button.click()

            time.sleep(1)

        except Exception:
            pass



    def preencher_campo_profissionais(self):
        """Preenche os campos dos profissionais."""
        try:

            cbo = str(self.df['CBO'].iloc[self.row_index])[:6]

            profissionais_tab = self.acessar_com_reattempt((By.XPATH, '//*[@id="ui-accordion-accordion-header-3"]'))
            profissionais_tab.click()

            time.sleep(1)

            inserir_button = self.acessar_com_reattempt((By.XPATH, '//*[@id="incluirProfissional"]'))
            inserir_button.click()

            time.sleep(2)

            grau_partic_input = self.acessar_com_reattempt((By.XPATH, '//*[@id="registroProfissionalGrauParticipacao"]/input'))
            grau_partic_input.click()

            time.sleep(2)

            for _ in range(5):
                grau_partic_input.send_keys(Keys.ARROW_DOWN)
                time.sleep(0.2)

            grau_partic_input.send_keys(Keys.RETURN)
            time.sleep(1)

            cod_profissional = self.get_excel_value('COD_PROFISSIONAL')
            profissional_codigo_input = self.acessar_com_reattempt((By.XPATH, '//*[@id="registroProfissionalCodigo"]/input'))
            profissional_codigo_input.send_keys(cod_profissional)
            profissional_codigo_input.send_keys(Keys.RETURN)

            time.sleep(2)

            cbo_input = self.acessar_com_reattempt((By.XPATH, '//*[@id="registroProfissionalCodCBO"]/input'))
            cbo_input.click()

            time.sleep(3)
            if cbo == "223605":
                cbo_input.send_keys(Keys.ARROW_DOWN)
                cbo_input.send_keys(Keys.RETURN)
            elif cbo == "223810":
                for _ in range(2):
                    cbo_input.send_keys(Keys.ARROW_DOWN)
                cbo_input.send_keys(Keys.RETURN)
            elif cbo == "225125":
                for _ in range(3):
                    cbo_input.send_keys(Keys.ARROW_DOWN)
                cbo_input.send_keys(Keys.RETURN)
            elif cbo == "225170":
                for _ in range(4):
                    cbo_input.send_keys(Keys.ARROW_DOWN)
                cbo_input.send_keys(Keys.RETURN)
            elif cbo == "251510":
                for _ in range(5):
                    cbo_input.send_keys(Keys.ARROW_DOWN)
                cbo_input.send_keys(Keys.RETURN)
            elif cbo == "223905":
                for _ in range(6):
                    cbo_input.send_keys(Keys.ARROW_DOWN)
                cbo_input.send_keys(Keys.RETURN)
            else:
                logging.warning(f"Não foi possível preencher o CBO de atendimento: {cbo}")

            cbo_input.send_keys(Keys.ESCAPE)
            cbo_input.send_keys(Keys.RETURN)

            time.sleep(2)

            # Clicar no botão 'Confirmar' após preencher o CBO
            confirmar_button = self.acessar_com_reattempt((By.XPATH, '//*[@id="confirmarEdicaoDeProfissional"]'))
            confirmar_button.click()

        except Exception:
            pass



    def preencher_observacao_justificativa(self):
        """Preenche o campo 'Observação/Justificativa' com dados do Excel."""
        try:
            justificativa = self.get_excel_value('JUSTIFICATIVA')

            
            observacao_tab = self.acessar_com_reattempt((By.ID, "ui-accordion-accordion-header-4"))
            self.driver.execute_script("arguments[0].scrollIntoView(true);", observacao_tab)
            time.sleep(1)

            observacao_tab.click()

           
            observacao_input = self.acessar_com_reattempt((By.ID, "observacao"))

            
            observacao_input.send_keys(justificativa)

            logging.info("Campo 'Observação/Justificativa' preenchido com sucesso com sucesso")

        except Exception:
            pass



    def selecionar_pedido_medico(self):
        """Seleciona o tipo de anexo usando teclas de navegação."""
        try:

            tipo_anexo_dropdown = self.acessar_com_reattempt((By.ID, "tipoAnexoGuiaUpload"))
            tipo_anexo_dropdown.click()
            time.sleep(1)

            # Simula a tecla para baixo várias vezes até chegar na opção desejada
            for _ in range(46): 
                tipo_anexo_dropdown.send_keys(Keys.ARROW_DOWN)

            tipo_anexo_dropdown.send_keys(Keys.RETURN)

        except Exception:
            pass



    def Anexando_RM(self):
        try:
            base_path = Path(r"G:\Meu Drive\IPASGO\1.RELATORIO MEDICO E CLINICO")
            nome_paciente = self.get_excel_value('PACIENTE')
            id_paciente = self.get_excel_value('CARTEIRA')

            if not nome_paciente or not id_paciente:
                return
            
            logging.info(f"Paciente: {nome_paciente}, ID: {id_paciente}")

            # Constrói o nome da pasta do paciente
            patient_folder_name = f"{nome_paciente}-{id_paciente}"
            patient_folder_path = base_path / patient_folder_name

            logging.info(f"O caminho do arquivo é: {patient_folder_path}")

            if not patient_folder_path.is_dir():
                logging.error(f"A pasta do paciente '{nome_paciente}' não foi encontrada.")
                return

            # Chamada corrigida da função
            arquivos_rm, _, _, _ = encontrar_arquivos_paciente(patient_folder_path, id_paciente)

            if not arquivos_rm:
                logging.error("Nenhum arquivo RM correspondente foi encontrado para o paciente.")
                return


            for arquivo_para_upload in arquivos_rm:
                input_file = self.driver.find_element(By.CSS_SELECTOR, 'input[type="file"]')

                input_file.send_keys(str(arquivo_para_upload))
                logging.info(f"Arquivo '{arquivo_para_upload}' inserido com sucesso.")

                time.sleep(2)
                break

            self.safe_click((By.XPATH, '//*[@id="upload_form"]/div/input[2]'))
            time.sleep(1)

             # Verificação se o arquivo foi anexado com sucesso
            nome_arquivo = arquivo_para_upload.name  # Obtém o nome do arquivo
            xpath_arquivo_anexado = '//*[@id="arquivos-enviados"]/tbody/tr[1]/td[3]/strong'

            # Espera explícita até que o elemento apareça e o texto corresponda
            try:
                elemento_anexo = WebDriverWait(self.driver, 10).until(
                    EC.text_to_be_present_in_element(
                        (By.XPATH, xpath_arquivo_anexado),
                        nome_arquivo
                    )
                )
                logging.info("Arquivo de RC está presente na caixa de diálogo determinada, podendo prosseguir a atividade.")
            except TimeoutException:
                raise Exception(f"Arquivo '{nome_arquivo}' não foi anexado com sucesso.")

        except Exception:
            pass



    def selecionar_relatorio_clinico(self):
        """Seleciona o tipo de anexo usando teclas de navegação."""
        try:

            tipo_anexo_dropdown = self.acessar_com_reattempt((By.ID, "tipoAnexoGuiaUpload"))
            tipo_anexo_dropdown.click()
            time.sleep(1)

            for _ in range(10):  
                tipo_anexo_dropdown.send_keys(Keys.ARROW_DOWN)

            tipo_anexo_dropdown.send_keys(Keys.RETURN)

            time.sleep(1)
            
        except Exception:
            pass


    def Anexando_RC(self):
        try:
            base_path = Path(r"G:\Meu Drive\IPASGO\1.RELATORIO MEDICO E CLINICO")
            nome_paciente = self.get_excel_value('PACIENTE')
            id_paciente = self.get_excel_value('CARTEIRA')
            cbo = self.get_excel_value('CBO')
            cbo = str(int(float(cbo)))

            if not nome_paciente or not id_paciente or not cbo:
                return
            logging.info(f"Paciente: {nome_paciente}, ID: {id_paciente}, CBO: {cbo}")

            patient_folder_name = f"{nome_paciente}-{id_paciente}"
            patient_folder_path = base_path / patient_folder_name

            logging.info(f"O caminho do arquivo é: {patient_folder_path}")

            if not patient_folder_path.is_dir():
                logging.error(f"A pasta do paciente '{nome_paciente}' não foi encontrada.")
                return

            # Chamada corrigida da função
            _, arquivos_rc_fono, arquivos_rc_psi, arquivos_rc_to = encontrar_arquivos_paciente(patient_folder_path, id_paciente)


            
            if cbo == "251510":  
                arquivos_rc = arquivos_rc_psi
                logging.info("CBO indica PSICOLOGIA. Selecionando arquivos PSI.")
            elif cbo == "223810":  
                arquivos_rc = arquivos_rc_fono
                logging.info("CBO indica FONOAUDIOLOGIA. Selecionando arquivos FONO.")
            elif cbo == "223905":  
                arquivos_rc = arquivos_rc_to
                logging.info("CBO indica TERAPIA OCUPACIONAL. Selecionando arquivos TO.")
            else:
                logging.error(f"CBO '{cbo}' não corresponde a uma especialidade conhecida.")
                return

            if not arquivos_rc:
                logging.error(f"Nenhum arquivo RC correspondente foi encontrado para o paciente e CBO: {cbo}.")
                return

            logging.info(f"Arquivos RC encontrados: {arquivos_rc}")

            
            for arquivo_para_upload in arquivos_rc:
                
                input_file = self.driver.find_element(By.CSS_SELECTOR, 'input[type="file"]')
                logging.info("Elemento de upload encontrado.")

                
                input_file.send_keys(str(arquivo_para_upload))
                logging.info(f"Arquivo '{arquivo_para_upload}' inserido com sucesso.")

                time.sleep(2)

                break  

            
            self.safe_click((By.XPATH, '//*[@id="upload_form"]/div/input[2]'))
            time.sleep(1)

            # Verificação se o arquivo foi anexado com sucesso
            nome_arquivo = arquivo_para_upload.name  # Obtém o nome do arquivo
            xpath_arquivo_anexado = '//*[@id="arquivos-enviados"]/tbody/tr[2]/td[3]/strong'

            # Espera explícita até que o elemento apareça e o texto corresponda
            try:
                elemento_anexo = WebDriverWait(self.driver, 10).until(
                    EC.text_to_be_present_in_element(
                        (By.XPATH, xpath_arquivo_anexado),
                        nome_arquivo
                    )
                )
                logging.info("Arquivo de RC está presente na caixa de diálogo determinada, podendo prosseguir a atividade.")
            except TimeoutException:
                raise Exception(f"Arquivo '{nome_arquivo}' não foi anexado com sucesso.")

        except Exception:
            pass


 
    def salvar_confirmar(self):
        try:
            # Tenta clicar no botão "Salvar"
            salvar_button = self.driver.find_element(By.XPATH, '//*[@id="btnGravar"]')
            self.driver.execute_script("arguments[0].scrollIntoView(true);", salvar_button)
            time.sleep(1.5)

            salvar_button.click()

            time.sleep(1.5)  

            # Tenta clicar no botão "Confirmar"
            try:
                # Espera até que o botão "Confirmar" esteja presente
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/div[8]/div[3]/div/button[1]'))
                )
                confirmar_button = self.driver.find_element(By.XPATH, '/html/body/div[8]/div[3]/div/button[1]')
                self.driver.execute_script("arguments[0].scrollIntoView(true);", confirmar_button)
                time.sleep(1)
                confirmar_button.click()
                logging.info("Confirmado com sucesso")

                # Chama a função para capturar e salvar o número da guia
                self.salvar_anotar_numero()

            except Exception:
                logging.error(f"Não foi possível clicar no botão 'Confirmar'")
                # Chama a função para tratar os erros
                self.tratar_erros()

        except Exception:
            logging.error(f"Erro ao clicar no botão 'Salvar'")
            # Opcionalmente, você pode chamar a função tratar_erros aqui ou tomar outra ação
            


    def tratar_erros(self):
        # Função para lidar com a captura e armazenamento de erros
        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)

        errors_found = []
        for error_name, error_xpath in self.ERROR_ELEMENTS.items():
            try:
                elemento = self.driver.find_element(By.XPATH, error_xpath)
                #self.driver.execute_script("arguments[0].scrollIntoView(true);", elemento)
                #time.sleep(0.5)  # Aguarda para garantir que o elemento esteja visível
                if elemento.is_displayed():
                    erro_texto = elemento.text.strip()
                    errors_found.append(f"{erro_texto}")
                    logging.warning(f"Erro detectado - {error_name}: {erro_texto}")
            except NoSuchElementException:
                continue  # Se o elemento não for encontrado, passa para o próximo

        if errors_found:
            # Processa e salva os erros encontrados
            try:
                nome_paciente = self.df['PACIENTE'].iloc[self.row_index]
            except KeyError:
                nome_paciente = 'Nome do paciente não encontrado'

            erros_formatados = "; ".join(errors_found)
            current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            # Salva os erros no arquivo TXT
            with open(self.txt_file_path, 'a', encoding='utf-8') as f:
                f.write(f"{current_time} - Paciente: {nome_paciente} - Erros: {erros_formatados}\n")

            logging.info(f"Erros salvos no arquivo 'numeros_guias.txt': {erros_formatados}")

            # Salva os erros no Excel
            self.salvar_numero_no_excel(lista_erros=erros_formatados)

            time.sleep(2)

            # Recarrega a página para processar a próxima linha
            self.driver.refresh()
            logging.info("Página recarregada devido aos erros encontrados.")

        else:
            pass
            #logging.error("Nenhum erro encontrado, mas o botão 'Confirmar' não pôde ser clicado.")
            #############################################################
            #criar exceção nessa parte para diminuir incidencia de erros no processamento headless
            #############################################################
            # Opcionalmente, você pode tomar alguma ação aqui, como levantar uma exceção




    def salvar_anotar_numero(self): 
        max_attempts = 3  
        for attempt in range(max_attempts):
            try:
                logging.info(f"Iniciando o processo de captura do número da guia... Tentativa {attempt + 1} de {max_attempts}")

                self.driver.execute_script("window.scrollTo(0, 0);")

                WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/div[8]'))
                )
                time.sleep(2)
                WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="ui-id-47"]'))
                )
                time.sleep(2)
                WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="dialogText"]'))
                )
                time.sleep(2)
                elemento_numero_guia = WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="dialogText"]/div[2]'))
                )

                time.sleep(1)
                numero_guia_completo = elemento_numero_guia.text.strip()
                numero_guia = numero_guia_completo.split("Nº Guia Operadora:")[-1].strip()
                logging.info(f"Número da Guia capturado: {numero_guia}")

                nome_paciente = self.df['PACIENTE'].iloc[self.row_index]
                logging.info(f"Nome do paciente capturado: {nome_paciente}")

                nome_especialidade = self.df['ESPECIALIDADE'].iloc[self.row_index]
                logging.info(f"Nome da especialidade capturada: {nome_especialidade}")

                time.sleep(1)
                self.salvar_numero_no_excel(numero_guia=numero_guia)    
                # Grava o número da guia e o nome do paciente em um arquivo txt
                try:
                    current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    with open(self.txt_file_path, 'a', encoding='utf-8') as f:
                        f.write(f"{current_time} - Paciente: {nome_paciente}  {nome_especialidade} - Nº Guia Operadora: {numero_guia}\n")
                    logging.info(f"Número da guia e nome do paciente salvos no arquivo 'numeros_guias.txt'.")

                    
                    ActionChains(self.driver).send_keys(Keys.ESCAPE).perform() 

                    time.sleep(5)

                    break  

                except Exception as e:
                    logging.error(f"Erro ao salvar no arquivo txt: {e}", exc_info=True)

            except Exception as e:
                if attempt < max_attempts - 1:
                    time.sleep(2)
                else:
                    
                    try:
                        current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        nome_paciente = self.df['PACIENTE'].iloc[self.row_index]
                    except KeyError:
                        nome_paciente = 'Nome do paciente não encontrado'
                    
                    # Escreve a mensagem de erro no arquivo txt
                    with open(self.txt_file_path, 'a', encoding='utf-8') as f:
                        f.write(f"{current_time} - Paciente: {nome_paciente} - Erro ao processar a linha\n")
                    
                    
                    self.salvar_numero_no_excel(lista_erros=f"Erro ao processar a linha")
                    raise e




    def salvar_numero_no_excel(self, numero_guia=None, lista_erros=None):
        try:
            col_name = 'GUIA_COD'  
            erro_col_name = 'ERRO'  

            
            if col_name not in self.df.columns:
                self.df[col_name] = ''
            if erro_col_name not in self.df.columns:
                self.df[erro_col_name] = ''

            # Atualiza o DataFrame na linha atual
            if lista_erros:
                self.df.at[self.row_index, col_name] = 'ERRO'
                self.df.at[self.row_index, erro_col_name] = lista_erros
                logging.info(f"Erro salvo na coluna '{erro_col_name}': {lista_erros}")
            else:
                if numero_guia:
                    self.df.at[self.row_index, col_name] = numero_guia
                    logging.info(f"Número da Guia '{numero_guia}' salvo com sucesso na coluna '{col_name}'.")
                else:           
                    self.df.at[self.row_index, col_name] = 'SUCESSO mas sem numero'
                    logging.info(f"não foi possível a captura do numero, mas foi bem sucedido. {col_name}'.")
                self.df.at[self.row_index, erro_col_name] = ''

            # Salva o DataFrame de volta no arquivo Excel
            self.df.to_excel(self.file_path, sheet_name=self.sheet_name, index=False)

        except Exception as e:
            logging.error(f"Erro ao salvar no arquivo Excel: {e}", exc_info=True)
            logging.warning("Execução correta, mas falha ao preencher o arquivo Excel.")




if __name__ == "__main__":
    try:
        
        root = tk.Tk()
        root.title("Automação de Solicitações IPASGO")

        # Define variables to store user inputs
        login_var = tk.StringVar()
        password_var = tk.StringVar()
        start_row_var = tk.StringVar()
        end_row_var = tk.StringVar()

        # Create labels and entry fields
        tk.Label(root, text="Login IPASGO").grid(row=0, column=0, padx=10, pady=5)
        login_entry = tk.Entry(root, textvariable=login_var)
        login_entry.grid(row=0, column=1, padx=10, pady=5)

        tk.Label(root, text="Senha IPASGO").grid(row=1, column=0, padx=10, pady=5)
        password_entry = tk.Entry(root, textvariable=password_var, show="*")
        password_entry.grid(row=1, column=1, padx=10, pady=5)

        tk.Label(root, text="Linha inicial do Excel").grid(row=2, column=0, padx=10, pady=5)
        start_row_entry = tk.Entry(root, textvariable=start_row_var)
        start_row_entry.grid(row=2, column=1, padx=10, pady=5)

        tk.Label(root, text="Linha final do Excel").grid(row=3, column=0, padx=10, pady=5)
        end_row_entry = tk.Entry(root, textvariable=end_row_var)
        end_row_entry.grid(row=3, column=1, padx=10, pady=5)

        #  Função para iniciar o processo de automação
        def iniciar_processo():
            login = login_var.get()
            password = password_var.get()
            start_row = start_row_var.get()
            end_row = end_row_var.get()

            if not login or not password or not start_row or not end_row:
                messagebox.showerror("Erro", "Por favor, preencha todos os campos.")
                return

            # Converter start_row e end_row em números inteiros
            try:
                start_row = int(start_row)
                end_row = int(end_row)
            except ValueError:
                messagebox.showerror("Erro", "As linhas inicial e final devem ser números inteiros.")
                return

            
            start_row_index = start_row - 2
            end_row_index = end_row - 2

            # Finaliza a janela Tkinte
            root.destroy()

            # Inicialize a classe de automação com entradas do usuário
            ipasgo = IpasgoAutomation(login, password, start_row_index, end_row_index)
            ipasgo.acessar_portal_ipasgo()

            # Valida o intervalo de linhas definido
            max_row = len(ipasgo.df) - 1
            if ipasgo.start_row < 0:
                print("A linha inicial não pode ser negativa.")
                return
            if ipasgo.end_row > max_row:
                print(f"A linha final não pode ser maior que {max_row + 2}.")
                return
            if ipasgo.start_row > ipasgo.end_row:
                print("A linha inicial não pode ser maior que a linha final.")
                return

        
            for idx in range(start_row_index, end_row_index + 1):
                ipasgo.row_index = idx
                logging.info(f"Iniciando o processamento da linha {idx + 2}")
                ipasgo.process_row()

            input("Pressione qualquer tecla para fechar o navegador...")
            ipasgo.close()

        
        iniciar_button = tk.Button(root, text="Iniciar", command=iniciar_processo)
        iniciar_button.grid(row=4, column=0, columnspan=2, pady=10)

                # Adiciona o filtro ao logger principal

        root.mainloop()

    except Exception as e:
        logging.error(f"Erro crítico: {e}")
      