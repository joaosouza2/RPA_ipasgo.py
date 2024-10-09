import logging
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException
import time
from selenium.webdriver.common.keys import Keys
from openpyxl import load_workbook
from selenium.webdriver.common.action_chains import ActionChains
import os
import re
from pathlib import Path
import shutil
import datetime
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def encontrar_arquivos_paciente(caminho_pasta, id_paciente, nome_paciente):
    p = Path(caminho_pasta)

    arquivos_rm = []
    arquivos_rc_fono = []
    arquivos_rc_psi = []
    arquivos_rc_to = []

    # Padrão para arquivos RM usando expressões regulares
    padrao_rm = rf'(?i)^RM.*{id_paciente}.*\..+$'

    # Padrões para os três tipos de documentação usando expressões regulares
    padrao_rc_fono = rf'(?i)^RC.*{id_paciente}.*FONO.*\..+$'
    padrao_rc_psi = rf'(?i)^RC.*{id_paciente}.*PSI.*\..+$'
    padrao_rc_to = rf'(?i)^RC.*{id_paciente}.*TO.*\..+$'

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
                time.sleep(1)
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


    def __init__(self):
        super().__init__()
        self.file_path = r"C:\Users\SUPERVISÃO ADM\.git\RPA_ipasgo.py\SOLICITACOES_AUTORIZACAO_FACPLAN.xlsx"
        self.copy_file_path = r"C:\Users\SUPERVISÃO ADM\.git\RPA_ipasgo.py\SOLICITACOES_AUTORIZACAO_FACPLAN_COPIA.csv"
        self.sheet_name = 'AUTORIZACOES'

        # Ler o arquivo Excel original
        self.df = pd.read_excel(
            self.file_path,
            sheet_name=self.sheet_name,
            header=0
        )

        self.df.columns = [col.upper().strip() for col in self.df.columns]
        if not os.path.exists(self.copy_file_path):
            self.df.to_csv(self.copy_file_path, index=False, encoding='utf-8')


        self.df = pd.read_csv(
            self.copy_file_path,
            header=0,    #linha do cabeçalho
            encoding='utf-8'
        )

        
        self.df.columns = [col.upper().strip() for col in self.df.columns]# Normaliza os nomes das colunas para maiúsculas
        
        # Adiciona a coluna 'ERRO' se não existir
        if 'ERRO' not in self.df.columns:
            self.df['ERRO'] = ''


        # Definição do intervalo de linhas
        self.start_row = 0
        self.end_row = len(self.df) - 1
        self.row_index = self.start_row
        
        # Definição dos elementos de erro


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
            matricula_input.send_keys("14898500")

            senha_input = self.acessar_com_reattempt((By.ID, "SilkUIFramework_wt13_block_wtPassword_wtPasswordInput"))
            senha_input.send_keys("Clmf2024")

            self.safe_click((By.ID, "SilkUIFramework_wt13_block_wtAction_wtLoginButton"))   
       
            self.wait_for_stability(timeout=10)

            link_portal_webplan = self.acessar_com_reattempt((By.XPATH, "//*[@id='IpasgoTheme_wt16_block_wtMainContent_wtSistemas_ctl10_SilkUIFramework_wt36_block_wtActions_wtModulos_SilkUIFramework_wt9_block_wtContent_wtModuloPortalTable_ctl04_wt2']"))
            self.driver.execute_script("arguments[0].scrollIntoView(true);", link_portal_webplan)
            time.sleep(2)
            link_portal_webplan.click()

            WebDriverWait(self.driver, 20).until(EC.number_of_windows_to_be(2))
            self.driver.switch_to.window(self.driver.window_handles[1])

            self.acessar_com_reattempt((By.ID, "menuPrincipal"))

            time.sleep(4)

            self.acessar_guias()

        except Exception as e:
            logging.error(f"Erro ao acessar o site ou preencher o formulário: {e}")
            return



    def acessar_guias(self):
        """Acessa a aba 'Guias' no menu principal."""
        try:
            guias_tab = self.acessar_com_reattempt((By.CSS_SELECTOR, ".menuLink.grupo-menu-guias-icon"))
            guias_tab.click()

            time.sleep(3)

            self.acessar_guia_sadt()

        except Exception as e:
            logging.error(f"Erro ao clicar na aba 'Guias': {e}")
            return



    def acessar_guia_sadt(self):
        """Acessa o 'Guia de SP/SADT' na aba 'Guias'."""
        try:
            guia_sadt_button = self.acessar_com_reattempt((By.CSS_SELECTOR, ".menuLink.guia-spsadt-icon"))
            guia_sadt_button.click()

            time.sleep(3)

        except Exception as e:
            logging.error(f"Erro ao processar o guia SP/SADT: {e}")
            return


    def process_row(self):
        """Processando uma única linha do Excel por vez."""
        try:
            # Lida com o alerta caso ele apareça
            self.lidar_com_alerta()

            # Preenche o número da carteira apenas após o alerta ser fechado
            self.preencher_numero_carteira()

            # Preenche o tipo de atendimento e quantas guias serão solicitadas
            self.preencher_carater_atendimento()

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

            # Armazenando o número em uma lista para print em txt ou excel
            lista_numeros = []  # Inicializa a lista para armazenar os números
            self.salvar_anotar_numero(lista_numeros)  # Salva na lista

        except Exception as e:
            logging.error(f"Erro ao processar a linha {self.row_index}: {e}")
            try:
                nome_paciente = self.df['PACIENTE'].iloc[self.row_index]
            except KeyError:
                nome_paciente = 'Nome do paciente não encontrado'

            # Salva a mensagem de erro no arquivo txt
            current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            with open('numeros_guias.txt', 'a', encoding='utf-8') as f:
                f.write(f"{current_time} - Paciente: {nome_paciente} - Erro ao processar a linha: {e}\n")

            # Salva uma entrada de erro no CSV
            self.salvar_numero_no_csv(lista_erros=f"Erro: {e}")

            # Aguarda 5 segundos
            time.sleep(5)

            # Passa para a próxima linha sem recarregar novamente
            logging.info("Passando para a próxima linha após erro.")
            pass  # Continua para a próxima iteração


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
        except Exception as e:
            logging.error(f"Erro ao lidar com o alerta: {e}")



    def preencher_numero_carteira(self):
        """Preenche o campo 'Número da Carteira' com dados do Excel e valida o nome do beneficiário."""
        try:    
            numero_carteira = self.get_excel_value('CARTEIRA')
            nome_esperado = self.get_excel_value('PACIENTE')

            numero_carteira_input = self.acessar_com_reattempt((By.ID, "numeroDaCarteira"))
            numero_carteira_input.send_keys(numero_carteira)
            logging.info(f"Campo 'Número da Carteira' preenchido com sucesso com o valor: {numero_carteira}")
            time.sleep(2)
            numero_carteira_input.send_keys(Keys.ARROW_DOWN)
            numero_carteira_input.send_keys(Keys.ENTER)
            

            # Aguarda até que o campo 'nomeDoBeneficiario' seja preenchido
            WebDriverWait(self.driver, 180).until(
                lambda driver: driver.find_element(By.XPATH, '//*[@id="nomeDoBeneficiario"]').get_attribute('value') != ''
            )

            # Valida se o nome do beneficiário é o esperado
            nome_preenchido = self.driver.find_element(By.XPATH, '//*[@id="nomeDoBeneficiario"]').get_attribute('value')
            if nome_preenchido.strip().lower() != nome_esperado.strip().lower():
                logging.error("O nome do beneficiário preenchido não corresponde ao esperado.")
                return


        except TimeoutException:
            logging.error("Tempo limite excedido ao esperar pela conclusão da etapa do token.")
        except Exception as e:
            logging.error(f"Erro ao preencher o campo 'Número da Carteira': {e}")



    def preencher_carater_atendimento(self):
        try:

            carater_atendimento_input = self.acessar_com_reattempt((By.ID, "caraterAtendimento"))

            carater_atendimento_input.click()

            carater_atendimento_input.send_keys(Keys.ARROW_DOWN)
            carater_atendimento_input.send_keys(Keys.ENTER)

            time.sleep(3)

        except Exception as e:
            logging.error(f"Erro ao preencher o campo 'Caráter do Atendimento': {e}")



    def preencher_indicacao_clinica(self):
        """Preenche o campo 'Indicação Clínica' com dados do Excel."""
        try:

            indicacao_clinica = self.get_excel_value('INDICACAO_CLINICA')

            indicacao_clinica_input = self.acessar_com_reattempt((By.ID, "indicacaoClinica"))
            indicacao_clinica_input.send_keys(indicacao_clinica)

            logging.info(f"Campo 'Indicação Clínica' preenchido com sucesso com o valor: {indicacao_clinica}")

        except Exception as e:
            logging.error(f"Erro ao preencher o campo 'Indicação Clínica': {e}")



    def acessar_procedimentos(self):
        """Clica no campo 'Procedimentos' para abrir a aba, garantindo que ele esteja visível."""
        try:

            procedimentos_tab = self.acessar_com_reattempt((By.ID, "ui-accordion-accordion-header-2"))

            self.driver.execute_script("arguments[0].scrollIntoView(true);", procedimentos_tab)

            time.sleep(1)

            procedimentos_tab.click()

            time.sleep(2)

        except Exception as e:
            logging.error(f"Erro ao abrir a aba 'Procedimentos': {e}")



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

            logging.info(f"Campo 'QUANTIDADE' preenchido com sucesso com o valor: {total}")

            confirmar_button = self.acessar_com_reattempt((By.XPATH, '//*[@id="confirmarEdicaoDeProcedimento"]'))

            confirmar_button.click()

            time.sleep(1)

        except Exception as e:
            logging.error(f"Erro ao preencher o procedimento: {e}")



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
                logging.warning(f"Nenhuma ação correspondente para o valor CBO: {cbo}")

            cbo_input.send_keys(Keys.ESCAPE)
            cbo_input.send_keys(Keys.RETURN)

            time.sleep(2)

            # Clicar no botão 'Confirmar' após preencher o CBO
            confirmar_button = self.acessar_com_reattempt((By.XPATH, '//*[@id="confirmarEdicaoDeProfissional"]'))
            confirmar_button.click()

        except Exception as e:
            logging.error(f"Erro ao preencher os campos dos profissionais: {e}")



    def preencher_observacao_justificativa(self):
        """Preenche o campo 'Observação/Justificativa' com dados do Excel."""
        try:
            justificativa = self.get_excel_value('JUSTIFICATIVA')

            # Rolagem para garantir que o campo esteja visível na tela
            observacao_tab = self.acessar_com_reattempt((By.ID, "ui-accordion-accordion-header-4"))
            self.driver.execute_script("arguments[0].scrollIntoView(true);", observacao_tab)
            time.sleep(1)

            observacao_tab.click()

            # Localiza o campo de observação
            observacao_input = self.acessar_com_reattempt((By.ID, "observacao"))

            # Preenche o campo com o valor da justificativa
            observacao_input.send_keys(justificativa)

            logging.info(f"Campo 'Observação/Justificativa' preenchido com sucesso com o valor: {justificativa}")

        except Exception as e:
            logging.error(f"Erro ao preencher o campo 'Observação/Justificativa': {e}")



    def selecionar_pedido_medico(self):
        """Seleciona o tipo de anexo usando teclas de navegação."""
        try:

            tipo_anexo_dropdown = self.acessar_com_reattempt((By.ID, "tipoAnexoGuiaUpload"))
            tipo_anexo_dropdown.click()
            time.sleep(1)

            # Simula a tecla para baixo várias vezes até chegar na opção desejada
            for _ in range(46):  # Ajuste o número de vezes conforme necessário
                tipo_anexo_dropdown.send_keys(Keys.ARROW_DOWN)

            tipo_anexo_dropdown.send_keys(Keys.RETURN)

        except Exception as e:
            logging.error(f"Erro ao selecionar o tipo de anexo: {e}")



    def Anexando_RM(self):
        try:

            # Define o caminho base
            base_path = Path(r"G:\Meu Drive\IPASGO\1.RELATORIO MEDICO E CLINICO")

            # Obtém 'Paciente' e 'CARTEIRA' da planilha Excel
            nome_paciente = self.get_excel_value('PACIENTE')
            id_paciente = self.get_excel_value('CARTEIRA')

            if not nome_paciente or not id_paciente:
                logging.error("Falha ao obter o nome ou o ID do paciente da planilha.")
                return
            logging.info(f"Paciente: {nome_paciente}, ID: {id_paciente}")

            # Constrói o nome da pasta do paciente
            patient_folder_name = f"{nome_paciente}-{id_paciente}"

            # Constrói o caminho completo para a pasta do paciente
            patient_folder_path = base_path / patient_folder_name

            logging.info(f"Caminho da pasta do paciente: {patient_folder_path}")

            # Verifica se a pasta do paciente existe
            if not patient_folder_path.is_dir():
                logging.error(f"A pasta do paciente '{patient_folder_path}' não foi encontrada.")
                return

            # Encontra os arquivos do paciente, focando apenas em RM
            arquivos_rm, _, _, _ = encontrar_arquivos_paciente(patient_folder_path, id_paciente, nome_paciente)

            if not arquivos_rm:
                logging.error("Nenhum arquivo RM correspondente foi encontrado para o paciente.")
                return

            logging.info(f"Arquivos RM encontrados: {arquivos_rm}")

            # Faz o upload do primeiro arquivo RM encontrado
            for arquivo_para_upload in arquivos_rm:
                # Localiza o elemento <input type="file">
                input_file = self.driver.find_element(By.CSS_SELECTOR, 'input[type="file"]')
                logging.info("Elemento de upload encontrado.")

                # Envia o caminho do arquivo para o elemento 'input'
                input_file.send_keys(str(arquivo_para_upload))
                logging.info(f"Arquivo '{arquivo_para_upload}' selecionado com sucesso.")

                time.sleep(2)

                break  # Encerra o loop após carregar o primeiro arquivo RM

            # Clica no botão 'Adicionar'
            self.safe_click((By.XPATH, '//*[@id="upload_form"]/div/input[2]'))

            time.sleep(2)

        except Exception as e:
            logging.error(f"Erro ao fazer upload dos arquivos RM do paciente: {e}")



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
            
        except Exception as e:
            logging.error(f"Erro ao selecionar o tipo de anexo: {e}")


    def Anexando_RC(self):
        try:
            
            base_path = Path(r"G:\Meu Drive\IPASGO\1.RELATORIO MEDICO E CLINICO")

            
            nome_paciente = self.get_excel_value('PACIENTE')
            id_paciente = self.get_excel_value('CARTEIRA')
            
            
            cbo = self.get_excel_value('CBO')
            cbo = str(int(float(cbo)))  

            if not nome_paciente or not id_paciente or not cbo:
                logging.error("Falha ao obter o nome, ID do paciente ou CBO da planilha.")
                return
            logging.info(f"Paciente: {nome_paciente}, ID: {id_paciente}, CBO: {cbo}")

            
            patient_folder_name = f"{nome_paciente}-{id_paciente}"

            
            patient_folder_path = base_path / patient_folder_name

            logging.info(f"Caminho da pasta do paciente: {patient_folder_path}")

            
            if not patient_folder_path.is_dir():
                logging.error(f"A pasta do paciente '{patient_folder_path}' não foi encontrada.")
                return

            
            _, arquivos_rc_fono, arquivos_rc_psi, arquivos_rc_to = encontrar_arquivos_paciente(patient_folder_path, id_paciente, nome_paciente)

            
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
                logging.info(f"Arquivo '{arquivo_para_upload}' selecionado com sucesso.")

                time.sleep(2)

                break  

            
            self.safe_click((By.XPATH, '//*[@id="upload_form"]/div/input[2]'))
            time.sleep(1)

        except Exception as e:
            logging.error(f"Erro ao fazer upload dos arquivos RC do paciente: {e}")



    def salvar_confirmar(self):
        try:
            salvar_button = self.driver.find_element(By.XPATH, '//*[@id="btnGravar"]')
            self.driver.execute_script("arguments[0].scrollIntoView(true);", salvar_button)
            time.sleep(1)
            salvar_button.click()

            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/div[8]/div[3]/div/button[1]'))
            )
            logging.info("Modal de confirmação detectado.")

            confirmar_button = self.driver.find_element(By.XPATH, '/html/body/div[8]/div[3]/div/button[1]')
            self.driver.execute_script("arguments[0].scrollIntoView(true);", confirmar_button)
            time.sleep(1)
            confirmar_button.click()

            # Definir os elementos de erro
            errors_found = []
            for error_name, error_xpath in self.ERROR_ELEMENTS.items():
                try:
                    elemento = self.driver.find_element(By.XPATH, error_xpath)
                    if elemento.is_displayed():
                        erro_texto = elemento.text.strip()
                        errors_found.append(f"{error_name}: {erro_texto}")
                        logging.warning(f"Erro detectado - {error_name}: {erro_texto}")
                except NoSuchElementException:
                    continue

            if errors_found:
                # Obtem o nome do paciente
                try:
                    nome_paciente = self.df['PACIENTE'].iloc[self.row_index]
                except KeyError:
                    nome_paciente = 'Nome do paciente não encontrado'

                # Formatar a mensagem de erro
                erros_formatados = "; ".join(errors_found)

                # Salvar no arquivo TXT
                current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                with open('numeros_guias.txt', 'a', encoding='utf-8') as f:
                    f.write(f"{current_time} - Paciente: {nome_paciente} - Erros: {erros_formatados}\n")

                logging.info(f"Erros salvos no arquivo 'numeros_guias.txt': {erros_formatados}")

                self.salvar_numero_no_csv(lista_erros=erros_formatados)# Salvar no CSV
            
                time.sleep(5)

                self.driver.refresh()# Recarregar a página
                logging.info("Página recarregada devido aos erros encontrados.")

                return  # Sai da função para que o 'process_row' possa continuar gerenciando a troca de linha ao final do código

        except Exception as e:
            logging.error(f"Erro ao tentar salvar e confirmar: {e}")
            # Salvar a mensagem de erro no TXT e CSV
            try:
                nome_paciente = self.df['PACIENTE'].iloc[self.row_index]
            except KeyError:
                nome_paciente = 'Nome do paciente não encontrado'

            current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            erro_texto = f"Erro ao salvar e confirmar: {e}"
            with open('numeros_guias.txt', 'a', encoding='utf-8') as f:
                f.write(f"{current_time} - Paciente: {nome_paciente} - {erro_texto}\n")

            # Salvar no CSV com o erro
            self.salvar_numero_no_csv(lista_erros=erro_texto)

            # Aguardar 5 segundos
            time.sleep(5)

            # Recarregar a página
            self.driver.refresh()
            logging.info("Página recarregada devido a erro inesperado.")

            # Levanta a exceção para ser tratada no process_row
            raise e




    def salvar_anotar_numero(self, lista_numeros): 
        max_attempts = 3  # Número máximo de tentativas
        for attempt in range(max_attempts):
            try:
                logging.info(f"Iniciando o processo de captura do número da guia... Tentativa {attempt + 1} de {max_attempts}")

                self.driver.execute_script("window.scrollTo(0, 0);")

                WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/div[8]'))
                )
                
                WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="ui-id-47"]'))
                )
                
                WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="dialogText"]'))
                )
                
                elemento_numero_guia = WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="dialogText"]/div[2]'))
                )

                # Extrai o número da guia
                numero_guia_completo = elemento_numero_guia.text.strip()
                numero_guia = numero_guia_completo.split("Nº Guia Operadora:")[-1].strip()
                logging.info(f"Número da Guia capturado: {numero_guia}")

                nome_paciente = self.df['PACIENTE'].iloc[self.row_index]
                logging.info(f"Nome do paciente capturado: {nome_paciente}")

                nome_especialidade = self.df['ESPECIALIDADE'].iloc[self.row_index]
                logging.info(f"Nome da especialidade capturada: {nome_especialidade}")

                lista_numeros.append(numero_guia)
                time.sleep(1)
                self.salvar_numero_no_csv(lista_numeros)    
                # Grava a lista de números e o nome do paciente em um arquivo txt
                try:
                    current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    with open('numeros_guias.txt', 'a', encoding='utf-8') as f:
                        f.write(f"Paciente: {current_time} - {nome_paciente}  {nome_especialidade} - Nº Guia Operadora: {numero_guia}\n")
                    logging.info(f"Números das guias e nomes dos pacientes salvos no arquivo 'numeros_guias.txt'.")

                    # Envia a tecla ESC para fechar o pop-up
                    ActionChains(self.driver).send_keys(Keys.ESCAPE).perform() 

                    time.sleep(5)

                    
                    break

                except Exception as e:
                    logging.error(f"Erro ao salvar no arquivo txt: {e}", exc_info=True)
                    logging.warning("Falha ao preencher o arquivo txt. Tentando novamente...")

            except Exception as e:
        
                if attempt < max_attempts - 1:
                    time.sleep(2)
                else:
                    logging.error("Número máximo de tentativas alcançado. Não foi possível capturar o número da guia.")
                    # Salva a mensagem de erro no txt e no CSV
                    try:
                        
                        nome_paciente = self.df['PACIENTE'].iloc[self.row_index]
                    except KeyError:
                        nome_paciente = 'Nome do paciente não encontrado'
                    
                    # Escreve a mensagem de erro no arquivo txt
                    with open('numeros_guias.txt', 'a', encoding='utf-8') as f:
                        f.write(f"Paciente: {nome_paciente} - Erro ao capturar o número da guia: {e}\n")
                    
                    lista_numeros.append('')
                    

                    self.salvar_numero_no_csv(lista_numeros)
                    raise e




    def salvar_numero_no_csv(self, lista_erros=None):
        try:
            col_name = 'GUIA_COD'  # Este é o nome exato da coluna
            erro_col_name = 'ERRO'  # Nova coluna para erros


            if col_name not in self.df.columns: # Verifica se a coluna 'GUIA_COD' existe; se não, cria
                self.df[col_name] = ''
            if erro_col_name not in self.df.columns: # Verifica se a coluna 'ERRO' existe; se não, cria
                self.df[erro_col_name] = ''

            numero_guia = self.df.at[self.row_index, col_name] #lê o valor N* guia antes de inserir

            # Atualiza o DataFrame na linha atual
            if lista_erros:
                self.df.at[self.row_index, col_name] = 'ERRO'
                self.df.at[self.row_index, erro_col_name] = lista_erros
                logging.info(f"Erro salvo no CSV na coluna '{erro_col_name}': {lista_erros}")
            else:
                numero_guia = 'SUCESSO' # Caso contrário, salva o número da guia normalmente, ou o valor correspondente
                self.df.at[self.row_index, col_name] = numero_guia
                self.df.at[self.row_index, erro_col_name] = ''
                logging.info(f"Número da Guia '{numero_guia}' salvo com sucesso no CSV na coluna '{col_name}'.")

            self.df.to_csv(self.copy_file_path, index=False, encoding='utf-8') # Salva o DataFrame de volta no arquivo CSV

        except Exception as e:
            logging.error(f"Erro ao salvar no CSV: {e}", exc_info=True)
            logging.warning("Execução correta, mas falha ao preencher o CSV.")

if __name__ == "__main__":
    try:
        # Create the GUI window
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

        # Function to start the automation process
        def iniciar_processo():
            login = login_var.get()
            password = password_var.get()
            start_row = start_row_var.get()
            end_row = end_row_var.get()

            if not login or not password or not start_row or not end_row:
                messagebox.showerror("Erro", "Por favor, preencha todos os campos.")
                return

            # Convert start_row and end_row to integers
            try:
                start_row = int(start_row)
                end_row = int(end_row)
            except ValueError:
                messagebox.showerror("Erro", "As linhas inicial e final devem ser números inteiros.")
                return

            # Close the GUI window
            root.destroy()

            # Initialize the automation class with user inputs
            ipasgo = IpasgoAutomation()
            ipasgo.acessar_portal_ipasgo()

            # Validate the range of rows defined
            max_row = len(ipasgo.df) - 1
            if ipasgo.start_row < 0:
                print("A linha inicial não pode ser negativa.")
                return
            if ipasgo.end_row > max_row:
                print(f"A linha final não pode ser maior que {max_row}.")
                return
            if ipasgo.start_row > ipasgo.end_row:
                print("A linha inicial não pode ser maior que a linha final.")
                return

            # Process each row in the specified range
            for idx in range(ipasgo.start_row, ipasgo.end_row + 1):
                ipasgo.row_index = idx
                logging.info(f"Iniciando o processamento da linha {idx}")
                ipasgo.process_row()

            input("Pressione qualquer tecla para fechar o navegador...")
            ipasgo.close()

        # Create the "Iniciar" button
        iniciar_button = tk.Button(root, text="Iniciar", command=iniciar_processo)
        iniciar_button.grid(row=4, column=0, columnspan=2, pady=10)

        # Start the GUI loop
        root.mainloop()

    except Exception as e:
        logging.error(f"Erro crítico: {e}")
       