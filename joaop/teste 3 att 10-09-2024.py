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
import xlwings as xw



# Configuração do logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Definição da função encontrar_arquivos_paciente
def encontrar_arquivos_paciente(caminho_pasta, id_paciente, nome_paciente):
    from pathlib import Path
    import re

    p = Path(caminho_pasta)

    # Variáveis para armazenar os arquivos RM e os diferentes tipos de RC
    arquivos_rm = []
    arquivos_rc_fono = []
    arquivos_rc_psi = []
    arquivos_rc_to = []

    # Padrão para arquivos RM (case-insensitive no início)
    padrao_rm = rf'(?i)^RM.*{id_paciente}.*\..+$'

    # Padrões para os três tipos de RC (simplificado para focar em FONO, PSI e TO)
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

    # Retorna os arquivos RM e as classificações dos arquivos RC
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
        logging.info("Página estabilizada.")

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
    
    def __init__(self):
        super().__init__()
        self.file_path = r"C:\Users\AUDITORIA_CLMF\Desktop\ipasgo.py\joaop\BASE_AUTORIZACAO - NOVO 08-08-2024.xlsx"
        self.sheet_name = 'AUTORIZACOES'
        self.df = pd.read_excel(self.file_path, sheet_name=self.sheet_name)
        self.row_index = 32 # Torne este valor dinâmico se necessário

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
            logging.info("Abrindo o site do IPASGO...")
            self.driver.get("https://portalos.ipasgo.go.gov.br/Portal_Dominio/PrestadorLogin.aspx")
            logging.info("Esperando a página carregar...")
            self.wait_for_stability(timeout=10)

            matricula_input = self.acessar_com_reattempt((By.ID, "SilkUIFramework_wt13_block_wtUsername_wtUserNameInput2"))
            matricula_input.send_keys("14898500")

            senha_input = self.acessar_com_reattempt((By.ID, "SilkUIFramework_wt13_block_wtPassword_wtPasswordInput"))
            senha_input.send_keys("Clmf2024")

            self.safe_click((By.ID, "SilkUIFramework_wt13_block_wtAction_wtLoginButton"))

            logging.info("Aguardando a nova página carregar após o login...")
            self.wait_for_stability(timeout=10)

            logging.info("Procurando pelo link Portal WebPlan...")
            link_portal_webplan = self.acessar_com_reattempt((By.XPATH, "//*[@id='IpasgoTheme_wt16_block_wtMainContent_wtSistemas_ctl10_SilkUIFramework_wt36_block_wtActions_wtModulos_SilkUIFramework_wt9_block_wtContent_wtModuloPortalTable_ctl04_wt2']"))
            self.driver.execute_script("arguments[0].scrollIntoView(true);", link_portal_webplan)
            time.sleep(2)
            link_portal_webplan.click()

            # Aguarda até que a nova página seja carregada
            WebDriverWait(self.driver, 20).until(EC.number_of_windows_to_be(2))
            self.driver.switch_to.window(self.driver.window_handles[1])

            # Espera pela presença do menu principal
            self.acessar_com_reattempt((By.ID, "menuPrincipal"))

            # Adiciona um tempo de espera para garantir que a página esteja completamente carregada
            time.sleep(4)

            self.acessar_guias()

        except Exception as e:
            logging.error(f"Erro ao acessar o site ou preencher o formulário: {e}")
            return

    def acessar_guias(self):
        """Acessa a aba 'Guias' no menu principal."""
        try:
            logging.info("Procurando e clicando na aba 'Guias'...")
            guias_tab = self.acessar_com_reattempt((By.CSS_SELECTOR, ".menuLink.grupo-menu-guias-icon"))
            guias_tab.click()
            logging.info("Clicou na aba 'Guias' com sucesso.")

            # Aguarda até que a aba de guias seja totalmente carregada
            time.sleep(3)

            self.acessar_guia_sadt()

        except Exception as e:
            logging.error(f"Erro ao clicar na aba 'Guias': {e}")
            return

    def acessar_guia_sadt(self):
        """Acessa o 'Guia de SP/SADT' na aba 'Guias'."""
        try:
            logging.info("Procurando o botão 'Guia de SP/SADT'...")
            guia_sadt_button = self.acessar_com_reattempt((By.CSS_SELECTOR, ".menuLink.guia-spsadt-icon"))
            guia_sadt_button.click()
            logging.info("Clicou em 'Guia de SP/SADT' com sucesso.")

            # Espera para garantir que a nova aba carregue completamente
            time.sleep(3)

            # Lida com o alerta caso ele apareça
            self.lidar_com_alerta()

            # Preenche o número da carteira apenas após o alerta ser fechado
            self.preencher_numero_carteira()

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

            self.selecionar_relatorio_clinico()

            self.Anexando_RC()

            self.salvar_confirmar()

            lista_numeros = []  # Inicializa a lista para armazenar os números
            self.salvar_anotar_numero(lista_numeros) #salva na lista

            self.salvar_numero_no_excel(lista_numeros[-1]) #armazena no excel [-1] é para pegar o último termo adicionado

            

        except Exception as e:
            logging.error(f"Erro ao processar o guia SP/SADT: {e}")
            return

    def lidar_com_alerta(self):
        """Lida com possíveis alertas que possam aparecer na página."""
        try:
            logging.info("Verificando se há um alerta na página...")
            alert_present = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "noty_top_layout_container"))
            )
            if alert_present:
                fechar_button = self.driver.find_element(By.ID, "button-1")
                fechar_button.click()
                logging.info("Alerta fechado com sucesso.")
                # Adiciona uma espera extra para garantir que o alerta tenha sido fechado corretamente
                time.sleep(2)
        except (TimeoutException, NoSuchElementException):
            logging.info("Nenhum alerta foi detectado.")
        except Exception as e:
            logging.error(f"Erro ao lidar com o alerta: {e}")

    def preencher_numero_carteira(self):
        """Preenche o campo 'Número da Carteira' com dados do Excel."""
        try:
            logging.info("Preenchendo o campo 'Número da Carteira'...")
            numero_carteira = self.get_excel_value('CARTEIRA')

            # Localiza o campo e preenche com o número da carteira
            numero_carteira_input = self.acessar_com_reattempt((By.ID, "numeroDaCarteira"))
            numero_carteira_input.send_keys(numero_carteira)

            logging.info(f"Campo 'Número da Carteira' preenchido com sucesso com o valor: {numero_carteira}")

            time.sleep(2)

            # Pressiona a seta para baixo para selecionar o paciente
            numero_carteira_input.send_keys(Keys.ARROW_DOWN)

            # Pressiona 'Enter' para confirmar a seleção
            numero_carteira_input.send_keys(Keys.ENTER)

            logging.info(f"Paciente selecionado com sucesso.")

        except Exception as e:
            logging.error(f"Erro ao preencher o campo 'Número da Carteira': {e}")

    def preencher_carater_atendimento(self):
        try:
            logging.info("Preenchendo o campo 'Caráter do Atendimento'...")

            # Localiza o campo "Caráter do Atendimento"
            carater_atendimento_input = self.acessar_com_reattempt((By.ID, "caraterAtendimento"))

            # Clica no campo
            carater_atendimento_input.click()

            # Seta para baixo e seleciona a opção sugerida
            carater_atendimento_input.send_keys(Keys.ARROW_DOWN)  # seta para baixo
            carater_atendimento_input.send_keys(Keys.ENTER)  # tecla enter

            logging.info("Campo 'Caráter do Atendimento' preenchido com sucesso.")

            time.sleep(3)

        except Exception as e:
            logging.error(f"Erro ao preencher o campo 'Caráter do Atendimento': {e}")

    def preencher_indicacao_clinica(self):
        """Preenche o campo 'Indicação Clínica' com dados do Excel."""
        try:
            logging.info("Preenchendo o campo 'Indicação Clínica'...")
            indicacao_clinica = self.get_excel_value('indicacao_clinica')

            # Localiza o campo e preenche com o valor da indicação clínica
            indicacao_clinica_input = self.acessar_com_reattempt((By.ID, "indicacaoClinica"))
            indicacao_clinica_input.send_keys(indicacao_clinica)

            logging.info(f"Campo 'Indicação Clínica' preenchido com sucesso com o valor: {indicacao_clinica}")

        except Exception as e:
            logging.error(f"Erro ao preencher o campo 'Indicação Clínica': {e}")

    def acessar_procedimentos(self):
        """Clica no campo 'Procedimentos' para abrir a aba, garantindo que ele esteja visível."""
        try:
            logging.info("Abrindo a aba 'Procedimentos'...")

            # Localiza o campo "Procedimentos" usando o ID do HTML
            procedimentos_tab = self.acessar_com_reattempt((By.ID, "ui-accordion-accordion-header-2"))

            # Rolagem para garantir que o campo esteja visível na tela
            self.driver.execute_script("arguments[0].scrollIntoView(true);", procedimentos_tab)

            # Aguarda um pouco para garantir que a rolagem tenha sido completada
            time.sleep(1)

            # Clica no campo para expandir a aba
            procedimentos_tab.click()

            time.sleep(2)

            logging.info("Aba 'Procedimentos' aberta com sucesso.")

        except Exception as e:
            logging.error(f"Erro ao abrir a aba 'Procedimentos': {e}")

    def clicar_inserir_e_preencher(self):
        """Clica no botão 'Inserir Procedimento' e preenche o campo subsequente com dados do Excel."""
        try:
            logging.info("Inserindo procedimento...")

            # Localiza o botão pelo ID
            inserir_button = self.acessar_com_reattempt((By.ID, "incluirProcedimento"))

            # Clica no botão
            inserir_button.click()
            logging.info("Botão 'Inserir Procedimento' clicado com sucesso.")

            # Aguarda 1 segundo para garantir que o campo de preenchimento esteja pronto
            time.sleep(1)

            # Pega o valor da coluna 'Status' como string, preservando os zeros à esquerda
            procedimento = self.get_excel_value('Status').zfill(8)  # Ajuste conforme o número de dígitos

            # Localiza o campo de preenchimento que é ativado após o clique no botão 'Inserir Procedimento'
            procedimento_input = self.acessar_com_reattempt((By.XPATH, '//*[@id="registroProcedimentoCodigo"]/input'))

            # Preenche o campo com o valor extraído da coluna 'Status', preservando todos os dígitos
            procedimento_input.send_keys(procedimento)

            time.sleep(2)

            procedimento_input.send_keys(Keys.RETURN)

            logging.info("Procedimento selecionado com sucesso.")

            # Limpa e preenche o campo 'Quantidade'
            total = self.get_excel_value('TOTAL')

            total_input = self.acessar_com_reattempt((By.XPATH, '//*[@id="registroProcedimentoQuantidade"]/input'))
            total_input.clear()
            total_input.send_keys(total)

            time.sleep(2)

            total_input.send_keys(Keys.RETURN)

            time.sleep(3)

            logging.info(f"Campo 'QUANTIDADE' preenchido com sucesso com o valor: {total}")

            # Localiza o botão 'Confirmar' pelo XPath
            confirmar_button = self.acessar_com_reattempt((By.XPATH, '//*[@id="confirmarEdicaoDeProcedimento"]'))

            # Clica no botão 'Confirmar' para salvar as alterações
            confirmar_button.click()

            time.sleep(2)

        except Exception as e:
            logging.error(f"Erro ao preencher o procedimento: {e}")

    def preencher_campo_profissionais(self):
        """Mantém a implementação original desta função devido às peculiaridades do site."""
        try:
            logging.info("Preenchendo campos dos profissionais...")

            # Abertura da planilha e extração do valor da coluna CBO
            cbo = str(self.df['CBO'].iloc[self.row_index])[:6]

            # Localiza e expande o campo 'Profissionais'
            profissionais_tab = self.acessar_com_reattempt((By.XPATH, '//*[@id="ui-accordion-accordion-header-3"]'))
            profissionais_tab.click()
            logging.info("Aba 'Profissionais' expandida com sucesso.")

            time.sleep(1)

            # Clica no botão 'Inserir'
            inserir_button = self.acessar_com_reattempt((By.XPATH, '//*[@id="incluirProfissional"]'))
            inserir_button.click()
            logging.info("Botão 'Inserir Profissional' clicado com sucesso.")

            time.sleep(2)

            # Localiza o campo 'Seq.Grau Partic.' e seleciona a opção (12) Clínico
            grau_partic_input = self.acessar_com_reattempt((By.XPATH, '//*[@id="registroProfissionalGrauParticipacao"]/input'))
            grau_partic_input.click()
            logging.info("Campo 'Seq.Grau Partic.' clicado com sucesso.")

            time.sleep(2)

            for _ in range(5):  # Ajuste o número de setas para baixo conforme necessário
                grau_partic_input.send_keys(Keys.ARROW_DOWN)
                time.sleep(0.2)

            grau_partic_input.send_keys(Keys.RETURN)
            time.sleep(1)

            # Preencher campo 'Profissional'
            cod_profissional = self.get_excel_value('COD_PROFISSIONAL')
            profissional_codigo_input = self.acessar_com_reattempt((By.XPATH, '//*[@id="registroProfissionalCodigo"]/input'))
            profissional_codigo_input.send_keys(cod_profissional)
            profissional_codigo_input.send_keys(Keys.RETURN)
            logging.info("Campo 'Profissional' preenchido com sucesso.")

            time.sleep(2)

            # Preencher campo 'CBO'
            cbo_input = self.acessar_com_reattempt((By.XPATH, '//*[@id="registroProfissionalCodCBO"]/input'))
            cbo_input.click()

            # Selecionar a opção correta baseada no CBO
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
            logging.info("Campo 'CBO' preenchido com sucesso.")

            time.sleep(2)

            # Clicar no botão 'Confirmar' após preencher o CBO
            confirmar_button = self.acessar_com_reattempt((By.XPATH, '//*[@id="confirmarEdicaoDeProfissional"]'))
            confirmar_button.click()
            logging.info("Botão 'Confirmar' clicado com sucesso.")

        except Exception as e:
            logging.error(f"Erro ao preencher os campos dos profissionais: {e}")


    def preencher_observacao_justificativa(self):
        """Preenche o campo 'Observação/Justificativa' com dados do Excel."""
        try:
            logging.info("Preenchendo campo 'Observação/Justificativa'...")
            justificativa = self.get_excel_value('JUSTIFICATIVA')

            # Rolagem para garantir que o campo esteja visível na tela
            observacao_tab = self.acessar_com_reattempt((By.ID, "ui-accordion-accordion-header-4"))
            self.driver.execute_script("arguments[0].scrollIntoView(true);", observacao_tab)
            time.sleep(1)

            observacao_tab.click()
            logging.info("Aba 'Observação/Justificativa' foi clicada com sucesso.")

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
            logging.info("Selecionando o tipo de anexo...")

            tipo_anexo_dropdown = self.acessar_com_reattempt((By.ID, "tipoAnexoGuiaUpload"))
            tipo_anexo_dropdown.click()
            time.sleep(1)

            # Simula a tecla para baixo várias vezes até chegar na opção desejada
            for _ in range(46):  # Ajuste o número de vezes conforme necessário
                tipo_anexo_dropdown.send_keys(Keys.ARROW_DOWN)

            tipo_anexo_dropdown.send_keys(Keys.RETURN)
            logging.info("Tipo de anexo selecionado com sucesso.")

        except Exception as e:
            logging.error(f"Erro ao selecionar o tipo de anexo: {e}")



    def Anexando_RM(self):
        try:
            logging.info("Iniciando o processo de upload dos arquivos RM do paciente...")

            # Define o caminho base
            base_path = Path(r"G:\Meu Drive\IPASGO\1.RELATORIO MEDICO E CLINICO")

            # Obtém 'Paciente' e 'CARTEIRA' da planilha Excel
            nome_paciente = self.get_excel_value('Paciente')
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

            time.sleep(1)

        except Exception as e:
            logging.error(f"Erro ao fazer upload dos arquivos RM do paciente: {e}")


    def selecionar_relatorio_clinico(self):
        """Seleciona o tipo de anexo usando teclas de navegação."""
        try:
            logging.info("Selecionando o tipo de anexo...")

            tipo_anexo_dropdown = self.acessar_com_reattempt((By.ID, "tipoAnexoGuiaUpload"))
            tipo_anexo_dropdown.click()
            time.sleep(1)

            # Simula a tecla para baixo várias vezes até chegar na opção desejada
            for _ in range(9):  # Ajuste o número de vezes conforme necessário
                tipo_anexo_dropdown.send_keys(Keys.ARROW_DOWN)

            tipo_anexo_dropdown.send_keys(Keys.RETURN)
            logging.info("Tipo de anexo selecionado com sucesso.")

        except Exception as e:
            logging.error(f"Erro ao selecionar o tipo de anexo: {e}")


    def Anexando_RC(self):
        try:
            logging.info("Iniciando o processo de upload dos arquivos RC do paciente...")

            # Define o caminho base
            base_path = Path(r"G:\Meu Drive\IPASGO\1.RELATORIO MEDICO E CLINICO")

            # Obtém 'Paciente', 'CARTEIRA' e 'CBO' da planilha Excel
            nome_paciente = self.get_excel_value('Paciente')
            id_paciente = self.get_excel_value('CARTEIRA')
            
            # Converte o valor de CBO para string, removendo qualquer ponto decimal
            cbo = self.get_excel_value('CBO')
            cbo = str(int(float(cbo)))  # Converte para inteiro e depois para string

            if not nome_paciente or not id_paciente or not cbo:
                logging.error("Falha ao obter o nome, ID do paciente ou CBO da planilha.")
                return
            logging.info(f"Paciente: {nome_paciente}, ID: {id_paciente}, CBO: {cbo}")

            # Constrói o nome da pasta do paciente
            patient_folder_name = f"{nome_paciente}-{id_paciente}"

            # Constrói o caminho completo para a pasta do paciente
            patient_folder_path = base_path / patient_folder_name

            logging.info(f"Caminho da pasta do paciente: {patient_folder_path}")

            # Verifica se a pasta do paciente existe
            if not patient_folder_path.is_dir():
                logging.error(f"A pasta do paciente '{patient_folder_path}' não foi encontrada.")
                return

            # Encontra os arquivos do paciente, focando apenas em RC
            _, arquivos_rc_fono, arquivos_rc_psi, arquivos_rc_to = encontrar_arquivos_paciente(patient_folder_path, id_paciente, nome_paciente)

            # Verifica o tipo de RC com base no CBO
            if cbo == "251510":  # PSI
                arquivos_rc = arquivos_rc_psi
                logging.info("CBO indica PSICOLOGIA. Selecionando arquivos PSI.")
            elif cbo == "223810":  # FONO
                arquivos_rc = arquivos_rc_fono
                logging.info("CBO indica FONOAUDIOLOGIA. Selecionando arquivos FONO.")
            elif cbo == "223905":  # TO
                arquivos_rc = arquivos_rc_to
                logging.info("CBO indica TERAPIA OCUPACIONAL. Selecionando arquivos TO.")
            else:
                logging.error(f"CBO '{cbo}' não corresponde a uma especialidade conhecida.")
                return

            if not arquivos_rc:
                logging.error(f"Nenhum arquivo RC correspondente foi encontrado para o paciente e CBO: {cbo}.")
                return

            logging.info(f"Arquivos RC encontrados: {arquivos_rc}")

            # Faz o upload do primeiro arquivo RC encontrado
            for arquivo_para_upload in arquivos_rc:
                # Localiza o elemento <input type="file">
                input_file = self.driver.find_element(By.CSS_SELECTOR, 'input[type="file"]')
                logging.info("Elemento de upload encontrado.")

                # Envia o caminho do arquivo para o elemento 'input'
                input_file.send_keys(str(arquivo_para_upload))
                logging.info(f"Arquivo '{arquivo_para_upload}' selecionado com sucesso.")

                time.sleep(2)

                break  # Encerra o loop após carregar o primeiro arquivo RC

            # Clica no botão 'Adicionar'
            self.safe_click((By.XPATH, '//*[@id="upload_form"]/div/input[2]'))
            time.sleep(1)

        except Exception as e:
            logging.error(f"Erro ao fazer upload dos arquivos RC do paciente: {e}")



    def salvar_confirmar(self):
        try:
            logging.info("Iniciando o processo de salvar...")

            # Localiza o botão 'Salvar' usando o XPATH
            salvar_button = self.driver.find_element(By.XPATH, '//*[@id="btnGravar"]')

            # Usa JavaScript para rolar até o botão 'Salvar'
            self.driver.execute_script("arguments[0].scrollIntoView(true);", salvar_button)
            logging.info("Botão 'Salvar' rolado para a visualização.")

            # Aguarda um pouco para garantir que o elemento esteja visível
            time.sleep(1)

            # Clica no botão 'Salvar'
            salvar_button.click()
            logging.info("Botão 'Salvar' clicado com sucesso.")

            # Aguarda até que o modal de confirmação apareça
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/div[8]/div[3]/div/button[1]'))
            )
            logging.info("Modal de confirmação detectado.")

            # Localiza o botão 'Sim' na janela de confirmação
            confirmar_button = self.driver.find_element(By.XPATH, '/html/body/div[8]/div[3]/div/button[1]')

            # Usa JavaScript para rolar até o botão 'Sim'
            self.driver.execute_script("arguments[0].scrollIntoView(true);", confirmar_button)
            logging.info("Botão 'Sim' rolado para a visualização.")

            # Aguarda um pouco para garantir que o elemento esteja visível
            time.sleep(1)

            # Clica no botão 'Sim' para confirmar a ação
            confirmar_button.click()
            logging.info("Botão 'Sim' clicado com sucesso.")
        except Exception as e:
            logging.error(f"Erro ao tentar salvar e confirmar, tente refazer o processo: {e}")


    def salvar_anotar_numero(self, lista_numeros):
        try:
            logging.info("Iniciando o processo de captura do número da guia...")

            # Primeiro, verificamos se o pop-up está visível
            pup_up = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/div[8]'))
            )
            logging.info("Pop-up localizado com sucesso.")

            # Depois, verificamos a tela de diálogo
            tela_dialogo = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="ui-id-47"]'))
            )
            logging.info("Tela de diálogo localizada com sucesso.")

            # Em seguida, localizamos o campo de texto que contém o número
            dialogo_texto = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="dialogText"]'))
            )
            logging.info("Campo de texto localizado com sucesso.")

            # Agora, localizamos o número da guia (o texto fora das tags <b>)
            elemento_numero_guia = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="dialogText"]/div[2]'))
            )
            logging.info("Elemento contendo o número da guia localizado com sucesso.")

            # O número está dentro do texto do 'div', após a tag <b>. Precisamos pegar o número extra.
            numero_guia_completo = elemento_numero_guia.text.strip()

            # Como o número está logo após o 'Nº Guia Operadora:', extraímos o número especificamente
            numero_guia = numero_guia_completo.split("Nº Guia Operadora:")[-1].strip()
            logging.info(f"Número da Guia capturado: {numero_guia}")

            # Adiciona o número à lista de números
            lista_numeros.append(numero_guia)

            # Grava a lista de números em um arquivo txt
            with open('numeros_guias.txt', 'w') as f:
                for numero in lista_numeros:
                    f.write(f"{numero}\n")
            logging.info(f"Números das guias salvos no arquivo 'numeros_guias.txt'.")

            # Após capturar o número, pressionar ESC para fechar o aviso
            logging.info("Tentando fechar o pop-up com a tecla ESC...")
            time.sleep(1)
            # Envia a tecla ESC para fechar o pop-up
            ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
            logging.info("Tecla ESC enviada com sucesso para fechar o pop-up.")

        except Exception as e:
            logging.error(f"Erro ao capturar o número da guia: {e}")


    def salvar_numero_no_excel(self, numero_guia):
        try:
            logging.info("Iniciando o processo de salvamento do número no Excel usando xlwings...")

            # Cria uma cópia do arquivo Excel original
            file_copy_path = r"C:\Users\AUDITORIA_CLMF\Desktop\ipasgo.py\joaop\BASE_AUTORIZACAO_COPIA.xlsx"

            # Abre o arquivo Excel existente com xlwings
            wb = xw.Book(self.file_path)
            ws = wb.sheets[self.sheet_name]

            # Define a linha a ser atualizada a partir de self.row_index, ajustando para a linha real (linha+2)
            row_to_update = self.row_index + 2  # Ajusta o índice para Excel (row_index no código -> linha no Excel)

            # Define a coluna onde deseja salvar o número
            col_to_update = 'P'  # Ajuste a coluna para onde o número deve ser salvo

            # Salva o número na célula específica
            ws.range(f'{col_to_update}{row_to_update}').value = numero_guia

            # Salva a cópia do arquivo
            wb.save(file_copy_path)
            wb.close()
            logging.info(f"Número da Guia {numero_guia} salvo na cópia do Excel na linha {row_to_update}, coluna {col_to_update}.")

        except Exception as e:
            logging.error(f"Erro ao salvar o número no Excel: {e}")

   






if __name__ == "__main__":
    ipasgo = IpasgoAutomation()
    ipasgo.acessar_portal_ipasgo()

    # Fechar o navegador apenas quando autorizado
    input("Pressione 'S' para fechar o navegador...")
    ipasgo.close()
