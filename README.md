# Solicitações de Guias Médicas para IPASGO

Este projeto foi desenvolvido para automatizar o processo de solicitação de guias médicas para o plano de saúde IPASGO utlizando o `*webdriver(selenium)`. O objetivo é otimizar o fluxo de requisições de guias médicas de maneira eficiente, com foco em minimizar o tempo de preenchimento e o risco de erros humanos. 

## Estrutura do Projeto

- **Arquivo Principal**: `joaop/teste 1 milhão.py` – Este script principal executa a automação das solicitações.
- **Arquivo de Referência**: `numeros_guias.txt` – Armazena números de guias que foram processadas.
- **Ambiente Virtual**: A pasta `joaop` contém os arquivos necessários para a configuração do ambiente virtual.

## Requisitos

Para rodar o projeto, você precisará instalar os seguintes requisitos:

- Python 3.x
- Bibliotecas do Python: 
  - `selenium.webdriver` (não esqueça de colocar o arquivo do selenium.webdriver no path de sua máquina)
  - Outras dependências podem ser instaladas via o arquivo de ambiente virtual na pasta `Lib`.

## Instruções de Instalação

1. Clone o repositório:

   ```bash
   git clone <https://github.com/joaosouza2/RPA_ipasgo.py.git>
   cd RPA_ipasgo.py-main
   ```

2. Crie e ative o ambiente virtual:

   ```bash
   python -m venv joaop
   source joaop/Scripts/activate
   ```

3. Instale as dependências:

   ```
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
      import inspect
   ```

## Como Executar

1. Verifique as estruturas dos seus dados e garanta que tenha todas as colunas necessárias para realização da solicitação da guia. Não esqueça de alterar endereço da sua planilha, e que ela seja `.xlsx`, caso contrário converta o pandas para o uso do csv como opção. 

2. Execute o script principal:

   ```bash
   python joaop/teste 1 milhão.py
   ```

O script irá processar as solicitações de guias médicas para o IPASGO automaticamente e salvará os resultados diretamente no arquivo `.xlsx` e no `numeros_guias.txt` ( No qual é responsável por armazenar os números de guias que foram processadas, ou os erros, contendo, horário da solicitação, paciente, especialidade e em qual etapa conteve erro)

## Licença

Este projeto é de propriedade de João Pedro Souza. Para dúvidas ou considerações, entre em contato via e-mail: joaosouza2@discente.ufg.br.