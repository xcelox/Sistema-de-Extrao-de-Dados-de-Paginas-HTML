from flask import Blueprint, request, render_template, jsonify, send_file
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException  # Importação de TimeoutException e StaleElementReferenceException
from bs4 import BeautifulSoup
import pandas as pd
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from openpyxl import Workbook
from openpyxl.styles import Font
import time
from time import sleep
import io
import os
from forms import esperar_e_clicar, esperar_e_enviar_chaves, extrair_dados

sec_bp = Blueprint('sec', __name__)

TEMP_FILE = 'temp_data.csv'

@sec_bp.route('/SEC')
def index():
    return render_template('SEC.html')

@sec_bp.route('/SEC/generate', methods=['POST'])
def generate():
    data = request.form['data']
    if not data.strip():
        return jsonify({'error': 'O formulário está vazio. Por favor, insira os dados.'}), 400

    lines = data.split('\n')[:500]  # Limita a 500 linhas
    df = pd.DataFrame(lines, columns=["matricula"])

    # Configurações para o Chrome
    options = Options()
    options.add_argument("--start-maximized")  # Iniciar maximizado
    servico = Service(executable_path='chromedriver.exe')  # Caminho para o chromedriver
    navegador = webdriver.Chrome(service=servico, options=options)

    link = 'https://achei.correios.com.br/app/empregados/index.php'
    sleep(5)
    navegador.get(link)

    # Criação do arquivo Excel
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Cadastros"

    # Adiciona cabeçalhos
    headers = ["MATRÍCULA", "NOME", "LOCAL DE EXERCÍCIO", "CPF", "CARGO", "FUNÇÃO", "funcaoCampo", "ESPECIALIDADE", "SE", "MCU", "LOTAÇÃO", "funcaoCampo", "REFERÊNCIA", "SITUAÇÃO", "JORNADA", "FÉRIAS", "ADMISSÃO", "TSCP"]
    sheet.append(headers)

    # Formata o cabeçalho em negrito
    for cell in sheet[1]:
        cell.font = Font(bold=True)

    # Carrega dados temporários se existirem
    if os.path.exists(TEMP_FILE):
        temp_df = pd.read_csv(TEMP_FILE)
        for row in temp_df.values:
            sheet.append(row.tolist())
        processed_matriculas = set(temp_df["MATRÍCULA"].tolist())
    else:
        processed_matriculas = set()

    total_consultas = 0

    for linha in df["matricula"]:
        if linha in processed_matriculas:
            continue  # Pula matrículas já processadas

        try:
            esperar_e_enviar_chaves(navegador, '//*[@id="query"]', linha)
            esperar_e_clicar(navegador, '//*[@id="botao-pesquisar"]')
            wait = WebDriverWait(navegador, 20)
            tabelaCadastro = wait.until(EC.visibility_of_element_located((By.ID, "registros")))

            try:
                tabelaCadastro = navegador.find_element(by=By.ID, value="registros")
                htmlContent = tabelaCadastro.get_attribute("outerHTML")
            except StaleElementReferenceException:
                tabelaCadastro = navegador.find_element(by=By.ID, value="registros")
                htmlContent = tabelaCadastro.get_attribute("outerHTML")

            soup = BeautifulSoup(htmlContent, "html.parser")
            cadastros = soup.find(name="table")
            if cadastros is None:
                continue  # Pula para a próxima iteração se nenhuma tabela for encontrada

            try:
                df_table = pd.read_html(str(cadastros), header=0)[0]
            except ValueError:
                continue  # Pula para a próxima iteração se nenhuma tabela for encontrada

            for row in df_table.values:
                sheet.append(row.tolist())

            # Salva dados temporários
            temp_df = pd.DataFrame(df_table)
            temp_df.to_csv(TEMP_FILE, mode='a', header=not os.path.exists(TEMP_FILE), index=False)

            # Limpa o campo de entrada usando Selenium
            campo_query = navegador.find_element(by=By.XPATH, value='//*[@id="query"]')
            campo_query.clear()
            sleep(0.2)

            total_consultas += 1
        except (TimeoutException, StaleElementReferenceException):
            # Limpa o campo de entrada e continua
            campo_query = navegador.find_element(by=By.XPATH, value='//*[@id="query"]')
            campo_query.clear()
            continue

    # Salva o arquivo Excel em um buffer
    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)

    # Fecha o navegador
    navegador.quit()

    # Remove o arquivo temporário
    if os.path.exists(TEMP_FILE):
        os.remove(TEMP_FILE)
        
    print("Os dados foram salvos no arquivo Relatorio_Cadastro.xlsx.")

    # Faz o download automático do arquivo
    return send_file(buffer, as_attachment=True, download_name="Relatorio_Cadastro.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    