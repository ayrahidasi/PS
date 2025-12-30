import os
import sys
import pyautogui
import xlwings
from datetime import datetime, timedelta
import re
from playwright.sync_api import Playwright, sync_playwright, expect
import customtkinter as ctk
import pdfplumber
from tkinter import Tk, filedialog


# Função para definir planilha e aba
def obter_planilha_e_aba(nome_aba):
    diretorio = os.path.dirname(__file__)
    arquivo = os.path.basename(__file__)[:-3]
    diretorio_arquivo = f"{diretorio}\\{arquivo}.xlsm"
    work_book = xlwings.Book(diretorio_arquivo)
    planilha = work_book.sheets[nome_aba]
    return work_book, planilha


# Função para extrair texto do PDF
def extract_text_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        pages_text = []
        for page in pdf.pages:
            pages_text.append(page.extract_text())
        return pages_text


def extract_table_from_pdf(pdf_path):
    tables_data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if len(table) > 2:
                    table = table[1:-1]
                    for row in table:
                        if len(row) == 5:
                            tables_data.append(row)
    return tables_data

    
# Função para encontrar informações específicas no texto da NFSE e inserir na planilha
def find_information_NFSE(text):
    info = {}
    lines = text.split('\n')

    # Lista de CNPJs que não são válidos
    cnpjs_invalidos = ["00.530.352/0001-59", "00530352000159", "03.610.650/0001-47"]
    fone_sefaz = "156"

    for i, line in enumerate(lines):
        # Verifica a linha atual
        cpf_cnpj_match = re.search(r'(CPF/CNPJ|CNPJ)\s*[:\-]?\s*([\d./-]{2,}\d{2})', line)
        if cpf_cnpj_match:
            cnpj = cpf_cnpj_match.group(2)
            if cnpj not in cnpjs_invalidos:
                info['CPF/CNPJ'] = cnpj
                #break
        else:
            # Verifica a linha seguinte, se existir
            if i + 1 < len(lines):
                next_line = lines[i + 1]
                cpf_cnpj_match = re.search(r'(CPF/CNPJ|CNPJ)\s*[:\-]?\s*([\d./-]{2,}\d{2})', next_line)
                if cpf_cnpj_match:
                    cnpj = cpf_cnpj_match.group(2)
                    if cnpj not in cnpjs_invalidos:
                        info['CPF/CNPJ'] = cnpj
                        #break

        if "Número da Nota Fiscal" in line or "NÚMERO DA NOTA FISCAL" in line:
            if i + 1 < len(lines):
                next_line = lines[i + 1]
                if fone_sefaz not in next_line:
                    nf_number_match = re.search(r'([\d]+)', next_line)
                    if nf_number_match:
                        info['Número da Nota Fiscal'] = nf_number_match.group(1)
        # if "Número da Nota Fiscal" in line or "NÚMERO DA NOTA FISCAL" in line:
        #     if i + 1 < len(lines):
        #         nf_number_match = re.search(r'([\d]+)', lines[i + 1])
        #         if nf_number_match:
        #             info['Número da Nota Fiscal'] = nf_number_match.group(1)

        
        if "Data de Geração da NFS-e" in line or "Data Emissão" in line or "Dt. de Emissão" in line or "DATA E HORA DE EMISSÃO" in line or "Data da Geração da NFS-e" in line:
            if i + 1 < len(lines):
                gen_date_match = re.search(r'([\d/]+ [\d:]+)', lines[i + 1])
                if gen_date_match:
                    info['Data de Geração da NFS-e'] = gen_date_match.group(1)[:10] 

        if "Vl. Total dos Serviços" in line or "VALOR TOTAL DA NOTA" in line or "VALOR TOTAL DOS SERVIÇOS" in line:
            if i + 1 < len(lines):
                value_match = re.search(r'R\$ [\d.,]+', lines[i + 1])
                if value_match:
                    info['Vl. Total dos Serviços'] = value_match.group()#.replace('.', '').replace('R', '').replace('$', '').replace(' ', '')
        
    return info


def extrair_pdf_NFSE(page):
    
    work_book, planilha = obter_planilha_e_aba('Dados')

    # Pede ao usuário para inserir o valor da linha inicial
    linha_inicial = pyautogui.prompt(text='Insira o valor da linha inicial:', title='Linha Inicial', default='2')
    linha_inicial = int(linha_inicial)

    # Pergunta ao usuário se deseja baixar a NFSE ou selecionar um arquivo local
    escolha = pyautogui.confirm(text='Você deseja baixar a NFSE ou selecionar um arquivo local?', buttons=['Baixar', 'Selecionar'])

    if escolha == 'Baixar':
        # Encontra o número do processo
        processo_edoc_texto = page.locator("span[id^='document_header_layout_form:'] > h1").text_content()
        processo_edoc = re.search(r"\d{6,7}/\d{4}", processo_edoc_texto).group(0).replace('/', '_')
        planilha.range('B1').value = processo_edoc.replace('_', '/')
        
        # Nomeia o arquivo baixado
        nome_arquivo = f'NF_{processo_edoc}.pdf'
        
        # Download da NFSE
        page.locator("img[src='icons2/documento-big.png']").click()
        with page.expect_download() as download_info:
            page.locator("img[alt='Download']").click()
        page.wait_for_timeout(1000)
        download = download_info.value
        download.save_as(nome_arquivo)
        page.locator("a[accesskey='V']").click()

        # Caminho para o seu arquivo PDF da NFSE
        pdf_path_nfse = f'{diretorio}\\{nome_arquivo}'

    else:
        # Seleciona um arquivo local usando customtkinter
        root = ctk.CTk()
        root.withdraw()  # Oculta a janela principal do Tkinter
        pdf_path_nfse = filedialog.askopenfilename(title="Selecione o arquivo PDF da NFSE", filetypes=[("PDF files", "*.pdf")])
        root.destroy()

    # Extraindo texto do PDF da NFSE
    text_nfse_pages = extract_text_from_pdf(pdf_path_nfse)

    # Encontrando informações específicas em cada página da NFSE e inserindo na planilha "Conferência"
    all_info_nfse = []
    for page_text in text_nfse_pages:
        info = find_information_NFSE(page_text)
        all_info_nfse.append(info)

    # Escrevendo os dados extraídos da NFSE na planilha
    for i, info in enumerate(all_info_nfse, start=linha_inicial):
        planilha.range(f'D{i}').value = info.get('CPF/CNPJ', '')
        planilha.range(f'F{i}').value = info.get('Número da Nota Fiscal', '')
        planilha.range(f'E{i}').value = info.get('Data de Geração da NFS-e', '')
        planilha.range(f'G{i}').value = info.get('Vl. Total dos Serviços', '')

    # Salvar o arquivo Excel
    work_book.save()


def find_information_ateste(text):
    info = {}
    lines = text.split('\n')

    for i, line in enumerate(lines):
        if "Prestador" in line:
            info['Prestador'] = line.split("Prestador")[-1].strip()
        if "Banco" in line:
            info['Banco'] = line.split("Banco")[-1].strip().replace(": ", "")
        if "Agência" in line:
            info['Agência'] = line.split("Agência")[-1].strip().replace(": ", "")[:4]  # Apenas os 4 primeiros digitos
        if "Conta" in line:
            info['Conta'] = line.split("Conta")[-1].strip().replace(": ", "")
        if "DV" in line:
            info['DV'] = line.split("DV")[-1].strip().replace(": ", "")
        if "Serviço realizado em" in line:            
            info['Serviço realizado em'] = line.split("Serviço realizado em")[-1].strip().replace(": ", "")
        if "Nota de empenho" in line:            
            info['Nota de empenho'] = line.split("Nota de empenho")[-1].strip().replace(": ", "")
        if "VALOR DE PAGTO BRUTO" in line:
            combined_line = line
            if i + 1 < len(lines):
                combined_line += " " + lines[i + 1]
            values = extract_values_from_line(combined_line)
            if len(values) == 7:
                info['VALOR DE PAGTO BRUTO'] = float(values[0].replace('.', '').replace(',', '.'))
                info['IR'] = float(values[1].replace('.', '').replace(',', '.'))
                info['PIS'] = float(values[2].replace('.', '').replace(',', '.'))
                info['COFINS'] = float(values[3].replace('.', '').replace(',', '.'))
                info['CSLL'] = float(values[4].replace('.', '').replace(',', '.'))
                info['ISS'] = float(values[5].replace('.', '').replace(',', '.'))
                info['VALOR PGTO LÍQUIDO'] = float(values[6].replace('.', '').replace(',', '.'))

        if "NF" in line and "NF's" not in line:
            info['NF'] = line.split("NF")[-1].strip()
            if i + 1 < len(lines):
                next_line = lines[i + 1]
                values = extract_values_from_line(next_line)
                if len(values) >= 4:
                    try:
                        info['VALOR APRESENTADO'] = float(values[1][2:].strip().replace('.', '').replace(',', '.'))  # Remove the first two characters (R$)
                        info['VALOR GLOSADO'] = float(values[2][2:].strip().replace('.', '').replace(',', '.'))  # Remove the first two characters (R$)
                        info['VALOR APRESENTADO DE RECURSO'] = float(values[3][2:].strip().replace('.', '').replace(',', '.'))  # Remove the first two characters (R$)
                        info['VALOR GLOSADO DE RECURSO'] = float(values[4][2:].strip().replace('.', '').replace(',', '.'))
                    except:
                        continue
    return info


def extrair_pdf_ateste(page):
    diretorio = os.path.dirname(__file__)
    arquivo = os.path.basename(__file__)[:-3]
    diretorio_arquivo = f"{diretorio}\\{arquivo}.xlsm"

    work_book = xlwings.Book(diretorio_arquivo)
    planilha_dados = work_book.sheets['Dados']
    planilha_bancos = work_book.sheets['Bancos']

    funcao_extracao = find_information_ateste

    # Pergunta ao usuário se deseja baixar a NFSE ou selecionar um arquivo local
    escolha = pyautogui.confirm(text='Você deseja baixar a NFSE ou selecionar um arquivo local?', buttons=['Baixar', 'Selecionar'])

    if escolha == 'Baixar':
        # Encontra o número do processo
        processo_edoc_texto = page.locator("span[id^='document_header_layout_form:'] > h1").text_content()
        processo_edoc = re.search(r"\d{6,7}/\d{4}", processo_edoc_texto).group(0).replace('/', '_')
        planilha_dados.range('B1').value = processo_edoc.replace('_', '/')
        
        # Nomeia o arquivo baixado
        nome_arquivo = f'R_{processo_edoc}.pdf'
        
        # Download do ateste
        page.locator("img[src='icons2/documento-big.png']").click()
        with page.expect_download() as download_info:
            page.locator("img[alt='Download']").click()
        page.wait_for_timeout(1000)
        download = download_info.value
        download.save_as(nome_arquivo)
        page.locator("a[accesskey='V']").click()

        # Caminho para o seu arquivo PDF 
        pdf_path = f'{diretorio}\\{nome_arquivo}'

    else:
        # Seleciona um arquivo local usando customtkinter
        root = ctk.CTk()
        root.withdraw()  # Oculta a janela principal do Tkinter
        pdf_path = filedialog.askopenfilename(title="Selecione o arquivo PDF da NFSE", filetypes=[("PDF files", "*.pdf")])
        root.destroy()

    # Extraindo texto do PDF da NFSE
    text_pages = extract_text_from_pdf(pdf_path)
    tables_data = extract_table_from_pdf(pdf_path)

    all_info = []
    for page_text in text_pages:
        info = funcao_extracao(page_text)
        all_info.append(info)

    for i, info in enumerate(all_info, start=22):
        prestador_full_text = info.get('Prestador', '').replace(": ", "")
        prestador_id = prestador_full_text[:14]  # First 14 characters
        prestador_name = prestador_full_text[17:]  # Characters after position 17
        municipio = info.get('Serviço realizado em', '').upper()
        if municipio == 'DISTRITO FEDERAL':
            municipio = 'DF'

        conta = info.get('Conta', '').replace(':', '')
        dv = info.get('DV', '').replace(':', '')
        conta_dv = f"{conta}{dv}"

        planilha_dados.range(f'C{i}').value = prestador_name
        planilha_dados.range(f'D{i}').value = prestador_id 
        planilha_dados.range(f'M{i}').value = municipio 
        #planilha_dados.range(f'E{i}').value = info.get('Nota de empenho', '')
        planilha_dados.range(f'P{i}').value = info.get('Banco', '')
        planilha_dados.range(f'Q{i}').value = info.get('Agência', '')
        planilha_dados.range(f'R{i}').value = conta_dv
        planilha_dados.range(f'G{i}').value = info.get('VALOR DE PAGTO BRUTO', '')
        planilha_dados.range(f'N{i}').value = info.get('VALOR PGTO LÍQUIDO', '')
        # planilha_dados.range(f'K{i}').value = info.get('IR', '')
        # planilha_dados.range(f'L{i}').value = info.get('PIS', '')
        # planilha_dados.range(f'M{i}').value = info.get('COFINS', '')
        # planilha_dados.range(f'N{i}').value = info.get('CSLL', '')
        # planilha_dados.range(f'O{i}').value = info.get('ISS', '')

    linha_excel = 22
    for row in tables_data:
    # Ignorar linhas inválidas
        if (
            not row or
            row[0] is None or row[0] == '' or
            any(isinstance(cell, str) and any(palavra in cell.upper() for palavra in ['APRESENTADO', 'GLOSADO', 'RECURSO']) for cell in row)):
            continue

        # Escreve os dados na planilha
        planilha_dados.range(f'F{linha_excel}').value = row[0]
        planilha_dados.range(f'H{linha_excel}').value = extrair_valor(row[1]) if len(row) > 1 else 0.0
        planilha_dados.range(f'I{linha_excel}').value = extrair_valor(row[2]) if len(row) > 2 else 0.0
        planilha_dados.range(f'J{linha_excel}').value = extrair_valor(row[3]) if len(row) > 3 else 0.0
        planilha_dados.range(f'K{linha_excel}').value = extrair_valor(row[4]) if len(row) > 4 else 0.0

        linha_excel += 1 # Só incrementa se a linha for válida


    update_bank_info(planilha_dados, planilha_bancos)

    work_book.save()


def find_information_ateste_cnpj_unico(text):
    info = {}

    # Extrair dados do prestador e banco
    match = re.search(r"Prestador:\s*(\d{14})\s*-\s*(.*?)\s+Banco:\s*(.*?)\s+Agência:\s*(\d{4})-?X?\s+Conta:\s*(\d+)\s+DV:\s*(\d+)", text)
    if match:
        info['Prestador_ID'] = match.group(1)
        info['Prestador_Nome'] = match.group(2).strip()
        info['Banco'] = match.group(3).strip()
        info['Agência'] = match.group(4).strip()
        info['Conta'] = match.group(5).strip()
        info['DV'] = match.group(6).strip()
        info['nota_de_empenho_ateste'] = match.group(6).strip()

    # Extrair valores principais
    valores = re.findall(r'R\$[\s]*[\d\.,]+', text)
    valores_float = [float(v.replace('R$', '').replace('.', '').replace(',', '.').strip()) for v in valores]

    if len(valores_float) >= 7:
        info['VALOR DE PAGTO BRUTO'] = valores_float[0]
        info['IR'] = valores_float[1]
        info['PIS'] = valores_float[2]
        info['CSLL'] = valores_float[3]
        info['ISS'] = valores_float[4]
        info['VALOR PGTO LÍQUIDO'] = valores_float[5]
        info['COFINS'] = valores_float[6]

    return info


def extract_values_from_line(line):
    values = re.findall(r'[\d.,]+', line)
    values = [value.strip() for value in values]
    return values

def update_bank_info(sheet_dados, sheet_bancos):
    bancos_dict = {}
    for row in sheet_bancos.range('A2:B' + str(sheet_bancos.cells.last_cell.row)).value:
        if row[1] is not None:
            bancos_dict[row[1].strip()] = row[0]

    last_row_dados = sheet_dados.range('P' + str(sheet_dados.cells.last_cell.row)).end('up').row

    for i in range(16, last_row_dados + 1):
        banco_name = sheet_dados.range(f'P{i}').value
        if banco_name is not None and banco_name.strip() != '':
            banco_name = banco_name.strip()
            if banco_name in bancos_dict:
                sheet_dados.range(f'P{i}').value = bancos_dict[banco_name]


def extrair_valor(valor):  
    if valor and isinstance(valor, str) and valor.startswith('R$'):
        try:
            return float(valor[2:].strip().replace('.', '').replace(',', '.'))
        except ValueError:
            return 0.0
    return 0.0


def extrair_pdf_ateste_cnpj_unico(page):
    diretorio = os.path.dirname(__file__)
    arquivo = os.path.basename(__file__)[:-3]
    diretorio_arquivo = f"{diretorio}\\{arquivo}.xlsm"
    work_book = xlwings.Book(diretorio_arquivo)
    planilha_dados = work_book.sheets['Dados']
    planilha_bancos = work_book.sheets['Bancos']

    # linha_inicial = pyautogui.prompt(text='Insira o valor da linha inicial:', title='Linha Inicial', default='2')
    # linha_inicial = int(linha_inicial)

    # # Escolha da função de extração
    # funcao_escolhida = pyautogui.confirm(
    # text='Qual função deseja usar para extrair os dados do PDF?',
    # buttons=['find_information_ateste', 'find_information_ateste_cnpj_unico'])

    # if funcao_escolhida == 'find_information_ateste':
    #     funcao_extracao = find_information_ateste
    # else:
    funcao_extracao = find_information_ateste_cnpj_unico

    # Pergunta ao usuário se deseja baixar a NFSE ou selecionar um arquivo local
    escolha = pyautogui.confirm(text='Você deseja baixar a NFSE ou selecionar um arquivo local?', buttons=['Baixar', 'Selecionar'])

    if escolha == 'Baixar':
        # Encontra o número do processo
        processo_edoc_texto = page.locator("span[id^='document_header_layout_form:'] > h1").text_content()
        processo_edoc = re.search(r"\d{6,7}/\d{4}", processo_edoc_texto).group(0).replace('/', '_')
        planilha_dados.range('B1').value = processo_edoc.replace('_', '/')
        
        # Nomeia o arquivo baixado
        nome_arquivo = f'R_{processo_edoc}.pdf'
        
        # Download do ateste
        page.locator("img[src='icons2/documento-big.png']").click()
        with page.expect_download() as download_info:
            page.locator("img[alt='Download']").click()
        page.wait_for_timeout(1000)
        download = download_info.value
        download.save_as(nome_arquivo)
        page.locator("a[accesskey='V']").click()

        # Caminho para o seu arquivo PDF 
        pdf_path = f'{diretorio}\\{nome_arquivo}'

    else:
        # Seleciona um arquivo local usando customtkinter
        root = ctk.CTk()
        root.withdraw()  # Oculta a janela principal do Tkinter
        pdf_path = filedialog.askopenfilename(title="Selecione o arquivo PDF da NFSE", filetypes=[("PDF files", "*.pdf")])
        root.destroy()

    # Extraindo texto do PDF da NFSE
    text_pages = extract_text_from_pdf(pdf_path)
    tables_data = extract_table_from_pdf(pdf_path)

    all_info = []
    for page_text in text_pages:
        info = funcao_extracao(page_text)
        all_info.append(info)

    for i, info in enumerate(all_info, start=22):
        prestador_full_text = info.get('Prestador', '').replace(": ", "")
        prestador_id = prestador_full_text[:14]  # First 14 characters
        prestador_name = prestador_full_text[17:]  # Characters after position 17

        conta = info.get('Conta', '').replace(':', '')
        dv = info.get('DV', '').replace(':', '')
        conta_dv = f"{conta}{dv}"


        # Preencher apenas se as células estiverem vazias
        # if not planilha_dados.range('W18').value:
        #     planilha_dados.range('W18').value = prestador_name
        # if not planilha_dados.range('W19').value:
        #     planilha_dados.range('W19').value = prestador_id
        if not planilha_dados.range('P2').value:
            planilha_dados.range('P2').value = info.get('Banco', '')
        if not planilha_dados.range('Q2').value:
            planilha_dados.range('Q2').value = info.get('Agência', '')
        if not planilha_dados.range('R2').value:
            planilha_dados.range('R2').value = conta_dv
        # if not planilha_dados.range('W17').value:
        #     planilha_dados.range('W17').value = info.get('VALOR DE PAGTO BRUTO', '')
        # if not planilha_dados.range('AB17').value:
        #     planilha_dados.range('AB17').value = info.get('VALOR PGTO LÍQUIDO', '')

        # planilha_dados.range(f'K{i}').value = info.get('IR', '')
        # planilha_dados.range(f'L{i}').value = info.get('PIS', '')
        # planilha_dados.range(f'M{i}').value = info.get('COFINS', '')
        # planilha_dados.range(f'N{i}').value = info.get('CSLL', '')
        # planilha_dados.range(f'O{i}').value = info.get('ISS', '')

    linha_excel = 22
    for row in tables_data:
    # Ignorar linhas inválidas
        if (
            not row or
            row[0] is None or row[0] == '' or
            any(isinstance(cell, str) and any(palavra in cell.upper() for palavra in ['APRESENTADO', 'GLOSADO', 'RECURSO']) for cell in row)):
            continue

        # Escreve os dados na planilha
        planilha_dados.range(f'V{linha_excel}').value = str(row[0])
        planilha_dados.range(f'W{linha_excel}').value = extrair_valor(row[1]) if len(row) > 1 else 0.0
        planilha_dados.range(f'X{linha_excel}').value = extrair_valor(row[2]) if len(row) > 2 else 0.0
        planilha_dados.range(f'Y{linha_excel}').value = extrair_valor(row[3]) if len(row) > 3 else 0.0
        planilha_dados.range(f'Z{linha_excel}').value = extrair_valor(row[4]) if len(row) > 4 else 0.0

        linha_excel += 1 # Só incrementa se a linha for válida


    update_bank_info(planilha_dados, planilha_bancos)

    work_book.save()


def conne(page):

    work_book, planilha = obter_planilha_e_aba('Dados')

    #define a ultima linha da planilha "NE"
    ultima_linha_NE = planilha['nota_de_empenho'].end('down').row

    for linha in range(2,ultima_linha_NE+1):
        #consulta a NE
        nota_de_empenho_planilha = planilha['nota_de_empenho'][linha-1].value
        nota_de_empenho = str(nota_de_empenho_planilha)
        num_nota_de_empenho = nota_de_empenho[-6:]
        ano_nota_de_empenho = nota_de_empenho[:4]
        page.locator("[id=\"frmMenu\\:acessoRapido\"]").fill("conne")
        page.locator("[id=\"frmMenu\\:acessoRapido\"]").press("Enter")
        #detalhamento NE
        page.frame_locator("iframe[title=\"SIAFI - Sistema Integrado de Administração Financeira do Governo Federal\"]").locator("input[id='ano']").fill(ano_nota_de_empenho)
        page.frame_locator("iframe[title=\"SIAFI - Sistema Integrado de Administração Financeira do Governo Federal\"]").locator("#numeroNE").get_by_role("textbox").fill(num_nota_de_empenho)
        page.frame_locator("iframe[title=\"SIAFI - Sistema Integrado de Administração Financeira do Governo Federal\"]").locator("#numeroNE").get_by_role("textbox").press("Enter")
        #page.frame_locator("iframe[title=\"SIAFI - Sistema Integrado de Administração Financeira do Governo Federal\"]").locator(":text-is('Pesquisar')").click()
        #extrair informações da NE
        localizador_natureza_da_despesa = page.frame_locator("iframe[title=\"SIAFI - Sistema Integrado de Administração Financeira do Governo Federal\"]").locator("div:below(span:text-is('Lista de Itens')) >> nth = 4").text_content()
        natureza_da_despesa = localizador_natureza_da_despesa[:6]
        planilha['natureza_da_despesa'][linha-1].value = natureza_da_despesa
        localizador_subelemento = page.frame_locator("iframe[title=\"SIAFI - Sistema Integrado de Administração Financeira do Governo Federal\"]").locator("span:has-text('Subelemento')>>nth=0").text_content()
        subelemento = localizador_subelemento[12:14]
        nome_NDD = localizador_subelemento[17:]
        planilha['subelemento'][linha-1].value = subelemento
        planilha['nome_NDD'][linha-1].value = nome_NDD
        cadastro = page.frame_locator("iframe[title=\"SIAFI - Sistema Integrado de Administração Financeira do Governo Federal\"]").locator("p.ng-star-inserted >> nth=8").text_content()
        planilha['cadastro'][linha-1].value = cadastro
        nome_favorecido = page.frame_locator("iframe[title=\"SIAFI - Sistema Integrado de Administração Financeira do Governo Federal\"]").locator("p:near(:text-is('Nome do Favorecido'))>>nth=0").text_content()
        planilha['nome_favorecido'][linha-1].value = nome_favorecido
        descricao = page.frame_locator("iframe[title=\"SIAFI - Sistema Integrado de Administração Financeira do Governo Federal\"]").locator("div:below(:text-is('Descrição:')) >> nth = 0").text_content()
        planilha['descricao'][linha-1].value = descricao

    # Salvar o arquivo Excel
    work_book.save()

    show_alert()


def liquidacao_pro_saude(page):
   
    work_book, planilha = obter_planilha_e_aba('Dados')

    #define a ultima linha da planilha
    #ultima_linha = planilha['emitente'].end('down').row
    ultima_linha_NE = work_book.sheets['NE']['nota_de_empenho'].end('down').row
    #pede ao usuário para inserir o valor da linha inicial
    linha_inicial = pyautogui.prompt(text='Insira o valor da linha inicial:', title='Linha Inicial', default='2')
    linha_inicial = int(linha_inicial)
    #pede ao usuário para inserir o valor da linha final
    linha_final = pyautogui.prompt(text='Insira o valor da linha final:', title='Linha Final',default=linha_inicial)
    linha_final = int(linha_final)

    #define NE e subelemento
    #nota_de_empenho = str(work_book.sheets['NE']['nota_de_empenho'][ultima_linha_NE-1].value)
    nota_de_empenho = str(planilha.range('D15').value)
    #subelemento = str(int(work_book.sheets['NE']['subelemento'][ultima_linha_NE-1].value))
    subelemento = str(int(planilha.range('F15').value))

    for linha in range(linha_inicial,linha_final+1):
       
        #define datas
        data_do_dia = datetime.now().strftime('%d/%m/%Y')
        data_de_vencimento_planilha = planilha.range('B2').value
        data_de_vencimento = data_de_vencimento_planilha.strftime('%d/%m/%Y')
        ateste_planilha = planilha.range('B3').value
        ateste = ateste_planilha.strftime('%d/%m/%Y')
        data_de_emissao_planilha = planilha['data_de_emissao'][linha-1].value
        data_de_emissao = data_de_emissao_planilha#.strftime('%d/%m/%Y')
        ano_emissao = data_de_emissao[-4:]
        mes_emissao = data_de_emissao[3:5]
        data_vencimento_DARF_planilha = planilha.range('B4').value
        data_vencimento_DARF = data_vencimento_DARF_planilha.strftime('%d/%m/%Y')
        data_vencimento_DAR_planilha = planilha.range('B5').value
        data_vencimento_DAR = data_vencimento_DAR_planilha.strftime('%d/%m/%Y')
       
        #define cada coluna da tabela
        processo_edoc = planilha.range('B1').value
        emitente = re.sub(r'[ ./-]','', planilha['emitente'][linha-1].value)
        try:
            num_doc_fiscal = str(int(planilha['num_doc_fiscal'][linha-1].value))
        except:
            num_doc_fiscal = str(planilha['num_doc_fiscal'][linha-1].value)
        valor_doc_fiscal_planilha = planilha['valor_doc_fiscal'][linha-1].value
        # Verifica se o valor é uma string
        if isinstance(valor_doc_fiscal_planilha, str):
            # Remove o símbolo de moeda e os separadores de milhar
            valor_doc_fiscal_planilha_clean = re.sub(r'[^\d,]', '', valor_doc_fiscal_planilha)
            # Substitui a vírgula decimal por um ponto
            valor_doc_fiscal_planilha_clean = valor_doc_fiscal_planilha_clean.replace(',', '.')
            # Converte para float
            valor_doc_fiscal_planilha_float = float(valor_doc_fiscal_planilha_clean)
        else:
            # Se já for um float, não precisa fazer nada
            valor_doc_fiscal_planilha_float = valor_doc_fiscal_planilha
        valor_doc_fiscal = f"{valor_doc_fiscal_planilha_float:.2f}"
        glosa_planilha = planilha['glosa'][linha-1].value
        formatted_glosa = "{:,.2f}".format(glosa_planilha)
        glosa = f"{formatted_glosa.replace(',', 'X').replace('.', ',').replace('X', '.')}"
        valor_PCO_planilha = planilha['valor_PCO'][linha-1].value
        valor_PCO = f"{valor_PCO_planilha:.2f}"
        natureza_rendimento = planilha['nat_rendimento_DARF'][linha-1].value
        valor_DARF_planilha = planilha['valor_DARF'][linha-1].value
        valor_DARF = f"{valor_DARF_planilha:.2f}"
        valor_DAR_planilha = planilha['valor_DAR'][linha-1].value
        valor_DAR = f"{valor_DAR_planilha:.2f}"
        banco = str(planilha['banco'][linha-1].value)
        agencia = str(planilha['agencia'][linha-1].value)
        conta_planilha = planilha['conta'][linha-1].value
        conta = re.sub(r'[ ./-]','', str(conta_planilha))
        
        #incdh
        page.locator("[id=\"frmMenu\\:acessoRapido\"]").fill("incdh")
        page.locator("[id=\"frmMenu\\:acessoRapido\"]").press("Enter")
        page.wait_for_timeout(1000)
        page.locator("[id$='codigoTipoDocHabil_input']").fill("NP")
        page.wait_for_timeout(500)
        page.locator("#form_manterDocumentoHabil\\:btnConfirmarTipoDoc").click()
        page.wait_for_timeout(500)

        #DADOS BÁSICOS
        page.locator("input[id='form_manterDocumentoHabil:dataVencimento_calendarInputDate']").fill(data_de_vencimento)
        page.wait_for_timeout(500)
        page.locator("input[id='form_manterDocumentoHabil:processo_input']").fill(processo_edoc)
        page.wait_for_timeout(500)
        page.locator("input[id='form_manterDocumentoHabil:dataAteste_calendarInputDate']").fill(ateste)
        page.wait_for_timeout(500)
        page.locator("input[id='form_manterDocumentoHabil:valorPrincipalDocumento_input']").click()
        page.locator("input[id='form_manterDocumentoHabil:valorPrincipalDocumento_input']").press_sequentially(valor_doc_fiscal)
        page.wait_for_timeout(500)
        page.locator("input[id='form_manterDocumentoHabil:credorDevedor_input']").fill(emitente)

        #Inclui Dados de Documentos de Origem
        page.locator("input[id='form_manterDocumentoHabil:tableDocsOrigem_painel_incluir']").click()
        page.wait_for_timeout(1000)
        #page.locator("input[id='form_manterDocumentoHabil:tableDocsOrigem:0:emitenteDocOrigem_input']").fill(emitente)
        page.wait_for_timeout(500)
        page.locator("input[id='form_manterDocumentoHabil:tableDocsOrigem:0:dataEmissaoDocOrigem_calendarInputDate']").fill(data_de_emissao)
        page.wait_for_timeout(500)
        page.locator("input[id='form_manterDocumentoHabil:tableDocsOrigem:0:numeroDocOrigem_input']").fill(num_doc_fiscal)
        page.wait_for_timeout(500)
        page.locator("input[id='form_manterDocumentoHabil:tableDocsOrigem:0:valorDocOrigem_input']").click()
        page.wait_for_timeout(1000)
        page.locator("input[id='form_manterDocumentoHabil:tableDocsOrigem:0:valorDocOrigem_input']").press_sequentially(valor_doc_fiscal)
        page.wait_for_timeout(500)
        #indica na planilha o nome do credor
        nome_emitente_siafi = page.locator("span[id='form_manterDocumentoHabil:nomeCredorDevedor']").text_content()
        nome_emitente = re.sub(r'[¬<&>"\';=%#]', '', nome_emitente_siafi)
        planilha['nome_emitente'][linha-1].value = nome_emitente 

        #define observações
        descricao = work_book.sheets['NE']['descricao'][ultima_linha_NE-1].value 
        codigo_recolhimento_DARF = planilha['codigo_recolhimento_DARF'][linha-1].value
        print(f'código_recolhimento_DARF: {codigo_recolhimento_DARF}')
        # if codigo_recolhimento_DARF and str(codigo_recolhimento_DARF).isdigit():
        #     codigo_recolhimento_DARF= int(codigo_recolhimento_DARF)
        if codigo_recolhimento_DARF == 6147:
            codigo_recolhimento_DARF = int(codigo_recolhimento_DARF)
            obs_DARF = f'RET. RFB 5,85PC - IN 1234/2012'
        elif codigo_recolhimento_DARF == 6190:
            codigo_recolhimento_DARF = int(codigo_recolhimento_DARF)
            obs_DARF = f'RET. RFB 9,45PC - IN 1234/2012'
        elif codigo_recolhimento_DARF == "ISENTO":
            obs_DARF = f'RET. RFB ISENTO'
        elif codigo_recolhimento_DARF == "IMUNE":
            obs_DARF = f'RET. RFB IMUNE'
        elif codigo_recolhimento_DARF == "SIMPLES NACIONAL":
            obs_DARF = f'RET. RFB OPT. SIMPLES NACIONAL'
        codigo_recolhimento_DARF_str = str(codigo_recolhimento_DARF)
        print(f'código_recolhimento_DARF: {codigo_recolhimento_DARF_str}')
        codigo_receita_DAR = planilha['codigo_receita_DAR'][linha-1].value
        print(f'código_receita_DAR: {codigo_receita_DAR}')
        if codigo_receita_DAR == 1782:
            codigo_receita_DAR = int(codigo_receita_DAR)
            obs_DAR = f'RET. ISS 2PC' 
        elif codigo_receita_DAR == "1782":
            codigo_receita_DAR = int(codigo_receita_DAR)
            obs_DAR = f'RET. ISS 2PC'
        elif codigo_receita_DAR == "ISENTO":
            obs_DAR = f'RET. ISS ISENTO'
        elif codigo_receita_DAR == "IMUNE":
            obs_DAR = f'RET. ISS IMUNE'  
        elif codigo_receita_DAR == "OUTRO MUNICÍPIO":
            obs_DAR = f'SERVIÇO PRESTADO EM OUTRO MUNICÍPIO'  
        elif codigo_receita_DAR == "SIMPLES NACIONAL":
            obs_DAR = f'RET. ISS CONFORME DESTACADO NO DOC. FISCAL'
        elif codigo_receita_DAR == "UNIPROFISSIONAL":
            obs_DAR = f'RET. ISS NÃO SE APLICA (SOC. UNIPROFISSIONAL)'
        observacao_DARF = str(f'NFSE {num_doc_fiscal} - {descricao} - {obs_DARF}').upper()
        observacao_DAR = str(f'NFSE {num_doc_fiscal} - {descricao} - {obs_DAR}').upper()
        observacao = str(f'NFSE {num_doc_fiscal} - {descricao} - PERIODO: {mes_emissao}/{ano_emissao} - {obs_DARF} - {obs_DAR} - EDOC: {processo_edoc}.').upper() 
        if glosa_planilha > 0:
            if (valor_PCO_planilha-valor_DAR_planilha-valor_DARF_planilha)> 0.02:
                observacao = str(f'NFSE {num_doc_fiscal} - {descricao} - GLOSA: R$ {glosa} - PERIODO: {mes_emissao}/{ano_emissao} - {obs_DARF} - {obs_DAR} - EDOC: {processo_edoc}.').upper()
            else:
                observacao = str(f'NFSE {num_doc_fiscal} - {descricao} - APENAS RETENÇÃO DE TRIBUTOS, PARA POSTERIOR PAGAMENTO DO VALOR LÍQUIDO COM RECURSOS PRÓPRIOS DO PRÓ-SAÚDE - PERIODO: {mes_emissao}/{ano_emissao} - {obs_DARF} - {obs_DAR} - EDOC: {processo_edoc}.').upper()
        nome_rascunho = str(f'{num_doc_fiscal} {processo_edoc}').upper() 

        #Inclui observação
        page.locator("textarea[id='form_manterDocumentoHabil:observacao']").fill(observacao)
        planilha['observacao'][linha-1].value = observacao
        #Confirma Dados Básicos
        page.locator("input[id='form_manterDocumentoHabil:btnConfirmarDadosBasicos']").click()
        page.wait_for_timeout(1000)
        
        #PRINCIPAL COM ORÇAMENTO
        page.locator("input[id='form_manterDocumentoHabil:abaPrincipalComOrcamentoId']").click()
        page.wait_for_timeout(1000)
        #Inclui Situação
        page.locator("#form_manterDocumentoHabil\\:campo_situacao_input").fill("DSP001")
        page.locator("input[id='form_manterDocumentoHabil:botao_ConfirmarSituacao']").click()
        page.wait_for_timeout(1000)
        #Preenche Situação DSP001
        #page.locator("select[class*='x-smal']").select_option(value='true')
        page.locator("input[class*='Número do Empenho']").fill(nota_de_empenho)
        page.wait_for_timeout(500)
        page.locator("input[class*='Subelemento']").click()
        page.locator("input[class*='Subelemento']").fill(subelemento)
        page.wait_for_timeout(500)
        page.locator("select[class*='Liquidado?']").select_option(value='true')
        page.wait_for_timeout(500)
        page.locator("input[class*='Conta Variação Patrimonial Diminutiva']").click()
        page.locator("input[class*='Conta Variação Patrimonial Diminutiva']").fill("3.3.2.3.1.01.00")
        page.wait_for_timeout(500)
        page.locator("input[class*='Contas a Pagar']").click()
        page.locator("input[class*='Contas a Pagar']").fill("2.1.3.1.1.04.00")
        page.wait_for_timeout(500)
        page.locator("input[class*='Valor']").click()
        page.locator("input[class*='Valor']").press_sequentially(valor_PCO)
        page.wait_for_timeout(500)
        #Confirma Situação
        page.locator("input[id='form_manterDocumentoHabil:lista_PCO_painel_confirmar']").click()
        page.wait_for_timeout(1000)

        if valor_DARF_planilha > 0:
            #DEDUÇÃO
            page.locator("input[id='form_manterDocumentoHabil:abaDeducaoId']").click()
            page.wait_for_timeout(1000)
            #inclui situação da dedução DDF025  
            page.locator("input[id='form_manterDocumentoHabil:campo_situacao_input']").fill("DDF025")
            page.locator("input[id='form_manterDocumentoHabil:botao_ConfirmarSituacao']").click()
            page.wait_for_timeout(1000)
            #preenche dados da situação da dedução
            page.locator("input[class*='Data de Vencimento']").fill(data_vencimento_DARF) 
            page.wait_for_timeout(500)
            page.locator("input[class*='Data de Pagamento']").fill(data_vencimento_DARF) 
            page.wait_for_timeout(500)
            page.locator("input[id='form_manterDocumentoHabil:listaDeducao:0:inputValorDeducao_input']").click()
            page.locator("input[id='form_manterDocumentoHabil:listaDeducao:0:inputValorDeducao_input']").press_sequentially(valor_DARF)
            page.wait_for_timeout(500)
            page.locator("input[id='form_manterDocumentoHabil:listaDeducao:0:campoInscricaoA_deducao_input_darf_input']").fill(codigo_recolhimento_DARF_str)
            page.locator("input[id='form_manterDocumentoHabil:listaDeducao:0:campoInscricaoB_deducao_input_inscr_input']").fill(natureza_rendimento)
            page.locator("input[id='form_manterDocumentoHabil:listaDeducao:0:deducao_tabela_recolhimentos_painel_incluir']").click()
            page.wait_for_timeout(500)
            page.locator("input[id='form_manterDocumentoHabil:listaDeducao:0:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_recolhedor_input']").fill(emitente)
            page.wait_for_timeout(500)
            page.locator("input[id='form_manterDocumentoHabil:listaDeducao:0:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_base_calculo_input']").click()
            page.locator("input[id='form_manterDocumentoHabil:listaDeducao:0:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_base_calculo_input']").press_sequentially(valor_doc_fiscal)
            page.wait_for_timeout(500)
            page.locator("input[id='form_manterDocumentoHabil:listaDeducao:0:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_valor_recolhimento_input']").click()
            page.locator("input[id='form_manterDocumentoHabil:listaDeducao:0:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_valor_recolhimento_input']").press_sequentially(valor_DARF)
            page.wait_for_timeout(500)
            #Confirma a Dedução DDF025
            page.locator("input[id='form_manterDocumentoHabil:listaDeducao_painel_confirmar']").click()
            page.wait_for_timeout(1000)
            #inclui Pré-Doc DDF025
            page.locator("input[class='botaoPredoc']").click()
            page.wait_for_timeout(1000)
            #Preenche Pré-Doc
            page.locator("select[class*='Recurso']").select_option(value='VINCULACAO_PAGAMENTO')
            page.wait_for_timeout(500)
            page.locator("input[class*='Período de Apuração']").fill(data_do_dia)
            page.wait_for_timeout(500)
            page.locator("input[id='form_manterDocumentoHabil:predocDarfModal_vinculacao_input']").fill("499")
            page.locator("textarea[id='form_manterDocumentoHabil:observacaoPredoc']").fill(observacao_DARF)
            page.wait_for_timeout(500)
            #Confirma o Pré-Doc
            page.locator("input[id='form_manterDocumentoHabil:btnConfirmarPredoc']").click()
            page.wait_for_timeout(1000)
        
        if valor_DAR_planilha > 0:
            if valor_DARF_planilha == 0:
                #DEDUÇÃO
                page.locator("input[id='form_manterDocumentoHabil:abaDeducaoId']").click()
                page.wait_for_timeout(1000)
            else:
                #Inclui nova Situação
                page.locator("input[id='form_manterDocumentoHabil:listaDeducao_painel_incluir']").click()
                page.wait_for_timeout(1000)
            #Inclui situação da dedução  DDR001
            page.locator("input[id='form_manterDocumentoHabil:campo_situacao_input']").fill("DDR001")
            page.locator("input[id='form_manterDocumentoHabil:botao_ConfirmarSituacao']").click()
            page.wait_for_timeout(1000)
            #preenche dados da situação da dedução
            page.locator("input[class*='Data de Vencimento']").fill(data_vencimento_DAR) 
            page.wait_for_timeout(500)
            page.locator("input[class*='Data de Pagamento']").fill(data_vencimento_DAR) 
            page.wait_for_timeout(500)
            page.locator("input[class*='Código do Município']").fill("9701") 
            page.locator("input[class*='Código de Receita']").fill("1782") 
            page.locator("input[id*='inputValorDeducao_input']").click()
            page.locator("input[id*='inputValorDeducao_input']").press_sequentially(valor_DAR)
            page.wait_for_timeout(500)
            page.locator("input[id*='deducao_tabela_recolhimentos_painel_incluir']").click()
            page.wait_for_timeout(1000)
            page.locator("input[id*='deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_recolhedor_input']").fill(emitente)
            page.wait_for_timeout(500)
            page.locator("input[class*='Valor Principal']").click()
            page.locator("input[class*='Valor Principal']").press_sequentially(valor_DAR) 
            page.wait_for_timeout(500)
            page.locator("input[id*='deducao_tabela_recolhimentos_painel_confirmar']").click()
            page.wait_for_timeout(1000)
            #Confirma a Dedução DDR001
            page.locator("input[id='form_manterDocumentoHabil:listaDeducao_painel_confirmar']").click()
            page.wait_for_timeout(1000)

            #inclui Pré-Doc DDR001
            page.locator(".botaoPredoc >> nth=-1").click()
            page.wait_for_timeout(1000)
            #Preenche Pré-Doc
            page.locator("select[class*='Tipo de Recurso']").select_option(value='VINCULACAO_PAGAMENTO')
            page.select_option("select[class*='Mês de Referência']", value=mes_emissao)
            page.select_option("select[class*='Ano de Referência']", value=ano_emissao)
            page.locator("input[id='form_manterDocumentoHabil:ugTomadoraDar_input']").fill("010001")
            page.locator("input[id='form_manterDocumentoHabil:municipioNFDar_input']").fill("9701")
            page.wait_for_timeout(500)
            page.locator("input[id='form_manterDocumentoHabil:nfRecibo_input']").fill(num_doc_fiscal)
            page.locator("input[class*='Data de Emissão da NF']").fill(data_de_emissao) 
            page.locator("input[id='form_manterDocumentoHabil:aliquotaNFDar_input']").click()
            page.locator("input[id='form_manterDocumentoHabil:aliquotaNFDar_input']").press_sequentially("2000")
            page.locator("input[id='form_manterDocumentoHabil:valorNfDar_input']").click()
            page.locator("input[id='form_manterDocumentoHabil:valorNfDar_input']").press_sequentially(valor_doc_fiscal) 
            page.wait_for_timeout(500)
            page.locator("textarea[id='form_manterDocumentoHabil:observacaoPredoc']").fill(observacao_DAR)
            page.wait_for_timeout(500)
            #Confirma o Pré-Doc
            page.locator("input[id='form_manterDocumentoHabil:btnConfirmarPredoc']").click()
            page.wait_for_timeout(1000)

        #CENTRO DE CUSTOS
        # Mapeamento dos números dos meses para os nomes dos meses
        meses = {"01": "Jan", "02": "Fev", "03": "Mar", "04": "Abr", "05": "Mai", "06": "Jun", "07": "Jul", "08": "Ago", "09": "Set", "10": "Out", "11": "Nov", "12": "Dez"}
        # Função para selecionar o mês
        def selecionar_mes(page, mes_num):
            mes_nome = meses[mes_num]
            page.locator(f"a.ui-monthpicker-month:has-text('{mes_nome}')").click()
        # Clica na aba "Detacustos"
        page.locator("input[id='form_manterDocumentoHabil:abaDetacustos']").click()
        page.wait_for_timeout(500)
        # Define Período de Referência
        page.locator("input[class='ng-tns-c48-21 ui-inputtext ui-widget ui-state-default ui-corner-all ng-star-inserted']").click()
        page.locator("select.ui-datepicker-year").select_option(ano_emissao)
        selecionar_mes(page, mes_emissao)
        page.wait_for_timeout(500)
        # Confirma o Centro de Custos
        page.locator("button[icon='pi pi-check'][ptooltip='Confirmar']").click()
        page.wait_for_timeout(500)
        
        # # page.locator("input[id='form_manterDocumentoHabil:abaDetacustos']").click()
        # # page.wait_for_timeout(500)
        # # # page.locator("input[id='form_manterDocumentoHabil:consolidado_dataTable:consolidado_checkAll']").check()
        # # # page.wait_for_timeout(500)
        # # per_ref = mes_emissao + "/" + ano_emissao
        # # page.locator("input[class='ng-tns-c48-21 ui-inputtext ui-widget ui-state-default ui-corner-all ng-star-inserted']").fill(per_ref)
        # # # page.locator("input[class*='Ano Referência']").fill(ano_emissao)
        # # # page.locator("input[id='form_manterDocumentoHabil:btnIncluirNovoVinculoCentroCusto']").click()
        # # page.locator("input[id='form_manterDocumentoHabil:abaDetacustos']").click()
        # # page.wait_for_timeout(1000)
        # Caso não haja valor líquido:
         
        #DADOS DE PAGAMENTO
        page.locator("input[id='form_manterDocumentoHabil:abaDadosPagRecId']").click()
        page.wait_for_timeout(1000)
            # Caso não haja valor líquido:
        if page.locator(":text-is('(ER0288) Não há itens liquidados com valor disponível para informar Dados de Pagamento/Recebimento.')").count() > 0:
            pass
        else:
            page.locator("input[id='form_manterDocumentoHabil:lista_DPgtoOB_painel_incluir']").click()
            page.wait_for_timeout(1000)
            page.locator("input[id='form_manterDocumentoHabil:lista_DPgtoOB_painel_confirmar']").click()
            page.wait_for_timeout(1000)
            #inclui Pré-Doc de Pagamento
            page.locator(".botaoPredoc").click()
            page.wait_for_timeout(1000)
            page.locator("select[id='form_manterDocumentoHabil:tipoOBSelect']").select_option(value='OB_CREDITO')
            page.wait_for_timeout(500)
            #inclui observação no Pré-Doc
            page.locator("textarea[id='form_manterDocumentoHabil:observacaoPredoc']").fill(observacao)
            page.wait_for_timeout(500)
            #digita dados bancários
            page.locator("input[id='form_manterDocumentoHabil:favorecido_banco_input']").fill(banco)
            page.locator("input[id='form_manterDocumentoHabil:favorecido_agencia_input']").fill(agencia)
            page.locator("input[id='form_manterDocumentoHabil:favorecido_conta_input']").fill(conta)
            #confirma o Pré-Doc
            page.locator("input[id='form_manterDocumentoHabil:btnConfirmarPredoc']").click()
            page.wait_for_timeout(500)        
            # no caso de erro
            if page.locator(":text-is('(0006) DOMICILIO BANCARIO DO CREDOR INEXISTENTE')").is_visible() == True:
                planilha['erro'][linha-1].value = "DADOS BANCÁRIOS NÃO ENCONTRADOS"
                #retorna o Pré-Doc
                page.locator("input[id='form_manterDocumentoHabil:btnRetornarPredoc']").click()
                page.wait_for_timeout(1000)
        #
        #salva rascunho
        page.locator("input[id='form_manterDocumentoHabil:salvarRascunho_botao']").click()
        page.locator("input[class*='Nome do rascunho']").fill(f"{nome_rascunho}")
        page.locator("input[id='formRascunho:btnConfirmar']").click()
        if page.locator(":text-is('(AT0002) Rascunho já existe! Deseja sobrescrever??')").is_visible() == True:
            page.locator("input[id='formRascunho:btnConfirmar']").click()
        page.wait_for_timeout(1000)
        #
        #registra a NS
        page.locator("input[id='form_manterDocumentoHabil:btnRegistrar']").click()
        #retorna o número da NS gerada
        expect(page.locator("a[id='form_manterDocumentoHabil:tableNsGeradas:0:acessoCondoc_lnkConsultaDoc']")).to_be_visible(timeout=60000)
        nota_de_sistema_ug = page.locator("a[id='form_manterDocumentoHabil:tableNsGeradas:0:acessoCondoc_lnkConsultaDoc']").text_content()
        nota_de_sistema = nota_de_sistema_ug[-12:]
        planilha['nota_de_sistema'][linha-1].value = nota_de_sistema 
        page.locator("input[id='form_manterDocumentoHabil:btnRetornarResultadoRegistrar']").click()
        page.wait_for_timeout(1000)

    show_alert()
    #page.locator("a[title='Sair do sistema']").click()

def inclui_NS_edoc(page):    
        
        work_book = xlwings.Book(f'{diretorio}\\{arquivo}.xlsm')
        planilha=work_book.sheets['Dados']       

        # Ponto= work_book.sheets['Dados'].api.OLEObjects("Ponto").Object.Value
        # Senha= work_book.sheets['Dados'].api.OLEObjects("Senha").Object.Value

        ultima_linha = planilha['emitente'].end('down').row
        processo_edoc = planilha.range('B1').value
        #pede ao usuário para inserir o valor da linha inicial
        linha_inicial = pyautogui.prompt(text='Insira o valor da linha inicial:', title='Linha Inicial', default='2')
        linha_inicial = int(linha_inicial)
        #pede ao usuário para inserir o valor da linha final
        linha_final = pyautogui.prompt(text='Insira o valor da linha final:', title='Linha Final',default=linha_inicial)
        linha_final = int(linha_final)

        
        #pesquisa processo
        page.locator("input[placeholder='Pesquisa de protocolo número/ano']").click()
        page.locator("input[placeholder='Pesquisa de protocolo número/ano']").fill(processo_edoc)
        page.locator("input[placeholder='Pesquisa de protocolo número/ano']").press("Enter")
        page.wait_for_timeout(1000)
        #clica no processo procurado
        page.locator(f"a:has-text('{processo_edoc}')").click(timeout=60000)
        page.wait_for_timeout(1000)

        for linha in range(linha_inicial,linha_final+1):

            nota_de_sistema = planilha['nota_de_sistema'][linha-1].value
            nome_arquivo = f"{nota_de_sistema}.pdf"

            #inclui documento
            page.locator(":text-is('Incluir')").click(timeout=90000)
            page.locator(":text-is('Criar usando arquivo do computador...')").click(timeout=60000)      
            page.wait_for_timeout(1000)   
            page.locator("input:right-of(span:text-is('Título')) >> nth=1").fill(nota_de_sistema)  
            #upload
            page.locator("label[class='labelInputFile']").set_input_files(f"{diretorio}\\{nome_arquivo}")  
            page.locator(":text-is('Salvar')").click(timeout=60000)
            # #assina odocumento
            # page.locator("[value='Assinar/Autenticar']").click(timeout=60000)
            # page.locator("[id*='txtPonto']").fill(Ponto)
            # page.locator("[id*='txtSenha']").fill(Senha)
            # #expect(page.locator("[id*='txtCargo']")).not_to_be_empty()
            # page.wait_for_timeout(1000)
            # page.locator("[value='Assinar']").click(timeout=60000)     
            # page.wait_for_timeout(2000)  
            
            #exclui a NS
            caminho_arquivo = os.path.join(diretorio, nome_arquivo)
            #os.path.exists(caminho_arquivo)
            os.remove(caminho_arquivo)
        
        #Salvar o arquivo Excel
        work_book.save()

        #Inclusão do relatório de retenções
        processo_formatado = processo_edoc.replace("/","-")
        print(processo_formatado)
        nome_relatorio = f"Relatório Retenções Processo {processo_formatado}.pdf"
        caminho_relatorio = os.path.join(diretorio, nome_relatorio)
        print(caminho_relatorio)
        if os.path.exists(caminho_relatorio):
            #inclui documento
            page.locator(":text-is('Incluir')").click(timeout=90000)
            page.locator(":text-is('Criar usando arquivo do computador...')").click(timeout=60000)      
            page.wait_for_timeout(1000)   
            page.locator("input:right-of(span:text-is('Título')) >> nth=1").fill("Relatório de retenções")  
            #upload
            page.locator("label[class='labelInputFile']").set_input_files(f"{diretorio}\\{nome_relatorio}")  
            page.locator(":text-is('Salvar')").click(timeout=60000)


        
        show_alert()


def conferir_liquidação(page):
    # sistema = page.locator("div[class='links']").text_content()
    # print(sistema)
    # pyautogui.alert(f'{sistema}')
    diretorio_arquivo = f"{diretorio}\\{arquivo}.xlsm"
    work_book = xlwings.Book(diretorio_arquivo)
    planilha = work_book.sheets['Dados']

    ano_atual = datetime.now().year
    
    # plan_empenho=work_book.sheets['NE']
    # nota_empenho=plan_empenho['nota_de_empenho'][1].value
    # descricao=plan_empenho['descricao'][1].value
    # nota_empenho=plan_empenho['nota_de_empenho'][1].value
    # linha_inicial = pyautogui.prompt(text='Insira o valor da linha inicial:', title='Linha Inicial', default='2')
    # linha_inicial = int(linha_inicial)
    # linha_final = pyautogui.prompt(text='Insira o valor da linha final:', title='Linha Final',default=linha_inicial)
    # linha_final = int(linha_final)

    # Percorre as linhas de 35 a 45
    for linha in range(35, 45):
        try:
            # Acessa o valor da coluna A (nota_de_sistema) na linha especificada
            nota_de_sistema_planilha = planilha.range(f'A{linha}').value
            nota_de_sistema = str(int(nota_de_sistema_planilha)) if nota_de_sistema_planilha and isinstance(nota_de_sistema_planilha, float) else str(nota_de_sistema_planilha) if nota_de_sistema_planilha else None
            
            # Acessa o valor da coluna B (nota_de_pagamento) na linha especificada
            nota_de_pagamento_planilha = planilha.range(f'B{linha}').value
            nota_de_pagamento = str(int(nota_de_pagamento_planilha)) if nota_de_pagamento_planilha and isinstance(nota_de_pagamento_planilha, float) else str(nota_de_pagamento_planilha) if nota_de_pagamento_planilha else None
            
            # Inicializa variáveis
            num_nota = None
            ano_nota = None
            
            # Verifica se a nota_de_sistema contém o ano
            if nota_de_sistema and len(nota_de_sistema) > 5:
                num_nota_de_sistema = nota_de_sistema[-6:]
                ano_nota_de_sistema = nota_de_sistema[:4]
            elif nota_de_sistema:
                num_nota_de_sistema = nota_de_sistema
                ano_nota_de_sistema = str(ano_atual)
            
            # Verifica se a nota_de_pagamento contém o ano
            if nota_de_pagamento and len(nota_de_pagamento) > 5:
                num_nota_de_pagamento = nota_de_pagamento[-6:]
                ano_nota_de_pagamento = nota_de_pagamento[:4]
            elif nota_de_pagamento:
                num_nota_de_pagamento = nota_de_pagamento
                ano_nota_de_pagamento = str(ano_atual)
            
            # Extração das informações do siafi
            page.locator("input[id='frmMenu:acessoRapido']").fill("condh")
            page.locator("input:right-of(input[id='frmMenu:acessoRapido'])").click()
            page.wait_for_timeout(1000)

            # Escolhe qual nota usar (priorizando nota_de_sistema se disponível)
            if nota_de_sistema:
                num_nota = num_nota_de_sistema.replace('N', '').replace('S', '')
                ano_nota = ano_nota_de_sistema
                # Seleciona o radio button
                page.click("input[name='formConsultarDH:rbtTipoConsulta'][value='false']")
                # Seleciona a opção "NS" no dropdown
                page.select_option("select[name='formConsultarDH:tipoDS']", "NS")
                # Preenche o número e o ano
                page.fill("input[name='formConsultarDH:numeroDS_input']", num_nota)
                page.fill("input[name='formConsultarDH:anoDS_input']", ano_nota)
                # Clica em Pesquisar
                page.locator("input:text-is('Pesquisar')>>nth=0").click()
                page.wait_for_timeout(1000)
                if page.locator("span.rich-messages-label:has-text('(AT0053) Não foi encontrado nenhum registro para o filtro selecionado.')").nth(0).is_visible():
                    planilha.range(f'B{linha}').value = "Não encontrado"
                    continue
                else:
                    nota_de_pagamento = page.locator("a[id='formConsultarDH:elementList:0:lnkDetalharConsulta']").text_content()
                    planilha.range(f'B{linha}').value = nota_de_pagamento
                    page.locator("a[id='formConsultarDH:elementList:0:lnkDetalharConsulta']").click()
                    page.wait_for_timeout(1000)
            
            elif nota_de_pagamento:
                num_nota = num_nota_de_pagamento
                ano_nota = ano_nota_de_pagamento
                page.locator("input[id='formConsultarDH:tipoDH_input']").fill('NP')
                page.locator("input:below(:text-is('Número'))>>nth=0").fill(num_nota)
                page.locator("input[id='formConsultarDH:anoDH_input']").fill(ano_nota)
                # Clica em Pesquisar
                page.locator("input:text-is('Pesquisar')>>nth=0").click()

            ## Extrai informações do cabeçalho
            data_vencimento_dh = page.locator("span[id='form_manterDocumentoHabil:dataVencimento_outputText']").text_content()
            # conferência['siafi'][2].value = data_vencimento_dh

            processo_cabeçalho = page.locator("span[id='form_manterDocumentoHabil:processo_outputText']").text_content()
            # conferência['siafi'][3].value = processo_cabeçalho

            ateste_cabeçalho = page.locator("span[id='form_manterDocumentoHabil:dataAteste_outputText']").text_content()
            # conferência['siafi'][4].value = ateste_cabeçalho

            valor_cabeçalho = page.locator("span[id='form_manterDocumentoHabil:valorPrincipalDocumento_output']").text_content()
            # valor_cabeçalho=valor_cabeçalho.replace('.','')
            # valor_cabeçalho=valor_cabeçalho.replace(',','.')
            # valor_cabeçalho=float(valor_cabeçalho)
            # conferência['siafi'][5].value = valor_cabeçalho

            CNPJ_cabeçalho = page.locator("span[id='form_manterDocumentoHabil:credorDevedor_output']").text_content()
            planilha.range(f'D{linha}').value = CNPJ_cabeçalho

            nome_credor = page.locator("span[id='form_manterDocumentoHabil:nomeCredorDevedor']").text_content()
            planilha.range(f'C{linha}').value = nome_credor


            ## Extrai informações dos Dados de Documentos de Origem
            CNPJ_origem = page.locator("span[id='form_manterDocumentoHabil:tableDocsOrigem:0:emitenteDocOrigem_output']").text_content()
            #conferência['siafi'][7].value = CNPJ_origem
            
            emissao_origem = page.locator("span[id='form_manterDocumentoHabil:tableDocsOrigem:0:dataEmissaoDocOrigem_outputText']").text_content()
            planilha.range(f'E{linha}').value = str(emissao_origem)

            numero_origem = page.locator("span[id='form_manterDocumentoHabil:tableDocsOrigem:0:numeroDocOrigem_outputText']").text_content()
            # Verifica se o valor é numérico
            if numero_origem.isdigit():
                numero_origem = int(numero_origem)
            else:
                numero_origem = numero_origem  # Mantém como string se não for numérico
            planilha.range(f'F{linha}').value = numero_origem

            valor_origem = page.locator("span[id='form_manterDocumentoHabil:tableDocsOrigem:0:valorDocOrigem_output']").text_content().replace(" ","")
            valor_origem = valor_origem.replace('.','')
            valor_origem = valor_origem.replace(',','.')        
            valor_origem = float(valor_origem)
            planilha.range(f'G{linha}').value = valor_origem

            obs_dados_básicos = page.locator("textarea[id='form_manterDocumentoHabil:observacao']").text_content()
            planilha.range(f'S{linha}').value = obs_dados_básicos

            ## Extrai informações do Principal com Orçamento
            page.locator("input:text-is('Principal Com Orçamento')").click()
            page.wait_for_timeout(2000)

            # Localiza todos os elementos com o texto 'Empenho'
            empenhos = page.locator("span[id^='form_manterDocumentoHabil:lista_PCO:'][id$=':PCO_item_num_empenho_header']").all()
            
            # Lista para armazenar os empenhos extraídos
            lista_empenhos = []
            # Itera por cada elemento encontrado
            for empenho in empenhos:
                conteudo = empenho.text_content()
                if conteudo:  # Verifica se o conteúdo não é nulo ou vazio
                    lista_empenhos.append(conteudo)
            # Junta os empenhos em uma string, separados por vírgula e espaço
            empenhos_concatenados = ', '.join(lista_empenhos)
            planilha.range(f'J{linha}').value = empenhos_concatenados

            if page.locator("span[id='form_manterDocumentoHabil:lista_PCO:0:PCO_item_campoClassificacaoA_output_classificacao_contabil']").is_visible():
                conta_vpd = page.locator("span[id='form_manterDocumentoHabil:lista_PCO:0:PCO_item_campoClassificacaoA_output_classificacao_contabil']").text_content()
                #conferência['siafi'][18].value = conta_vpd

            if page.locator("span[id='form_manterDocumentoHabil:lista_PCO:0:PCO_item_campoClassificacaoB_output_classificacao_contabil']").is_visible():
                conta_passivo = page.locator("span[id='form_manterDocumentoHabil:lista_PCO:0:PCO_item_campoClassificacaoB_output_classificacao_contabil']").text_content()
                #conferência['siafi'][19].value = conta_passivo

            valor_PCO = page.locator("span[id='form_manterDocumentoHabil:lista_PCO:0:painel_collapse_PCO_valor_cabecalho']").text_content()
            valor_PCO = valor_PCO.replace('.','')
            valor_PCO = valor_PCO.replace(',','.')
            planilha.range(f'I{linha}').value = valor_PCO


            ## Extrai informações das Deduções
            if page.locator("input:text-is('Dedução')").is_visible():
                page.locator("input:text-is('Dedução')").click()
                page.wait_for_timeout(2000)

                elementos = page.locator("span[id^='form_manterDocumentoHabil:listaDeducao:'][id$=':situacaoDeducaoHeader']").all()

                for index, elemento in enumerate(elementos):
                    texto = elemento.inner_text()
                    if "DDF025 - RETENÇÃO IMPOSTOS E CONTRIBUIÇÕES - IN RFB 1234-2012 - EFD-REINF R-4020" in texto:
                        elemento.click()

                        nat_rend = page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:campoInscricaoB_deducao_output']").text_content()
                        #conferência['siafi'][26].value = nat_rend

                        vencimento_ddf025 = page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:dataVencimentoDeducao_outputText']").text_content()
                        #conferência['siafi'][27].value = vencimento_ddf025

                        cod_darf = page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:campoInscricaoA_deducao_output']").text_content()
                        #conferência['siafi'][28].value = cod_darf

                        cnpj_darf = page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_recolhedor_outputText']").text_content()
                        #conferência['siafi'][29].value = cnpj_darf

                        bc_darf = page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_base_calculo_output']").text_content()
                        bc_darf = bc_darf.replace('.','')
                        bc_darf = bc_darf.replace(',','.')
                        bc_darf = float(bc_darf)
                        #conferência['siafi'][30].value = bc_darf

                        valor_darf = page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:panelItemDeducao_valor_cabecalho']").text_content()
                        valor_darf = valor_darf.replace('.','')
                        valor_darf = valor_darf.replace(',','.')
                        valor_darf = float(valor_darf)
                        planilha.range(f'K{linha}').value= valor_darf

                        page.locator(f"input[id='form_manterDocumentoHabil:listaDeducao:{index}:btnPredoc']").click(timeout=6000)
                        page.wait_for_timeout(2000)

                        obs_darf_siafi=page.locator("textarea[id='form_manterDocumentoHabil:observacaoPredoc']").text_content()
                        #conferência['siafi'][32].value = obs_darf_siafi

                        page.locator("input[id='form_manterDocumentoHabil:btnRetornarPredoc']").click(timeout=6000)  

                        elemento.click()

                        print(f"""
                            nat_rend: {nat_rend}
                            vencimento_ddf025: {vencimento_ddf025}
                            cod_darf: {cod_darf}
                            cnpj_darf: {cnpj_darf}
                            bc_darf: {bc_darf}
                            valor_darf: {valor_darf}
                            obs_darf_siafi: {obs_darf_siafi}
                            """)
                        # break

                    if "DDR001 - RETENÇÕES DE IMPOSTOS RECOLHÍVEIS POR DAR" in texto:
                        elemento.click()

                        ug_dar= page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:inputUgPagadoraDeducao_output']").text_content()
                        #planilha['siafi1'][51].value=ug_dar 

                        vencimento_dar = page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:dataVencimentoDeducao_outputText']").text_content()
                        #planilha['siafi1'][58].value=vencimento_dar
                        
                        municipio_favorecido_dar = page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:campoInscricaoA_deducao_output']").text_content()
                        #planilha['siafi1'][59].value=municipio_favorecido_dar
                        
                        receita_dar = page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:campoInscricaoB_deducao_output']").text_content()
                        #planilha['siafi1'][60].value=receita_dar
                        
                        cnpj_dar = page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_recolhedor_outputText']").text_content()
                        #planilha['siafi1'][61].value=cnpj_dar
                        
                        valor_dar = page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_valor_recolhimento_output']").text_content()
                        valor_dar = valor_dar.replace('.','')
                        valor_dar = valor_dar.replace(',','.')
                        valor_dar = float(valor_dar)
                        planilha.range(f'M{linha}').value = valor_dar
                    
                        page.locator(f"input[id='form_manterDocumentoHabil:listaDeducao:{index}:btnPredoc']").click()
                        page.wait_for_timeout(2000)

                        ref_dar = page.locator("span[id='form_manterDocumentoHabil:referenciaDar_Output']").text_content()
                        #planilha['siafi1'][50].value=ref_dar

                        municipio_nf_dar = page.locator("span[id='form_manterDocumentoHabil:municipioNFDar_outputText']").text_content()
                        #planilha['siafi1'][52].value=municipio_nf_dar
                        
                        numero_nf_dar = page.locator("span[id='form_manterDocumentoHabil:nfRecibo_outputText']").text_content()
                        #planilha['siafi1'][53].value=numero_nf_dar
                        
                        emissao_nf_dar = page.locator("span[id='form_manterDocumentoHabil:dataEmissaoNfDar_outputText']").text_content()
                        #planilha['siafi1'][54].value=emissao_nf_dar
                        
                        aliquota_dar = page.locator("span[id='form_manterDocumentoHabil:aliquotaNFDar_output']").text_content()
                        #planilha['siafi1'][55].value=aliquota_dar
                        
                        valor_nf_dar = page.locator("span[id='form_manterDocumentoHabil:valorNfDar_output']").text_content()
                        #planilha['siafi1'][56].value=valor_nf_dar
                        
                        obs_dar = page.locator("textarea[id='form_manterDocumentoHabil:observacaoPredoc']").text_content()
                        #planilha['siafi1'][57].value=obs_dar
                        
                        page.locator("input[id='form_manterDocumentoHabil:btnRetornarPredoc']").click()

                        elemento.click()

                        print(f"""
                            ref_dar: {ref_dar}
                            ug_dar: {ug_dar}
                            municipio_nf_dar: {municipio_nf_dar}
                            numero_nf_dar: {numero_nf_dar}
                            emissao_nf_dar: {emissao_nf_dar}
                            aliquota_dar: {aliquota_dar}
                            valor_nf_dar: {valor_nf_dar}
                            obs_dar: {obs_dar}
                            vencimento_dar: {vencimento_dar}
                            municipio_favorecido_dar: {municipio_favorecido_dar}
                            receita_dar: {receita_dar}
                            cnpj_dar: {cnpj_dar}
                            valor_dar: {valor_dar}
                        """)

            else:
                planilha.range(f'K{linha}').value = "ISENTO/IMUNE"
                planilha.range(f'M{linha}').value = "ISENTO/IMUNE"
                #continue

            ## Extrai informações dos Dados de Pagamento
            page.locator("input:text-is('Dados de Pagamento')").click()
            page.wait_for_timeout(2000)

            cnpj_pagamento_siafi = page.locator("span[id='form_manterDocumentoHabil:lista_DPgtoOB:0:codigoFavorecido_output']").text_content()
            #conferência['siafi'][52].value = cnpj_pagamento_siafi

            valor_pagamento_siafi = page.locator("span[id='form_manterDocumentoHabil:vlLiquidoPagar']").text_content()
            valor_pagamento_siafi=valor_pagamento_siafi.replace('.','')
            valor_pagamento_siafi=valor_pagamento_siafi.replace(',','.')
            valor_pagamento_siafi=float(valor_pagamento_siafi)
            planilha.range(f'N{linha}').value = valor_pagamento_siafi

            valor_favorecido_siafi = page.locator("span[id='form_manterDocumentoHabil:lista_DPgtoOB:0:valorPredoc_output']").text_content()
            valor_favorecido_siafi=valor_favorecido_siafi.replace('.','')
            valor_favorecido_siafi=valor_favorecido_siafi.replace(',','.')
            valor_favorecido_siafi=float(valor_favorecido_siafi)
            #conferência['siafi'][59].value = valor_favorecido_siafi
            
            page.locator("input[id='form_manterDocumentoHabil:lista_DPgtoOB:0:btnPredoc']").click()
            page.wait_for_timeout(2000)
            
            processo_pagamento_siafi = page.locator("span[id='form_manterDocumentoHabil:processoid_outputText']").text_content()
            #conferência['siafi'][53].value = processo_pagamento_siafi
            
            banco_siafi = page.locator("span[id='form_manterDocumentoHabil:favorecido_banco_outputText']").text_content()
            planilha.range(f'P{linha}').value= banco_siafi
            
            agencia_siafi = page.locator("span[id='form_manterDocumentoHabil:favorecido_agencia_outputText']").text_content()
            planilha.range(f'Q{linha}').value = agencia_siafi
            
            conta_siafi = page.locator("span[id='form_manterDocumentoHabil:favorecido_conta_outputText']").text_content()
            planilha.range(f'R{linha}').value = conta_siafi
            
            osb_pagamento_siafi = page.locator("textarea[id='form_manterDocumentoHabil:observacaoPredoc']").text_content()
            #conferência['siafi'][57].value = osb_pagamento_siafi
            
            page.locator("input[id='form_manterDocumentoHabil:btnRetornarPredoc']").click()

            ## Extrai informações do Centro de Cusatos
            page.locator("input:text-is('Detacustos')").click()
            page.wait_for_timeout(500)
            page.locator("span[class='ng-tns-c47-9 pi pi-plus']").click()

            custo_siafi = page.locator("div[class='p-text-center ng-star-inserted']").nth(-2).text_content()
            planilha.range(f'O{linha}').value = custo_siafi
            #pyautogui.alert('Pausa para conferência adicional')
            
        except:
            continue
   
    # Salvar o arquivo Excel
    work_book.save()
    
    show_alert()

            ###################################################################
            ####  DIFICULDADE DE ENCONTRAR DEDUÇÕES EM POSIÇÕES DIFERENTES  ####

            #     # if page.locator(":text-is('DDF021')").count()==2:
            #     #     posicao_inss=i
            #     #     apuracao_inss=page.locator("span[id='form_manterDocumentoHabil:predocDarfModal_periodoApuracao_outputText']").text_content()
            #     #     planilha['siafi1'][40].value=apuracao_inss
            #     #     obs_inss=page.locator("textarea[id='form_manterDocumentoHabil:observacaoPredoc']").text_content()
            #     #     planilha['siafi1'][41].value=obs_inss
            #     #     vencimento_inss=page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{posicao_inss}:dataVencimentoDeducao_outputText']").text_content()
            #     #     planilha['siafi1'][42].value=vencimento_inss
            #     #     cod_inss=page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{posicao_inss}:campoInscricaoA_deducao_output']").text_content()
            #     #     planilha['siafi1'][43].value=cod_inss
            #     #     cnpj_inss=page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{posicao_inss}:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_recolhedor_outputText']").text_content()
            #     #     planilha['siafi1'][44].value=cnpj_inss
            #     #     bc_inss=page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{posicao_inss}:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_base_calculo_output']").text_content()
            #     #     planilha['siafi1'][45].value=bc_inss
            #     #     valor_inss=page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{posicao_inss}:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_valor_recolhimento_output']").text_content()
            #     #     planilha['siafi1'][46].value=valor_inss                   
            #     #     page.locator("input[id='form_manterDocumentoHabil:btnRetornarPredoc']").click(timeout=6000)
                    


        #     try:
        #         page.locator("input:text-is('Dedução')").click(timeout=6000)
        #         page.wait_for_timeout(2000)
        #     except:
        #         conferência['siafi'][26:33].value = "RET. RFB ISENTO/IMUNE"
        #         conferência['siafi'][36:49].value = "RET. ISS ISENTO/IMUNE"

        #     #if codigo_recolhimento_DARF == 6147 and codigo_receita_DAR == 1782:
        #     if page.locator(":has-text('DDF025')").count()>=1 and page.locator(":has-text('DDR001')").count()>=1:
        #         nat_rend=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:campoInscricaoB_deducao_output']").text_content()
        #         conferência['siafi'][26].value = nat_rend
        #         vencimento_ddf025=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:dataVencimentoDeducao_outputText']").text_content()
        #         conferência['siafi'][27].value = vencimento_ddf025
        #         cod_darf=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:campoInscricaoA_deducao_output']").text_content()
        #         conferência['siafi'][28].value = cod_darf
        #         cnpj_darf=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_recolhedor_outputText']").text_content()
        #         conferência['siafi'][29].value = cnpj_darf
        #         bc_darf=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_base_calculo_output']").text_content()
        #         bc_darf=bc_darf.replace('.','')
        #         bc_darf=bc_darf.replace(',','.')
        #         bc_darf=float(bc_darf)
        #         conferência['siafi'][30].value = bc_darf
        #         valor_darf=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:panelItemDeducao_valor_cabecalho']").text_content()
        #         valor_darf=valor_darf.replace('.','')
        #         valor_darf=valor_darf.replace(',','.')
        #         valor_darf=float(valor_darf)
        #         conferência['siafi'][31].value = valor_darf
        #         page.locator("input[id='form_manterDocumentoHabil:listaDeducao:0:btnPredoc']").click(timeout=6000)
        #         page.wait_for_timeout(2000)
        #         obs_darf_siafi=page.locator("textarea[id='form_manterDocumentoHabil:observacaoPredoc']").text_content()
        #         conferência['siafi'][32].value = obs_darf_siafi
        #         page.locator("input[id='form_manterDocumentoHabil:btnRetornarPredoc']").click(timeout=6000)
        #         vencimento_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:listaDeducao:1:dataVencimentoDeducao_outputText']").text_content()
        #         conferência['siafi'][36].value = vencimento_dar_siafi
        #         cod_municipio_siafi = page.locator("span[id='form_manterDocumentoHabil:listaDeducao:1:campoInscricaoA_deducao_output']").text_content()
        #         conferência['siafi'][37].value = cod_municipio_siafi
        #         cod_receita_siafi = page.locator("span[id='form_manterDocumentoHabil:listaDeducao:1:campoInscricaoB_deducao_output']").text_content()
        #         conferência['siafi'][38].value = cod_receita_siafi
        #         cnpj_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:listaDeducao:1:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_recolhedor_outputText']").text_content()
        #         conferência['siafi'][39].value = cnpj_dar_siafi
        #         valor_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:listaDeducao:1:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_valor_recolhimento_output']").text_content()
        #         valor_dar_siafi=valor_dar_siafi.replace('.','')
        #         valor_dar_siafi=valor_dar_siafi.replace(',','.')
        #         valor_dar_siafi=float(valor_dar_siafi)
        #         conferência['siafi'][40].value = valor_dar_siafi
        #         page.locator("input[id='form_manterDocumentoHabil:listaDeducao:1:btnPredoc']").click(timeout=6000)
        #         page.wait_for_timeout(2000)
        #         ref_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:referenciaDar_Output']").text_content()
        #         conferência['siafi'][41].value = ref_dar_siafi
        #         ug_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:ugTomadoraDar_output']").text_content()
        #         conferência['siafi'][42].value = ug_dar_siafi
        #         municipio_nf_dar = page.locator("span[id='form_manterDocumentoHabil:municipioNFDar_outputText']").text_content()
        #         conferência['siafi'][43].value = municipio_nf_dar
        #         numero_nf_dar = page.locator("span[id='form_manterDocumentoHabil:nfRecibo_outputText']").text_content()
        #         numero_nf_dar=int(numero_nf_dar)
        #         conferência['siafi'][44].value = numero_nf_dar
        #         emissao_nf_dar = page.locator("span[id='form_manterDocumentoHabil:dataEmissaoNfDar_outputText']").text_content()
        #         conferência['siafi'][45].value = emissao_nf_dar
        #         aliquota_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:aliquotaNFDar_output']").text_content()
        #         conferência['siafi'][46].value = aliquota_dar_siafi
        #         valor_nf_dar = page.locator("span[id='form_manterDocumentoHabil:valorNfDar_output']").text_content()
        #         valor_nf_dar=valor_nf_dar.replace('.','')
        #         valor_nf_dar=valor_nf_dar.replace(',','.')
        #         valor_nf_dar=float(valor_nf_dar)
        #         conferência['siafi'][47].value = valor_nf_dar
        #         historico_dar_siafi = page.locator("textarea[id='form_manterDocumentoHabil:observacaoPredoc']").text_content()
        #         conferência['siafi'][48].value = historico_dar_siafi
        #         page.locator("input[id='form_manterDocumentoHabil:btnRetornarPredoc']").click(timeout=6000)
        #     if page.locator(":has-text('DDF025')").count()==0 and page.locator(":has-text('DDR001')").count()>=1:
        #         conferência['siafi'][26:33].value = "RET. RFB ISENTO/IMUNE"
        #         vencimento_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:dataVencimentoDeducao_outputText']").text_content()
        #         conferência['siafi'][36].value = vencimento_dar_siafi
        #         cod_municipio_siafi = page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:campoInscricaoA_deducao_output']").text_content()
        #         conferência['siafi'][37].value = cod_municipio_siafi
        #         cod_receita_siafi = page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:campoInscricaoB_deducao_output']").text_content()
        #         conferência['siafi'][38].value = cod_receita_siafi
        #         cnpj_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_recolhedor_outputText']").text_content()
        #         conferência['siafi'][39].value = cnpj_dar_siafi
        #         valor_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_valor_recolhimento_output']").text_content()
        #         conferência['siafi'][40].value = valor_dar_siafi
        #         page.locator("input[id='form_manterDocumentoHabil:listaDeducao:0:btnPredoc']").click(timeout=6000)
        #         page.wait_for_timeout(2000)
        #         ref_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:referenciaDar_Output']").text_content()
        #         conferência['siafi'][41].value = ref_dar_siafi
        #         ug_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:ugTomadoraDar_output']").text_content()
        #         conferência['siafi'][42].value = ug_dar_siafi
        #         municipio_nf_dar = page.locator("span[id='form_manterDocumentoHabil:municipioNFDar_outputText']").text_content()
        #         conferência['siafi'][43].value = municipio_nf_dar
        #         numero_nf_dar = page.locator("span[id='form_manterDocumentoHabil:nfRecibo_outputText']").text_content()
        #         conferência['siafi'][44].value = numero_nf_dar
        #         emissao_nf_dar = page.locator("span[id='form_manterDocumentoHabil:dataEmissaoNfDar_outputText']").text_content()
        #         conferência['siafi'][45].value = emissao_nf_dar
        #         aliquota_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:aliquotaNFDar_output']").text_content()
        #         conferência['siafi'][46].value = aliquota_dar_siafi
        #         valor_nf_dar = page.locator("span[id='form_manterDocumentoHabil:valorNfDar_output']").text_content()
        #         conferência['siafi'][47].value = valor_nf_dar
        #         historico_dar_siafi = page.locator("textarea[id='form_manterDocumentoHabil:observacaoPredoc']").text_content()
        #         conferência['siafi'][48].value = historico_dar_siafi
        #         page.locator("input[id='form_manterDocumentoHabil:btnRetornarPredoc']").click(timeout=6000)
        #     if page.locator(":has-text('DDF025')").count()>=1 and page.locator(":has-text('DDR001')").count()==0:
        #         conferência['siafi'][36:49].value = "RET. ISS ISENTO/IMUNE"
        #         nat_rend=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:campoInscricaoB_deducao_output']").text_content()
        #         conferência['siafi'][26].value = nat_rend
        #         vencimento_ddf025=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:dataVencimentoDeducao_outputText']").text_content()
        #         conferência['siafi'][27].value = vencimento_ddf025
        #         cod_darf=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:campoInscricaoA_deducao_output']").text_content()
        #         conferência['siafi'][28].value = cod_darf
        #         cnpj_darf=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_recolhedor_outputText']").text_content()
        #         conferência['siafi'][29].value = cnpj_darf
        #         bc_darf=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_base_calculo_output']").text_content()
        #         conferência['siafi'][30].value = bc_darf
        #         valor_darf=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:panelItemDeducao_valor_cabecalho']").text_content()
        #         conferência['siafi'][31].value = valor_darf
        #         page.locator("input[id='form_manterDocumentoHabil:listaDeducao:0:btnPredoc']").click(timeout=6000)
        #         page.wait_for_timeout(2000)
        #         obs_darf_siafi=page.locator("textarea[id='form_manterDocumentoHabil:observacaoPredoc']").text_content()
        #         conferência['siafi'][32].value = obs_darf_siafi
        #         page.locator("input[id='form_manterDocumentoHabil:btnRetornarPredoc']").click(timeout=6000)


    #     #Preenche as planilhas de conferência com os dados já verificados
    #     conferência=work_book.sheets[f'Conferência{linha-1}']
    #     conferência["J3"].value = num_np
    #     conferência['verificado'][2].value = planilha["B2"].value #data de vencimento
    #     conferência['verificado'][3].value = planilha["B1"].value #número do processo
    #     conferência['verificado'][4].value = planilha["B3"].value #data do ateste
    #     conferência['verificado'][5].value = planilha['valor_doc_fiscal'][linha-1].value #valor do documento fiscal, a ser comparado com o valor do documento no cabeçalho de DB
    #     CNPJ_planilha = planilha['emitente'][linha-1].value
    #     CNPJ_planilha = CNPJ_planilha.replace('.','')
    #     CNPJ_planilha = CNPJ_planilha.replace('/','')
    #     CNPJ_planilha = CNPJ_planilha.replace('-','')
    #     #conferência['verificado'][6].value = planilha['emitente'][linha-1].value #CNPJ, a ser comparado com o CNPJ no cabeçalho de DB
    #     conferência['verificado'][6].value = CNPJ_planilha
        
    #     #conferência['verificado'][7].value = planilha['emitente'][linha-1].value #CNPJ, a ser comparado com o CNPJ em dados do doc de origem em DB
    #     conferência['verificado'][7].value = CNPJ_planilha
    #     conferência['verificado'][8].value = planilha['data_de_emissao'][linha-1].value #data de emissão, a ser comparada com a data de emissão em DB
    #     conferência['verificado'][9].value = planilha['num_doc_fiscal'][linha-1].value # n doc
    #     conferência['verificado'][10].value = planilha['valor_doc_fiscal'][linha-1].value #valor do doc, a ser comparado com o valor em dados do doc de origem em DB
    #     #Procedimentos para obter os textos das observações
    #     num_doc_fiscal = str(int(planilha['num_doc_fiscal'][linha-1].value))
    #     processo_edoc = planilha.range('B1').value
    #     data_de_emissao_planilha = planilha['data_de_emissao'][linha-1].value
    #     data_de_emissao = str(data_de_emissao_planilha)
    #     print(data_de_emissao)
    #     ano_emissao = data_de_emissao[-4:]
    #     mes_emissao = data_de_emissao[3:5]
    #     #descricao = work_book.sheets['NE']['descricao'][3].value 
    #     codigo_recolhimento_DARF = planilha['codigo_recolhimento_DARF'][linha-1].value
    #     if codigo_recolhimento_DARF == 6147:
    #         obs_DARF = f'RET. RFB 5,85PC - IN 1234/2012'
    #     elif codigo_recolhimento_DARF == "ISENTO":
    #         obs_DARF = f'RET. RFB ISENTO'
    #     elif codigo_recolhimento_DARF == "IMUNE":
    #         obs_DARF = f'RET. RFB IMUNE'
    #     codigo_receita_DAR = planilha['codigo_receita_DAR'][linha-1].value
    #     if codigo_receita_DAR == 1782:
    #         obs_DAR = f'RET. ISS 2PC' 
    #     elif codigo_receita_DAR == "ISENTO":
    #         obs_DAR = f'RET. ISS ISENTO'
    #     elif codigo_receita_DAR == "IMUNE":
    #         obs_DAR = f'RET. ISS IMUNE'  
    #     elif codigo_receita_DAR == "OUTRO MUNICÍPIO":
    #         obs_DAR = f'SERVIÇO PRESTADO EM OUTRO MUNICÍPIO'  
    #     observacao_DARF = str(f'NFSE {num_doc_fiscal} - {descricao} - {obs_DARF}').upper()
    #     observacao_DAR = str(f'NFSE {num_doc_fiscal} - {descricao} - {obs_DAR}').upper()
    #     observacao = str(f'NFSE {num_doc_fiscal} - {descricao} - PERIODO: {mes_emissao}/{ano_emissao} - {obs_DARF} - {obs_DAR} - EDOC: {processo_edoc}.').upper()
    #     #Prossegue-se com o preenchimento da coluna "verificado" na planilha de conferência"
    #     conferência['verificado'][11].value = observacao #histórico que deve aparecer em dados básicos, a ser comparado com o do siafi
    #     conferência['verificado'][12].value = planilha['nome_emitente'][linha-1].value
    #     conferência['verificado'][17].value = nota_empenho
    #     conferência['verificado'][18].value = "3.3.2.3.1.01.00"
    #     conferência['verificado'][19].value = "2.1.3.1.1.04.00"
    #     conferência['verificado'][20].value = planilha['valor_PCO'][linha-1].value
    #     if codigo_recolhimento_DARF == 6147:
    #         conferência['verificado'][26].value = "17005"
    #         conferência['verificado'][27].value = planilha["B4"].value #data de vencimento da ddf025
    #         conferência['verificado'][28].value = "6147"
    #         conferência['verificado'][29].value = CNPJ_planilha
    #         conferência['verificado'][30].value = planilha['valor_doc_fiscal'][linha-1].value
    #         conferência['verificado'][31].value = planilha['valor_DARF'][linha-1].value
    #         conferência['verificado'][32].value = observacao_DARF

    #     else:
    #         conferência['verificado'][26:33].value = "RET. RFB ISENTO/IMUNE"
        
    #     if codigo_receita_DAR == 1782:
    #         conferência['verificado'][36].value = planilha["B5"].value #data de vencimento da DDR001
    #         conferência['verificado'][37].value = "9701"
    #         conferência['verificado'][38].value = "1782"
    #         conferência['verificado'][39].value = CNPJ_planilha
    #         conferência['verificado'][40].value = planilha['valor_DAR'][linha-1].value
    #         conferência['verificado'][41].value = f'{mes_emissao}/{ano_emissao}'
    #         conferência['verificado'][42].value = "010001"
    #         conferência['verificado'][43].value = "9701"
    #         conferência['verificado'][44].value = planilha['num_doc_fiscal'][linha-1].value
    #         conferência['verificado'][45].value = planilha['data_de_emissao'][linha-1].value
    #         conferência['verificado'][46].value = "2,000"
    #         conferência['verificado'][47].value = planilha['valor_doc_fiscal'][linha-1].value
    #         conferência['verificado'][48].value = observacao_DAR
    #     else:
    #         conferência['verificado'][36:49].value = "RET. ISS ISENTO/IMUNE"
        
    #     conferência['verificado'][52].value = CNPJ_planilha
    #     conferência['verificado'][53].value = planilha["B1"].value
    #     conferência['verificado'][54].value = planilha['banco'][linha-1].value
    #     conferência['verificado'][55].value = planilha['agencia'][linha-1].value
    #     conferência['verificado'][56].value = planilha['conta'][linha-1].value
    #     conferência['verificado'][57].value = observacao
    #     conferência['verificado'][58].value = planilha['valor_pagamento'][linha-1].value
    #     conferência['verificado'][59].value = planilha['valor_pagamento'][linha-1].value

    #     conferência['verificado'][61].value = f'{mes_emissao}/{ano_emissao}'

    #     #Aqui começa a extração das informações do siafi.

    #     page.locator("input[id='frmMenu:acessoRapido']").fill("condh")
    #     page.locator("input:right-of(input[id='frmMenu:acessoRapido'])").click(timeout=6000)
    #     #page.locator("input:below(:text-is('Tipo'))>>nth=0").fill('NP')
    #     page.locator("input[id='formConsultarDH:tipoDH_input']").fill('NP')
    #     #Número_NP="4953"
    #     #LINHA ABAIXO PARA TESTE NAS NPS DE 2024, COMENTAR DEPOIS.
    #     #page.locator("input[id='formConsultarDH:anoDH_input']").fill("2024")
    #     page.locator("input:below(:text-is('Número'))>>nth=0").fill(num_np)
    #     page.wait_for_timeout(2000)
    #     page.locator("input:text-is('Pesquisar')>>nth=0").click(timeout=6000)
    #     data_vencimento_dh=page.locator("span[id='form_manterDocumentoHabil:dataVencimento_outputText']").text_content()
    #     conferência['siafi'][2].value = data_vencimento_dh
    #     processo_cabeçalho=page.locator("span[id='form_manterDocumentoHabil:processo_outputText']").text_content()
    #     conferência['siafi'][3].value = processo_cabeçalho
    #     ateste_cabeçalho=page.locator("span[id='form_manterDocumentoHabil:dataAteste_outputText']").text_content()
    #     conferência['siafi'][4].value = ateste_cabeçalho
    #     valor_cabeçalho=page.locator("span[id='form_manterDocumentoHabil:valorPrincipalDocumento_output']").text_content()
    #     valor_cabeçalho=valor_cabeçalho.replace('.','')
    #     valor_cabeçalho=valor_cabeçalho.replace(',','.')
    #     valor_cabeçalho=float(valor_cabeçalho)
        
    #     conferência['siafi'][5].value = valor_cabeçalho
    #     CNPJ_cabeçalho=page.locator("span[id='form_manterDocumentoHabil:credorDevedor_output']").text_content()
    #     conferência['siafi'][6].value = CNPJ_cabeçalho
    #     CNPJ_origem=page.locator("span[id='form_manterDocumentoHabil:tableDocsOrigem:0:emitenteDocOrigem_output']").text_content()
    #     conferência['siafi'][7].value = CNPJ_origem
    #     emissão_origem=page.locator("span[id='form_manterDocumentoHabil:tableDocsOrigem:0:dataEmissaoDocOrigem_outputText']").text_content()
    #     conferência['siafi'][8].value = emissão_origem
    #     número_origem=page.locator("span[id='form_manterDocumentoHabil:tableDocsOrigem:0:numeroDocOrigem_outputText']").text_content()
    #     número_origem=int(número_origem)
    #     conferência['siafi'][9].value = número_origem
    #     valor_origem=page.locator("span[id='form_manterDocumentoHabil:tableDocsOrigem:0:valorDocOrigem_output']").text_content().replace(" ","")
    #     valor_origem=valor_origem.replace('.','')
    #     valor_origem=valor_origem.replace(',','.')
        
    #     valor_origem=float(valor_origem)
    #     conferência['siafi'][10].value = valor_origem
    #     obs_dados_básicos=page.locator("textarea[id='form_manterDocumentoHabil:observacao']").text_content()
    #     conferência['siafi'][11].value = obs_dados_básicos
    #     nome_credor = page.locator("span[id='form_manterDocumentoHabil:nomeCredorDevedor']").text_content()
    #     conferência['siafi'][12].value = nome_credor

    #     page.locator("input:text-is('Principal Com Orçamento')").click(timeout=6000)

    #     empenho_siafi=page.locator("span[id='form_manterDocumentoHabil:lista_PCO:0:PCO_item_num_empenho_header']").text_content()
    #     conferência['siafi'][17].value = empenho_siafi
    #     conta_vpd=page.locator("span[id='form_manterDocumentoHabil:lista_PCO:0:PCO_item_campoClassificacaoA_output_classificacao_contabil']").text_content()
    #     conferência['siafi'][18].value = conta_vpd
    #     conta_passivo=page.locator("span[id='form_manterDocumentoHabil:lista_PCO:0:PCO_item_campoClassificacaoB_output_classificacao_contabil']").text_content()
    #     conferência['siafi'][19].value = conta_passivo
    #     valor_PCO=page.locator("span[id='form_manterDocumentoHabil:lista_PCO:0:painel_collapse_PCO_valor_cabecalho']").text_content()
    #     valor_PCO=valor_PCO.replace('.','')
    #     valor_PCO=valor_PCO.replace(',','.')
    #     conferência['siafi'][20].value = valor_PCO

    #     try:
    #         page.locator("input:text-is('Dedução')").click(timeout=6000)
    #         page.wait_for_timeout(2000)
    #     except:
    #         conferência['siafi'][26:33].value = "RET. RFB ISENTO/IMUNE"
    #         conferência['siafi'][36:49].value = "RET. ISS ISENTO/IMUNE"
    #     #if codigo_recolhimento_DARF == 6147 and codigo_receita_DAR == 1782:
    #     if page.locator(":has-text('DDF025')").count()>=1 and page.locator(":has-text('DDR001')").count()>=1:
    #         nat_rend=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:campoInscricaoB_deducao_output']").text_content()
    #         conferência['siafi'][26].value = nat_rend
    #         vencimento_ddf025=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:dataVencimentoDeducao_outputText']").text_content()
    #         conferência['siafi'][27].value = vencimento_ddf025
    #         cod_darf=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:campoInscricaoA_deducao_output']").text_content()
    #         conferência['siafi'][28].value = cod_darf
    #         cnpj_darf=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_recolhedor_outputText']").text_content()
    #         conferência['siafi'][29].value = cnpj_darf
    #         bc_darf=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_base_calculo_output']").text_content()
    #         bc_darf=bc_darf.replace('.','')
    #         bc_darf=bc_darf.replace(',','.')
    #         bc_darf=float(bc_darf)
    #         conferência['siafi'][30].value = bc_darf
    #         valor_darf=page.locator("span[id='form_manterDocumentoHabil:lista_PCO:0:painel_collapse_PCO_valor_cabecalho']").text_content()
    #         valor_darf=valor_darf.replace('.','')
    #         valor_darf=valor_darf.replace(',','.')
    #         valor_darf=float(valor_darf)
    #         conferência['siafi'][31].value = valor_darf
    #         page.locator("input[id='form_manterDocumentoHabil:listaDeducao:0:btnPredoc']").click(timeout=6000)
    #         page.wait_for_timeout(2000)
    #         obs_darf_siafi=page.locator("textarea[id='form_manterDocumentoHabil:observacaoPredoc']").text_content()
    #         conferência['siafi'][32].value = obs_darf_siafi
    #         page.locator("input[id='form_manterDocumentoHabil:btnRetornarPredoc']").click(timeout=6000)
    #         vencimento_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:listaDeducao:1:dataVencimentoDeducao_outputText']").text_content()
    #         conferência['siafi'][36].value = vencimento_dar_siafi
    #         cod_municipio_siafi = page.locator("span[id='form_manterDocumentoHabil:listaDeducao:1:campoInscricaoA_deducao_output']").text_content()
    #         conferência['siafi'][37].value = cod_municipio_siafi
    #         cod_receita_siafi = page.locator("span[id='form_manterDocumentoHabil:listaDeducao:1:campoInscricaoB_deducao_output']").text_content()
    #         conferência['siafi'][38].value = cod_receita_siafi
    #         cnpj_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:listaDeducao:1:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_recolhedor_outputText']").text_content()
    #         conferência['siafi'][39].value = cnpj_dar_siafi
    #         valor_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:listaDeducao:1:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_valor_recolhimento_output']").text_content()
    #         valor_dar_siafi=valor_dar_siafi.replace('.','')
    #         valor_dar_siafi=valor_dar_siafi.replace(',','.')
    #         valor_dar_siafi=float(valor_dar_siafi)
    #         conferência['siafi'][40].value = valor_dar_siafi
    #         page.locator("input[id='form_manterDocumentoHabil:listaDeducao:1:btnPredoc']").click(timeout=6000)
    #         page.wait_for_timeout(2000)
    #         ref_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:referenciaDar_Output']").text_content()
    #         conferência['siafi'][41].value = ref_dar_siafi
    #         ug_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:ugTomadoraDar_output']").text_content()
    #         conferência['siafi'][42].value = ug_dar_siafi
    #         municipio_nf_dar = page.locator("span[id='form_manterDocumentoHabil:municipioNFDar_outputText']").text_content()
    #         conferência['siafi'][43].value = municipio_nf_dar
    #         numero_nf_dar = page.locator("span[id='form_manterDocumentoHabil:nfRecibo_outputText']").text_content()
    #         numero_nf_dar=int(numero_nf_dar)
    #         conferência['siafi'][44].value = numero_nf_dar
    #         emissao_nf_dar = page.locator("span[id='form_manterDocumentoHabil:dataEmissaoNfDar_outputText']").text_content()
    #         conferência['siafi'][45].value = emissao_nf_dar
    #         aliquota_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:aliquotaNFDar_output']").text_content()
    #         conferência['siafi'][46].value = aliquota_dar_siafi
    #         valor_nf_dar = page.locator("span[id='form_manterDocumentoHabil:valorNfDar_output']").text_content()
    #         valor_nf_dar=valor_nf_dar.replace('.','')
    #         valor_nf_dar=valor_nf_dar.replace(',','.')
    #         valor_nf_dar=float(valor_nf_dar)
    #         conferência['siafi'][47].value = valor_nf_dar
    #         historico_dar_siafi = page.locator("textarea[id='form_manterDocumentoHabil:observacaoPredoc']").text_content()
    #         conferência['siafi'][48].value = historico_dar_siafi
    #         page.locator("input[id='form_manterDocumentoHabil:btnRetornarPredoc']").click(timeout=6000)
    #     if page.locator(":has-text('DDF025')").count()==0 and page.locator(":has-text('DDR001')").count()>=1:
    #         conferência['siafi'][26:33].value = "RET. RFB ISENTO/IMUNE"
    #         vencimento_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:dataVencimentoDeducao_outputText']").text_content()
    #         conferência['siafi'][36].value = vencimento_dar_siafi
    #         cod_municipio_siafi = page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:campoInscricaoA_deducao_output']").text_content()
    #         conferência['siafi'][37].value = cod_municipio_siafi
    #         cod_receita_siafi = page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:campoInscricaoB_deducao_output']").text_content()
    #         conferência['siafi'][38].value = cod_receita_siafi
    #         cnpj_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_recolhedor_outputText']").text_content()
    #         conferência['siafi'][39].value = cnpj_dar_siafi
    #         valor_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_valor_recolhimento_output']").text_content()
    #         conferência['siafi'][40].value = valor_dar_siafi
    #         page.locator("input[id='form_manterDocumentoHabil:listaDeducao:0:btnPredoc']").click(timeout=6000)
    #         page.wait_for_timeout(2000)
    #         ref_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:referenciaDar_Output']").text_content()
    #         conferência['siafi'][41].value = ref_dar_siafi
    #         ug_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:ugTomadoraDar_output']").text_content()
    #         conferência['siafi'][42].value = ug_dar_siafi
    #         municipio_nf_dar = page.locator("span[id='form_manterDocumentoHabil:municipioNFDar_outputText']").text_content()
    #         conferência['siafi'][43].value = municipio_nf_dar
    #         numero_nf_dar = page.locator("span[id='form_manterDocumentoHabil:nfRecibo_outputText']").text_content()
    #         conferência['siafi'][44].value = numero_nf_dar
    #         emissao_nf_dar = page.locator("span[id='form_manterDocumentoHabil:dataEmissaoNfDar_outputText']").text_content()
    #         conferência['siafi'][45].value = emissao_nf_dar
    #         aliquota_dar_siafi = page.locator("span[id='form_manterDocumentoHabil:aliquotaNFDar_output']").text_content()
    #         conferência['siafi'][46].value = aliquota_dar_siafi
    #         valor_nf_dar = page.locator("span[id='form_manterDocumentoHabil:valorNfDar_output']").text_content()
    #         conferência['siafi'][47].value = valor_nf_dar
    #         historico_dar_siafi = page.locator("textarea[id='form_manterDocumentoHabil:observacaoPredoc']").text_content()
    #         conferência['siafi'][48].value = historico_dar_siafi
    #         page.locator("input[id='form_manterDocumentoHabil:btnRetornarPredoc']").click(timeout=6000)
    #     if page.locator(":has-text('DDF025')").count()>=1 and page.locator(":has-text('DDR001')").count()==0:
    #         conferência['siafi'][36:49].value = "RET. ISS ISENTO/IMUNE"
    #         nat_rend=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:campoInscricaoB_deducao_output']").text_content()
    #         conferência['siafi'][26].value = nat_rend
    #         vencimento_ddf025=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:dataVencimentoDeducao_outputText']").text_content()
    #         conferência['siafi'][27].value = vencimento_ddf025
    #         cod_darf=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:campoInscricaoA_deducao_output']").text_content()
    #         conferência['siafi'][28].value = cod_darf
    #         cnpj_darf=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_recolhedor_outputText']").text_content()
    #         conferência['siafi'][29].value = cnpj_darf
    #         bc_darf=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_base_calculo_output']").text_content()
    #         conferência['siafi'][30].value = bc_darf
    #         valor_darf=page.locator("span[id='form_manterDocumentoHabil:listaDeducao:0:panelItemDeducao_valor_cabecalho']").text_content()
    #         conferência['siafi'][31].value = valor_darf
    #         page.locator("input[id='form_manterDocumentoHabil:listaDeducao:0:btnPredoc']").click(timeout=6000)
    #         page.wait_for_timeout(2000)
    #         obs_darf_siafi=page.locator("textarea[id='form_manterDocumentoHabil:observacaoPredoc']").text_content()
    #         conferência['siafi'][32].value = obs_darf_siafi
    #         page.locator("input[id='form_manterDocumentoHabil:btnRetornarPredoc']").click(timeout=6000)


    #     page.locator("input:text-is('Dados de Pagamento')").click(timeout=6000)
    #     page.wait_for_timeout(2000)
    #     cnpj_pagamento_siafi = page.locator("span[id='form_manterDocumentoHabil:lista_DPgtoOB:0:codigoFavorecido_output']").text_content()
    #     conferência['siafi'][52].value = cnpj_pagamento_siafi
    #     page.locator("input[id='form_manterDocumentoHabil:lista_DPgtoOB:0:btnPredoc']").click(timeout=6000)
    #     page.wait_for_timeout(2000)
    #     processo_pagamento_siafi = page.locator("span[id='form_manterDocumentoHabil:processoid_outputText']").text_content()

    #     conferência['siafi'][53].value = processo_pagamento_siafi
    #     banco_siafi = page.locator("span[id='form_manterDocumentoHabil:favorecido_banco_outputText']").text_content()
    #     conferência['siafi'][54].value = banco_siafi
    #     agencia_siafi = page.locator("span[id='form_manterDocumentoHabil:favorecido_agencia_outputText']").text_content()
    #     conferência['siafi'][55].value = agencia_siafi
    #     conta_siafi = page.locator("span[id='form_manterDocumentoHabil:favorecido_conta_outputText']").text_content()
    #     conferência['siafi'][56].value = conta_siafi
    #     osb_pagamento_siafi = page.locator("textarea[id='form_manterDocumentoHabil:observacaoPredoc']").text_content()
    #     conferência['siafi'][57].value = osb_pagamento_siafi
        
    #     page.locator("input[id='form_manterDocumentoHabil:btnRetornarPredoc']").click(timeout=6000)
    #     valor_pagamento_siafi = page.locator("span[id='form_manterDocumentoHabil:vlLiquidoPagar']").text_content()
    #     valor_pagamento_siafi=valor_pagamento_siafi.replace('.','')
    #     valor_pagamento_siafi=valor_pagamento_siafi.replace(',','.')
    #     valor_pagamento_siafi=float(valor_pagamento_siafi)
    #     conferência['siafi'][58].value = valor_pagamento_siafi
    #     valor_favorecido_siafi = page.locator("span[id='form_manterDocumentoHabil:lista_DPgtoOB:0:valorPredoc_output']").text_content()
    #     valor_favorecido_siafi=valor_favorecido_siafi.replace('.','')
    #     valor_favorecido_siafi=valor_favorecido_siafi.replace(',','.')
    #     valor_favorecido_siafi=float(valor_favorecido_siafi)
    #     conferência['siafi'][59].value = valor_favorecido_siafi

    #     page.locator("input:text-is('Centro de Custo')").click(timeout=6000)

    #     page.wait_for_timeout(2000)

    #     custo_siafi = page.locator("span[id='form_manterDocumentoHabil:vinculosCentroCusto:0:referencia']").text_content()
    #     conferência['siafi'][61].value = custo_siafi
    #     pyautogui.alert('Pausa para conferência adicional')


    # show_alert()


def interigir_abas(context, paginas):
    def on_option_click(action):
        root.destroy()  # Fechar a janela atual de opções
        aba_aberta = False
        if action == "SIAFI":
            for index, pagina in enumerate(paginas):
                if action.lower() in pagina.lower():  # Neste caso a aba do Siafi está aberta em alguma das abas
                    page = context.pages[index]
                    aba_aberta = True
                    # Aqui é preciso verificar se o usuário está logado corretamente.
                    # É preciso garantir que a página do Siafi está na posição inicial: Gercomp
                    mostra_opcoes_siafi(context, page, paginas)  # Passa a página correta aqui
                    break
            if aba_aberta == False:
                # fazer o tratamento aqui para o caso de não ter uma janela aberta no Siafi.
                pass
        elif action == "EDOC":
            for index, pagina in enumerate(paginas):
                if action.lower() in pagina.lower():  # Neste caso a aba do Edoc está aberta em alguma das abas
                    page = context.pages[index]
                    aba_aberta = True
                    mostra_opcoes_edoc(context, page, paginas)  # Passa a página correta aqui
                    break
                else:
                    pass  # Aqui, fazer o tratamento para abrir uma aba dedicada ao Edoc

    root = ctk.CTk()
    root.geometry("400x200")
    root.title("Escolha uma Opção")

    label = ctk.CTkLabel(root, text="Escolha uma opção:")
    label.pack(pady=20)

    botao_DJSID = ctk.CTkButton(
        root, text="Trabalhar com o Siafi", command=lambda: on_option_click("SIAFI"))
    botao_DJSID.pack(pady=10)

    botao_edoc = ctk.CTkButton(
        root, text="Trabalhar com o Edoc", command=lambda: on_option_click("EDOC"))
    botao_edoc.pack(pady=10)

    root.mainloop()


def mostra_opcoes_siafi(context, page, paginas):
    def on_option_click(action):
        root.destroy()  # Fechar a janela atual de opções
        if action == "incdh":
            # caminho_arquivo = f"{os.path.dirname(__file__)}\\banco2.txt"
            liquidacao_pro_saude(page)
        elif action == "conne":
            conne(page)
        elif action == "conferir liquidação":
            conferir_liquidação(page)
            return
        # Mostrar as opções novamente
        mostra_opcoes_siafi(context, page, paginas)

    root = ctk.CTk()
    root.geometry("400x500")
    root.title("Escolha uma Opção")

    label = ctk.CTkLabel(root, text="Escolha uma opção:")
    label.pack(pady=20)

    botao_dados_bancarios = ctk.CTkButton(
        root, text="Incluir Documento Hábil", command=lambda: on_option_click("incdh"))
    botao_dados_bancarios.pack(pady=10)

    botao_dados_despachos = ctk.CTkButton(
        root, text="Consultar Nota de Empenho", command=lambda: on_option_click("conne"))
    botao_dados_despachos.pack(pady=10)

    botao_dados_liquid = ctk.CTkButton(
        root, text="Conferir Liquidação", command=lambda: on_option_click("conferir liquidação"))
    botao_dados_liquid.pack(pady=10)

    root.mainloop()


def mostra_opcoes_edoc(context, page, paginas):
    def on_option_click(action):
        root.destroy()  # Fechar a janela atual de opções
        if action == "pdf_NFSE":
            extrair_pdf_NFSE(page)
        elif action == "pdf_ateste":
            extrair_pdf_ateste(page)
        elif action == "pdf_ateste_cnpj":
            extrair_pdf_ateste_cnpj_unico(page)
        elif action == "NS_edoc":
            inclui_NS_edoc(page)
            return
        # Mostrar as opções novamente
        mostra_opcoes_edoc(context, page, paginas)

    root = ctk.CTk()
    root.geometry("400x500")
    root.title("Escolha uma Opção")

    label = ctk.CTkLabel(root, text="Escolha uma opção:")
    label.pack(pady=20)

    botao_dados_bancarios = ctk.CTkButton(
        root, text="Extrair dados da NFSE", command=lambda: on_option_click("pdf_NFSE"))
    botao_dados_bancarios.pack(pady=10)

    botao_extrair_ateste = ctk.CTkButton(
        root, text="Extrair dados do Relatório", command=lambda: on_option_click("pdf_ateste"))
    botao_extrair_ateste.pack(pady=10)

    botao_extrair_ateste_cnpj = ctk.CTkButton(
        root, text="Extrair Relatório (CNPJ único)", command=lambda: on_option_click("pdf_ateste_cnpj"))
    botao_extrair_ateste_cnpj.pack(pady=10)

    botao_dados_despachos = ctk.CTkButton(
        root, text="Incluir NS ao processo eDoc", command=lambda: on_option_click("NS_edoc"))
    botao_dados_despachos.pack(pady=10)

    root.mainloop()


def show_alert():
    alert = ctk.CTk()
    alert.title("Alerta")
    alert.geometry("300x150")

    label = ctk.CTkLabel(alert, text="Sistema Executado")
    label.pack(pady=20)

    button = ctk.CTkButton(alert, text="OK", command=alert.destroy)
    button.pack(pady=10)

    alert.mainloop()


def run(playwright):
    browser = playwright.chromium.connect_over_cdp("http://localhost:9230")
    context = browser.contexts[0]

    # Obtenha todas as páginas abertas no contexto
    pages = context.pages
    paginas = []
    for index, pagina in enumerate(pages):
        paginas.append(f"{index + 1}:{pagina.url}")
    interigir_abas(context, paginas)

    browser.close()


def main():
    with sync_playwright() as playwright:
        run(playwright)


if getattr(sys, 'frozen', False):
    diretorio = os.path.dirname(sys.executable)
    arquivo = os.path.basename(sys.executable)[:-4]
    main()
elif __file__ and __name__ == '__main__':  # Add this condition
    diretorio = os.path.dirname(__file__)
    arquivo = os.path.basename(__file__)[:-3]
    # encontra a planilha Excel
    diretorio_arquivo = f"{diretorio}\\{arquivo}.xlsm"
    # diretorio_arquivo = r"Z:\MACROS\MACROS_SIAFI\PROGRAMAS\DEP_JUDICIAL_SEM_ID\DEP_JUDICIAL_SEM_ID.xlsm"
    main()