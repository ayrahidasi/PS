import os
import sys

# Variáveis globais para armazenar parâmetros do Flask
_web_mode_params = {}

def get_pro_saude_directory():
    """Obtém o diretório PRO_SAUDE na área de trabalho do usuário"""
    user_home = os.path.expanduser('~')
    pro_saude_dir = os.path.join(user_home, 'Desktop', 'PRO_SAUDE')
    
    # Cria o diretório se não existir
    if not os.path.exists(pro_saude_dir):
        try:
            os.makedirs(pro_saude_dir, exist_ok=True)
        except Exception as e:
            print(f"Aviso: Não foi possível criar o diretório {pro_saude_dir}: {e}")
            # Fallback para o diretório atual se não conseguir criar
            return os.path.dirname(os.path.abspath(__file__))
    
    return pro_saude_dir

def set_web_mode(enabled=True):
    """Define se o sistema está rodando em modo web (Flask)"""
    global _web_mode_params
    _web_mode_params['_WEB_MODE'] = enabled

def is_web_mode():
    """Verifica se está em modo web"""
    return _web_mode_params.get('_WEB_MODE', False)

def set_web_params(**params):
    """Define parâmetros para uso em modo web"""
    global _web_mode_params
    _web_mode_params.update(params)

def get_web_param(key, default=None):
    """Obtém parâmetro definido para modo web"""
    return _web_mode_params.get(key, default)

# Definir funções mock primeiro
def mock_prompt(text='', title='', default=''):
    if 'linha inicial' in text.lower():
        return str(get_web_param('linha_inicial', default or '2'))
    elif 'linha final' in text.lower():
        return str(get_web_param('linha_final', default or get_web_param('linha_inicial', '2')))
    return str(default)

def mock_confirm(text='', buttons=None):
    if 'baixar' in text.lower() and 'selecionar' in text.lower():
        escolha = get_web_param('escolha', 'Baixar')
        return escolha
    elif buttons and len(buttons) > 0:
        return get_web_param('escolha', buttons[0])
    return 'OK'

def mock_alert(text=''):
    pass  # Não faz nada em modo web

def mock_filedialog_open(**kwargs):
    """Substituto que retorna o arquivo uploadado se disponível"""
    uploaded_file = get_web_param('uploaded_file_path')
    if uploaded_file and os.path.exists(uploaded_file):
        return uploaded_file
    else:
        raise ValueError("Nenhum arquivo foi uploadado ou arquivo não encontrado")

# Classes mock simples - SEM herança de Mock para evitar recursão
class SimpleMockPyAutoGUI:
    def prompt(self, *args, **kwargs):
        return mock_prompt(*args, **kwargs)
    
    def confirm(self, *args, **kwargs):
        return mock_confirm(*args, **kwargs)
    
    def alert(self, *args, **kwargs):
        return mock_alert(*args, **kwargs)

class SimpleMockTkinterRoot:
    def __init__(self):
        pass
    def withdraw(self):
        pass
    def destroy(self):
        pass
    def mainloop(self):
        pass

class SimpleMockCustomTkinter:
    @staticmethod
    def CTk():
        return SimpleMockTkinterRoot()

class SimpleMockFileDialog:
    @staticmethod
    def askopenfilename(**kwargs):
        return mock_filedialog_open(**kwargs)

class SimpleMockTkinter:
    Tk = SimpleMockTkinterRoot
    filedialog = SimpleMockFileDialog

# Mock das dependências problemáticas ANTES de qualquer import
sys.modules['pyautogui'] = SimpleMockPyAutoGUI()
sys.modules['customtkinter'] = SimpleMockCustomTkinter()
sys.modules['tkinter'] = SimpleMockTkinter()
sys.modules['tkinter.filedialog'] = SimpleMockFileDialog()

# Set environment variable to prevent GUI execution
os.environ['FLASK_WEB_MODE'] = '1'

# Safely import PRO_SAUDE by preventing main execution
old_argv = sys.argv
try:
    # Temporarily modify sys.argv to prevent main() execution paths
    sys.argv = ['web_mode']
    import PRO_SAUDE
    
    # Define as variáveis globais que normalmente seriam definidas no main()
    # Essas são necessárias para as funções que acessam arquivos Excel
    # ALTERAÇÃO: Usar o diretório PRO_SAUDE na área de trabalho
    PRO_SAUDE.diretorio = get_pro_saude_directory()
    PRO_SAUDE.arquivo = 'PRO_SAUDE'  # Nome base do arquivo
    
    # PATCHES IMEDIATOS para todas as funções que usam diretório local
    def create_patched_function(original_func, func_name):
        """Cria uma versão patcheada de uma função que substitui referências ao diretório"""
        def patched_func(*args, **kwargs):
            # Salva os valores originais se existirem
            old_diretorio = getattr(PRO_SAUDE, 'diretorio', None)
            old_arquivo = getattr(PRO_SAUDE, 'arquivo', None)
            
            try:
                # Define os valores corretos
                PRO_SAUDE.diretorio = get_pro_saude_directory()
                PRO_SAUDE.arquivo = 'PRO_SAUDE'
                
                # Executa a função original
                return original_func(*args, **kwargs)
            finally:
                # Restaura valores se necessário (mas mantenha os novos)
                PRO_SAUDE.diretorio = get_pro_saude_directory()
                PRO_SAUDE.arquivo = 'PRO_SAUDE'
        
        return patched_func
    
    # PATCH IMEDIATO: Substituir a função obter_planilha_e_aba
    if hasattr(PRO_SAUDE, 'obter_planilha_e_aba'):
        def patched_obter_planilha_e_aba_immediate(nome_aba):
            """Versão patcheada imediata da função obter_planilha_e_aba"""
            import xlwings
            diretorio = get_pro_saude_directory()
            arquivo = 'PRO_SAUDE'
            diretorio_arquivo = f"{diretorio}\\{arquivo}.xlsm"
            
            # Verifica se o arquivo existe
            if not os.path.exists(diretorio_arquivo):
                raise FileNotFoundError(
                    f"Arquivo não encontrado: {diretorio_arquivo}\n"
                    f"Por favor, certifique-se de que o arquivo PRO_SAUDE.xlsm está localizado em: {diretorio}"
                )
            
            work_book = xlwings.Book(diretorio_arquivo)
            planilha = work_book.sheets[nome_aba]
            return work_book, planilha
        
        PRO_SAUDE.obter_planilha_e_aba = patched_obter_planilha_e_aba_immediate
    
    # PATCH para extrair_pdf_ateste
    if hasattr(PRO_SAUDE, 'extrair_pdf_ateste'):        
        def patched_extrair_pdf_ateste(page):
            import xlwings
            import pyautogui
            import re
            from tkinter import filedialog
            import customtkinter as ctk
            
            # USA O DIRETÓRIO CORRETO
            diretorio = get_pro_saude_directory()
            arquivo = 'PRO_SAUDE'
            diretorio_arquivo = f"{diretorio}\\{arquivo}.xlsm"

            work_book = xlwings.Book(diretorio_arquivo)
            planilha_dados = work_book.sheets['Dados']
            planilha_bancos = work_book.sheets['Bancos']

            funcao_extracao = PRO_SAUDE.find_information_ateste

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
                download.save_as(f'{diretorio}\\{nome_arquivo}')  # CORREÇÃO APLICADA
                page.locator("a[accesskey='V']").click()

                # CAMINHO CORRETO PARA O PDF - USA O DIRETÓRIO CERTO
                pdf_path = f'{diretorio}\\{nome_arquivo}'

            else:
                # Seleciona um arquivo local usando customtkinter
                root = ctk.CTk()
                root.withdraw()  # Oculta a janela principal do Tkinter
                pdf_path = filedialog.askopenfilename(title="Selecione o arquivo PDF da NFSE", filetypes=[("PDF files", "*.pdf")])
                root.destroy()

            # Extraindo texto do PDF da NFSE
            text_pages = PRO_SAUDE.extract_text_from_pdf(pdf_path)
            tables_data = PRO_SAUDE.extract_table_from_pdf(pdf_path)

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
                planilha_dados.range(f'E{i}').value = info.get('Nota de empenho', '')
                planilha_dados.range(f'P{i}').value = info.get('Banco', '')
                planilha_dados.range(f'Q{i}').value = info.get('Agência', '')
                planilha_dados.range(f'R{i}').value = conta_dv
                planilha_dados.range(f'G{i}').value = info.get('VALOR DE PAGTO BRUTO', '')
                planilha_dados.range(f'N{i}').value = info.get('VALOR PGTO LÍQUIDO', '')

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
                planilha_dados.range(f'H{linha_excel}').value = PRO_SAUDE.extrair_valor(row[1]) if len(row) > 1 else 0.0
                planilha_dados.range(f'I{linha_excel}').value = PRO_SAUDE.extrair_valor(row[2]) if len(row) > 2 else 0.0
                planilha_dados.range(f'J{linha_excel}').value = PRO_SAUDE.extrair_valor(row[3]) if len(row) > 3 else 0.0
                planilha_dados.range(f'K{linha_excel}').value = PRO_SAUDE.extrair_valor(row[4]) if len(row) > 4 else 0.0

                linha_excel += 1 # Só incrementa se a linha for válida

            PRO_SAUDE.update_bank_info(planilha_dados, planilha_bancos)
            work_book.save()
        
        PRO_SAUDE.extrair_pdf_ateste = patched_extrair_pdf_ateste
    
    # PATCH para extrair_pdf_ateste_cnpj_unico - REESCRITA COMPLETA
    if hasattr(PRO_SAUDE, 'extrair_pdf_ateste_cnpj_unico'):        
        def patched_extrair_pdf_ateste_cnpj_unico(page):
            import xlwings
            import pyautogui
            import re
            from tkinter import filedialog
            import customtkinter as ctk
            
            # USA O DIRETÓRIO CORRETO
            diretorio = get_pro_saude_directory()
            arquivo = 'PRO_SAUDE'
            diretorio_arquivo = f"{diretorio}\\{arquivo}.xlsm"
            
            work_book = xlwings.Book(diretorio_arquivo)
            planilha_dados = work_book.sheets['Dados']
            planilha_bancos = work_book.sheets['Bancos']

            funcao_extracao = PRO_SAUDE.find_information_ateste_cnpj_unico

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
                download.save_as(f'{diretorio}\\{nome_arquivo}')  # CORREÇÃO APLICADA
                page.locator("a[accesskey='V']").click()

                # CAMINHO CORRETO PARA O PDF - USA O DIRETÓRIO CERTO
                pdf_path = f'{diretorio}\\{nome_arquivo}'

            else:
                # Seleciona um arquivo local usando customtkinter
                root = ctk.CTk()
                root.withdraw()  # Oculta a janela principal do Tkinter
                pdf_path = filedialog.askopenfilename(title="Selecione o arquivo PDF da NFSE", filetypes=[("PDF files", "*.pdf")])
                root.destroy()

            # Extraindo texto do PDF da NFSE
            text_pages = PRO_SAUDE.extract_text_from_pdf(pdf_path)
            tables_data = PRO_SAUDE.extract_table_from_pdf(pdf_path)

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
                if not planilha_dados.range('P2').value:
                    planilha_dados.range('P2').value = info.get('Banco', '')
                if not planilha_dados.range('Q2').value:
                    planilha_dados.range('Q2').value = info.get('Agência', '')
                if not planilha_dados.range('R2').value:
                    planilha_dados.range('R2').value = conta_dv

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
                planilha_dados.range(f'W{linha_excel}').value = PRO_SAUDE.extrair_valor(row[1]) if len(row) > 1 else 0.0
                planilha_dados.range(f'X{linha_excel}').value = PRO_SAUDE.extrair_valor(row[2]) if len(row) > 2 else 0.0
                planilha_dados.range(f'Y{linha_excel}').value = PRO_SAUDE.extrair_valor(row[3]) if len(row) > 3 else 0.0
                planilha_dados.range(f'Z{linha_excel}').value = PRO_SAUDE.extrair_valor(row[4]) if len(row) > 4 else 0.0

                linha_excel += 1 # Só incrementa se a linha for válida

            PRO_SAUDE.update_bank_info(planilha_dados, planilha_bancos)
            work_book.save()
        
        PRO_SAUDE.extrair_pdf_ateste_cnpj_unico = patched_extrair_pdf_ateste_cnpj_unico
    
    # PATCH para extrair_pdf_NFSE - REESCRITA COMPLETA
    if hasattr(PRO_SAUDE, 'extrair_pdf_NFSE'):        
        def patched_extrair_pdf_NFSE(page):
            import pyautogui
            import re
            from tkinter import filedialog
            import customtkinter as ctk
            
            # USA O DIRETÓRIO CORRETO
            diretorio = get_pro_saude_directory()
            arquivo = 'PRO_SAUDE'
            
            work_book, planilha = PRO_SAUDE.obter_planilha_e_aba('Dados')

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
                download.save_as(f'{diretorio}\\{nome_arquivo}')  # CORREÇÃO APLICADA
                page.locator("a[accesskey='V']").click()

                # CAMINHO CORRETO PARA O PDF - USA O DIRETÓRIO CERTO
                pdf_path_nfse = f'{diretorio}\\{nome_arquivo}'

            else:
                # Seleciona um arquivo local usando customtkinter
                root = ctk.CTk()
                root.withdraw()  # Oculta a janela principal do Tkinter
                pdf_path_nfse = filedialog.askopenfilename(title="Selecione o arquivo PDF da NFSE", filetypes=[("PDF files", "*.pdf")])
                root.destroy()

            # Extraindo texto do PDF da NFSE
            text_nfse_pages = PRO_SAUDE.extract_text_from_pdf(pdf_path_nfse)

            # Encontrando informações específicas em cada página da NFSE e inserindo na planilha "Conferência"
            all_info_nfse = []
            for page_text in text_nfse_pages:
                info = PRO_SAUDE.find_information_NFSE(page_text)
                all_info_nfse.append(info)

            # Escrevendo os dados extraídos da NFSE na planilha
            for i, info in enumerate(all_info_nfse, start=linha_inicial):
                planilha.range(f'D{i}').value = info.get('CPF/CNPJ', '')
                planilha.range(f'F{i}').value = info.get('Número da Nota Fiscal', '')
                planilha.range(f'E{i}').value = info.get('Data de Geração da NFS-e', '')
                planilha.range(f'G{i}').value = info.get('Vl. Total dos Serviços', '')

            # Salvar o arquivo Excel
            work_book.save()
        
        PRO_SAUDE.extrair_pdf_NFSE = patched_extrair_pdf_NFSE
    
    # PATCH para liquidacao_pro_saude
    if hasattr(PRO_SAUDE, 'liquidacao_pro_saude'):
        original_liquidacao = PRO_SAUDE.liquidacao_pro_saude
        
        def patched_liquidacao_pro_saude(page):
            diretorio = get_pro_saude_directory()
            arquivo = 'PRO_SAUDE'
            
            # Define as variáveis globais
            PRO_SAUDE.diretorio = diretorio
            PRO_SAUDE.arquivo = arquivo
            
            return original_liquidacao(page)
        
        PRO_SAUDE.liquidacao_pro_saude = patched_liquidacao_pro_saude
    
    # PATCH para inclui_NS_edoc - REESCRITA COMPLETA
    if hasattr(PRO_SAUDE, 'inclui_NS_edoc'):        
        def patched_inclui_NS_edoc(page):
            import xlwings
            import pyautogui
            import os
            
            # USA O DIRETÓRIO CORRETO
            diretorio = get_pro_saude_directory()
            arquivo = 'PRO_SAUDE'
            
            work_book = xlwings.Book(f'{diretorio}\\{arquivo}.xlsm')
            planilha = work_book.sheets['Dados']       

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
                
                #upload - USA O DIRETÓRIO CORRETO
                arquivo_path = f"{diretorio}\\{nome_arquivo}"
                page.locator("label[class='labelInputFile']").set_input_files(arquivo_path)  
                page.locator(":text-is('Salvar')").click(timeout=60000)
                
                #exclui a NS - USA O DIRETÓRIO CORRETO
                caminho_arquivo = os.path.join(diretorio, nome_arquivo)
                if os.path.exists(caminho_arquivo):
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
            if hasattr(PRO_SAUDE, 'show_alert'):
                PRO_SAUDE.show_alert()
        
        PRO_SAUDE.inclui_NS_edoc = patched_inclui_NS_edoc
    
    # PATCH para conferir_liquidação - REESCRITA COMPLETA
    if hasattr(PRO_SAUDE, 'conferir_liquidação'):        
        def patched_conferir_liquidacao(page):
            import xlwings
            from datetime import datetime
            
            # USA O DIRETÓRIO CORRETO
            diretorio = get_pro_saude_directory()
            arquivo = 'PRO_SAUDE'
            diretorio_arquivo = f"{diretorio}\\{arquivo}.xlsm"
            
            work_book = xlwings.Book(diretorio_arquivo)
            planilha = work_book.sheets['Dados']

            ano_atual = datetime.now().year
            
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
                    processo_cabeçalho = page.locator("span[id='form_manterDocumentoHabil:processo_outputText']").text_content()
                    ateste_cabeçalho = page.locator("span[id='form_manterDocumentoHabil:dataAteste_outputText']").text_content()
                    valor_cabeçalho = page.locator("span[id='form_manterDocumentoHabil:valorPrincipalDocumento_output']").text_content()
                    CNPJ_cabeçalho = page.locator("span[id='form_manterDocumentoHabil:credorDevedor_output']").text_content()
                    planilha.range(f'D{linha}').value = CNPJ_cabeçalho
                    nome_credor = page.locator("span[id='form_manterDocumentoHabil:nomeCredorDevedor']").text_content()
                    planilha.range(f'C{linha}').value = nome_credor

                    ## Extrai informações dos Dados de Documentos de Origem
                    CNPJ_origem = page.locator("span[id='form_manterDocumentoHabil:tableDocsOrigem:0:emitenteDocOrigem_output']").text_content()
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

                    if page.locator("span[id='form_manterDocumentoHabil:lista_PCO:0:PCO_item_campoClassificacaoB_output_classificacao_contabil']").is_visible():
                        conta_passivo = page.locator("span[id='form_manterDocumentoHabil:lista_PCO:0:PCO_item_campoClassificacaoB_output_classificacao_contabil']").text_content()

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
                                vencimento_ddf025 = page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:dataVencimentoDeducao_outputText']").text_content()
                                cod_darf = page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:campoInscricaoA_deducao_output']").text_content()
                                cnpj_darf = page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_recolhedor_outputText']").text_content()
                                bc_darf = page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_base_calculo_output']").text_content()
                                bc_darf = bc_darf.replace('.','')
                                bc_darf = bc_darf.replace(',','.')
                                bc_darf = float(bc_darf)
                                valor_darf = page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:panelItemDeducao_valor_cabecalho']").text_content()
                                valor_darf = valor_darf.replace('.','')
                                valor_darf = valor_darf.replace(',','.')
                                valor_darf = float(valor_darf)
                                planilha.range(f'K{linha}').value= valor_darf

                                page.locator(f"input[id='form_manterDocumentoHabil:listaDeducao:{index}:btnPredoc']").click(timeout=6000)
                                page.wait_for_timeout(2000)

                                obs_darf_siafi=page.locator("textarea[id='form_manterDocumentoHabil:observacaoPredoc']").text_content()
                                page.locator("input[id='form_manterDocumentoHabil:btnRetornarPredoc']").click(timeout=6000)  
                                elemento.click()

                            if "DDR001 - RETENÇÕES DE IMPOSTOS RECOLHÍVEIS POR DAR" in texto:
                                elemento.click()

                                ug_dar= page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:inputUgPagadoraDeducao_output']").text_content()
                                vencimento_dar = page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:dataVencimentoDeducao_outputText']").text_content()
                                municipio_favorecido_dar = page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:campoInscricaoA_deducao_output']").text_content()
                                receita_dar = page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:campoInscricaoB_deducao_output']").text_content()
                                cnpj_dar = page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_recolhedor_outputText']").text_content()
                                valor_dar = page.locator(f"span[id='form_manterDocumentoHabil:listaDeducao:{index}:deducao_tabela_recolhimentos:0:deducao_tabela_recolhimentos_valor_recolhimento_output']").text_content()
                                valor_dar = valor_dar.replace('.','')
                                valor_dar = valor_dar.replace(',','.')
                                valor_dar = float(valor_dar)
                                planilha.range(f'M{linha}').value = valor_dar
                            
                                page.locator(f"input[id='form_manterDocumentoHabil:listaDeducao:{index}:btnPredoc']").click()
                                page.wait_for_timeout(2000)

                                ref_dar = page.locator("span[id='form_manterDocumentoHabil:referenciaDar_Output']").text_content()
                                municipio_nf_dar = page.locator("span[id='form_manterDocumentoHabil:municipioNFDar_outputText']").text_content()
                                numero_nf_dar = page.locator("span[id='form_manterDocumentoHabil:nfRecibo_outputText']").text_content()
                                emissao_nf_dar = page.locator("span[id='form_manterDocumentoHabil:dataEmissaoNfDar_outputText']").text_content()
                                aliquota_dar = page.locator("span[id='form_manterDocumentoHabil:aliquotaNFDar_output']").text_content()
                                valor_nf_dar = page.locator("span[id='form_manterDocumentoHabil:valorNfDar_output']").text_content()
                                obs_dar = page.locator("textarea[id='form_manterDocumentoHabil:observacaoPredoc']").text_content()
                                
                                page.locator("input[id='form_manterDocumentoHabil:btnRetornarPredoc']").click()
                                elemento.click()

                    else:
                        planilha.range(f'K{linha}').value = "ISENTO/IMUNE"
                        planilha.range(f'M{linha}').value = "ISENTO/IMUNE"

                    ## Extrai informações dos Dados de Pagamento
                    page.locator("input:text-is('Dados de Pagamento')").click()
                    page.wait_for_timeout(2000)

                    cnpj_pagamento_siafi = page.locator("span[id='form_manterDocumentoHabil:lista_DPgtoOB:0:codigoFavorecido_output']").text_content()
                    valor_pagamento_siafi = page.locator("span[id='form_manterDocumentoHabil:vlLiquidoPagar']").text_content()
                    valor_pagamento_siafi=valor_pagamento_siafi.replace('.','')
                    valor_pagamento_siafi=valor_pagamento_siafi.replace(',','.')
                    valor_pagamento_siafi=float(valor_pagamento_siafi)
                    planilha.range(f'N{linha}').value = valor_pagamento_siafi

                    valor_favorecido_siafi = page.locator("span[id='form_manterDocumentoHabil:lista_DPgtoOB:0:valorPredoc_output']").text_content()
                    valor_favorecido_siafi=valor_favorecido_siafi.replace('.','')
                    valor_favorecido_siafi=valor_favorecido_siafi.replace(',','.')
                    valor_favorecido_siafi=float(valor_favorecido_siafi)
                    
                    page.locator("input[id='form_manterDocumentoHabil:lista_DPgtoOB:0:btnPredoc']").click()
                    page.wait_for_timeout(2000)
                    
                    processo_pagamento_siafi = page.locator("span[id='form_manterDocumentoHabil:processoid_outputText']").text_content()
                    banco_siafi = page.locator("span[id='form_manterDocumentoHabil:favorecido_banco_outputText']").text_content()
                    planilha.range(f'P{linha}').value= banco_siafi
                    
                    agencia_siafi = page.locator("span[id='form_manterDocumentoHabil:favorecido_agencia_outputText']").text_content()
                    planilha.range(f'Q{linha}').value = agencia_siafi
                    
                    conta_siafi = page.locator("span[id='form_manterDocumentoHabil:favorecido_conta_outputText']").text_content()
                    planilha.range(f'R{linha}').value = conta_siafi
                    
                    osb_pagamento_siafi = page.locator("textarea[id='form_manterDocumentoHabil:observacaoPredoc']").text_content()
                    page.locator("input[id='form_manterDocumentoHabil:btnRetornarPredoc']").click()

                    ## Extrai informações do Centro de Cusatos
                    page.locator("input:text-is('Detacustos')").click()
                    page.wait_for_timeout(500)
                    page.locator("span[class='ng-tns-c47-9 pi pi-plus']").click()

                    custo_siafi = page.locator("div[class='p-text-center ng-star-inserted']").nth(-2).text_content()
                    planilha.range(f'O{linha}').value = custo_siafi
                    
                except:
                    continue
           
            # Salvar o arquivo Excel
            work_book.save()
            
            if hasattr(PRO_SAUDE, 'show_alert'):
                PRO_SAUDE.show_alert()
        
        PRO_SAUDE.conferir_liquidação = patched_conferir_liquidacao
        
finally:
    sys.argv = old_argv

# Aplicar patches globalmente quando em modo web para evitar overhead
_patches_applied = False

def patched_obter_planilha_e_aba(nome_aba):
    """Versão patcheada da função obter_planilha_e_aba que usa o diretório PRO_SAUDE"""
    import xlwings
    diretorio = get_pro_saude_directory()
    arquivo = 'PRO_SAUDE'
    diretorio_arquivo = f"{diretorio}\\{arquivo}.xlsm"
    
    # Verifica se o arquivo existe
    if not os.path.exists(diretorio_arquivo):
        raise FileNotFoundError(
            f"Arquivo não encontrado: {diretorio_arquivo}\n"
            f"Por favor, certifique-se de que o arquivo PRO_SAUDE.xlsm está localizado em: {diretorio}"
        )
    
    work_book = xlwings.Book(diretorio_arquivo)
    planilha = work_book.sheets[nome_aba]
    return work_book, planilha

def apply_global_patches():
    """Aplica patches globalmente uma única vez"""
    global _patches_applied
    if _patches_applied:
        return
    
    # Patch do pyautogui com mocks específicos
    import pyautogui
    pyautogui.prompt = mock_prompt
    pyautogui.confirm = mock_confirm
    pyautogui.alert = mock_alert
    
    # Patch do customtkinter
    try:
        import customtkinter
        customtkinter.CTk = lambda: SimpleMockTkinterRoot()
    except ImportError:
        pass
    
    # PATCH CRÍTICO: Substitui a função obter_planilha_e_aba
    if hasattr(PRO_SAUDE, 'obter_planilha_e_aba'):
        PRO_SAUDE.obter_planilha_e_aba = patched_obter_planilha_e_aba
    
    # Substitui a função show_alert se existir
    if hasattr(PRO_SAUDE, 'show_alert'):
        PRO_SAUDE.show_alert = mock_alert
    
    _patches_applied = True

# Wrappers otimizados - sem overhead de patch por execução
def extrair_pdf_NFSE(page, web_mode=False, **kwargs):
    if web_mode or is_web_mode():
        apply_global_patches()
        set_web_params(**kwargs)
        
        # Se um arquivo foi uploadado, configura o caminho
        uploaded_file_path = kwargs.get('uploaded_file_path')
        if uploaded_file_path:
            set_web_params(uploaded_file_path=uploaded_file_path)
    
    try:
        result = PRO_SAUDE.extrair_pdf_NFSE(page)
        
        # Limpa arquivo temporário após processamento
        if web_mode or is_web_mode():
            uploaded_file_path = get_web_param('uploaded_file_path')
            if uploaded_file_path and os.path.exists(uploaded_file_path):
                try:
                    os.remove(uploaded_file_path)
                except:
                    pass  # Ignora erros de limpeza
        
        return result
    except TypeError as e:
        if "missing 1 required positional argument: 'value'" in str(e):
            raise ValueError(
                f"Erro: Não foi possível acessar os dados da planilha Excel. "
                f"Verifique se o arquivo PRO_SAUDE.xlsm existe no diretório {get_pro_saude_directory()}."
            )
        else:
            raise e
    except Exception as e:
        # Limpa arquivo temporário mesmo em caso de erro
        if web_mode or is_web_mode():
            uploaded_file_path = get_web_param('uploaded_file_path')
            if uploaded_file_path and os.path.exists(uploaded_file_path):
                try:
                    os.remove(uploaded_file_path)
                except:
                    pass
        raise e

def extrair_pdf_ateste(page, web_mode=False, **kwargs):
    if web_mode or is_web_mode():
        apply_global_patches()
        set_web_params(**kwargs)
        
        # Suporte a upload para pdf_ateste também
        uploaded_file_path = kwargs.get('uploaded_file_path')
        if uploaded_file_path:
            set_web_params(uploaded_file_path=uploaded_file_path)
    
    try:
        result = PRO_SAUDE.extrair_pdf_ateste(page)
        
        # Limpa arquivo temporário após processamento
        if web_mode or is_web_mode():
            uploaded_file_path = get_web_param('uploaded_file_path')
            if uploaded_file_path and os.path.exists(uploaded_file_path):
                try:
                    os.remove(uploaded_file_path)
                except:
                    pass
        
        return result
    except TypeError as e:
        if "missing 1 required positional argument: 'value'" in str(e):
            raise ValueError(
                f"Erro: Não foi possível acessar os dados da planilha Excel. "
                f"Verifique se o arquivo PRO_SAUDE.xlsm existe no diretório {get_pro_saude_directory()}."
            )
        else:
            raise e
    except Exception as e:
        # Limpa arquivo temporário mesmo em caso de erro
        if web_mode or is_web_mode():
            uploaded_file_path = get_web_param('uploaded_file_path')
            if uploaded_file_path and os.path.exists(uploaded_file_path):
                try:
                    os.remove(uploaded_file_path)
                except:
                    pass
        raise e

def extrair_pdf_ateste_cnpj_unico(page, web_mode=False, **kwargs):
    if web_mode or is_web_mode():
        apply_global_patches()
        set_web_params(**kwargs)
        
        uploaded_file_path = kwargs.get('uploaded_file_path')
        if uploaded_file_path:
            set_web_params(uploaded_file_path=uploaded_file_path)
    
    try:
        result = PRO_SAUDE.extrair_pdf_ateste_cnpj_unico(page)
        
        if web_mode or is_web_mode():
            uploaded_file_path = get_web_param('uploaded_file_path')
            if uploaded_file_path and os.path.exists(uploaded_file_path):
                try:
                    os.remove(uploaded_file_path)
                except:
                    pass
        
        return result
    except TypeError as e:
        if "missing 1 required positional argument: 'value'" in str(e):
            raise ValueError(
                f"Erro: Não foi possível acessar os dados da planilha Excel. "
                f"Verifique se o arquivo PRO_SAUDE.xlsm existe no diretório {get_pro_saude_directory()}."
            )
        else:
            raise e
    except Exception as e:
        if web_mode or is_web_mode():
            uploaded_file_path = get_web_param('uploaded_file_path')
            if uploaded_file_path and os.path.exists(uploaded_file_path):
                try:
                    os.remove(uploaded_file_path)
                except:
                    pass
        raise e

def liquidacao_pro_saude(page, web_mode=False, **kwargs):
    if web_mode or is_web_mode():
        apply_global_patches()
        set_web_params(**kwargs)
    
    try:
        return PRO_SAUDE.liquidacao_pro_saude(page)
    except TypeError as e:
        error_msg = str(e)
        if "missing 1 required positional argument: 'value'" in error_msg:
            # Verificar se é especificamente problema com processo_edoc (célula B1)
            try:
                # Tentar acessar a planilha para verificar a célula B1
                work_book, planilha = PRO_SAUDE.obter_planilha_e_aba('Dados')
                processo_edoc = planilha.range('B1').value
                
                if processo_edoc is None or str(processo_edoc).strip() == '':
                    raise ValueError("Não está preenchido na planilha o número do processo (célula B1)")
                else:
                    # Se B1 não está vazio, o problema é outro
                    raise ValueError(
                        f"Erro: Não foi possível acessar os dados da planilha Excel. "
                        f"Verifique se o arquivo PRO_SAUDE.xlsm existe no diretório {get_pro_saude_directory()}."
                    )
            except Exception as excel_error:
                if "Não está preenchido na planilha" in str(excel_error):
                    raise excel_error  # Re-raise nossa mensagem específica
                else:
                    # Problema ao acessar Excel
                    raise ValueError(
                        f"Erro: Não foi possível acessar o arquivo Excel. "
                        f"Verifique se o arquivo PRO_SAUDE.xlsm existe no diretório {get_pro_saude_directory()}."
                    )
        else:
            raise e

def inclui_NS_edoc(page, web_mode=False, **kwargs):
    if web_mode or is_web_mode():
        apply_global_patches()
        set_web_params(**kwargs)
    return PRO_SAUDE.inclui_NS_edoc(page)

def conne(page, web_mode=False, **kwargs):
    if web_mode or is_web_mode():
        apply_global_patches()
    return PRO_SAUDE.conne(page)

# Tenta encontrar a função de conferir liquidação
conferir_liquidacao = None
for name in ['conferir_liquidacao', 'conferir_liquidação']:
    if hasattr(PRO_SAUDE, name):
        original_func = getattr(PRO_SAUDE, name)
        def make_conferir_wrapper(func):
            def conferir_liquidacao_wrapper(page, web_mode=False, **kwargs):
                if web_mode or is_web_mode():
                    apply_global_patches()
                
                try:
                    return func(page)
                except NameError as e:
                    error_msg = str(e)
                    if "name 'diretorio' is not defined" in error_msg:
                        raise ValueError(
                            f"Erro: Variáveis de diretório não foram definidas corretamente. "
                            f"Verifique se o arquivo PRO_SAUDE.xlsm está no diretório {get_pro_saude_directory()}."
                        )
                    elif "name 'arquivo' is not defined" in error_msg:
                        raise ValueError(
                            f"Erro: Variáveis de arquivo não foram definidas corretamente. "
                            f"Verifique se o arquivo PRO_SAUDE.xlsm está no diretório {get_pro_saude_directory()}."
                        )
                    else:
                        raise ValueError(f"Erro: Variável não definida - {error_msg}")
                except TypeError as e:
                    if "missing 1 required positional argument: 'value'" in str(e):
                        raise ValueError(
                            f"Erro: Não foi possível acessar os dados da planilha Excel. "
                            f"Verifique se o arquivo PRO_SAUDE.xlsm existe no diretório {get_pro_saude_directory()}."
                        )
                    else:
                        raise e
            return conferir_liquidacao_wrapper
        conferir_liquidacao = make_conferir_wrapper(original_func)
        break

if not conferir_liquidacao:
    def conferir_liquidacao(*args, **kwargs):
        raise NotImplementedError("Função conferir_liquidacao não encontrada")

# Exporta funções auxiliares úteis
obter_planilha_e_aba = getattr(PRO_SAUDE, 'obter_planilha_e_aba', None)
extract_text_from_pdf = getattr(PRO_SAUDE, 'extract_text_from_pdf', None)
extract_table_from_pdf = getattr(PRO_SAUDE, 'extract_table_from_pdf', None)
find_information_NFSE = getattr(PRO_SAUDE, 'find_information_NFSE', None)
find_information_ateste = getattr(PRO_SAUDE, 'find_information_ateste', None)
find_information_ateste_cnpj_unico = getattr(PRO_SAUDE, 'find_information_ateste_cnpj_unico', None)
extract_values_from_line = getattr(PRO_SAUDE, 'extract_values_from_line', None)
update_bank_info = getattr(PRO_SAUDE, 'update_bank_info', None)
extrair_valor = getattr(PRO_SAUDE, 'extrair_valor', None)

# Modificação para usar o diretório PRO_SAUDE na área de trabalho
if getattr(sys, 'frozen', False):
    diretorio = get_pro_saude_directory()
    arquivo = 'PRO_SAUDE'
elif __file__ and __name__ == '__main__':
    diretorio = get_pro_saude_directory()
    arquivo = 'PRO_SAUDE'
    # encontra a planilha Excel
    diretorio_arquivo = f"{diretorio}\\{arquivo}.xlsm"
    
    # Verifica se o arquivo existe e informa ao usuário
    if not os.path.exists(diretorio_arquivo):
        print(f"AVISO: Arquivo {diretorio_arquivo} não encontrado!")
        print(f"Por favor, certifique-se de que o arquivo PRO_SAUDE.xlsm está localizado em: {diretorio}")
    else:
        print(f"Arquivo encontrado: {diretorio_arquivo}")