from flask import Flask, render_template, request, redirect, jsonify, send_from_directory, make_response, send_file, abort, flash, url_for
import pandas as pd
import os
import numpy as np
import logging
from pathlib import Path
import subprocess
import sys
from playwright.sync_api import sync_playwright
import ADAPTADOR_PRO_SAUDE as PRO_SAUDE
import threading
import time
import uuid

app = Flask(__name__)
app.secret_key = 'your-secret-key-here-change-in-production'
app.logger.setLevel(logging.INFO)
logging.basicConfig(level=logging.INFO)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Configurações
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Dicionário para acompanhar status das operações
operation_status = {}

def get_browser_connection():
    """Estabelece conexão com o navegador via CDP"""
    try:
        playwright = sync_playwright().start()
        browser = playwright.chromium.connect_over_cdp("http://localhost:9230")
        context = browser.contexts[0]
        pages = context.pages
        return playwright, browser, context, pages
    except Exception as e:
        app.logger.error(f"Erro ao conectar com o navegador: {e}")
        return None, None, None, None

def update_operation_status(operation_id, status, progress, completed=False, error=None):
    """Helper para atualizar status das operações"""
    operation_status[operation_id] = {
        'status': status,
        'progress': progress,
        'completed': completed,
        'error': error
    }

def execute_function_with_browser(func_name, operation_id, page_filter=None, **kwargs):
    """Executa uma função que precisa de conexão com o navegador"""
    try:
        update_operation_status(operation_id, 'Conectando com navegador...', 10)
        
        # Definir modo web globalmente
        PRO_SAUDE.set_web_mode(True)
        
        playwright, browser, context, pages = get_browser_connection()
        
        if not pages:
            update_operation_status(operation_id, 'Erro: Nenhuma página do navegador encontrada', 0, True, 
                                  'Certifique-se de que o navegador está aberto com CDP ativo.')
            return {"error": "Nenhuma página do navegador encontrada"}
        
        update_operation_status(operation_id, 'Navegador conectado. Executando função...', 30)
        
        # Seleciona a página adequada
        if page_filter:
            page = None
            for p in pages:
                if page_filter.lower() in p.url.lower():
                    page = p
                    break
            if not page:
                page = pages[0]
        else:
            page = pages[0]
        
        update_operation_status(operation_id, 'Processando dados...', 50)
        
        # Executa a função específica
        if func_name == "extrair_pdf_NFSE":
            PRO_SAUDE.extrair_pdf_NFSE(
                page, 
                web_mode=True,
                linha_inicial=kwargs.get('linha_inicial_param', 2),
                escolha=kwargs.get('escolha_param', 'Baixar'),
                uploaded_file_path=kwargs.get('uploaded_file_path')
            )
        elif func_name == "extrair_pdf_ateste":
            PRO_SAUDE.extrair_pdf_ateste(
                page, 
                web_mode=True,
                escolha=kwargs.get('escolha_param', 'Baixar'),
                uploaded_file_path=kwargs.get('uploaded_file_path')
            )
        elif func_name == "extrair_pdf_ateste_cnpj_unico":
            PRO_SAUDE.extrair_pdf_ateste_cnpj_unico(
                page, 
                web_mode=True,
                escolha=kwargs.get('escolha_param', 'Baixar'),
                uploaded_file_path=kwargs.get('uploaded_file_path')
            )
        elif func_name == "liquidacao_pro_saude":
            PRO_SAUDE.liquidacao_pro_saude(
                page, 
                web_mode=True,
                linha_inicial=kwargs.get('linha_inicial_param', 2),
                linha_final=kwargs.get('linha_final_param', 2)
            )
        elif func_name == "conne":
            PRO_SAUDE.conne(page, web_mode=True)
        elif func_name == "conferir_liquidacao":
            PRO_SAUDE.conferir_liquidacao(page, web_mode=True)
        elif func_name == "inclui_NS_edoc":
            PRO_SAUDE.inclui_NS_edoc(
                page, 
                web_mode=True,
                linha_inicial=kwargs.get('linha_inicial_param', 2),
                linha_final=kwargs.get('linha_final_param', 2)
            )
        else:
            raise ValueError(f"Função {func_name} não reconhecida")
        
        update_operation_status(operation_id, 'Concluído com sucesso!', 100, True)
        return {"success": "Função executada com sucesso!"}
        
    except Exception as e:
        error_msg = f"Erro ao executar função: {str(e)}"
        app.logger.error(f"Erro em {func_name}: {e}", exc_info=True)
        update_operation_status(operation_id, f'Erro: {str(e)}', 0, True, str(e))
        return {"error": error_msg}
    finally:
        # Cleanup
        if 'browser' in locals() and browser:
            try:
                browser.close()
            except:
                pass
        if 'playwright' in locals() and playwright:
            try:
                playwright.stop()
            except:
                pass

def allowed_file(filename):
    """Verifica se o arquivo é permitido"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() == 'pdf'

def save_uploaded_file(file):
    """Salva arquivo uploadado e retorna o caminho"""
    if file and allowed_file(file.filename):
        # Nome único para evitar conflitos
        filename = f"uploaded_{int(time.time())}_{file.filename}"
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(file_path)
        return file_path
    return None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/pro_saude')
def pro_saude():
    return render_template('pro_saude.html')

@app.route('/status/<operation_id>')
def get_status(operation_id):
    """Retorna o status de uma operação"""
    status = operation_status.get(operation_id, {
        'status': 'Operação não encontrada',
        'progress': 0,
        'completed': True,
        'error': 'Operação não encontrada'
    })
    return jsonify(status)

@app.route('/status_page/<operation_id>')
def status_page(operation_id):
    """Página de status para acompanhar operação"""
    return render_template('status.html', operation_id=operation_id)

# === ROTAS SIAFI ===

@app.route('/incdh', methods=['GET', 'POST'])
def incdh():
    """Incluir Documento Hábil (Liquidação Pro Saúde)"""
    if request.method == 'POST':
        try:
            linha_inicial = int(request.form.get('linha_inicial', 2))
            linha_final = int(request.form.get('linha_final', linha_inicial))
            
            operation_id = str(uuid.uuid4())
            
            def run_operation():
                execute_function_with_browser(
                    "liquidacao_pro_saude", 
                    operation_id,
                    "siafi",
                    linha_inicial_param=linha_inicial,
                    linha_final_param=linha_final
                )
            
            thread = threading.Thread(target=run_operation)
            thread.daemon = True
            thread.start()
            
            return redirect(url_for('status_page', operation_id=operation_id))
            
        except ValueError as e:
            flash(f'Erro nos parâmetros: {e}', 'error')
            return render_template('form_incdh.html')
    
    return render_template('form_incdh.html')

@app.route('/conne')
def conne():
    """Consultar Nota de Empenho"""
    operation_id = str(uuid.uuid4())
    
    def run_operation():
        execute_function_with_browser("conne", operation_id, "siafi")
    
    thread = threading.Thread(target=run_operation)
    thread.daemon = True
    thread.start()
    
    return redirect(url_for('status_page', operation_id=operation_id))

@app.route('/conferir_liquidacao')
def conferir_liquidacao():
    """Conferir Liquidação"""
    operation_id = str(uuid.uuid4())
    
    def run_operation():
        execute_function_with_browser("conferir_liquidacao", operation_id, "siafi")
    
    thread = threading.Thread(target=run_operation)
    thread.daemon = True
    thread.start()
    
    return redirect(url_for('status_page', operation_id=operation_id))

# === ROTAS EDOC ===

@app.route('/pdf_NFSE', methods=['GET', 'POST'])
def pdf_nfse():
    """Extrair dados da NFSE"""
    if request.method == 'POST':
        try:
            linha_inicial = int(request.form.get('linha_inicial', 2))
            escolha = request.form.get('escolha', 'Baixar')
            
            uploaded_file_path = None
            
            # Processa upload se necessário
            if escolha == 'Selecionar':
                if 'pdf_file' not in request.files:
                    flash('Nenhum arquivo foi selecionado', 'error')
                    return render_template('form_nfse.html')
                
                file = request.files['pdf_file']
                if file.filename == '':
                    flash('Nenhum arquivo foi selecionado', 'error')
                    return render_template('form_nfse.html')
                
                uploaded_file_path = save_uploaded_file(file)
                if not uploaded_file_path:
                    flash('Arquivo inválido. Por favor, selecione um arquivo PDF.', 'error')
                    return render_template('form_nfse.html')
            
            operation_id = str(uuid.uuid4())
            
            def run_operation():
                execute_function_with_browser(
                    "extrair_pdf_NFSE", 
                    operation_id,
                    "edoc", 
                    linha_inicial_param=linha_inicial, 
                    escolha_param=escolha,
                    uploaded_file_path=uploaded_file_path
                )
            
            thread = threading.Thread(target=run_operation)
            thread.daemon = True
            thread.start()
            
            return redirect(url_for('status_page', operation_id=operation_id))
            
        except ValueError as e:
            flash(f'Erro nos parâmetros: {e}', 'error')
            return render_template('form_nfse.html')
    
    return render_template('form_nfse.html')

@app.route('/pdf_ateste', methods=['GET', 'POST'])
def pdf_ateste():
    """Extrair dados do Relatório de Ateste"""
    if request.method == 'POST':
        try:
            escolha = request.form.get('escolha', 'Baixar')
            uploaded_file_path = None
            
            if escolha == 'Selecionar':
                if 'pdf_file' in request.files:
                    file = request.files['pdf_file']
                    if file.filename != '':
                        uploaded_file_path = save_uploaded_file(file)
                        if not uploaded_file_path:
                            flash('Arquivo inválido. Por favor, selecione um arquivo PDF.', 'error')
                            return render_template('form_ateste.html')
            
            operation_id = str(uuid.uuid4())
            
            def run_operation():
                execute_function_with_browser(
                    "extrair_pdf_ateste", 
                    operation_id,
                    "edoc", 
                    escolha_param=escolha,
                    uploaded_file_path=uploaded_file_path
                )
            
            thread = threading.Thread(target=run_operation)
            thread.daemon = True
            thread.start()
            
            return redirect(url_for('status_page', operation_id=operation_id))
            
        except Exception as e:
            flash(f'Erro: {e}', 'error')
            return render_template('form_ateste.html')
    
    return render_template('form_ateste.html')

@app.route('/pdf_ateste_cnpj')
def pdf_ateste_cnpj():
    """Extrair Relatório (CNPJ único)"""
    operation_id = str(uuid.uuid4())
    
    def run_operation():
        execute_function_with_browser(
            "extrair_pdf_ateste_cnpj_unico", 
            operation_id,
            "edoc", 
            escolha_param='Baixar'
        )
    
    thread = threading.Thread(target=run_operation)
    thread.daemon = True
    thread.start()
    
    return redirect(url_for('status_page', operation_id=operation_id))

@app.route('/ns_edoc', methods=['GET', 'POST'])
def ns_edoc():
    """Incluir NS ao processo eDoc"""
    if request.method == 'POST':
        try:
            linha_inicial = int(request.form.get('linha_inicial', 2))
            linha_final = int(request.form.get('linha_final', linha_inicial))
            
            operation_id = str(uuid.uuid4())
            
            def run_operation():
                execute_function_with_browser(
                    "inclui_NS_edoc", 
                    operation_id,
                    "edoc", 
                    linha_inicial_param=linha_inicial, 
                    linha_final_param=linha_final
                )
            
            thread = threading.Thread(target=run_operation)
            thread.daemon = True
            thread.start()
            
            return redirect(url_for('status_page', operation_id=operation_id))
            
        except ValueError as e:
            flash(f'Erro nos parâmetros: {e}', 'error')
            return render_template('form_ns_edoc.html')
    
    return render_template('form_ns_edoc.html')

# === ROTAS UTILITÁRIAS ===

@app.route('/static/<path:filename>')
def custom_static(filename):
    """Serve arquivos estáticos"""
    return send_from_directory('static', filename)

@app.route('/health')
def health_check():
    """Verifica se a aplicação está funcionando"""
    active_ops = len([op for op in operation_status.values() if not op.get('completed', True)])
    return jsonify({
        'status': 'healthy',
        'timestamp': time.time(),
        'active_operations': active_ops,
        'total_operations': len(operation_status)
    })

@app.route('/cleanup')
def cleanup_operations():
    """Remove operações antigas do status"""
    global operation_status
    
    operations_to_remove = []
    current_time = time.time()
    
    for op_id, op_data in operation_status.items():
        if op_data.get('completed', False):
            operations_to_remove.append(op_id)
    
    removed_count = len(operations_to_remove)
    for op_id in operations_to_remove:
        operation_status.pop(op_id, None)
    
    # Limpa arquivos de upload antigos também
    cleanup_old_uploads()
    
    return jsonify({
        'message': f'Removidas {removed_count} operações antigas',
        'active_operations': len([op for op in operation_status.values() if not op.get('completed', True)])
    })

def cleanup_old_uploads():
    """Remove arquivos de upload antigos (mais de 1 hora)"""
    try:
        current_time = time.time()
        for filename in os.listdir(UPLOAD_FOLDER):
            file_path = os.path.join(UPLOAD_FOLDER, filename)
            if os.path.isfile(file_path):
                # Remove arquivos mais antigos que 1 hora
                if current_time - os.path.getctime(file_path) > 3600:
                    os.remove(file_path)
    except Exception as e:
        app.logger.error(f"Erro ao limpar uploads antigos: {e}")

# === TRATAMENTO DE ERROS ===

@app.errorhandler(404)
def not_found(error):
    return render_template('error.html', 
                         error="Página não encontrada", 
                         error_code=404), 404

@app.errorhandler(500)
def internal_error(error):
    return render_template('error.html', 
                         error="Erro interno do servidor", 
                         error_code=500), 500

@app.errorhandler(413)
def file_too_large(error):
    flash('Arquivo muito grande. Tamanho máximo: 16MB', 'error')
    return redirect(request.url)

def initialize_app():
    """Inicialização da aplicação"""
    PRO_SAUDE.set_web_mode(True)
    app.logger.info("Aplicação iniciada em modo web")
    
    # Limpa uploads antigos na inicialização
    cleanup_old_uploads()

if __name__ == "__main__":
    # Inicialização da aplicação
    initialize_app()
    
    # Configurações
    debug_mode = os.environ.get('FLASK_DEBUG', 'True').lower() == 'true'
    host = os.environ.get('FLASK_HOST', '127.0.0.1')
    port = int(os.environ.get('FLASK_PORT', 5000))
    
    app.logger.info(f"Iniciando aplicação em {host}:{port} (debug={debug_mode})")
    
    # Executa a aplicação
    app.run(
        debug=debug_mode, 
        use_reloader=False,  # Importante: desabilita reloader para evitar problemas com threads
        host=host,
        port=port,
        threaded=True
    )