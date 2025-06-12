from flask import Flask, render_template, request, send_file, flash, redirect, url_for, jsonify
import os
import pandas as pd
from utils.processor import process_files, FileValidationError, ProcessingError
from config import Config
import re

app = Flask(__name__)
app.config.from_object(Config)

# Ensure directories exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['PROCESSED_FOLDER'], exist_ok=True)

def allowed_file(filename):
    return Config.is_allowed_file(filename)

def validate_numeric_input(value):
    """Valida se o valor é numérico permitindo vírgulas e pontos"""
    if not value or not value.strip():
        return False
    cleaned = value.strip().replace(',', '.').replace('.', '').replace('-', '')
    return cleaned.isnumeric()

def clean_numeric_value(value):
    """Limpa e converte valor numérico"""
    cleaned = "".join([v for v in value if v.isnumeric() or ( v in [',', '.'] and value.index(v) == len(value) - 3) or ( v == '-' and value.index(v) == 0) ]).strip()
    cleaned = cleaned.replace(',', '.', -1)
    return float(cleaned)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/health')
def health_check():
    return {'status': 'healthy', 'message': 'Application is running'}, 200

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        # Verificar se todos os arquivos foram enviados
        required_files = ['arquivo_xlsx', 'arquivo_pdf', 'parametros_xlsx']
        uploaded_files = {}
        
        for file_key in required_files:
            if file_key not in request.files:
                return jsonify({"error": f"Arquivo {file_key.replace('_', ' ')} não encontrado"}), 400
            
            file = request.files[file_key]
            if file.filename == '':
                return jsonify({"error": f"Nenhum arquivo selecionado para {file_key.replace('_', ' ')}"}), 400
            
            if not allowed_file(file.filename):
                return jsonify({"error": f"Tipo de arquivo inválido para {file_key.replace('_', ' ')}. Use apenas .xlsx ou .pdf"}), 400
            
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            uploaded_files[file_key] = filepath
          # Validar dados do formulário
        try:
            nome_cliente = request.form.get('nome_cliente', '').strip()
            if not nome_cliente:
                return jsonify({"error": "Nome do cliente é obrigatório"}), 400
            
            mes_ano = request.form.get('mes_ano', '').strip()
            if len(mes_ano) != 7 or "/" not in mes_ano:
                return jsonify({"error": "Formato de data inválido. Use o formato: 01/2025"}), 400
            
            transferencia = request.form.get('transferencia', '').strip()
            if not validate_numeric_input(transferencia):
                return jsonify({"error": "Valor de transferência inválido"}), 400
            transferencia = clean_numeric_value(transferencia)
            
            pagamentos = request.form.get('pagamentos', '').strip()
            if not validate_numeric_input(pagamentos):
                return jsonify({"error": "Valor de pagamentos inválido"}), 400
            pagamentos = clean_numeric_value(pagamentos)
            
            saldo_inicial = request.form.get('saldo_inicial', '').strip()
            if not validate_numeric_input(saldo_inicial):
                return jsonify({"error": "Valor de saldo inicial inválido"}), 400
            saldo_inicial = clean_numeric_value(saldo_inicial)
            
            saldo_final = request.form.get('saldo_final', '').strip()
            if not validate_numeric_input(saldo_final):
                return jsonify({"error": "Valor de saldo final inválido"}), 400
            saldo_final = clean_numeric_value(saldo_final)
            
            diario_n = request.form.get('diario_n', '').strip()
            if not diario_n.isnumeric():
                return jsonify({"error": "Número do diário deve ser um valor numérico"}), 400
            diario_n = int(diario_n)
            
        except Exception as e:
            return jsonify({"error": f"Erro ao validar dados do formulário: {str(e)}"}), 400
          # Processar arquivos
        try:
            output_file = process_files(
                uploaded_files['arquivo_xlsx'],
                uploaded_files['arquivo_pdf'],
                uploaded_files['parametros_xlsx'],
                nome_cliente,
                mes_ano,
                transferencia,
                pagamentos,
                saldo_inicial,
                saldo_final,
                diario_n,
                app.config['PROCESSED_FOLDER']
            )
            
            return send_file(output_file, as_attachment=True)
            
        except FileValidationError as e:
            return jsonify({"error": f"Erro de validação: {str(e)}"}), 400
        except ProcessingError as e:
            return jsonify({"error": f"Erro de processamento: {str(e)}"}), 500
        except Exception as e:
            app.logger.error(f"Erro inesperado: {str(e)}")
            return jsonify({"error": f"Erro inesperado: {str(e)}"}), 500
        
        finally:
            # Limpar arquivos temporários
            for filepath in uploaded_files.values():
                try:
                    if os.path.exists(filepath):
                        os.remove(filepath)
                except:
                    pass
        
    except Exception as e:
        app.logger.error(f"Error processing files: {str(e)}")
        return jsonify({"error": f"Erro ao processar arquivos: {str(e)}"}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port, debug=False)