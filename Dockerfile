# Use Python 3.11 slim image
FROM python:3.11-slim

# Define o diretório de trabalho
WORKDIR /app

# Instala dependências do sistema
RUN apt-get update && apt-get install -y \
    gcc \
    g++ \
    && rm -rf /var/lib/apt/lists/*

# Copia arquivo de dependências primeiro (para cache do Docker)
COPY requirements.txt .

# Atualiza pip e instala dependências Python
RUN pip install --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Copia todo o código da aplicação
COPY . .

# Cria diretórios necessários e define permissões
RUN mkdir -p uploads processed && \
    chmod -R 755 uploads processed

# Criar usuário não-root para segurança
RUN adduser --disabled-password --gecos '' appuser && \
    chown -R appuser:appuser /app
USER appuser

# Expõe a porta que a aplicação vai usar
EXPOSE $PORT

# Define variáveis de ambiente
ENV FLASK_APP=app.py
ENV FLASK_ENV=production
ENV PYTHONPATH=/app

# Comando para iniciar a aplicação (Render fornece a variável PORT)
CMD gunicorn --bind 0.0.0.0:$PORT --workers 2 --timeout 300 app:app