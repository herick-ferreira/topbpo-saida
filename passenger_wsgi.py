import sys
import os

# Adiciona o diretório da aplicação ao path
sys.path.insert(0, os.path.dirname(__file__))

from app import app as application

if __name__ == "__main__":
    application.run()