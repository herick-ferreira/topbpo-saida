import os

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'topbpo-secret-key-2025'
    UPLOAD_FOLDER = os.environ.get('UPLOAD_FOLDER') or 'uploads'
    PROCESSED_FOLDER = os.environ.get('PROCESSED_FOLDER') or 'processed'
    PORT = int(os.environ.get('PORT', 5000))
    ALLOWED_EXTENSIONS = {'xlsx', 'pdf'}
    MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB max file size

    @staticmethod
    def is_allowed_file(filename):
        return '.' in filename and filename.rsplit('.', 1)[1].lower() in Config.ALLOWED_EXTENSIONS