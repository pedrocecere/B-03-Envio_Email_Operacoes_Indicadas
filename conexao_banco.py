from pydantic import BaseSettings
from dotenv import load_dotenv

# Carrega as variáveis do arquivo .env
load_dotenv()

class Settings(BaseSettings):
    db_server: str
    db_name: str
    db_user: str
    db_password: str
    db_trusted_connection: str = "no"  # Define um valor padrão caso não esteja no .env

    class Config:
        env_file = ".env"  # Indica o arquivo .env como fonte das variáveis

# Instancia as configurações
settings = Settings()
