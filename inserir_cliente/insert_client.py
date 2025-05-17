import os
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from dotenv import load_dotenv

from create_tables.create_tables_ import Cliente, Base

# Carregar variáveis de ambiente
load_dotenv()

# Configuração de conexão com PostgreSQL
postgres_user = os.getenv('POSTGRES_USER')
postgres_password = os.getenv('POSTGRES_PASSWORD')
postgres_host = os.getenv('POSTGRES_HOST', 'localhost')
postgres_port = os.getenv('POSTGRES_PORT', '5432')
postgres_db = os.getenv('POSTGRES_DB')

# Criar Engine
engine = create_engine(f'postgresql+psycopg2://{postgres_user}:{postgres_password}@{postgres_host}:{postgres_port}/{postgres_db}')
Session = sessionmaker(bind=engine)

def inserir_cliente(razao_social, cnpj):
    """
    Função para inserir um cliente no banco de dados.
    
    :param razao_social: Nome da empresa
    :param cnpj: CNPJ da empresa
    """
    # Criar uma nova sessão
    session = Session()
    try:
        # Criar o objeto Cliente
        novo_cliente = Cliente(razao_social=razao_social, cnpj=cnpj)
        
        # Adicionar à sessão
        session.add(novo_cliente)
        
        # Commitar a transação
        session.commit()
        
        print(f"Cliente '{razao_social}' inserido com sucesso.")
    except Exception as e:
        # Caso ocorra erro, rollback para evitar inconsistência
        session.rollback()
        print(f"Erro ao inserir cliente: {e}")
    finally:
        # Fechar a sessão para liberar recursos
        session.close()

#