import psycopg2

class PostgresConnection:
    def __init__(self):
        self.host = ""
        self.database = ""
        self.user = ""
        self.password = ""
        self.port = 5432

    def connect(self):
        try:
            connection = psycopg2.connect(
                host=self.host,
                database=self.database,
                user=self.user,
                password=self.password,
                port=self.port
            )
            print("Conexão estabelecida com sucesso ao PostgreSQL!")
            return connection
        except Exception as e:
            print(f"Erro ao conectar ao banco de dados PostgreSQL: {e}")
            return None

    def disconnect(self, connection):
        if connection:
            connection.close()
            print("Conexão fechada com sucesso ao PostgreSQL.")


