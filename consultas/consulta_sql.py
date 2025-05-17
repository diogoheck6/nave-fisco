from firebird_connection.connection import FirebirdConnectionODBC
from queries.lctofissai_query import get_lctofissai_query
from queries.lctofisent_query import get_lctofisent_query

def fazer_consulta_sql_entradas(insc_federal, start_date, end_date):
    # Criar uma instância da classe de conexão
    fb_conn = FirebirdConnectionODBC()
    connection = fb_conn.connect()

    if connection:
        try:
            # Execute a consulta
            cursor = connection.cursor()
            query = get_lctofisent_query()

            # Execute a consulta com parâmetros
            cursor.execute(query, (insc_federal, insc_federal, start_date, end_date))
            rows = cursor.fetchall()
            
            # Processe os resultados
            for row in rows:
                # print(row)
                pass
        finally:
            fb_conn.disconnect(connection)

    
def fazer_consulta_sql_saidas(insc_federal, start_date, end_date):
    # Criar uma instância da classe de conexão
    fb_conn = FirebirdConnectionODBC()
    connection = fb_conn.connect()

    if connection:
        try:
            # Execute a consulta
            cursor = connection.cursor()
            query = get_lctofissai_query()

            # Execute a consulta com parâmetros
            cursor.execute(query, (insc_federal, insc_federal, start_date, end_date))
            rows = cursor.fetchall()

            # Obtenha o cabeçalho (nomes das colunas)
            header = [column[0] for column in cursor.description]
            # print("Cabeçalho:", header)

            # Processe os resultados
            for row in rows:
                # print(row)
                pass
        finally:
            fb_conn.disconnect(connection)
