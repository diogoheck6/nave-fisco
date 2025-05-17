import psycopg2

def ler_clientes_banco():
    try:
        conexao = psycopg2.connect(
            host="",
            database="",
            user="",
            password="",
            port=5432
        )
        cursor = conexao.cursor()
        consulta = """
            SELECT
                est.INSCRFEDERAL AS cnpj,
                emp.NOMEEMPRESA AS razao_social
            FROM 
                EMPRESA emp
            INNER JOIN 
                ESTAB est ON emp.CODIGOEMPRESA = est.CODIGOEMPRESA
            WHERE 
                est.DATAENCERATIV = '2100-12-31'
                AND emp.NOMEEMPRESA NOT LIKE '%%Empresa%%Padr%%o%%Questor%%'
                AND est.SIGLAESTADO = 'SC'
            ORDER BY
                emp.NOMEEMPRESA
        """
        cursor.execute(consulta)
        resultado = cursor.fetchall()
        dicionario_clientes = {cnpj: razao_social for cnpj, razao_social in resultado}
        cursor.close()
        conexao.close()
        return dicionario_clientes
    except Exception as e:
        print(f"Ocorreu um erro ao consultar o banco: {e}")
        return {}

if __name__ == '__main__':
# Uso:
    clientes = ler_clientes_banco()
    print(clientes)