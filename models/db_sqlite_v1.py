import sqlite3
import os

class NumerosSolicitacaoCTE:

    def __init__(self, db_path=None):
        if not db_path:
            db_path = os.path.join(os.getcwd(), 'NumerosSolicitacaoCTE.db')
        self.db_path = db_path
        self._criar_tabelas()

    def _criar_tabelas(self):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS clientes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    cnpj TEXT NOT NULL UNIQUE,
                    razao_social TEXT NOT NULL
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS competencias (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    cliente_id INTEGER NOT NULL,
                    competencia TEXT NOT NULL,
                    id_cte TEXT,
                    operacao TEXT NOT NULL,
                    FOREIGN KEY(cliente_id) REFERENCES clientes(id),
                    UNIQUE(cliente_id, competencia, operacao)
                )
            ''')

    def adicionar_cliente(self, cnpj, razao_social):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT OR IGNORE INTO clientes (cnpj, razao_social) VALUES (?, ?)
            ''', (cnpj, razao_social))
            conn.commit()

    def adicionar_competencia(self, cnpj, competencia, operacao, id_cte=None):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT id FROM clientes WHERE cnpj = ?', (cnpj,))
            cliente_id = cursor.fetchone()

            if not cliente_id:
                print("Cliente não encontrado.")
                return

            try:
                cursor.execute('''
                    INSERT INTO competencias (cliente_id, competencia, id_cte, operacao)
                    VALUES (?, ?, ?, ?)
                    ON CONFLICT(cliente_id, competencia, operacao)
                    DO UPDATE SET id_cte=excluded.id_cte
                ''', (cliente_id[0], competencia, id_cte, operacao))
                conn.commit()

            except sqlite3.IntegrityError as e:
                print(f"Erro de integridade: {e}")

    def id_existe(self, cte_id, operacao):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT 1 FROM competencias WHERE id_cte = ? AND operacao = ?
            ''', (cte_id, operacao))
            return cursor.fetchone() is not None
        
    def deletar_cliente_por_id(self, cliente_id):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            try:
                # Deletar competências associadas ao cliente
                cursor.execute('DELETE FROM competencias WHERE cliente_id = ?', (cliente_id,))
                # Deletar o cliente
                cursor.execute('DELETE FROM clientes WHERE id = ?', (cliente_id,))
                conn.commit()
                print(f"Cliente e suas competências deletados para o ID: {cliente_id}")

            except sqlite3.Error as e:
                print(f"Erro ao deletar cliente e competências: {e}")

    def recuperar_ids(self, cnpj, competencia, operacao=None):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            sql = '''
                SELECT id_cte FROM competencias
                JOIN clientes ON competencias.cliente_id = clientes.id
                WHERE clientes.cnpj = ? AND competencia = ?
            '''
            params = [cnpj, competencia]
            if operacao:
                sql += ' AND operacao = ?'
                params.append(operacao)

            cursor.execute(sql, params)
            ids = cursor.fetchall()
            return [i[0] for i in ids if i[0] is not None]

    def listar_competencias(self, cnpj):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT DISTINCT competencia FROM competencias
                JOIN clientes ON competencias.cliente_id = clientes.id
                WHERE clientes.cnpj = ?
            ''', (cnpj,))
            competencias = cursor.fetchall()
            return [c[0] for c in competencias]


if __name__ == '__main__':
    pass