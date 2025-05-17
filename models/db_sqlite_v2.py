import sqlite3
import os
from datetime import datetime
import pytz
from dateutil import parser

br_timezone = pytz.timezone('America/Sao_Paulo')

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
                    data_inserida DATETIME DEFAULT CURRENT_TIMESTAMP,
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
        current_time = datetime.now(br_timezone)
        
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT id FROM clientes WHERE cnpj = ?', (cnpj,))
            cliente_id = cursor.fetchone()

            if not cliente_id:
                print("Cliente não encontrado.")
                return

            try:
                cursor.execute('''
                    INSERT INTO competencias (cliente_id, competencia, id_cte, operacao, data_inserida)
                    VALUES (?, ?, ?, ?, ?)
                    ON CONFLICT(cliente_id, competencia, operacao)
                    DO UPDATE SET id_cte=excluded.id_cte, data_inserida=excluded.data_inserida
                ''', (cliente_id[0], competencia, id_cte, operacao, current_time))
                conn.commit()

            except sqlite3.IntegrityError as e:
                print(f"Erro de integridade: {e}")


    def deve_pular_solicitacao_cte(self, cnpj, competencia):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT MAX(data_inserida) FROM competencias
                JOIN clientes ON competencias.cliente_id = clientes.id
                WHERE clientes.cnpj = ? AND competencia = ?
            ''', (cnpj, competencia))
            last_entry_date = cursor.fetchone()[0]

            if last_entry_date:
                # Usa o dateutil para lidar com o sufixo de fuso-horário
                last_entry_date = parser.isoparse(last_entry_date)
                return last_entry_date.date() == datetime.now(pytz.timezone('America/Sao_Paulo')).date()
            return False

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
                cursor.execute('DELETE FROM competencias WHERE cliente_id = ?', (cliente_id,))
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


def processar_dados(cnpj, competencia, operacao, id_cte=None):
    obj_base = NumerosSolicitacaoCTE()

    if obj_base.deve_pular_solicitacao_cte(cnpj, competencia):
        print("Pular solicitação de CTE. Isso já foi realizado hoje para esta competência.")
        return

    # Lógica para processar a chamada solicitar_ctes
    print("Chamando solicitar_ctes...")

    obj_base.adicionar_competencia(cnpj, competencia, operacao, id_cte)
    print("Competência processada e armazenada com sucesso.")


