import mysql.connector
import openpyxl

# Conexão com o MySQL
conn = mysql.connector.connect(
    host="localhost",
    user="root",
    password="",
    database="reeducandos_db"
)
cursor = conn.cursor()

# Carrega a planilha geral
wb = openpyxl.load_workbook("planilha_reeducandos.xlsx")
ws = wb.active

# Pula o cabeçalho
for row in ws.iter_rows(min_row=2, values_only=True):
    nome, cpf, telefone, frequencia, arquivo_origem = row

    sql = """
    INSERT INTO reeducandos (nome, cpf, telefone, frequencia, arquivo_origem)
    VALUES (%s, %s, %s, %s, %s)
    """
    cursor.execute(sql, (nome, cpf, telefone, frequencia, arquivo_origem))

conn.commit()
cursor.close()
conn.close()

print(" Dados inseridos com sucesso no banco de dados.")
