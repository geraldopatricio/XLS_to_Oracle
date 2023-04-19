import cx_Oracle
import xlrd

# Conexão com o banco de dados Oracle
dsn_tns = cx_Oracle.makedsn('host', 'port', service_name='service_name')  # substitua pelos valores corretos
conn = cx_Oracle.connect(user='usuario', password='senha', dsn=dsn_tns) # substitua pelos valores corretos

# Leitura do arquivo Excel
arquivo_excel = xlrd.open_workbook('c:/pastapublica/bra.xls')
planilha = arquivo_excel.sheet_by_index(0)
conteudo = planilha.cell_value(rowx=1, colx=1)
material = planilha.cell_value(rowx=8, colx=3)

# Inserção no banco de dados
cursor = conn.cursor()
cursor.execute("INSERT INTO usuario (conteudo, material) VALUES (:conteudo, :material)", [conteudo, material])
conn.commit()
cursor.close()

# Encerramento da conexão com o banco de dados
conn.close()
