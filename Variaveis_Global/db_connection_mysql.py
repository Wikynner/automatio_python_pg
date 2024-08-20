import mysql.connector

# Variável global para a conexão
db_connection = None

def connect_to_database():
    global db_connection
    if db_connection is None:
        try:
            db_connection = mysql.connector.connect(
                host='localhost',
                user='root',
                password='',
                database='alertas_bi'
            )
            print("Conexão estabelecida com sucesso.")
        except mysql.connector.Error as err:
            print(f"Erro ao conectar ao banco de dados: {err}")
    else:
        print("Conexão já está estabelecida.")

def close_database_connection():
    global db_connection
    if db_connection is not None and db_connection.is_connected():
        db_connection.close()
        db_connection = None
        print("Conexão fechada com sucesso.")
    else:
        print("Nenhuma conexão ativa encontrada.")

def get_connection():
    global db_connection
    if db_connection is None:
        connect_to_database()  # Estabeleça a conexão se não estiver conectada
    return db_connection

#Remova ou comente as linhas abaixo para evitar execuções desnecessárias ao importar o módulo
# connect_to_database()
# close_database_connection()
