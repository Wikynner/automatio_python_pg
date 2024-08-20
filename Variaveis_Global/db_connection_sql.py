import pyodbc

# Variável global para a conexão
db_connection = None

def connect_to_database():
    global db_connection
    if db_connection is None:
        try:
            # String de conexão para SQL Server
            connection_string = (
                'DRIVER={SQL Server};'
                'SERVER=192.168.100.105;'
                'DATABASE=agrimanager;'
                'UID=controladoria;'
                'PWD=Senha@2022'
            )
            db_connection = pyodbc.connect(connection_string)
            print("Conexão estabelecida com sucesso.")
        except pyodbc.Error as err:
            print(f"Erro ao conectar ao banco de dados: {err}")
    else:
        print("Conexão já está estabelecida.")

def close_database_connection():
    global db_connection
    if db_connection is not None:
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

# # Testando a conexão
# connect_to_database()
# close_database_connection()
