# Importe a variável global de config.py
from config import consulta_query

def execute_query():
    global consulta_query
    if consulta_query:
        print(f"Executando a consulta: {consulta_query}")
        # Aqui você pode conectar ao banco de dados e executar a consulta
        # Exemplo:
        # cursor.execute(consulta_query)
    else:
        print("select * from f_adtos ")

# Teste a execução da consulta
execute_query()
