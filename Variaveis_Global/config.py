# main.py

# Importe a variável global de config.py
from config import consulta_query

# Defina a consulta SQL que deseja armazenar globalmente
consulta_query = """
WITH RECEBER AS (
            SELECT *,
            CASE 
                WHEN VENCIDO = 'S' THEN TIMESTAMPDIFF(DAY, DATA_VCTO, CAST(SYSDATE() AS DATE))
                ELSE 0
            END DIAS_VENCIDO
            FROM f_titulos_receber ftr 
            WHERE 1=1
            AND PESSOA_FIN NOT IN ('MAURO FERNANDO SCHAEDLER','IEDA WEBLER SCHAEDLER')
        )
        SELECT 
            PESSOA_FIN AS Fornecedor,
            MOEDA_ORIGINAL AS MOEDA,
            CASE 
                WHEN DIAS_VENCIDO > 0 AND DIAS_VENCIDO < 31 THEN 'VENCIDO 0 - 30 DIAS'
                WHEN DIAS_VENCIDO >= 31 AND DIAS_VENCIDO < 61 THEN 'VENCIDO 31 - 60 DIAS'
                WHEN DIAS_VENCIDO >= 61 AND DIAS_VENCIDO < 91 THEN 'VENCIDO 61 - 90 DIAS'
                WHEN DIAS_VENCIDO >= 91 AND DIAS_VENCIDO < 121 THEN 'VENCIDO 91 - 120 DIAS'
                WHEN DIAS_VENCIDO >= 121 AND DIAS_VENCIDO < 181 THEN 'VENCIDO 121 - 180 DIAS'
                WHEN DIAS_VENCIDO >= 181 AND DIAS_VENCIDO < 366 THEN 'VENCIDO 181 - 365 DIAS'
                WHEN DIAS_VENCIDO >= 365 THEN 'VENCIDO 365+ DIAS'
                ELSE 'A VENCER'
            END AS TEMPO_VENCIDO,
            ROUND(SUM(VALOR_ORIGINAL), 2) AS VALOR_ORIGINAL,
            DA.DEPARTAMENTO AS DEPARTAMENTO,
           -- IFNULL(DA.EMAIL, 'financeiro@trescoqueiros.com') AS EMAIL
           'Controladoria4@trescoqueiros.com,Controladoria2@trescoqueiros.com' AS EMAIL
        FROM RECEBER
        LEFT JOIN departamentos_adto DA ON DA.CODIGO = RECEBER.DEPARTAMENTO
        GROUP BY PESSOA_FIN, MOEDA_ORIGINAL
        ORDER BY PESSOA_FIN asc
        limit 1
"""

def execute_query():
    global consulta_query
    if consulta_query:
        print(f"Executando a consulta: {consulta_query}")
        # Aqui você pode executar a consulta no banco de dados
        # Exemplo simples:
        # cursor.execute(consulta_query)
    else:
        print("Nenhuma consulta SQL definida.")

# Teste a execução da consulta
execute_query()
