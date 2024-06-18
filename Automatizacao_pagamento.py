
import pyodbc
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
import os


# Carregar variáveis de ambiente do arquivo .env
load_dotenv()

# Configurações de e-mail a partir do .env
EMAIL_ADDRESS = os.getenv('EMAIL_ADDRESS')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
EMAIL_RECIPIENT = os.getenv('EMAIL_RECIPIENT')
EMAIL_CC = os.getenv('EMAIL_CC')
DEFAULT_EMAIL = os.getenv('DEFAULT_EMAIL', 'controladoria2@trescoqueiros.com')  # Definir o e-mail padrão aqui

# Função para enviar e-mail com HTML
def send_email(to_address, subject, html_body, cc_addresses=None):
    from_address = EMAIL_ADDRESS
    password = EMAIL_PASSWORD

    msg = MIMEMultipart()
    msg['From'] = from_address
    msg['To'] = to_address
    msg['Subject'] = subject

    # Adicionar os endereços CC ao cabeçalho do e-mail, se fornecidos
    if cc_addresses:
        msg['Cc'] = ', '.join(cc_addresses)

    # Anexar a parte HTML ao e-mail
    msg.attach(MIMEText(html_body, 'html'))

    # Configurar o servidor SMTP
    server = smtplib.SMTP('smtp.office365.com', 587)
    server.starttls()
    server.login(from_address, password)

    # Enviar o e-mail, incluindo os endereços CC na lista de destinatários
    recipients = [to_address] + (cc_addresses or [])
    server.sendmail(from_address, recipients, msg.as_string())
    server.quit()

# Configurações de conexão com o banco de dados a partir do .env
DB_SERVER = os.getenv('DB_SERVER')
DB_DATABASE = os.getenv('DB_DATABASE')
DB_USERNAME = os.getenv('DB_USERNAME')
DB_PASSWORD = os.getenv('DB_PASSWORD')

# Query SQL para obter os dados dos pagamentos processados e os e-mails dos fornecedores
query = '''
SELECT TL.numeroboleto                            NUMERO_TITULO,
       ps.codigo                                  ID_PESSOA,
       PS.NOME                                    FORNECEDOR,
       RTRIM(T2.DESCRICAO)                        FORMA_PAGAMENTO,
       TD.DESCRICAO                               TIPO_DOC,
       TL.emissao                                  DT_EMISSAO,
       TL.primvcto                                PRIM_DT_VCTO,
       TL.datavcto                                DT_VCTO,
       CASE
         WHEN TL.saldo = 0 THEN TL.ultpgto
       END                                        DT_PAGTO,
       Isnull(TL.datafaixa1, TL.datafaixa2)       DT_LIBERACAO,
       TL.valor                                   VALOR_TITULO,
       TL.saldo                                   SALDO_PEDENTE,
       T1.MASCARA                                 MOEDA,
       TL.valor2                                  VALOR_ORIGINAL,
       TL.saldo2                                  SALDO_ORIGINAL,
       --NULLIF ( Ps.email,'')                      EMAIL
       CASE 
            WHEN ps.CODIGO = 6281145 THEN 'controladoria4@trescoqueiros.com'
            WHEN PS.CODIGO = 6808  THEN 'controladoria4@trescoqueiros.com'
            WHEN PS.CODIGO = 10229 THEN 'controladoria4@trescoqueiros.com'
       END                                       EMAIL
       
FROM   titulos TL

       INNER JOIN documentos DC
               ON TL.seqdoc = DC.sequencial
       INNER JOIN tiposdoc TD
               ON TD.codigo = DC.tipodoc
       INNER JOIN PESSOAS PS
                ON PS.CODIGO = DC.representante
       INNER JOIN TABELAS T1
                ON T1.CODIGO = TL.moeda2
                   AND T1.TIPO = 10
       INNER JOIN TABELAS T2
                ON T2.Codigo = TL.FORMAPAGTO
                   AND T2.Tipo = 218
WHERE  1 = 1
       AND ISNULL(DC.SituacaoEF, 'N') <> 'S'
       --AND (CASE WHEN TL.saldo = 0 THEN TL.ultpgto END ) =  CAST( DATEADD( DAY , -1,SYSDATETIME()) AS DATE ) -- 
       AND (CASE WHEN TL.saldo = 0 THEN TL.ultpgto END ) = '2024-06-14 00:00:00.000'
       AND TL.formapagto NOT IN (12)    
       AND TD.tipo = 'S'
       AND DC.TIPODOC NOT IN (  
                                67, 75,
                                71,  213, 85, 233, 234, 203  --FOLHA
                             )
       AND TL.FORMAPAGTO NOT IN (8)
       AND ps.codigo  IN (6808,10229,3670)
       
ORDER  BY DC.emissao DESC
'''
try:
    # Conectar ao banco de dados
    conn = pyodbc.connect('DRIVER={SQL Server};SERVER='+DB_SERVER+';DATABASE='+DB_DATABASE+';UID='+DB_USERNAME+';PWD='+DB_PASSWORD)
    cursor = conn.cursor()

    # Executar a consulta SQL
    cursor.execute(query)
    rows = cursor.fetchall()

    # Dicionário para armazenar informações de e-mail por fornecedor
    fornecedores_email = {}

    # Preencher o dicionário com os dados obtidos
    for row in rows:
        # Verificar se o e-mail está diretamente especificado na consulta SQL
        fornecedor_email = row.EMAIL.strip() if hasattr(row, 'EMAIL') and row.EMAIL else None
        
        # Se o e-mail não estiver diretamente especificado, usar o e-mail padrão definido no .env
        if not fornecedor_email:
            fornecedor_email = DEFAULT_EMAIL

        if fornecedor_email:
            if fornecedor_email not in fornecedores_email:
                fornecedores_email[fornecedor_email] = []
            fornecedores_email[fornecedor_email].append(row)

    # Verificar se há resultados
    if len(fornecedores_email) > 0:
        for email_fornecedor, info_pagamentos in fornecedores_email.items():
            # Construir a tabela HTML com os resultados da consulta para este fornecedor
            tabela_html = """
            <table>
                <thead>
                    <tr>
                        <th>Número do Título</th>
                        <th>Fornecedor</th>
                        <th>Tipo de Documento</th>
                        <th>Data de Emissão</th>
                        <th>Data de Vencimento</th>
                        <th>Data de Pagamento</th>
                        <th>Data de Liberação</th>
                        <th>Valor</th>
                        <th>Saldo Pendente</th>
                        <th>Moeda</th>
                        <th>Valor Original</th>
                        <th>Saldo Original</th>
                    </tr>
                </thead>
                <tbody>
            """

            for row in info_pagamentos:
                tabela_html += f"""
                    <tr>
                        <td>{row.NUMERO_TITULO}</td>
                        <td>{row.FORNECEDOR}</td>
                        <td>{row.TIPO_DOC}</td>
                        <td>{row.DT_EMISSAO}</td>
                        <td>{row.DT_VCTO}</td>
                        <td>{row.DT_PAGTO if row.DT_PAGTO is not None else 'Não Pago'}</td>
                        <td>{row.DT_LIBERACAO}</td>
                        <td>{row.VALOR_TITULO}</td>
                        <td>{row.SALDO_PEDENTE}</td>
                        <td>{row.MOEDA}</td>
                        <td>{row.VALOR_ORIGINAL}</td>
                        <td>{row.SALDO_ORIGINAL}</td>
                    </tr>
                """

            tabela_html += """
                </tbody>
            </table>
            """

            # Corpo do e-mail com a tabela de resultados para este fornecedor
            mensagem_html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        body {{
            font-family: Arial, sans-serif;
            background-color: #f3f3f3;
            margin: 20px 0;
            padding: 0;
        }}
        .container {{
            background-color: white;
            max-width: 90%;
            margin: 20px auto;
            margin-top: 5% ;
            border: 1px solid #ddd;
            border-radius: 8px;
            overflow: hidden;
        }}
        .will {{
        margin-top: 5% ;
        }}
        .header {{
            background-color: rgb(1,85,26);
            padding: 20px;
            text-align: center;
            color: white;
        }}
        .header h1 {{
            margin: 0;
            font-size: 24px;
        }}
        .content {{
            padding: 20px;
            text-align: left;
        }}
        .content p {{
            font-size: 16px;
            line-height: 1.5;
        }}
        .table-container {{
            margin: 20px 0;
            overflow-x: auto;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }}
        table, th, td {{
            border: 1px solid #ddd;
        }}
        th, td {{
            padding: 12px;
            text-align: left;
        }}
        th {{
            background-color: #f4f4f4;
        }}
        .footer {{
            text-align: center;
            padding: 20px;
            background-color: #f9f9f9;
            border-top: 1px solid #ddd;
        }}
    </style>
</head>
<body>
    <div class="container">
    <div class="will">
        <div class="header">
            <h1>ATUALIZAÇÃO DE PAGAMENTOS</h1>
        </div>
        <div class="content">
            <p>Gostaríamos de informá-lo que os seguintes pagamentos de título foram processados recentemente. Por favor, verifique os detalhes abaixo:</p>
            <div class="table-container">
                {tabela_html}
            </div>
        </div>
        </div>
        <div class="footer"></div>
    </div>
</body>
</html>"""

            # Preparar a lista de endereços CC a partir da variável de ambiente
            cc_addresses = EMAIL_CC.split(',') if EMAIL_CC else None

            # Enviar o e-mail para o fornecedor atual
            send_email(email_fornecedor, "Pagamentos Processados", mensagem_html, cc_addresses=cc_addresses)
            print(f"Email enviado com os pagamentos processados para {email_fornecedor}.")

    else:
        print("Não há pagamentos processados.")

except Exception as e:
    print(f"Erro ao conectar ao banco de dados ou enviar e-mails: {str(e)}")

finally:
    try:
        cursor.close()
        conn.close()
    except Exception as e:
        print(f"Erro ao fechar conexão: {str(e)}")
