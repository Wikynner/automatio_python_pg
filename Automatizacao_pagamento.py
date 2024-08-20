import pyodbc
from email.mime.image import MIMEImage
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import re
import sys
import os

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), 'Variaveis_Global')))
#caminho_logo = "C:\Users\wikynner.pires\Documents\java\IDENTLAN_Duplicados\logo_3coqueiros.png"
try:
   import db_connection_sql as db_module
except ImportError as e :
    print(f"Erro ao importar os módulos:{e}")
    sys.exit(1)
# # Configurações de e-mail
# DB_SERVER = '192.168.100.105'
# DB_DATABASE = 'agrimanager'
# DB_USERNAME = 'controladoria'
# DB_PASSWORD = 'Senha@2022'
EMAIL_ADDRESS = 'noreply@trescoqueiros.com'
EMAIL_PASSWORD = 'Lak41871'
EMAIL_CC = 'aprendizfinanceiro@trescoqueiros.com,financeiro5@trescoqueiros.com,financeiro3@trescoqueiros.com,financeiro4@trescoqueiros.com,compras@trescoqueiros.com'
EMAIL_BCC = 'controladoria2@trescoqueiros.com,controladoria3@trescoqueiros.com,controladoria4@trescoqueiros.com,controladoria@trescoqueiros.com'
DEFAULT_EMAIL ='compras@trescoqueiros.com'

# Função para validar o formato do e-mail
def is_valid_email(email):
    regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(regex, email) is not None

# Função para enviar e-mail com HTML
def send_email(to_addresses, subject, html_body, email_type='correto', cc_addresses=None, bcc_addresses=None, error_emails=None):
    from_address = EMAIL_ADDRESS
    password = EMAIL_PASSWORD

    msg = MIMEMultipart()
    msg['From'] = from_address
    msg['To'] = ', '.join(to_addresses)
    msg['Subject'] = subject

    # Adicionar os endereços CC ao cabeçalho do e-mail, se fornecidos
    if cc_addresses:
        msg['Cc'] = ', '.join(cc_addresses)

    # Adicionar os endereços BCC ao cabeçalho do e-mail, se fornecidos
    if bcc_addresses:
        msg['Bcc'] = ', '.join(bcc_addresses)
        
  # Adicionando a imagem da logo
    caminho_logo = r"C:\Controladoria\Automação\Automatio_pg_python\logo_3coqueiros.png" #  caminho da log
    with open(caminho_logo, 'rb') as img:
        logo = MIMEImage(img.read())
        logo.add_header('Content-ID', '<logo>')  # ID para referenciar a imagem no HTML
        logo.add_header('Content-Disposition', 'inline', filename=os.path.basename(caminho_logo))
        msg.attach(logo)


    # Construir o corpo do e-mail baseado no tipo (correto, erro)
    if email_type == 'correto':
        msg.attach(MIMEText(html_body, 'html'))
    elif email_type == 'erro':
        # Corpo do e-mail para avisos de erro
        error_body = f"""
        <html>
        <body>
            <p>Alguns endereços de e-mail fornecidos são inválidos:</p>
            <table>
                <thead>
                    <tr>
                        <th>E-mail Inválido</th>
                    </tr>
                </thead>
                <tbody>
                    {error_emails}
                </tbody>
            </table>
            <p>Por favor, verifique e atualize os e-mails corretamente para garantir o recebimento de futuras comunicações.</p>
        </body>
        </html>
        """
        msg.attach(MIMEText(error_body, 'html'))

    # Configurar o servidor SMTP
    server = smtplib.SMTP('smtp.office365.com', 587)
    server.starttls()
    server.login(from_address, password)

    # Enviar o e-mail, incluindo os endereços CC e BCC na lista de destinatários
    recipients = to_addresses + (cc_addresses or []) + (bcc_addresses or [])
    server.sendmail(from_address, recipients, msg.as_string())
    server.quit()

try:
    # Query SQL para obter os dados dos pagamentos processados e os e-mails dos fornecedores
    query = '''
   SELECT 
        TL.numeroboleto                            NUMERO_TITULO,
        ps.codigo                                  ID_PESSOA,
        PS.NOME                                    FORNECEDOR,
        RTRIM(T2.DESCRICAO)                        FORMA_PAGAMENTO,
        TD.DESCRICAO                               TIPO_DOC,
        TL.emissao								   DT_EMISSAO,
        TL.primvcto    							   PRIM_DT_VCTO,
        CONVERT(VARCHAR(10), TL.ultpgto, 103)      DT_VCTO,
        CASE
            WHEN TL.saldo = 0 THEN CONVERT(VARCHAR(10), TL.ultpgto, 103)
        END                                        DT_PAGTO,
        TL.valor                                   VALOR_TITULO,
        TL.saldo                                   SALDO_PEDENTE,
        T1.MASCARA                                 MOEDA,
        TL.valor2                                  VALOR_ORIGINAL,
        TL.saldo2                                  SALDO_ORIGINAL,
        case 
        	when ps.Email is null     then 'SEM EMAIL'
        	WHEN RTRIM(ps.Email) = '' THEN 'SEM EMAIL'
        	else RTRIM(ps.Email) 
         end                                        EMAIL
        -- 'controladoria4@trescoqueiros.com'            EMAIL  
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
           AND (CASE WHEN TL.saldo = 0 THEN TL.ultpgto END ) =  CAST( DATEADD( DAY , -1,SYSDATETIME()) AS DATE )
           --AND (CASE WHEN TL.saldo = 0 THEN CAST( TL.ultpgto AS DATE ) END ) = '2024-06-25'
           AND TL.formapagto NOT IN (8,11,12)    
           AND TD.tipo = 'S'
           AND DC.TIPODOC NOT IN (  
                                    67,                         -- ARREND.
                                    119,218,75,147,214,112,99,74,116,
                                                                -- GUIA IMPOSTO  
                                    240,                        -- NF3 ENERGIA
                                    62, 63, 64, 65,             -- CONSORCIO
                                    71, 213, 85, 233, 234, 203, -- FOLHA
                                    14, 15, 23                  -- FINANCIAMENTO
                                    )
        ORDER BY PS.nome ASC 
    '''

    # Conectar ao banco de dados
    #conn = pyodbc.connect('DRIVER={SQL Server};SERVER='+DB_SERVER+';DATABASE='+DB_DATABASE+';UID='+DB_USERNAME+';PWD='+DB_PASSWORD)
    conn = db_module.get_connection()
    cursor = conn.cursor()

    # Executar a consulta SQL
    cursor.execute(query)
    rows = cursor.fetchall()

    # Dicionário para armazenar informações de e-mail por ID_PESSOA
    fornecedores_notas = {}
    emails_invalidos = set()  # Usar um conjunto para rastrear e-mails inválidos já processados

    # Preencher o dicionário com os dados obtidos
    for row in rows:
        id_pessoa = row.ID_PESSOA
        fornecedor_emails = row.EMAIL.strip().split(',') if hasattr(row, 'EMAIL') and row.EMAIL else None

        if fornecedor_emails is None or all(not is_valid_email(email.strip()) for email in fornecedor_emails):
            if id_pessoa not in emails_invalidos:
                # Se todos os e-mails não forem válidos, enviar e-mail de aviso (somente uma vez por fornecedor)
                subject = f"Problema com endereço de e-mail para {row.FORNECEDOR}"
                html_body = f"""
                <html>
                <body>
                    <p>Os endereços de e-mail fornecidos para {row.FORNECEDOR} são inválidos:</p>
                    <ul>
                        {''.join(f'<li>{email}</li>' for email in fornecedor_emails)}
                    </ul>
                    <p>Por favor, verifique e atualize os e-mails corretamente para garantir o recebimento de futuras comunicações.</p>
                </body>
                </html>
                """
                send_email(DEFAULT_EMAIL.split(','), subject, html_body, email_type='erro', cc_addresses=EMAIL_CC.split(','), bcc_addresses=EMAIL_BCC.split(','))
                print(f"E-mail de aviso enviado para {', '.join(DEFAULT_EMAIL.split(','))} sobre o problema com o e-mail de {row.FORNECEDOR}.")
                emails_invalidos.add(id_pessoa)  # Adicionar fornecedor à lista de e-mails inválidos processados
        else:
            valid_emails = [email.strip() for email in fornecedor_emails if is_valid_email(email.strip())]
            if id_pessoa not in fornecedores_notas:
                fornecedores_notas[id_pessoa] = {
                    'emails': valid_emails,
                    'notas': []
                }
            
            fornecedores_notas[id_pessoa]['notas'].append(row)

    # Verificar se há resultados para enviar e-mails com atualizações de pagamento
    if len(fornecedores_notas) > 0:
        for id_pessoa, info in fornecedores_notas.items():
            emails_fornecedor = info['emails']
            notas = info['notas']

            # Construir a tabela HTML com os resultados da consulta para este fornecedor
            tabela_html = """
            <table>
                <thead>
                    <tr>
                        <th>Número do Título</th>
                        <th>Fornecedor</th>
                        <th>Data de Vencimento</th>
                        <th>Data de Pagamento</th>
                        <th>Valor</th>
                        <th>Forma de Pagamento</th>
                    </tr>
                </thead>
                <tbody>
            """
            
            
            for row in notas:
                VALOR_TITULO =row.VALOR_TITULO
                VALOR_TITULO = '{:,.2f}'.format(VALOR_TITULO)
                VALOR_TITULO = str(VALOR_TITULO)
                VALOR_TITULO = VALOR_TITULO.replace(',','!')
                VALOR_TITULO = VALOR_TITULO.replace('.',',')
                VALOR_TITULO = 'R$ ' + VALOR_TITULO.replace('!','.')
                
                tabela_html += f"""
                    <tr>
                        <td>{row.NUMERO_TITULO}</td>
                        <td>{row.FORNECEDOR}</td>
                        <td>{row.DT_VCTO}</td>
                        <td>{row.DT_PAGTO if row.DT_PAGTO is not None else 'Não Pago'}</td>
                        <td>{VALOR_TITULO }</td>
                        <td>{row.FORMA_PAGAMENTO}</td>
                    </tr>
                """

            tabela_html += """
                </tbody>
            </table>
            """
            # Corpo do e-mail com a tabela de resultados para este fornecedor
            mensagem_html = f"""
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
            margin-top: 5%;
            border: 1px solid #ddd;
            border-radius: 8px;
            overflow: hidden;
        }}
        .will {{
            margin-top: 5%;
        }}
        .header {{
            background-color: rgb(1, 85, 26);
            padding: 20px;
            align-items: center;
            color: white;
            display: flex;
            justify-content: flex-start; 
        }}
        .logo3c {{
            flex: 0 0 18%;
            text-align: left;
            width: 90px;
            height: auto;
        }}
        .titulo {{
            flex: 1; /* Equivale a 80% */
            text-align: left; /* Alinha o texto à esquerda */
            margin-left: 10px;
        }}
        .titulo h1 {{
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
                <div class="logo3c">
                    <img src="cid:logo" class="logo3c" alt="Logo Três Coqueiros"/>
                </div>
                <div class="titulo">
                    <h1>Pagamentos Processados</h1>
                </div>
            </div>
            <div class="content">
                <p>Olá {notas[0].FORNECEDOR},</p>
                <p>Gostaríamos de informá-lo que os seguintes pagamentos de título foram processados recentemente. Por favor, verifique os detalhes abaixo:</p>
                <div class="table-container">
                    {tabela_html}
                </div>
            </div>
        </div>
        <div class="footer">
            <p>Este é um e-mail automático, por favor não responda.</p>
        </div>
    </div>
</body>
</html>

            """

            # Preparar a lista de endereços CC a partir da variável de ambiente
            cc_addresses = EMAIL_CC.split(',') if EMAIL_CC else None
            bcc_addresses = EMAIL_BCC.split(',') if EMAIL_BCC else None

            # Definir um assunto específico para o e-mail
            subject = f"PAGAMENTOS PROCESSADOS PARA {notas[0].FORNECEDOR} EM {notas[0].DT_PAGTO or 'recente'}"

            # Enviar o e-mail para o fornecedor atual
            send_email(emails_fornecedor, subject, mensagem_html, email_type='correto', cc_addresses=cc_addresses, bcc_addresses=bcc_addresses)
            print(f"Email enviado com os pagamentos processados para {', '.join(emails_fornecedor)}.")

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
