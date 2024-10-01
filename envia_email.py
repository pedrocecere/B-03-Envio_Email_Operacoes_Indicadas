import pyodbc
from dotenv import load_dotenv
import os
import pandas as pd
from datetime import datetime
import win32com.client as win32

# Carregar variáveis de ambiente do arquivo .env
load_dotenv()  # Este comando lê o arquivo .env automaticamente

# Definindo os parâmetros de conexão a partir das variáveis de ambiente
server = os.getenv('DB_SERVER')
database = os.getenv('DB_DATABASE')
username = os.getenv('DB_USERNAME')
password = os.getenv('DB_PASSWORD')

# String de conexão com os dados lidos do .env
connection_string = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'

# Inicializa o Outlook
outlook = win32.Dispatch('Outlook.Application')
remetente = 'pedro.cecere@ramaadvogados.com.br'
assunto = 'Relatório de casos cadastrados'

# Dicionário de e-mails dos gestores e as carteiras que eles gerenciam
email_gestores = {
    "Pedro Vieira Cecere": "pedro.cecere@ramaadvogados.com.br",
    "Rafael Rama e Silva": "pedro.cecere@ramaadvogados.com.br",
    "Daniella": "pedro.cecere@ramaadvogados.com.br",
    "Sirlei Rama": "pedro.cecere@ramaadvogados.com.br",
    "Ellen Stella Rama": "pedro.cecere@ramaadvogados.com.br",
    "Rodrigo Rama e Silva" : "pedro.cecere@ramaadvogados.com.br"
}

# Dicionário que mapeia carteiras para gestores
carteiras_por_gestor = {
    "Pedro Vieira Cecere": ["Massificado PF", "Massificado PJ", "Autos", "Alto ticket", "Núcleo Massificado"],
    "Rafael Rama e Silva": ["Credito Imobiliario"],
    "Daniella": ["Massificado PF", "Massificado PJ", "Autos", "Alto ticket", "Núcleo Massificado"],
    "Sirlei Rama": ["Massificado PF", "Massificado PJ", "Autos", "Alto ticket", "Núcleo Massificado"],
    "Helen Rama": ["Recuperação Judicial", "Judicial Especializado"],
    "Rodrigo Rama e Silva" : ["Massificado PF", "Massificado PJ", "Autos", "Alto ticket", "Núcleo Massificado"]
}

def leitura_banco_de_dados(connection_string):
    # Definir o DataFrame antes do try para evitar problemas em caso de erro
    df = pd.DataFrame()

    # Conectando ao banco de dados
    try:
        conn = pyodbc.connect(connection_string)
        print("Conexão bem-sucedida!")

        # Criando o cursor para executar a consulta
        cursor = conn.cursor()
        
        # Definindo a consulta SQL
        query = """
        SELECT
            c.F13577 AS criado_em,
            d.F00689 AS nome,
            a.F31768 AS operacao,
            f.F00091 AS devedor,
            f.F27086 AS documento,
            g.F26297 AS carteira,
            CASE
                WHEN a.F16778 = 1 THEN 'Liquidado'
                WHEN a.F16778 = 2 THEN 'Em aberto'
                WHEN a.F16778 = 3 THEN 'Com pendência'
                WHEN a.F16778 = 4 THEN 'Em negociação'
                WHEN a.F16778 = 5 THEN 'Negociado'
                WHEN a.F16778 = 6 THEN 'Liquidado via negociação'
                WHEN a.F16778 = 7 THEN 'Devolvido para o cliente'
                ELSE 'Sem status'
            END AS situacao,
            c.F13661 AS processo
        FROM ramaprod.dbo.T01167 AS a
        LEFT JOIN ramaprod.dbo.T00041 AS b ON a.F35050 = b.ID
        LEFT JOIN ramaprod.dbo.T01166 AS c ON a.F13700 = c.ID
        LEFT JOIN ramaprod.dbo.T00003 AS d ON c.F13576 = d.ID
        LEFT JOIN ramaprod.dbo.T01889 AS e ON c.F26866 = e.ID
        LEFT JOIN ramaprod.dbo.T00030 AS f ON e.F26827 = f.ID
        LEFT JOIN ramaprod.dbo.T01859 AS g ON c.F26458 = g.ID
        WHERE 
            a.F31768 IS NOT NULL
            AND MONTH(c.F13577) = MONTH(GETDATE()) AND YEAR(c.F13577) = YEAR(GETDATE()) AND DAY(c.F13577) = DAY(GETDATE())
        ORDER BY c.F13577 DESC;
        """ 
        
        # Executando a consulta
        cursor.execute(query)
        
        # Buscando os resultados
        rows = cursor.fetchall()

        # Criando um DataFrame a partir dos resultados
        df = pd.DataFrame.from_records(rows, columns=['criado_em', 'nome', 'operacao', 'devedor', 'documento', 'carteira', 'situacao', 'processo'])

    except pyodbc.Error as e:
        print("Erro na conexão:", e)

    finally:
        # Fechando a conexão
        if 'conn' in locals():
            conn.close()
            print("Conexão encerrada.")
    
    return df

def tratamento_df_por_gestor(df, carteiras_por_gestor):
    # Filtrar e agrupar o DataFrame por carteiras de cada gestor
    arquivos_gerados = {}

    for gestor, carteiras in carteiras_por_gestor.items():
        df_gestor = df[df['carteira'].isin(carteiras)]
        
        if not df_gestor.empty:
            # Nome do arquivo Excel para o gestor
            nome_arquivo = 'Operações cadastradas.xlsx'
            df_gestor.to_excel(nome_arquivo, index=False)
            arquivos_gerados[gestor] = nome_arquivo
            print(f'Arquivo {nome_arquivo} criado para o gestor {gestor}')

    return arquivos_gerados

def enviar_email(arquivos_gerados, email_gestores):
    # Enviar e-mails com os arquivos Excel gerados
    for gestor, arquivo in arquivos_gerados.items():
        destinatario = email_gestores.get(gestor)
        if destinatario:
            mail = outlook.CreateItem(0)
            mail.Subject = assunto
            mail.To = destinatario
            mail.Body = f"Olá {gestor},\n\nSegue em anexo o relatório das carteiras sob sua responsabilidade.\n\nAtenciosamente,\n\nPedro Cecere - Analista de Dados"
            mail.Attachments.Add(os.path.abspath(arquivo))
            mail.Send()
            print(f'E-mail enviado para {destinatario} com o arquivo {arquivo}')

# Executando as funções
df = leitura_banco_de_dados(connection_string)
arquivos_gerados = tratamento_df_por_gestor(df, carteiras_por_gestor)
enviar_email(arquivos_gerados, email_gestores)

