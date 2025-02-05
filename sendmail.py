import smtplib
import os
import xml.etree.ElementTree as ET
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import time
import mysql.connector
from mysql.connector import Error
import mysql.connector
from mysql.connector import Error

try:
    con = mysql.connector.connect(host='localhost', database='adalto', user='root', password='clara02')
    print('Conexão Realizada com Sucesso!!')
    consulta_sql = "SELECT emailRemet, senha, smtp, porta FROM sendmail"
    cursor = con.cursor()
    cursor.execute(consulta_sql)
    linhas = cursor.fetchall()
    print(f'Número total de registros retornados: {cursor.rowcount}')

    if linhas:
        remetente = linhas[0][0]
        senha = linhas[0][1]
        server_smtp = linhas[0][2]
        porta = linhas [0][3]

        senha_oculta = '*' * len(senha)

        print("\nMonstrando os dados do email")
        print(f'Email: {remetente}')
        print(f'Senha: {senha_oculta}')
        print(f'smtp: {server_smtp}')
        print(f'porta: {porta} \n')
    else:
        print("Nenhum Registro Encontrado")
except Error as e:
    print(f'Erro ao acessar tabela MySQL {e}')
finally:
    if con.is_connected():
        cursor.close()
        con.close()
        print('Conexão Encerrada \n')


# Diretórios onde estão localizados os arquivos XML e os arquivos a serem enviados
pasta_xml = r"C:\Check6\Novo Sendmail"  # Pasta onde novos XMLs chegam
pasta_arquivos = r"C:\Check6\NFe\PROTOCOLADAS"  # Pasta onde estão os arquivos das notas fiscais

def enviar_email(destinatario, assunto, corpo, arquivos_anexos):
    try:
        mensagem = MIMEMultipart()
        mensagem['From'] = remetente
        mensagem['To'] = destinatario
        mensagem['Subject'] = assunto
        mensagem.attach(MIMEText(corpo, 'plain'))

        for arquivo_anexo in arquivos_anexos:
            with open(arquivo_anexo, 'rb') as file:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(file.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(arquivo_anexo)}')
                mensagem.attach(part)

        server = smtplib.SMTP(server_smtp, porta)
        server.starttls()
        server.login(remetente, senha)
        server.sendmail(remetente, destinatario, mensagem.as_string())
        server.quit()

        print(f"Email enviado com sucesso para {destinatario}")
    except Exception as e:
        print(f"Houve um erro ao enviar o email: {e}")

def obter_dados_do_xml(caminho_xml):
    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()
        chave_acesso = None
        email_cliente = None

        for param in root.findall(".//msg-parameters/parameter"):
            if param.get('key') == '$cChave07':
                chave_acesso = param.get('value')
                break
        
        email_tag = root.find(".//addresses/address")
        if email_tag is not None:
            email_cliente = email_tag.get('value')

        return chave_acesso, email_cliente
    except Exception as e:
        print(f"Erro ao ler o arquivo XML: {e}")
        return None, None

def localizar_arquivos_com_chave(chave_acesso):
    arquivos_encontrados = []
    for ext in ['.pdf', '.xml']:
        arquivo = os.path.join(pasta_arquivos, f"{chave_acesso}-procNfe{ext}")
        if os.path.exists(arquivo):
            arquivos_encontrados.append(arquivo)
    return arquivos_encontrados

def processar_e_enviar(caminho_xml):
    chave_acesso, email_cliente = obter_dados_do_xml(caminho_xml)
    
    if chave_acesso and email_cliente:
        arquivos_anexos = localizar_arquivos_com_chave(chave_acesso)
        
        if arquivos_anexos:
            assunto = "Nota Fiscal Eletrônica"
            corpo = "Segue em anexo, XML e PDF da Nota Fiscal Eletrônica."
            enviar_email(email_cliente, assunto, corpo, arquivos_anexos)
            os.remove(caminho_xml)
            print(f"Arquivo XML {caminho_xml} enviado e excluído.")
        else:
            print(f"Arquivos com a chave {chave_acesso} não encontrados na pasta de arquivos.")
    else:
        print(f"Erro ao extrair dados do XML {caminho_xml}.")

# Função para buscar XMLs no diretório
def buscar_xml_no_diretorio():
    arquivos_xml = [os.path.join(pasta_xml, f) for f in os.listdir(pasta_xml) if f.endswith(".xml")]
    if arquivos_xml:
        arquivos_xml.sort(key=os.path.getmtime, reverse=True)  # Ordena pelo mais recente
        print(f"Arquivo XML mais recente encontrado: {arquivos_xml[0]}")
        return arquivos_xml[0]
    print("Nenhum arquivo XML encontrado no diretório.")
    return None

# Monitoramento e processamento principal
def monitorar_pasta():
    while True:
        caminho_xml = buscar_xml_no_diretorio()
        if caminho_xml:
            processar_e_enviar(caminho_xml)
        time.sleep(30)  # Verifica a cada 30 segundos

# Inicia o monitoramento da pasta XML
monitorar_pasta()
