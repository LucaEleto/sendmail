import smtplib
import os
import xml.etree.ElementTree as ET
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from segredo import senha
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time
import logging

#Configuração log
logging.basicConfig(
    filename='log_sendmail', #nome do arquivo
    level=logging.INFO, #nivel do log
    format='%(asctime)s - %(levelname)s - %(message)s' #formato da mensagem
)

# Configurações do servidor SMTP
server_smtp = "smtp.gmail.com"
porta = 587
remetente = "lucavilarino321@gmail.com"
senha = senha

# Diretórios onde estão localizados os arquivos XML e os arquivos a serem enviados
pasta_xml = r"C:\Users\lucav\OneDrive\Documentos\GitHub\sendmail"
pasta_arquivos = r"C:\Users\lucav\OneDrive\Área de Trabalho\Notas"

def enviar_email(destinatario, assunto, corpo, arquivos_anexos):
    try:
        # Criação do email
        mensagem = MIMEMultipart()
        mensagem['From'] = remetente
        mensagem['To'] = destinatario
        mensagem['Subject'] = assunto
        mensagem.attach(MIMEText(corpo, 'plain'))

        # Anexando os arquivos
        for arquivo_anexo in arquivos_anexos:
            with open(arquivo_anexo, 'rb') as file:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(file.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(arquivo_anexo)}')
                mensagem.attach(part)

        # Conectando ao servidor SMTP e enviando o email
        server = smtplib.SMTP(server_smtp, porta)
        server.starttls()
        server.login(remetente, senha)
        server.sendmail(remetente, destinatario, mensagem.as_string())
        server.quit()
        logging.info(f'Email enviado com sucesso para o {destinatario}')
        print(f"Email enviado com sucesso para {destinatario}")
    except Exception as e:
        logging.info(f"Houve um erro ao enviar o email: {e}")
        print(f"Houve um erro ao enviar o email: {e}")

def obter_dados_do_xml(caminho_xml):
    try:
        logging.info(f'Lendo o arquivo XML: {caminho_xml}')
        tree = ET.parse(caminho_xml)
        root = tree.getroot()
        
        # Procura a chave de acesso e o e-mail do cliente no XML
        chave_acesso = None
        email_cliente = None

        # Encontra a chave de acesso dentro dos parâmetros
        for param in root.findall(".//msg-parameters/parameter"):
            if param.get('key') == '$cChave07':
                chave_acesso = param.get('value')
                break  # Encontra apenas uma vez
        
        # Encontra o e-mail do cliente dentro das addresses
        email_tag = root.find(".//addresses/address")
        if email_tag is not None:
            email_cliente = email_tag.get('value')
                
        return chave_acesso, email_cliente
    except Exception as e:
        logging.info(f'Erro ao ler o arquivo XML {caminho_xml}: {e}')
        print(f"Erro ao ler o arquivo XML: {e}")
        return None, None

def localizar_arquivos_com_chave(chave_acesso):
    arquivos_encontrados = []
    for ext in ['.pdf', '.xml']:
        arquivo = os.path.join(pasta_arquivos, f"{chave_acesso}-procNfe{ext}")
        if os.path.exists(arquivo):
            arquivos_encontrados.append(arquivo)
            logging.info(f'Arquivo encontrado: {arquivo}')
    return arquivos_encontrados

def processar_e_enviar(caminho_xml):
    chave_acesso, email_cliente = obter_dados_do_xml(caminho_xml)
    
    if chave_acesso and email_cliente:
        arquivos_anexos = localizar_arquivos_com_chave(chave_acesso)
        
        if arquivos_anexos:
            assunto = "Nota Fiscal Eletrônica"
            corpo = "Segue em anexo, XML e PDF da Nota Fiscal Eletrônica."
            enviar_email(email_cliente, assunto, corpo, arquivos_anexos)
            os.remove(caminho_xml)  # Exclui o arquivo XML após o envio
            logging.info(f"Arquivo XML {caminho_xml} enviado e excluído.")
            print(f"Arquivo XML {caminho_xml} enviado e excluído.")
        else:
            logging.info(f"Arquivos com a chave {chave_acesso} não encontrados na pasta de arquivos.")
            print(f"Arquivos com a chave {chave_acesso} não encontrados na pasta de arquivos.")
    else:
        logging.info(f"Erro ao extrair dados do XML {caminho_xml}.")
        print(f"Erro ao extrair dados do XML {caminho_xml}.")

# Classe de evento para monitorar novos arquivos XML na pasta
class MonitoramentoHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return
        if event.src_path.endswith(".xml"):
            logging.info(f"Novo arquivo detectado: {event.src_path}")
            print(f"Novo arquivo detectado: {event.src_path}")
            processar_e_enviar(event.src_path)

# Função principal para iniciar o monitoramento
def monitorar_pasta():
    observer = Observer()
    event_handler = MonitoramentoHandler()
    observer.schedule(event_handler, pasta_xml, recursive=False)
    observer.start()
    logging.info(f"Monitorando a pasta {pasta_xml} para novos arquivos XML...")
    print(f"Monitorando a pasta {pasta_xml} para novos arquivos XML...")

    try:
        while True:
            time.sleep(1)  # Mantém o monitoramento ativo
    except KeyboardInterrupt:
        observer.stop()
        logging.info('Monitoramento interrompido pelo usuário')
    observer.join()

# Inicia o monitoramento da pasta XML
monitorar_pasta()
