# -*- coding: utf-8 -*-

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
import os
import json
import time

CONFIG_FILE = "config.json"

def save_smtp_config():
    print("Configuração do Servidor SMTP:")
    smtp_host = input("Servidor SMTP: ")
    smtp_port = input("Porta SMTP: ")
    smtp_user = input("Usuário: ")
    smtp_password = input("Senha: ")

    config = {
        "smtp_host": smtp_host,
        "smtp_port": smtp_port,
        "smtp_user": smtp_user,
        "smtp_password": smtp_password
    }

    with open(CONFIG_FILE, "w") as file:
        json.dump(config, file)

    print("Configurações salvas no arquivo 'config.json'.\n")


def load_smtp_config():
    if not os.path.exists(CONFIG_FILE):
        print("Arquivo de configuração não encontrado. Configure o servidor SMTP primeiro.")
        save_smtp_config()

    with open(CONFIG_FILE, "r") as file:
        return json.load(file)


def load_spreadsheet():
    file_path = input("Digite o caminho para a planilha (Excel): ")
    try:
        data = pd.read_excel(file_path)
        if not {"A", "B", "C"}.issubset(data.columns):
            raise ValueError("A planilha deve conter as colunas A, B e C.")
        return data, file_path
    except Exception as e:
        print(f"Erro ao carregar a planilha: {e}")
        return None, None


def send_emails(config, data, file_path):
    smtp_host = config["smtp_host"]
    smtp_port = config["smtp_port"]
    smtp_user = config["smtp_user"]
    smtp_password = config["smtp_password"]

    # Configurar o número de e-mails por lote e tempo de pausa
    batch_size = 5
    pause_time = 60  # Tempo de pausa entre os lotes em segundos

    try:
        # Função para inicializar o servidor SMTP
        def init_smtp_client():
            server = smtplib.SMTP(smtp_host, smtp_port)
            server.login(smtp_user, smtp_password)
            return server

        # Enviar e-mails em lotes
        email_count = 0
        server = init_smtp_client()

        for index, row in data.iterrows():
            email_to = row["A"]
            subject = row["B"]
            body = row["C"]

            try:
                msg = MIMEMultipart()
                msg["From"] = smtp_user
                msg["To"] = email_to
                msg["Subject"] = subject

                msg.attach(MIMEText(body, "html"))

                server.sendmail(smtp_user, email_to, msg.as_string())
                print(f"E-mail enviado com sucesso para: {email_to}")
                data.at[index, "Status"] = "Sucesso"
            except Exception as e:
                print(f"Erro ao enviar e-mail para {email_to}: {e}")
                data.at[index, "Status"] = f"Erro: {e}"

            # Após cada 5 e-mails, reiniciar a conexão SMTP e dar pausa
            email_count += 1
            if email_count % batch_size == 0:
                print(f"Pausa de {pause_time} segundos após enviar {batch_size} e-mails.")
                time.sleep(pause_time)

                # Fechar e reiniciar o servidor SMTP
                server.quit()
                print(f"Reiniciando a conexão SMTP após enviar {email_count} e-mails.")
                server = init_smtp_client()

        # Fechar a conexão SMTP
        server.quit()

        # Atualizar a planilha com os resultados
        data.to_excel(file_path, index=False)
        print("Processo concluído. Resultados salvos na planilha.")
    except Exception as e:
        print(f"Erro na configuração do servidor SMTP: {e}")


def main():
    print("Sistema de Disparo de E-mails\n")
    smtp_config = load_smtp_config()

    while True:
        print("1. Configurar Servidor SMTP")
        print("2. Carregar Planilha e Enviar E-mails")
        print("3. Sair")
        choice = input("Escolha uma opção: ")

        if choice == "1":
            save_smtp_config()
        elif choice == "2":
            data, file_path = load_spreadsheet()
            if data is not None:
                proceed = input("Deseja prosseguir com os disparos? (1-Sim / 2-Não): ")
                if proceed == "1":
                    send_emails(smtp_config, data, file_path)
        elif choice == "3":
            print("Encerrando o sistema.")
            break
        else:
            print("Opção inválida. Tente novamente.\n")


if __name__ == "__main__":
    main()
