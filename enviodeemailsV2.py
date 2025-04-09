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


def send_emails(config, data, file_path, hourly_limit, pause_minutes):
    smtp_host = config["smtp_host"]
    smtp_port = config["smtp_port"]
    smtp_user = config["smtp_user"]
    smtp_password = config["smtp_password"]

    pause_duration = pause_minutes * 60  # Converter minutos para segundos
    email_count = 0  # Contador de e-mails enviados

    try:
        # Função para inicializar o servidor SMTP
        def init_smtp_client():
            server = smtplib.SMTP(smtp_host, smtp_port)
            server.starttls()
            server.login(smtp_user, smtp_password)
            return server

        # Inicializar o cliente SMTP
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
                email_count += 1

            except Exception as e:
                print(f"Erro ao enviar e-mail para {email_to}: {e}")
                data.at[index, "Status"] = f"Erro: {e}"

            # Verificar se o limite de e-mails foi atingido
            if email_count == hourly_limit:
                print(f"Limite de {hourly_limit} e-mails atingido. Pausando por {pause_minutes} minutos...")
                server.quit()
                time.sleep(pause_duration)
                email_count = 0  # Reiniciar o contador
                server = init_smtp_client()  # Reiniciar a conexão SMTP

        # Fechar a conexão SMTP ao final
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
                try:
                    # Solicitar limites de envio e tempo de pausa
                    hourly_limit = int(input("Digite o limite de e-mails por período (ex.: 50): "))
                    pause_minutes = int(input("Digite o tempo de pausa em minutos (ex.: 60): "))
                    
                    proceed = input("Deseja prosseguir com os disparos? (1-Sim / 2-Não): ")
                    if proceed == "1":
                        send_emails(smtp_config, data, file_path, hourly_limit, pause_minutes)
                except ValueError:
                    print("Por favor, insira valores válidos para o limite e o tempo de pausa.")
        elif choice == "3":
            print("Encerrando o sistema.")
            break
        else:
            print("Opção inválida. Tente novamente.\n")


if __name__ == "__main__":
    main()
