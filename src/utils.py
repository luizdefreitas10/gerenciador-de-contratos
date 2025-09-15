import os
import smtplib
from dotenv import load_dotenv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
from email.message import EmailMessage
from datetime import datetime


ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ENV_PATH = os.path.join(ROOT_DIR, ".env")
load_dotenv(ENV_PATH)


SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")
DRY_RUN = os.getenv("DRY_RUN", "False").lower() == "true"


PLANILHA_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "Planilha de Contratos 2025 - CÓPIA.xlsx",
)


def enviar_email(destinatario: str, assunto: str, mensagem: str):
    if DRY_RUN:
        print("--- DRY RUN: não envia ---")
        print(f"Para: {destinatario} | Assunto: {assunto}")
        return True

    msg = EmailMessage()
    msg["From"] = EMAIL_USER
    msg["To"] = destinatario
    msg["Subject"] = assunto
    msg.set_content(mensagem)

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) as s:
            s.set_debuglevel(0)
            s.ehlo()
            s.starttls()
            s.ehlo()
            s.login(EMAIL_USER, EMAIL_PASS)
            s.send_message(msg)
        print(f"✅ [OK] Email enviado para {destinatario}")
        return True
    except smtplib.SMTPAuthenticationError as e:
        print("[ERRO] Autenticação SMTP falhou:", e)
        with open("email_error.log", "a", encoding="utf-8") as f:
            f.write(
                f"{datetime.now().isoformat()} AUTH_FAIL for {EMAIL_USER} -> {destinatario}\n{e}\n"
            )
        return False
    except Exception as e:
        print("[ERRO] Falha ao enviar email:", type(e), e)
        with open("email_error.log", "a", encoding="utf-8") as f:
            f.write(f"{datetime.now().isoformat()} SEND_FAIL -> {destinatario}\n{e}\n")
        return False


def get_email_gestor(nome_gestor: str) -> str:
    """Busca o email do gestor pelo nome na aba 'Gestores'"""
    df_gestores = pd.read_excel(PLANILHA_PATH, sheet_name="Gestores", header=3)
    linha = df_gestores[
        df_gestores["Gestor do Contrato"].str.lower() == str(nome_gestor).lower()
    ]
    if not linha.empty:
        return linha["Email "].values[0]
    return None
