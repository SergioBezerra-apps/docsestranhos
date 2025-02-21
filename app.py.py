#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import streamlit as st
import pandas as pd
import smtplib
import os
from email.message import EmailMessage

# =============================================================================
# CONFIGURAÇÕES ( ajuste seu SMTP, usuário e senha do app )
# =============================================================================
smtp_server = 'smtp.gmail.com'
smtp_port = 465
smtp_username = 'sergiolbezerralj@gmail.com'
smtp_password = 'dimwpnhowxxeqbes'  # verifique se está correto

# Lista de naturezas "habituais"
NATUREZAS_HABITUAIS = [
    "APOSENTADORIA",
    "CONCURSO PÚBLICO",
    "CONCURSO PÚBLICO (DOC)",
    "CONCURSO PÚBLICO (RETIFICAÇÃO)",
    "CONTRATAÇÃO DE PESSOAL POR PRAZO DETERMINADO",
    "CONTRATAÇÃO DE PESSOAL POR PRAZO DETERMINADO (RETIFICAÇÃO)",
    "PEDIDO",
    "PENSÃO",
    "PROMOÇÃO",
    "REFORMA",
    "RESPOSTA A OFÍCIO",
    "REVISÃO DE PENSÃO",
    "REVISÃO DE PROVENTOS",
    "TRANSFERÊNCIA PARA RESERVA REMUNERADA",
]

# =============================================================================
# FUNÇÃO PARA ENVIO DE EMAIL COM ANEXOS
# =============================================================================
def send_email_with_attachments(to_emails, subject, body, attachment_paths):
    """
    Envia e-mail com anexos usando SSL (biblioteca padrão).
    """
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = SMTP_USERNAME
    msg['To'] = ', '.join(to_emails)
    msg.set_content(body)

    # Adiciona os anexos
    for path in attachment_paths:
        with open(path, 'rb') as file:
            filename = os.path.basename(path)
            msg.add_attachment(file.read(),
                               maintype='application',
                               subtype='octet-stream',
                               filename=filename)

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, timeout=10) as server:
            server.login(SMTP_USERNAME, SMTP_PASSWORD)
            server.send_message(msg)
            print("E-mail enviado com sucesso!")
    except smtplib.SMTPConnectError as e:
        print(f"Erro de conexão SMTP: {e}")
    except smtplib.SMTPAuthenticationError as e:
        print(f"Erro de autenticação SMTP: {e}")
    except smtplib.SMTPException as e:
        print(f"Erro SMTP: {e}")
    except Exception as e:
        print(f"Erro geral: {e}")

# =============================================================================
# APLICAÇÃO STREAMLIT
# =============================================================================
def main():
    st.title("Filtragem de Documentos e Envio por E-mail")
    st.write("""
    **Como funciona**:
    1. Selecione um arquivo **.xlsx** (com colunas: nrdoc, dvdoc, andoc, nrprinc, dctramita, dcgrnatureza, etc.).
    2. Informe os e-mails de destino, separados por vírgula.
    3. Clique em **Processar e Enviar**.
    4. A aplicação gera dois arquivos:
       - **docs_principais.xlsx**: convergindo com a junção (nrdoc-dvdoc/últimos2dígitosAno) e filtrando (nrdoc <= 99999, dctramita='PRINCIPAL').
       - **natureza_nao_habitual.xlsx**: processos cuja 'dcgrnatureza' não faz parte da lista de naturezas habituais.
    5. Esses dois arquivos são enviados via e-mail.
    """)

    # 1) Upload do arquivo Excel
    uploaded_file = st.file_uploader("Selecione o arquivo Excel", type=["xlsx"])

    # 2) Campo para digitar e-mails
    emails_input = st.text_input(
        "Digite os e-mails de destino, separados por vírgula",
        value=""
    )

    # Botão de processamento
    if st.button("Processar e Enviar E-mails"):
        if not uploaded_file:
            st.error("Por favor, faça o upload de um arquivo Excel (.xlsx).")
            return

        if not emails_input.strip():
            st.error("Por favor, insira ao menos um endereço de e-mail.")
            return

        # Converte string de e-mails em lista
        to_emails = [email.strip() for email in emails_input.split(",")]

        # 3) Ler o DataFrame a partir do arquivo enviado
        try:
            df = pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"Erro ao ler o arquivo: {e}")
            return

        # ---------------------------------------------------------------------
        # 4) Criar a coluna "nrprinc_formatado": nrdoc-dvdoc/yy
        # ---------------------------------------------------------------------
        def format_nrprinc(row):
            ano_str = str(row["andoc"])
            ano2d = ano_str[-2:]  # pega últimos 2 dígitos
            return f"{row['nrdoc']}-{row['dvdoc']}/{ano2d}"

        df["nrprinc_formatado"] = df.apply(format_nrprinc, axis=1)

        # ---------------------------------------------------------------------
        # 5) Filtrar "Docs Principais"
        #     - nrdoc <= 99999
        #     - dctramita == 'PRINCIPAL'
        #     - nrprinc_formatado == nrprinc
        # ---------------------------------------------------------------------
        mask_principal = (
            (df["nrdoc"] <= 99999) &
            (df["dctramita"] == "PRINCIPAL") &
            (df["nrprinc_formatado"] == df["nrprinc"])
        )
        df_principais = df.loc[mask_principal].copy()

        # ---------------------------------------------------------------------
        # 6) Filtrar naturezas NÃO habituais
        # ---------------------------------------------------------------------
        mask_nao_habitual = ~df["dcgrnatureza"].isin(NATUREZAS_HABITUAIS)
        df_nao_habitual = df.loc[mask_nao_habitual].copy()

        # ---------------------------------------------------------------------
        # 7) Gerar os dois arquivos .xlsx
        #    Obs.: usamos nomes fixos, mas poderíamos personalizar com data/hora
        # ---------------------------------------------------------------------
        file_principais = "docs_principais.xlsx"
        file_nao_habitual = "natureza_nao_habitual.xlsx"

        if not df_principais.empty:
            df_principais.to_excel(file_principais, index=False)
        else:
            pd.DataFrame(columns=df.columns).to_excel(file_principais, index=False)

        if not df_nao_habitual.empty:
            df_nao_habitual.to_excel(file_nao_habitual, index=False)
        else:
            pd.DataFrame(columns=df.columns).to_excel(file_nao_habitual, index=False)

        # ---------------------------------------------------------------------
        # 8) Envio de e-mail
        # ---------------------------------------------------------------------
        subject = "Resultados - DOCS PRINCIPAIS e NATUREZA NÃO HABITUAL"
        body = (
            "Segue em anexo:\n\n"
            "1) docs_principais.xlsx => Registros que batem com nrdoc-dvdoc/yy, "
            "tendo (nrdoc <= 99999) e dctramita='PRINCIPAL'.\n"
            "2) natureza_nao_habitual.xlsx => Contém processos cuja 'dcgrnatureza' "
            "não faz parte da lista de naturezas habituais.\n\n"
            "Atenciosamente,\nSistema Streamlit"
        )

        attachment_paths = [file_principais, file_nao_habitual]
        send_email_with_attachments(
            to_emails=to_emails,
            subject=subject,
            body=body,
            attachment_paths=attachment_paths
        )

        st.success("Processamento concluído e e-mails enviados com sucesso!")


if __name__ == "__main__":
    main()
