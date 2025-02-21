#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import streamlit as st
import pandas as pd
import smtplib
import os
from email.message import EmailMessage

# =============================================================================
# 1. CONFIGURAÇÕES E LISTA DE NATUREZAS HABITUAIS
# =============================================================================
smtp_server = 'smtp.gmail.com'
smtp_port = 465
smtp_username = 'sergiolbezerralj@gmail.com'
smtp_password = 'dimwpnhowxxeqbes'  # verifique se está correto


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
# 2. FUNÇÃO DE ENVIO DE E-MAIL COM ANEXOS (MANTIDA DA LÓGICA ANTERIOR)
# =============================================================================
def send_email_with_attachments(
    to_emails,
    subject,
    body,
    attachment_paths
):
    """
    Envia e-mail com anexos usando SSL.
    `to_emails` deve ser lista de e-mails.
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
            msg.add_attachment(
                file.read(),
                maintype='application',
                subtype='octet-stream',
                filename=filename
            )

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
# 3. APLICAÇÃO STREAMLIT
# =============================================================================
def main():
    st.title("Filtragem de Documentos e Envio por E-mail")
    st.write("""
    Esta aplicação:
    1. Lê o arquivo **cap3.xlsx**.
    2. Compara a concatenação (nrdoc-dvdoc/últimos2dígitosAno) com a coluna nrprinc.
    3. Filtra docs principais (nrdoc <= 99999 e dctramita == 'PRINCIPAL') e verifica se o nrprinc confere.
    4. Separa também as naturezas **não** habituais em arquivo distinto.
    5. Gera dois arquivos .xlsx e envia para os e-mails fornecidos.
    """)

    # Campo para digitar os e-mails de destino
    emails_input = st.text_input(
        "Digite os e-mails de destino, separados por vírgula",
        value=""
    )

    # Botão de processamento
    if st.button("Processar e Enviar E-mails"):
        if not emails_input.strip():
            st.error("Por favor, insira ao menos um endereço de e-mail.")
            return

        # Separa os e-mails por vírgula
        to_emails = [e.strip() for e in emails_input.split(",")]

        # Carrega o arquivo Excel
        try:
            df = pd.read_excel("cap3.xlsx")
        except Exception as e:
            st.error(f"Erro ao ler cap3.xlsx: {e}")
            return

        # -----------------------------------------------------------------------------
        # Passo 1: Criar a coluna "nrprinc_formatado" concatenando nrdoc-dvdoc/yy
        # -----------------------------------------------------------------------------
        def format_nrprinc(row):
            # row["andoc"] espera-se ser 4 dígitos (ex: 2012 -> "12")
            ano = str(row["andoc"])
            ano2d = ano[-2:]  # Pega últimos 2 dígitos
            return f"{row['nrdoc']}-{row['dvdoc']}/{ano2d}"

        df["nrprinc_formatado"] = df.apply(format_nrprinc, axis=1)

        # -----------------------------------------------------------------------------
        # Passo 2: Filtrar docs principais
        #    Regras:
        #      - nrdoc <= 99999
        #      - dctramita == 'PRINCIPAL'
        #      - nrprinc_formatado == nrprinc (significa que “bateu” com a junção)
        # -----------------------------------------------------------------------------
        mask_principal = (
            (df["nrdoc"] <= 99999) &
            (df["dctramita"] == "PRINCIPAL") &
            (df["nrprinc_formatado"] == df["nrprinc"])
        )
        df_principais = df.loc[mask_principal].copy()

        # -----------------------------------------------------------------------------
        # Passo 3: Filtrar naturezas não habituais
        #    - Precisamos de todos os registros CUJA 'dcgrnatureza' não esteja na lista
        #    - MAS o enunciado sugere filtrar somente na "CargaPrincipal"? Ou tudo?
        #      Aqui, vamos assumir que é geral: se "dcgrnatureza" não estiver na lista,
        #      vai para outro DataFrame.
        # -----------------------------------------------------------------------------
        mask_nao_habitual = ~df["dcgrnatureza"].isin(NATUREZAS_HABITUAIS)
        df_nao_habitual = df.loc[mask_nao_habitual].copy()

        # -----------------------------------------------------------------------------
        # Passo 4: Gerar os dois arquivos XLSX (principal e não-habitual)
        # -----------------------------------------------------------------------------
        file_principais = "docs_principais.xlsx"
        file_nao_habitual = "natureza_nao_habitual.xlsx"

        if not df_principais.empty:
            df_principais.to_excel(file_principais, index=False)
        else:
            # Cria um arquivo vazio com cabeçalho, caso não haja resultados
            pd.DataFrame(columns=df.columns).to_excel(file_principais, index=False)

        if not df_nao_habitual.empty:
            df_nao_habitual.to_excel(file_nao_habitual, index=False)
        else:
            pd.DataFrame(columns=df.columns).to_excel(file_nao_habitual, index=False)

        # -----------------------------------------------------------------------------
        # Passo 5: Enviar e-mail com dois anexos
        # -----------------------------------------------------------------------------
        subject = "Resultados - DOCS PRINCIPAIS e NATUREZA NÃO HABITUAL"
        body = (
            "Segue em anexo:\n\n"
            "1) docs_principais.xlsx => Contém os registros que batem com a junção "
            "(nrdoc-dvdoc/yy) e que possuem nrdoc <= 99999 e dctramita='PRINCIPAL'.\n"
            "2) natureza_nao_habitual.xlsx => Contém processos cuja 'dcgrnatureza' "
            "não faz parte da lista de naturezas habituais.\n\n"
            "Att,\nScript Automático"
        )

        attachment_paths = [file_principais, file_nao_habitual]
        send_email_with_attachments(
            to_emails=to_emails,
            subject=subject,
            body=body,
            attachment_paths=attachment_paths
        )

        # Mensagem de sucesso na tela
        st.success("Processo concluído e e-mails enviados com sucesso!")

# Executa a aplicação streamlit
if __name__ == "__main__":
    main()

