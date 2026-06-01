import os
import json
import smtplib
import tempfile
from datetime import datetime
from email.mime.text import MIMEText

import gspread
from google.oauth2.service_account import Credentials

PRODUTOS = [
    "IM7S",
    "IB7S",
    "IE60P",
    "LS10E",
    "OE8GH",
    "ME23S",
    "PE12P",
    "WD11A4453BX",
    "DPS161IX"
]

PLANILHA = "iateste"

EMAIL_DESTINO = "matheusdias441@gmail.com"


def conectar_google():
    creds_json = os.environ["GOOGLE_CREDENTIALS"]

    with tempfile.NamedTemporaryFile(delete=False, suffix=".json") as f:
        f.write(creds_json.encode())
        caminho = f.name

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = Credentials.from_service_account_file(
        caminho,
        scopes=scopes
    )

    return gspread.authorize(creds)


def buscar_preco(produto):
    """
    IMPLEMENTAÇÃO INICIAL.

    Substituir futuramente por scraping/API.

    Retorna:
    (loja, preco, link)
    """

    precos_mock = {
        "IM7S": 5499,
        "IB7S": 4699,
        "IE60P": 2299,
        "LS10E": 3099,
        "OE8GH": 1849,
        "ME23S": 599,
        "PE12P": 699,
        "WD11A4453BX": 4199,
        "DPS161IX": 349
    }

    return (
        "Pesquisa Web",
        precos_mock.get(produto, 0),
        f"https://www.google.com/search?q={produto}"
    )


def enviar_email(produto, preco, loja, link, menor_antigo):
    msg = MIMEText(
        f"""
Novo menor preço encontrado.

Produto: {produto}

Preço atual: R$ {preco}

Menor preço anterior: R$ {menor_antigo}

Loja: {loja}

Link:
{link}
"""
    )

    msg["Subject"] = f"Novo menor preço - {produto}"
    msg["From"] = os.environ["EMAIL_USER"]
    msg["To"] = EMAIL_DESTINO

    servidor = smtplib.SMTP("smtp.gmail.com", 587)
    servidor.starttls()
    servidor.login(
        os.environ["EMAIL_USER"],
        os.environ["EMAIL_PASSWORD"]
    )
    servidor.send_message(msg)
    servidor.quit()


def main():

    gc = conectar_google()

    planilha = gc.open(PLANILHA)

    try:
        historico = planilha.worksheet("Historico")
    except:
        historico = planilha.add_worksheet(
            title="Historico",
            rows=5000,
            cols=10
        )

        historico.append_row(
            ["Data", "Produto", "Loja", "Preco", "Link"]
        )

    try:
        menores = planilha.worksheet("Menores Precos")
    except:
        menores = planilha.add_worksheet(
            title="Menores Precos",
            rows=100,
            cols=10
        )

        menores.append_row(
            ["Produto", "Menor Preco", "Loja", "Data"]
        )

    registros = menores.get_all_records()

    mapa = {}

    for r in registros:
        mapa[r["Produto"]] = r

    hoje = datetime.now().strftime("%d/%m/%Y")

    for produto in PRODUTOS:

        loja, preco, link = buscar_preco(produto)

        historico.append_row(
            [hoje, produto, loja, preco, link]
        )

        if produto not in mapa:

            menores.append_row(
                [produto, preco, loja, hoje]
            )

        else:

            menor_antigo = float(mapa[produto]["Menor Preco"])

            if preco < menor_antigo:

                celula = menores.find(produto)

                linha = celula.row

                menores.update(
                    f"A{linha}:D{linha}",
                    [[produto, preco, loja, hoje]]
                )

                enviar_email(
                    produto,
                    preco,
                    loja,
                    link,
                    menor_antigo
                )


if __name__ == "__main__":
    main()
