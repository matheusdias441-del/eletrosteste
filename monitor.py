import os
import smtplib
import tempfile
from datetime import datetime
from email.mime.text import MIMEText

import gspread
from google.oauth2.service_account import Credentials
from serpapi import GoogleSearch

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

    try:

        params = {
            "engine": "google_shopping",
            "q": produto,
            "gl": "br",
            "hl": "pt-br",
            "api_key": os.environ["SERPAPI_KEY"]
        }

        search = GoogleSearch(params)
        results = search.get_dict()

        shopping_results = results.get("shopping_results", [])

        if not shopping_results:
            return ("Não encontrado", 0, "")

        melhor_item = None
        menor_preco = None

        for item in shopping_results:

            preco_str = str(item.get("price", ""))

            preco_limpo = (
                preco_str
                .replace("R$", "")
                .replace(".", "")
                .replace(",", ".")
                .strip()
            )

            try:
                preco = float(preco_limpo)
            except:
                continue

            if menor_preco is None or preco < menor_preco:
                menor_preco = preco
                melhor_item = item

        if melhor_item is None:
            return ("Não encontrado", 0, "")

        loja = melhor_item.get("source", "Desconhecida")
        link = melhor_item.get("link", "")

        return (
            loja,
            menor_preco,
            link
        )

    except Exception as e:

        print(f"Erro ao buscar {produto}: {e}")

        return (
            "Erro",
            0,
            ""
        )


def enviar_email(
    produto,
    preco,
    loja,
    link,
    menor_antigo
):

    corpo = f"""
Novo menor preço histórico encontrado.

Produto: {produto}

Preço atual: R$ {preco:.2f}

Menor preço anterior: R$ {menor_antigo:.2f}

Loja: {loja}

Link:
{link}
"""

    msg = MIMEText(corpo)

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


def obter_ou_criar_aba(planilha, nome, linhas, colunas):

    try:
        return planilha.worksheet(nome)

    except:

        aba = planilha.add_worksheet(
            title=nome,
            rows=linhas,
            cols=colunas
        )

        return aba


def main():

    gc = conectar_google()

    planilha = gc.open(PLANILHA)

    historico = obter_ou_criar_aba(
        planilha,
        "Historico",
        5000,
        10
    )

    menores = obter_ou_criar_aba(
        planilha,
        "Menores Precos",
        100,
        10
    )

    if len(historico.get_all_values()) == 0:
        historico.append_row(
            [
                "Data",
                "Produto",
                "Loja",
                "Preco",
                "Link"
            ]
        )

    if len(menores.get_all_values()) == 0:
        menores.append_row(
            [
                "Produto",
                "Menor Preco",
                "Loja",
                "Data"
            ]
        )

    registros = menores.get_all_records()

    mapa = {}

    for registro in registros:

        produto = registro.get("Produto")

        if produto:
            mapa[produto] = registro

    hoje = datetime.now().strftime("%d/%m/%Y")

    for produto in PRODUTOS:

        print(f"Pesquisando {produto}...")

        loja, preco, link = buscar_preco(produto)

        if preco <= 0:
            continue

        historico.append_row(
            [
                hoje,
                produto,
                loja,
                preco,
                link
            ]
        )

        if produto not in mapa:

            menores.append_row(
                [
                    produto,
                    preco,
                    loja,
                    hoje
                ]
            )

            print(
                f"Primeiro registro: {produto} - R$ {preco}"
            )

        else:

            menor_antigo = float(
                mapa[produto]["Menor Preco"]
            )

            if preco < menor_antigo:

                celula = menores.find(produto)

                linha = celula.row

                menores.update(
                    f"A{linha}:D{linha}",
                    [[
                        produto,
                        preco,
                        loja,
                        hoje
                    ]]
                )

                enviar_email(
                    produto,
                    preco,
                    loja,
                    link,
                    menor_antigo
                )

                print(
                    f"Novo menor preço: {produto} - R$ {preco}"
                )

            else:

                print(
                    f"Sem novo menor preço: {produto}"
                )


if __name__ == "__main__":
    main()
