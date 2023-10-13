#!/usr/bin/env python
# coding: utf-8

# <h2>AUTOMACAO - IPTU - SJRP</h2>

#
# <li>Automação feita para coletar Valor, Unidade, Bloco e Cidade do IPTU de São José do Rio Preto</li>
# Importando bibliotecas

import tabula
from tabula.io import read_pdf
from IPython.display import display
import pandas as pd
import openpyxl
from openpyxl import Workbook, load_workbook
import datetime
import os
from pypdf import PdfMerger
import csv


def coletaiptu(IPTU):
    global df1

    # Coletando dados
    lista_tabelas = tabula.read_pdf(f"{IPTU}", pages="all")
    tabelaCabecalho = lista_tabelas[0]
    tabelaPagamento = lista_tabelas[4]

    verificador = tabelaCabecalho.iloc[1, 2]

    if verificador == "Lote":
        ### Coletando Lote Cabecalho
        tabelaCabecalho.columns = tabelaCabecalho.iloc[0]

        tabelaCabecalho[[0, 1]] = tabelaCabecalho.iloc[0].str.split("\r", expand=True)
        tabelaCabecalho[[2, 3]] = tabelaCabecalho.iloc[1].str.split("\r", expand=True)
        tabelaCabecalho.columns = tabelaCabecalho.iloc[0]

        tabelaCabecalho = tabelaCabecalho.dropna(axis=1)

        rua = tabelaCabecalho.iloc[1, 0]

        rua = rua.split("\r")

        rua = rua[1].split(" - ")

        unidade = rua[2]

        unidade = unidade.split("+")

        unidade = unidade[0]

        unidade = unidade.split("B")

        rua = rua[0]
        unidadeU = unidade[0]
        bloco = "B" + unidade[1]

        ###Coletando Valor do Documento
        tabelaPagamento = tabelaPagamento.dropna(axis=0)

        valor = tabelaPagamento.iloc[1, 4]
        valor = valor.split("\r")
        valor = valor[1]

        valorTotal = f"R$ {valor}"

        agora = datetime.datetime.now()

        ### Gerando Dataframe
        dfCabecalho = df = pd.DataFrame(
            {
                "index": [i],
                "Rua": [rua],
                "Cidade": "São José do Rio Preto",
                "Bloco": [bloco],
                "Unidade": [unidadeU],
                "Valor": [valorTotal],
                "Data": [agora],
            }
        )
    else:
        tabelaCabecalho.columns = tabelaCabecalho.iloc[0]

        tabelaCabecalho[[0, 1]] = tabelaCabecalho.iloc[0].str.split("\r", expand=True)
        tabelaCabecalho.columns = tabelaCabecalho.iloc[0]

        rua = tabelaCabecalho.iloc[1, 0]

        rua = rua.split("\r")

        rua = rua[1]

        bloco = ""

        bloco = tabelaCabecalho.iloc[1, 1]

        bloco = bloco.split("\r")

        bloco = bloco[0] + " " + bloco[1]

        unidade = tabelaCabecalho.iloc[1, 2]

        unidade = unidade.split("\r")

        unidade = unidade[0] + " " + unidade[1]

        valor = tabelaPagamento.iloc[2, 4]

        valor = valor.split("\r")

        valor = f"R$ {valor[1]}"

        agora = datetime.datetime.now()

        ### Gerando Dataframe
        dfCabecalho = df = pd.DataFrame(
            {
                "index": [i],
                "Rua": [rua],
                "Cidade": "São José do Rio Preto",
                "Bloco": [bloco],
                "Unidade": [unidade],
                "Valor": [valor],
                "Data": [agora],
            }
        )

    if i == 0:
        df1 = dfCabecalho
    else:
        df1 = pd.concat([df1, dfCabecalho])

    return df1


caminhos = [
    os.path.join(
        r"caminho dos PDFs",
        nome,
    )
    for nome in os.listdir(
        r"caminho dos PDFs"
    )
]

arquivos = [arq for arq in caminhos if os.path.isfile(arq)]

pdfs = [arq for arq in arquivos if arq.endswith(".pdf")]

pdfs = list(
    map(
        lambda x: x.replace(
            "caminho dos PDFs",
            "",
        ),
        pdfs,
    )
)


def exportExcel(writer):
    writer = pd.ExcelWriter("SJRP.xlsx")

    # write dataframe to excel
    df1.to_excel(writer, index=False)

    # save the excel
    writer.close()


i = 0
df1 = 0
while i < len(pdfs):
    df1 = coletaiptu(pdfs[i])
    i += 1

exportExcel(df1)

agora = datetime.datetime.now()

ano = agora.year
dia = agora.day
mês = agora.month
hora = agora.hour
minuto = agora.minute

merger = PdfMerger()

for pdf in pdfs:
    merger.append(pdf)

merger.write(f"IPTU_{dia}-{mês}-{ano}_{hora}-{minuto}_Unit.pdf")
merger.close()


tabela = pd.read_excel("SJRP.xlsx")
display(tabela)
