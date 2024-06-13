import pandas as pd
import openpyxl
from openpyxl import load_workbook
import os
import xlwings as xw

# Metodo com a lógica de filtragem
def Filtragem(empresa, cliente):
    tec_cliente_1 = {"BANRISUL", "BRBPAY", "C&A MODAS",
                     "CONDUCTOR VOXCRED", "ENTRE WHITE"}
    tec_cliente_2 = {"VERO – BANRISUL", "ZIPBANK"}
    dominix_cliente = {"ACQIO", "CIELO", "GRANITO", "HYPERLOCAL"}
    ogea_cliente = {"ACQIO", "BANCO PAN", "BANCO SICOOB", "BOLT CARD", "C6 Bank", "CIELO", "CLARO DECODER",
                    "CrefisaPay", "ENTRE PAYMENT", "GETNET", "GILBARCO", "GPON", "GRANITO", "JUSTA PAGAMENTOS",
                    "KONCENTRI", "MAGALU", "PAGBEM", "PAGSEGURO", "PUNTO", "QUERO 2 PAY", "R4 Bank", "REDE",
                    "STONE", "TNS", "TON PAGAMENTOS", "TRIPAG"}

    if empresa == "TECNOLOGIA":
        if cliente in tec_cliente_1:
            return "PDF TEC"
        elif cliente in tec_cliente_2:
            return "PDF TEC (COMODATO)"
    elif empresa == "DOMINIX":
        if cliente in dominix_cliente:
            return "PDF DOMINIX"
    elif empresa == "ÓGEA":
        if cliente in ogea_cliente:
            return "PDF OGEA"
    return "Cliente ou empresa não encontrados"
# Metodo que encontra o primeiro arquivo .xlsx

def encontrar_xlsx():
    for arquivo in os.listdir('.'):
        if arquivo.endswith('.xlsx'):
            return arquivo
    raise FileNotFoundError(
        "Nenhum arquivo .xlsx encontrado no diretório atual.")


# Nomes dos arquivos e etc
nome_arquivo = encontrar_xlsx()
sheet_name1 = "Modelo Gerar Saldo"

tabela1 = pd.read_excel(nome_arquivo, sheet_name=sheet_name1)

# capturando dados da tabela
i = 0
for index,  row in tabela1.iterrows():
    # pegando atributos da tabela ,-,
    empresa = row['Empresa']
    cliente = row['Cliente']
    prestador = row['Prestador']
    o_r = row['OR']
    # Metodo de filtragem
    resultado = Filtragem(empresa, cliente)
    # metodo "prencher"
    tabela2 = load_workbook(nome_arquivo)
    nome_sheet = tabela2[resultado]

    nome_sheet['E7(obs:mudar esse dado)'] = o_r
    nome_sheet['F7(obs:mudar esse dado)'] = prestador
    # Salvar as mudanças no arquivo Excel
    tabela2.save(f'{o_r}.xlsx')
    # convertendo .xlsx para .pdf
    wb = xw.Book(f'{o_r}.xlsx')
    sheet = wb.sheets[resultado]
    sheet.to_pdf(f'{o_r}.pdf')
    wb.close()
    i += 1
