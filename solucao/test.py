# from mouseinfo import mouseInfo
#
# mouseInfo()
import re
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

outPut = """
HomeCadastros Suprimentos Vendas Finanças
+99
Benicio
Relatório de Movimentação de um Produto
Imprimir
Visualizar
Período
Do mês

De um período
Mês
10/2023
Produto
Verniz Tripla Proteção Suvinil 900ml  Escolha Sua Cor Cor:Mogno
Filtrar depósitos

Todos
Visualizar por:
Por dia
Movimentação de produto
Saldo
no dia
Est…
Saldo
médio
do
perí…
14/10/2023
22/10/2023
31/10/2023
4
5
6
7
Task	Saldo no dia	Estoque mínimo	Saldo médio do período
14/10/2023	7,00	4,00	5,67
22/10/2023	6,00	4,00	5,67
31/10/2023	4,00	4,00	5,67
Saldo antes de 01/10/2023	8,00
Período	Entradas	Saídas	Saldo parcial
14/10/2023	0,000	1,000	7,000
22/10/2023	0,000	1,000	6,000
31/10/2023	0,000	2,000	4,000
Totais	0,000	4,000
Organisys SoftwareVersão 17/01/2024
Verniz Tripla Proteção Suvinil 900ml Escolha Sua Cor Cor:Mogno
"""

padrao = re.compile(r'Período	Entradas(.*?)Organisys', re.DOTALL)
resultado = re.search(padrao, outPut)

# Verificar se a correspondência foi encontrada
if resultado:
    parte_desejada = "Produto: ProdutoX \n" + r"Período	Entradas	" + resultado.group(1).strip()
    print(parte_desejada)

    linhas = parte_desejada.split('\n')

    # Obtém o produto e remove a palavra "produto:" do início
    produto = linhas[0].replace("Produto:", "").strip()

    # Obtém os dados em um formato DataFrame do Pandas
    dados = [linha.split('\t') for linha in linhas[1:]]
    df = pd.DataFrame(dados, columns=['Período', 'Entradas', 'Saídas', 'Saldo parcial'])

    # Cria um novo arquivo Excel
    arquivo_excel = Workbook()
    planilha = arquivo_excel.active

    # Adiciona os dados do produto à planilha
    planilha['A1'] = 'Produto'
    planilha['B1'] = produto

    # Adiciona os títulos das colunas
    for col_num, titulo in enumerate(df.columns, start=2):
        planilha.cell(row=2, column=col_num, value=titulo)

    # Adiciona os dados à planilha
    for linha in dataframe_to_rows(df, index=False, header=False):
        planilha.append(linha)

    arquivo_excel.active.title = produto

    # Salva o arquivo Excel
    arquivo_excel.save('saida_excel.xlsx')
