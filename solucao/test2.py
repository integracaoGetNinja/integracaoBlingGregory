# from mouseinfo import mouseInfo
#
# mouseInfo()
import re
from io import BytesIO
import openpyxl

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
    parte_desejada = """
Período	Entradas	Saídas	Saldo parcial
14/10/2023	3,000	1,000	7,000
22/10/2023	4,000	1,000	6,000
31/10/2023	5,000	2,000	4,000
Totais	4,000	4,000
SEPARACAODEPRODUTOS
Período	Entradas	Saídas	Saldo parcial
14/10/2023	0,000	41,000	37,000
22/10/2023	0,000	51,000	26,000
31/10/2023	0,000	62,000	14,000
Totais	10,000	44,000
SEPARACAODEPRODUTOS
Período	Entradas	Saídas	Saldo parcial
14/10/2023	0,000	3,000	8,000
22/10/2023	0,000	4,000	9,000
31/10/2023	0,000	5,000	10,000
Totais	11,000	4,000
    """
    print(parte_desejada)

    produtos_raw = parte_desejada.split("SEPARACAODEPRODUTOS")

    # Remove linhas vazias resultantes da divisão
    produtos = [produto.strip() for produto in produtos_raw if produto.strip()]

    # Cria um arquivo Excel
    wb = openpyxl.Workbook()

    # Itera sobre cada produto
    for produto in produtos:
        # Divide o produto em linhas
        linhas_produto = produto.split("\n")

        # Obtém o nome do produto (a linha que contém "produto:")
        nome_produto = "produtoTeste"

        # Cria uma nova planilha para o produto
        planilha = wb.create_sheet(title=nome_produto)

        # Adiciona os cabeçalhos da tabela
        cabecalhos = ["Período", "Entradas", "Saídas", "Saldo parcial"]
        planilha.append(cabecalhos)

        # Adiciona os dados à tabela
        for linha_dados in linhas_produto[1:]:
            dados = linha_dados.split("\t")
            planilha.append(dados)

    # Salva o arquivo Excel
    with BytesIO() as excel_buffer:
        wb.save(excel_buffer)
        excel_content = excel_buffer.getvalue()

        # Escreve o conteúdo do Excel para um arquivo
        with open("saida_excel.xlsx", "wb") as excel_file:
            excel_file.write(excel_content)

    print("Arquivo Excel criado com sucesso!")