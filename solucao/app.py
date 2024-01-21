import pyautogui
import time
import pyperclip
import re
from io import BytesIO
import openpyxl

edt_produto_x = 502
edt_produto_y = 263

bt_visualizar_x = 991
bt_visualizar_y = 203

bt_selecao_x = 515
bt_seleccao_y = 291

# > pyinstaller  spec.spec

input("""
                               ##############
                                # ATENÇÃO #
                               #############
   - Esteja na página "Relatório de Movimentação de um Produto"
   - Garante que o arquivo "produtos.txt" esteja na mesma pasta do programa.
   - Coloque as configações do relatório que deseja buscar dos produtos.
   
   @@@@@@@@@@@@@@@@@@@@@@ Aperte ENTER para começar @@@@@@@@@@@@@@@@@@@@@@@@@@
   
""")


def limpar_nome(nome):
    return re.sub(r'[\/:*?"<>|]', '_', nome)

with open('produtos.txt', 'r', encoding='utf-8') as arquivo:
    linhas_do_arquivo = arquivo.readlines()

produtos = []
for linha in linhas_do_arquivo:

    time.sleep(2)

    pyautogui.click(x=edt_produto_x, y=edt_produto_y)

    time.sleep(1)

    # Seleciona todo o texto na caixa de pesquisa e pressiona Delete para limpar
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')

    # Escreve o novo texto na caixa de pesquisa
    pyperclip.copy(linha)
    pyautogui.hotkey('ctrl', 'v')

    time.sleep(2.5)
    pyautogui.click(x=bt_selecao_x, y=bt_seleccao_y)

    time.sleep(1)
    pyautogui.click(x=bt_visualizar_x, y=bt_visualizar_y)
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')

    conteudo = pyperclip.paste()
    padrao = re.compile(r'Período	Entradas(.*?)Organisys', re.DOTALL)

    resultado = re.search(padrao, conteudo)

    if resultado:
        parte_desejada = r"Período	Entradas	" + resultado.group(1).strip()

        produtos.append({
            "produto": linha,
            "parte_desejada": parte_desejada
        })

wb = openpyxl.Workbook()

# Itera sobre cada produto
for produto in produtos:
    # Divide o produto em linhas
    linhas_produto = produto.get('parte_desejada').split("\n")

    # Obtém o nome do produto (a linha que contém "produto:")
    nome_produto = limpar_nome(produto.get('produto'))

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
    with open("novo_relatório_movimentação.xlsx", "wb") as excel_file:
        excel_file.write(excel_content)

input("Arquivo Excel Criado com Sucesso!\nEnter para Sair")
