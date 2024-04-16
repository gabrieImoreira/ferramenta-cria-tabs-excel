from openpyxl import load_workbook
# variáveis

aba_principal = 'aba principal'
tabela_principal = 'TabelaBase'

if __name__ == "__main__":
    # Ler o arquivo de entrada
    with open('params.txt', 'r') as file:
        for line in file:
            if '=' in line:
                key, value = line.strip().split('=')
                if key.strip() == 'ARQUIVO_ENTRADA':
                    arquivo_entrada = value.strip()
                elif key.strip() == 'ARQUIVO_SAIDA':
                    arquivo_saida = value.strip()

    if not arquivo_entrada.endswith('.xlsx'):
        raise ValueError('O arquivo de entrada deve ser um arquivo Excel (.xlsx)')
    # Abrindo a primeira aba pra coleta de informações
    wb = load_workbook(arquivo_entrada, data_only=True)
    ws = wb[aba_principal]
    # print(ws.tables.items())
    table_range = ws.tables[tabela_principal]
    print(table_range.ref)
    linhas_tabela = []
    # Iterar sobre as linhas da tabela
    for linha_excel in ws[table_range.ref]:
        # Lista para armazenar os valores da linha
        linha = []
        # Iterar sobre as células da linha
        for cell in linha_excel:
            if cell.data_type == 'f':
                print('yesss', cell.value)
                valor = cell.get_explicit_value()
                linha.append(valor)
            else:
                linha.append(cell.value)
        # Adicionar a lista de valores da linha à lista de linhas da tabela
        linhas_tabela.append(linha)

    # Imprimir as listas de cada linha da tabela
    print(linhas_tabela[0])
    print(linhas_tabela[1])
    print(linhas_tabela[2])
    # for linha in linhas_tabela:
    #     print(linha)
    #     break
