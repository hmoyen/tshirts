import gspread
import pandas as pd
from gspread_dataframe import set_with_dataframe
import gspread_formatting as gsf

# Autenticar no Google Sheets
gc = gspread.service_account(filename='/home/lena/tshirts/camisetas-421115-29a129e1e422.json')

# Abrir a planilha
spreadsheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/1pOwDcS8K-NmGAzcw7x8UO1V3FJPFbpU5AVHpKdlJzao/edit#gid=441120012')

# Obter a primeira planilha da planilha
sheet = spreadsheet.sheet1
data = sheet.get_all_records()

# Apagar as planilhas existentes
for worksheet in spreadsheet.worksheets():
    if worksheet.title != "respostas":
        spreadsheet.del_worksheet(worksheet)

# Definir os cabeçalhos esperados
expected_headers = ['Carimbo de data/hora', 'Endereço de e-mail', 'Nome completo',
                    'CAMISETA MODALIDADE TÊNIS AZUL E BRANCA', 'Se personalizada (1), qual o nome atrás?',
                    'CAMISETA MODALIDADE TÊNIS AZUL E AMARELA', 'Se personalizada (2), qual o nome atrás?',
                    'CAMISETA MODALIDADE TÊNIS PRETA E AMARELA', 'Se personalizada (3), qual o nome atrás?',
                    'CAMISETA RAQUETES PRETA', 'CAMISETA FEDERER BRANCA']

# Criar dicionários para armazenar os dados processados
camisetas_personalizadas = []
total_camisetas = {}
pedido_por_pessoa = {}

# Processar os dados
for row in data:
    # Normalizar o nome completo
    nome_completo = row['Nome completo'].strip().lower()
    
    # Verificar se há personalização em cada camiseta e adicioná-la à lista de camisetas personalizadas
    personalizacao_azul_branca = row['Se personalizada (1), qual o nome atrás?']
    personalizacao_azul_amarela = row['Se personalizada (2), qual o nome atrás?']
    personalizacao_preta_amarela = row['Se personalizada (3), qual o nome atrás?']
    
    if personalizacao_azul_branca:
        camisetas_personalizadas.append({'Nome': nome_completo, 'Camiseta': 'CAMISETA MODALIDADE TÊNIS AZUL E BRANCA',
                                         'Tamanho': row['CAMISETA MODALIDADE TÊNIS AZUL E BRANCA'],
                                         'Personalização': personalizacao_azul_branca})
    if personalizacao_azul_amarela:
        camisetas_personalizadas.append({'Nome': nome_completo, 'Camiseta': 'CAMISETA MODALIDADE TÊNIS AZUL E AMARELA',
                                         'Tamanho': row['CAMISETA MODALIDADE TÊNIS AZUL E AMARELA'],
                                         'Personalização': personalizacao_azul_amarela})
    if personalizacao_preta_amarela:
        camisetas_personalizadas.append({'Nome': nome_completo, 'Camiseta': 'CAMISETA MODALIDADE TÊNIS PRETA E AMARELA',
                                         'Tamanho': row['CAMISETA MODALIDADE TÊNIS PRETA E AMARELA'],
                                         'Personalização': personalizacao_preta_amarela})
    
    # Calcular o total de cada modelo de camiseta
    for camiseta in row.keys():
        if camiseta.startswith('CAMISETA'):
            tamanho_camiseta = row[camiseta]
            if tamanho_camiseta:
                if camiseta not in total_camisetas:
                    total_camisetas[camiseta] = {}
                if tamanho_camiseta not in total_camisetas[camiseta]:
                    total_camisetas[camiseta][tamanho_camiseta] = 0
                total_camisetas[camiseta][tamanho_camiseta] += 1
    
    # Criar pedido por pessoa
    if nome_completo not in pedido_por_pessoa:
        pedido_por_pessoa[nome_completo] = {'Nome completo': nome_completo, 'Total camisetas': 0, 'Total personalizacoes': 0,
                                             'Valor camisetas': 0, 'Valor personalizacoes': 0,
                                             'CAMISETA MODALIDADE TÊNIS AZUL E BRANCA': '',
                                             'CAMISETA MODALIDADE TÊNIS AZUL E AMARELA': '',
                                             'CAMISETA MODALIDADE TÊNIS PRETA E AMARELA': '',
                                             'CAMISETA RAQUETES PRETA': '',
                                             'CAMISETA FEDERER BRANCA': ''}
    for camiseta, tamanho in row.items():
        if camiseta.startswith('CAMISETA') and tamanho:
            if pedido_por_pessoa[nome_completo][camiseta]:
                pedido_por_pessoa[nome_completo][camiseta] += ',' + tamanho
            else:
                pedido_por_pessoa[nome_completo][camiseta] = tamanho
            pedido_por_pessoa[nome_completo]['Total camisetas'] += 1

# Calcular o número de personalizações para cada pessoa
for pessoa in pedido_por_pessoa.values():
    nome_pessoa = pessoa['Nome completo']
    num_personalizacoes = sum(1 for camiseta in camisetas_personalizadas if camiseta['Nome'] == nome_pessoa)
    pessoa['Total personalizacoes'] = num_personalizacoes

# Calcular o valor total do pedido por pessoa
valor_camiseta = 40.60
valor_personalizacao = 15

for pessoa in pedido_por_pessoa.values():
    pessoa['Valor camisetas'] = pessoa['Total camisetas'] * valor_camiseta
    pessoa['Valor personalizacoes'] = pessoa['Total personalizacoes'] * valor_personalizacao
    pessoa['Valor total'] = pessoa['Valor camisetas'] + pessoa['Valor personalizacoes']

# Converter a lista de camisetas personalizadas em DataFrame
camisetas_personalizadas_df = pd.DataFrame(camisetas_personalizadas)

# Converter o total de camisetas em DataFrame
total_camisetas_df = pd.DataFrame(total_camisetas)

# Adicionar coluna 'Tamanho' ao DataFrame 'total_camisetas_df'
total_camisetas_df['Tamanho'] = total_camisetas_df.index

# Renomear índice para 'Tamanho' no DataFrame 'total_camisetas_df'
total_camisetas_df = total_camisetas_df.rename_axis('Tamanho', axis='columns')

# Converter o pedido por pessoa em DataFrame
pedido_por_pessoa_df = pd.DataFrame.from_dict(pedido_por_pessoa, orient='index')

# Criar páginas na planilha para cada DataFrame
worksheet_personalizadas = spreadsheet.add_worksheet(title="Camisetas Personalizadas", rows=1000, cols=20)
set_with_dataframe(worksheet_personalizadas, camisetas_personalizadas_df)

worksheet_total_camisetas = spreadsheet.add_worksheet(title="Total de Camisetas", rows=1000, cols=20)
set_with_dataframe(worksheet_total_camisetas, total_camisetas_df)

worksheet_pedido_pessoa = spreadsheet.add_worksheet(title="Pedido por Pessoa", rows=1000, cols=20)
set_with_dataframe(worksheet_pedido_pessoa, pedido_por_pessoa_df)

# Definir tons de azul para as cores de fundo
background_colors = [gsf.Color(0.8, 0.9, 1), gsf.Color(0.7, 0.85, 1)]

# Aplicar formatação
gsf.format_cell_range(worksheet_personalizadas, 'A1:Z1', gsf.cellFormat(textFormat=gsf.textFormat(bold=True)))
gsf.format_cell_range(worksheet_total_camisetas, 'A1:Z1', gsf.cellFormat(textFormat=gsf.textFormat(bold=True)))
gsf.format_cell_range(worksheet_pedido_pessoa, 'A1:Z1', gsf.cellFormat(textFormat=gsf.textFormat(bold=True)))

gsf.format_cell_range(worksheet_personalizadas, 'A2:Z1000', gsf.cellFormat(backgroundColor=background_colors[0]))
gsf.format_cell_range(worksheet_total_camisetas, 'A2:Z1000', gsf.cellFormat(backgroundColor=background_colors[0]))
gsf.format_cell_range(worksheet_pedido_pessoa, 'A2:Z1000', gsf.cellFormat(backgroundColor=background_colors[0]))

gsf.format_cell_range(worksheet_personalizadas, 'A:Z', gsf.cellFormat(textFormat=gsf.textFormat(fontFamily="Nunito")))
gsf.format_cell_range(worksheet_total_camisetas, 'A:Z', gsf.cellFormat(textFormat=gsf.textFormat(fontFamily="Nunito")))
gsf.format_cell_range(worksheet_pedido_pessoa, 'A:Z', gsf.cellFormat(textFormat=gsf.textFormat(fontFamily="Nunito")))

gsf.format_cell_range(worksheet_personalizadas, 'A:Z', gsf.cellFormat(borders=gsf.Borders(top=gsf.Border(style='SOLID'), 
                                                                                            bottom=gsf.Border(style='SOLID'), 
                                                                                            left=gsf.Border(style='SOLID'), 
                                                                                            right=gsf.Border(style='SOLID'))))
gsf.format_cell_range(worksheet_total_camisetas, 'A:Z', gsf.cellFormat(borders=gsf.Borders(top=gsf.Border(style='SOLID'), 
                                                                                             bottom=gsf.Border(style='SOLID'), 
                                                                                             left=gsf.Border(style='SOLID'), 
                                                                                             right=gsf.Border(style='SOLID'))))
gsf.format_cell_range(worksheet_pedido_pessoa, 'A:Z', gsf.cellFormat(borders=gsf.Borders(top=gsf.Border(style='SOLID'), 
                                                                                         bottom=gsf.Border(style='SOLID'), 
                                                                                         left=gsf.Border(style='SOLID'), 
                                                                                         right=gsf.Border(style='SOLID'))))

# Intercale as linhas com cores diferentes
for worksheet in [worksheet_pedido_pessoa]:
    rows = worksheet.get_all_values()
    for i, row in enumerate(rows):
        if i % 2 == 0:
            gsf.format_cell_range(worksheet, f'A{i+1}:Z{i+1}', gsf.cellFormat(backgroundColor=background_colors[1]))

