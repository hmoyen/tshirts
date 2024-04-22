import gspread
import pandas as pd

gc = gspread.service_account(filename='/home/lena/camisetas/camisetas-421115-29a129e1e422.json')

sheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/1pOwDcS8K-NmGAzcw7x8UO1V3FJPFbpU5AVHpKdlJzao/edit#gid=441120012').sheet1


# Definir os cabeçalhos esperados
expected_headers = ['Carimbo de data/hora', 'Endereço de e-mail', 'Nome completo',
                    'CAMISETA MODALIDADE TÊNIS AZUL E BRANCA', 'Se personalizada (1), qual o nome atrás?',
                    'CAMISETA MODALIDADE TÊNIS AZUL E AMARELA', 'Se personalizada (2), qual o nome atrás?',
                    'CAMISETA MODALIDADE TÊNIS PRETA E AMARELA', 'Se personalizada (3), qual o nome atrás?',
                    'CAMISETA RAQUETES PRETA', 'CAMISETA FEDERER BRANCA']

# Obter todos os dados
data = sheet.get_all_records(expected_headers=expected_headers)

# Manipular os dados para criar os CSVs
camisetas_personalizadas = []
total_camisetas = {}
pedido_por_pessoa = {}

for row in data:
    # Verifica se há personalização em cada camiseta e adiciona ao CSV de camisetas personalizadas
    personalizacao_azul_branca = row['Se personalizada (1), qual o nome atrás?']
    personalizacao_azul_amarela = row['Se personalizada (2), qual o nome atrás?']
    personalizacao_preta_amarela = row['Se personalizada (3), qual o nome atrás?']
    
    if personalizacao_azul_branca:
        camisetas_personalizadas.append({'Nome': row['Nome completo'], 'Camiseta': 'CAMISETA MODALIDADE TÊNIS AZUL E BRANCA',
                                         'Tamanho': row['CAMISETA MODALIDADE TÊNIS AZUL E BRANCA'],
                                         'Personalização': personalizacao_azul_branca})
    if personalizacao_azul_amarela:
        camisetas_personalizadas.append({'Nome': row['Nome completo'], 'Camiseta': 'CAMISETA MODALIDADE TÊNIS AZUL E AMARELA',
                                         'Tamanho': row['CAMISETA MODALIDADE TÊNIS AZUL E AMARELA'],
                                         'Personalização': personalizacao_azul_amarela})
    if personalizacao_preta_amarela:
        camisetas_personalizadas.append({'Nome': row['Nome completo'], 'Camiseta': 'CAMISETA MODALIDADE TÊNIS PRETA E AMARELA',
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
    nome_completo = row['Nome completo']
    if nome_completo not in pedido_por_pessoa:
        pedido_por_pessoa[nome_completo] = {'Total camisetas': 0, 'Total personalizacoes': 0,
                                             'Valor camisetas': 0, 'Valor personalizacoes': 0}
    for camiseta, tamanho in row.items():
        if camiseta.startswith('CAMISETA') and tamanho:
            pedido_por_pessoa[nome_completo][camiseta] = tamanho
            pedido_por_pessoa[nome_completo]['Total camisetas'] += 1
            if 'personalizacao' in camiseta.lower():
                pedido_por_pessoa[nome_completo]['Total personalizacoes'] += 1

# Criar DataFrame para camisetas personalizadas e total de camisetas
camisetas_personalizadas_df = pd.DataFrame(camisetas_personalizadas)
total_camisetas_df = pd.DataFrame(total_camisetas)

# Calcular valor total do pedido por pessoa
valor_camiseta = 40
valor_personalizacao = 15

# Carregar o CSV das camisetas personalizadas
contagem_nomes = camisetas_personalizadas_df['Nome'].value_counts()

# Calcular o número total de personalizações para cada pessoa
for nome, contagem in contagem_nomes.items():
    if nome in pedido_por_pessoa:
        pedido_por_pessoa[nome]['Total personalizacoes'] = contagem

# Recalcular o valor total do pedido por pessoa
for nome, pedido in pedido_por_pessoa.items():
    pedido_por_pessoa[nome]['Valor camisetas'] = pedido['Total camisetas'] * valor_camiseta
    pedido_por_pessoa[nome]['Valor personalizacoes'] = pedido['Total personalizacoes'] * valor_personalizacao
    pedido_por_pessoa[nome]['Valor total'] = pedido['Valor camisetas'] + pedido['Valor personalizacoes']

# Converter pedido_por_pessoa em DataFrame
pedido_por_pessoa_df = pd.DataFrame.from_dict(pedido_por_pessoa, orient='index')

# Salvar em CSV
camisetas_personalizadas_df.to_csv('camisetas_personalizadas.csv', index=False)
total_camisetas_df.to_csv('total_camisetas.csv')
pedido_por_pessoa_df.to_csv('pedido_por_pessoa.csv')