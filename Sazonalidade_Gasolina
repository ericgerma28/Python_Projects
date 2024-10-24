import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Carregar o arquivo CSV (altere o caminho para o seu arquivo de dados)
df = pd.read_csv('//content//Preços semestrais - AUTOMOTIVOS_2024.01.csv', sep=';', on_bad_lines='skip')

# Exibir as primeiras linhas para verificar os dados
print(df.head())

# Substituir vírgulas por pontos apenas nas colunas que são do tipo string
df['Valor de Venda'] = df['Valor de Venda'].astype(str).str.replace(',', '.')

# Converter a coluna "Valor de Venda" para numérico, forçando erros para NaN
df['Valor de Venda'] = pd.to_numeric(df['Valor de Venda'], errors='coerce')

# Remover linhas com valores ausentes ou não numéricos em "Valor de Venda"
df = df.dropna(subset=['Valor de Venda'])

# Verificar se há dados válidos na coluna Data da Coleta
print(df[['Data da Coleta', 'Valor de Venda']].head())

# Converter a coluna "Data da Coleta" para o tipo datetime
df['Data da Coleta'] = pd.to_datetime(df['Data da Coleta'], format='%d/%m/%Y', errors='coerce')

# Remover linhas com valores ausentes em "Data da Coleta"
df = df.dropna(subset=['Data da Coleta'])

# Extrair o mês e o ano da data de coleta para fazer a análise de sazonalidade
df['AnoMes'] = df['Data da Coleta'].dt.to_period('M')

# Converter a coluna 'AnoMes' para string para facilitar o gráfico
df['AnoMes'] = df['AnoMes'].astype(str)

# Agrupar por Ano e Mês, e calcular o preço médio de venda
preco_sazonalidade = df.groupby('AnoMes')['Valor de Venda'].mean().reset_index()

# Exibir a tabela de preço médio por mês
print(preco_sazonalidade)

# Verificar se existem dados válidos antes de criar o gráfico
if not preco_sazonalidade.empty:
    # Visualizar a variação dos preços ao longo do tempo
    plt.figure(figsize=(12, 8))
    sns.lineplot(x='AnoMes', y='Valor de Venda', data=preco_sazonalidade, marker='o', color='b')
    plt.title('Análise de Sazonalidade: Preço Médio de Venda ao Longo do Tempo')
    plt.ylabel('Preço Médio (R$)')
    plt.xlabel('Ano e Mês')
    plt.xticks(rotation=45)

    plt.grid(True)
    plt.tight_layout()
    plt.show()
else:
    print("Nenhum dado válido encontrado para gerar o gráfico.")
