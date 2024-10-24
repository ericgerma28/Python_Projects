import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Carregar o arquivo CSV (altere o caminho para o seu arquivo de dados)
df = pd.read_csv('//content//Preços semestrais - AUTOMOTIVOS_2024.01.csv', sep=';', on_bad_lines='skip')

# Exibir as primeiras linhas para verificar os dados
print(df.head())

# Substituir vírgulas por pontos apenas nas colunas que são do tipo string
df['Valor de Venda'] = df['Valor de Venda'].astype(str).str.replace(',', '.')
df['Valor de Compra'] = df['Valor de Compra'].astype(str).str.replace(',', '.')

# Converter colunas de interesse para o tipo correto
df['Valor de Venda'] = pd.to_numeric(df['Valor de Venda'], errors='coerce')

# Remover linhas com valores ausentes em "Valor de Venda"
df = df.dropna(subset=['Valor de Venda'])

# Converter a coluna 'Data da Coleta' para formato de data
df['Data da Coleta'] = pd.to_datetime(df['Data da Coleta'], format='%d/%m/%Y', errors='coerce')

# Verificar novamente se há dados válidos
print(df[['Data da Coleta', 'Valor de Venda']].head())

# Agrupar por data e calcular a média do preço de venda
preco_temporal = df.groupby('Data da Coleta')['Valor de Venda'].mean().reset_index()

# Exibir a tabela de preço médio por data
print(preco_temporal)

# Verificar se existem dados válidos antes de criar o gráfico
if not preco_temporal.empty:
    # Visualizar com gráficos de linha a variação dos preços ao longo do tempo
    plt.figure(figsize=(10, 6))
    sns.lineplot(x='Data da Coleta', y='Valor de Venda', data=preco_temporal, marker='o')
    plt.title('Variação do Preço Médio de Venda ao Longo do Tempo')
    plt.ylabel('Preço Médio (R$)')
    plt.xlabel('Data da Coleta')
    plt.xticks(rotation=45)
    plt.grid(True)
    plt.tight_layout()
    plt.show()
else:
    print("Nenhum dado válido encontrado para gerar o gráfico.")
