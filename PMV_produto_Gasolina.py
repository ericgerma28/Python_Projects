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

# Verificar novamente se há dados válidos
print(df[['Produto', 'Bandeira', 'Valor de Venda']].head())

# Agrupar por Produto e Bandeira, e calcular o preço médio de venda
preco_produto_bandeira = df.groupby(['Produto', 'Bandeira'])['Valor de Venda'].mean().reset_index()

# Exibir a tabela de preço médio por produto e bandeira
print(preco_produto_bandeira)

# Verificar se existem dados válidos antes de criar o gráfico
if not preco_produto_bandeira.empty:
    # Visualizar com gráficos de barras a variação dos preços por Produto e Bandeira
    plt.figure(figsize=(12, 8))
    sns.barplot(x='Produto', y='Valor de Venda', hue='Bandeira', data=preco_produto_bandeira, palette='viridis')
    plt.title('Preço Médio de Venda por Produto e Bandeira')
    plt.ylabel('Preço Médio (R$)')
    plt.xlabel('Produto')
    plt.xticks(rotation=45)

    # Ajustar a posição da legenda para ficar abaixo do gráfico
    plt.legend(title='Bandeira', loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=6)  # Ajuste da posição da legenda

    plt.grid(True)
    plt.tight_layout()
    plt.show()
else:
    print("Nenhum dado válido encontrado para gerar o gráfico.")
