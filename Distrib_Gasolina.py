import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Carregar o arquivo CSV (altere o caminho para o seu arquivo de dados)
df = pd.read_csv('//content//Preços semestrais - AUTOMOTIVOS_2024.01.csv', sep=';', on_bad_lines='skip')

# Substituir vírgulas por pontos nas colunas de valores
df['Valor de Venda'] = df['Valor de Venda'].astype(str).str.replace(',', '.')
df['Valor de Compra'] = df['Valor de Compra'].astype(str).str.replace(',', '.')

# Converter as colunas "Valor de Venda" para numérico
df['Valor de Venda'] = pd.to_numeric(df['Valor de Venda'], errors='coerce')

# Remover linhas com valores ausentes em "Valor de Venda"
df = df.dropna(subset=['Valor de Venda'])

# Exibir as primeiras linhas para verificar os dados
print(df[['Regiao - Sigla', 'Produto', 'Valor de Venda']].head())

# Análise da distribuição de preços por Região
plt.figure(figsize=(12, 6))
sns.histplot(df, x='Valor de Venda', hue='Regiao - Sigla', multiple='stack', bins=30, palette='viridis')
plt.title('Distribuição dos Preços de Venda por Região')
plt.xlabel('Preço de Venda (R$)')
plt.ylabel('Frequência')
plt.grid(True)
plt.tight_layout()
plt.show()

# Análise da distribuição de preços por Produto
plt.figure(figsize=(12, 6))
sns.histplot(df, x='Valor de Venda', hue='Produto', multiple='stack', bins=30, palette='coolwarm')
plt.title('Distribuição dos Preços de Venda por Produto')
plt.xlabel('Preço de Venda (R$)')
plt.ylabel('Frequência')
plt.grid(True)
plt.tight_layout()
plt.show()
