import matplotlib.pyplot as plt

# Agrupar por ano e contar a quantidade de músicas lançadas por ano
songs_per_year = df.groupby('released_year').size()

# Agrupar por mês e contar a quantidade de músicas lançadas por mês
songs_per_month = df.groupby('released_month').size()

# Exibir as informações antes de plotar
print("\nQuantidade de músicas lançadas por ano:")
print(songs_per_year)

print("\nQuantidade de músicas lançadas por mês:")
print(songs_per_month)

# Plotar as tendências de lançamento por ano
plt.figure(figsize=(12, 6))
plt.plot(songs_per_year.index, songs_per_year.values, marker='o', linestyle='-', color='b')
plt.title('Tendência de Lançamentos por Ano')
plt.xlabel('Ano')
plt.ylabel('Quantidade de Músicas Lançadas')
plt.grid(True)
plt.show()

# Plotar as tendências de lançamento por mês
plt.figure(figsize=(12, 6))
plt.bar(songs_per_month.index, songs_per_month.values, color='g')
plt.title('Tendência de Lançamentos por Mês')
plt.xlabel('Mês')
plt.ylabel('Quantidade de Músicas Lançadas')
plt.xticks(range(1, 13))  # Para mostrar os meses de 1 a 12
plt.grid(True)
plt.show()
